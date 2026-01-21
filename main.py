import os
import re
import base64
from io import BytesIO
from typing import Dict, Optional, Tuple, List
from datetime import datetime
from copy import deepcopy

from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
import boto3
from botocore.config import Config
from docx import Document
from docx.shared import Inches, Pt, Twips
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image

app = FastAPI(title="IDS Template Filler", version="1.1.0")

# S3 Configuration
S3_ENDPOINT = "https://nyc3.digitaloceanspaces.com"
S3_BUCKET = "fam.workspace"
S3_REGION = "nyc3"

# Initialize S3 client
s3_client = boto3.client(
    "s3",
    endpoint_url=S3_ENDPOINT,
    aws_access_key_id=os.getenv("S3_ACCESS_KEY"),
    aws_secret_access_key=os.getenv("S3_SECRET_KEY"),
    region_name=S3_REGION,
    config=Config(s3={'addressing_style': 'path'})  # Handle dot in bucket name
)

# Image dimension constraints
MAX_WIDTH_INCHES = 6.5   # Content area of letter page with 1" margins
MAX_HEIGHT_INCHES = 8.0  # Leave room for captions/spacing and prevent page breaks

# Image width configuration (fallback values)
IMAGE_WIDTHS = {
    "IMAGE_SOURCES_USES": 6.5,
    "IMAGE_CAPITAL_STACK_CLOSING": 6.5,
    "IMAGE_LOAN_TO_COST": 6.0,
    "IMAGE_LTV_LTC": 6.0,
    "IMAGE_AERIAL_MAP": 3.25,      # Was 5.0
    "IMAGE_LOCATION_MAP": 3.25,    # Was 5.0
    "IMAGE_REGIONAL_MAP": 3.25,    # Was 5.0
    "IMAGE_SITE_PLAN": 5.5,
    "IMAGE_PILOT_SCHEDULE": 6.0,
    "IMAGE_TAKEOUT_SIZING": 6.0,
    "IMAGE_STREET_VIEW": 5.5,      # Add this line
}

# Font formatting constants
FONT_NAME = "Times New Roman"
FONT_SIZE_11PT = Pt(11)  # Default for all body text placeholders


def calculate_image_dimensions(image_bytes: bytes, preferred_width: float) -> Tuple[float, float]:
    """
    Calculate optimal image dimensions that fit within page constraints.
    
    Args:
        image_bytes: The image data
        preferred_width: Preferred width in inches
        
    Returns:
        Tuple of (width_inches, height_inches) that maintains aspect ratio
        and fits within page constraints
    """
    try:
        # Get original image dimensions
        image = Image.open(BytesIO(image_bytes))
        original_width, original_height = image.size
        
        # Calculate original aspect ratio
        aspect_ratio = original_height / original_width
        
        # Start with preferred width, but constrain to max width
        width_inches = min(preferred_width, MAX_WIDTH_INCHES)
        height_inches = width_inches * aspect_ratio
        
        # If height is too tall, scale down to fit max height
        if height_inches > MAX_HEIGHT_INCHES:
            height_inches = MAX_HEIGHT_INCHES
            width_inches = height_inches / aspect_ratio
            
            # Make sure width still fits after height adjustment
            if width_inches > MAX_WIDTH_INCHES:
                width_inches = MAX_WIDTH_INCHES
                height_inches = width_inches * aspect_ratio
        
        return width_inches, height_inches
        
    except Exception as e:
        # Fallback to safe defaults if image processing fails
        print(f"Warning: Could not process image dimensions: {e}")
        return min(preferred_width, MAX_WIDTH_INCHES), min(4.0, MAX_HEIGHT_INCHES)


class FillRequest(BaseModel):
    placeholders: Dict[str, str] = {}
    images: Dict[str, str] = {}
    template_key: str = "_Templates/IDS_Template_Fairbridge.docx"
    output_filename: str = "IDS_Generated.docx"


class FillAndUploadRequest(BaseModel):
    placeholders: Dict[str, str] = {}
    images: Dict[str, str] = {}
    template_key: str = "_Templates/IDS_Template_Fairbridge.docx"
    output_key: str


@app.get("/health")
async def health_check():
    return {"status": "ok", "version": "1.1.0"}


def download_template(template_key: str) -> bytes:
    """Download template from S3."""
    try:
        response = s3_client.get_object(Bucket=S3_BUCKET, Key=template_key)
        return response['Body'].read()
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"Template not found: {template_key}")


def get_unique_output_key(s3_client, bucket: str, output_key: str) -> str:
    """
    If output_key exists, return IDS_Generated_2.docx, _3.docx, etc.
    """
    # Check if file exists
    try:
        s3_client.head_object(Bucket=bucket, Key=output_key)
    except:
        # Doesn't exist, use as-is
        return output_key
    
    # File exists - find next available number
    base, ext = os.path.splitext(output_key)
    
    # Check if already has a number suffix
    match = re.match(r'(.+)_(\d+)$', base)
    if match:
        base = match.group(1)
        start = int(match.group(2)) + 1
    else:
        start = 2
    
    for i in range(start, 1000):
        new_key = f"{base}_{i}{ext}"
        try:
            s3_client.head_object(Bucket=bucket, Key=new_key)
        except:
            return new_key
    
    # Fallback with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return f"{base}_{timestamp}{ext}"


def upload_to_s3(content: bytes, key: str) -> str:
    """Upload content to S3 and return URL."""
    try:
        s3_client.put_object(
            Bucket=S3_BUCKET,
            Key=key,
            Body=content,
            ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        return f"{S3_ENDPOINT}/{S3_BUCKET}/{key}"
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to upload to S3: {str(e)}")


def sanitize_text_content(text: str) -> str:
    """
    Clean up LLM-generated text content.
    - Replace 3+ newlines with 2 (single paragraph break)
    - Replace multiple spaces with single space
    - Strip leading/trailing whitespace
    """
    if not text:
        return text
    # Replace 3+ newlines with 2 (single paragraph break)
    text = re.sub(r'\n{3,}', '\n\n', text)
    # Replace multiple spaces with single space
    text = re.sub(r' {2,}', ' ', text)
    # Strip leading/trailing whitespace from each line
    lines = text.split('\n')
    lines = [line.strip() for line in lines]
    text = '\n'.join(lines)
    # Strip leading/trailing whitespace from entire text
    text = text.strip()
    return text


def apply_font_formatting(run, font_size: Pt, bold: bool = False):
    """Apply font formatting to a run."""
    run.font.name = FONT_NAME
    run.font.size = font_size
    run.font.bold = bold


def apply_default_font_formatting(run):
    """Apply default 11pt Times New Roman formatting to a run."""
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE_11PT


def create_paragraph_after(paragraph, doc=None):
    """
    Create a new paragraph element after the given paragraph.
    Returns the new paragraph object.
    """
    new_p = OxmlElement('w:p')
    paragraph._p.addnext(new_p)
    
    # We need to wrap this in a Paragraph object
    # Find it in the document
    from docx.text.paragraph import Paragraph
    new_para = Paragraph(new_p, paragraph._parent)
    return new_para


def set_paragraph_spacing(paragraph, space_after=0, line_spacing=240):
    """Set paragraph spacing (space_after in Pt, line_spacing in twips - 240 = single)."""
    pPr = paragraph._p.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:after'), str(int(space_after * 20)))  # Convert Pt to twips
    spacing.set(qn('w:line'), str(line_spacing))
    spacing.set(qn('w:lineRule'), 'auto')


def parse_sponsor_section(content: str) -> List[dict]:
    """
    Parse SPONSOR_SECTION content into structured list of paragraphs.
    
    Expected input format:
    Sponsorship – B+
    
    Company Name
    Company description paragraph...
    
    Person Name – Title
    Person bio paragraph...
    
    Returns list of dicts: [{"text": "...", "bold": True/False, "is_blank": True/False}]
    """
    if not content:
        return []
    
    content = sanitize_text_content(content)
    lines = content.split('\n')
    paragraphs = []
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Empty line = blank paragraph for spacing
        if not line:
            paragraphs.append({"text": "", "bold": False, "is_blank": True})
            i += 1
            continue
        
        # Determine if this line should be bold (header)
        is_header = False
        
        # Grade line: "Sponsorship – X" or "Sponsorship - X"
        if line.lower().startswith("sponsorship"):
            is_header = True
        # Person line: Contains " – " (em-dash) with title
        elif " – " in line:
            is_header = True
        # Person line: Contains " - " (hyphen) with title pattern like "Name - Title"
        elif " - " in line and len(line) < 80:
            # Check if it looks like "Name - Title" pattern
            parts = line.split(" - ", 1)
            if len(parts) == 2 and len(parts[0]) < 40 and len(parts[1]) < 40:
                is_header = True
        # Company name: Short line followed by longer description
        elif len(line) < 60 and not line.endswith('.'):
            # Look ahead for description
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                # If next line is long text (description), this is a header
                if next_line and len(next_line) > 80:
                    is_header = True
        
        paragraphs.append({"text": line, "bold": is_header, "is_blank": False})
        i += 1
    
    return paragraphs


def parse_risks_section(content: str) -> List[dict]:
    """
    Parse RISKS_SECTION content into structured list.
    
    Expected input format (tab-separated):
    Risk Name\tMitigant text paragraph...
    
    Risk Name 2\tMitigant text paragraph...
    
    Returns list of dicts with risk info and blank line indicators.
    """
    if not content:
        return []
    
    content = sanitize_text_content(content)
    
    # Split by double newline to get individual risks
    risk_blocks = re.split(r'\n\n+', content)
    
    paragraphs = []
    
    for idx, block in enumerate(risk_blocks):
        block = block.strip()
        if not block:
            continue
        
        # Check for tab separator
        if '\t' in block:
            parts = block.split('\t', 1)
            risk_name = parts[0].strip()
            mitigant = parts[1].strip() if len(parts) > 1 else ""
            paragraphs.append({
                "type": "risk",
                "risk_name": risk_name,
                "mitigant": mitigant
            })
        else:
            # No tab - might be legacy format or malformed
            # Try to detect "Risk Name    Description" pattern (multiple spaces)
            space_match = re.match(r'^([A-Z][^.]{5,40}?)\s{2,}(.+)$', block, re.DOTALL)
            if space_match:
                paragraphs.append({
                    "type": "risk",
                    "risk_name": space_match.group(1).strip(),
                    "mitigant": space_match.group(2).strip()
                })
            else:
                # Just treat as plain text
                paragraphs.append({
                    "type": "text",
                    "text": block
                })
        
        # Add blank line between risks (except after last)
        if idx < len(risk_blocks) - 1:
            paragraphs.append({"type": "blank"})
    
    return paragraphs


def insert_sponsor_paragraphs(paragraph, content: str):
    """
    Replace a paragraph with multiple paragraphs for SPONSOR_SECTION.
    The original paragraph is used for the first content, then new paragraphs are inserted after.
    """
    parsed = parse_sponsor_section(content)
    
    if not parsed:
        return
    
    # Clear the original paragraph
    paragraph.clear()
    
    # Use the original paragraph for the first non-blank item
    current_para = paragraph
    first_content = True
    
    for item in parsed:
        if first_content and not item.get("is_blank", False):
            # Use original paragraph for first content
            if item.get("text"):
                run = current_para.add_run(item["text"])
                apply_font_formatting(run, FONT_SIZE_11PT, bold=item.get("bold", False))
            set_paragraph_spacing(current_para, space_after=0, line_spacing=240)
            first_content = False
        else:
            # Create new paragraph after current
            new_para = create_paragraph_after(current_para)
            
            if item.get("is_blank", False):
                # Empty paragraph for spacing
                set_paragraph_spacing(new_para, space_after=0, line_spacing=240)
            else:
                run = new_para.add_run(item.get("text", ""))
                apply_font_formatting(run, FONT_SIZE_11PT, bold=item.get("bold", False))
                set_paragraph_spacing(new_para, space_after=0, line_spacing=240)
            
            current_para = new_para


def insert_risks_paragraphs(paragraph, content: str):
    """
    Replace a paragraph with multiple paragraphs for RISKS_SECTION.
    Each risk gets its own paragraph with bold risk name, tab, then mitigant text.
    """
    parsed = parse_risks_section(content)
    
    if not parsed:
        return
    
    # Clear the original paragraph
    paragraph.clear()
    
    current_para = paragraph
    first_content = True
    
    for item in parsed:
        if item.get("type") == "blank":
            # Create blank paragraph for spacing
            new_para = create_paragraph_after(current_para)
            set_paragraph_spacing(new_para, space_after=0, line_spacing=240)
            current_para = new_para
            continue
        
        if item.get("type") == "risk":
            if first_content:
                # Use original paragraph
                target_para = current_para
                first_content = False
            else:
                # Create new paragraph
                target_para = create_paragraph_after(current_para)
                current_para = target_para
            
            # Set hanging indent for risk format (matches template: left=2160, hanging=2160)
            pPr = target_para._p.get_or_add_pPr()
            ind = pPr.find(qn('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind')
                pPr.append(ind)
            ind.set(qn('w:left'), '2160')      # 1.5 inches in twips
            ind.set(qn('w:hanging'), '2160')   # Hanging indent
            
            # Set justified alignment
            jc = pPr.find(qn('w:jc'))
            if jc is None:
                jc = OxmlElement('w:jc')
                pPr.append(jc)
            jc.set(qn('w:val'), 'both')
            
            # Add bold risk name
            risk_run = target_para.add_run(item["risk_name"])
            apply_font_formatting(risk_run, FONT_SIZE_11PT, bold=True)
            
            # Add tab
            tab_run = target_para.add_run("\t")
            
            # Add mitigant text (not bold)
            mitigant_run = target_para.add_run(item["mitigant"])
            apply_font_formatting(mitigant_run, FONT_SIZE_11PT, bold=False)
            
            set_paragraph_spacing(target_para, space_after=0, line_spacing=240)
        
        elif item.get("type") == "text":
            # Plain text paragraph
            if first_content:
                target_para = current_para
                first_content = False
            else:
                target_para = create_paragraph_after(current_para)
                current_para = target_para
            
            run = target_para.add_run(item.get("text", ""))
            apply_font_formatting(run, FONT_SIZE_11PT, bold=False)
            set_paragraph_spacing(target_para, space_after=0, line_spacing=240)


def replace_placeholders_in_paragraph(paragraph, placeholders: Dict[str, str]) -> bool:
    """
    Replace {{PLACEHOLDER}} patterns across run boundaries.
    Handles Word's tendency to split text across multiple runs.
    Applies proper formatting based on placeholder type.
    """
    # Build full text and map character positions to runs
    full_text = ""
    char_to_run = []  # Maps each character index to (run_index, char_index_in_run)
    
    for run_idx, run in enumerate(paragraph.runs):
        for char_idx, char in enumerate(run.text):
            char_to_run.append((run_idx, char_idx))
        full_text += run.text
    
    if not full_text or "{{" not in full_text:
        return False
    
    # Find all placeholders
    modified = False
    pattern = re.compile(r'\{\{([A-Z0-9_]+)\}\}')
    
    # Track replacements to apply formatting after
    replacements_to_format = []
    
    # Process replacements from end to start (so positions stay valid)
    matches = list(pattern.finditer(full_text))
    for match in reversed(matches):
        placeholder_key = match.group(1)
        if placeholder_key in placeholders:
            replacement = str(placeholders[placeholder_key])
            
            # Special handling for SPONSOR_SECTION - insert multiple paragraphs
            if placeholder_key == "SPONSOR_SECTION":
                insert_sponsor_paragraphs(paragraph, replacement)
                modified = True
                continue
            
            # Special handling for RISKS_SECTION - insert multiple paragraphs with formatting
            if placeholder_key == "RISKS_SECTION":
                insert_risks_paragraphs(paragraph, replacement)
                modified = True
                continue
            
            # Sanitize text content for other placeholders
            replacement = sanitize_text_content(replacement)
            
            start_pos = match.start()
            end_pos = match.end()
            
            # Find which runs are affected
            if char_to_run:  # Check we have characters to map
                start_run_idx, start_char_idx = char_to_run[start_pos]
                end_run_idx, end_char_idx = char_to_run[end_pos - 1]
                
                if start_run_idx == end_run_idx:
                    # Placeholder is within a single run - simple case
                    run = paragraph.runs[start_run_idx]
                    run.text = run.text[:start_char_idx] + replacement + run.text[end_char_idx + 1:]
                    # Store for formatting
                    replacements_to_format.append((start_run_idx, placeholder_key, replacement))
                else:
                    # Placeholder spans multiple runs
                    # Put replacement in first run, clear the rest
                    first_run = paragraph.runs[start_run_idx]
                    last_run = paragraph.runs[end_run_idx]
                    
                    # Text before placeholder in first run + replacement
                    first_run.text = first_run.text[:start_char_idx] + replacement
                    
                    # Text after placeholder in last run
                    last_run.text = last_run.text[end_char_idx + 1:]
                    
                    # Clear runs in between
                    for run_idx in range(start_run_idx + 1, end_run_idx):
                        paragraph.runs[run_idx].text = ""
                    
                    # Store for formatting
                    replacements_to_format.append((start_run_idx, placeholder_key, replacement))
                
                modified = True
                
                # Rebuild the mapping for next iteration (since text changed)
                full_text = ""
                char_to_run = []
                for run_idx, run in enumerate(paragraph.runs):
                    for char_idx, char in enumerate(run.text):
                        char_to_run.append((run_idx, char_idx))
                    full_text += run.text
    
    # Apply formatting to replaced text
    if modified and replacements_to_format:
        # Apply default 11pt Times New Roman formatting to all runs with replacements
        for run_idx, placeholder_key, replacement_text in replacements_to_format:
            if run_idx < len(paragraph.runs):
                run = paragraph.runs[run_idx]
                apply_default_font_formatting(run)
    
    return modified


def replace_image_placeholders_in_paragraph(paragraph, images: Dict[str, str]) -> bool:
    """Replace image placeholders in a paragraph, handling split runs."""
    # Build full text to check for image placeholders
    full_text = ""
    char_to_run = []
    
    for run_idx, run in enumerate(paragraph.runs):
        for char_idx, char in enumerate(run.text):
            char_to_run.append((run_idx, char_idx))
        full_text += run.text
    
    if not full_text or "{{IMAGE_" not in full_text:
        return False
    
    # Find image placeholders
    pattern = re.compile(r'\{\{(IMAGE_[A-Z0-9_]+)\}\}')
    matches = list(pattern.finditer(full_text))
    
    for match in reversed(matches):
        placeholder_key = match.group(1)
        if placeholder_key in images:
            start_pos = match.start()
            end_pos = match.end()
            
            # Find which runs are affected
            if char_to_run:
                start_run_idx, start_char_idx = char_to_run[start_pos]
                end_run_idx, end_char_idx = char_to_run[end_pos - 1]
                
                # Clear the placeholder text
                if start_run_idx == end_run_idx:
                    run = paragraph.runs[start_run_idx]
                    run.text = run.text[:start_char_idx] + run.text[end_char_idx + 1:]
                else:
                    first_run = paragraph.runs[start_run_idx]
                    last_run = paragraph.runs[end_run_idx]
                    first_run.text = first_run.text[:start_char_idx]
                    last_run.text = last_run.text[end_char_idx + 1:]
                    for run_idx in range(start_run_idx + 1, end_run_idx):
                        paragraph.runs[run_idx].text = ""
                
                # Decode and insert image with proper dimensions
                try:
                    image_bytes = base64.b64decode(images[placeholder_key])
                    
                    # Calculate optimal dimensions that fit page constraints
                    preferred_width = IMAGE_WIDTHS.get(placeholder_key, 6.0)
                    width_inches, height_inches = calculate_image_dimensions(image_bytes, preferred_width)
                    
                    # Create fresh image stream for insertion
                    image_stream = BytesIO(image_bytes)
                    
                    # Add the image to the first affected run with calculated dimensions
                    paragraph.runs[start_run_idx].add_picture(
                        image_stream, 
                        width=Inches(width_inches),
                        height=Inches(height_inches)
                    )
                    
                    print(f"Inserted {placeholder_key}: {width_inches:.2f}\" x {height_inches:.2f}\"")
                    return True
                    
                except Exception as e:
                    raise HTTPException(status_code=400, detail=f"Failed to process image {placeholder_key}: {str(e)}")
    
    return False


def process_paragraphs(paragraphs, placeholders: Dict[str, str], images: Dict[str, str]):
    """Process paragraphs for text and image replacements."""
    # Convert to list to avoid issues with modifying during iteration
    para_list = list(paragraphs)
    for paragraph in para_list:
        # Try text replacement first
        replace_placeholders_in_paragraph(paragraph, placeholders)
        # Then try image replacement
        replace_image_placeholders_in_paragraph(paragraph, images)


def replace_placeholders_in_table(table, placeholders: Dict[str, str], images: Dict[str, str]) -> bool:
    """Replace placeholders in all cells of a table."""
    modified = False
    for row in table.rows:
        for cell in row.cells:
            process_paragraphs(cell.paragraphs, placeholders, images)
            modified = True
    return modified


def fill_template(template_bytes: bytes, placeholders: Dict[str, str], images: Dict[str, str]) -> bytes:
    """Fill the template with placeholders and images."""
    doc = Document(BytesIO(template_bytes))
    
    # Process main document paragraphs
    process_paragraphs(doc.paragraphs, placeholders, images)
    
    # Process tables
    for table in doc.tables:
        replace_placeholders_in_table(table, placeholders, images)
    
    # Process headers and footers for all section types
    for section in doc.sections:
        # Process different header types
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header:
                process_paragraphs(header.paragraphs, placeholders, images)
                for table in header.tables:
                    replace_placeholders_in_table(table, placeholders, images)
        
        # Process different footer types
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                process_paragraphs(footer.paragraphs, placeholders, images)
                for table in footer.tables:
                    replace_placeholders_in_table(table, placeholders, images)
    
    # Save to bytes
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


@app.post("/fill")
async def fill_template_endpoint(request: FillRequest):
    """Fill template and return as download."""
    # Download template
    template_bytes = download_template(request.template_key)
    
    # Fill template
    filled_bytes = fill_template(template_bytes, request.placeholders, request.images)
    
    # Return as download
    return StreamingResponse(
        BytesIO(filled_bytes),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={request.output_filename}"}
    )


@app.post("/fill-and-upload")
async def fill_and_upload_endpoint(request: FillAndUploadRequest):
    """Fill template and upload to S3."""
    # Download template
    template_bytes = download_template(request.template_key)
    
    # Fill template
    filled_bytes = fill_template(template_bytes, request.placeholders, request.images)
    
    # Get unique filename if original exists
    output_key = get_unique_output_key(s3_client, S3_BUCKET, request.output_key)
    
    # Upload to S3
    output_url = upload_to_s3(filled_bytes, output_key)
    
    return {
        "success": True,
        "output_key": output_key,  # Return actual filename used
        "output_url": output_url,
        "original_key": request.output_key  # Also return what was requested
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
