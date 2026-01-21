import os
import re
import base64
from io import BytesIO
from typing import Dict, Optional, Tuple
from datetime import datetime

from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
import boto3
from botocore.config import Config
from docx import Document
from docx.shared import Inches, Pt
from PIL import Image

app = FastAPI(title="IDS Template Filler", version="1.0.0")

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
    "IMAGE_AERIAL_MAP": 5.0,
    "IMAGE_LOCATION_MAP": 5.0,
    "IMAGE_REGIONAL_MAP": 5.0,
    "IMAGE_SITE_PLAN": 5.5,
    "IMAGE_PILOT_SCHEDULE": 6.0,
    "IMAGE_TAKEOUT_SIZING": 6.0,
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
    return {"status": "ok"}

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
    - Strip leading/trailing whitespace
    """
    if not text:
        return text
    # Replace 3+ newlines with 2 (single paragraph break)
    text = re.sub(r'\n{3,}', '\n\n', text)
    # Strip leading/trailing whitespace
    text = text.strip()
    return text

def apply_font_formatting(run, font_size: Pt, bold: bool = False):
    """Apply font formatting to a run (used for SPONSOR_SECTION special formatting)."""
    run.font.name = FONT_NAME
    run.font.size = font_size
    run.font.bold = bold

def apply_default_font_formatting(run):
    """Apply default 11pt Times New Roman formatting to a run."""
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE_11PT

def format_sponsor_section_content(content: str, paragraph):
    """
    Format SPONSOR_SECTION content with special rules:
    - First line (Sponsorship – [Grade]): 11pt Bold
    - Company/sponsor name lines: 11pt Bold
    - Bio/description paragraphs: 11pt Regular
    """
    if not content:
        return
    
    # Sanitize whitespace first
    content = sanitize_text_content(content)
    
    # Split into lines
    lines = content.split('\n')
    
    # Clear existing runs
    paragraph.clear()
    
    if not lines:
        return
    
    # Process first line (should be "Sponsorship – [Grade]")
    first_run = paragraph.add_run(lines[0])
    apply_font_formatting(first_run, FONT_SIZE_11PT, bold=True)
    
    # Process remaining lines
    i = 1
    while i < len(lines):
        line = lines[i].strip()
        
        if not line:
            # Empty line - add line break for paragraph separation
            if i < len(lines) - 1:  # Don't add break after last line
                paragraph.add_run().add_break()
            i += 1
            continue
        
        # Heuristic to detect company/sponsor name headers:
        # - Line is relatively short (< 80 chars)
        # - Next non-empty line is significantly longer (at least 1.5x)
        # - OR line looks like a name (contains common name patterns)
        is_header = False
        
        # Check if it looks like a name/company (contains capital letters, no sentence-ending punctuation)
        looks_like_name = (
            len(line) < 80 and
            not line.endswith('.') and
            not line.endswith('!') and
            not line.endswith('?') and
            any(c.isupper() for c in line)
        )
        
        if looks_like_name:
            # Look ahead to next non-empty line to confirm
            for j in range(i + 1, len(lines)):
                next_line = lines[j].strip()
                if next_line:
                    # If next line is much longer, this is likely a header
                    if len(next_line) > len(line) * 1.5:
                        is_header = True
                    # If next line starts with lowercase or is a sentence, this is a header
                    elif next_line and next_line[0].islower():
                        is_header = True
                    break
        
        # Add the line
        run = paragraph.add_run(line)
        apply_font_formatting(run, FONT_SIZE_11PT, bold=is_header)
        
        # Add line break if not last line
        if i < len(lines) - 1:
            paragraph.add_run().add_break()
        
        i += 1
    
    # Set paragraph spacing
    ppr = paragraph.paragraph_format
    ppr.line_spacing_rule = 1  # Single spacing
    ppr.space_after = Pt(0)  # No extra space after

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
            
            # Special handling for SPONSOR_SECTION
            if placeholder_key == "SPONSOR_SECTION":
                # Sanitize and format the entire paragraph with special rules
                format_sponsor_section_content(replacement, paragraph)
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
    
    # Set paragraph spacing properties for consistent formatting
    if modified:
        ppr = paragraph.paragraph_format
        # Set line spacing: single spacing (lineRule="auto")
        # This ensures consistent paragraph spacing
        ppr.line_spacing_rule = 1  # Single spacing
        ppr.space_after = Pt(0)  # No extra space after
    
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
    for paragraph in paragraphs:
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