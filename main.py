import os
import re
import base64
from io import BytesIO
from typing import Dict, Optional

from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
import boto3
from botocore.config import Config
from docx import Document
from docx.shared import Inches

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

# Image width configuration
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

def replace_placeholders_in_paragraph(paragraph, placeholders: Dict[str, str]) -> bool:
    """
    Replace {{PLACEHOLDER}} patterns across run boundaries.
    Handles Word's tendency to split text across multiple runs.
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
    
    # Process replacements from end to start (so positions stay valid)
    matches = list(pattern.finditer(full_text))
    for match in reversed(matches):
        placeholder_key = match.group(1)
        if placeholder_key in placeholders:
            replacement = str(placeholders[placeholder_key])
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
                
                modified = True
                
                # Rebuild the mapping for next iteration (since text changed)
                full_text = ""
                char_to_run = []
                for run_idx, run in enumerate(paragraph.runs):
                    for char_idx, char in enumerate(run.text):
                        char_to_run.append((run_idx, char_idx))
                    full_text += run.text
    
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
                
                # Decode and insert image
                try:
                    image_bytes = base64.b64decode(images[placeholder_key])
                    image_stream = BytesIO(image_bytes)
                    
                    # Add the image to the first affected run
                    width = IMAGE_WIDTHS.get(placeholder_key, 6.0)
                    paragraph.runs[start_run_idx].add_picture(image_stream, width=Inches(width))
                    
                    return True
                except Exception as e:
                    raise HTTPException(status_code=400, detail=f"Failed to decode image {placeholder_key}: {str(e)}")
    
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
    
    # Upload to S3
    output_url = upload_to_s3(filled_bytes, request.output_key)
    
    return {
        "success": True,
        "output_key": request.output_key,
        "output_url": output_url
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)