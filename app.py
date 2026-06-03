from flask import Flask, request, send_file
from flask_cors import CORS
from docx import Document
import io
import re
import base64
import logging
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import simpleSplit
import tempfile
import os

app = Flask(__name__)
CORS(app)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def extract_fields_from_docx(docx_bytes):
    """Extract all fields from DOCX (placeholders like {{field}}, [field])"""
    try:
        doc = Document(io.BytesIO(docx_bytes))
        text = '\n'.join([para.text for para in doc.paragraphs])
        
        # Also check tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        text += '\n' + para.text
        
        fields = []
        patterns = [
            r'\{\{([^}]+)\}\}',
            r'\[([^\]]+)\]',
            r'__([^_]+)__',
            r'\$([^$]+)\$'
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                field_name = match.strip()
                if field_name and field_name not in fields:
                    fields.append(field_name)
        
        logger.info(f'✅ Extracted {len(fields)} fields from DOCX')
        return fields
        
    except Exception as e:
        logger.error(f'❌ Error extracting fields: {str(e)}')
        return []

def replace_fields_in_docx(docx_bytes, field_values):
    """Replace placeholders in DOCX with actual values"""
    try:
        doc = Document(io.BytesIO(docx_bytes))
        
        logger.info(f'📝 Replacing {len(field_values)} field values')
        
        # Process paragraphs
        for paragraph in doc.paragraphs:
            original_text = paragraph.text
            new_text = original_text
            
            for field_name, field_value in field_values.items():
                if field_name and field_value:
                    field_str = str(field_value)
                    patterns = [
                        f'{{{{{field_name}}}}}',
                        f'[{field_name}]',
                        f'__{field_name}__',
                        f'${field_name}$'
                    ]
                    for pattern in patterns:
                        if pattern in new_text:
                            new_text = new_text.replace(pattern, field_str)
                            logger.info(f'  ✓ Replaced {pattern}')
            
            if new_text != original_text:
                paragraph.text = new_text
        
        # Process tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        original_text = paragraph.text
                        new_text = original_text
                        
                        for field_name, field_value in field_values.items():
                            if field_name and field_value:
                                field_str = str(field_value)
                                patterns = [
                                    f'{{{{{field_name}}}}}',
                                    f'[{field_name}]',
                                    f'__{field_name}__',
                                    f'${field_name}$'
                                ]
                                for pattern in patterns:
                                    if pattern in new_text:
                                        new_text = new_text.replace(pattern, field_str)
                        
                        if new_text != original_text:
                            paragraph.text = new_text
        
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        logger.info('✅ Fields replaced successfully')
        return output
        
    except Exception as e:
        logger.error(f'❌ Error replacing fields: {str(e)}')
        return None

def convert_text_to_pdf(text_content, title):
    """Convert text content to PDF using reportlab"""
    try:
        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4
        
        # Add title
        pdf.setFont("Helvetica-Bold", 16)
        pdf.drawCentredString(width/2, height - 50, title)
        
        # Add content
        pdf.setFont("Helvetica", 10)
        y = height - 80
        line_height = 14
        
        # Split text into lines
        lines = text_content.split('\n')
        for line in lines:
            if y < 50:
                pdf.showPage()
                y = height - 50
                pdf.setFont("Helvetica", 10)
            
            # Wrap long lines
            wrapped_lines = simpleSplit(line, pdf._fontname, pdf._fontsize, width - 100)
            for wrapped_line in wrapped_lines:
                if y < 50:
                    pdf.showPage()
                    y = height - 50
                    pdf.setFont("Helvetica", 10)
                pdf.drawString(50, y, wrapped_line)
                y -= line_height
        
        # Add signature footer
        y -= 20
        pdf.setFont("Helvetica-Oblique", 8)
        pdf.drawString(50, y, f"Electronically signed document - Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        pdf.save()
        buffer.seek(0)
        return buffer.getvalue()
        
    except Exception as e:
        logger.error(f'❌ PDF conversion error: {str(e)}')
        return None

@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return {
        'status': 'ok',
        'service': 'DocSign PDF Service',
        'timestamp': datetime.now().isoformat()
    }

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    """Parse DOCX and detect fields"""
    try:
        logger.info('📄 Parse DOCX request received')
        
        if 'file' not in request.files:
            return {'success': False, 'error': 'No file uploaded'}, 400
        
        file = request.files['file']
        if file.filename == '':
            return {'success': False, 'error': 'No file selected'}, 400
        
        docx_bytes = file.read()
        logger.info(f'📦 File size: {len(docx_bytes)} bytes')
        
        fields = extract_fields_from_docx(docx_bytes)
        
        return {
            'success': True,
            'fields': [{'name': f, 'label': f.replace('_', ' ').title(), 'type': 'text'} for f in fields],
            'message': f'Found {len(fields)} fields'
        }
    
    except Exception as e:
        logger.error(f'❌ Parse error: {str(e)}')
        return {'success': False, 'error': str(e)}, 500

@app.route('/fill-and-pdf', methods=['POST'])
def fill_and_pdf():
    """Fill fields in DOCX and return PDF"""
    try:
        logger.info('🔵 Fill and PDF request received')
        
        data = request.json
        if not data:
            return {'success': False, 'error': 'No JSON data provided'}, 400
        
        docx_base64 = data.get('docxBase64')
        field_values = data.get('fieldValues', {})
        title = data.get('title', 'signed_document')
        
        logger.info(f'📊 Title: {title}')
        logger.info(f'📋 Fields: {list(field_values.keys())}')
        
        if not docx_base64:
            return {'success': False, 'error': 'No document provided'}, 400
        
        # Decode base64
        try:
            docx_bytes = base64.b64decode(docx_base64)
            logger.info(f'✅ Decoded DOCX: {len(docx_bytes)} bytes')
        except Exception as e:
            return {'success': False, 'error': f'Failed to decode DOCX: {str(e)}'}, 400
        
        # Fill fields in DOCX
        filled_docx_io = replace_fields_in_docx(docx_bytes, field_values)
        
        if not filled_docx_io:
            return {'success': False, 'error': 'Failed to replace fields'}, 500
        
        # Extract text from filled DOCX
        filled_doc = Document(io.BytesIO(filled_docx_io.getvalue()))
        text_content = '\n'.join([para.text for para in filled_doc.paragraphs])
        
        # Convert to PDF
        pdf_bytes = convert_text_to_pdf(text_content, title)
        
        if not pdf_bytes:
            logger.warning('⚠️ PDF conversion failed, returning DOCX as fallback')
            return send_file(
                io.BytesIO(filled_docx_io.getvalue()),
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                as_attachment=True,
                download_name=f'{title}_filled.docx'
            )
        
        logger.info(f'📤 Returning PDF: {title}_signed.pdf ({len(pdf_bytes)} bytes)')
        
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'{title}_signed.pdf'
        )
    
    except Exception as e:
        logger.error(f'❌ Error: {str(e)}')
        return {'success': False, 'error': str(e)}, 500

@app.route('/fill-and-return-docx', methods=['POST'])
def fill_and_return_docx():
    """Fill fields and return DOCX only"""
    try:
        logger.info('🔵 Fill and return DOCX request received')
        
        data = request.json
        if not data:
            return {'success': False, 'error': 'No JSON data provided'}, 400
        
        docx_base64 = data.get('docxBase64')
        field_values = data.get('fieldValues', {})
        title = data.get('title', 'signed_document')
        
        if not docx_base64:
            return {'success': False, 'error': 'No document provided'}, 400
        
        docx_bytes = base64.b64decode(docx_base64)
        filled_docx_io = replace_fields_in_docx(docx_bytes, field_values)
        
        if not filled_docx_io:
            return {'success': False, 'error': 'Failed to replace fields'}, 500
        
        return send_file(
            io.BytesIO(filled_docx_io.getvalue()),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f'{title}_filled.docx'
        )
    
    except Exception as e:
        logger.error(f'❌ Error: {str(e)}')
        return {'success': False, 'error': str(e)}, 500

if __name__ == '__main__':
    logger.info('🚀 Starting DocSign PDF Service')
    app.run(host='0.0.0.0', port=5001, debug=False)
