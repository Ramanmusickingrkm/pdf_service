from flask import Flask, request, send_file
from flask_cors import CORS
import mammoth
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import tempfile

app = Flask(__name__)
CORS(app)

def extract_fields_from_docx(docx_bytes):
    """Extract all fields from DOCX (placeholders like {{field}}, [field])"""
    doc = Document(io.BytesIO(docx_bytes))
    text = '\n'.join([para.text for para in doc.paragraphs])
    
    fields = []
    # Find all placeholders
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
    
    return fields

def replace_fields_in_docx(docx_bytes, field_values):
    """Replace placeholders in DOCX with actual values"""
    doc = Document(io.BytesIO(docx_bytes))
    
    for paragraph in doc.paragraphs:
        for field_name, field_value in field_values.items():
            patterns = [
                f'{{{{{field_name}}}}}',
                f'[{field_name}]',
                f'__{field_name}__',
                f'${field_name}$'
            ]
            for pattern in patterns:
                if pattern in paragraph.text:
                    paragraph.text = paragraph.text.replace(pattern, str(field_value))
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for field_name, field_value in field_values.items():
                        patterns = [
                            f'{{{{{field_name}}}}}',
                            f'[{field_name}]',
                            f'__{field_name}__',
                            f'${field_name}$'
                        ]
                        for pattern in patterns:
                            if pattern in paragraph.text:
                                paragraph.text = paragraph.text.replace(pattern, str(field_value))
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

def convert_docx_to_pdf(docx_bytes):
    """Convert DOCX to PDF using reportlab"""
    try:
        from pdf2docx import Converter
        import tempfile
        
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_docx:
            temp_docx.write(docx_bytes)
            temp_docx_path = temp_docx.name
        
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
            temp_pdf_path = temp_pdf.name
        
        cv = Converter(temp_docx_path)
        cv.convert(temp_pdf_path, start=0, end=None)
        cv.close()
        
        with open(temp_pdf_path, 'rb') as f:
            pdf_bytes = f.read()
        
        os.unlink(temp_docx_path)
        os.unlink(temp_pdf_path)
        
        return pdf_bytes
        
    except ImportError:
        # Fallback: Return DOCX as is
        return docx_bytes

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    """Parse DOCX and detect fields"""
    try:
        file = request.files['file']
        if not file:
            return {'success': False, 'error': 'No file uploaded'}, 400
        
        docx_bytes = file.read()
        fields = extract_fields_from_docx(docx_bytes)
        
        return {
            'success': True,
            'fields': [{'name': f, 'label': f.replace('_', ' ').title(), 'type': 'text'} for f in fields],
            'message': f'Found {len(fields)} fields'
        }
    
    except Exception as e:
        return {'success': False, 'error': str(e)}, 500

@app.route('/fill-and-pdf', methods=['POST'])
def fill_and_pdf():
    """Fill fields in DOCX and return PDF"""
    try:
        data = request.json
        docx_base64 = data.get('docxBase64')
        field_values = data.get('fieldValues', {})
        title = data.get('title', 'signed_document')
        
        if not docx_base64:
            return {'success': False, 'error': 'No document provided'}, 400
        
        import base64
        docx_bytes = base64.b64decode(docx_base64)
        
        # Fill fields in DOCX
        filled_docx = replace_fields_in_docx(docx_bytes, field_values)
        
        # Convert to PDF
        pdf_bytes = convert_docx_to_pdf(filled_docx.getvalue())
        
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'{title}_signed.pdf'
        )
    
    except Exception as e:
        return {'success': False, 'error': str(e)}, 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001)