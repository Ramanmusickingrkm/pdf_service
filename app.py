from flask import Flask, request, send_file
from flask_cors import CORS
from docx import Document
import io
import re
import base64
import subprocess
import tempfile
import os

app = Flask(__name__)
CORS(app)

def extract_fields_from_docx(docx_bytes):
    """Extract all fields from DOCX"""
    doc = Document(io.BytesIO(docx_bytes))
    text = '\n'.join([para.text for para in doc.paragraphs])
    
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
    
    return fields

def replace_fields_in_docx(docx_bytes, field_values):
    """Replace placeholders in DOCX"""
    doc = Document(io.BytesIO(docx_bytes))
    
    def replace_in_paragraph(paragraph):
        for field_name, field_value in field_values.items():
            if not field_value:
                continue
            patterns = [
                f'{{{{{field_name}}}}}',
                f'[{field_name}]',
                f'__{field_name}__',
                f'${field_name}$'
            ]
            for pattern in patterns:
                if pattern in paragraph.text:
                    paragraph.text = paragraph.text.replace(pattern, str(field_value))
    
    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph)
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

def convert_docx_to_pdf_libreoffice(docx_bytes):
    """Convert DOCX to PDF using LibreOffice (headless)"""
    temp_docx_path = None
    temp_pdf_path = None
    
    try:
        # Create temporary files
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
            f.write(docx_bytes)
            temp_docx_path = f.name
        
        output_dir = tempfile.mkdtemp()
        
        # Convert using LibreOffice
        result = subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'pdf',
            '--outdir', output_dir, temp_docx_path
        ], capture_output=True, text=True, timeout=60)
        
        if result.returncode != 0:
            print(f"LibreOffice error: {result.stderr}")
            return None
        
        # Expected PDF path
        pdf_name = os.path.basename(temp_docx_path).replace('.docx', '.pdf')
        temp_pdf_path = os.path.join(output_dir, pdf_name)
        
        if os.path.exists(temp_pdf_path):
            with open(temp_pdf_path, 'rb') as f:
                pdf_bytes = f.read()
            return pdf_bytes
        else:
            print("PDF file not created")
            return None
            
    except subprocess.TimeoutExpired:
        print("LibreOffice conversion timeout")
        return None
    except Exception as e:
        print(f"Conversion error: {e}")
        return None
    finally:
        # Cleanup
        if temp_docx_path and os.path.exists(temp_docx_path):
            os.unlink(temp_docx_path)
        if temp_pdf_path and os.path.exists(temp_pdf_path):
            os.unlink(temp_pdf_path)

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
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
    """Fill fields and return PDF"""
    try:
        data = request.json
        docx_base64 = data.get('docxBase64')
        field_values = data.get('fieldValues', {})
        title = data.get('title', 'signed_document')
        
        if not docx_base64:
            return {'success': False, 'error': 'No document provided'}, 400
        
        docx_bytes = base64.b64decode(docx_base64)
        filled_docx = replace_fields_in_docx(docx_bytes, field_values)
        
        # Convert to PDF
        pdf_bytes = convert_docx_to_pdf_libreoffice(filled_docx.getvalue())
        
        if pdf_bytes:
            return send_file(
                io.BytesIO(pdf_bytes),
                mimetype='application/pdf',
                as_attachment=True,
                download_name=f'{title}_signed.pdf'
            )
        else:
            # Fallback: Return DOCX
            return send_file(
                io.BytesIO(filled_docx.getvalue()),
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                as_attachment=True,
                download_name=f'{title}_signed.docx'
            )
    
    except Exception as e:
        print(f"Error: {str(e)}")
        return {'success': False, 'error': str(e)}, 500

@app.route('/fill-and-return-docx', methods=['POST'])
def fill_and_return_docx():
    """Return DOCX only (no PDF conversion)"""
    try:
        data = request.json
        docx_base64 = data.get('docxBase64')
        field_values = data.get('fieldValues', {})
        title = data.get('title', 'signed_document')
        
        if not docx_base64:
            return {'success': False, 'error': 'No document provided'}, 400
        
        docx_bytes = base64.b64decode(docx_base64)
        filled_docx = replace_fields_in_docx(docx_bytes, field_values)
        
        return send_file(
            io.BytesIO(filled_docx.getvalue()),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f'{title}_filled.docx'
        )
    
    except Exception as e:
        return {'success': False, 'error': str(e)}, 500

@app.route('/health', methods=['GET'])
def health():
    return {'status': 'ok', 'service': 'docx-pdf-service'}

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001)
