from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docx import Document
import io
import re
import os
import base64
import tempfile

app = Flask(__name__)
CORS(app)

# PDF support flag
PDF_SUPPORT = False
try:
    from pdf2docx import Converter
    PDF_SUPPORT = True
    print("✅ pdf2docx loaded, PDF support enabled")
except ImportError:
    print("⚠️ pdf2docx not available, will return DOCX format")

def extract_fields_from_docx(docx_bytes):
    """Extract all fields from DOCX"""
    doc = Document(io.BytesIO(docx_bytes))
    text = '\n'.join([para.text for para in doc.paragraphs])
    
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
    """Replace placeholders with actual values"""
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
    """Convert DOCX to PDF using pdf2docx"""
    if not PDF_SUPPORT:
        return docx_bytes
    
    try:
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
        
    except Exception as e:
        print(f"PDF conversion error: {e}")
        return docx_bytes

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({
        'status': 'ok',
        'service': 'docx-pdf-service',
        'pdf_support': PDF_SUPPORT
    })

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        docx_bytes = file.read()
        fields = extract_fields_from_docx(docx_bytes)
        
        return jsonify({
            'success': True,
            'fields': [{'name': f, 'label': f.replace('_', ' ').title(), 'type': 'text'} for f in fields],
            'message': f'Found {len(fields)} fields'
        })
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/fill-and-pdf', methods=['POST'])
def fill_and_pdf():
    try:
        data = request.get_json()
        docx_base64 = data.get('docxBase64')
        field_values = data.get('fieldValues', {})
        title = data.get('title', 'signed_document')
        
        if not docx_base64:
            return jsonify({'success': False, 'error': 'No document provided'}), 400
        
        docx_bytes = base64.b64decode(docx_base64)
        filled_docx = replace_fields_in_docx(docx_bytes, field_values)
        filled_bytes = filled_docx.getvalue()
        
        output_bytes = convert_docx_to_pdf(filled_bytes)
        
        if output_bytes != filled_bytes:
            content_type = 'application/pdf'
            extension = 'pdf'
        else:
            content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            extension = 'docx'
        
        return send_file(
            io.BytesIO(output_bytes),
            mimetype=content_type,
            as_attachment=True,
            download_name=f'{title.replace("/", "_")}_signed.{extension}'
        )
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/fill-and-return-docx', methods=['POST'])
def fill_and_return_docx():
    try:
        data = request.get_json()
        docx_base64 = data.get('docxBase64')
        field_values = data.get('fieldValues', {})
        title = data.get('title', 'signed_document')
        
        if not docx_base64:
            return jsonify({'success': False, 'error': 'No document provided'}), 400
        
        docx_bytes = base64.b64decode(docx_base64)
        filled_docx = replace_fields_in_docx(docx_bytes, field_values)
        
        return send_file(
            io.BytesIO(filled_docx.getvalue()),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f'{title.replace("/", "_")}_filled.docx'
        )
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
