from flask import Flask, request, send_file
from flask_cors import CORS
from docx import Document
import io
import re
import base64
import logging
from datetime import datetime

app = Flask(__name__)
CORS(app)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def replace_fields_in_docx(docx_bytes, field_values):
    """Replace placeholders in DOCX while preserving formatting"""
    try:
        doc = Document(io.BytesIO(docx_bytes))
        logger.info(f'📝 Replacing fields in DOCX: {list(field_values.keys())}')
        
        # Process all paragraphs
        for paragraph in doc.paragraphs:
            original_text = paragraph.text
            if not original_text:
                continue
                
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
                            logger.info(f'  ✓ Replaced {pattern} -> {field_str[:30]}')
            
            if new_text != original_text:
                paragraph.text = new_text
        
        # Process tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        original_text = paragraph.text
                        if not original_text:
                            continue
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
        import traceback
        traceback.print_exc()
        return None

@app.route('/health', methods=['GET'])
def health():
    return {
        'status': 'ok',
        'service': 'DocSign DOCX Filler Service',
        'timestamp': datetime.now().isoformat()
    }

@app.route('/fill-and-pdf', methods=['POST'])
def fill_and_pdf():
    """Fill fields in DOCX and return filled DOCX"""
    try:
        data = request.json
        if not data:
            return {'success': False, 'error': 'No JSON data provided'}, 400
        
        docx_base64 = data.get('docxBase64')
        field_values = data.get('fieldValues', {})
        title = data.get('title', 'signed_document')
        
        logger.info(f'📊 Title: {title}')
        logger.info(f'📋 Field values received: {list(field_values.keys())}')
        
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
        
        logger.info(f'✅ Returning filled DOCX: {len(filled_docx_io.getvalue())} bytes')
        
        return send_file(
            io.BytesIO(filled_docx_io.getvalue()),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f'{title}_filled.docx'
        )
    
    except Exception as e:
        logger.error(f'❌ Error: {str(e)}')
        import traceback
        traceback.print_exc()
        return {'success': False, 'error': str(e)}, 500

@app.route('/fill-and-return-docx', methods=['POST'])
def fill_and_return_docx():
    return fill_and_pdf()

if __name__ == '__main__':
    logger.info('🚀 Starting DocSign DOCX Filler Service')
    app.run(host='0.0.0.0', port=5001, debug=False)
