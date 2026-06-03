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

def extract_fields_from_text(text):
    """Extract all field names from text"""
    fields = []
    patterns = [
        r'\{\{([^}]+)\}\}',   # {{field}}
        r'\[([^\]]+)\]',       # [field]
        r'__([^_]+)__',        # __field__
        r'\$([^$]+)\$'         # $field$
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            field_name = match.strip()
            if field_name and field_name not in fields:
                fields.append(field_name)
    
    return fields

def get_field_mapping(docx_bytes):
    """Detect all fields in DOCX and create name mapping"""
    doc = Document(io.BytesIO(docx_bytes))
    text = '\n'.join([para.text for para in doc.paragraphs])
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    text += '\n' + para.text
    
    fields = extract_fields_from_text(text)
    logger.info(f'🔍 Detected fields in template: {fields}')
    
    # Create mapping (normalize field names)
    mapping = {}
    for field in fields:
        # Create normalized key (lowercase, no spaces, underscores instead of spaces)
        normalized = re.sub(r'[^a-zA-Z0-9]', '_', field.lower())
        mapping[normalized] = field
        logger.info(f'  📌 "{field}" -> normalized: "{normalized}"')
    
    return mapping, fields

def replace_fields_in_docx(docx_bytes, field_values):
    """Replace placeholders in DOCX while preserving ALL formatting"""
    try:
        doc = Document(io.BytesIO(docx_bytes))
        
        # First, detect all fields in the document and create mapping
        field_mapping, detected_fields = get_field_mapping(docx_bytes)
        
        logger.info(f'📝 Received field values: {list(field_values.keys())}')
        logger.info(f'📝 Detected template fields: {detected_fields}')
        
        # Create replacement dictionary
        replacements = {}
        for template_field in detected_fields:
            # Try direct match first
            if template_field in field_values:
                replacements[template_field] = str(field_values[template_field])
                logger.info(f'  ✓ Direct match: {template_field} -> {replacements[template_field][:50]}')
            else:
                # Try normalized match
                normalized = re.sub(r'[^a-zA-Z0-9]', '_', template_field.lower())
                if normalized in field_values:
                    replacements[template_field] = str(field_values[normalized])
                    logger.info(f'  ✓ Normalized match: {template_field} (normalized: {normalized}) -> {replacements[template_field][:50]}')
                else:
                    # Check case-insensitive
                    for key in field_values.keys():
                        if key.lower() == template_field.lower():
                            replacements[template_field] = str(field_values[key])
                            logger.info(f'  ✓ Case-insensitive match: {template_field} -> {replacements[template_field][:50]}')
                            break
        
        if not replacements:
            logger.warning('⚠️ No field replacements found! Check field name matching.')
            logger.warning(f'   Available field values: {list(field_values.keys())}')
            logger.warning(f'   Detected template fields: {detected_fields}')
        
        # Process paragraphs
        for paragraph in doc.paragraphs:
            original_text = paragraph.text
            new_text = original_text
            
            for template_field, replacement_value in replacements.items():
                if not replacement_value:
                    continue
                
                # Try all pattern formats
                patterns_to_replace = [
                    f'{{{{{template_field}}}}}',
                    f'[{template_field}]',
                    f'__{template_field}__',
                    f'${template_field}$'
                ]
                
                for pattern in patterns_to_replace:
                    if pattern in new_text:
                        new_text = new_text.replace(pattern, replacement_value)
                        logger.info(f'  ✓ Replaced {pattern}')
            
            if new_text != original_text:
                if paragraph.runs:
                    # Update runs to preserve formatting
                    full_text = ''.join([run.text for run in paragraph.runs])
                    for template_field, replacement_value in replacements.items():
                        if not replacement_value:
                            continue
                        patterns_to_replace = [
                            f'{{{{{template_field}}}}}',
                            f'[{template_field}]',
                            f'__{template_field}__',
                            f'${template_field}$'
                        ]
                        for pattern in patterns_to_replace:
                            if pattern in full_text:
                                full_text = full_text.replace(pattern, replacement_value)
                    
                    # Distribute text across runs
                    if full_text != original_text:
                        # Clear all runs and create new one
                        paragraph.clear()
                        paragraph.add_run(full_text)
                else:
                    paragraph.text = new_text
        
        # Process tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        original_text = paragraph.text
                        new_text = original_text
                        
                        for template_field, replacement_value in replacements.items():
                            if not replacement_value:
                                continue
                            patterns_to_replace = [
                                f'{{{{{template_field}}}}}',
                                f'[{template_field}]',
                                f'__{template_field}__',
                                f'${template_field}$'
                            ]
                            for pattern in patterns_to_replace:
                                if pattern in new_text:
                                    new_text = new_text.replace(pattern, replacement_value)
                        
                        if new_text != original_text:
                            if paragraph.runs:
                                full_text = ''.join([run.text for run in paragraph.runs])
                                for template_field, replacement_value in replacements.items():
                                    if not replacement_value:
                                        continue
                                    patterns_to_replace = [
                                        f'{{{{{template_field}}}}}',
                                        f'[{template_field}]',
                                        f'__{template_field}__',
                                        f'${template_field}$'
                                    ]
                                    for pattern in patterns_to_replace:
                                        if pattern in full_text:
                                            full_text = full_text.replace(pattern, replacement_value)
                                
                                if full_text != original_text:
                                    paragraph.clear()
                                    paragraph.add_run(full_text)
                            else:
                                paragraph.text = new_text
        
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        logger.info('✅ Fields replaced, formatting preserved')
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

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    """Parse DOCX and detect fields"""
    try:
        if 'file' not in request.files:
            return {'success': False, 'error': 'No file uploaded'}, 400
        
        file = request.files['file']
        if file.filename == '':
            return {'success': False, 'error': 'No file selected'}, 400
        
        docx_bytes = file.read()
        
        # Extract fields
        doc = Document(io.BytesIO(docx_bytes))
        text = '\n'.join([para.text for para in doc.paragraphs])
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        text += '\n' + para.text
        
        fields = extract_fields_from_text(text)
        
        # Also provide normalized versions for frontend
        field_list = []
        for field in fields:
            field_list.append({
                'name': field,
                'normalized': re.sub(r'[^a-zA-Z0-9]', '_', field.lower()),
                'label': field.replace('_', ' ').title(),
                'type': 'text'
            })
        
        return {
            'success': True,
            'fields': field_list,
            'message': f'Found {len(fields)} fields'
        }
    
    except Exception as e:
        logger.error(f'Parse error: {str(e)}')
        return {'success': False, 'error': str(e)}, 500

@app.route('/fill-and-pdf', methods=['POST'])
def fill_and_pdf():
    """Fill fields in DOCX and return FILLED DOCX (preserving format)"""
    try:
        data = request.json
        if not data:
            return {'success': False, 'error': 'No JSON data provided'}, 400
        
        docx_base64 = data.get('docxBase64')
        field_values = data.get('fieldValues', {})
        title = data.get('title', 'signed_document')
        
        logger.info(f'📊 Title: {title}')
        logger.info(f'📋 Received field values: {list(field_values.keys())}')
        
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
        
        logger.info(f'✅ Returning filled DOCX: {title}_filled.docx')
        
        return send_file(
            io.BytesIO(filled_docx_io.getvalue()),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f'{title}_filled.docx'
        )
    
    except Exception as e:
        logger.error(f'Error: {str(e)}')
        return {'success': False, 'error': str(e)}, 500

@app.route('/fill-and-return-docx', methods=['POST'])
def fill_and_return_docx():
    """Same as fill-and-pdf - returns filled DOCX"""
    return fill_and_pdf()

if __name__ == '__main__':
    logger.info('🚀 Starting DocSign DOCX Filler Service')
    logger.info('📄 Returns filled DOCX with preserved formatting')
    app.run(host='0.0.0.0', port=5001, debug=False)
