from flask import Flask, request, send_file
from flask_cors import CORS
import mammoth
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import re
import os
import base64
import logging
from datetime import datetime

# PDF conversion imports
try:
    from pdf2docx import Converter
    HAS_PDF2DOCX = True
except ImportError:
    HAS_PDF2DOCX = False

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

app = Flask(__name__)
CORS(app)
logging.basicConfig(level=logging.INFO)

def extract_fields_from_docx(docx_bytes):
    """Extract all fields from DOCX (placeholders like {{field}}, [field])"""
    try:
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
        
        app.logger.info(f'✅ Extracted {len(fields)} fields from DOCX')
        return fields
        
    except Exception as e:
        app.logger.error(f'❌ Error extracting fields: {str(e)}')
        return []

def replace_fields_in_docx(docx_bytes, field_values):
    """Replace placeholders in DOCX with actual values - Proper method"""
    try:
        doc = Document(io.BytesIO(docx_bytes))
        
        app.logger.info(f'📝 Replacing {len(field_values)} field values in DOCX')
        
        # Replace in paragraphs
        for paragraph in doc.paragraphs:
            for field_name, field_value in field_values.items():
                if field_name and field_value:
                    field_str = str(field_value)
                    
                    # Check multiple placeholder patterns
                    patterns = [
                        (f'{{{{{field_name}}}}}', field_str),  # {{field}}
                        (f'[{field_name}]', field_str),         # [field]
                        (f'__{field_name}__', field_str),       # __field__
                        (f'${field_name}$', field_str),         # $field$
                    ]
                    
                    for pattern, replacement in patterns:
                        if pattern in paragraph.text:
                            app.logger.info(f'  ✓ Replaced {pattern} with {replacement[:50]}')
                            # Replace in runs to preserve formatting
                            if pattern in paragraph.text:
                                paragraph.text = paragraph.text.replace(pattern, replacement)
        
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
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
                                    if pattern in paragraph.text:
                                        paragraph.text = paragraph.text.replace(pattern, field_str)
        
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        app.logger.info('✅ Fields replaced successfully')
        return output
        
    except Exception as e:
        app.logger.error(f'❌ Error replacing fields: {str(e)}')
        return None

def convert_docx_to_pdf_advanced(docx_bytes, title='signed_document'):
    """
    Convert DOCX to PDF using python-docx with better formatting
    Falls back to reportlab if needed
    """
    try:
        app.logger.info('🔄 Converting DOCX to PDF...')
        
        # Try using python-docx + reportlab for conversion
        doc = Document(io.BytesIO(docx_bytes))
        
        # Extract content
        pdf_elements = []
        styles = getSampleStyleSheet()
        
        # Add title
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            textColor=colors.HexColor('#1e2d4a'),
            spaceAfter=12,
            alignment=1  # CENTER
        )
        
        # Process paragraphs
        for para in doc.paragraphs:
            if para.text.strip():
                # Determine style
                if para.style.name.startswith('Heading'):
                    style = styles['Heading2']
                else:
                    style = styles['Normal']
                
                para_obj = Paragraph(para.text, style)
                pdf_elements.append(para_obj)
                pdf_elements.append(Spacer(1, 6))
        
        # Process tables
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_data = [cell.text for cell in row.cells]
                table_data.append(row_data)
            
            if table_data:
                table_obj = Table(table_data)
                table_obj.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e2d4a')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('GRID', (0, 0), (-1, -1), 1, colors.grey),
                ]))
                pdf_elements.append(table_obj)
                pdf_elements.append(Spacer(1, 12))
        
        # Add signature info
        pdf_elements.append(Spacer(1, 20))
        sig_style = ParagraphStyle(
            'SignatureInfo',
            parent=styles['Normal'],
            fontSize=9,
            textColor=colors.HexColor('#666666')
        )
        sig_text = f'<b>Document:</b> {title}<br/><b>Generated:</b> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}<br/><b>Status:</b> Electronically Signed'
        pdf_elements.append(Paragraph(sig_text, sig_style))
        
        # Create PDF
        output = io.BytesIO()
        pdf_doc = SimpleDocTemplate(output, pagesize=A4, topMargin=20, bottomMargin=20)
        pdf_doc.build(pdf_elements)
        
        output.seek(0)
        app.logger.info(f'✅ PDF created successfully, size: {len(output.getvalue())} bytes')
        return output.getvalue()
        
    except Exception as e:
        app.logger.error(f'❌ PDF conversion error: {str(e)}')
        return None

@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return {
        'status': 'ok',
        'service': 'DocSign PDF Service',
        'capabilities': {
            'parse_docx': True,
            'fill_fields': True,
            'convert_pdf': True,
            'reportlab': HAS_REPORTLAB,
            'pdf2docx': HAS_PDF2DOCX
        }
    }

@app.route('/parse-docx', methods=['POST'])
def parse_docx():
    """Parse DOCX and detect fields"""
    try:
        app.logger.info('📄 Parse DOCX request received')
        
        file = request.files.get('file')
        if not file:
            return {'success': False, 'error': 'No file uploaded'}, 400
        
        docx_bytes = file.read()
        app.logger.info(f'📦 File size: {len(docx_bytes)} bytes')
        
        fields = extract_fields_from_docx(docx_bytes)
        
        return {
            'success': True,
            'fields': [{'name': f, 'label': f.replace('_', ' ').title(), 'type': 'text'} for f in fields],
            'message': f'Found {len(fields)} fields'
        }
    
    except Exception as e:
        app.logger.error(f'❌ Parse error: {str(e)}')
        return {'success': False, 'error': str(e)}, 500

@app.route('/fill-and-pdf', methods=['POST'])
def fill_and_pdf():
    """Fill fields in DOCX and return PDF"""
    try:
        app.logger.info('🔵 Fill and PDF request received')
        
        data = request.json
        docx_base64 = data.get('docxBase64')
        field_values = data.get('fieldValues', {})
        title = data.get('title', 'signed_document')
        
        app.logger.info(f'📊 Title: {title}')
        app.logger.info(f'📋 Field count: {len(field_values)}')
        app.logger.info(f'📋 Fields: {list(field_values.keys())}')
        
        if not docx_base64:
            return {'success': False, 'error': 'No document provided'}, 400
        
        # Decode base64
        try:
            docx_bytes = base64.b64decode(docx_base64)
            app.logger.info(f'✅ Decoded DOCX: {len(docx_bytes)} bytes')
        except Exception as e:
            return {'success': False, 'error': f'Failed to decode DOCX: {str(e)}'}, 400
        
        # Fill fields in DOCX
        filled_docx_io = replace_fields_in_docx(docx_bytes, field_values)
        if not filled_docx_io:
            return {'success': False, 'error': 'Failed to replace fields'}, 500
        
        filled_docx_bytes = filled_docx_io.getvalue()
        app.logger.info(f'✅ DOCX filled: {len(filled_docx_bytes)} bytes')
        
        # Convert to PDF
        pdf_bytes = convert_docx_to_pdf_advanced(filled_docx_bytes, title)
        
        if not pdf_bytes:
            app.logger.warning('⚠️ PDF conversion failed, returning DOCX as fallback')
            pdf_bytes = filled_docx_bytes
            mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            filename = f'{title}_filled.docx'
        else:
            mimetype = 'application/pdf'
            filename = f'{title}_signed.pdf'
        
        app.logger.info(f'📤 Returning file: {filename} ({len(pdf_bytes)} bytes)')
        
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype=mimetype,
            as_attachment=True,
            download_name=filename
        )
    
    except Exception as e:
        app.logger.error(f'❌ Error: {str(e)}')
        return {'success': False, 'error': str(e)}, 500

@app.route('/docx-to-pdf', methods=['POST'])
def docx_to_pdf():
    """Convert DOCX to PDF directly"""
    try:
        app.logger.info('🔄 DOCX to PDF conversion request')
        
        data = request.json
        docx_base64 = data.get('docxBase64')
        title = data.get('title', 'document')
        
        if not docx_base64:
            return {'success': False, 'error': 'No document provided'}, 400
        
        docx_bytes = base64.b64decode(docx_base64)
        pdf_bytes = convert_docx_to_pdf_advanced(docx_bytes, title)
        
        if not pdf_bytes:
            return {'success': False, 'error': 'PDF conversion failed'}, 500
        
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'{title}.pdf'
        )
    
    except Exception as e:
        app.logger.error(f'❌ Error: {str(e)}')
        return {'success': False, 'error': str(e)}, 500

if __name__ == '__main__':
    app.logger.info('🚀 Starting DocSign PDF Service')
    app.logger.info(f'✓ ReportLab available: {HAS_REPORTLAB}')
    app.logger.info(f'✓ PDF2DOCX available: {HAS_PDF2DOCX}')
    app.run(host='0.0.0.0', port=5001, debug=True)
