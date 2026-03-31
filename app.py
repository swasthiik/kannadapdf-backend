from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import subprocess
import os
import tempfile
import uuid
import fitz  # PyMuPDF

app = Flask(__name__)
CORS(app)  # Allow your frontend to call this API

UPLOAD_FOLDER = tempfile.gettempdir()

# ─────────────────────────────────────────
# 1. WORD TO PDF
# ─────────────────────────────────────────
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'ಫೈಲ್ ಕಂಡುಬಂದಿಲ್ಲ'}), 400

    file = request.files['file']
    if not file.filename.endswith(('.doc', '.docx')):
        return jsonify({'error': 'DOC ಅಥವಾ DOCX ಫೈಲ್ ಮಾತ್ರ ಸ್ವೀಕಾರ'}), 400

    # Save uploaded file
    input_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}_{file.filename}")
    file.save(input_path)

    # Convert using LibreOffice
    output_dir = UPLOAD_FOLDER
    try:
        subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'pdf',
            '--outdir', output_dir, input_path
        ], check=True, timeout=30)

        output_path = input_path.rsplit('.', 1)[0] + '.pdf'

        response = send_file(
            output_path,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='converted.pdf'
        )
        return response

    except subprocess.CalledProcessError:
        return jsonify({'error': 'ಪರಿವರ್ತನೆ ವಿಫಲವಾಗಿದೆ'}), 500
    finally:
        # Clean up files after sending
        if os.path.exists(input_path):
            os.remove(input_path)


# ─────────────────────────────────────────
# 2. PDF TO WORD
# ─────────────────────────────────────────
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files:
        return jsonify({'error': 'ಫೈಲ್ ಕಂಡುಬಂದಿಲ್ಲ'}), 400

    file = request.files['file']
    if not file.filename.endswith('.pdf'):
        return jsonify({'error': 'PDF ಫೈಲ್ ಮಾತ್ರ ಸ್ವೀಕಾರ'}), 400

    input_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}.pdf")
    file.save(input_path)

    output_dir = UPLOAD_FOLDER
    try:
        subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'docx',
            '--outdir', output_dir, input_path
        ], check=True, timeout=30)

        output_path = input_path.replace('.pdf', '.docx')

        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='converted.docx'
        )

    except subprocess.CalledProcessError:
        return jsonify({'error': 'ಪರಿವರ್ತನೆ ವಿಫಲವಾಗಿದೆ'}), 500
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)


# ─────────────────────────────────────────
# 3. PDF SPLIT
# ─────────────────────────────────────────
@app.route('/pdf-split', methods=['POST'])
def pdf_split():
    if 'file' not in request.files:
        return jsonify({'error': 'ಫೈಲ್ ಕಂಡುಬಂದಿಲ್ಲ'}), 400

    file = request.files['file']
    pages_param = request.form.get('pages', '')  # e.g. "1-3" or "2,4,6"

    if not file.filename.endswith('.pdf'):
        return jsonify({'error': 'PDF ಫೈಲ್ ಮಾತ್ರ ಸ್ವೀಕಾರ'}), 400

    input_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}.pdf")
    file.save(input_path)

    try:
        doc = fitz.open(input_path)
        total_pages = len(doc)

        # Parse page range
        pages_to_extract = parse_pages(pages_param, total_pages)

        # Create new PDF with selected pages
        new_doc = fitz.open()
        for page_num in pages_to_extract:
            new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)

        output_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}_split.pdf")
        new_doc.save(output_path)
        new_doc.close()
        doc.close()

        return send_file(
            output_path,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='split.pdf'
        )

    except Exception as e:
        return jsonify({'error': f'ದೋಷ: {str(e)}'}), 500
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)


def parse_pages(pages_param, total_pages):
    """Parse page string like '1-3' or '1,3,5' into 0-indexed list"""
    if not pages_param.strip():
        # Return all pages if no param given
        return list(range(total_pages))

    pages = set()
    parts = pages_param.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            start, end = part.split('-')
            for p in range(int(start)-1, int(end)):
                if 0 <= p < total_pages:
                    pages.add(p)
        else:
            p = int(part) - 1
            if 0 <= p < total_pages:
                pages.add(p)
    return sorted(pages)


# ─────────────────────────────────────────
# 4. HEALTH CHECK
# ─────────────────────────────────────────
@app.route('/')
def health():
    return jsonify({
        'status': 'ok',
        'message': 'ಕನ್ನಡ PDF API ಚಾಲನೆಯಲ್ಲಿದೆ',
        'tools': ['word-to-pdf', 'pdf-to-word', 'pdf-split']
    })


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
