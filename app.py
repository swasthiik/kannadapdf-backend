from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from pypdf import PdfWriter, PdfReader
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import os, tempfile, uuid, io

app = Flask(__name__)
CORS(app)

TMP = tempfile.gettempdir()

def tmp_path(ext):
    return os.path.join(TMP, f"{uuid.uuid4()}.{ext}")

def cleanup(*paths):
    for p in paths:
        try:
            if p and os.path.exists(p): os.remove(p)
        except: pass

@app.route('/')
def health():
    return jsonify({'status': 'ok', 'message': 'ಕನ್ನಡ PDF API ಚಾಲನೆಯಲ್ಲಿದೆ'})

@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'ಫೈಲ್ ಕಂಡುಬಂದಿಲ್ಲ'}), 400
    file = request.files['file']
    if not file.filename.lower().endswith(('.doc', '.docx')):
        return jsonify({'error': 'DOC ಅಥವಾ DOCX ಮಾತ್ರ ಸ್ವೀಕಾರ'}), 400
    input_path = tmp_path('docx')
    file.save(input_path)
    try:
        doc = Document(input_path)
        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=A4)
        width, height = A4
        y = height - 50
        c.setFont("Helvetica", 11)
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                y -= 10
                continue
            words = text.split()
            line = ""
            for word in words:
                test = line + " " + word if line else word
                if c.stringWidth(test, "Helvetica", 11) < width - 80:
                    line = test
                else:
                    c.drawString(40, y, line)
                    y -= 18
                    line = word
                    if y < 60:
                        c.showPage()
                        c.setFont("Helvetica", 11)
                        y = height - 50
            if line:
                c.drawString(40, y, line)
                y -= 18
            if y < 60:
                c.showPage()
                c.setFont("Helvetica", 11)
                y = height - 50
        c.save()
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name='converted.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cleanup(input_path)

@app.route('/pdf-merge', methods=['POST'])
def pdf_merge():
    files = request.files.getlist('files')
    if len(files) < 2:
        return jsonify({'error': 'ಕನಿಷ್ಠ 2 PDF ಬೇಕು'}), 400
    saved = []
    try:
        writer = PdfWriter()
        for file in files:
            path = tmp_path('pdf')
            file.save(path)
            saved.append(path)
            for page in PdfReader(path).pages:
                writer.add_page(page)
        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name='merged.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cleanup(*saved)

@app.route('/pdf-split', methods=['POST'])
def pdf_split():
    if 'file' not in request.files:
        return jsonify({'error': 'ಫೈಲ್ ಕಂಡುಬಂದಿಲ್ಲ'}), 400
    file = request.files['file']
    pages_param = request.form.get('pages', '').strip()
    input_path = tmp_path('pdf')
    file.save(input_path)
    try:
        reader = PdfReader(input_path)
        total = len(reader.pages)
        pages_to_extract = parse_pages(pages_param, total)
        writer = PdfWriter()
        for i in pages_to_extract:
            writer.add_page(reader.pages[i])
        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name='split.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cleanup(input_path)

@app.route('/img-to-pdf', methods=['POST'])
def img_to_pdf():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'ಚಿತ್ರಗಳು ಕಂಡುಬಂದಿಲ್ಲ'}), 400
    saved = []
    try:
        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=A4)
        w, h = A4
        for file in files:
            path = tmp_path('img')
            file.save(path)
            saved.append(path)
            img = Image.open(path)
            img_w, img_h = img.size
            ratio = min(w / img_w, h / img_h)
            new_w, new_h = img_w * ratio, img_h * ratio
            c.drawImage(path, (w-new_w)/2, (h-new_h)/2, new_w, new_h)
            c.showPage()
        c.save()
        buf.seek(0)
        return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name='images.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cleanup(*saved)

@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files:
        return jsonify({'error': 'ಫೈಲ್ ಕಂಡುಬಂದಿಲ್ಲ'}), 400
    file = request.files['file']
    input_path = tmp_path('pdf')
    file.save(input_path)
    try:
        reader = PdfReader(input_path)
        doc = Document()
        doc.add_heading('Converted from PDF', 0)
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                doc.add_heading(f'Page {i+1}', level=2)
                doc.add_paragraph(text)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', as_attachment=True, download_name='converted.docx')
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cleanup(input_path)

def parse_pages(pages_param, total):
    if not pages_param:
        return list(range(total))
    pages = set()
    for part in pages_param.split(','):
        part = part.strip()
        if '-' in part:
            s, e = part.split('-')
            for p in range(int(s)-1, int(e)):
                if 0 <= p < total: pages.add(p)
        else:
            p = int(part) - 1
            if 0 <= p < total: pages.add(p)
    return sorted(pages)


# ─────────────────────────────────────────
# GEMINI AI — shared setup
# ─────────────────────────────────────────
import requests as req_lib, json

GEMINI_KEY = os.environ.get('GEMINI_API_KEY', '')
GEMINI_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent'
GEMINI_CHAT_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent'

def call_gemini(prompt):
    headers = {'Content-Type': 'application/json'}
    data = {'contents': [{'parts': [{'text': prompt}]}]}
    r = req_lib.post(f'{GEMINI_URL}?key={GEMINI_KEY}', headers=headers, json=data, timeout=40)
    r.raise_for_status()
    return r.json()['candidates'][0]['content']['parts'][0]['text']


# ─────────────────────────────────────────
# AI CHAT — Resume Chatbot (NEW)
# ─────────────────────────────────────────
@app.route('/ai-chat', methods=['POST'])
def ai_chat():
    """Conversational AI chatbot for resume building"""
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No data provided'}), 400

    messages  = data.get('messages', [])
    system    = data.get('system', '')

    if not messages:
        return jsonify({'error': 'No messages provided'}), 400

    try:
        # Build Gemini contents array from chat history
        contents = []

        # Add system prompt as first user turn (Gemini doesn't have a system field)
        if system:
            contents.append({
                'role': 'user',
                'parts': [{'text': f'[SYSTEM INSTRUCTIONS]\n{system}\n[END SYSTEM]\n\nAcknowledge you understand.'}]
            })
            contents.append({
                'role': 'model',
                'parts': [{'text': 'Understood! I am your AI Resume Assistant. I will help users build professional resumes through conversation.'}]
            })

        # Add conversation history
        for msg in messages:
            role = 'user' if msg.get('role') == 'user' else 'model'
            contents.append({
                'role': role,
                'parts': [{'text': msg.get('content', '')}]
            })

        headers = {'Content-Type': 'application/json'}
        payload = {'contents': contents}

        r = req_lib.post(
            f'{GEMINI_CHAT_URL}?key={GEMINI_KEY}',
            headers=headers,
            json=payload,
            timeout=40
        )
        r.raise_for_status()
        reply = r.json()['candidates'][0]['content']['parts'][0]['text']
        return jsonify({'reply': reply})

    except req_lib.exceptions.HTTPError as e:
        return jsonify({'error': f'Gemini API error: {str(e)}'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ─────────────────────────────────────────
# AI RESUME BUILDER — Gemini AI
# ─────────────────────────────────────────
def build_resume_pdf(text):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        rightMargin=0.65*inch, leftMargin=0.65*inch,
        topMargin=0.65*inch, bottomMargin=0.65*inch)

    name_s  = ParagraphStyle('n', fontSize=18, fontName='Helvetica-Bold', alignment=TA_CENTER, spaceAfter=3, textColor=colors.HexColor('#1a1a2e'))
    cont_s  = ParagraphStyle('c', fontSize=9,  fontName='Helvetica', alignment=TA_CENTER, spaceAfter=10, textColor=colors.HexColor('#555'))
    head_s  = ParagraphStyle('h', fontSize=11, fontName='Helvetica-Bold', spaceBefore=8, spaceAfter=3, textColor=colors.HexColor('#e63946'))
    body_s  = ParagraphStyle('b', fontSize=9.5, fontName='Helvetica', spaceAfter=3, leading=14, textColor=colors.HexColor('#222'))
    bull_s  = ParagraphStyle('bl', fontSize=9.5, fontName='Helvetica', spaceAfter=2, leading=13, leftIndent=14, textColor=colors.HexColor('#333'))

    story = []
    lines = [l.rstrip() for l in text.strip().split('\n')]
    name_done = False

    for line in lines:
        if not line.strip():
            story.append(Spacer(1, 3))
            continue
        if line.startswith('##'):
            txt = line.replace('##','').replace('**','').strip()
            story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#e63946'), spaceAfter=4))
            story.append(Paragraph(txt.upper(), head_s))
        elif line.startswith('#') and not name_done:
            story.append(Paragraph(line.replace('#','').replace('**','').strip(), name_s))
            name_done = True
        elif line.startswith('- ') or line.startswith('• '):
            story.append(Paragraph('• ' + line[2:].replace('**','').strip(), bull_s))
        elif any(x in line for x in ['@','|','+91','linkedin','github','http']) and not name_done:
            story.append(Paragraph(line.replace('**','').strip(), cont_s))
            name_done = True
        elif line.startswith('**') and line.endswith('**'):
            story.append(Paragraph(line.replace('**','').strip(), ParagraphStyle('sb', fontSize=9.5, fontName='Helvetica-Bold', spaceAfter=2, textColor=colors.HexColor('#222'))))
        else:
            clean = line.replace('**','').replace('##','').replace('#','').strip()
            if clean:
                if not name_done and len(clean) < 50:
                    story.append(Paragraph(clean, name_s))
                    name_done = True
                else:
                    story.append(Paragraph(clean, body_s))

    doc.build(story)
    buf.seek(0)
    return buf

@app.route('/ai-resume-build', methods=['POST'])
def ai_resume_build():
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No data provided'}), 400

    name          = data.get('name','')
    email         = data.get('email','')
    phone         = data.get('phone','')
    linkedin      = data.get('linkedin','')
    summary       = data.get('summary','')
    skills        = data.get('skills','')
    education     = data.get('education','')
    experience    = data.get('experience','')
    projects      = data.get('projects','')
    certifications= data.get('certifications','')

    prompt = f"""Create a professional resume in clean plain text format for:

Name: {name}
Email: {email}
Phone: {phone}
LinkedIn: {linkedin}
Summary: {summary}
Skills: {skills}
Education: {education}
Experience: {experience}
Projects: {projects}
Certifications: {certifications}

Format rules:
- First line: just the person name
- Second line: contact info separated by | symbols
- Use ## for section headings (OBJECTIVE, EDUCATION, EXPERIENCE, PROJECTS, SKILLS, CERTIFICATIONS)
- Use - for bullet points
- Use **text** for bold job titles or company names
- Keep it professional, clean, ATS-friendly
- Only output the resume text, nothing else"""

    try:
        resume_text = call_gemini(prompt)
        pdf_buf = build_resume_pdf(resume_text)
        return send_file(pdf_buf, mimetype='application/pdf', as_attachment=True, download_name='resume.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/ai-resume-improve', methods=['POST'])
def ai_resume_improve():
    if 'file' not in request.files:
        return jsonify({'error': 'File not found'}), 400

    file = request.files['file']
    job_role = request.form.get('job_role', 'Software Engineer')
    input_path = tmp_path('pdf' if file.filename.lower().endswith('.pdf') else 'docx')
    file.save(input_path)

    try:
        if input_path.endswith('.pdf'):
            reader = PdfReader(input_path)
            resume_text = ' '.join([p.extract_text() or '' for p in reader.pages])
        else:
            doc = Document(input_path)
            resume_text = ' '.join([p.text for p in doc.paragraphs if p.text.strip()])

        prompt = f"""You are a professional resume writer. Improve and reformat this resume for the role of "{job_role}".

Original resume:
{resume_text}

Instructions:
- Keep all original information, just improve wording and structure
- First line: just the person name
- Second line: contact info with | separators
- Use ## for section headings (OBJECTIVE, EDUCATION, EXPERIENCE, PROJECTS, SKILLS, CERTIFICATIONS)
- Use - for bullet points
- Use **text** for bold job titles or company names
- Make it ATS-friendly and professional
- Only output the improved resume, nothing else"""

        improved_text = call_gemini(prompt)
        pdf_buf = build_resume_pdf(improved_text)
        return send_file(pdf_buf, mimetype='application/pdf', as_attachment=True, download_name='improved_resume.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        cleanup(input_path)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)