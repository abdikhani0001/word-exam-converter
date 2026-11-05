import os
import random
from flask import Flask, render_template, request, send_file, redirect
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pdfplumber

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)

    if file:
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        school_name = request.form.get('school_name', 'AL-AMIIN PRIMARY & SECONDARY SCHOOL')
        exam_title = request.form.get('exam_title', 'END OF TERM EXAMINATION')
        subject = request.form.get('subject', 'Saynis')
        class_grade = request.form.get('class_grade', 'Fasalka 1aad')

        extracted_text = ""
        ext = file.filename.lower().split('.')[-1]
        if ext == 'pdf':
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        extracted_text += t + "\n"
        elif ext == 'docx':
            from docx import Document as DocxReader
            doc_reader = DocxReader(file_path)
            for p in doc_reader.paragraphs:
                extracted_text += p.text + "\n"

        doc = Document()
        for s in doc.sections:
            s.top_margin = Inches(1)
            s.bottom_margin = Inches(1)
            s.left_margin = Inches(1)
            s.right_margin = Inches(1)

        # Header Block
        p_school = doc.add_paragraph()
        p_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p_school.add_run(school_name.upper())
        r1.bold = True
        r1.font.size = Pt(14)

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p_title.add_run(exam_title.upper())
        r2.bold = True
        r2.font.size = Pt(12)

        p_meta = doc.add_paragraph()
        p_meta.add_run(f"SUBJECT: {subject.upper()}\t\tCLASS: {class_grade}\n")
        p_meta.add_run("DATE: ________________________\t\tMARKS: /50\n")
        p_meta.add_run("=" * 55)

        for line in extracted_text.split('\n'):
            if line.strip():
                doc.add_paragraph(line)

        output_path = os.path.join(OUTPUT_FOLDER, "Formatted_Exam.docx")
        doc.save(output_path)
        return send_file(output_path, as_attachment=True)

@app.route('/generate-smart-exam', methods=['POST'])
def generate_smart_exam():
    if 'textbook_file' not in request.files:
        return redirect(request.url)
    file = request.files['textbook_file']
    if file.filename == '':
        return redirect(request.url)

    if file:
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        school_name = request.form.get('school_name', 'AL-AMIIN PRIMARY & SECONDARY SCHOOL')
        subject = request.form.get('subject', 'Saynis')
        class_grade = request.form.get('class_grade', 'Fasalka 1aad')
        
        sec_a_title = request.form.get('sec_a_title', 'SECTION A: MULTIPLE CHOICE')
        sec_a_count = int(request.form.get('sec_a_count', 5))
        
        sec_b_title = request.form.get('sec_b_title', 'SECTION B: STRUCTURED QUESTIONS')
        sec_b_count = int(request.form.get('sec_b_count', 5))
        
        footer_note = request.form.get('footer_note', 'Answer all questions.')

        # Extract text content from textbook
        text_content = ""
        ext = file.filename.lower().split('.')[-1]
        if ext == 'pdf':
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        text_content += t + "\n"
        elif ext == 'docx':
            from docx import Document as DocxReader
            doc_reader = DocxReader(file_path)
            for p in doc_reader.paragraphs:
                text_content += p.text + "\n"

        sentences = [s.strip() for s in text_content.replace('\n', '. ').split('.') if len(s.strip()) > 25]
        random.shuffle(sentences)

        doc = Document()
        for s in doc.sections:
            s.top_margin = Inches(1)
            s.bottom_margin = Inches(1)
            s.left_margin = Inches(1)
            s.right_margin = Inches(1)

        # Institutional Header
        p_school = doc.add_paragraph()
        p_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p_school.add_run(school_name.upper())
        r1.bold = True
        r1.font.name = 'Arial'
        r1.font.size = Pt(14)

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p_title.add_run(f"END OF TERM EXAMINATION — {subject.upper()}")
        r2.bold = True
        r2.font.name = 'Arial'
        r2.font.size = Pt(12)

        p_meta = doc.add_paragraph()
        p_meta.add_run(f"CLASS: {class_grade}\t\t\t\tTIME: 1 HOUR 30 MINS\n")
        p_meta.add_run("STUDENT NAME: _____________________________________\tMARKS: /50\n")
        p_meta.add_run("=" * 55)

        total_questions_needed = sec_a_count + sec_b_count
        if len(sentences) < total_questions_needed:
            # Pad sentences if textbook sample is small
            sentences = sentences * (total_questions_needed // max(len(sentences), 1) + 1)

        idx = 0

        # --- SECTION A ---
        p_seca = doc.add_paragraph()
        r_seca = p_seca.add_run(sec_a_title)
        r_seca.bold = True
        r_seca.font.size = Pt(11)

        for i in range(1, sec_a_count + 1):
            concept = sentences[idx] if idx < len(sentences) else "core fundamental concept"
            idx += 1
            doc.add_paragraph(f"{i}. Which of the following best describes or relates to: {concept}?")
            doc.add_paragraph("   A) True statement regarding the concept\n   B) Incorrect alternative option\n   C) Partial contradictory factor\n   D) None of the above\n")

        # --- SECTION B ---
        p_secb = doc.add_paragraph()
        r_secb = p_secb.add_run(sec_b_title)
        r_secb.bold = True
        r_secb.font.size = Pt(11)

        for i in range(1, sec_b_count + 1):
            concept = sentences[idx] if idx < len(sentences) else "subject matter principle"
            idx += 1
            doc.add_paragraph(f"{i}. Explain the significance of {concept} in detail.")
            doc.add_paragraph("Answer: ___________________________________________________________________\n___________________________________________________________________________\n")

        # Footer Instruction Note
        p_foot = doc.add_paragraph()
        p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_foot = p_foot.add_run(f"\n*** {footer_note} ***")
        r_foot.italic = True
        r_foot.font.size = Pt(10)

        output_path = os.path.join(OUTPUT_FOLDER, "Structured_AI_Exam.docx")
        doc.save(output_path)
        return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
