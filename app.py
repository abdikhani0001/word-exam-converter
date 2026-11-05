import os
from flask import Flask, render_template, request, send_file, redirect, url_for
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pdfplumber
from PIL import Image
import pytesseract

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

        # Form fields
        school_name = request.form.get('school_name', 'AL-AMIIN PRIMARY & SECONDARY SCHOOL')
        exam_title = request.form.get('exam_title', 'END OF TERM EXAMINATION')
        subject = request.form.get('subject', 'Saynis')
        class_grade = request.form.get('class_grade', 'Fasalka 1aad')
        layout_style = request.form.get('layout_style', 'Traditional Kenya')

        # Extract text based on file extension
        extracted_text = ""
        ext = file.filename.lower().split('.')[-1]
        
        if ext == 'pdf':
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        extracted_text += text + "\n"
                    else:
                        img = page.to_image(resolution=150).original
                        extracted_text += pytesseract.image_to_string(img) + "\n"
        elif ext in ['jpg', 'jpeg', 'png']:
            img = Image.open(file_path)
            extracted_text = pytesseract.image_to_string(img)
        elif ext == 'docx':
            from docx import Document as DocxReader
            doc_reader = DocxReader(file_path)
            for p in doc_reader.paragraphs:
                extracted_text += p.text + "\n"

        # Create Word Document conforming to Kenyan/Traditional Exam Layout
        doc = Document()
        
        # Header Section
        p_school = doc.add_paragraph()
        p_school.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_school = p_school.add_run(school_name.upper())
        run_school.bold = True
        run_school.font.size = Pt(14)

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_title = p_title.add_run(exam_title.upper())
        run_title.bold = True
        run_title.font.size = Pt(12)

        p_meta = doc.add_paragraph()
        p_meta.add_run(f"SUBJECT: {subject.upper()}\t\tCLASS: {class_grade}\n")
        p_meta.add_run("DATE: ________________________\t\tMARKS: /50\n")
        p_meta.add_run("-" * 55)

        # Add extracted content body
        for line in extracted_text.split('\n'):
            doc.add_paragraph(line)

        output_filename = "Converted_Exam.docx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        doc.save(output_path)

        return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
