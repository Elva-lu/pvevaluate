
from flask import Flask, request, send_file
import os
import io
from docx import Document
import fitz # PyMuPDF

app = Flask(__name__)

def load_evaluation_logic():
    logic_doc = Document("CC2025016_MA評估回饋20250417 1.docx")
    return "\n".join([p.text for p in logic_doc.paragraphs if p.text.strip()])

def extract_pdf_text(pdf_path):
    doc = fitz.open(pdf_path)
    return "".join([page.get_text() for page in doc])

def generate_report(pdf_text, logic_text):
    doc = Document()
    doc.add_heading("評估回饋報告", level=1)
    doc.add_heading("📄 PDF 文件內容摘要", level=2)
    doc.add_paragraph(pdf_text)
    doc.add_heading("🧠 評估邏輯", level=2)
    doc.add_paragraph(logic_text)
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream
@app.route("/evaluate", methods=["POST"])
def evaluate():
    if "file" not in request.files:
        return "請提供 PDF 檔案", 400
    file = request.files["file"]
    file.save("uploaded.pdf")
    logic = load_evaluation_logic()
    pdf_text = extract_pdf_text("uploaded.pdf")
    report = generate_report(pdf_text, logic)
    os.remove("uploaded.pdf")
    return send_file(report, as_attachment=True, download_name="評估回饋報告.docx")

if __name__ == "__main__":
    app.run()
