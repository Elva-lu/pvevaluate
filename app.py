from flask import Flask, request, send_file
import os
import io
from docx import Document

app = Flask(__name__)

def load_evaluation_logic():
    logic_doc = Document("CC2025016_MA評估回饋20250417 1.docx")
    return "\n".join([p.text for p in logic_doc.paragraphs if p.text.strip()])

def extract_text_file(txt_path):
    with open(txt_path, "r", encoding="utf-8") as f:
        return f.read()

def generate_report(text_content, logic_text):
    doc = Document()
    doc.add_heading("評估回饋報告", level=1)
    doc.add_heading("📄 文字檔內容摘要", level=2)
    doc.add_paragraph(text_content)
    doc.add_heading("🧠 評估邏輯", level=2)
    doc.add_paragraph(logic_text)
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

@app.route("/evaluate", methods=["POST"])
@app.route("/")
def home():
    return "👋 Flask 服務已上線，請使用 POST /evaluate 上傳 TXT"

@app.route("/evaluate", methods=["POST"])
def evaluate():
    if "file" not in request.files:
        return "請提供 TXT 檔案", 400
    file = request.files["file"]
    file.save("uploaded.txt")
    logic = load_evaluation_logic()
    text_content = extract_text_file("uploaded.txt")
    report = generate_report(text_content, logic)
    os.remove("uploaded.txt")
    return send_file(report, as_attachment=True, download_name="評估回饋報告.docx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
