from flask import Flask, request, send_file
import os
import io
from docx import Document

app = Flask(__name__)

TEMPLATE_PATH = "CC2025016_MA評估回饋20250417.docx"

# 預先載入邏輯
if not os.path.exists(TEMPLATE_PATH):
    LOGIC_TEXT = None
    STATUS_MESSAGE = "⚠️ 找不到範本檔案"
else:
    logic_doc = Document(TEMPLATE_PATH)
    LOGIC_TEXT = "\n".join([p.text for p in logic_doc.paragraphs if p.text.strip()])
    STATUS_MESSAGE = f"依 {os.path.basename(TEMPLATE_PATH)} 邏輯"

def generate_report(text_content, logic_text, status_message):
    doc = Document()
    doc.add_heading("評估回饋報告", level=1)
    doc.add_paragraph(f"📌 {status_message}")
    doc.add_heading("📄 文字內容摘要", level=2)
    doc.add_paragraph(text_content)
    doc.add_heading("🧠 評估邏輯", level=2)
    doc.add_paragraph(logic_text)
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

@app.route("/", methods=["GET"])
def home():
    return "👋 Flask 服務已上線，請使用 POST /evaluate 傳送文字"

@app.route("/evaluate", methods=["POST"])
def evaluate():
    text_content = request.form.get("text")
    if not text_content:
        return "❌ 請提供文字內容（欄位名稱為 text）", 400

    if not LOGIC_TEXT:
        return STATUS_MESSAGE, 500

    report = generate_report(text_content, LOGIC_TEXT, STATUS_MESSAGE)
    response = send_file(report, as_attachment=True, download_name="評估回饋報告.docx")
    response.headers["X-Status-Message"] = STATUS_MESSAGE
    return response

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
