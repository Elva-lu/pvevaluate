from flask import Flask, request, send_file
import io
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import re

app = Flask(__name__)

# 判斷是否為中文
def is_chinese(char):
    return '\u4e00' <= char <= '\u9fff'

# 自動套用字型（中文：標楷體，英文數字：Times New Roman）
def add_paragraph_with_fonts(doc, text):
    p = doc.add_paragraph()
    run = None
    current_font = None

    for char in text:
        if is_chinese(char):
            font = '標楷體'
        elif char.isalpha() or char.isdigit():
            font = 'Times New Roman'
        else:
            font = current_font or 'Times New Roman'

        if font != current_font:
            run = p.add_run()
            run.font.size = Pt(12)
            run.font.name = font
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font)
            current_font = font

        run.add_text(char)

    return p

@app.route("/evaluate", methods=["POST"])
def evaluate():
    try:
        raw_text = request.form.get("text")
        if not raw_text:
            return "❌ 缺少 text 欄位", 400

        # 抽出 part_number
        match = re.search(r"\{[^\}]*part_number\s*[:：]\s*(\S+)[^\}]*\}", raw_text)
        part_number = match.group(1) if match else "NA"

        # 移除 JSON 標記那一行
        content = re.sub(r"\{[^\}]*part_number\s*[:：]\s*\S+[^\}]*\}\s*", "", raw_text).strip()

        # 建立 Word 檔案
        doc = Document()
        add_paragraph_with_fonts(doc, f"料品號: {part_number}")
        doc.add_paragraph()  # 空行
        for line in content.splitlines():
            add_paragraph_with_fonts(doc, line)

        stream = io.BytesIO()
        doc.save(stream)
        stream.seek(0)

        filename = f"ADR_Report_{part_number}.docx"
        return send_file(stream, as_attachment=True, download_name=filename)

    except Exception as e:
        return f"❌ 發生錯誤: {str(e)}", 500

@app.route("/", methods=["GET"])
def home():
    return "✅ 上傳報告到 /evaluate via POST"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, threaded=True)
