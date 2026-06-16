import os
import json
import datetime
import tempfile
from io import BytesIO
from flask import Flask, render_template, request, jsonify, send_file
import pdfplumber
import pandas as pd

try:
    from pdf2image import convert_from_bytes
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

try:
    from docx import Document
except ImportError:
    pass

try:
    from pdf2docx import Converter as PDF2Docx
    PDF2DOCX_AVAILABLE = True
except ImportError:
    PDF2DOCX_AVAILABLE = False

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB


# ──────────────────────────────────────────────
# Analysis
# ──────────────────────────────────────────────
def analyze_pdf(pdf_bytes: bytes) -> dict:
    table_pages = text_pages = total_pages = 0
    try:
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            total_pages = len(pdf.pages)
            for page in pdf.pages[:min(5, total_pages)]:
                tables = page.extract_tables() or page.extract_tables(
                    table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
                if tables:
                    table_pages += 1
                if (page.extract_text() or "").strip():
                    text_pages += 1
    except Exception:
        pass

    if text_pages == 0 and table_pages == 0:
        return {"format": "word", "badge": "🔍 掃描版 PDF",
                "reason": "未偵測到可擷取的文字，建議啟用 OCR 再轉為 Word。", "use_ocr": True}
    elif table_pages >= text_pages * 0.6:
        return {"format": "excel", "badge": "📊 表格型 PDF",
                "reason": f"前 {min(5, total_pages)} 頁中有 {table_pages} 頁含表格，最適合匯出為 Excel。", "use_ocr": False}
    elif table_pages > 0:
        return {"format": "word", "badge": "📝 混合型 PDF",
                "reason": f"文字與表格混合（{table_pages} 頁有表格），建議轉為 Word。", "use_ocr": False}
    else:
        return {"format": "markdown", "badge": "📄 純文字 PDF",
                "reason": "純文字內容，Markdown 最輕量且易讀。", "use_ocr": False}


# ──────────────────────────────────────────────
# Extraction
# ──────────────────────────────────────────────
def _unique_cols(headers: list) -> list:
    seen: dict = {}
    out = []
    for h in headers:
        key = h if h else "col"
        if key in seen:
            seen[key] += 1
            out.append(f"{key}_{seen[key]}")
        else:
            seen[key] = 0
            out.append(key)
    return out


def extract_tables(pdf_bytes: bytes) -> dict:
    result = {}
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables(table_settings={
                "vertical_strategy": "lines", "horizontal_strategy": "lines"})
            if not tables:
                tables = page.extract_tables(table_settings={
                    "vertical_strategy": "text", "horizontal_strategy": "text"})
            if tables:
                dfs = []
                for t in tables:
                    cleaned = [[c if c is not None else "" for c in row] for row in t]
                    if not cleaned:
                        continue
                    if len(cleaned) > 1:
                        df = pd.DataFrame(cleaned[1:], columns=_unique_cols(cleaned[0]))
                    else:
                        df = pd.DataFrame(cleaned)
                    dfs.append(df.astype(str).replace("None", ""))
                if dfs:
                    try:
                        result[f"Page {i+1}"] = pd.concat(dfs, ignore_index=True)
                    except Exception:
                        normed = [df.rename(columns=str) for df in dfs]
                        result[f"Page {i+1}"] = pd.concat(normed, ignore_index=True)
    return result


def extract_ocr(pdf_bytes: bytes) -> dict:
    images = convert_from_bytes(pdf_bytes, dpi=300)
    result = {}
    for i, img in enumerate(images):
        try:
            text = pytesseract.image_to_string(img, lang="chi_tra+eng")
        except Exception:
            text = pytesseract.image_to_string(img)
        if text.strip():
            result[f"Page {i+1}"] = text.strip()
    return result


def extract_text_pages(pdf_bytes: bytes) -> dict:
    result = {}
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if text.strip():
                result[f"Page {i+1}"] = text.strip()
    return result


# ──────────────────────────────────────────────
# Build output
# ──────────────────────────────────────────────
def build_excel(data: dict) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sheet, df in data.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
    return out.getvalue()


def build_word(pdf_bytes: bytes, is_ocr: bool, ocr_data: dict = None) -> bytes:
    if is_ocr and ocr_data:
        doc = Document()
        doc.add_heading("PDF 轉換結果（OCR）", level=0)
        for page, text in ocr_data.items():
            doc.add_heading(page, level=1)
            doc.add_paragraph(text)
        out = BytesIO()
        doc.save(out)
        return out.getvalue()

    if PDF2DOCX_AVAILABLE:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
            f.write(pdf_bytes)
            tmp_pdf = f.name
        tmp_docx = tmp_pdf.replace(".pdf", ".docx")
        try:
            cv = PDF2Docx(tmp_pdf)
            cv.convert(tmp_docx)
            cv.close()
            with open(tmp_docx, "rb") as f:
                return f.read()
        finally:
            os.unlink(tmp_pdf)
            if os.path.exists(tmp_docx):
                os.unlink(tmp_docx)

    doc = Document()
    doc.add_heading("PDF 轉換結果", level=0)
    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def build_word_editable(pdf_bytes: bytes, table_data: dict = None) -> bytes:
    pages = extract_text_pages(pdf_bytes)
    doc = Document()
    if pages:
        for label, text in pages.items():
            doc.add_heading(label, level=1)
            for line in text.splitlines():
                if line.strip():
                    doc.add_paragraph(line.strip())
            doc.add_paragraph()
    elif table_data:
        for label, df in table_data.items():
            doc.add_heading(label, level=1)
            for _, row in df.iterrows():
                line = "　".join(str(v) for v in row if str(v).strip())
                if line.strip():
                    doc.add_paragraph(line)
            doc.add_paragraph()
    else:
        doc.add_paragraph("⚠️ 此 PDF 無法擷取文字，請改用 OCR 模式。")
    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def build_markdown(data: dict, is_ocr: bool) -> str:
    parts = ["# PDF 轉換結果\n"]
    for page, content in data.items():
        parts.append(f"\n## {page}\n")
        parts.append(content if is_ocr else content.to_markdown(index=False))
        parts.append("\n")
    return "\n".join(parts)


def build_json(data: dict, is_ocr: bool) -> str:
    payload = data if is_ocr else {p: df.to_dict(orient="records") for p, df in data.items()}
    return json.dumps(payload, ensure_ascii=False, indent=2)


# ──────────────────────────────────────────────
# Routes
# ──────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html", ocr_available=OCR_AVAILABLE)


@app.route("/analyze", methods=["POST"])
def analyze_route():
    if "pdf" not in request.files:
        return jsonify({"error": "No file"}), 400
    rec = analyze_pdf(request.files["pdf"].read())
    return jsonify(rec)


@app.route("/convert", methods=["POST"])
def convert_route():
    if "pdf" not in request.files:
        return jsonify({"error": "No file"}), 400

    pdf_bytes = request.files["pdf"].read()
    fmt = request.form.get("format", "excel")
    use_ocr = request.form.get("ocr") == "true"
    word_editable = request.form.get("word_editable") == "true"
    base_name = request.form.get("filename", "converted")
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    if use_ocr and OCR_AVAILABLE:
        data = extract_ocr(pdf_bytes)
        is_ocr = True
    else:
        data = extract_tables(pdf_bytes)
        is_ocr = False

    if fmt == "excel":
        if not data:
            return jsonify({"error": "no_data"}), 422
        file_bytes = build_excel(data)
        filename = f"{base_name}_{ts}.xlsx"
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    elif fmt == "word":
        if word_editable:
            file_bytes = build_word_editable(pdf_bytes, table_data=data if not is_ocr else None)
        else:
            file_bytes = build_word(pdf_bytes, is_ocr, ocr_data=data if is_ocr else None)
        filename = f"{base_name}_{ts}.docx"
        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    elif fmt == "markdown":
        if not data:
            return jsonify({"error": "no_data"}), 422
        file_bytes = build_markdown(data, is_ocr).encode("utf-8")
        filename = f"{base_name}_{ts}.md"
        mime = "text/markdown"

    elif fmt == "json":
        if not data:
            return jsonify({"error": "no_data"}), 422
        file_bytes = build_json(data, is_ocr).encode("utf-8")
        filename = f"{base_name}_{ts}.json"
        mime = "application/json"

    else:
        return jsonify({"error": "Unknown format"}), 400

    return send_file(BytesIO(file_bytes), mimetype=mime,
                     as_attachment=True, download_name=filename)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
