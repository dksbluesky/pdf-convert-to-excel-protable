import os
import json
import datetime
import tempfile
import concurrent.futures
from urllib.parse import quote
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

GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY', '')
GEMINI_AVAILABLE = bool(GEMINI_API_KEY)
if GEMINI_AVAILABLE:
    try:
        import google.generativeai as genai
        genai.configure(api_key=GEMINI_API_KEY)
    except ImportError:
        GEMINI_AVAILABLE = False

GEMINI_MAX_BYTES = 20 * 1024 * 1024  # 20 MB inline limit

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 300 * 1024 * 1024  # 300 MB total (AI mode batch uploads)


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
# Routes — general mode
# ──────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html",
                           ocr_available=OCR_AVAILABLE,
                           gemini_available=GEMINI_AVAILABLE)


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


# ──────────────────────────────────────────────
# Routes — AI mode (Gemini)
# ──────────────────────────────────────────────
PROMPT_AI_PLAIN = """你是一個專業的資料輸入員。請將這份圖片或 PDF 中的表格轉換為純文字資料。

【嚴格規則】
1. 每一欄之間，請使用 "###" 作為分隔符號。
2. 每一列資料換一行。
3. 第一行必須是表頭。
4. 不要輸出任何 Markdown 標記，只要純文字。
5. 金額請保留千分位符號，不要隨意移除。
6. 若遇到跨頁，請自動合併。
7. 底部若有付款條件、稅金等資訊，請整理在表格最下方。"""

PROMPT_AI_TRANSLATE = """你是一個專業的雙語資料輸入員。請將這份圖片或 PDF 中的表格轉換為純文字資料，並附上中文翻譯。

【嚴格規則】
1. 每一欄之間，請使用 "###" 作為分隔符號。
2. 每一列資料換一行，第一行必須是表頭。
3. 請以「儲存格」為單位逐一判斷，不要只看欄位標題或其他列的內容：只要某個儲存格包含非中文的文字說明（例如店名、地址、品名、付款方式、備註等），就必須在該欄位右側緊接著新增一欄，欄名為「原欄名_中文」，填入該儲存格的繁體中文翻譯。
4. 只要同一欄中有任何一列符合規則 3，該欄就必須建立對應的「_中文」欄位；同一欄中屬於純數字、金額、日期、編號的儲存格，其對應的「_中文」欄位留空即可，不需要翻譯。
5. 特別注意：店名、分店名稱、付款方式（如信用卡、現金）等即使出現在看起來像「數值欄」的位置，仍然是文字內容，必須翻譯，不可因為欄位標題像數字欄而跳過。
6. 每一列的欄位數量必須一致（包含新增的翻譯欄位）。
7. 不要輸出任何 Markdown 標記，只要純文字。
8. 金額請保留千分位符號，不要隨意移除。
9. 若遇到跨頁，請自動合併。
10. 底部若有付款條件、稅金等資訊，請整理在表格最下方的列，文字部分同樣依規則 3、4、5 判斷是否翻譯。"""


def _gemini_mime(filename: str):
    fname = filename.lower()
    if fname.endswith(".pdf"):
        return "application/pdf"
    elif fname.endswith((".jpg", ".jpeg")):
        return "image/jpeg"
    elif fname.endswith(".png"):
        return "image/png"
    return None


def _excel_safe_sheet_name(name: str, used: set) -> str:
    invalid = "[]:*?/\\"
    clean = "".join(c for c in name if c not in invalid).strip() or "Sheet"
    clean = clean[:31]
    base, i = clean, 1
    while clean in used:
        suffix = f"_{i}"
        clean = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(clean)
    return clean


def _extract_invoice(file_bytes: bytes, mime_type: str, translate: bool) -> pd.DataFrame:
    model = genai.GenerativeModel("gemini-2.5-flash")
    prompt = PROMPT_AI_TRANSLATE if translate else PROMPT_AI_PLAIN
    response = model.generate_content([{"mime_type": mime_type, "data": file_bytes}, prompt])

    raw = response.text.replace("```csv", "").replace("```", "").strip()
    lines = [ln for ln in raw.split("\n") if ln.strip()]
    if not lines:
        raise ValueError("no_data")

    headers = [h.strip() for h in lines[0].split("###")]
    rows = []
    for line in lines[1:]:
        row = [c.strip() for c in line.split("###")]
        if len(row) < len(headers):
            row += [""] * (len(headers) - len(row))
        elif len(row) > len(headers):
            row = row[:len(headers)]
        rows.append(row)

    if not rows:
        raise ValueError("no_data")

    return pd.DataFrame(rows, columns=headers)


@app.route("/detect-language", methods=["POST"])
def detect_language_route():
    if not GEMINI_AVAILABLE:
        return jsonify({"language": "zh"})

    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file"}), 400

    file_bytes = f.read()
    if len(file_bytes) > GEMINI_MAX_BYTES:
        return jsonify({"language": "zh"})

    mime_type = _gemini_mime(f.filename)
    if not mime_type:
        return jsonify({"language": "zh"})

    try:
        model = genai.GenerativeModel("gemini-2.5-flash")
        prompt = ("請判斷這份文件中表格內容主要使用的語言。"
                  "只回答一個代碼，不要有其他文字或標點：zh／en／ja／ko／other")
        response = model.generate_content([{"mime_type": mime_type, "data": file_bytes}, prompt])
        code = response.text.strip().lower()
        if code not in ("zh", "en", "ja", "ko", "other"):
            code = "other"
        return jsonify({"language": code})
    except Exception:
        return jsonify({"language": "zh"})


@app.route("/convert-ai", methods=["POST"])
def convert_ai_route():
    if not GEMINI_AVAILABLE:
        return jsonify({"error": "AI 模式未啟用（缺少 GEMINI_API_KEY）"}), 503

    files = request.files.getlist("file")
    if not files:
        return jsonify({"error": "No file"}), 400

    translate = request.form.get("translate") == "true"
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    # ── Optional: merge into an existing Excel file ──
    existing_sheets = {}
    existing_label = None
    existing_file = request.files.get("existing_excel")
    if existing_file and existing_file.filename:
        try:
            existing_sheets = pd.read_excel(BytesIO(existing_file.read()), sheet_name=None)
            existing_label = existing_file.filename.rsplit(".", 1)[0]
        except Exception:
            existing_sheets = {}
            existing_label = None

    jobs = []
    for f in files:
        file_bytes = f.read()
        if len(file_bytes) > GEMINI_MAX_BYTES:
            jobs.append((f.filename, None, "too_large"))
            continue
        mime_type = _gemini_mime(f.filename)
        if not mime_type:
            jobs.append((f.filename, None, "bad_format"))
            continue
        jobs.append((f.filename, file_bytes, mime_type))

    def _run(job):
        filename, file_bytes, mime_type = job
        if file_bytes is None:
            return filename, None, mime_type  # reason stored in mime_type slot
        try:
            return filename, _extract_invoice(file_bytes, mime_type, translate), None
        except Exception as e:
            return filename, None, ("429" if "429" in str(e) else "no_data")

    results, failures = {}, []
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(4, len(jobs))) as ex:
        for filename, df, reason in ex.map(_run, jobs):
            if df is not None:
                results[filename] = df
            else:
                failures.append((filename, reason))

    if not results:
        return jsonify({"error": "no_data"}), 422

    used_sheet_names = set(existing_sheets.keys())
    sheets = {}
    combined_parts = []
    for filename, df in results.items():
        sheet_name = _excel_safe_sheet_name(filename.rsplit(".", 1)[0], used_sheet_names)
        sheets[sheet_name] = df
        tagged = df.copy()
        tagged.insert(0, "來源檔案", filename)
        combined_parts.append(tagged)

    new_combined = (pd.concat(combined_parts, ignore_index=True, sort=False)
                    if len(combined_parts) > 1 else combined_parts[0])

    merge_note = None
    if existing_sheets:
        all_sheets = dict(existing_sheets)
        if "全部明細" in all_sheets:
            append_target = "全部明細"
        elif len(existing_sheets) == 1:
            append_target = next(iter(existing_sheets.keys()))
        else:
            append_target = None

        if append_target:
            old_df = all_sheets.pop(append_target)
            if "來源檔案" not in old_df.columns:
                old_df = old_df.copy()
                old_df.insert(0, "來源檔案", "(原有資料)")
            all_sheets["全部明細"] = pd.concat([old_df, new_combined], ignore_index=True, sort=False)
        else:
            all_sheets["全部明細（新增）"] = new_combined
            merge_note = "原檔案有多個分頁且找不到「全部明細」，新資料已新增為獨立分頁"
        all_sheets.update(sheets)
        filename_out = f"{existing_label}_更新_{ts}.xlsx"
    elif len(results) > 1:
        all_sheets = {"全部明細": new_combined, **sheets}
        filename_out = f"批次轉換_{len(results)}筆_{ts}.xlsx"
    else:
        all_sheets = sheets
        only_name = next(iter(results.keys())).rsplit(".", 1)[0]
        filename_out = f"{only_name}_{ts}.xlsx"

    excel_bytes = build_excel(all_sheets)
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    resp = send_file(BytesIO(excel_bytes), mimetype=mime,
                     as_attachment=True, download_name=filename_out)
    resp.headers["X-Success-Count"] = str(len(results))
    resp.headers["X-Total-Count"] = str(len(jobs))
    if failures:
        resp.headers["X-Convert-Warnings"] = ";".join(
            f"{quote(fn)}:{reason}" for fn, reason in failures)
    if merge_note:
        resp.headers["X-Merge-Note"] = quote(merge_note)
    return resp


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
