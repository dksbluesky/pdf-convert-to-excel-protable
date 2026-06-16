import streamlit as st
import pdfplumber
import pandas as pd
import json
from io import BytesIO

try:
    from pdf2image import convert_from_bytes
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

FORMATS = ["Excel (.xlsx)", "Word (.docx)", "Markdown (.md)", "JSON (.json)"]

st.set_page_config(page_title="PDF 轉換器", page_icon="📄", layout="wide")

# Hide Streamlit branding
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

st.title("📄 PDF 轉換器")
st.caption("上傳 PDF，自動偵測最佳格式並立即轉換。支援文字型與掃描版（OCR）。")

# ──────────────────────────────────────────────
# Analysis (cached — runs once per file)
# ──────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def analyze_pdf(pdf_bytes: bytes) -> dict:
    """Sample up to 5 pages to recommend the best output format."""
    table_pages, text_pages, total_pages = 0, 0, 0
    try:
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            total_pages = len(pdf.pages)
            for page in pdf.pages[:min(5, total_pages)]:
                if page.extract_tables():
                    table_pages += 1
                if (page.extract_text() or "").strip():
                    text_pages += 1
    except Exception:
        pass

    if text_pages == 0 and table_pages == 0:
        return {
            "format": "Word (.docx)",
            "badge": "🔍 掃描版 PDF",
            "reason": "未偵測到可擷取的文字，建議啟用 OCR 再轉為 Word。",
            "use_ocr": True,
        }
    elif table_pages >= text_pages * 0.6:
        return {
            "format": "Excel (.xlsx)",
            "badge": "📊 表格型 PDF",
            "reason": f"前 {min(5,total_pages)} 頁中有 {table_pages} 頁含表格，最適合匯出為 Excel。",
            "use_ocr": False,
        }
    elif table_pages > 0:
        return {
            "format": "Word (.docx)",
            "badge": "📝 混合型 PDF",
            "reason": f"文字與表格混合（{table_pages} 頁有表格），建議轉為 Word 保留排版。",
            "use_ocr": False,
        }
    else:
        return {
            "format": "Markdown (.md)",
            "badge": "📄 純文字 PDF",
            "reason": "純文字內容，Markdown 最輕量且易讀。",
            "use_ocr": False,
        }


# ──────────────────────────────────────────────
# Extraction (cached — re-runs only when file or mode changes)
# ──────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def extract_tables(pdf_bytes: bytes) -> dict:
    result = {}
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables(table_settings={
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
            })
            if tables:
                dfs = []
                for t in tables:
                    cleaned = [[c if c is not None else "" for c in row] for row in t]
                    if len(cleaned) > 1:
                        df = pd.DataFrame(cleaned[1:], columns=cleaned[0])
                    else:
                        df = pd.DataFrame(cleaned)
                    dfs.append(df.astype(str).replace("None", ""))
                if dfs:
                    result[f"Page {i+1}"] = pd.concat(dfs, ignore_index=True)
    return result


@st.cache_data(show_spinner=False)
def extract_ocr(pdf_bytes: bytes) -> dict:
    images = convert_from_bytes(pdf_bytes, dpi=300)
    result = {}
    for i, img in enumerate(images):
        try:
            text = pytesseract.image_to_string(img, lang="chi_tra+eng")
        except pytesseract.TesseractError:
            text = pytesseract.image_to_string(img)
        if text.strip():
            result[f"Page {i+1}"] = text.strip()
    return result


# ──────────────────────────────────────────────
# Build output files
# ──────────────────────────────────────────────
def build_excel(data: dict) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sheet, df in data.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
    return out.getvalue()


def build_word(data: dict, is_ocr: bool) -> bytes:
    doc = Document()
    doc.add_heading("PDF 轉換結果", level=0)
    for page, content in data.items():
        doc.add_heading(page, level=1)
        if is_ocr:
            doc.add_paragraph(content)
        else:
            df: pd.DataFrame = content
            tbl = doc.add_table(rows=1 + len(df), cols=len(df.columns))
            tbl.style = "Table Grid"
            for j, col in enumerate(df.columns):
                tbl.rows[0].cells[j].text = str(col)
            for i, row in df.iterrows():
                for j, val in enumerate(row):
                    tbl.rows[i + 1].cells[j].text = str(val)
            doc.add_paragraph()
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
# Main UI
# ──────────────────────────────────────────────
uploaded = st.file_uploader("📂 上傳 PDF", type=["pdf"])

if uploaded:
    pdf_bytes = uploaded.read()
    base_name = uploaded.name.rsplit(".", 1)[0]
    timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")

    # Analyze first (cached, fast)
    with st.spinner("分析 PDF 結構中..."):
        rec = analyze_pdf(pdf_bytes)

    rec_index = FORMATS.index(rec["format"])

    # Sidebar — show recommendation + allow override
    with st.sidebar:
        st.header("⚙️ 設定")
        st.success(f"**{rec['badge']}**\n\n{rec['reason']}")
        output_format = st.radio("輸出格式", FORMATS, index=rec_index)

        st.divider()
        use_ocr = rec.get("use_ocr", False)
        if OCR_AVAILABLE:
            use_ocr = st.checkbox(
                "🔍 OCR 模式（掃描版 PDF）",
                value=rec.get("use_ocr", False),
                help="將每頁轉成影像再辨識文字。表格結構可能不完整。"
            )
        else:
            if rec.get("use_ocr"):
                st.warning("⚠️ 偵測到掃描版，但 OCR 套件未安裝。")

    # Extract (cached)
    st.divider()
    with st.spinner("轉換中，請稍候..."):
        if use_ocr and OCR_AVAILABLE:
            data = extract_ocr(pdf_bytes)
            is_ocr = True
        else:
            data = extract_tables(pdf_bytes)
            is_ocr = False

    if not data:
        if not use_ocr:
            st.warning("⚠️ 找不到表格結構。若為掃描版 PDF，請在左側勾選 **OCR 模式** 再試。")
        else:
            st.warning("⚠️ OCR 未能辨識出文字，請確認 PDF 影像品質。")
        st.stop()

    st.success(f"✅ 成功處理 {len(data)} 頁")

    # Preview
    fmt = output_format
    st.subheader("👁️ 預覽")

    if fmt == "Excel (.xlsx)":
        for page, df in data.items():
            with st.expander(page, expanded=True):
                st.dataframe(df, use_container_width=True)
        file_bytes = build_excel(data)
        file_name = f"{base_name}_{timestamp}.xlsx"
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    elif fmt == "Word (.docx)":
        for page, content in data.items():
            with st.expander(page, expanded=True):
                st.text(content) if is_ocr else st.dataframe(content, use_container_width=True)
        file_bytes = build_word(data, is_ocr)
        file_name = f"{base_name}_{timestamp}.docx"
        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    elif fmt == "Markdown (.md)":
        md_text = build_markdown(data, is_ocr)
        with st.expander("Markdown 預覽", expanded=True):
            st.markdown(md_text)
        file_bytes = md_text.encode("utf-8")
        file_name = f"{base_name}_{timestamp}.md"
        mime = "text/markdown"

    elif fmt == "JSON (.json)":
        json_str = build_json(data, is_ocr)
        with st.expander("JSON 預覽", expanded=True):
            st.json(json.loads(json_str))
        file_bytes = json_str.encode("utf-8")
        file_name = f"{base_name}_{timestamp}.json"
        mime = "application/json"

    # Download
    st.divider()
    st.download_button(
        label=f"📥 下載 {fmt}",
        data=file_bytes,
        file_name=file_name,
        mime=mime,
        type="primary",
        use_container_width=True,
    )

else:
    with st.sidebar:
        st.header("⚙️ 設定")
        st.info("上傳 PDF 後將自動偵測最佳格式。")
        st.radio("輸出格式", FORMATS, index=0, disabled=True)
    st.info("👆 請上傳 PDF 開始轉換")
