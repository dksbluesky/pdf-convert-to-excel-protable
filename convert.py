import streamlit as st
import pdfplumber
import pandas as pd
import json
from io import BytesIO

# OCR support — needs Tesseract + poppler installed (or packages.txt on Streamlit Cloud)
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

# ──────────────────────────────────────────────
# Page config
# ──────────────────────────────────────────────
st.set_page_config(page_title="萬用 PDF 轉換器", page_icon="📑", layout="wide")
st.title("📑 萬用 PDF 轉換器")
st.markdown("上傳 PDF，選擇格式，預覽後下載。支援文字型與掃描版 PDF（OCR）。")

# ──────────────────────────────────────────────
# Sidebar options
# ──────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 設定")

    output_format = st.radio(
        "輸出格式",
        ["Excel (.xlsx)", "Word (.docx)", "Markdown (.md)", "JSON (.json)"],
        index=0,
    )

    st.divider()

    use_ocr = False
    if OCR_AVAILABLE:
        use_ocr = st.checkbox("🔍 OCR 模式（掃描版 PDF）", value=False)
        if use_ocr:
            st.info(
                "OCR 模式會將每頁轉成影像再辨識文字。"
                "表格結構可能不完整，但至少能取得文字內容。"
            )
    else:
        st.warning(
            "**OCR 功能未啟用**\n\n"
            "需要系統安裝：\n"
            "- Tesseract OCR\n"
            "- poppler\n\n"
            "並執行：\n"
            "```\npip install pytesseract pdf2image\n```"
        )

# ──────────────────────────────────────────────
# Extraction helpers
# ──────────────────────────────────────────────
def extract_tables(pdf_bytes: bytes) -> dict:
    """Extract tables from a text-based PDF. Returns {page_label: DataFrame}."""
    result = {}
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        total = len(pdf.pages)
        bar = st.progress(0, text="掃描頁面中...")
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
                    result[f"Page {i + 1}"] = pd.concat(dfs, ignore_index=True)
            bar.progress((i + 1) / total, text=f"處理第 {i + 1} / {total} 頁...")
        bar.empty()
    return result


def extract_ocr(pdf_bytes: bytes) -> dict:
    """Convert each PDF page to image and OCR it. Returns {page_label: text}."""
    images = convert_from_bytes(pdf_bytes, dpi=300)
    result = {}
    bar = st.progress(0, text="OCR 辨識中...")
    for i, img in enumerate(images):
        # Try Traditional Chinese + English; falls back gracefully if lang pack missing
        try:
            text = pytesseract.image_to_string(img, lang="chi_tra+eng")
        except pytesseract.TesseractError:
            text = pytesseract.image_to_string(img)
        if text.strip():
            result[f"Page {i + 1}"] = text.strip()
        bar.progress((i + 1) / len(images), text=f"OCR 第 {i + 1} / {len(images)} 頁...")
    bar.empty()
    return result

# ──────────────────────────────────────────────
# Build output helpers
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
        if is_ocr:
            parts.append(content)
        else:
            parts.append(content.to_markdown(index=False))
        parts.append("\n")
    return "\n".join(parts)


def build_json(data: dict, is_ocr: bool) -> str:
    if is_ocr:
        payload = data
    else:
        payload = {page: df.to_dict(orient="records") for page, df in data.items()}
    return json.dumps(payload, ensure_ascii=False, indent=2)

# ──────────────────────────────────────────────
# Main UI
# ──────────────────────────────────────────────
uploaded = st.file_uploader("📂 上傳 PDF", type=["pdf"])

if uploaded:
    pdf_bytes = uploaded.read()
    base_name = uploaded.name.rsplit(".", 1)[0]
    timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")

    st.divider()

    with st.spinner("處理中，請稍候..."):
        if use_ocr:
            data = extract_ocr(pdf_bytes)
            is_ocr = True
        else:
            data = extract_tables(pdf_bytes)
            is_ocr = False

    if not data:
        if not use_ocr:
            st.warning(
                "⚠️ 找不到任何表格結構。\n\n"
                "若為掃描版 PDF，請在左側勾選 **OCR 模式** 再重試。"
            )
        else:
            st.warning("⚠️ OCR 未能辨識出任何文字，請確認 PDF 影像品質。")
        st.stop()

    st.success(f"✅ 成功處理 {len(data)} 頁")

    # ── Build output & preview ─────────────────
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

    # ── Download ───────────────────────────────
    st.divider()
    st.download_button(
        label=f"📥 確認並下載 {fmt}",
        data=file_bytes,
        file_name=file_name,
        mime=mime,
        type="primary",
        use_container_width=True,
    )

else:
    st.info("👆 請上傳 PDF 開始轉換")
