import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="高鐵 PDF 轉 Excel 工具", page_icon="📂")
st.title("📂 高鐵 PDF 轉 Excel 雲端版")
st.markdown("手機專用：上傳 PDF，轉換後下載 Excel，再拿去查詢系統使用。")

# 1. 檔案上傳
uploaded_file = st.file_uploader("請上傳 Download.pdf", type=["pdf"])

if uploaded_file is not None:
    st.info("正在讀取並轉換中，請稍候...")
    
    try:
        # 使用 BytesIO 在記憶體中建立 Excel 檔，不存到硬碟
        output = BytesIO()
        
        with pdfplumber.open(uploaded_file) as pdf, pd.ExcelWriter(output, engine='openpyxl') as writer:
            has_tables = False
            
            # 進度條
            progress_bar = st.progress(0)
            total_pages = len(pdf.pages)
            
            for i, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                
                if tables:
                    has_tables = True
                    page_df_list = []
                    for table in tables:
                        # 強制轉成文字格式，避免錯誤
                        df = pd.DataFrame(table).astype(str)
                        page_df_list.append(df)
                    
                    if page_df_list:
                        page_df = pd.concat(page_df_list, ignore_index=True)
                        
                        # 判斷方向
                        text = page.extract_text() or ""
                        sheet_name = f"Page_{i+1}"
                        if "Southbound" in text or "南下" in text:
                            sheet_name = f"Page_{i+1}_南下"
                        elif "Northbound" in text or "北上" in text:
                            sheet_name = f"Page_{i+1}_北上"
                        
                        # 寫入 Excel
                        page_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                
                # 更新進度
                progress_bar.progress((i + 1) / total_pages)
            
            if has_tables:
                st.success("✅ 轉換成功！請點擊下方按鈕下載。")
                
                # 重置游標位置，準備下載
                output.seek(0)
                
                # 產生當前時間檔名
                timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
                file_name = f"thsr_schedule_{timestamp}.xlsx"
                
                st.download_button(
                    label="📥 下載 Excel 檔案",
                    data=output,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("⚠️ 在 PDF 中找不到表格，請確認檔案是否正確。")
                
    except Exception as e:
        st.error(f"發生錯誤：{e}")