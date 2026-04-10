import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# ==========================================
# 頁面設定
# ==========================================
st.set_page_config(page_title="萬用 PDF 轉 Excel 工具", page_icon="📑")
st.title("📑 萬用 PDF 轉 Excel 轉換器")
st.markdown("""
**長久需求專用版**：不限高鐵，任何 PDF 報表皆可嘗試轉換。
程式會自動偵測每一頁的表格，並將其存為 Excel 的不同工作表。
""")

# ==========================================
# 檔案上傳
# ==========================================
uploaded_file = st.file_uploader("📂 請上傳包含表格的 PDF 檔", type=["pdf"])

if uploaded_file is not None:
    st.info("🔄 正在分析文件結構，請稍候...")
    
    try:
        all_page_data = [] # 用來暫存每一頁抓到的 DataFrame 資訊
        
        with pdfplumber.open(uploaded_file) as pdf:
            total_pages = len(pdf.pages)
            progress_bar = st.progress(0)
            
            for i, page in enumerate(pdf.pages):
                # 這裡加入了 table_settings，對於表單類型的 PDF 較友善
                tables = page.extract_tables(table_settings={
                    "vertical_strategy": "lines", 
                    "horizontal_strategy": "lines",
                })
                
                if tables:
                    page_df_list = []
                    for table in tables:
                        df = pd.DataFrame(table).astype(str)
                        df = df.replace("None", "")
                        page_df_list.append(df)
                    
                    if page_df_list:
                        combined_df = pd.concat(page_df_list, ignore_index=True)
                        all_page_data.append((f"Page_{i+1}", combined_df))
                
                progress_bar.progress((i + 1) / total_pages)

        # =======================================================
        # 判斷是否有抓到資料，才執行 Excel 寫入
        # =======================================================
        if all_page_data:
            output = BytesIO()
            # 只有在確定有資料時，才開啟 ExcelWriter
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet_name, df in all_page_data:
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            
            st.success(f"✅ 轉換完成！共成功處理 {len(all_page_data)} 頁表格。")
            
            output.seek(0)
            timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
            original_name = uploaded_file.name.rsplit('.', 1)[0]
            file_name = f"{original_name}_converted_{timestamp}.xlsx"
            
            st.download_button(
                label="📥 下載 Excel 檔案",
                data=output,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ 掃描了整份文件，但找不到任何像表格的結構。這可能是因為 PDF 是掃描影像（圖片）而非文字格式。")
            
    except Exception as e:
        st.error(f"❌ 發生錯誤：{e}")

else:
    st.info("👆 請上傳 PDF 以開始轉換")
