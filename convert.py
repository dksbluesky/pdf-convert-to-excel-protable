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
        # 使用 BytesIO 在記憶體中建立 Excel
        output = BytesIO()
        pages_converted = 0
        
        # 轉換核心邏輯
        # 1. 建立 ExcelWriter 物件
        with pdfplumber.open(uploaded_file) as pdf, pd.ExcelWriter(output, engine='openpyxl') as writer:
            total_pages = len(pdf.pages)
            progress_bar = st.progress(0)
            
            for i, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                
                if tables:
                    page_df_list = []
                    for table in tables:
                        # 轉為文字格式，並清洗 None
                        df = pd.DataFrame(table).astype(str)
                        df = df.replace("None", "")
                        page_df_list.append(df)
                    
                    if page_df_list:
                        page_df = pd.concat(page_df_list, ignore_index=True)
                        sheet_name = f"Page_{i+1}"
                        page_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                        pages_converted += 1
                
                progress_bar.progress((i + 1) / total_pages)
        
        # =======================================================
        # 關鍵修正：這裡的程式碼必須在 `with` 區塊結束之後執行
        # 確保 ExcelWriter 已經 .close() 並將資料完全寫入 output
        # =======================================================
        
        if pages_converted > 0:
            st.success(f"✅ 轉換完成！共成功處理 {pages_converted} 頁表格。")
            
            # 準備下載：將游標移回檔案開頭
            output.seek(0)
            
            # 設定檔名
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
            st.warning("⚠️ 掃描了整份文件，但找不到任何像表格的結構。")
            
    except Exception as e:
        st.error(f"❌ 發生錯誤：{e}")

else:
    st.info("👆 請上傳 PDF 以開始轉換")
