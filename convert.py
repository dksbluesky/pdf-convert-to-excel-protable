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
        
        # 轉換核心邏輯
        with pdfplumber.open(uploaded_file) as pdf, pd.ExcelWriter(output, engine='openpyxl') as writer:
            total_pages = len(pdf.pages)
            pages_converted = 0
            progress_bar = st.progress(0)
            
            for i, page in enumerate(pdf.pages):
                # 嘗試抓取該頁面所有的表格
                tables = page.extract_tables()
                
                if tables:
                    # 將這一頁找到的所有表格合併 (有些頁面可能有多個表格)
                    page_df_list = []
                    for table in tables:
                        # 全部轉為文字格式，避免數字/日期格式判讀錯誤
                        df = pd.DataFrame(table).astype(str)
                        
                        # 簡單清洗：把 None 轉為空字串
                        df = df.replace("None", "")
                        page_df_list.append(df)
                    
                    if page_df_list:
                        # 合併該頁所有小表格
                        page_df = pd.concat(page_df_list, ignore_index=True)
                        
                        # 命名工作表：Page_1, Page_2...
                        sheet_name = f"Page_{i+1}"
                        
                        # 寫入 Excel (不帶入預設的 0,1,2 索引與欄位名，保留原始樣貌)
                        page_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                        pages_converted += 1
                
                # 更新進度條
                progress_bar.progress((i + 1) / total_pages)
            
            # 結果判定
            if pages_converted > 0:
                st.success(f"✅ 轉換完成！共成功處理 {pages_converted} 頁表格。")
                
                # 準備下載
                output.seek(0)
                timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
                # 檔名使用 original_converted_時間.xlsx
                original_name = uploaded_file.name.rsplit('.', 1)[0]
                file_name = f"{original_name}_converted_{timestamp}.xlsx"
                
                st.download_button(
                    label="📥 下載 Excel 檔案",
                    data=output,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ 掃描了整份文件，但找不到任何像表格的結構。請確認 PDF 是否為掃描圖片檔。")
                st.caption("提示：此工具僅能處理「文字版 PDF」，若是「照片/掃描檔」需要使用 OCR 技術。")
                
    except Exception as e:
        st.error(f"❌ 發生錯誤：{e}")

else:
    st.info("👆 請上傳 PDF 以開始轉換")
