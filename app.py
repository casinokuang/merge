import streamlit as st
import pandas as pd
import io

# --- 頁面設定 ---
st.set_page_config(page_title="布料索引自動比對系統", layout="wide")

st.title("🧵 布料索引自動比對系統 (Web版)")
st.markdown("""
此工具將執行原本 Excel VBA `MergeAndCompareWithFabricIndex` 的功能：
1. **合併** 主檔的 A欄 與 D欄。
2. **比對** `Fabric name index` (A欄)。
3. **填寫** 結果至 E欄 (若比對成功則標記黃色)。
4. **保持 H 欄為數字格式**。
""")

# --- 側邊欄：檔案上傳區 ---
st.sidebar.header("📂 檔案上傳區")
uploaded_main = st.sidebar.file_uploader("1. 上傳主工作表 (需處理的檔案)", type=["xlsx", "xlsm"])
uploaded_index = st.sidebar.file_uploader("2. 上傳 Fabric name index (索引檔)", type=["xlsx", "xlsm"])

# --- 輔助函式：強力清洗鍵值 ---
def clean_key_func(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip().upper()
    if s.endswith(".0"):
        s = s[:-2]
    if s == "NAN":
        return ""
    return s

# --- 核心邏輯函數 ---
def process_data(main_df, index_df):
    # 1. 建立索引字典
    index_keys = index_df.iloc[:, 0].apply(clean_key_func)
    index_vals = index_df.iloc[:, 1]
    index_dict = dict(zip(index_keys, index_vals))
    
    # 2. 準備主檔數據
    df_result = main_df.copy()
    
    # --- 關鍵修正：將 H 欄 (Index 7) 轉回數字格式 ---
    # errors='coerce' 會將無法轉換的文字變為 NaN，再用 fillna(0) 補齊
    if df_result.shape[1] >= 8:
        df_result.iloc[:, 7] = pd.to_numeric(df_result.iloc[:, 7], errors='coerce').fillna(0)
    
    # 3. 執行合併與比對邏輯 (A欄 + D欄)
    main_keys = df_result.iloc[:, 0].apply(clean_key_func) + df_result.iloc[:, 3].apply(clean_key_func)
    
    new_e_column = []
    highlight_mask = [] 
    
    for idx, key in enumerate(main_keys):
        if key in index_dict:
            new_e_column.append(index_dict[key])
            highlight_mask.append(True)
        else:
            new_e_column.append(key)
            highlight_mask.append(False)
            
    # 寫入 E 欄 (Index 4)
    while df_result.shape[1] < 5:
        df_result[f'Col_{df_result.shape[1]}'] = None
    df_result.iloc[:, 4] = new_e_column
        
    return df_result, highlight_mask

# --- Excel 匯出函式 ---
def convert_df_to_excel_with_highlight(df, mask):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Result')
        
        workbook = writer.book
        worksheet = writer.sheets['Result']
        
        # 定義黃色格式
        yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
        
        # 遍歷 mask 上色 E 欄
        for idx, is_match in enumerate(mask):
            if is_match:
                value_to_write = df.iloc[idx, 4]
                if pd.isna(value_to_write):
                    value_to_write = ""
                worksheet.write(idx + 1, 4, value_to_write, yellow_format)
                
        # 額外確保 H 欄在 Excel 中的格式 (Column index 7)
        # 如果需要特定的小數點位數，可以在此設定
        num_format = workbook.add_format({'num_format': '#,##0.00'})
        worksheet.set_column(7, 7, None, num_format)

    output.seek(0)
    return output.getvalue()

# --- 主程式執行區 ---
if uploaded_main and uploaded_index:
    try:
        # 讀取時主檔仍用 str 以利 A, D 欄比對，但在 process_data 中會把 H 轉回數字
        df_main = pd.read_excel(uploaded_main, header=0, dtype=str)
        df_index = pd.read_excel(uploaded_index, header=0, dtype=str)
        
        st.success(f"✅ 檔案讀取成功！準備處理 {len(df_main)} 筆資料。")
        
        if st.button("🚀 執行合併與比對"):
            with st.spinner('正在處理中...'):
                result_df, mask = process_data(df_main, df_index)
                
                st.info(f"📊 處理完成：{sum(mask)} 筆比對成功。H 欄已轉換為數字格式。")
                
                # 預覽
                st.subheader("結果預覽 (前 10 筆)")
                st.dataframe(result_df.head(10))
                
                # 下載
                excel_data = convert_df_to_excel_with_highlight(result_df, mask)
                
                st.download_button(
                    label="📥 下載 merge.xlsx",
                    data=excel_data,
                    file_name="merge.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"發生錯誤：{str(e)}")
else:
    st.info("👈 請從左側選單上傳兩個 Excel 檔案以開始使用。")
