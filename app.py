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
3. **填寫** 結果至 E欄 (若比對成功則填入對應值並**標記黃色**，否則填入合併字串)。
""")

# --- 側邊欄：檔案上傳區 ---
st.sidebar.header("📂 檔案上傳區")
uploaded_main = st.sidebar.file_uploader("1. 上傳主工作表 (需處理的檔案)", type=["xlsx", "xlsm"])
uploaded_index = st.sidebar.file_uploader("2. 上傳 Fabric name index (索引檔)", type=["xlsx", "xlsm"])

# --- 輔助函式：強力清洗鍵值 (解決 123 vs 123.0 問題) ---
def clean_key_func(val):
    """
    將任何輸入值轉為標準化的比對鍵值 (去除空白、轉大寫、處理浮點數)
    """
    if pd.isna(val) or val is None:
        return ""
    
    # 強制轉字串並去空白、轉大寫
    s = str(val).strip().upper()
    
    # 處理 Excel 讀取整數時可能出現的 .0 (例如 "123.0" -> "123")
    if s.endswith(".0"):
        s = s[:-2]
        
    # 處理 Pandas 讀取空值可能產生的 "NAN" 字串
    if s == "NAN":
        return ""
        
    return s

# --- 核心邏輯函數 ---
def process_data(main_df, index_df):
    # 1. 建立索引字典
    # 使用 clean_key_func 確保比對精準
    index_keys = index_df.iloc[:, 0].apply(clean_key_func)
    index_vals = index_df.iloc[:, 1]
    
    # 轉成字典 { 'KEY': 'Value' }
    index_dict = dict(zip(index_keys, index_vals))
    
    # 2. 準備主檔數據
    df_result = main_df.copy()
    
    # 3. 執行合併與比對邏輯 (A欄 + D欄)
    # VBA: mergeText = Cells(i, 1) & Cells(i, 4)
    main_keys = df_result.iloc[:, 0].apply(clean_key_func) + df_result.iloc[:, 3].apply(clean_key_func)
    
    # 建立結果列表與顏色標記列表
    new_e_column = []
    highlight_mask = [] # True = 要變黃色, False = 不變色
    
    for idx, key in enumerate(main_keys):
        if key in index_dict:
            # Match Found: 取出對應值
            new_e_column.append(index_dict[key])
            highlight_mask.append(True) # 標記為需要上色
        else:
            # No Match: 使用合併後的 Key
            new_e_column.append(key)
            highlight_mask.append(False)
            
    # 將結果寫入 E 欄 (Index 4)
    # 確保 DataFrame 至少有 5 欄
    while df_result.shape[1] < 5:
        df_result[f'Col_{df_result.shape[1]}'] = None
        
    df_result.iloc[:, 4] = new_e_column
        
    return df_result, highlight_mask

# --- Excel 匯出函式 (含黃色標記) ---
def convert_df_to_excel_with_highlight(df, mask):
    output = io.BytesIO()
    
    # 使用 XlsxWriter 引擎
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Result')
        
        workbook = writer.book
        worksheet = writer.sheets['Result']
        
        # 定義黃色格式 (對應 VBA: RGB(255, 255, 0))
        yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
        
        # 遍歷 mask，如果為 True 則將該列的 E 欄 (Column 4) 上色
        # ExcelWriter 的 row 0 是標題，所以資料從 row 1 開始
        # column 4 對應 Excel 的 E 欄
        for idx, is_match in enumerate(mask):
            if is_match:
                value_to_write = df.iloc[idx, 4]
                if pd.isna(value_to_write):
                    value_to_write = ""
                # 寫入儲存格並套用格式
                worksheet.write(idx + 1, 4, value_to_write, yellow_format)
                
    output.seek(0)
    return output.getvalue()

# --- 主程式執行區 ---
if uploaded_main and uploaded_index:
    try:
        # 讀取 Excel 檔案 (使用 dtype=str 以避免數字格式問題)
        df_main = pd.read_excel(uploaded_main, header=0, dtype=str)
        df_index = pd.read_excel(uploaded_index, header=0, dtype=str)
        
        st.success(f"✅ 檔案讀取成功！準備比對 {len(df_main)} 筆資料。")
        
        if st.button("🚀 執行合併與比對"):
            with st.spinner('正在處理中...'):
                # 1. 執行運算
                result_df, mask = process_data(df_main, df_index)
                
                # 2. 顯示統計
                match_count = sum(mask)
                st.info(f"📊 處理完成：共 {len(result_df)} 筆，其中 {match_count} 筆比對成功 (已標示為黃色)。")
                
                # 3. 網頁預覽 (模擬黃色底色)
                st.subheader("結果預覽 (E欄)")
                
                def highlight_rows(row):
                    if row.name < len(mask) and mask[row.name]:
                        return ['background-color: #FFFFE0'] * len(row)
                    return [''] * len(row)

                st.dataframe(result_df.head(10).style.apply(highlight_rows, axis=1))
                
                # 4. 產生並提供下載
                excel_data = convert_df_to_excel_with_highlight(result_df, mask)
                
                # --- 修改處：檔名設定為 merge.xlsx ---
                st.download_button(
                    label="📥 下載 merge.xlsx",
                    data=excel_data,
                    file_name="merge.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"發生錯誤：{str(e)}")
        st.warning("請確認 Excel 格式：主檔需有 A-E 欄，索引檔需有 A-B 欄。")

else:
    st.info("👈 請從左側選單上傳兩個 Excel 檔案以開始使用。")
