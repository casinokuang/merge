import streamlit as st
import pandas as pd
import io

# 設定網頁標題與寬度
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

# --- 核心邏輯函數 ---
def process_data(main_df, index_df):
    # 1. 建立索引字典 (Hash Map) - 對應 VBA 的 Fabric Index 讀取
    # 假設 Index 檔：A欄是 Key, B欄是 Value
    # 轉成字典 { 'Key1': 'Value1', 'Key2': 'Value2' } 以加速比對
    index_dict = dict(zip(
        index_df.iloc[:, 0].astype(str).str.strip(), 
        index_df.iloc[:, 1]
    ))
    
    # 2. 準備主檔數據
    # 複製一份以免影響原始數據
    df_result = main_df.copy()
    
    # 3. 執行合併與比對邏輯 (取代 VBA 的 For Loop)
    # VBA: mergeText = Cells(i, 1) & Cells(i, 4)
    # Pandas: 向量化操作，速度極快
    merge_keys = df_result.iloc[:, 0].astype(str).str.strip() + df_result.iloc[:, 3].astype(str).str.strip()
    
    # 建立結果列表與標記列表
    new_e_column = []
    highlight_mask = [] # 用來記錄哪一行需要變黃色
    
    for key in merge_keys:
        if key in index_dict:
            # Match Found: 取出對應值 (VBA: wsFabric.Cells(j, 2).Value)
            new_e_column.append(index_dict[key])
            highlight_mask.append(True) # 標記為需要上色
        else:
            # No Match: 使用合併字串 (VBA: wsCurrent.Cells(i, 5).Value = mergeText)
            new_e_column.append(key)
            highlight_mask.append(False)
            
    # 將結果寫入 E 欄 (Index 4)
    # 如果原始檔案沒有 E 欄，Pandas 會自動新增
    if df_result.shape[1] < 5:
        df_result['Result'] = new_e_column
    else:
        df_result.iloc[:, 4] = new_e_column
        
    return df_result, highlight_mask

# --- 主程式 ---
if uploaded_main and uploaded_index:
    try:
        # 讀取 Excel 檔案
        df_main = pd.read_excel(uploaded_main)
        df_index = pd.read_excel(uploaded_index)
        
        st.success(f"檔案讀取成功！主檔共 {len(df_main)} 筆，索引檔共 {len(df_index)} 筆。")
        
        if st.button("🚀 開始執行比對 (Run Merge & Compare)"):
            with st.spinner('正在處理中...'):
                # 執行運算
                result_df, mask = process_data(df_main, df_index)
                
                # --- 顯示預覽結果 ---
                st.subheader("📊 結果預覽")
                
                # 在網頁上模擬黃色底色顯示
                def highlight_rows(row):
                    # 取得該行的 index
                    idx = row.name 
                    if idx < len(mask) and mask[idx]:
                        return ['background-color: #FFFF00'] * len(row)
                    return [''] * len(row)

                st.dataframe(result_df.style.apply(highlight_rows, axis=1), use_container_width=True)
                
                # --- 產生下載檔案 (包含黃色底色) ---
                output = io.BytesIO()
                
                # 使用 XlsxWriter 引擎來寫入顏色格式
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='Processed_Data')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Processed_Data']
                    
                    # 定義黃色格式 (對應 VBA: RGB(255, 255, 0))
                    yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
                    
                    # 取得 E 欄的索引 (Excel 是 1-based, 但寫程式通常處理 Column index)
                    # 假設我們要標記整行，或者只標記 E 欄
                    # 這裡模擬 VBA：整行的 E 欄 (第 5 欄) 變色
                    
                    # 遍歷 mask，如果為 True 則將該列的 E 欄 (Column 4) 上色
                    # 注意：ExcelWriter 的 row 0 是標題，所以資料從 row 1 開始
                    for idx, is_match in enumerate(mask):
                        if is_match:
                            # 寫入該儲存格並套用格式
                            value_to_write = result_df.iloc[idx, 4] # E欄的值
                            # (Row, Col, Data, Format)
                            worksheet.write(idx + 1, 4, value_to_write, yellow_format)
                            
                output.seek(0)
                
                st.download_button(
                    label="📥 下載處理後的 Excel 檔案",
                    data=output,
                    file_name="processed_fabric_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"發生錯誤：{str(e)}")
        st.info("請確認上傳的 Excel 檔案格式是否正確 (主檔需有A-E欄，索引檔需有A-B欄)。")

else:
    st.info("👈 請從左側選單上傳兩個 Excel 檔案以開始使用。")