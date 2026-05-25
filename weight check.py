import streamlit as st
import pandas as pd
import io

# 設定網頁標題與分頁標籤（這是 Streamlit 的優勢，可以做得很像官網）
st.set_page_config(page_title="出貨重量自動核對工具", page_icon="🚚", layout="wide")

def clean_and_format_key(series):
    """
    【專用大清洗工具】
    移除所有空格、去掉小數點(.0)、並且把字尾的「/箱」或「箱」通通拔掉，確保完全對齊！
    """
    return (series.astype(str)
            .str.strip()
            .str.replace(r'\.0$', '', regex=True)
            .str.replace(r'\s+', '', regex=True)
            .str.replace(r'/箱$', '', regex=True)  # 拔掉字尾的 /箱
            .str.replace(r'箱$', '', regex=True))   # 拔掉字尾的 箱

def process_shipping_data(shipping_file, weight_file):
    try:
        # 💡 Streamlit 的優化：直接讀取上傳的檔案物件，移除原本的 .name（雲端才不會出錯）
        # 1. 讀取出貨明細表 (標頭在第二列 header=1)
        df_ship = pd.read_excel(shipping_file, header=1)
        
        # 2. 讀取包裝重量表 (標頭在第一列 header=0)
        df_weight = pd.read_excel(weight_file, header=0)
        
        # 清理兩張表的欄位名稱空格
        df_ship.columns = df_ship.columns.str.strip()
        df_weight.columns = df_weight.columns.str.strip()
        
        # 必要欄位檢查
        required_ship = ['編號', '卷/箱', '箱數']
        required_weight = ['編號', '卷/箱', '重量(箱)']
        
        for col in required_ship:
            if col not in df_ship.columns:
                return None, f"❌ 出貨明細表 缺少必要欄位：【{col}】"
        for col in required_weight:
            if col not in df_weight.columns:
                return None, f"❌ 包裝重量表 缺少必要欄位：【{col}】"

        # 💡 為了在輸出的 J 欄維持你原本想要的漂亮外觀，我們保留格式做顯示
        df_ship['編號_卷/箱'] = df_ship['編號'].astype(str).str.strip() + "_" + df_ship['卷/箱'].astype(str).str.strip()

        # 3. 幕後悄悄進行「大去骨清洗」，用來當作真正對齊的鎖匙
        ship_no_clean = clean_and_format_key(df_ship['編號'])
        ship_box_clean = clean_and_format_key(df_ship['卷/箱'])
        df_ship['幕後對照Key'] = ship_no_clean + "_" + ship_box_clean
        
        weight_no_clean = clean_and_format_key(df_weight['編號'])
        weight_box_clean = clean_and_format_key(df_weight['卷/箱'])
        df_weight['幕後對照Key'] = weight_no_clean + "_" + weight_box_clean
        
        # 唯獨保留對照需要的欄位，並去除重複值
        df_weight_clean = df_weight[['幕後對照Key', '重量(箱)']].drop_duplicates(subset=['幕後對照Key'])
        
        # 4. 開始 VLOOKUP 比對對齊
        df_merged = pd.merge(df_ship, df_weight_clean, on='幕後對照Key', how='left')
        
        # 5. K 欄名字改為 "重量(箱)"
        df_merged['重量(箱)_y'] = df_merged['重量(箱)_y'] if '重量(箱)_y' in df_merged.columns else df_merged['重量(箱)']
        df_merged['重量(箱)'] = pd.to_numeric(df_merged['重量(箱)_y'], errors='coerce')
        
        # 6. L 欄名字改為 "總重量" = 重量(箱) * 箱數
        df_merged['箱數'] = pd.to_numeric(df_merged['箱數'], errors='coerce')
        df_merged['總重量'] = df_merged['重量(箱)'] * df_merged['箱數']
        
        # 7. 清理掉中間產生的暫存幕後工具欄位，保持表格乾淨
        drop_cols = ['幕後對照Key', '重量(箱)_x', '重量(箱)_y']
        for col in drop_cols:
            if col in df_merged.columns:
                df_merged = df_merged.drop(columns=[col])
                
        # 💡 雲端優化：將 Excel 存入記憶體快取，方便 Streamlit 提供下載按鈕
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_merged.to_excel(writer, index=False)
        xlsx_data = output.getvalue()
        
        success_count = df_merged['重量(箱)'].notna().sum()
        total_count = len(df_merged)
        
        msg = f"🎉 【終極防呆版】處理完成！\n\n📊 總共比對 {total_count} 筆資料，其中【{success_count} 筆】已完美消除「/箱」字尾差異並成功抓到重量！"
        return xlsx_data, msg
        
    except Exception as e:
        return None, f"❌ 程式執行發生錯誤，原因：{str(e)}"

# --------------------------------------------------
# 【Streamlit 視覺 UI 介面設計】
# --------------------------------------------------
st.title("🚚 出貨明細與包裝重量自動核對系統 (Streamlit 雲端版)")
st.markdown("### ✨ 此版本會自動相容『1R/箱』與『1R』的寫法差異，確保兩邊代碼完美對齊。")
st.write("---")

# 建立左右兩欄，放上傳按鈕
col1, col2 = st.columns(2)

with col1:
    file_ship = st.file_uploader("請放入【出貨明細 Excel】(主要欄位在第 2 列)", type=["xlsx", "xls"])

with col2:
    file_weight = st.file_uploader("請放入【包裝重量 Excel】(主要欄位在第 1 列)", type=["xlsx", "xls"])

st.write("---")

# 當兩張表都丟進來後，顯示開始比對按鈕
if file_ship and file_weight:
    if st.button("🚀 開始自動比對並計算重量", type="primary", use_container_width=True):
        
        # 執行計算
        with st.spinner("系統正在後台大清洗並對齊資料中，請稍候..."):
            result_data, status_msg = process_shipping_data(file_ship, file_weight)
        
        if result_data:
            st.success("處理成功！")
            st.info(status_msg)
            
            # 製作下載按鈕
            st.download_button(
                label="📥 點擊此處下載自動計算結果 (.xlsx)",
                data=result_data,
                file_name="出貨明細_已完成重量核對.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error(status_msg)
else:
    st.warning("💡 請先在上方【同時上傳】兩份 Excel 檔案，系統按鈕就會出現囉！")
