import gradio as gr
import pandas as pd

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
        # 1. 讀取出貨明細表 (標頭在第二列 header=1)
        df_ship = pd.read_excel(shipping_file.name, header=1)
        
        # 2. 讀取包裝重量表 (標頭在第一列 header=0)
        df_weight = pd.read_excel(weight_file.name, header=0)
        
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
                
        # 8. 匯出成全新的 Excel 檔案
        output_path = "出貨明細_已完成重量核對.xlsx"
        df_merged.to_excel(output_path, index=False)
        
        success_count = df_merged['重量(箱)'].notna().sum()
        total_count = len(df_merged)
        
        return output_path, f"🎉 【終極防呆版】處理完成！\n📊 總共比對 {total_count} 筆資料，其中【{success_count} 筆】已完美消除「/箱」字尾差異並成功抓到重量！\n✨ 欄位名稱已就緒。"
        
    except Exception as e:
        return None, f"❌ 程式執行發生錯誤，原因：{str(e)}"

# --------------------------------------------------
# 【視覺 UI 介面設計】
# --------------------------------------------------
with gr.Blocks(title="出貨明細重量自動核對工具") as demo:
    gr.Markdown("# 🚚 出貨明細與包裝重量自動核對系統 (終極防呆版)")
    gr.Markdown("此版本會自動相容『1R/箱』與『1R』的寫法差異，確保兩邊代碼完美對齊。")
    
    with gr.Row():
        file_ship = gr.File(label="請放入【出貨明細 Excel】(主要欄位在第 2 列)", file_types=[".xlsx", ".xls"])
        file_weight = gr.File(label="請放入【包裝重量 Excel】(主要欄位在第 1 列)", file_types=[".xlsx", ".xls"])
        
    btn = gr.Button("🚀 開始自動比對並計算重量", variant="primary")
    
    with gr.Row():
        status_msg = gr.Textbox(label="系統處理狀態通知", lines=3)
        output_file = gr.File(label="✨ 點擊此處下載自動計算結果")
        
    btn.click(
        fn=process_shipping_data, 
        inputs=[file_ship, file_weight], 
        outputs=[output_file, status_msg]
    )

demo.launch(server_port=7861, share=False)