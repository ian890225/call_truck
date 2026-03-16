import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# 設定網頁標題與寬度
st.set_page_config(page_title="叫車排程系統", page_icon="🚛", layout="wide")

st.title("🚛自動叫車排程系統")
st.markdown("請在下方上傳最新的 **構件管制表** 與 **出貨聯絡單**，系統將自動為您產出今日與本週的叫車總表。")

# 建立左右兩個上傳區塊
col1, col2 = st.columns(2)
with col1:
    ctrl_files = st.file_uploader("📂 1. 請上傳「構件管制表」(可多選)", type=['xlsx', 'csv'], accept_multiple_files=True)
with col2:
    ship_files = st.file_uploader("📂 2. 請上傳「出貨聯絡單」(可多選)", type=['xlsx', 'csv'], accept_multiple_files=True)

if st.button("🚀 開始產生叫車排程", use_container_width=True, type="primary"):
    if not ctrl_files or not ship_files:
        st.warning("⚠️ 請確認「管制表」和「出貨單」都已經上傳喔！")
    else:
        with st.spinner("系統正在拼命讀取與運算中，請稍候..."):
            ctrl_dfs = []
            ship_dfs = []

            # 1. 處理管制表
            for file in ctrl_files:
                try:
                    # 判斷副檔名
                    if file.name.endswith('.csv'):
                        df_temp = pd.read_csv(file, nrows=20, header=None)
                        flat_str = ' '.join(df_temp.fillna('').astype(str).values.flatten())
                        if '構件編號' in flat_str and '組立日期' in flat_str:
                            header_idx = df_temp[df_temp.apply(lambda r: r.astype(str).str.contains('構件編號').any(), axis=1)].index
                            if len(header_idx) > 0:
                                file.seek(0)
                                df_full = pd.read_csv(file, skiprows=header_idx[0])
                                ctrl_dfs.append(df_full)
                    else:
                        xls_dict = pd.read_excel(file, sheet_name=None, header=None)
                        for sheet_name, df in xls_dict.items():
                            if df.empty: continue
                            head_df = df.head(20).fillna('').astype(str)
                            flat_str = ' '.join(head_df.values.flatten())
                            if '構件編號' in flat_str and '組立日期' in flat_str:
                                header_idx = head_df[head_df.apply(lambda r: r.str.contains('構件編號').any(), axis=1)].index
                                if len(header_idx) > 0:
                                    df_full = pd.read_excel(file, sheet_name=sheet_name, skiprows=header_idx[0])
                                    ctrl_dfs.append(df_full)
                except Exception as e:
                    st.error(f"讀取管制表時發生錯誤：{e}")

            # 2. 處理出貨聯絡單
            for file in ship_files:
                try:
                    if file.name.endswith('.csv'):
                        df_temp = pd.read_csv(file, nrows=20, header=None)
                        flat_str = ' '.join(df_temp.fillna('').astype(str).values.flatten())
                        if '實際交貨日期' in flat_str and '批號' in flat_str:
                            header_idx = df_temp[df_temp.apply(lambda r: r.astype(str).str.contains('實際交貨日期').any(), axis=1)].index
                            if len(header_idx) > 0:
                                file.seek(0)
                                df_full = pd.read_csv(file, skiprows=header_idx[0])
                                ship_dfs.append(df_full)
                    else:
                        xls_dict = pd.read_excel(file, sheet_name=None, header=None)
                        for sheet_name, df in xls_dict.items():
                            if df.empty: continue
                            head_df = df.head(20).fillna('').astype(str)
                            flat_str = ' '.join(head_df.values.flatten())
                            if '實際交貨日期' in flat_str and '批號' in flat_str:
                                header_idx = head_df[head_df.apply(lambda r: r.str.contains('實際交貨日期').any(), axis=1)].index
                                if len(header_idx) > 0:
                                    df_full = pd.read_excel(file, sheet_name=sheet_name, skiprows=header_idx[0])
                                    ship_dfs.append(df_full)
                except Exception as e:
                    st.error(f"讀取出貨單時發生錯誤：{e}")

            df_ctrl = pd.concat(ctrl_dfs, ignore_index=True) if ctrl_dfs else pd.DataFrame()
            df_ship = pd.concat(ship_dfs, ignore_index=True) if ship_dfs else pd.DataFrame()

            if df_ship.empty or df_ctrl.empty:
                st.error("❌ 未找到有效的資料，請確認上傳的檔案內容是否正確。")
            else:
                # 3. 清理資料與合併
                df_ship.columns = df_ship.columns.astype(str).str.strip()
                df_ship = df_ship[['批號', '實際交貨日期']].dropna(subset=['批號', '實際交貨日期'])
                df_ship['批號'] = df_ship['批號'].astype(str).str.strip()
                df_ship = df_ship[~df_ship['批號'].str.contains('批號', na=False)]
                df_ship = df_ship.drop_duplicates(subset=['批號'])
                df_ship = df_ship.rename(columns={'實際交貨日期': '工地需求日'})
                df_ship['工地需求日'] = pd.to_datetime(df_ship['工地需求日'], errors='coerce')
                df_ship = df_ship.dropna(subset=['工地需求日'])

                df_ctrl.columns = df_ctrl.columns.astype(str).str.strip()
                df_ctrl['批號'] = df_ctrl['批號'].astype(str).str.strip()
                df_merged = pd.merge(df_ctrl, df_ship, on='批號', how='inner')

                if df_merged.empty:
                    st.error("❌ 管制表與出貨單之間比對不到相同的批號！")
                else:
                    # 4. 日期推算邏輯
                    def calculate_loading_date(delivery_date):
                        if pd.isnull(delivery_date): return pd.NaT
                        loading_date = delivery_date - pd.Timedelta(days=1)
                        if loading_date.weekday() == 6: # 週日改週六
                            loading_date = delivery_date - pd.Timedelta(days=2)
                        return loading_date

                    def calculate_calling_date(loading_date):
                        if pd.isnull(loading_date): return pd.NaT
                        calling_date = loading_date - pd.Timedelta(days=1)
                        if calling_date.weekday() == 6: calling_date -= pd.Timedelta(days=2)
                        elif calling_date.weekday() == 5: calling_date -= pd.Timedelta(days=1)
                        return calling_date

                    df_merged['裝車日期'] = df_merged['工地需求日'].apply(calculate_loading_date)
                    df_merged['叫車日期'] = df_merged['裝車日期'].apply(calculate_calling_date)

                    # 塗裝廠商判斷
                    if '噴漆包商' not in df_merged.columns: df_merged['噴漆包商'] = ''
                    if '預計塗裝廠商' not in df_merged.columns:
                        df_merged['預計塗裝廠商'] = df_merged['一次預定廠商'] if '一次預定廠商' in df_merged.columns else ''

                    df_merged['噴漆包商'] = df_merged['噴漆包商'].astype(str).replace(['nan', 'NaN'], '').str.strip()
                    df_merged['預計塗裝廠商'] = df_merged['預計塗裝廠商'].astype(str).replace(['nan', 'NaN'], '').str.strip()
                    df_merged['塗裝廠商'] = np.where(df_merged['噴漆包商'] != '', df_merged['噴漆包商'], df_merged['預計塗裝廠商'])
                    df_merged['塗裝廠商'] = df_merged['塗裝廠商'].replace('', '未指定')

                    df_merged = df_merged.rename(columns={'工程案號': '工程編號', '節數': '節次'})

                    # 日期篩選 (今天與本週)
                    date_cols = ['組立日期', '焊接日期', '二檢日期', '噴漆施工日期', '叫車日期', '裝車日期', '工地需求日']
                    for c in date_cols:
                        if c in df_merged.columns:
                            df_merged[c] = pd.to_datetime(df_merged[c], errors='coerce')

                    # 針對伺服器時區問題，自動抓取當前台灣時間作為今天
                    today = pd.Timestamp.utcnow().tz_convert('Asia/Taipei').normalize().tz_localize(None)
                    start_of_week = today - pd.Timedelta(days=today.weekday())
                    end_of_week = start_of_week + pd.Timedelta(days=6)

                    df_today = df_merged[df_merged['叫車日期'] == today].copy()
                    df_week = df_merged[(df_merged['叫車日期'] >= start_of_week) & (df_merged['叫車日期'] <= end_of_week)].copy()

                    for df_target in [df_today, df_week]:
                        for c in date_cols:
                            if c in df_target.columns:
                                df_target[c] = df_target[c].dt.strftime('%Y/%m/%d').fillna('')

                    if df_today.empty and df_week.empty:
                        st.info("🎉 太棒了！今天與本週目前都沒有需要叫車的排程。")
                    else:
                        if df_today.empty:
                            st.info(f"今天 ({today.strftime('%Y/%m/%d')}) 沒有需叫車排程，以下為您產出「本週」的清單。")
                        else:
                            st.success(f"✅ 成功計算完成！找到今天與本週的叫車排程。")

                        # 整理 Sheet 1: 今日叫車構件明細
                        sheet1_cols = ['工程編號', '工程區', '節次', '構件編號', '組立日期', '焊接日期', '二檢日期', '噴漆施工日期', '批號', '叫車日期', '裝車日期', '工地需求日']
                        for c in sheet1_cols:
                            if c not in df_today.columns: df_today[c] = ''
                        df_sheet1 = df_today[sheet1_cols]

                        # 整理 Sheet 2: 今日叫車廠商及批號
                        df_today['BOM數量'] = pd.to_numeric(df_today['BOM數量'] if 'BOM數量' in df_today.columns else 1, errors='coerce').fillna(0)
                        group_cols = ['塗裝廠商', '工程編號', '批號', '叫車日期', '裝車日期', '工地需求日']
                        df_sheet2 = df_today.groupby(group_cols, dropna=False, as_index=False)['BOM數量'].sum().rename(columns={'BOM數量': '構件支數'})
                        sheet2_cols = ['塗裝廠商', '工程編號', '批號', '構件支數', '叫車日期', '裝車日期', '工地需求日']
                        df_sheet2 = df_sheet2[sheet2_cols]

                        # 整理 Sheet 3: 本週叫車廠商及批號
                        df_week['BOM數量'] = pd.to_numeric(df_week['BOM數量'] if 'BOM數量' in df_week.columns else 1, errors='coerce').fillna(0)
                        df_sheet3 = df_week.groupby(group_cols, dropna=False, as_index=False)['BOM數量'].sum().rename(columns={'BOM數量': '構件支數'})
                        if not df_sheet3.empty: 
                            df_sheet3 = df_sheet3.sort_values(by=['叫車日期', '塗裝廠商'])
                        df_sheet3 = df_sheet3[sheet2_cols]

                        # 顯示在網頁上供預覽 (已補上今日廠商總表)
                        if not df_sheet1.empty:
                            st.subheader("📌 今日叫車構件明細 (預覽)")
                            st.dataframe(df_sheet1, use_container_width=True)
                        if not df_sheet2.empty:
                            st.subheader("🚚 今日叫車廠商總表 (預覽)")
                            st.dataframe(df_sheet2, use_container_width=True)
                        if not df_sheet3.empty:
                            st.subheader("📅 本週叫車總表 (預覽)")
                            st.dataframe(df_sheet3, use_container_width=True)

                        # 將 Excel 寫入記憶體中供下載
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_sheet1.to_excel(writer, index=False, sheet_name='今日叫車構件明細')
                            df_sheet2.to_excel(writer, index=False, sheet_name='今日叫車廠商及批號')
                            df_sheet3.to_excel(writer, index=False, sheet_name='本週叫車廠商及批號')
                            
                            contractors = df_sheet2['塗裝廠商'].unique()
                            for contractor in contractors:
                                safe_sheet_name = str(contractor)[:31].replace('/', '_').replace('*', '_').replace('[', '').replace(']', '')
                                df_contractor = df_sheet2[df_sheet2['塗裝廠商'] == contractor]
                                df_contractor.to_excel(writer, index=False, sheet_name=safe_sheet_name)
                        output.seek(0)

                        st.markdown("---")
                        st.download_button(
                            label="📥 下載完整 Excel 安排表",
                            data=output,
                            file_name=f"{today.strftime('%Y%m%d')}_叫車安排表.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
