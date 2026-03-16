import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import json
import re

# Google API 套件
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# 設定網頁標題與寬度
st.set_page_config(page_title="叫車排程系統", page_icon="🚛", layout="wide")

# =====================================================================
# ⬇️ 請在這裡填入您 Google Drive 的資料夾與檔案 ID ⬇️
CTRL_FOLDER_ID = '1Cd8OWf6unmQP0qZax6jXMp3AdHaaEUDG'
SHIP_FOLDER_ID = '1yLo56xotUbdirQGvuITPYketISSZXsYq'
MASTER_FILE_ID = '18plINq_UXDovZGCPT4XrqKHRRXYBaevK'  # 例如: '1B3d4E5f6G7h8I9j0K1lM'
# =====================================================================

st.title("🚛自動叫車排程系統")

# 檢查是否已設定 Google 金鑰
if "GCP_KEY_JSON" not in st.secrets:
    st.error("⚠️ 尚未設定 Google Drive API 金鑰！請至 Streamlit 後台的 Secrets 設定。")
    st.stop()

@st.cache_resource
def get_gdrive_service():
    """初始化 Google Drive API 服務"""
    try:
        key_dict = json.loads(st.secrets["GCP_KEY_JSON"])
        creds = service_account.Credentials.from_service_account_info(
            key_dict, scopes=['https://www.googleapis.com/auth/drive.readonly']
        )
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.error(f"連線 Google 服務失敗，請確認 Secrets 格式是否正確。錯誤訊息：{e}")
        st.stop()

@st.cache_data(ttl=1800) # 快取 30 分鐘，避免每次切換名字都重新下載一覽表
def load_master_file(_service, file_id):
    """讀取主辦一覽表，支援 Excel, CSV 與 Google 試算表"""
    if file_id == '這裡貼上一覽表檔案ID':
        return None
    try:
        file_info = _service.files().get(fileId=file_id, fields="name, mimeType").execute()
        mime_type = file_info.get('mimeType', '')
        
        if mime_type == 'application/vnd.google-apps.spreadsheet':
            request = _service.files().export_media(
                fileId=file_id, 
                mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            request = _service.files().get_media(fileId=file_id)
            
        fh = BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        fh.seek(0)
        
        if file_info['name'].endswith('.csv'):
            try:
                df = pd.read_csv(fh, encoding='utf-8')
            except UnicodeDecodeError:
                fh.seek(0)
                df = pd.read_csv(fh, encoding='big5')
        else:
            df = pd.read_excel(fh)
        return df
    except Exception as e:
        st.error(f"❌ 讀取一覽表失敗，請確認 ID 是否正確或是否有權限。錯誤：{e}")
        return None

def fetch_files_from_drive(folder_id, service, target_project_ids=None):
    """下載指定資料夾內的檔案，如果 target_project_ids 有值，就只下載符合案號的檔案"""
    query = f"'{folder_id}' in parents and trashed=false"
    try:
        results = service.files().list(q=query, fields="files(id, name, mimeType)", pageSize=1000).execute()
        items = results.get('files', [])
        downloaded_files = []
        for item in items:
            file_name = item['name']
            mime_type = item['mimeType']
            
            if file_name.endswith(('.xlsx', '.csv', '.xls')):
                
                # 【效能優化】：過濾掉不相關的檔案，不下載不處理！
                if target_project_ids is not None:
                    if not any(str(pid).strip() in file_name for pid in target_project_ids):
                        continue 
                        
                if mime_type == 'application/vnd.google-apps.spreadsheet':
                    request = service.files().export_media(
                        fileId=item['id'], 
                        mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                else:
                    request = service.files().get_media(fileId=item['id'])
                
                if request is None: continue
                fh = BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                fh.seek(0)
                downloaded_files.append({'name': file_name, 'content': fh})
        return downloaded_files
    except Exception as e:
        st.error(f"❌ 讀取資料夾失敗！錯誤：{e}")
        return []

service = get_gdrive_service()

# 讀取一覽表並產生互動選單
df_master = load_master_file(service, MASTER_FILE_ID)
selected_projects = None

if df_master is not None and not df_master.empty and '生管' in df_master.columns and '工程編號' in df_master.columns:
    st.markdown("### 🧑‍💻 第一步：選擇生管主辦")
    # 萃取出乾淨、不重複的名字清單
    names_series = df_master['生管'].dropna().astype(str)
    all_names = set()
    for n in names_series:
        clean_n = re.sub(r'(主辦:|協辦:|\n|-)', ' ', n)
        for p in clean_n.split():
            if p.strip() and p.strip() != 'nan':
                all_names.add(p.strip())
    manager_list = ["請選擇..."] + sorted(list(all_names))
    
    selected_manager = st.selectbox("🙋‍♂️ 請選擇您的名字：", manager_list)
    
    if selected_manager != "請選擇...":
        st.markdown("### 🏗️ 第二步：確認工程案號")
        
        # 尋找該主辦的所有工程案，並排除已結案的
        mask_manager = df_master['生管'].astype(str).str.contains(selected_manager, na=False)
        if '結案' in df_master.columns:
            mask_active = df_master['結案'].isna() | (df_master['結案'].astype(str).str.strip() == '') | (df_master['結案'].astype(str).str.lower() == 'nan')
            active_projects = df_master[mask_manager & mask_active]['工程編號'].dropna().astype(str).tolist()
        else:
            active_projects = df_master[mask_manager]['工程編號'].dropna().astype(str).tolist()
            
        active_projects = [str(p).strip() for p in active_projects if str(p).strip()]
        all_options = sorted(df_master['工程編號'].dropna().astype(str).unique().tolist())
        safe_active_projects = [p for p in active_projects if p in all_options]

        # 讓使用者可以微調要排程的案子
        selected_projects = st.multiselect(
            "📌 系統已為您自動勾選「未結案」的工程，您可手動新增/刪除要排程的案子：",
            options=all_options,
            default=safe_active_projects
        )
else:
    st.info("尚未設定主辦一覽表 ID 或欄位格式錯誤。系統將不篩選案號，直接讀取資料夾內所有檔案。")


st.markdown("---")

# 執行按鈕
if st.button("🚀 確認案號無誤，開始執行排程", use_container_width=True, type="primary"):
    
    if CTRL_FOLDER_ID == '這裡貼上管制表資料夾ID' or SHIP_FOLDER_ID == '這裡貼上出貨單資料夾ID':
        st.error("⚠️ 請先在程式碼第 18~20 行填入您真實的 Google Drive ID！")
        st.stop()

    if selected_projects is not None and len(selected_projects) == 0:
        st.warning("⚠️ 案號清單目前為空，請確認您有勾選要運算的工程案喔！")
        st.stop()

    with st.spinner("🔄 系統正在精準鎖定並下載您的專屬檔案，請稍候..."):
        # 雲端下載檔案 (只抓選取的案子！)
        ctrl_files = fetch_files_from_drive(CTRL_FOLDER_ID, service, target_project_ids=selected_projects)
        ship_files = fetch_files_from_drive(SHIP_FOLDER_ID, service, target_project_ids=selected_projects)

        if not ctrl_files:
            st.warning("⚠️ 在「管制表」資料夾中找不到符合上述案號的檔案。")
        if not ship_files:
            st.warning("⚠️ 在「出貨單」資料夾中找不到符合上述案號的檔案。")

        ctrl_dfs, ship_dfs = [], []

        # 1. 處理管制表
        for file_data in ctrl_files:
            try:
                file_name, file_content = file_data['name'], file_data['content']
                if file_name.endswith('.csv'):
                    df_temp = pd.read_csv(file_content, nrows=20, header=None)
                    flat_str = ' '.join(df_temp.fillna('').astype(str).values.flatten())
                    if '構件編號' in flat_str and '組立日期' in flat_str:
                        header_idx = df_temp[df_temp.apply(lambda r: r.astype(str).str.contains('構件編號').any(), axis=1)].index
                        if len(header_idx) > 0:
                            file_content.seek(0)
                            ctrl_dfs.append(pd.read_csv(file_content, skiprows=header_idx[0]))
                else:
                    xls_dict = pd.read_excel(file_content, sheet_name=None, header=None)
                    for sheet_name, df in xls_dict.items():
                        if df.empty: continue
                        head_df = df.head(20).fillna('').astype(str)
                        flat_str = ' '.join(head_df.values.flatten())
                        if '構件編號' in flat_str and '組立日期' in flat_str:
                            header_idx = head_df[head_df.apply(lambda r: r.str.contains('構件編號').any(), axis=1)].index
                            if len(header_idx) > 0:
                                ctrl_dfs.append(pd.read_excel(file_content, sheet_name=sheet_name, skiprows=header_idx[0]))
            except Exception as e:
                st.error(f"讀取 {file_name} 時發生錯誤：{e}")

        # 2. 處理出貨聯絡單
        for file_data in ship_files:
            try:
                file_name, file_content = file_data['name'], file_data['content']
                if file_name.endswith('.csv'):
                    df_temp = pd.read_csv(file_content, nrows=20, header=None)
                    flat_str = ' '.join(df_temp.fillna('').astype(str).values.flatten())
                    if '實際交貨日期' in flat_str and '批號' in flat_str:
                        header_idx = df_temp[df_temp.apply(lambda r: r.astype(str).str.contains('實際交貨日期').any(), axis=1)].index
                        if len(header_idx) > 0:
                            file_content.seek(0)
                            ship_dfs.append(pd.read_csv(file_content, skiprows=header_idx[0]))
                else:
                    xls_dict = pd.read_excel(file_content, sheet_name=None, header=None)
                    for sheet_name, df in xls_dict.items():
                        if df.empty: continue
                        head_df = df.head(20).fillna('').astype(str)
                        flat_str = ' '.join(head_df.values.flatten())
                        if '實際交貨日期' in flat_str and '批號' in flat_str:
                            header_idx = head_df[head_df.apply(lambda r: r.str.contains('實際交貨日期').any(), axis=1)].index
                            if len(header_idx) > 0:
                                ship_dfs.append(pd.read_excel(file_content, sheet_name=sheet_name, skiprows=header_idx[0]))
            except Exception as e:
                st.error(f"讀取 {file_name} 時發生錯誤：{e}")

        df_ctrl = pd.concat(ctrl_dfs, ignore_index=True) if ctrl_dfs else pd.DataFrame()
        df_ship = pd.concat(ship_dfs, ignore_index=True) if ship_dfs else pd.DataFrame()

        if df_ship.empty or df_ctrl.empty:
            st.error("❌ 未能合併有效資料，請確認您選取的工程案在 Drive 中是否有對應的檔案。")
        else:
            # 3. 清理與合併
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
                # 4. 運算邏輯
                def calculate_loading_date(delivery_date):
                    if pd.isnull(delivery_date): return pd.NaT
                    loading_date = delivery_date - pd.Timedelta(days=1)
                    if loading_date.weekday() == 6: loading_date -= pd.Timedelta(days=2)
                    return loading_date

                def calculate_calling_date(loading_date):
                    if pd.isnull(loading_date): return pd.NaT
                    calling_date = loading_date - pd.Timedelta(days=1)
                    if calling_date.weekday() == 6: calling_date -= pd.Timedelta(days=2)
                    elif calling_date.weekday() == 5: calling_date -= pd.Timedelta(days=1)
                    return calling_date

                df_merged['裝車日期'] = df_merged['工地需求日'].apply(calculate_loading_date)
                df_merged['叫車日期'] = df_merged['裝車日期'].apply(calculate_calling_date)

                if '噴漆包商' not in df_merged.columns: df_merged['噴漆包商'] = ''
                if '預計塗裝廠商' not in df_merged.columns: df_merged['預計塗裝廠商'] = df_merged.get('一次預定廠商', '')

                df_merged['噴漆包商'] = df_merged['噴漆包商'].astype(str).replace(['nan', 'NaN'], '').str.strip()
                df_merged['預計塗裝廠商'] = df_merged['預計塗裝廠商'].astype(str).replace(['nan', 'NaN'], '').str.strip()
                df_merged['塗裝廠商'] = np.where(df_merged['噴漆包商'] != '', df_merged['噴漆包商'], df_merged['預計塗裝廠商'])
                df_merged['塗裝廠商'] = df_merged['塗裝廠商'].replace('', '未指定')

                df_merged = df_merged.rename(columns={'工程案號': '工程編號', '節數': '節次'})

                date_cols = ['組立日期', '焊接日期', '二檢日期', '噴漆施工日期', '叫車日期', '裝車日期', '工地需求日']
                for c in date_cols:
                    if c in df_merged.columns: df_merged[c] = pd.to_datetime(df_merged[c], errors='coerce')

                today = pd.Timestamp.utcnow().tz_convert('Asia/Taipei').normalize().tz_localize(None)
                start_of_week = today - pd.Timedelta(days=today.weekday())
                end_of_week = start_of_week + pd.Timedelta(days=6)

                df_today = df_merged[df_merged['叫車日期'] == today].copy()
                df_week = df_merged[(df_merged['叫車日期'] >= start_of_week) & (df_merged['叫車日期'] <= end_of_week)].copy()

                for df_target in [df_today, df_week]:
                    for c in date_cols:
                        if c in df_target.columns: df_target[c] = df_target[c].dt.strftime('%Y/%m/%d').fillna('')

                if df_today.empty and df_week.empty:
                    st.info("🎉 您的專屬排程已同步完畢！今天與本週目前都沒有需要叫車的案子。")
                else:
                    st.success(f"✅ 成功為您計算完專屬的排程！")

                    sheet1_cols = ['工程編號', '工程區', '節次', '構件編號', '組立日期', '焊接日期', '二檢日期', '噴漆施工日期', '批號', '叫車日期', '裝車日期', '工地需求日']
                    for c in sheet1_cols:
                        if c not in df_today.columns: df_today[c] = ''
                    df_sheet1 = df_today[sheet1_cols]

                    df_today['BOM數量'] = pd.to_numeric(df_today['BOM數量'] if 'BOM數量' in df_today.columns else 1, errors='coerce').fillna(0)
                    group_cols = ['塗裝廠商', '工程編號', '批號', '叫車日期', '裝車日期', '工地需求日']
                    df_sheet2 = df_today.groupby(group_cols, dropna=False, as_index=False)['BOM數量'].sum().rename(columns={'BOM數量': '構件支數'})
                    sheet2_cols = ['塗裝廠商', '工程編號', '批號', '構件支數', '叫車日期', '裝車日期', '工地需求日']
                    df_sheet2 = df_sheet2[sheet2_cols]

                    df_week['BOM數量'] = pd.to_numeric(df_week['BOM數量'] if 'BOM數量' in df_week.columns else 1, errors='coerce').fillna(0)
                    df_sheet3 = df_week.groupby(group_cols, dropna=False, as_index=False)['BOM數量'].sum().rename(columns={'BOM數量': '構件支數'})
                    if not df_sheet3.empty: df_sheet3 = df_sheet3.sort_values(by=['叫車日期', '塗裝廠商'])
                    df_sheet3 = df_sheet3[sheet2_cols]

                    if not df_sheet2.empty:
                        st.subheader("🚚 1. 今日叫車廠商總表 (預覽)")
                        st.dataframe(df_sheet2, use_container_width=True)
                    if not df_sheet3.empty:
                        st.subheader("📅 2. 本週叫車總表 (預覽)")
                        st.dataframe(df_sheet3, use_container_width=True)
                    if not df_sheet1.empty:
                        st.subheader("📌 3. 今日叫車構件明細 (預覽)")
                        st.dataframe(df_sheet1, use_container_width=True)

                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_sheet1.to_excel(writer, index=False, sheet_name='今日叫車構件明細')
                        df_sheet2.to_excel(writer, index=False, sheet_name='今日叫車廠商及批號')
                        df_sheet3.to_excel(writer, index=False, sheet_name='本週叫車廠商及批號')
                        for contractor in df_sheet2['塗裝廠商'].unique():
                            safe_name = str(contractor)[:31].replace('/', '_').replace('*', '_').replace('[', '').replace(']', '')
                            df_contractor = df_sheet2[df_sheet2['塗裝廠商'] == contractor]
                            df_contractor.to_excel(writer, index=False, sheet_name=safe_name)
                    output.seek(0)

                    st.markdown("---")
                    st.download_button(
                        label="📥 下載完整 Excel 安排表",
                        data=output,
                        file_name=f"{today.strftime('%Y%m%d')}_叫車安排表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
