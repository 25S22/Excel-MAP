import pandas as pd
import requests
import urllib3
from datetime import datetime
import time
import os
import win32com.client  # For creating draft emails in Outlook

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
INPUT_EXCEL_PATH = r'C:\path\to\your\input.xlsx'
SHEETS_TO_PROCESS = ['Sheet1', 'Sheet2']  # or ['all'] for all sheets
LOGSOURCE_COLUMN = 'log source name'
IP_COLUMN = 'IP'
QRADAR_HOST = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL = False
DRAFT_OUTPUT_PATH = os.path.join(os.path.dirname(INPUT_EXCEL_PATH), 'inactive_and_errors.xlsx')
# ─── END CONFIGURATION ─────────────────────────────────────────────────────────

def test_qradar_connection(qradar_host, username, password):
    """Test if we can connect to the QRadar API."""
    print("🔗 Testing QRadar connection...")
    qradar_host = qradar_host.rstrip('/')
    endpoint = f"{qradar_host}/api/help/versions"
    try:
        resp = requests.get(endpoint, auth=(username, password), verify=VERIFY_SSL,
                            timeout=10, headers={'Accept': 'application/json'})
        if resp.status_code == 200:
            print("✅ QRadar connection successful!")
            return True
        elif resp.status_code == 401:
            print("❌ Authentication failed! Check username/password.")
            return False
        else:
            print(f"⚠️ Unexpected response: {resp.status_code}")
            return False
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False

def _empty_details():
    return {'qradar_id': '', 'protocol_type': '', 'enabled': '', 'last_seen': '', 'activity_status': ''}

def _start_aql_search(qradar_host, username, password, query):
    endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches"
    try:
        resp = requests.post(endpoint, auth=(username, password), verify=VERIFY_SSL,
                             timeout=30,
                             headers={'Accept': 'application/json', 'Content-Type': 'application/json'},
                             json={'query_expression': query})
        if resp.status_code == 201:
            return resp.json().get('search_id')
    except Exception:
        pass
    return None

def _get_search_results(qradar_host, username, password, search_id):
    endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches/{search_id}/results"
    for _ in range(10):
        time.sleep(2)
        resp = requests.get(endpoint, auth=(username, password), verify=VERIFY_SSL,
                            timeout=30,
                            headers={'Accept': 'application/json'})
        if resp.status_code == 200:
            data = resp.json()
            if 'events' in data:
                return data
    return {}

def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    filter_key = 'ip_address' if is_ip else 'name'
    query_filter = f'{filter_key}="{identifier}"'
    ls_endpoint = f"{qradar_host.rstrip('/')}/api/config/event_sources/log_source_management/log_sources"
    try:
        resp = requests.get(ls_endpoint, params={'filter': query_filter},
                            auth=(username, password), verify=VERIFY_SSL,
                            timeout=30, headers={'Accept': 'application/json'})
        if resp.status_code != 200:
            return {'status': f'API Error {resp.status_code}', **_empty_details()}
        ls_data = resp.json()
        if not ls_data:
            return {'status': 'Not Found', **_empty_details()}
        log_source = ls_data[0]
        ls_id = log_source.get('id')
        aql = f"SELECT MAX(starttime) FROM events WHERE logsourceid={ls_id} LAST 30 DAYS"
        search_id = _start_aql_search(qradar_host, username, password, aql)
        last_seen = ''
        activity_status = 'No Activity'
        if search_id:
            results = _get_search_results(qradar_host, username, password, search_id)
            events = results.get('events', [])
            if events and events[0].get('MAX(starttime)'):
                epoch_ms = events[0]['MAX(starttime)']
                last_seen = datetime.fromtimestamp(epoch_ms/1000).strftime('%Y-%m-%d %H:%M:%S')
                activity_status = 'Active'
            else:
                last_seen = 'No events in last 30 days'
        return {
            'status': 'Found', 'qradar_id': ls_id,
            'protocol_type': log_source.get('protocol_type', ''),
            'enabled': log_source.get('enabled', ''),
            'last_seen': last_seen, 'activity_status': activity_status
        }
    except Exception as e:
        return {'status': f'Error: {e}', **_empty_details()}

def process_sheet(df, sheet_name, qradar_host, username, password, logsource_column, ip_column):
    print(f"\n📋 Processing sheet: {sheet_name}")
    for col in [logsource_column, ip_column]:
        if col not in df.columns:
            print(f"❌ Column '{col}' not found in {sheet_name}! Available: {list(df.columns)}")
            return df
    for col in ['status','qradar_id','protocol_type','enabled','last_seen','activity_status']:
        if col not in df.columns:
            df[col] = ''
    total = len(df)
    print(f"Found {total} rows to process...")
    for idx, row in df.iterrows():
        name_val = str(row[logsource_column]).strip()
        details = None
        if name_val and name_val.lower() not in ['nan','none','']:
            print(f"[{idx+1}/{total}] Lookup by name: {name_val}")
            details = get_log_source_details(qradar_host, username, password, name_val, is_ip=False)
        if not details or details['status']=='Not Found':
            ip_val = str(row[ip_column]).strip()
            if ip_val and ip_val.lower() not in ['nan','none','']:
                print(f"   🔁 Fallback to IP: {ip_val}")
                details = get_log_source_details(qradar_host, username, password, ip_val, is_ip=True)
        if not details:
            details = {'status':'Empty/Invalid', **_empty_details()}
        for k,v in details.items(): df.at[idx,k]=v
        print(f"   → {details['status']} | Last Seen: {details['last_seen']}")
        time.sleep(0.5)
    return df

def create_outlook_draft(attachment_path, subject, body):
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(attachment_path)
    mail.Save()
    print(f"✉️ Draft created: {attachment_path}")

def filter_and_email(df_dict, draft_path):
    frames=[]
    for name,df in df_dict.items():
        if 'status' in df.columns and 'activity_status' in df.columns:
            mask=(df['activity_status']!='Active')|(df['status'].str.startswith('API Error'))
            subset=df[mask].copy()
            if not subset.empty:
                subset['remark']='Check log source name'
                subset['sheet_name']=name
                frames.append(subset)
    if not frames:
        print("✅ No inactive/API errors; skipping email.")
        return
    result_df=pd.concat(frames,ignore_index=True)
    total=len(result_df)
    num_inactive=result_df[result_df['activity_status']!='Active'].shape[0]
    num_api_errors=result_df[result_df['status'].str.startswith('API Error')].shape[0]
    result_df.to_excel(draft_path,index=False)
    print(f"💾 Saved flagged rows to: {draft_path}")
    subject="Inactive Log Sources"
    body=(
        f"Hello,\n\n"
        f"Attached is the QRadar log source report as of {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.\n"
        f"Total flagged: {total}\n"
        f"Inactive: {num_inactive}\n"
        f"API errors: {num_api_errors}\n\n"
        "Rows are tagged 'Check log source name'.\n\n"
        "Best regards,\n"
        "QRadar Automation Bot"
    )
    create_outlook_draft(draft_path,subject,body)

def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    print("🚀 Starting QRadar Log Source Checker...")
    if not test_qradar_connection(QRADAR_HOST,QRADAR_USERNAME,QRADAR_PASSWORD): return
    print(f"\n📖 Reading Excel: {INPUT_EXCEL_PATH}")
    all_sheets=pd.read_excel(INPUT_EXCEL_PATH,sheet_name=None)
    sheets=list(all_sheets.keys())
    print(f"Sheets found: {sheets}")
    to_proc=sheets if SHEETS_TO_PROCESS==['all'] else SHEETS_TO_PROCESS
    for sheet in to_proc:
        if sheet in all_sheets:
            all_sheets[sheet]=process_sheet(
                all_sheets[sheet],sheet,QRADAR_HOST,QRADAR_USERNAME,QRADAR_PASSWORD,
                LOGSOURCE_COLUMN,IP_COLUMN
            )
        else:
            print(f"⚠️ Skipping missing sheet: {sheet}")
    print(f"\n💾 Saving updates to original Excel...")
    with pd.ExcelWriter(INPUT_EXCEL_PATH,engine='openpyxl') as writer:
        for name,df in all_sheets.items(): df.to_excel(writer,sheet_name=name,index=False)
    print("✅ Original updated.")
    filter_and_email(all_sheets,DRAFT_OUTPUT_PATH)
    for sheet in to_proc:
        df=all_sheets.get(sheet)
        if df is not None:
            print(f"\n📊 Summary for {sheet}: {df['status'].value_counts().to_dict()} | {df['activity_status'].value_counts().to_dict()}")

if __name__=='__main__':
    main()
