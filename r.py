import pandas as pd
import requests
import urllib3
from datetime import datetime, timedelta
import time
import os
import win32com.client  # For creating draft emails in Outlook

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
# 1) Path to your existing Excel file:
INPUT_EXCEL_PATH = r'C:\path\to\your\input.xlsx'
# 2) List of sheet names to process (or ['all'] for all):
SHEETS_TO_PROCESS = ['Sheet1', 'Sheet2']
# 3) Column containing log source names:
LOGSOURCE_COLUMN = 'log source name'
# 4) Column containing IP addresses:
IP_COLUMN = 'IP'
# 5) QRadar API details:
QRADAR_HOST = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
# 6) SSL verification:
VERIFY_SSL = False
# 7) Path for filtered report:
DRAFT_OUTPUT_PATH = os.path.join(os.path.dirname(INPUT_EXCEL_PATH), 'inactive_and_errors.xlsx')
# ─── END CONFIGURATION ─────────────────────────────────────────────────────────


def test_qradar_connection(qradar_host, username, password):
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
    return {'qradar_id':'','protocol_type':'','enabled':'','last_seen':'','activity_status':''}


def _start_aql_search(qradar_host, username, password, query):
    endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches"
    try:
        resp = requests.post(endpoint, auth=(username,password), verify=VERIFY_SSL,
                             timeout=30,
                             headers={'Accept':'application/json','Content-Type':'application/json'},
                             json={'query_expression':query})
        if resp.status_code==201:
            return resp.json().get('search_id')
    except:
        pass
    return None


def _get_search_results(qradar_host, username, password, search_id):
    endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches/{search_id}/results"
    for _ in range(10):
        time.sleep(2)
        resp = requests.get(endpoint, auth=(username,password), verify=VERIFY_SSL,
                            timeout=30, headers={'Accept':'application/json'})
        if resp.status_code==200:
            data = resp.json()
            if 'events' in data:
                return data
    return {}


def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    filter_key='ip_address' if is_ip else 'name'
    query_filter=f'{filter_key}="{identifier}"'
    ls_endpoint=f"{qradar_host.rstrip('/')}/api/config/event_sources/log_source_management/log_sources"
    try:
        resp=requests.get(ls_endpoint, params={'filter':query_filter}, auth=(username,password),
                          verify=VERIFY_SSL, timeout=30, headers={'Accept':'application/json'})
        if resp.status_code!=200:
            return {'status':f'API Error {resp.status_code}',**_empty_details()}
        ls_data=resp.json()
        if not ls_data:
            return {'status':'Not Found',**_empty_details()}
        log_source=ls_data[0]
        ls_id=log_source.get('id')
        aql=f"SELECT MAX(starttime) FROM events WHERE logsourceid={ls_id} LAST 30 DAYS"
        search_id=_start_aql_search(qradar_host,username,password,aql)
        last_seen=''
        activity_status='No Activity'
        if search_id:
            results=_get_search_results(qradar_host,username,password,search_id)
            events=results.get('events',[])
            if events and events[0].get('MAX(starttime)'):
                epoch_ms=events[0]['MAX(starttime)']
                last_seen=datetime.fromtimestamp(epoch_ms/1000).strftime('%Y-%m-%d %H:%M:%S')
                activity_status='Active'
            else:
                last_seen='No events in last 30 days'
        return {'status':'Found','qradar_id':ls_id,'protocol_type':log_source.get('protocol_type',''),
                'enabled':log_source.get('enabled',''),'last_seen':last_seen,'activity_status':activity_status}
    except Exception as e:
        return {'status':f'Error: {e}',**_empty_details()}


def process_sheet(df, sheet_name, qradar_host, username, password, logsource_column, ip_column):
    print(f"\n📋 Processing sheet: {sheet_name}")
    for col in [logsource_column,ip_column]:
        if col not in df.columns:
            print(f"❌ Column '{col}' not found in {sheet_name}")
            return df
    for col in ['status','qradar_id','protocol_type','enabled','last_seen','activity_status']:
        if col not in df.columns:
            df[col]=''
    total=len(df)
    print(f"Found {total} rows...")
    for idx,row in df.iterrows():
        name_val=str(row[logsource_column]).strip()
        details=None
        if name_val and name_val.lower() not in ['nan','none','']:
            print(f"[{idx+1}/{total}] Lookup by name: {name_val}")
            details=get_log_source_details(qradar_host,username,password,name_val,False)
        if not details or details['status']=='Not Found':
            ip_val=str(row[ip_column]).strip()
            if ip_val and ip_val.lower() not in ['nan','none','']:
                print(f"   🔁 Fallback to IP: {ip_val}")
                details=get_log_source_details(qradar_host,username,password,ip_val,True)
        if not details:
            details={'status':'Empty/Invalid',**_empty_details()}
        for k,v in details.items(): df.at[idx,k]=v
        print(f"   → {details['status']} | Last Seen: {details['last_seen']} | Activity: {details['activity_status']}")
        time.sleep(0.5)
    return df


def create_outlook_draft(attachment_path, subject, body):
    outlook=win32com.client.Dispatch('Outlook.Application')
    mail=outlook.CreateItem(0)
    mail.Subject=subject
    mail.Body=body
    mail.Attachments.Add(attachment_path)
    mail.Display()  # pop-up window
    print(f"✉️ Draft displayed with attachment: {attachment_path}")


def filter_and_email(df_dict, draft_path):
    frames=[]
    for name,df in df_dict.items():
        # inactive only
        inactive= df['activity_status']!='Active'
        errors= df['status'].str.startswith('API Error')
        if inactive.any():
            sub=df[inactive].copy()
            sub['remark']=''
            sub['sheet_name']=name
            frames.append(sub)
        if errors.any():
            sub_err=df[errors].copy()
            sub_err['remark']='Check log source name'
            sub_err['sheet_name']=name
            frames.append(sub_err)
    if not frames:
        print("✅ No flagged rows; skipping email.")
        return
    report=pd.concat(frames,ignore_index=True)
    report.to_excel(draft_path,index=False)
    total=len(report)
    inactive_count=(report['activity_status']!='Active').sum()
    error_count=report['status'].str.startswith('API Error').sum()
    subject="Inactive Log Sources"
    body=(f"Hello,\n\nAttached is the QRadar report as of {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.\n"
          f"Total flagged rows: {total} (Inactive: {inactive_count}, API errors: {error_count}).\n"
          "Remarks only on API errors.\n\nBest regards,\nQRadar Automation Bot")
    create_outlook_draft(draft_path,subject,body)


def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    print("🚀 Starting QRadar checker...")
    if not test_qradar_connection(QRADAR_HOST,QRADAR_USERNAME,QRADAR_PASSWORD): return
    print(f"\n📖 Reading Excel: {INPUT_EXCEL_PATH}")
    sheets=pd.read_excel(INPUT_EXCEL_PATH,sheet_name=None)
    names=list(sheets.keys())
    to_proc=names if SHEETS_TO_PROCESS==['all'] else SHEETS_TO_PROCESS
    for name in to_proc:
        if name in sheets:
            sheets[name]=process_sheet(sheets[name],name,QRADAR_HOST,QRADAR_USERNAME,QRADAR_PASSWORD,LOGSOURCE_COLUMN,IP_COLUMN)
        else:
            print(f"⚠️ Missing sheet: {name}")
    with pd.ExcelWriter(INPUT_EXCEL_PATH,engine='openpyxl') as writer:
        for name,df in sheets.items(): df.to_excel(writer,sheet_name=name,index=False)
    filter_and_email(sheets,DRAFT_OUTPUT_PATH)
    for name,df in sheets.items():
        print(f"📊 {name}: Found={df['status'].eq('Found').sum()}, Inactive={df['activity_status'].eq('Inactive').sum()}, Errors={df['status'].str.startswith('API Error').sum()}")

if __name__=='__main__': main()
