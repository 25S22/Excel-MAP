import pandas as pd
import requests
import urllib3
from datetime import datetime, timedelta
import time

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────

INPUT_EXCEL_PATH = r'C:\path\to\your\input.xlsx'
SHEETS_TO_PROCESS = ['Sheet1', 'Sheet2']
LOGSOURCE_COLUMN = 'log source name'
IP_COLUMN = 'IP'
QRADAR_HOST = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL = False
INACTIVITY_THRESHOLD_DAYS = 30

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────

# Create a session for reuse
session = requests.Session()

def debug_get(session, url, **kwargs):
    req = requests.Request('GET', url, **kwargs)
    prepped = session.prepare_request(req)
    print(f"→ GET {prepped.url}")
    if 'params' in kwargs:
        print(f"   params: {kwargs['params']}")
    resp = session.send(prepped, verify=kwargs.get('verify', True), timeout=kwargs.get('timeout', 30))
    print(f"   ← Response {resp.status_code}: {resp.text}")
    if resp.status_code == 422:
        try:
            print("   422 details JSON:", resp.json())
        except Exception:
            pass
    return resp

def debug_post(session, url, **kwargs):
    req = requests.Request('POST', url, **kwargs)
    prepped = session.prepare_request(req)
    print(f"→ POST {prepped.url}")
    if 'json' in kwargs:
        print(f"   JSON body: {kwargs['json']}")
    resp = session.send(prepped, verify=kwargs.get('verify', True), timeout=kwargs.get('timeout', 30))
    print(f"   ← Response {resp.status_code}: {resp.text}")
    if resp.status_code == 422:
        try:
            print("   422 details JSON:", resp.json())
        except Exception:
            pass
    return resp

def test_qradar_connection(qradar_host, username, password):
    print("🔗 Testing QRadar connection...")
    resp = session.get(
        f"{qradar_host.rstrip('/')}/api/help/versions",
        auth=(username, password),
        verify=VERIFY_SSL,
        timeout=10,
        headers={'Accept': 'application/json'}
    )
    print(f"  → GET {resp.url} returned {resp.status_code}")
    return resp.status_code == 200

def _empty_details():
    return {'status':'','qradar_id':'','protocol_type':'','enabled':'',
            'last_seen':'','activity_status':'','days_since_last_event':None,
            'might_be_disabled_alert':''}

def _start_aql_search(qradar_host, username, password, query):
    endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches"
    resp = debug_post(session, endpoint,
                      auth=(username, password),
                      headers={'Accept': 'application/json','Content-Type': 'application/json'},
                      json={'query_expression': query},
                      verify=VERIFY_SSL, timeout=30)
    if resp.status_code == 201:
        return resp.json().get('search_id')
    return None

def _get_search_results(qradar_host, username, password, search_id):
    endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches/{search_id}/results"
    for i in range(15):
        time.sleep(2)
        resp = debug_get(session, endpoint,
                         auth=(username, password),
                         headers={'Accept':'application/json'},
                         verify=VERIFY_SSL, timeout=30)
        if resp.status_code == 200 and 'events' in resp.json():
            return resp.json()
    return {}

def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    filter_key = 'ip_address' if is_ip else 'name'
    query_filter = f'{filter_key}="{identifier}"'
    url = f"{qradar_host.rstrip('/')}/api/config/event_sources/log_source_management/log_sources"
    resp = debug_get(session, url,
                     params={'filter': query_filter},
                     auth=(username, password),
                     headers={'Accept':'application/json'},
                     verify=VERIFY_SSL, timeout=30)
    if resp.status_code != 200:
        return {**_empty_details(), 'status':f"API Error {resp.status_code}"}
    data = resp.json()
    if not isinstance(data, list) or not data:
        return {**_empty_details(), 'status':'Not Found'}
    ls = data[0]
    ls_id = ls.get('id')
    if not ls_id:
        return {**_empty_details(), 'status':'No ID'}

    aql = (f"SELECT MAX(starttime) AS max_starttime FROM events "
           f"WHERE logsourceid={ls_id} LAST {INACTIVITY_THRESHOLD_DAYS} DAYS")
    sid = _start_aql_search(qradar_host, username, password, aql)
    if not sid:
        return {**_empty_details(), 'status':'AQL Start Error'}
    res = _get_search_results(qradar_host, username, password, sid)
    events = res.get('events', [])
    last_seen = f"No events in last {INACTIVITY_THRESHOLD_DAYS} days"
    activity_status = 'Inactive'
    days_since = None
    alert = 'Yes'
    if events and events[0].get('max_starttime'):
        dt = datetime.fromtimestamp(events[0]['max_starttime']/1000)
        last_seen = dt.strftime('%Y-%m-%d %H:%M:%S')
        days_since = (datetime.now() - dt).days
        if days_since <= INACTIVITY_THRESHOLD_DAYS:
            activity_status = 'Active'
            alert = 'No'
    return {
        'status':'Found',
        'qradar_id': ls_id,
        'protocol_type': ls.get('protocol_type',''),
        'enabled': ls.get('enabled',''),
        'last_seen': last_seen,
        'activity_status': activity_status,
        'days_since_last_event': days_since,
        'might_be_disabled_alert': alert
    }

def process_sheet(df, sheet_name, qradar_host, username, password, logsrc_col, ip_col):
    print(f"\n📋 Processing sheet: {sheet_name}")
    for col in [logsrc_col, ip_col]:
        if col not in df.columns:
            print(f"❌ Column '{col}' missing.")
            return df
    for col in ['status','qradar_id','protocol_type','enabled',
                'last_seen','activity_status','days_since_last_event','might_be_disabled_alert']:
        if col not in df.columns:
            df[col] = ''
    for idx, row in df.iterrows():
        identifier = str(row[logsrc_col]).strip()
        details = None
        if identifier and identifier.lower() not in ['nan','none','']:
            print(f"[{idx+1}] Lookup by name: {identifier}")
            details = get_log_source_details(qradar_host, username, password, identifier, is_ip=False)
        if not details or details.get('status') == 'Not Found':
            ip = str(row[ip_col]).strip()
            if ip and ip.lower() not in ['nan','none','']:
                print(f"  🔁 Fallback by IP: {ip}")
                details = get_log_source_details(qradar_host, username, password, ip, is_ip=True)
        if not details:
            details = {**_empty_details(), 'status':'Empty'}
        for k,v in details.items():
            df.at[idx, k] = v
        print(f"  → {details['status']} | LastSeen: {details['last_seen']} | Alert: {details['might_be_disabled_alert']}")
        time.sleep(0.5)
    return df

def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings()
    if not test_qradar_connection(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD):
        print("❌ Connection failed.")
        return
    print(f"📖 Loading Excel: {INPUT_EXCEL_PATH}")
    try:
        sheets = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=None)
    except Exception as e:
        print("❌ Excel load error:", e)
        return
    to_process = list(sheets.keys()) if SHEETS_TO_PROCESS == ['all'] else SHEETS_TO_PROCESS
    for sheet in to_process:
        if sheet in sheets:
            sheets[sheet] = process_sheet(sheets[sheet], sheet, QRADAR_HOST,
                                          QRADAR_USERNAME, QRADAR_PASSWORD,
                                          LOGSOURCE_COLUMN, IP_COLUMN)
        else:
            print(f"⚠️ Sheet '{sheet}' not found.")
    print("💾 Saving results...")
    try:
        with pd.ExcelWriter(INPUT_EXCEL_PATH, engine='openpyxl') as w:
            for name, df in sheets.items():
                df.to_excel(w, sheet_name=name, index=False)
        print("✅ Done.")
    except Exception as e:
        print("❌ Save error:", e)

if __name__ == '__main__':
    main()
