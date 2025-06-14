import pandas as pd
import requests
import urllib3
from datetime import datetime, timedelta
import time

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────

# 1) Path to your existing Excel file (use forward slashes or raw string):
INPUT_EXCEL_PATH = r'C:\path\to\your\input.xlsx'

# 2) List of sheet names to process (or use ['all'] to process all sheets):
SHEETS_TO_PROCESS = ['Sheet1', 'Sheet2']  # or ['all'] for all sheets

# 3) The name of the column containing log source names:
LOGSOURCE_COLUMN = 'log source name'

# 4) The name of the column containing IP addresses (fallback):
IP_COLUMN = 'IP'

# 5) QRadar details:
QRADAR_HOST = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'

# 6) SSL verification (set to False for testing):
VERIFY_SSL = False

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────


def test_qradar_connection(qradar_host, username, password):
    """Test if we can connect to the QRadar API."""
    print("🔗 Testing QRadar connection...")
    qradar_host = qradar_host.rstrip('/')
    test_endpoint = f"{qradar_host}/api/help/versions"

    try:
        resp = requests.get(
            test_endpoint,
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=10,
            headers={'Accept': 'application/json'}
        )
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
    return {
        'qradar_id': '',
        'protocol_type': '',
        'enabled': '',
        'last_seen': '',
        'activity_status': ''
    }


def _start_aql_search(qradar_host, username, password, query):
    endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches"
    try:
        resp = requests.post(
            endpoint,
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=30,
            headers={'Accept': 'application/json', 'Content-Type': 'application/json'},
            json={'query_expression': query}
        )
        if resp.status_code == 201:
            return resp.json().get('search_id')
    except Exception:
        pass
    return None


def _get_search_results(qradar_host, username, password, search_id):
    endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches/{search_id}/results"
    for _ in range(10):
        time.sleep(2)
        resp = requests.get(
            endpoint,
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=30,
            headers={'Accept': 'application/json'}
        )
        if resp.status_code == 200:
            data = resp.json()
            if 'events' in data:
                return data
    return {}


def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    """
    Lookup a log source by name or IP and return its details,
    including the actual last event timestamp.
    """
    filter_key = 'ip_address' if is_ip else 'name'
    query_filter = f'{filter_key}="{identifier}"'
    ls_endpoint = f"{qradar_host.rstrip('/')}/api/config/event_sources/log_source_management/log_sources"

    try:
        resp = requests.get(
            ls_endpoint,
            params={'filter': query_filter},
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=30,
            headers={'Accept': 'application/json'}
        )
        if resp.status_code != 200:
            return {'status': f'API Error {resp.status_code}', **_empty_details()}

        ls_data = resp.json()
        if not ls_data:
            return {'status': 'Not Found', **_empty_details()}

        log_source = ls_data[0]
        ls_id = log_source.get('id')

        # Build AQL to get the MAX(starttime) for last 30 days
        aql = f"SELECT MAX(starttime) FROM events WHERE logsourceid={ls_id} LAST 30 DAYS"
        search_id = _start_aql_search(qradar_host, username, password, aql)

        last_seen = ''
        activity_status = 'No Activity'

        if search_id:
            results = _get_search_results(qradar_host, username, password, search_id)
            events = results.get('events', [])
            if events and events[0].get('MAX(starttime)'):
                epoch_ms = events[0]['MAX(starttime)']
                last_seen = datetime.fromtimestamp(epoch_ms / 1000).strftime('%Y-%m-%d %H:%M:%S')
                activity_status = 'Active'
            else:
                last_seen = 'No events in last 30 days'

        return {
            'status': 'Found',
            'qradar_id': ls_id,
            'protocol_type': log_source.get('protocol_type', ''),
            'enabled': log_source.get('enabled', ''),
            'last_seen': last_seen,
            'activity_status': activity_status
        }

    except Exception as e:
        return {'status': f'Error: {e}', **_empty_details()}


def process_sheet(df, sheet_name, qradar_host, username, password, logsource_column, ip_column):
    """Process a single DataFrame sheet, with fallback to IP lookup."""
    print(f"\n📋 Processing sheet: {sheet_name}")

    # Ensure the required columns exist
    for col in [logsource_column, ip_column]:
        if col not in df.columns:
            print(f"❌ Column '{col}' not found in {sheet_name}! Available columns: {list(df.columns)}")
            return df

    # Prepare result columns
    for col in ['status', 'qradar_id', 'protocol_type', 'enabled', 'last_seen', 'activity_status']:
        if col not in df.columns:
            df[col] = ''

    total = len(df)
    print(f"Found {total} rows to process...")

    for idx, row in df.iterrows():
        name_val = str(row[logsource_column]).strip()
        details = None

        if name_val and name_val.lower() not in ['nan', 'none', '']:
            print(f"[{idx+1}/{total}] Lookup by name: {name_val}")
            details = get_log_source_details(qradar_host, username, password, name_val, is_ip=False)

        if not details or details['status'] == 'Not Found':
            ip_val = str(row[ip_column]).strip() if ip_column in df.columns else ''
            if ip_val and ip_val.lower() not in ['nan', 'none', '']:
                print(f"   🔁 Name not found; fallback to IP: {ip_val}")
                details = get_log_source_details(qradar_host, username, password, ip_val, is_ip=True)

        if not details:
            details = {
                'status': 'Empty/Invalid',
                'qradar_id': '',
                'protocol_type': '',
                'enabled': '',
                'last_seen': '',
                'activity_status': ''
            }

        # Write back to DataFrame
        for k, v in details.items():
            df.at[idx, k] = v

        print(f"   → {details['status']} | Last Seen: {details['last_seen']}")

        time.sleep(0.5)  # Rate-limit

    return df


def main():
    # Disable SSL warnings if VERIFY_SSL is False
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("🚀 Starting QRadar Log Source Checker with IP fallback and last-event timestamp...")

    # Test connection
    if not test_qradar_connection(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD):
        print("❌ Cannot connect to QRadar. Exiting.")
        return

    # Load Excel
    print(f"\n📖 Reading Excel file: {INPUT_EXCEL_PATH}")
    all_sheets = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=None)
    sheet_names = list(all_sheets.keys())
    print(f"Found sheets: {sheet_names}")

    # Determine which to process
    sheets = sheet_names if SHEETS_TO_PROCESS == ['all'] else SHEETS_TO_PROCESS

    for sheet in sheets:
        if sheet in all_sheets:
            all_sheets[sheet] = process_sheet(
                all_sheets[sheet],
                sheet,
                QRADAR_HOST,
                QRADAR_USERNAME,
                QRADAR_PASSWORD,
                LOGSOURCE_COLUMN,
                IP_COLUMN
            )
        else:
            print(f"⚠️ Sheet '{sheet}' not in workbook; skipping.")

    # Save back
    print(f"\n💾 Saving results to Excel...")
    with pd.ExcelWriter(INPUT_EXCEL_PATH, engine='openpyxl') as writer:
        for name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    print("✅ All done!")

    # Optional summary
    for sheet in sheets:
        df = all_sheets.get(sheet)
        if df is not None and 'status' in df.columns:
            print(f"\n📊 Summary for {sheet}:")
            print(df['status'].value_counts().to_dict())
            print(df['activity_status'].value_counts().to_dict())


if __name__ == '__main__':
    main()
