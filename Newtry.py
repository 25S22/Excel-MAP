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

# 6) SSL verification (set to False for testing; in prod, you should verify SSL):
VERIFY_SSL = False

# 7) Inactivity threshold (days) to consider a log source “might be disabled”
INACTIVITY_THRESHOLD_DAYS = 30  # e.g., 30 days ~ 1 month

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
        print(f"  → GET {test_endpoint} returned {resp.status_code}")
        if resp.status_code == 200:
            print("✅ QRadar connection successful!")
            return True
        elif resp.status_code == 401:
            print("❌ Authentication failed! Check username/password.")
            try:
                print("    Response JSON:", resp.json())
            except:
                print("    Response text:", resp.text)
            return False
        else:
            print(f"⚠️ Unexpected response: {resp.status_code} - {resp.text}")
            return False
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False


def _empty_details():
    return {
        'status': '',
        'qradar_id': '',
        'protocol_type': '',
        'enabled': '',
        'last_seen': '',
        'activity_status': '',
        'days_since_last_event': None,
        'might_be_disabled_alert': ''  # Yes/No/Unknown
    }


def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    """
    Lookup a log source by name or IP and return its details,
    including the last_event_time from config API, days since last event,
    and a flag if inactive beyond threshold. No AQL used.
    Uses GET /api/config/event_sources/log_source_management/log_sources?filter=... .
    Expects the API to return 'last_event_time' in milliseconds since epoch. 2
    """
    filter_key = 'ip_address' if is_ip else 'name'
    query_filter = f'{filter_key}="{identifier}"'
    ls_endpoint = f"{qradar_host.rstrip('/')}/api/config/event_sources/log_source_management/log_sources"

    try:
        # Correct: exactly one 'filter' parameter
        resp = requests.get(
            ls_endpoint,
            params={'filter': query_filter},
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=30,
            headers={'Accept': 'application/json'}
        )
        # Print for debugging; you can reduce verbosity if desired
        print(f"  → GET {ls_endpoint}?filter={query_filter} returned {resp.status_code}")
        if resp.status_code != 200:
            if resp.status_code == 422:
                try:
                    print("    422 response details:", resp.json())
                except Exception:
                    print("    422 response (non-JSON):", resp.text)
            return {'status': f'API Error {resp.status_code}', **_empty_details()}

        ls_data = resp.json()
        if not isinstance(ls_data, list) or not ls_data:
            return {'status': 'Not Found', **_empty_details()}

        # Use first matching log source
        log_source = ls_data[0]
        ls_id = log_source.get('id')
        if ls_id is None:
            return {'status': 'No ID in response', **_empty_details()}

        # Extract last_event_time if present (milliseconds since epoch)
        last_event_ms = log_source.get('last_event_time')
        if last_event_ms is None:
            # If the field is missing or null, treat as no events seen or unknown
            last_seen = f'No last_event_time field'
            activity_status = 'Unknown'
            days_since_last_event = None
            might_be_disabled = 'Unknown'
        else:
            try:
                last_dt = datetime.fromtimestamp(last_event_ms / 1000.0)
                last_seen = last_dt.strftime('%Y-%m-%d %H:%M:%S')
                delta = datetime.now() - last_dt
                days_since_last_event = delta.days
                if delta.days <= INACTIVITY_THRESHOLD_DAYS:
                    activity_status = 'Active'
                    might_be_disabled = 'No'
                else:
                    activity_status = 'Inactive'
                    might_be_disabled = 'Yes'
            except Exception as e:
                print(f"    ⚠️ Error parsing last_event_time: {e}")
                last_seen = 'Error parsing last_event_time'
                activity_status = 'Unknown'
                days_since_last_event = None
                might_be_disabled = 'Unknown'

        # You can also read additional fields if needed, e.g., protocol_type, enabled, etc.
        protocol_type = log_source.get('protocol_type', '')
        enabled = log_source.get('enabled', '')

        return {
            'status': 'Found',
            'qradar_id': ls_id,
            'protocol_type': protocol_type,
            'enabled': enabled,
            'last_seen': last_seen,
            'activity_status': activity_status,
            'days_since_last_event': days_since_last_event,
            'might_be_disabled_alert': might_be_disabled
        }

    except Exception as e:
        print(f"  ⚠️ Exception in get_log_source_details for identifier={identifier}: {e}")
        return {'status': f'Error: {e}', **_empty_details()}


def process_sheet(df, sheet_name, qradar_host, username, password, logsource_column, ip_column):
    """Process a single DataFrame sheet, with fallback to IP lookup."""
    print(f"\n📋 Processing sheet: {sheet_name}")

    # Ensure the required columns exist
    for col in [logsource_column, ip_column]:
        if col not in df.columns:
            print(f"❌ Column '{col}' not found in {sheet_name}! Available columns: {list(df.columns)}")
            return df

    # Prepare result columns if not present
    for col in ['status', 'qradar_id', 'protocol_type', 'enabled', 'last_seen',
                'activity_status', 'days_since_last_event', 'might_be_disabled_alert']:
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

        if not details or details.get('status') == 'Not Found':
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
                'activity_status': '',
                'days_since_last_event': None,
                'might_be_disabled_alert': ''
            }

        # Write back to DataFrame
        for k, v in details.items():
            df.at[idx, k] = v

        print(f"   → {details.get('status')} | Last Seen: {details.get('last_seen')} | MightBeDisabled: {details.get('might_be_disabled_alert')}")

        time.sleep(0.5)  # Rate-limit

    return df


def main():
    # Disable SSL warnings if VERIFY_SSL is False
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("🚀 Starting QRadar Log Source Checker using last_event_time (no AQL)...")

    # Test connection
    if not test_qradar_connection(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD):
        print("❌ Cannot connect to QRadar. Exiting.")
        return

    # Load Excel
    print(f"\n📖 Reading Excel file: {INPUT_EXCEL_PATH}")
    try:
        all_sheets = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=None)
    except Exception as e:
        print(f"❌ Failed to read Excel: {e}")
        return

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
    try:
        with pd.ExcelWriter(INPUT_EXCEL_PATH, engine='openpyxl') as writer:
            for name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)
        print("✅ All done!")
    except Exception as e:
        print(f"❌ Failed to write Excel: {e}")

    # Optional summary
    for sheet in sheets:
        df = all_sheets.get(sheet)
        if df is not None and 'status' in df.columns:
            print(f"\n📊 Summary for {sheet}:")
            print("  Status counts:", df['status'].value_counts().to_dict())
            if 'activity_status' in df.columns:
                print("  Activity status counts:", df['activity_status'].value_counts().to_dict())
            if 'might_be_disabled_alert' in df.columns:
                print("  Might-be-disabled flag counts:", df['might_be_disabled_alert'].value_counts().to_dict())


if __name__ == '__main__':
    main()
