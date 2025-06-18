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
        'status': '',
        'qradar_id': '',
        'protocol_type': '',
        'enabled': '',
        'last_seen': '',
        'activity_status': '',
        'days_since_last_event': None,
        'might_be_disabled_alert': ''  # Yes/No
    }


def _start_aql_search(qradar_host, username, password, query):
    """
    Start an Ariel (AQL) search. Returns search_id or None.
    """
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
        else:
            print(f"⚠️ AQL search start returned status {resp.status_code}: {resp.text}")
    except Exception as e:
        print(f"⚠️ Exception starting AQL search: {e}")
    return None


def _get_search_results(qradar_host, username, password, search_id, poll_interval=2, max_polls=15):
    """
    Poll for Ariel search results. Returns the JSON dict when 'events' found or empty dict.
    """
    endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches/{search_id}/results"
    for attempt in range(max_polls):
        time.sleep(poll_interval)
        try:
            resp = requests.get(
                endpoint,
                auth=(username, password),
                verify=VERIFY_SSL,
                timeout=30,
                headers={'Accept': 'application/json'}
            )
            if resp.status_code == 200:
                data = resp.json()
                # If search still running, QRadar may return status but no 'events' yet.
                if 'events' in data:
                    return data
                # Else keep polling
            else:
                print(f"⚠️ Polling AQL search {search_id}, status {resp.status_code}")
        except Exception as e:
            print(f"⚠️ Exception polling AQL search: {e}")
    print(f"⚠️ AQL search {search_id} did not return events after polling.")
    return {}


def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    """
    Lookup a log source by name or IP and return its details,
    including the actual last event timestamp, days since last event,
    and a flag if inactive beyond threshold.
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

        # Assume first match is desired
        log_source = ls_data[0]
        ls_id = log_source.get('id')

        # Build AQL to get the MAX(starttime) for last N days (we choose threshold days, but we query a bit larger window if desired).
        # Here we query LAST INACTIVITY_THRESHOLD_DAYS days; if no events, we return “No events in last N days.”
        # You may choose to query a longer window or entire history; but large windows may be slow.
        aql = f"SELECT MAX(starttime) AS max_starttime FROM events WHERE logsourceid={ls_id} LAST {INACTIVITY_THRESHOLD_DAYS} DAYS"
        search_id = _start_aql_search(qradar_host, username, password, aql)

        last_seen = ''
        activity_status = 'No Activity'
        days_since_last_event = None
        might_be_disabled = 'Yes'  # default if no recent events

        if search_id:
            results = _get_search_results(qradar_host, username, password, search_id)
            events = results.get('events', [])
            if events and events[0].get('max_starttime') is not None:
                epoch_ms = events[0]['max_starttime']
                # Convert epoch (ms) to datetime
                last_dt = datetime.fromtimestamp(epoch_ms / 1000)
                last_seen = last_dt.strftime('%Y-%m-%d %H:%M:%S')
                # Calculate days since:
                delta = datetime.now() - last_dt
                days_since_last_event = delta.days
                # Active if within threshold
                if delta <= timedelta(days=INACTIVITY_THRESHOLD_DAYS):
                    activity_status = 'Active'
                    might_be_disabled = 'No'
                else:
                    activity_status = 'Inactive'
                    might_be_disabled = 'Yes'
            else:
                # No events returned in last threshold days
                last_seen = f'No events in last {INACTIVITY_THRESHOLD_DAYS} days'
                activity_status = 'Inactive'
                days_since_last_event = None
                might_be_disabled = 'Yes'
        else:
            # Could not start/search AQL
            last_seen = 'AQL Error/Timeout'
            activity_status = 'Unknown'
            days_since_last_event = None
            might_be_disabled = 'Unknown'

        return {
            'status': 'Found',
            'qradar_id': ls_id,
            'protocol_type': log_source.get('protocol_type', ''),
            'enabled': log_source.get('enabled', ''),
            'last_seen': last_seen,
            'activity_status': activity_status,
            'days_since_last_event': days_since_last_event,
            'might_be_disabled_alert': might_be_disabled
        }

    except Exception as e:
        return {'status': f'Error: {e}', **_empty_details()}


def get_qradar_health_metrics(qradar_host, username, password):
    """
    Example: Fetch QRadar system health metric definitions and current values.
    You may refine which metrics to retrieve (e.g., CPU, Memory, Disk, etc.)
    """
    base = qradar_host.rstrip('/')
    # 1) List all system metrics
    sys_metrics_endpoint = f"{base}/api/health/metrics/system_metrics"
    try:
        resp = requests.get(
            sys_metrics_endpoint,
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=30,
            headers={'Accept': 'application/json'}
        )
        if resp.status_code != 200:
            print(f"⚠️ Unable to list system metrics: {resp.status_code}")
            return {}
        # Sample response: list of metric definitions, each with 'id', 'name', 'description', etc.
        metrics_list = resp.json()
    except Exception as e:
        print(f"⚠️ Exception fetching system_metrics list: {e}")
        return {}

    # 2) For demo: fetch first N metrics or filter by name keyword (e.g., 'CPU' or 'Memory')
    #    In practice, pick the IDs you need.
    results = {}
    for metric in metrics_list:
        metric_id = metric.get('id')
        metric_name = metric.get('name')
        # Example filter: only CPU or Memory metrics (adjust as needed)
        if metric_name and any(keyword in metric_name.lower() for keyword in ['cpu', 'memory', 'disk', 'uptime']):
            try:
                m_resp = requests.get(
                    f"{sys_metrics_endpoint}/{metric_id}",
                    auth=(username, password),
                    verify=VERIFY_SSL,
                    timeout=30,
                    headers={'Accept': 'application/json'}
                )
                if m_resp.status_code == 200:
                    # The returned JSON may include current value(s), timestamp, etc.
                    results[metric_name] = m_resp.json()
                else:
                    print(f"⚠️ Failed to fetch metric {metric_id}: {m_resp.status_code}")
            except Exception as e:
                print(f"⚠️ Exception fetching metric {metric_id}: {e}")
        # Optionally break after some metrics if too many
        # if len(results) >= 5:
        #     break

    return results


def process_sheet(df, sheet_name, qradar_host, username, password, logsource_column, ip_column):
    """Process a single DataFrame sheet, with fallback to IP lookup."""
    print(f"\n📋 Processing sheet: {sheet_name}")

    # Ensure the required columns exist
    for col in [logsource_column, ip_column]:
        if col not in df.columns:
            print(f"❌ Column '{col}' not found in {sheet_name}! Available columns: {list(df.columns)}")
            return df

    # Prepare result columns
    for col in ['status', 'qradar_id', 'protocol_type', 'enabled', 'last_seen',
                'activity_status', 'days_since_last_event', 'might_be_disabled_alert']:
        if col not in df.columns:
            df[col] = ''

    total = len(df)
    print(f"Found {total} rows to process...")

    for idx, row in df.iterrows():
        name_val = str(row[logsource_column]).strip()
        details = None

        # Try lookup by name first
        if name_val and name_val.lower() not in ['nan', 'none', '']:
            print(f"[{idx+1}/{total}] Lookup by name: {name_val}")
            details = get_log_source_details(qradar_host, username, password, name_val, is_ip=False)

        # If not found or error, fallback to IP
        if not details or details.get('status') == 'Not Found':
            ip_val = str(row[ip_column]).strip() if ip_column in df.columns else ''
            if ip_val and ip_val.lower() not in ['nan', 'none', '']:
                print(f"   🔁 Name not found; fallback to IP: {ip_val}")
                details = get_log_source_details(qradar_host, username, password, ip_val, is_ip=True)

        # If still no details
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

    print("🚀 Starting QRadar Log Source Checker with IP fallback, last-event timestamp, inactivity flag, and health metrics...")

    # Test connection
    if not test_qradar_connection(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD):
        print("❌ Cannot connect to QRadar. Exiting.")
        return

    # Optionally: fetch and print some health metrics
    print("\n🔍 Fetching sample QRadar health metrics (CPU/Memory/etc)...")
    health_metrics = get_qradar_health_metrics(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD)
    if health_metrics:
        for name, data in health_metrics.items():
            # The structure depends on QRadar version; print or log key parts
            print(f"Metric: {name} -> {data}")
    else:
        print("⚠️ No health metrics retrieved or no matching filters.")

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
            print("Status counts:", df['status'].value_counts().to_dict())
            if 'activity_status' in df.columns:
                print("Activity status counts:", df['activity_status'].value_counts().to_dict())
            if 'might_be_disabled_alert' in df.columns:
                print("Might-be-disabled flag counts:", df['might_be_disabled_alert'].value_counts().to_dict())


if __name__ == '__main__':
    main()
