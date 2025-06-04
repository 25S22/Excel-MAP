import pandas as pd
import requests
import urllib3

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────

# 1) Path to your existing Excel file (use forward slashes or raw string):
INPUT_EXCEL_PATH = r'C:\path\to\your\input.xlsx'  # Use raw string or forward slashes

# 2) Which sheet in the Excel file to read:
EXCEL_SHEET_NAME = 0  # or 'Sheet1'

# 3) The name of the column containing log source names:
LOGSOURCE_COLUMN = 'log source name'

# 4) QRadar details:
QRADAR_HOST = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'

# 5) SSL verification (set to False for testing):
VERIFY_SSL = False

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────


def get_log_source_details(qradar_host, username, password, log_source_name):
    """Query QRadar for log source details."""
    
    # Clean URL and prepare endpoint
    qradar_host = qradar_host.rstrip('/')
    endpoint = f"{qradar_host}/api/config/event_sources/log_source_management/log_sources"
    
    # Prepare filter
    params = {'filter': f'name="{log_source_name}"'}
    
    # Simple auth using requests built-in basic auth
    try:
        response = requests.get(
            endpoint,
            params=params,
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=30,
            headers={'Accept': 'application/json'}
        )
        
        if response.status_code == 401:
            return {'status': 'Auth Failed', 'qradar_id': '', 'protocol_type': '', 'enabled': ''}
        
        if response.status_code != 200:
            return {'status': f'Error {response.status_code}', 'qradar_id': '', 'protocol_type': '', 'enabled': ''}
        
        data = response.json()
        
        if data and len(data) > 0:
            ls = data[0]
            return {
                'status': 'Found',
                'qradar_id': ls.get('id', ''),
                'protocol_type': ls.get('protocol_type', ''),
                'enabled': ls.get('enabled', '')
            }
        else:
            return {'status': 'Not Found', 'qradar_id': '', 'protocol_type': '', 'enabled': ''}
            
    except Exception as e:
        return {'status': f'Error: {str(e)[:50]}', 'qradar_id': '', 'protocol_type': '', 'enabled': ''}


def main():
    # Disable SSL warnings if needed
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    print("Starting QRadar Log Source Checker...")
    
    # Read Excel file
    print(f"Reading: {INPUT_EXCEL_PATH}")
    df = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=EXCEL_SHEET_NAME)
    
    # Check if column exists
    if LOGSOURCE_COLUMN not in df.columns:
        print(f"Error: Column '{LOGSOURCE_COLUMN}' not found!")
        print(f"Available columns: {list(df.columns)}")
        return
    
    # Add new columns
    df['status'] = ''
    df['qradar_id'] = ''
    df['protocol_type'] = ''
    df['enabled'] = ''
    
    total = len(df)
    print(f"Processing {total} log sources...")
    
    # Process each row
    for idx, row in df.iterrows():
        log_source_name = str(row[LOGSOURCE_COLUMN]).strip()
        
        if not log_source_name or log_source_name == 'nan':
            details = {'status': 'Empty', 'qradar_id': '', 'protocol_type': '', 'enabled': ''}
        else:
            print(f"[{idx+1}/{total}] Checking: {log_source_name}")
            details = get_log_source_details(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD, log_source_name)
        
        # Update DataFrame
        df.at[idx, 'status'] = details['status']
        df.at[idx, 'qradar_id'] = details['qradar_id']
        df.at[idx, 'protocol_type'] = details['protocol_type']
        df.at[idx, 'enabled'] = details['enabled']
        
        print(f"   Result: {details['status']}")
    
    # Save back to Excel
    print("Saving results...")
    df.to_excel(INPUT_EXCEL_PATH, index=False, sheet_name=EXCEL_SHEET_NAME)
    
    print("✅ Done! Results saved to the same Excel file.")
    
    # Show summary
    status_counts = df['status'].value_counts()
    print("\n📊 Summary:")
    for status, count in status_counts.items():
        print(f"   {status}: {count}")


if __name__ == '__main__':
    main()
