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

# 4) QRadar details:
QRADAR_HOST = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'

# 5) SSL verification (set to False for testing):
VERIFY_SSL = False

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────


def test_qradar_connection(qradar_host, username, password):
    """Test if we can connect to QRadar API."""
    print("🔗 Testing QRadar connection...")
    
    qradar_host = qradar_host.rstrip('/')
    test_endpoint = f"{qradar_host}/api/help/versions"
    
    try:
        response = requests.get(
            test_endpoint,
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=10,
            headers={'Accept': 'application/json'}
        )
        
        if response.status_code == 200:
            print("✅ QRadar connection successful!")
            return True
        elif response.status_code == 401:
            print("❌ Authentication failed! Check username/password.")
            return False
        else:
            print(f"⚠️ Unexpected response: {response.status_code}")
            print(f"Response: {response.text[:200]}")
            return False
            
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False


def get_log_source_details(qradar_host, username, password, log_source_name):
    """Get log source details and recent activity from QRadar."""
    
    qradar_host = qradar_host.rstrip('/')
    
    # Step 1: Find the log source
    ls_endpoint = f"{qradar_host}/api/config/event_sources/log_source_management/log_sources"
    
    try:
        # Search for log source by name
        ls_response = requests.get(
            ls_endpoint,
            params={'filter': f'name="{log_source_name}"'},
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=30,
            headers={'Accept': 'application/json'}
        )
        
        if ls_response.status_code != 200:
            return {
                'status': f'API Error {ls_response.status_code}',
                'qradar_id': '',
                'protocol_type': '',
                'enabled': '',
                'last_seen': '',
                'activity_status': ''
            }
        
        ls_data = ls_response.json()
        
        if not ls_data or len(ls_data) == 0:
            return {
                'status': 'Not Found',
                'qradar_id': '',
                'protocol_type': '',
                'enabled': '',
                'last_seen': '',
                'activity_status': 'Not Found'
            }
        
        # Log source found
        log_source = ls_data[0]
        log_source_id = log_source.get('id')
        
        # Step 2: Check recent activity (last 7 days)
        seven_days_ago = int((datetime.now() - timedelta(days=7)).timestamp() * 1000)
        
        # Query for recent events from this log source
        search_endpoint = f"{qradar_host}/api/ariel/searches"
        
        # AQL query to find recent events from this log source
        aql_query = f"SELECT COUNT(*) FROM events WHERE logsourceid={log_source_id} LAST 7 DAYS"
        
        search_response = requests.post(
            search_endpoint,
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=30,
            headers={'Accept': 'application/json', 'Content-Type': 'application/json'},
            json={'query_expression': aql_query}
        )
        
        last_seen = "Unknown"
        activity_status = "Unknown"
        
        if search_response.status_code == 201:
            search_data = search_response.json()
            search_id = search_data.get('search_id')
            
            # Wait for search to complete and get results
            time.sleep(2)  # Give search time to complete
            
            results_endpoint = f"{qradar_host}/api/ariel/searches/{search_id}/results"
            results_response = requests.get(
                results_endpoint,
                auth=(username, password),
                verify=VERIFY_SSL,
                timeout=30,
                headers={'Accept': 'application/json'}
            )
            
            if results_response.status_code == 200:
                results_data = results_response.json()
                if results_data.get('events') and len(results_data['events']) > 0:
                    event_count = results_data['events'][0].get('COUNT(*)', 0)
                    if event_count > 0:
                        activity_status = "Active (Last 7 days)"
                        last_seen = "Within 7 days"
                    else:
                        activity_status = "Inactive (7+ days)"
                        last_seen = "More than 7 days ago"
                else:
                    activity_status = "No Recent Activity"
                    last_seen = "More than 7 days ago"
        
        return {
            'status': 'Found',
            'qradar_id': log_source_id,
            'protocol_type': log_source.get('protocol_type', ''),
            'enabled': log_source.get('enabled', ''),
            'last_seen': last_seen,
            'activity_status': activity_status
        }
        
    except Exception as e:
        return {
            'status': f'Error: {str(e)[:50]}',
            'qradar_id': '',
            'protocol_type': '',
            'enabled': '',
            'last_seen': '',
            'activity_status': 'Error'
        }


def process_sheet(df, sheet_name, qradar_host, username, password, logsource_column):
    """Process a single sheet."""
    print(f"\n📋 Processing sheet: {sheet_name}")
    
    # Check if column exists
    if logsource_column not in df.columns:
        print(f"❌ Column '{logsource_column}' not found in {sheet_name}!")
        print(f"Available columns: {list(df.columns)}")
        return df
    
    # Add new columns if they don't exist
    new_columns = ['status', 'qradar_id', 'protocol_type', 'enabled', 'last_seen', 'activity_status']
    for col in new_columns:
        if col not in df.columns:
            df[col] = ''
    
    total = len(df)
    print(f"Found {total} rows to process...")
    
    # Process each row
    for idx, row in df.iterrows():
        log_source_name = str(row[logsource_column]).strip()
        
        if not log_source_name or log_source_name.lower() in ['nan', 'none', '']:
            details = {
                'status': 'Empty',
                'qradar_id': '',
                'protocol_type': '',
                'enabled': '',
                'last_seen': '',
                'activity_status': 'N/A'
            }
        else:
            print(f"[{idx+1}/{total}] Checking: {log_source_name}")
            details = get_log_source_details(qradar_host, username, password, log_source_name)
        
        # Update DataFrame
        for key, value in details.items():
            df.at[idx, key] = value
        
        print(f"   Result: {details['status']} | Activity: {details.get('activity_status', 'Unknown')}")
        
        # Small delay to avoid overwhelming QRadar
        time.sleep(0.5)
    
    return df


def main():
    # Disable SSL warnings if needed
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    print("🚀 Starting QRadar Log Source Checker with Activity Detection...")
    
    # Test connection first
    if not test_qradar_connection(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD):
        print("❌ Cannot connect to QRadar. Please check your configuration.")
        return
    
    # Read Excel file with all sheets
    print(f"\n📖 Reading Excel file: {INPUT_EXCEL_PATH}")
    
    try:
        # Read all sheets to preserve them
        all_sheets = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=None)  # This reads ALL sheets
        print(f"Found sheets: {list(all_sheets.keys())}")
        
        # Determine which sheets to process
        if SHEETS_TO_PROCESS == ['all']:
            sheets_to_process = list(all_sheets.keys())
        else:
            sheets_to_process = SHEETS_TO_PROCESS
        
        print(f"Will process sheets: {sheets_to_process}")
        
        # Process each specified sheet
        for sheet_name in sheets_to_process:
            if sheet_name in all_sheets:
                all_sheets[sheet_name] = process_sheet(
                    all_sheets[sheet_name], 
                    sheet_name, 
                    QRADAR_HOST, 
                    QRADAR_USERNAME, 
                    QRADAR_PASSWORD, 
                    LOGSOURCE_COLUMN
                )
            else:
                print(f"⚠️ Sheet '{sheet_name}' not found in Excel file!")
        
        # Save all sheets back to Excel (this preserves all sheets)
        print(f"\n💾 Saving results back to Excel...")
        with pd.ExcelWriter(INPUT_EXCEL_PATH, engine='openpyxl') as writer:
            for sheet_name, sheet_df in all_sheets.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("✅ Done! All sheets preserved, processed sheets updated.")
        
        # Show summary for processed sheets
        for sheet_name in sheets_to_process:
            if sheet_name in all_sheets:
                print(f"\n📊 Summary for {sheet_name}:")
                if 'status' in all_sheets[sheet_name].columns:
                    status_counts = all_sheets[sheet_name]['status'].value_counts()
                    for status, count in status_counts.items():
                        print(f"   {status}: {count}")
                
                if 'activity_status' in all_sheets[sheet_name].columns:
                    print(f"\n🔄 Activity Summary for {sheet_name}:")
                    activity_counts = all_sheets[sheet_name]['activity_status'].value_counts()
                    for activity, count in activity_counts.items():
                        print(f"   {activity}: {count}")
        
    except Exception as e:
        print(f"❌ Error processing Excel file: {e}")
        raise


if __name__ == '__main__':
    main()
