import pandas as pd
import requests
import urllib3
from datetime import datetime, timedelta
import time
import os
import json
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
ACTIVITY_THRESHOLD_DAYS = 7  # Consider log source inactive if no events in X days
REQUEST_TIMEOUT = 30
MAX_SEARCH_RETRIES = 20  # Increased from 10
SEARCH_RETRY_DELAY = 3   # Increased from 2
# ─── END CONFIGURATION ─────────────────────────────────────────────────────────


def test_qradar_connection(qradar_host, username, password):
    """Test QRadar connection and validate credentials"""
    print("🔗 Testing QRadar connection...")
    qradar_host = qradar_host.rstrip('/')
    endpoint = f"{qradar_host}/api/help/versions"
    
    try:
        resp = requests.get(
            endpoint,
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'}  # Added API version
        )
        
        if resp.status_code == 200:
            print("✅ QRadar connection successful!")
            version_info = resp.json()
            print(f"   QRadar Version: {version_info[0].get('version', 'Unknown')}")
            return True
        elif resp.status_code == 401:
            print("❌ Authentication failed! Check username/password.")
            return False
        else:
            print(f"⚠️ Unexpected response: {resp.status_code} - {resp.text}")
            return False
            
    except requests.exceptions.Timeout:
        print("❌ Connection timeout! Check network connectivity.")
        return False
    except requests.exceptions.ConnectionError:
        print("❌ Connection error! Check QRadar host URL.")
        return False
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return False


def _empty_details():
    """Return empty details structure"""
    return {
        'qradar_id': '',
        'protocol_type': '',
        'enabled': '',
        'last_seen': '',
        'activity_status': '',
        'event_count_30d': 0,
        'average_eps': 0
    }


def _start_aql_search(qradar_host, username, password, query):
    """Start an AQL search and return search ID"""
    endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches"
    
    try:
        print(f"   🔍 Executing AQL: {query}")
        
        resp = requests.post(
            endpoint,
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'Version': '14.0'  # Added API version
            },
            json={'query_expression': query}
        )
        
        if resp.status_code == 201:
            search_data = resp.json()
            search_id = search_data.get('search_id')
            print(f"   ✅ Search started with ID: {search_id}")
            return search_id
        else:
            print(f"   ❌ Failed to start search: {resp.status_code} - {resp.text}")
            return None
            
    except Exception as e:
        print(f"   ❌ AQL search error: {e}")
        return None


def _get_search_results(qradar_host, username, password, search_id):
    """Get search results with improved error handling and status checking"""
    endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches/{search_id}/results"
    status_endpoint = f"{qradar_host.rstrip('/')}/api/ariel/searches/{search_id}"
    
    print(f"   ⏳ Waiting for search {search_id} to complete...")
    
    for attempt in range(MAX_SEARCH_RETRIES):
        try:
            # Check search status first
            status_resp = requests.get(
                status_endpoint,
                auth=(username, password),
                verify=VERIFY_SSL,
                timeout=REQUEST_TIMEOUT,
                headers={'Accept': 'application/json', 'Version': '14.0'}
            )
            
            if status_resp.status_code == 200:
                status_data = status_resp.json()
                search_status = status_data.get('status', 'UNKNOWN')
                
                print(f"   📊 Search status: {search_status} (attempt {attempt + 1}/{MAX_SEARCH_RETRIES})")
                
                if search_status == 'COMPLETED':
                    # Get results
                    results_resp = requests.get(
                        endpoint,
                        auth=(username, password),
                        verify=VERIFY_SSL,
                        timeout=REQUEST_TIMEOUT,
                        headers={'Accept': 'application/json', 'Version': '14.0'}
                    )
                    
                    if results_resp.status_code == 200:
                        data = results_resp.json()
                        print(f"   ✅ Search completed successfully")
                        return data
                    else:
                        print(f"   ❌ Failed to get results: {results_resp.status_code}")
                        return {}
                        
                elif search_status == 'ERROR':
                    print(f"   ❌ Search failed with error")
                    return {}
                
                elif search_status in ['WAIT', 'EXECUTE', 'SORTING']:
                    time.sleep(SEARCH_RETRY_DELAY)
                    continue
                else:
                    print(f"   ⚠️ Unknown search status: {search_status}")
                    time.sleep(SEARCH_RETRY_DELAY)
                    
            else:
                print(f"   ❌ Failed to check search status: {status_resp.status_code}")
                time.sleep(SEARCH_RETRY_DELAY)
                
        except Exception as e:
            print(f"   ❌ Error checking search results: {e}")
            time.sleep(SEARCH_RETRY_DELAY)
    
    print(f"   ⏰ Search timed out after {MAX_SEARCH_RETRIES} attempts")
    return {}


def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    """
    Enhanced log source lookup with better activity detection
    """
    filter_key = 'ip_address' if is_ip else 'name'
    query_filter = f'{filter_key}="{identifier}"'
    ls_endpoint = f"{qradar_host.rstrip('/')}/api/config/event_sources/log_source_management/log_sources"

    try:
        # Get log source details
        resp = requests.get(
            ls_endpoint,
            params={'filter': query_filter},
            auth=(username, password),
            verify=VERIFY_SSL,
            timeout=REQUEST_TIMEOUT,
            headers={'Accept': 'application/json', 'Version': '14.0'}
        )
        
        if resp.status_code != 200:
            print(f"   ❌ Log source API error: {resp.status_code} - {resp.text}")
            return {'status': f'API Error {resp.status_code}', **_empty_details()}

        ls_data = resp.json()
        if not ls_data:
            return {'status': 'Not Found', **_empty_details()}

        log_source = ls_data[0]
        ls_id = log_source.get('id')
        ls_name = log_source.get('name', identifier)
        
        print(f"   📋 Found log source: {ls_name} (ID: {ls_id})")

        # Multiple AQL queries for comprehensive activity check
        queries = [
            # Query 1: Last event timestamp and count for last 30 days
            f"SELECT MAX(starttime) as last_event, COUNT(*) as event_count FROM events WHERE logsourceid={ls_id} LAST 30 DAYS",
            
            # Query 2: Recent activity check (last 7 days)
            f"SELECT COUNT(*) as recent_count FROM events WHERE logsourceid={ls_id} LAST {ACTIVITY_THRESHOLD_DAYS} DAYS"
        ]
        
        last_seen = ''
        activity_status = 'No Activity'
        event_count_30d = 0
        recent_event_count = 0
        
        # Execute first query (last event + 30-day count)
        search_id = _start_aql_search(qradar_host, username, password, queries[0])
        if search_id:
            results = _get_search_results(qradar_host, username, password, search_id)
            events = results.get('events', [])
            
            if events and len(events) > 0:
                event_data = events[0]
                
                # Get last event timestamp
                last_event_ms = event_data.get('last_event')
                if last_event_ms and last_event_ms != 'NULL':
                    try:
                        last_seen = datetime.fromtimestamp(last_event_ms / 1000).strftime('%Y-%m-%d %H:%M:%S')
                        
                        # Check if recent enough to be considered active
                        last_event_time = datetime.fromtimestamp(last_event_ms / 1000)
                        threshold_time = datetime.now() - timedelta(days=ACTIVITY_THRESHOLD_DAYS)
                        
                        if last_event_time > threshold_time:
                            activity_status = 'Active'
                        else:
                            activity_status = 'Inactive'
                            
                    except (ValueError, TypeError) as e:
                        print(f"   ⚠️ Error parsing timestamp: {e}")
                        last_seen = 'Invalid timestamp'
                
                # Get event count
                event_count_30d = event_data.get('event_count', 0)
                if event_count_30d == 'NULL':
                    event_count_30d = 0
                    
            else:
                last_seen = 'No events in last 30 days'
                event_count_30d = 0
        
        # Execute second query (recent activity)
        search_id = _start_aql_search(qradar_host, username, password, queries[1])
        if search_id:
            results = _get_search_results(qradar_host, username, password, search_id)
            events = results.get('events', [])
            
            if events and len(events) > 0:
                recent_event_count = events[0].get('recent_count', 0)
                if recent_event_count == 'NULL':
                    recent_event_count = 0
        
        # Calculate average EPS over 30 days
        average_eps = round(event_count_30d / (30 * 24 * 3600), 2) if event_count_30d > 0 else 0
        
        # Final activity determination
        if recent_event_count > 0:
            activity_status = 'Active'
        elif event_count_30d > 0:
            activity_status = 'Inactive'
        else:
            activity_status = 'No Activity'
            
        print(f"   📊 Events: {event_count_30d} (30d), {recent_event_count} ({ACTIVITY_THRESHOLD_DAYS}d), Status: {activity_status}")

        return {
            'status': 'Found',
            'qradar_id': ls_id,
            'protocol_type': log_source.get('protocol_type', ''),
            'enabled': log_source.get('enabled', ''),
            'last_seen': last_seen,
            'activity_status': activity_status,
            'event_count_30d': event_count_30d,
            'average_eps': average_eps
        }

    except Exception as e:
        print(f"   ❌ Unexpected error: {e}")
        return {'status': f'Error: {str(e)[:50]}...', **_empty_details()}


def process_sheet(df, sheet_name, qradar_host, username, password, logsource_column, ip_column):
    """Process a single sheet with enhanced logging and error handling"""
    print(f"\n📋 Processing sheet: {sheet_name}")
    
    # Validate columns exist
    missing_cols = []
    for col in [logsource_column, ip_column]:
        if col not in df.columns:
            missing_cols.append(col)
    
    if missing_cols:
        print(f"❌ Missing columns in {sheet_name}: {missing_cols}")
        print(f"   Available columns: {list(df.columns)}")
        return df
    
    # Add result columns if they don't exist
    result_columns = ['status', 'qradar_id', 'protocol_type', 'enabled', 'last_seen', 'activity_status', 'event_count_30d', 'average_eps']
    for col in result_columns:
        if col not in df.columns:
            df[col] = ''
    
    total = len(df)
    processed = 0
    found_count = 0
    
    print(f"Found {total} rows to process...")
    
    for idx, row in df.iterrows():
        processed += 1
        print(f"\n[{processed}/{total}] Processing row {idx + 1}")
        
        name_val = str(row[logsource_column]).strip()
        details = None
        
        # Try lookup by name first
        if name_val and name_val.lower() not in ['nan', 'none', '', 'null']:
            print(f"   🔍 Lookup by name: '{name_val}'")
            details = get_log_source_details(qradar_host, username, password, name_val, is_ip=False)
        
        # Fallback to IP lookup if name lookup failed
        if not details or details['status'] == 'Not Found':
            ip_val = str(row[ip_column]).strip()
            if ip_val and ip_val.lower() not in ['nan', 'none', '', 'null']:
                print(f"   🔁 Fallback to IP: '{ip_val}'")
                details = get_log_source_details(qradar_host, username, password, ip_val, is_ip=True)
        
        # Use empty details if nothing found
        if not details:
            details = {'status': 'Empty/Invalid', **_empty_details()}
        
        # Update DataFrame
        for k, v in details.items():
            df.at[idx, k] = v
        
        if details['status'] == 'Found':
            found_count += 1
            
        print(f"   ✅ Result: {details['status']} | Activity: {details['activity_status']} | Events: {details.get('event_count_30d', 0)}")
        
        # Add delay to avoid overwhelming QRadar
        time.sleep(0.5)
    
    print(f"\n📊 Sheet {sheet_name} completed: {found_count}/{total} log sources found")
    return df


def create_outlook_draft(attachment_path, subject, body):
    """Create Outlook draft with error handling"""
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.Body = body
        mail.Attachments.Add(attachment_path)
        mail.Display()  # Pop up the draft window
        print(f"✉️ Draft created and displayed: {attachment_path}")
    except Exception as e:
        print(f"❌ Failed to create Outlook draft: {e}")
        print(f"   Email would have been: {subject}")
        print(f"   Attachment: {attachment_path}")


def filter_and_email(df_dict, draft_path):
    """Filter inactive/error log sources and create email draft"""
    frames = []
    
    for name, df in df_dict.items():
        if 'status' in df.columns and 'activity_status' in df.columns:
            # Filter inactive sources
            mask_inactive = (df['activity_status'] == 'Inactive') | (df['activity_status'] == 'No Activity')
            
            # Filter API errors
            mask_errors = df['status'].str.startswith('API Error', na=False)
            
            # Filter not found
            mask_not_found = df['status'] == 'Not Found'

            # Add inactive sources
            if mask_inactive.any():
                sub = df[mask_inactive].copy()
                sub['remark'] = 'Inactive or no activity detected'
                sub['sheet_name'] = name
                frames.append(sub)

            # Add error sources
            if mask_errors.any():
                sub_err = df[mask_errors].copy()
                sub_err['remark'] = 'API error - check log source configuration'
                sub_err['sheet_name'] = name
                frames.append(sub_err)
                
            # Add not found sources
            if mask_not_found.any():
                sub_nf = df[mask_not_found].copy()
                sub_nf['remark'] = 'Log source not found in QRadar'
                sub_nf['sheet_name'] = name
                frames.append(sub_nf)

    if not frames:
        print("✅ No inactive, error, or not found log sources; skipping email.")
        return

    result_df = pd.concat(frames, ignore_index=True)
    total = len(result_df)
    inactive_count = ((result_df['activity_status'] == 'Inactive') | (result_df['activity_status'] == 'No Activity')).sum()
    error_count = result_df['status'].str.startswith('API Error', na=False).sum()
    not_found_count = (result_df['status'] == 'Not Found').sum()

    # Save filtered results
    result_df.to_excel(draft_path, index=False)
    print(f"💾 Filtered report saved to: {draft_path}")

    # Create email
    subject = f"QRadar Log Source Status Report - {total} Issues Found"
    body = f"""Hello,

Attached is the QRadar log source status report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.

Summary:
- Total flagged log sources: {total}
- Inactive/No Activity: {inactive_count}
- API Errors: {error_count}
- Not Found: {not_found_count}

Please review the attached Excel file for detailed information and take appropriate action.

Activity threshold: {ACTIVITY_THRESHOLD_DAYS} days
Report includes log sources with no events in the last {ACTIVITY_THRESHOLD_DAYS} days or API/lookup errors.

Best regards,
QRadar Automation System
"""
    
    create_outlook_draft(draft_path, subject, body)


def main():
    """Main execution function"""
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("🚀 Starting QRadar Log Source Checker (Enhanced Version)...")
    print(f"⚙️ Configuration:")
    print(f"   - QRadar Host: {QRADAR_HOST}")
    print(f"   - Activity Threshold: {ACTIVITY_THRESHOLD_DAYS} days")
    print(f"   - Request Timeout: {REQUEST_TIMEOUT}s")
    print(f"   - SSL Verification: {VERIFY_SSL}")
    
    # Test connection
    if not test_qradar_connection(QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD):
        print("❌ Connection test failed. Please check your configuration.")
        return

    # Read Excel file
    print(f"\n📖 Reading Excel file: {INPUT_EXCEL_PATH}")
    try:
        all_sheets = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=None)
        sheets = list(all_sheets.keys())
        print(f"📄 Sheets found: {sheets}")
    except Exception as e:
        print(f"❌ Failed to read Excel file: {e}")
        return

    # Process sheets
    to_process = sheets if SHEETS_TO_PROCESS == ['all'] else SHEETS_TO_PROCESS
    print(f"📋 Sheets to process: {to_process}")
    
    for sheet in to_process:
        if sheet in all_sheets:
            print(f"\n{'='*50}")
            all_sheets[sheet] = process_sheet(
                all_sheets[sheet], sheet,
                QRADAR_HOST, QRADAR_USERNAME, QRADAR_PASSWORD,
                LOGSOURCE_COLUMN, IP_COLUMN
            )
        else:
            print(f"⚠️ Sheet '{sheet}' not found in Excel file. Skipping...")

    # Save updated Excel
    print(f"\n💾 Saving updates to original Excel file...")
    try:
        with pd.ExcelWriter(INPUT_EXCEL_PATH, engine='openpyxl') as writer:
            for name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)
        print("✅ Original Excel file updated successfully.")
    except Exception as e:
        print(f"❌ Failed to save Excel file: {e}")

    # Generate filtered report and email
    print(f"\n📧 Generating filtered report...")
    filter_and_email(all_sheets, DRAFT_OUTPUT_PATH)

    # Final summary
    print(f"\n📊 FINAL SUMMARY:")
    print("=" * 60)
    
    total_processed = 0
    total_found = 0
    total_active = 0
    total_inactive = 0
    total_errors = 0
    
    for sheet in to_process:
        if sheet in all_sheets:
            df = all_sheets[sheet]
            if 'status' in df.columns:
                sheet_total = len(df)
                sheet_found = (df['status'] == 'Found').sum()
                sheet_active = (df['activity_status'] == 'Active').sum()
                sheet_inactive = ((df['activity_status'] == 'Inactive') | (df['activity_status'] == 'No Activity')).sum()
                sheet_errors = (df['status'].str.startswith('API Error', na=False) | (df['status'] == 'Not Found')).sum()
                
                print(f"📋 {sheet}:")
                print(f"   Total: {sheet_total}, Found: {sheet_found}, Active: {sheet_active}, Inactive: {sheet_inactive}, Errors: {sheet_errors}")
                
                total_processed += sheet_total
                total_found += sheet_found
                total_active += sheet_active
                total_inactive += sheet_inactive
                total_errors += sheet_errors
    
    print("=" * 60)
    print(f"🎯 OVERALL TOTALS:")
    print(f"   Processed: {total_processed}")
    print(f"   Found: {total_found}")
    print(f"   Active: {total_active}")
    print(f"   Inactive: {total_inactive}")
    print(f"   Errors: {total_errors}")
    
    if total_processed > 0:
        success_rate = (total_found / total_processed) * 100
        active_rate = (total_active / total_found) * 100 if total_found > 0 else 0
        print(f"   Success Rate: {success_rate:.1f}%")
        print(f"   Active Rate: {active_rate:.1f}%")
    
    print(f"\n✅ QRadar Log Source Checker completed!")
    print(f"📁 Updated Excel: {INPUT_EXCEL_PATH}")
    print(f"📧 Filtered Report: {DRAFT_OUTPUT_PATH}")


if __name__ == '__main__':
    main()
