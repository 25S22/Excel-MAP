import pandas as pd
import requests
import urllib3
from datetime import datetime, timedelta
import time
import os
import win32com.client  # For creating draft emails in Outlook

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

# 7) Path to save the filtered Excel draft attachment:
DRAFT_OUTPUT_PATH = os.path.join(os.path.dirname(INPUT_EXCEL_PATH), 'inactive_and_errors.xlsx')

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

# ... existing helper functions (_empty_details, _start_aql_search, _get_search_results, get_log_source_details, process_sheet) remain unchanged ...

# ─── NEW FEATURE: EMAIL DRAFT GENERATION ────────────────────────────────────────

def create_outlook_draft(attachment_path, subject, body):
    """
    Create an Outlook draft email with the given attachment, subject, and body.
    """
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # 0: olMailItem
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(attachment_path)
    mail.Save()  # Saves to Drafts folder
    print(f"✉️ Draft email created with attachment: {attachment_path}")


def filter_and_email(df_dict, draft_path):
    """
    From the processed sheets, extract rows that are inactive or API errors,
    add remark, save to a new Excel, and create an email draft with counts.
    """
    frames = []
    for name, df in df_dict.items():
        if 'status' in df.columns and 'activity_status' in df.columns:
            mask = (
                (df['activity_status'] != 'Active') |
                (df['status'].str.startswith('API Error'))
            )
            subset = df[mask].copy()
            if not subset.empty:
                subset['remark'] = 'Check log source name'
                subset['sheet_name'] = name
                frames.append(subset)

    if frames:
        result_df = pd.concat(frames, ignore_index=True)
        # Calculate counts
        total = len(result_df)
        num_inactive = result_df[result_df['activity_status'] != 'Active'].shape[0]
        num_api_errors = result_df[result_df['status'].str.startswith('API Error')].shape[0]

        # Save filtered rows to new Excel
        result_df.to_excel(draft_path, index=False)
        print(f"💾 Filtered rows saved to: {draft_path}")

        # Prepare email content
        subject = "Inactive Log Sources"
        body = (
            f"Hello,\n\n"
            f"Please find attached the QRadar log source report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.\n"
            f"Summary:\n"
            f"  • Total flagged rows: {total}\n"
            f"  • Inactive sources: {num_inactive}\n"
            f"  • API error sources: {num_api_errors}\n\n"
            "Rows are tagged with 'Check log source name' for your review.\n\n"
            "Best regards,\n"
            "Your QRadar Automation Bot"
        )
        create_outlook_draft(draft_path, subject, body)
    else:
        print("✅ No inactive or API error rows found; no email draft created.")


def main():
    # Disable SSL warnings if VERIFY_SSL is False
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    print("🚀 Starting QRadar Log Source Checker with draft email feature...")

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

    # Save back the original file
    print(f"\n💾 Saving results to Excel..." )
    with pd.ExcelWriter(INPUT_EXCEL_PATH, engine='openpyxl') as writer:
        for name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    print("✅ Original file updated.")

    # New feature: filter and create draft email
    filter_and_email(all_sheets, DRAFT_OUTPUT_PATH)

    # Optional summary
    for sheet in sheets:
        df = all_sheets.get(sheet)
        if df is not None and 'status' in df.columns:
            print(f"\n📊 Summary for {sheet}:")
            print(df['status'].value_counts().to_dict())
            print(df['activity_status'].value_counts().to_dict())


if __name__ == '__main__':
    main()
