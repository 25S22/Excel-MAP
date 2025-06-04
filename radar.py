import pandas as pd
import requests
import urllib3
import time
from typing import Dict, Any
import logging
from pathlib import Path

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────

# 1) Path to your existing Excel file:
INPUT_EXCEL_PATH = 'input.xlsx'

# 2) Which sheet in the Excel file to read (by name or index):
EXCEL_SHEET_NAME = 0  # or e.g. 'Sheet1'

# 3) The name of the column containing log source names:
LOGSOURCE_COLUMN = 'log source name'

# 4) Full URL (including protocol) of your QRadar console:
QRADAR_HOST = 'https://your-qradar-host'

# 5) Your QRadar API token (generate under Admin → Authorized Services → "Add"):
QRADAR_API_TOKEN = 'YOUR_QRADAR_API_TOKEN_HERE'

# 6) Path to your SSL certificate (PEM) to verify the QRadar TLS connection.
#    If you want to skip verification (not recommended), set VERIFY_SSL=False.
SSL_CERT_PATH = '/path/to/your/qradar_cert.pem'
VERIFY_SSL = True  # set to False ONLY for testing/self‐signed (not recommended in prod)

# 7) Rate limiting - delay between API calls (seconds)
API_DELAY = 0.5  # Adjust based on your QRadar's rate limits

# 8) Retry configuration
MAX_RETRIES = 3
RETRY_DELAY = 2

# ─── END CONFIGURATION ─────────────────────────────────────────────────────────

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('qradar_log_source_check.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def validate_configuration() -> bool:
    """Validate the configuration before running."""
    errors = []
    
    # Check if Excel file exists
    if not Path(INPUT_EXCEL_PATH).exists():
        errors.append(f"Excel file not found: {INPUT_EXCEL_PATH}")
    
    # Check QRadar host format
    if not QRADAR_HOST.startswith(('http://', 'https://')):
        errors.append("QRADAR_HOST must include protocol (http:// or https://)")
    
    # Check API token
    if QRADAR_API_TOKEN == 'YOUR_QRADAR_API_TOKEN_HERE':
        errors.append("Please set your actual QRadar API token")
    
    # Check SSL certificate if verification is enabled
    if VERIFY_SSL and not Path(SSL_CERT_PATH).exists():
        logger.warning(f"SSL certificate not found at {SSL_CERT_PATH}. Consider setting VERIFY_SSL=False for testing.")
    
    if errors:
        for error in errors:
            logger.error(error)
        return False
    
    return True


def get_log_source_details(qradar_host: str,
                           api_token: str,
                           log_source_name: str,
                           ssl_cert: str,
                           verify_ssl: bool,
                           retries: int = 0) -> Dict[str, Any]:
    """
    Queries QRadar for a log source matching exactly `log_source_name`.
    Enhanced with better error handling, retries, and logging.
    """
    # Clean the QRadar host URL (remove trailing slash if present)
    qradar_host = qradar_host.rstrip('/')
    
    endpoint = f"{qradar_host}/api/config/event_sources/log_source_management/log_sources"
    
    # Properly escape the log source name for the filter
    safe_name = log_source_name.replace('"', '\\"').replace("'", "\\'")
    params = {
        'filter': f'name="{safe_name}"'
    }
    
    headers = {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'SEC': api_token,
        'Version': '12.0'  # Specify API version for consistency
    }

    # Determine SSL verification setting
    if verify_ssl and Path(ssl_cert).exists():
        ssl_verify = ssl_cert
    elif verify_ssl:
        ssl_verify = True  # Use default CA bundle
    else:
        ssl_verify = False

    try:
        logger.debug(f"Querying QRadar for log source: {log_source_name}")
        
        resp = requests.get(
            endpoint,
            headers=headers,
            params=params,
            verify=ssl_verify,
            timeout=30  # Increased timeout
        )
        
        # Check for rate limiting
        if resp.status_code == 429:
            logger.warning(f"Rate limited for '{log_source_name}'. Waiting before retry...")
            time.sleep(RETRY_DELAY * 2)  # Wait longer for rate limits
            if retries < MAX_RETRIES:
                return get_log_source_details(qradar_host, api_token, log_source_name, 
                                            ssl_cert, verify_ssl, retries + 1)
        
        resp.raise_for_status()
        
    except requests.exceptions.SSLError as e:
        logger.error(f"SSL ERROR for '{log_source_name}': {e}")
        if retries < MAX_RETRIES:
            logger.info(f"Retrying ({retries + 1}/{MAX_RETRIES})...")
            time.sleep(RETRY_DELAY)
            return get_log_source_details(qradar_host, api_token, log_source_name, 
                                        ssl_cert, verify_ssl, retries + 1)
        return {
            'status': 'SSL Error',
            'qradar_id': '',
            'protocol_type': '',
            'protocol_name': '',
            'enabled': '',
            'description': '',
            'error_details': str(e)
        }
        
    except requests.exceptions.Timeout as e:
        logger.error(f"TIMEOUT ERROR for '{log_source_name}': {e}")
        if retries < MAX_RETRIES:
            logger.info(f"Retrying ({retries + 1}/{MAX_RETRIES})...")
            time.sleep(RETRY_DELAY)
            return get_log_source_details(qradar_host, api_token, log_source_name, 
                                        ssl_cert, verify_ssl, retries + 1)
        return {
            'status': 'Timeout Error',
            'qradar_id': '',
            'protocol_type': '',
            'protocol_name': '',
            'enabled': '',
            'description': '',
            'error_details': str(e)
        }
        
    except requests.exceptions.RequestException as e:
        logger.error(f"REQUEST ERROR for '{log_source_name}': {e}")
        if retries < MAX_RETRIES:
            logger.info(f"Retrying ({retries + 1}/{MAX_RETRIES})...")
            time.sleep(RETRY_DELAY)
            return get_log_source_details(qradar_host, api_token, log_source_name, 
                                        ssl_cert, verify_ssl, retries + 1)
        return {
            'status': 'Request Error',
            'qradar_id': '',
            'protocol_type': '',
            'protocol_name': '',
            'enabled': '',
            'description': '',
            'error_details': str(e)
        }

    try:
        data = resp.json()
    except ValueError as e:
        logger.error(f"JSON decode error for '{log_source_name}': {e}")
        return {
            'status': 'JSON Error',
            'qradar_id': '',
            'protocol_type': '',
            'protocol_name': '',
            'enabled': '',
            'description': '',
            'error_details': 'Invalid JSON response'
        }

    # Handle the response data
    if isinstance(data, list) and len(data) > 0:
        # Take the first match (assuming names should be unique)
        ls = data[0]
        
        # Log if multiple matches found
        if len(data) > 1:
            logger.warning(f"Multiple log sources found for '{log_source_name}'. Using first match (ID: {ls.get('id')})")
        
        return {
            'status': 'Exists',
            'qradar_id': ls.get('id', ''),
            'protocol_type': ls.get('protocol_type', ''),
            'protocol_name': ls.get('protocol_name', ''),
            'enabled': ls.get('enabled', ''),
            'description': ls.get('description', '') or '',
            'error_details': ''
        }
    else:
        return {
            'status': 'Not Found',
            'qradar_id': '',
            'protocol_type': '',
            'protocol_name': '',
            'enabled': '',
            'description': '',
            'error_details': ''
        }


def backup_excel_file(file_path: str) -> str:
    """Create a backup of the Excel file before modifying it."""
    backup_path = f"{Path(file_path).stem}_backup{Path(file_path).suffix}"
    import shutil
    shutil.copy2(file_path, backup_path)
    logger.info(f"Backup created: {backup_path}")
    return backup_path


def main():
    logger.info("Starting QRadar Log Source Checker")
    
    # Validate configuration
    if not validate_configuration():
        logger.error("Configuration validation failed. Please check your settings.")
        return
    
    # Suppress only InsecureRequestWarning if VERIFY_SSL=False
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        logger.warning("SSL verification disabled - not recommended for production")

    try:
        # Create backup
        backup_path = backup_excel_file(INPUT_EXCEL_PATH)
        
        # Step 1: Read the Excel file into a DataFrame
        logger.info(f"Reading Excel file: {INPUT_EXCEL_PATH}")
        df = pd.read_excel(INPUT_EXCEL_PATH, sheet_name=EXCEL_SHEET_NAME)
        logger.info(f"Loaded {len(df)} rows from Excel file")

        # Step 2: Ensure the "log source name" column exists
        if LOGSOURCE_COLUMN not in df.columns:
            available_cols = ', '.join(df.columns.tolist())
            raise KeyError(f"Column '{LOGSOURCE_COLUMN}' not found. Available columns: {available_cols}")

        # Step 3: Prepare new columns
        new_cols = ['status', 'qradar_id', 'protocol_type', 'protocol_name', 'enabled', 'description', 'error_details']
        for col in new_cols:
            if col not in df.columns:
                df[col] = ''
            else:
                logger.info(f"Column '{col}' already exists - will be overwritten")

        total = len(df)
        successful = 0
        failed = 0
        
        # Step 4: Iterate over each row
        logger.info(f"Processing {total} log sources...")
        
        for idx, row in df.iterrows():
            log_source_name = str(row[LOGSOURCE_COLUMN]).strip()
            
            if not log_source_name or log_source_name.lower() in ['nan', 'none', '']:
                details = {
                    'status': 'No log source provided',
                    'qradar_id': '',
                    'protocol_type': '',
                    'protocol_name': '',
                    'enabled': '',
                    'description': '',
                    'error_details': 'Empty or null log source name'
                }
            else:
                details = get_log_source_details(
                    qradar_host=QRADAR_HOST,
                    api_token=QRADAR_API_TOKEN,
                    log_source_name=log_source_name,
                    ssl_cert=SSL_CERT_PATH,
                    verify_ssl=VERIFY_SSL
                )

            # Write back into DataFrame
            for key, val in details.items():
                df.at[idx, key] = val

            # Update counters
            if details['status'] == 'Exists':
                successful += 1
            elif details['status'] not in ['No log source provided']:
                failed += 1

            logger.info(f"[{idx+1}/{total}] '{log_source_name}': {details['status']}")
            
            # Rate limiting
            if idx < total - 1:  # Don't sleep after the last request
                time.sleep(API_DELAY)

        # Step 5: Save the updated DataFrame
        logger.info("Saving updated Excel file...")
        
        # Use ExcelWriter to preserve formatting better
        with pd.ExcelWriter(INPUT_EXCEL_PATH, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=EXCEL_SHEET_NAME if isinstance(EXCEL_SHEET_NAME, str) else 'Sheet1')
        
        # Summary
        logger.info(f"\n✅ Process completed successfully!")
        logger.info(f"📊 Summary:")
        logger.info(f"   - Total processed: {total}")
        logger.info(f"   - Found in QRadar: {successful}")
        logger.info(f"   - Not found: {total - successful - failed}")
        logger.info(f"   - Errors: {failed}")
        logger.info(f"   - Results saved to: {INPUT_EXCEL_PATH}")
        logger.info(f"   - Backup saved as: {backup_path}")

    except Exception as e:
        logger.error(f"Fatal error: {e}")
        raise


if __name__ == '__main__':
    main()
