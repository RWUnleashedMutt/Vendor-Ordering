import time
import tomllib
import os
import sys
import argparse
from datetime import datetime
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- CONSTANTS ---
CREDENTIALS_FILE = './credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

store_map = {
    'Current Quantity City Market: DTR': 'CM',
    'Current Quantity Crabtree Valley Mall': 'CVM',
    'Current Quantity Crescent Commons': 'CC',
    'Current Quantity Downtown Durham': 'DTD',
    'Current Quantity Front Street': 'MF',
    'Current Quantity Lake Boone': 'LB',
    'Current Quantity Landfall Shopping Center': 'LF',
    'Current Quantity Parkway Plaza': 'PP',
    'Current Quantity Southport - Tidewater': 'SP',
    'Current Quantity Stonehenge Market': 'SH',
    'Current Quantity The Streets at Southpoint': 'SS',
    'Current Quantity HQ': 'HQ'
}

# Set up error logging
LOG_DIR = './Data/log'
os.makedirs(LOG_DIR, exist_ok=True)
ERROR_LOG_FILE = os.path.join(
    LOG_DIR, f"batch_total_max_errors_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")


def log_error(vendor, error_msg):
    """Write error to both console and error log file."""
    msg = f"[{vendor}] {error_msg}"
    print(f"  ✗ {msg}")
    with open(ERROR_LOG_FILE, 'a') as f:
        f.write(f"{datetime.now().isoformat()} - {msg}\n")


def get_google_client():
    """Authenticate and return a gspread client."""
    creds = Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


def load_vendors_from_config():
    """Load vendors from config.toml"""
    try:
        with open('./.streamlit/secrets.toml', 'rb') as f:
            config = tomllib.load(f)
        vendors = config.get('sheet_ids', {})
        if not vendors:
            print("Error: No vendors found in config.toml")
            sys.exit(1)
        return vendors
    except FileNotFoundError:
        print("Error: config.toml not found")
        sys.exit(1)


def pull_from_sheets(client, sheet_id):
    """Pull the current edited matrix from Google Sheets by sheet ID."""
    try:
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.sheet1
        data = worksheet.get_all_records(
            value_render_option='UNFORMATTED_VALUE')
        df = pd.DataFrame(data)
        df['SKU'] = df['SKU'].astype(str)
        print(f"Pulled {len(df)} SKUs from Google Sheets.")
        return df
    except gspread.SpreadsheetNotFound:
        print(f"Error: Could not find Google Sheet with ID '{sheet_id}'.")
        sys.exit(1)


def push_to_sheets(spreadsheet, rules_df):
    """Push the rules matrix DataFrame to Google Sheets."""
    worksheet = spreadsheet.sheet1
    worksheet.clear()

    # Replace NaN with empty string for Sheets compatibility
    rules_df = rules_df.fillna('')
    data = [rules_df.columns.tolist()] + rules_df.values.tolist()
    worksheet.update(data)
    print(f"Pushed {len(rules_df)} SKUs to Google Sheets.")
    print(f"Sheet URL: {spreadsheet.url}")


def calculate_total_max(rules_df):
    """Sum every store's *_Max column into a single 'Total Max' column.

    Run this AFTER any DNO zeroing sync logic so the total reflects
    the final, synced Max values (DNO'd stores contribute 0).
    """
    max_cols = [f'{code}_Max' for code in store_map.values()
                if f'{code}_Max' in rules_df.columns]

    if not max_cols:
        print("Total Max: no store Max columns found, skipping.")
        rules_df['Total Max'] = 0
        return rules_df

    rules_df['Total Max'] = rules_df[max_cols].apply(
        pd.to_numeric, errors='coerce').fillna(0).sum(axis=1).astype(int)

    print(f"Total Max calculated across {len(max_cols)} store column(s).")
    return rules_df


def load_local_matrix(matrix_path):
    """Load the local matrix file if it exists, else return None."""
    if os.path.exists(matrix_path):
        df = pd.read_excel(matrix_path, dtype={'SKU': str})
        return df
    return None


def process_vendor(vendor, sheet_id, client, dry_run=False):
    """Pull a vendor's sheet, recalculate Total Max, and push/save the result."""
    matrix_path = f'./Data/Rules/{vendor} Rules Matrix.xlsx'

    sheets_df = pull_from_sheets(client, sheet_id)
    sheets_df = calculate_total_max(sheets_df)

    max_cols = [f'{code}_Max' for code in store_map.values()
                if f'{code}_Max' in sheets_df.columns]

    if dry_run:
        return {
            'status': 'ok',
            'num_skus': len(sheets_df),
            'num_max_cols_found': len(max_cols),
            'total_sum': int(sheets_df['Total Max'].sum())
        }

    # Open spreadsheet once and push the updated data (with Total Max) back
    spreadsheet = client.open_by_key(sheet_id)
    push_to_sheets(spreadsheet, sheets_df)

    # Keep the local matrix file consistent too, if one exists
    if os.path.exists(matrix_path):
        sheets_df.to_excel(matrix_path, index=False)

    return {
        'status': 'ok',
        'num_skus': len(sheets_df),
        'num_max_cols_found': len(max_cols),
        'total_sum': int(sheets_df['Total Max'].sum())
    }


def run_batch(dry_run=False):
    vendors = load_vendors_from_config()
    print(f"\nLoaded {len(vendors)} vendors from config.toml")

    print("Connecting to Google Sheets...")
    client = get_google_client()
    print("✓ Connected\n")

    print("="*60)
    print("RECALCULATING TOTAL MAX" + (" (DRY RUN)" if dry_run else ""))
    print("="*60 + "\n")

    results = {}
    processed = 0
    errors = 0

    vendor_items = list(vendors.items())
    for i, (vendor, sheet_id) in enumerate(vendor_items):
        print(f"Processing: {vendor}...", end=" ")
        try:
            result = process_vendor(vendor, sheet_id, client, dry_run=dry_run)
            results[vendor] = result
            print(
                f"OK — {result['num_skus']} SKUs, "
                f"{result['num_max_cols_found']} store Max cols, "
                f"sum={result['total_sum']}"
            )
            processed += 1
        except gspread.SpreadsheetNotFound:
            log_error(vendor, f"Sheet ID '{sheet_id}' not found")
            results[vendor] = {'status': 'error', 'error': 'Sheet not found'}
            errors += 1
        except Exception as e:
            log_error(vendor, str(e))
            results[vendor] = {'status': 'error', 'error': str(e)}
            errors += 1

        # Avoid hammering the Sheets API, same pacing as batch_apply.py
        if i < len(vendor_items) - 1:
            time.sleep(7)

    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print(f"\nVendors processed: {processed}/{len(vendors)}")
    print(f"Errors:            {errors}")

    if errors:
        print(f"\nSee {ERROR_LOG_FILE} for details")

    if dry_run:
        print("\n[DRY RUN MODE] - No changes were pushed to Sheets or saved locally")

    return results


def main():
    parser = argparse.ArgumentParser(
        description='Recalculate Total Max across all vendor Rules Matrix sheets'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Calculate and report totals without pushing to Sheets or saving locally'
    )
    args = parser.parse_args()

    try:
        run_batch(dry_run=args.dry_run)
        print(f"\n{'='*60}")
        print("Batch Total Max calculation complete!")
        print('='*60)
    except KeyboardInterrupt:
        print("\n\nBatch processing cancelled by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\nFatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
