import time
import tomllib
import os
import sys
import argparse
from datetime import datetime
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from apply_changes import (
    get_google_client,
    load_local_matrix,
    compare_and_apply,
    export_change_log
)

CREDENTIALS_FILE = './credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]


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


# Set up error logging
LOG_DIR = './Data/log'
os.makedirs(LOG_DIR, exist_ok=True)
ERROR_LOG_FILE = os.path.join(
    LOG_DIR, f"batch_apply_errors_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")


def log_error(vendor, error_msg):
    """Write error to both console and error log file."""
    msg = f"[{vendor}] {error_msg}"
    print(f"  ✗ {msg}")
    with open(ERROR_LOG_FILE, 'a') as f:
        f.write(f"{datetime.now().isoformat()} - {msg}\n")


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


def collect_all_changes(vendors, client, dry_run=False):
    """
    Collect changes for all vendors without applying them yet.
    Returns a dict of results per vendor and overall totals.
    vendors is a dict of {vendor_name: sheet_id}
    """
    results = {}
    summary = {
        'total_vendors': len(vendors),
        'processed': 0,
        'errors': 0,
        'total_new_skus': 0,
        'total_removed_skus': 0,
        'total_value_changes': 0,
        'vendors_with_changes': 0
    }

    print("\n" + "="*60)
    print("SCANNING FOR CHANGES" + (" (DRY RUN)" if dry_run else ""))
    print("="*60 + "\n")

    for vendor, sheet_id in vendors.items():
        print(f"Scanning: {vendor}...", end=" ")
        try:
            matrix_path = f'./Data/Rules/{vendor} Rules Matrix.xlsx'

            sheets_df = pull_from_sheets(client, sheet_id)
            time.sleep(5)
            local_df = load_local_matrix(matrix_path)

            updated_df, changes_df, added, removed = compare_and_apply(
                local_df, sheets_df)

            results[vendor] = {
                'status': 'ok',
                'matrix_path': matrix_path,
                'updated_df': updated_df,
                'changes_df': changes_df,
                'sheets_df': sheets_df,
                'local_df': local_df,
                'added': added,
                'removed': removed,
                'num_changes': len(changes_df),
                'num_added': len(added),
                'num_removed': len(removed)
            }

            if len(changes_df) == 0 and not added and not removed:
                print("No changes")
            else:
                summary['vendors_with_changes'] += 1
                summary['total_new_skus'] += len(added)
                summary['total_removed_skus'] += len(removed)
                summary['total_value_changes'] += len(changes_df)
                print(
                    f"Found {len(changes_df)} value changes, {len(added)} new, {len(removed)} removed")

            summary['processed'] += 1

        except Exception as e:
            results[vendor] = {'status': 'error', 'error': str(e)}
            log_error(vendor, str(e))
            summary['errors'] += 1
            summary['processed'] += 1

    return results, summary


def print_summary(results, summary, dry_run=False):
    """Print a detailed summary of what was/would be changed."""
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print(
        f"\nVendors scanned:         {summary['processed']}/{summary['total_vendors']}")
    print(f"Vendors with changes:    {summary['vendors_with_changes']}")
    print(f"Errors encountered:      {summary['errors']}")

    if summary['vendors_with_changes'] > 0:
        print(f"\nTOTAL CHANGES ACROSS ALL VENDORS:")
        print(f"  New SKUs:              {summary['total_new_skus']}")
        print(f"  Removed SKUs:          {summary['total_removed_skus']}")
        print(f"  Value Changes:         {summary['total_value_changes']}")

        print(f"\nCHANGES BY VENDOR:")
        for vendor, result in results.items():
            if result['status'] == 'ok':
                num_changes = result['num_changes']
                num_added = result['num_added']
                num_removed = result['num_removed']
                if num_changes > 0 or num_added > 0 or num_removed > 0:
                    print(f"  {vendor}:")
                    print(f"    - {num_changes} value changes")
                    print(f"    - {num_added} new SKUs")
                    print(f"    - {num_removed} removed SKUs")

    if summary['errors'] > 0:
        print(f"\nERRORS (see {ERROR_LOG_FILE} for details)")

    if dry_run:
        print("\n[DRY RUN MODE] - No changes were applied")


def apply_all_changes(results):
    """Apply all collected changes to local files and export logs."""
    print("\n" + "="*60)
    print("APPLYING CHANGES")
    print("="*60 + "\n")

    applied_count = 0
    for vendor, result in results.items():
        if result['status'] != 'ok':
            continue

        if result['num_changes'] == 0 and result['num_added'] == 0 and result['num_removed'] == 0:
            print(f"{vendor}: No changes to apply")
            continue

        try:
            # Save updated local matrix
            result['updated_df'].to_excel(result['matrix_path'], index=False)

            # Export change log
            export_change_log(
                result['changes_df'],
                result['added'],
                result['removed'],
                result['sheets_df'],
                result['local_df'],
                vendor
            )

            print(f"✓ {vendor}: Applied successfully")
            applied_count += 1

        except Exception as e:
            log_error(vendor, f"Failed to apply changes: {e}")

    return applied_count


def main():
    parser = argparse.ArgumentParser(
        description='Batch apply changes from Google Sheets to all vendors'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Scan for changes without applying them'
    )
    args = parser.parse_args()

    try:
        # Load config
        vendors = load_vendors_from_config()
        print(f"\nLoaded {len(vendors)} vendors from config.toml")

        # Single auth for all vendors
        print("Connecting to Google Sheets...")
        client = get_google_client()
        print("✓ Connected")

        # Collect changes
        results, summary = collect_all_changes(
            vendors, client, dry_run=args.dry_run)

        # Show summary
        print_summary(results, summary, dry_run=args.dry_run)

        # Ask for confirmation (unless dry-run)
        if not args.dry_run and summary['vendors_with_changes'] > 0:
            confirm = input(
                f"\n{'='*60}\nApply {summary['total_value_changes']} value changes, "
                f"{summary['total_new_skus']} new SKUs, and {summary['total_removed_skus']} removed SKUs? (yes/no): "
            ).strip().lower()

            if confirm != 'yes':
                print("Changes not applied.")
                return

            # Apply all changes
            applied_count = apply_all_changes(results)
            print(f"\n✓ Applied changes to {applied_count} vendors")

        elif args.dry_run:
            print("\nTo apply these changes, run without --dry-run flag")

        print(f"\n{'='*60}")
        print("Batch processing complete!")
        print('='*60)

    except KeyboardInterrupt:
        print("\n\nBatch processing cancelled by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\nFatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
