# Building batch_apply.py: A Step-by-Step Tutorial

## Overview

`batch_apply.py` is a script that automates the process of syncing Google Sheets changes to multiple local vendor rule matrices. Instead of running `apply_changes.py` for each vendor one at a time, this script processes all vendors in batch mode with a single Google Sheets authentication.

**Key Features:**

- Batch processes all vendors from a config file
- Dry-run mode to preview changes without applying
- Confirmation prompt before making changes
- Error logging to file
- Summary reporting across all vendors

---

## Step 1: Set Up Imports and Configuration

Start by importing all the libraries we'll need.

```python
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
```

**What each import does:**

- `tomllib`: Read TOML configuration files
- `os`, `sys`: File and system operations
- `argparse`: Parse command-line arguments (like `--dry-run`)
- `datetime`: Create timestamps for logs
- `pandas`: DataFrames for data manipulation
- `gspread`: Google Sheets API client
- `google.oauth2`: Google authentication
- `apply_changes`: Reuse existing functions

**Why not import `pull_from_sheets`?** We'll create our own version that works with sheet IDs instead of vendor names.

---

## Step 2: Define Configuration Constants

```python
CREDENTIALS_FILE = './credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

LOG_DIR = './Data/log'
os.makedirs(LOG_DIR, exist_ok=True)
ERROR_LOG_FILE = os.path.join(LOG_DIR, f"batch_apply_errors_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
```

**What this does:**

- Points to where Google credentials are stored
- Defines which Google APIs we need access to
- Creates a log directory if it doesn't exist
- Generates a unique error log filename with timestamp

---

## Step 3: Create the Pull from Sheets Function

This is customized to use sheet IDs instead of vendor names.

```python
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
```

**Why use `sheet_id` instead of vendor name?**

- More reliable: IDs don't change, sheet names might
- Faster: Direct access vs. searching by name
- Cleaner separation: Config file stores IDs, not duplicated logic

---

## Step 4: Error Logging Function

A simple utility to log errors to both console and a file.

```python
def log_error(vendor, error_msg):
    """Write error to both console and error log file."""
    msg = f"[{vendor}] {error_msg}"
    print(f"  ✗ {msg}")
    with open(ERROR_LOG_FILE, 'a') as f:
        f.write(f"{datetime.now().isoformat()} - {msg}\n")
```

**Why log to both?**

- Console: You see errors immediately
- File: You can review errors later, find patterns

---

## Step 5: Load Vendors from Config

```python
def load_vendors_from_config():
    """Load vendors from config.toml"""
    try:
        with open('config.toml', 'rb') as f:
            config = tomllib.load(f)
        vendors = config.get('vendors', {})
        if not vendors:
            print("Error: No vendors found in config.toml")
            sys.exit(1)
        return vendors
    except FileNotFoundError:
        print("Error: config.toml not found")
        sys.exit(1)
```

**What this expects in `config.toml`:**

```toml
[vendors]
"Canine Caviar" = "1TJXe9V_aF1A1wm_O9XK_iWJNU119iH3ZBBorUlX_0ss"
"Fluff & Tuff" = "1nGWM9Lt34e3vpqaETjPeMVsCTKVC9kIEQ3VVx1mEUqY"
# ... more vendors
```

Returns a dictionary: `{"Canine Caviar": "sheet_id", "Fluff & Tuff": "sheet_id", ...}`

---

## Step 6: Collect All Changes (The Core Logic)

This function scans all vendors WITHOUT applying changes yet. The key insight: **collect first, confirm second, apply third**.

```python
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
            local_df = load_local_matrix(matrix_path)

            updated_df, changes_df, added, removed = compare_and_apply(local_df, sheets_df)

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
                print(f"Found {len(changes_df)} value changes, {len(added)} new, {len(removed)} removed")

            summary['processed'] += 1

        except Exception as e:
            results[vendor] = {'status': 'error', 'error': str(e)}
            log_error(vendor, str(e))
            summary['errors'] += 1
            summary['processed'] += 1

    return results, summary
```

**Why collect first?**

- If an error happens partway through, you know which vendors failed
- User can review all changes before committing any
- You can generate a comprehensive summary before asking confirmation

**What gets stored in `results`:**

- `updated_df`: The new local matrix data
- `changes_df`: All the changes that happened
- `sheets_df`, `local_df`: Both versions for comparison
- `num_changes`, `num_added`, `num_removed`: Counts for summary

---

## Step 7: Print Summary Report

```python
def print_summary(results, summary, dry_run=False):
    """Print a detailed summary of what was/would be changed."""
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print(f"\nVendors scanned:         {summary['processed']}/{summary['total_vendors']}")
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
```

**Output example:**

```
============================================================
SUMMARY
============================================================

Vendors scanned:         13/13
Vendors with changes:    3
Errors encountered:      0

TOTAL CHANGES ACROSS ALL VENDORS:
  New SKUs:              12
  Removed SKUs:          2
  Value Changes:         45

CHANGES BY VENDOR:
  Canine Caviar:
    - 20 value changes
    - 5 new SKUs
    - 1 removed SKU
  Fluff & Tuff:
    - 15 value changes
    - 4 new SKUs
    - 1 removed SKU
```

---

## Step 8: Apply All Changes

Only called after user confirms. This actually writes files to disk.

```python
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
```

**Why separate from collection?**

- File I/O is slow; don't do it until you're sure
- If a write fails midway, you know which vendors didn't get updated
- User has time to cancel before any files are touched

---

## Step 9: The Main Function (Orchestration)

This ties everything together. It's the "conductor" of the script.

```python
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
        results, summary = collect_all_changes(vendors, client, dry_run=args.dry_run)

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
```

**Flow:**

1. Parse command-line arguments
2. Load vendors from config
3. Authenticate with Google (once)
4. Collect all changes
5. Show summary
6. Ask for confirmation (unless `--dry-run`)
7. Apply changes if confirmed
8. Handle errors gracefully

---

## Step 10: The Entry Point

```python
if __name__ == "__main__":
    main()
```

This ensures `main()` only runs when the script is executed directly, not when imported.

---

## Full Script

Here's the complete `batch_apply.py`:

```python
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

LOG_DIR = './Data/log'
os.makedirs(LOG_DIR, exist_ok=True)
ERROR_LOG_FILE = os.path.join(LOG_DIR, f"batch_apply_errors_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")


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


def log_error(vendor, error_msg):
    """Write error to both console and error log file."""
    msg = f"[{vendor}] {error_msg}"
    print(f"  ✗ {msg}")
    with open(ERROR_LOG_FILE, 'a') as f:
        f.write(f"{datetime.now().isoformat()} - {msg}\n")


def load_vendors_from_config():
    """Load vendors from config.toml"""
    try:
        with open('config.toml', 'rb') as f:
            config = tomllib.load(f)
        vendors = config.get('vendors', {})
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
            local_df = load_local_matrix(matrix_path)

            updated_df, changes_df, added, removed = compare_and_apply(local_df, sheets_df)

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
                print(f"Found {len(changes_df)} value changes, {len(added)} new, {len(removed)} removed")

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
    print(f"\nVendors scanned:         {summary['processed']}/{summary['total_vendors']}")
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
        results, summary = collect_all_changes(vendors, client, dry_run=args.dry_run)

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
```

---

## How to Use

### Setup

1. Create `config.toml` in the same folder as `batch_apply.py`
2. Add your vendor sheet IDs:
   ```toml
   [vendors]
   "Canine Caviar" = "1TJXe9V_aF1A1wm_O9XK_iWJNU119iH3ZBBorUlX_0ss"
   "Fluff & Tuff" = "1nGWM9Lt34e3vpqaETjPeMVsCTKVC9kIEQ3VVx1mEUqY"
   # ... add all vendors
   ```

### Running It

**Preview mode (no changes applied):**

```bash
python batch_apply.py --dry-run
```

**Scan and apply with confirmation:**

```bash
python batch_apply.py
```

### Example Output

```
Loaded 13 vendors from config.toml
Connecting to Google Sheets...
✓ Connected

============================================================
SCANNING FOR CHANGES
============================================================

Scanning: Canine Caviar... Found 20 value changes, 5 new, 1 removed
Scanning: Fluff & Tuff... Found 15 value changes, 4 new, 1 removed
Scanning: SE... No changes
Scanning: Front Porch Pets... Found 10 value changes, 3 new, 0 removed
...

============================================================
SUMMARY
============================================================

Vendors scanned:         13/13
Vendors with changes:    3
Errors encountered:      0

TOTAL CHANGES ACROSS ALL VENDORS:
  New SKUs:              12
  Removed SKUs:          2
  Value Changes:         45

============================================================
Apply 45 value changes, 12 new SKUs, and 2 removed SKUs? (yes/no): yes

============================================================
APPLYING CHANGES
============================================================

✓ Canine Caviar: Applied successfully
✓ Fluff & Tuff: Applied successfully
✓ Front Porch Pets: Applied successfully

✓ Applied changes to 3 vendors

============================================================
Batch processing complete!
============================================================
```

---

## Key Design Decisions

### 1. **Three-Phase Approach (Collect → Summarize → Apply)**

Instead of processing vendors one at a time, we:

1. **Collect** all changes from all vendors
2. **Summarize** what changed globally
3. **Ask for confirmation** with complete information
4. **Apply** once user agrees

**Benefits:** Safe, transparent, reversible until the last moment.

### 2. **Error Handling Without Stopping**

If one vendor fails to scan, we log it and continue with the rest.

```python
except Exception as e:
    log_error(vendor, str(e))
    summary['errors'] += 1
    summary['processed'] += 1
```

**Why?** You want to know about all problems at once, not stop at the first error.

### 3. **Dry-Run Mode**

Run with `--dry-run` to see what _would_ happen without actually doing it.

```python
if not args.dry_run and summary['vendors_with_changes'] > 0:
    # Only ask for confirmation in non-dry-run mode
```

**Why?** Safety and confidence before committing changes.

### 4. **Sheet IDs Instead of Names**

Uses Google Sheet IDs from `config.toml` instead of constructing sheet names.

```python
spreadsheet = client.open_by_key(sheet_id)  # Direct and reliable
```

**Why?**

- IDs never change; names might be edited by accident
- Faster lookup (direct vs. searching)
- Cleaner architecture

---

## Extending the Script

### Add Slack Notifications

After applying changes, send a summary to Slack:

```python
import requests

def notify_slack(summary, applied_count):
    webhook_url = "https://hooks.slack.com/services/YOUR/WEBHOOK/URL"
    message = {
        'text': f'Batch sync complete: {applied_count} vendors updated',
        'blocks': [
            {
                'type': 'section',
                'text': {
                    'type': 'mrkdwn',
                    'text': f'*Batch Sync Report*\n• Vendors updated: {applied_count}\n• Value changes: {summary["total_value_changes"]}\n• New SKUs: {summary["total_new_skus"]}'
                }
            }
        ]
    }
    requests.post(webhook_url, json=message)
```

### Add Email Reports

Export the summary to a CSV and email it:

```python
def email_summary(summary, results):
    import smtplib
    from email.mime.text import MIMEText
    # ... email code
```

### Schedule with Cron

Run automatically every morning:

```bash
0 9 * * * /usr/bin/python3 /path/to/batch_apply.py >> /var/log/batch_apply.log 2>&1
```

---

## Troubleshooting

| Problem                       | Solution                                                                         |
| ----------------------------- | -------------------------------------------------------------------------------- |
| `config.toml not found`       | Make sure `config.toml` is in the same folder as `batch_apply.py`                |
| `Could not find Google Sheet` | Verify sheet IDs in `config.toml` are correct and you have access                |
| `Matrix file not found`       | Check that local matrix files exist at `./Data/Rules/{vendor} Rules Matrix.xlsx` |
| `Credential error`            | Ensure `credentials.json` exists and is valid                                    |
| Script hangs                  | Press Ctrl+C to cancel. Any applied changes are already saved.                   |

---

## Summary

`batch_apply.py` is built on these principles:

1. **Modularity**: Each function does one thing well
2. **Safety**: Collect, review, confirm, apply (never silently change files)
3. **Observability**: Clear output, detailed logging, error tracking
4. **Flexibility**: Dry-run mode, command-line args, easy to extend

It transforms a tedious 13-vendor manual sync process into a single command with full visibility and control.
