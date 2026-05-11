import pandas as pd
import os
import sys
import re
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

CREDENTIALS_FILE = './credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]


def sanitize_vendor_name(name):
    return re.sub(r'[<>:"/\\|?*]', '', name).strip()


def get_google_client():
    creds = Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


def pull_from_sheets(client, vendor):
    """Pull the current edited matrix from Google Sheets."""
    sheet_name = f'{vendor} Rules Matrix'
    try:
        spreadsheet = client.open(sheet_name)
        worksheet = spreadsheet.sheet1
        data = worksheet.get_all_records(
            value_render_option='UNFORMATTED_VALUE')
        df = pd.DataFrame(data)
        df['SKU'] = df['SKU'].astype(str)
        print(f"Pulled {len(df)} SKUs from Google Sheets.")
        return df
    except gspread.SpreadsheetNotFound:
        print(f"Error: Could not find Google Sheet '{sheet_name}'.")
        sys.exit(1)


def load_local_matrix(path):
    """Load the original local rules matrix."""
    if not os.path.exists(path):
        print(f"Error: Local matrix not found at {path}")
        sys.exit(1)
    df = pd.read_excel(path, dtype={'SKU': str})
    print(f"Loaded local matrix: {len(df)} SKUs")
    return df


def compare_and_apply(local_df, sheets_df):
    """Compare Sheets version to local, log changes, and return updated DataFrame."""
    local_skus = set(local_df['SKU'].tolist())
    sheets_skus = set(sheets_df['SKU'].tolist())

    added = sheets_skus - local_skus
    removed = local_skus - sheets_skus
    change_records = []

    # Find value changes on common SKUs
    common_skus = local_skus & sheets_skus
    local_common = local_df[local_df['SKU'].isin(common_skus)].set_index('SKU')
    sheets_common = sheets_df[sheets_df['SKU'].isin(
        common_skus)].set_index('SKU')

    value_cols = [c for c in local_df.columns if c.endswith('_Min')
                  or c.endswith('_Max') or c.endswith('_DNO')]

    dno_cols = [c for c in value_cols if c.endswith('_DNO')]

    # Work on a copy so we can mutate it
    updated_df = sheets_df.copy().set_index('SKU')

    for sku in common_skus:
        item_name = sheets_common.loc[sku,
                                      'Item Name'] if 'Item Name' in sheets_common.columns else ''

        # --- DNO → True: zero out Min/Max first ---
        for dno_col in dno_cols:
            if dno_col not in local_common.columns or dno_col not in sheets_common.columns:
                continue

            old_dno = local_common.loc[sku, dno_col]
            new_dno = sheets_common.loc[sku, dno_col]
            dno_became_true = (str(old_dno).upper() != 'TRUE') and (
                str(new_dno).upper() == 'TRUE' or new_dno is True)

            if dno_became_true:
                store = dno_col.rsplit('_', 1)[0]
                min_col = f'{store}_Min'
                max_col = f'{store}_Max'

                for zero_col in (min_col, max_col):
                    if zero_col in updated_df.columns:
                        old_zero_val = updated_df.loc[sku, zero_col]
                        updated_df.loc[sku, zero_col] = 0
                        if str(old_zero_val) != '0':
                            change_records.append({
                                'SKU': sku,
                                'Item Name': item_name,
                                'Store': store,
                                'Field': zero_col.rsplit('_', 1)[1],
                                'Old Value': old_zero_val,
                                'New Value': 0
                            })

        # --- Normal value-change detection ---
        for col in value_cols:
            if col not in local_common.columns or col not in sheets_common.columns:
                continue

            old_val = local_common.loc[sku, col]
            new_val = updated_df.loc[sku, col]
            if str(old_val) != str(new_val):
                store = col.rsplit('_', 1)[0]
                field = col.rsplit('_', 1)[1]
                already_logged = any(
                    r['SKU'] == sku and r['Store'] == store and r['Field'] == field
                    for r in change_records
                )
                if not already_logged:
                    change_records.append({
                        'SKU': sku,
                        'Item Name': item_name,
                        'Store': store,
                        'Field': field,
                        'Old Value': old_val,
                        'New Value': new_val
                    })

    updated_df = updated_df.reset_index()
    changes_df = pd.DataFrame(change_records)

    # Summary
    print(f"\n--- Changes Found ---")
    print(f"  New SKUs:        {len(added)}")
    print(f"  Removed SKUs:    {len(removed)}")
    print(f"  Value Changes:   {len(changes_df)}")

    return updated_df, changes_df, added, removed


def export_change_log(changes_df, added, removed, sheets_df, local_df, vendor):
    """Export a change log Excel report."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f'./Data/Rules/{vendor}_ChangeLog_{timestamp}.xlsx'
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    added_df = sheets_df[sheets_df['SKU'].isin(
        added)][['SKU', 'Item Name']].copy()
    added_df.insert(0, 'Change', 'Added')

    removed_df = local_df[local_df['SKU'].isin(
        removed)][['SKU', 'Item Name']].copy()
    removed_df.insert(0, 'Change', 'Removed')

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        summary_df = pd.DataFrame({
            'Category': ['New SKUs Added', 'SKUs Removed', 'Value Changes (Min/Max/DNO)'],
            'Count': [len(added), len(removed), len(changes_df)]
        })
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        if not added_df.empty:
            added_df.to_excel(writer, sheet_name='New SKUs', index=False)
        if not removed_df.empty:
            removed_df.to_excel(writer, sheet_name='Removed SKUs', index=False)
        if not changes_df.empty:
            changes_df.to_excel(
                writer, sheet_name='Value Changes', index=False)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_length = max((len(str(cell.value))
                                 for cell in col if cell.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(
                    max_length + 4, 50)

    print(f"\nChange log saved to: {output_path}")


def apply_changes(vendor):
    """Main logic to pull from Sheets, apply changes, and save locally."""
    try:
        matrix_path = f'./Data/Rules/{vendor} Rules Matrix.xlsx'

        print("\nConnecting to Google Sheets...")
        client = get_google_client()
        sheets_df = pull_from_sheets(client, vendor)
        local_df = load_local_matrix(matrix_path)

        updated_df, changes_df, added, removed = compare_and_apply(
            local_df, sheets_df)

        if len(changes_df) == 0 and not added and not removed:
            print("\nNo changes detected. Local file unchanged.")
            return

        confirm = input(
            "\nApply these changes to your local matrix? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Changes not applied.")
            return

        # Save updated local matrix
        updated_df.to_excel(matrix_path, index=False)
        print(f"\nLocal matrix updated at {matrix_path}")

        # Export change log
        export_change_log(changes_df, added, removed,
                          sheets_df, local_df, vendor)

        print("\nDone!")

    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    Vendor = sanitize_vendor_name(input('Input Vendor: (Ex: Southeast): '))
    if not Vendor:
        print("Error: Vendor name is invalid or empty.")
        sys.exit(1)

    apply_changes(Vendor)
