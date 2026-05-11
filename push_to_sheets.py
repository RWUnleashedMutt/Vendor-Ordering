import pandas as pd
import os
import sys
import re
import tkinter as tk
from tkinter import filedialog
import gspread
from google.oauth2.service_account import Credentials

# --- CONSTANTS ---
CREDENTIALS_FILE = './credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

SHEET_IDS = {
    'All Vendors': '1iX-LpiavNqcyZqe1r068DmziafQbsDuugmdszqS89Tw',
    # Add a line for each vendor
}


def get_file_path(title="Select File", file_types=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(title=title, filetypes=file_types)
    root.destroy()
    return file_path


def sanitize_vendor_name(name):
    return re.sub(r'[<>:"/\\|?*]', '', name).strip()


def get_google_client():
    """Authenticate and return a gspread client."""
    creds = Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


def load_local_matrix(path):
    """Load the local rules matrix."""
    if not os.path.exists(path):
        print(f"Error: File not found at {path}")
        sys.exit(1)
    df = pd.read_excel(path, dtype={'SKU': str})
    print(f"Loaded local matrix: {len(df)} SKUs")
    return df


def pull_current_sheet(client, vendor):
    """Pull the current Google Sheet data for comparison."""
    if vendor not in SHEET_IDS:
        print(f"Error: No Sheet ID found for vendor '{vendor}'.")
        print("Please add the Sheet ID to the SHEET_IDS dictionary.")
        sys.exit(1)
    try:
        spreadsheet = client.open_by_key(SHEET_IDS[vendor])
        worksheet = spreadsheet.sheet1
        data = worksheet.get_all_records(
            value_render_option='UNFORMATTED_VALUE')
        df = pd.DataFrame(data)
        df['SKU'] = df['SKU'].astype(str)
        print(f"Loaded Google Sheet: {len(df)} SKUs")
        return spreadsheet, worksheet, df
    except gspread.SpreadsheetNotFound:
        print(f"Error: Could not find Google Sheet for vendor '{vendor}'.")
        sys.exit(1)


def compare_local_to_sheets(local_df, sheets_df):
    """Compare local matrix to current Sheets version and summarize differences."""
    local_skus = set(local_df['SKU'].tolist())
    sheets_skus = set(sheets_df['SKU'].tolist())

    new_in_local = local_skus - sheets_skus
    removed_from_local = sheets_skus - local_skus
    common_skus = local_skus & sheets_skus

    value_cols = [c for c in local_df.columns if c.endswith('_Min')
                  or c.endswith('_Max') or c.endswith('_DNO')]

    local_common = local_df[local_df['SKU'].isin(common_skus)].set_index('SKU')
    sheets_common = sheets_df[sheets_df['SKU'].isin(
        common_skus)].set_index('SKU')

    change_count = 0
    for sku in common_skus:
        for col in value_cols:
            if col in local_common.columns and col in sheets_common.columns:
                if str(local_common.loc[sku, col]) != str(sheets_common.loc[sku, col]):
                    change_count += 1

    return new_in_local, removed_from_local, change_count


def push_to_sheets(worksheet, local_df):
    """Push the local matrix to Google Sheets."""
    local_df = local_df.fillna('')
    data = [local_df.columns.tolist()] + local_df.values.tolist()
    worksheet.clear()
    worksheet.update(data)
    print(f"Pushed {len(local_df)} SKUs to Google Sheets.")


def run_push():
    """Main logic to push local matrix to Google Sheets."""
    try:
        # --- Select local file ---
        print("\nPlease select the local Rules Matrix to push...")
        matrix_path = get_file_path(title="Select Local Rules Matrix")
        if not matrix_path:
            print("No file selected. Exiting.")
            sys.exit(0)

        # --- Get vendor name ---
        Vendor = sanitize_vendor_name(input('\nInput Vendor name: '))
        if not Vendor:
            print("Error: Vendor name is invalid or empty.")
            sys.exit(1)

        # --- Load local matrix ---
        local_df = load_local_matrix(matrix_path)

        # --- Connect to Sheets and compare ---
        print("\nConnecting to Google Sheets...")
        client = get_google_client()
        spreadsheet, worksheet, sheets_df = pull_current_sheet(client, Vendor)

        new_in_local, removed_from_local, change_count = compare_local_to_sheets(
            local_df, sheets_df)

        # --- Show summary ---
        print(f"\n--- Push Summary ---")
        print(f"  SKUs to be added to Sheets:       {len(new_in_local)}")
        print(f"  SKUs to be removed from Sheets:   {len(removed_from_local)}")
        print(f"  Value changes (Min/Max/DNO):       {change_count}")
        print(f"\nSheet URL: {spreadsheet.url}")

        if len(new_in_local) == 0 and len(removed_from_local) == 0 and change_count == 0:
            print("\nNo differences found. Google Sheet is already up to date.")
            return

        # --- Confirm ---
        confirm = input(
            "\nPush these changes to Google Sheets? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Push cancelled. No changes made.")
            return

        # --- Push ---
        push_to_sheets(worksheet, local_df)
        print("\nGoogle Sheet successfully updated!")
        print(f"View it here: {spreadsheet.url}")

    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    run_push()
