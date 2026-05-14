import pandas as pd
import os
import sys
import re
import tkinter as tk
from tkinter import filedialog
import gspread
from google.oauth2.service_account import Credentials

# --- CONSTANTS ---
DEFAULT_MIN = 1
DEFAULT_MAX = 2
DEFAULT_ORDER_QTY = 1
REQUIRED_CATALOG_COLS = {'SKU', 'Item Name', 'Reporting Category'}
CREDENTIALS_FILE = './credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

EXCLUDED_SKUS = {
    # Add SKUs here to exclude them from the matrix entirely
    # Example:
    # '12345',
    # 'SKU-ABC',
}

SHEET_IDS = {
    'All Vendors': '1iX-LpiavNqcyZqe1r068DmziafQbsDuugmdszqS89Tw',
    'Bradley Caldwell': '1eqENDXTdDJVKdos-VUXYNYMNM806rNcDrv63Q654nyc',
    'Fluff & Tuff': '1nGWM9Lt34e3vpqaETjPeMVsCTKVC9kIEQ3VVx1mEUqY',
    'SE': '1O6HWGeLgtdScnJ0_pQc8asaSj3-L4pP9vjCvvXa26vQ'
    # Add a line for each vendor
}

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


def get_or_create_sheet(client, vendor):
    """Open an existing Google Sheet by ID."""
    if vendor not in SHEET_IDS:
        print(f"Error: No Sheet ID found for vendor '{vendor}'.")
        print(
            "Please create the Sheet manually in Google Drive and add the ID to SHEET_IDS.")
        sys.exit(1)

    spreadsheet = client.open_by_key(SHEET_IDS[vendor])
    print(f"Connected to Google Sheet for {vendor}")
    return spreadsheet


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


def load_catalog(path):
    """Load and validate the catalog file."""
    # NOTE: header=1 assumes the catalog has a blank/title row before the actual headers.
    # Change to header=0 if your headers are on row 1.
    catalog = pd.read_excel(path, header=1, usecols=list(REQUIRED_CATALOG_COLS),
                            dtype={'SKU': str, 'Item Name': str})
    catalog = catalog.dropna(subset=['SKU']).drop_duplicates(subset=['SKU'])

    missing_cols = REQUIRED_CATALOG_COLS - set(catalog.columns)
    if missing_cols:
        print(f"Error: Catalog is missing required columns: {missing_cols}")
        sys.exit(1)

    return catalog


def load_or_create_matrix(path, catalog):
    """Load existing rules matrix or create a fresh one from the catalog."""
    if os.path.exists(path):
        rules_df = pd.read_excel(path, dtype={'SKU': str})

        if 'Order_Qty' in rules_df.columns:
            rules_df = rules_df.rename(
                columns={'Order_Qty': 'Order In Quantities'})

        rules_df = rules_df.drop(
            columns=['Item Name', 'Reporting Category'], errors='ignore')
        rules_df = pd.merge(
            rules_df, catalog[['SKU', 'Item Name', 'Reporting Category']], on='SKU', how='left')

        if 'Order In Quantities' not in rules_df.columns:
            rules_df['Order In Quantities'] = DEFAULT_ORDER_QTY
    else:
        rules_df = catalog.copy()
        rules_df['Order In Quantities'] = DEFAULT_ORDER_QTY

    return rules_df


def remove_excluded_skus(rules_df):
    """Remove any SKUs that are on the exclusion list."""
    excluded = rules_df[rules_df['SKU'].isin(EXCLUDED_SKUS)]

    if not excluded.empty:
        print(f"\nRemoving {len(excluded)} excluded SKU(s):")
        for sku in excluded['SKU'].tolist():
            print(f"  - {sku}")
        rules_df = rules_df[~rules_df['SKU'].isin(
            EXCLUDED_SKUS)].reset_index(drop=True)
    else:
        print("No excluded SKUs found in matrix.")

    return rules_df


def remove_discontinued_skus(rules_df, catalog):
    """Remove SKUs from the matrix that are no longer in the catalog."""
    catalog_skus = set(catalog['SKU'].tolist())
    discontinued = rules_df[~rules_df['SKU'].isin(catalog_skus)]

    if not discontinued.empty:
        print(f"\nRemoving {len(discontinued)} discontinued SKU(s):")
        for sku in discontinued['SKU'].tolist():
            print(f"  - {sku}")

        confirm = input(
            "\nConfirm removal of discontinued SKUs? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Removal cancelled. Discontinued SKUs will be kept.")
            return rules_df

        rules_df = rules_df[rules_df['SKU'].isin(
            catalog_skus)].reset_index(drop=True)
        print("Discontinued SKUs removed.")
    else:
        print("No discontinued SKUs found.")

    return rules_df


def ensure_store_columns(rules_df):
    """Add any missing store columns with default values."""
    for code in store_map.values():
        if f'{code}_DNO' not in rules_df.columns:
            rules_df[f'{code}_DNO'] = False
        if f'{code}_Min' not in rules_df.columns:
            rules_df[f'{code}_Min'] = DEFAULT_MIN
        if f'{code}_Max' not in rules_df.columns:
            rules_df[f'{code}_Max'] = DEFAULT_MAX
    return rules_df


def append_new_skus(rules_df, catalog):
    """Add any SKUs from the catalog that aren't already in the matrix."""
    existing_skus = set(rules_df['SKU'].tolist())
    new_items = catalog[~catalog['SKU'].isin(existing_skus)].copy()

    if not new_items.empty:
        new_items['Order In Quantities'] = DEFAULT_ORDER_QTY
        for code in store_map.values():
            new_items[f'{code}_DNO'] = False
            new_items[f'{code}_Min'] = DEFAULT_MIN
            new_items[f'{code}_Max'] = DEFAULT_MAX
        rules_df = pd.concat([rules_df, new_items], ignore_index=True)
        print(f"Added {len(new_items)} new SKU(s) to the matrix.")
    else:
        print("No new SKUs found.")

    return rules_df


def apply_dno_zeroing(rules_df):
    """For any store where DNO is True, set the corresponding Min and Max to 0."""
    zeroed_count = 0

    for code in store_map.values():
        dno_col = f'{code}_DNO'
        min_col = f'{code}_Min'
        max_col = f'{code}_Max'

        if dno_col not in rules_df.columns:
            continue

        # Match both boolean True and string 'TRUE'
        dno_mask = rules_df[dno_col].apply(
            lambda v: str(v).upper() == 'TRUE' or v is True
        )

        if dno_mask.any():
            if min_col in rules_df.columns:
                zeroed_count += (rules_df.loc[dno_mask, min_col] != 0).sum()
                rules_df.loc[dno_mask, min_col] = 0
            if max_col in rules_df.columns:
                zeroed_count += (rules_df.loc[dno_mask, max_col] != 0).sum()
                rules_df.loc[dno_mask, max_col] = 0

    if zeroed_count:
        print(
            f"DNO zeroing applied: {zeroed_count} Min/Max value(s) set to 0.")
    else:
        print("DNO zeroing: no Min/Max values needed adjustment.")

    return rules_df


def sync_rules_matrix(vendor, catalog_path, matrix_path):
    """Main logic to sync the rules matrix and push to Google Sheets."""
    try:
        catalog = load_catalog(catalog_path)
        # Filter excluded SKUs from catalog
        catalog = catalog[~catalog['SKU'].isin(EXCLUDED_SKUS)]
        rules_df = load_or_create_matrix(matrix_path, catalog)
        # Remove excluded SKUs from matrix
        rules_df = remove_excluded_skus(rules_df)
        rules_df = remove_discontinued_skus(rules_df, catalog)
        rules_df = ensure_store_columns(rules_df)
        rules_df = append_new_skus(rules_df, catalog)
        rules_df = apply_dno_zeroing(rules_df)

        matrix_columns = ['SKU', 'Item Name',
                          'Reporting Category', 'Order In Quantities']
        for code in store_map.values():
            matrix_columns.extend(
                [f'{code}_DNO', f'{code}_Min', f'{code}_Max'])

        final_cols = [c for c in matrix_columns if c in rules_df.columns]
        rules_df = rules_df[final_cols]

        # Save locally
        os.makedirs(os.path.dirname(matrix_path), exist_ok=True)
        rules_df.to_excel(matrix_path, index=False)
        print(f"\nLocal matrix saved at {matrix_path}")

        # Push to Google Sheets
        print("\nConnecting to Google Sheets...")
        client = get_google_client()
        spreadsheet = get_or_create_sheet(client, vendor)
        push_to_sheets(spreadsheet, rules_df)

        print(f"\nSuccess! Matrix synced locally and to Google Sheets.")

    except FileNotFoundError as e:
        print(f"Error: File not found — {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    Vendor = sanitize_vendor_name(input('Input Vendor: (Ex: Southeast): '))
    if not Vendor:
        print("Error: Vendor name is invalid or empty.")
        sys.exit(1)

    # --- WARNING ---
    print("\n" + "="*60)
    print("  WARNING: Before uploading a new catalog, make sure you")
    print("  have pulled the latest edits from Google Sheets first.")
    print("  If you skip this, any unsaved team edits will be lost.")
    print("="*60)
    print("\nHave you already run apply_changes.py to pull team edits? (yes/no): ", end="")
    confirm = input().strip().lower()

    if confirm != 'yes':
        print("\nPlease run apply_changes.py first, then come back and run this script.")
        print("Exiting.")
        sys.exit(0)

    # --- PROCEED ---
    print("\nPlease select the Catalog Excel file...")
    CATALOG_PATH = get_file_path(title=f"Select Catalog for {Vendor}")
    if not CATALOG_PATH:
        print("No file selected. Exiting.")
        sys.exit(0)

    RULES_MATRIX_PATH = f'./Data/Rules/{Vendor} Rules Matrix.xlsx'
    sync_rules_matrix(Vendor, CATALOG_PATH, RULES_MATRIX_PATH)
