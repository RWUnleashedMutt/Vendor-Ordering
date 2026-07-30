import pandas as pd
import tomllib
import sys
import tkinter as tk
from tkinter import filedialog
import gspread
from google.oauth2.service_account import Credentials
from datetime import date

# --- CONSTANTS ---
DEFAULT_MIN = 0
DEFAULT_MAX = 0
DEFAULT_ORDER_QTY = 1
REQUIRED_CATALOG_COLS = {'Token', 'SKU', 'Item Name', 'Reporting Category'}
CREDENTIALS_FILE = './credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
with open('./.streamlit/secrets.toml', 'rb') as f:
    config = tomllib.load(f)

SHEET_IDS = config['sheet_ids']

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


def select_vendor_interactively():
    """Prompt the user to pick a vendor from config.toml's sheet_ids, by
    number or exact name — same UX as the vendor picker in batch_apply.py
    and batch_calculate_total_max.py."""
    vendor_names = list(SHEET_IDS.keys())

    if not vendor_names:
        print("Error: No vendors found in config.")
        sys.exit(1)

    print("\nAvailable vendors:")
    for i, name in enumerate(vendor_names, start=1):
        print(f"  {i}. {name}")

    while True:
        selection = input(
            f"\nEnter a vendor number (1-{len(vendor_names)}) or name: "
        ).strip()

        if selection.isdigit():
            idx = int(selection)
            if 1 <= idx <= len(vendor_names):
                return vendor_names[idx - 1]
            print(f"  Invalid number. Please enter 1-{len(vendor_names)}.")
            continue

        matches = [name for name in vendor_names if name.lower() ==
                   selection.lower()]
        if matches:
            return matches[0]

        print(f"  '{selection}' not found. Try again.")


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


def load_excluded_skus(spreadsheet):
    """Load excluded SKUs from 'Excluded SKUs' sheet within the Rules Matrix workbook.

    Exclusions stay keyed on SKU (not Token) — this list is hand-maintained
    by staff, and typing a 24-character Token isn't practical. Token-based
    rename protection is for the automated catalog/matrix sync; exclusions
    are a separate, manually-curated concern.
    """
    try:
        worksheet = spreadsheet.worksheet("Excluded SKUs")
        skus = worksheet.col_values(1)  # Get column A

        # Skip header row and empty cells, strip whitespace
        excluded = {sku.strip() for sku in skus[1:] if sku.strip()}

        print(
            f"✓ Loaded {len(excluded)} excluded SKUs from 'Excluded SKUs' sheet")
        return excluded
    except Exception as e:
        print(f"⚠ Warning: Could not load 'Excluded SKUs' sheet: {e}")
        print("  Proceeding with empty exclusion list.")
        print("  Note: Create a sheet named 'Excluded SKUs' in your Rules Matrix workbook.")
        return set()


def col_letter(n):
    """Convert a 1-indexed column number to its A1 letter (1 -> A, 27 -> AA)."""
    letters = ''
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _to_native(val):
    """Convert numpy scalar types (int64, float64, bool_, etc.) — which
    pandas hands back from things like .iloc[] — into plain Python types.
    gspread/requests can't JSON-serialize numpy scalars directly."""
    if hasattr(val, 'item'):
        try:
            return val.item()
        except Exception:
            return val
    return val


def push_to_sheets(spreadsheet, rules_df):
    """Push the rules matrix DataFrame to Google Sheets.

    'Total Max' is written as a live SUM formula (not a value) referencing
    that row's own store Max columns, positioned right before 'Date Added'.
    Because it's a formula, it recalculates automatically in Sheets whenever
    anyone edits a store's Max value directly — no script run required.

    Each store's DNO column is likewise written as a live formula:
    TRUE when that store's Min and Max are both 0, FALSE otherwise. Lock
    the DNO columns in Sheets (Data > Protected sheets and ranges) so staff
    can only edit Min/Max — DNO then follows automatically. The Token
    column should be locked too — it's an internal join key, not something
    staff should hand-edit.
    """
    worksheet = spreadsheet.sheet1

    rules_df = rules_df.fillna('')
    columns = rules_df.columns.tolist()

    max_cols = [f'{code}_Max' for code in store_map.values()
                if f'{code}_Max' in columns]

    insert_at = columns.index(
        'Date Added') if 'Date Added' in columns else len(columns)
    output_columns = columns[:insert_at] + ['Total Max'] + columns[insert_at:]

    max_col_letters = [col_letter(output_columns.index(c) + 1)
                       for c in max_cols]

    # Map each store's DNO column to its own Min/Max column letters, so the
    # DNO formula can reference them directly.
    dno_formula_refs = {}
    for code in store_map.values():
        dno_col, min_col, max_col = f'{code}_DNO', f'{code}_Min', f'{code}_Max'
        if dno_col in output_columns and min_col in output_columns and max_col in output_columns:
            dno_formula_refs[dno_col] = (
                col_letter(output_columns.index(min_col) + 1),
                col_letter(output_columns.index(max_col) + 1),
            )

    header = output_columns
    body = []
    for i in range(len(rules_df)):
        sheet_row = i + 2  # +1 for header, +1 for 1-indexing
        row_vals = rules_df.iloc[i]
        row_out = []
        for col in output_columns:
            if col == 'Total Max':
                refs = ','.join(
                    f'{letter}{sheet_row}' for letter in max_col_letters)
                row_out.append(f'=SUM({refs})' if refs else 0)
            elif col in dno_formula_refs:
                min_letter, max_letter = dno_formula_refs[col]
                row_out.append(
                    f'=IF(AND({min_letter}{sheet_row}=0,{max_letter}{sheet_row}=0),TRUE,FALSE)')
            else:
                row_out.append(_to_native(row_vals[col]))
        body.append(row_out)

    data = [header] + body
    last_col = col_letter(len(output_columns))
    last_row = len(body) + 1

    # Write the full payload to an explicit range WITHOUT clearing first —
    # if the write fails partway (bad data, network blip, etc.), the sheet's
    # existing data is left untouched instead of wiped. Only after a
    # confirmed-successful write do we trim any leftover rows/columns from
    # a previous, larger matrix.
    worksheet.update(f'A1:{last_col}{last_row}', data,
                     value_input_option='USER_ENTERED')

    # update() only writes values/formulas — it never touches cell
    # formatting. Because this sheet is overwritten in place rather than
    # rebuilt, the Total Max column can land on cells that still carry
    # leftover formatting from whatever used to sit there in a prior run
    # (most often Date Added's date format, since Total Max sits right next
    # to it). Force plain-number formatting on the Total Max column every
    # push so it can never inherit a stale format.
    total_max_letter = col_letter(output_columns.index('Total Max') + 1)
    try:
        worksheet.format(
            f'{total_max_letter}2:{total_max_letter}{last_row}',
            {'numberFormat': {'type': 'NUMBER', 'pattern': '0'}}
        )
    except Exception as e:
        print(
            f"⚠ Warning: wrote Total Max values, but couldn't reset its number format: {e}")

    try:
        worksheet.resize(rows=last_row, cols=len(output_columns))
    except Exception as e:
        print(
            f"⚠ Warning: wrote the data successfully, but couldn't trim leftover rows/columns from a previous run: {e}")

    print(
        f"Pushed {len(rules_df)} SKUs to Google Sheets (DNO and Total Max are live).")
    print(f"Sheet URL: {spreadsheet.url}")


def load_catalog(path):
    """Load and validate the catalog file."""
    # NOTE: header=1 assumes the catalog has a blank/title row before the actual headers.
    # Change to header=0 if your headers are on row 1.
    # Read all columns first so a missing required column is reported with a
    # clear message instead of pandas' usecols ValueError.
    catalog = pd.read_excel(path, header=1, dtype={
                            'SKU': str, 'Token': str, 'Item Name': str})

    missing_cols = REQUIRED_CATALOG_COLS - set(catalog.columns)
    if missing_cols:
        print(f"Error: Catalog is missing required columns: {missing_cols}")
        sys.exit(1)

    catalog = catalog[list(REQUIRED_CATALOG_COLS)]
    catalog = catalog.dropna(
        subset=['Token']).drop_duplicates(subset=['Token'])

    return catalog


def _migrate_to_token(existing_df, catalog):
    """One-time migration for sheets synced before Token support was added.

    Matches each existing row to the catalog by SKU (the old join key) just
    once, to backfill a Token column. From the next sync onward, that
    sheet has Token and never needs this fallback again.

    Prints a summary so the user has a clean list of anything that didn't
    match automatically, rather than having to hunt through console output.
    """
    print("No Token column found on sheet — migrating from SKU-based matching...")

    existing_df = pd.merge(
        existing_df, catalog[['SKU', 'Token']], on='SKU', how='left')

    matched_mask = existing_df['Token'].notna() & (existing_df['Token'] != '')
    matched = existing_df[matched_mask]
    unmatched = existing_df[~matched_mask]

    print("\n" + "="*60)
    print("  TOKEN MIGRATION SUMMARY")
    print("="*60)
    print(f"  Rows matched to a Token:   {len(matched)}/{len(existing_df)}")
    print(f"  Rows NOT matched:          {len(unmatched)}/{len(existing_df)}")

    if not unmatched.empty:
        name_col = 'Item Name' if 'Item Name' in unmatched.columns else None
        print("\n  Unmatched rows (kept, but still SKU-matched only until")
        print("  their SKU is confirmed against a future catalog):")
        for _, row in unmatched.iterrows():
            label = f"{row['SKU']}"
            if name_col and pd.notna(row.get(name_col)):
                label += f" — {row[name_col]}"
            print(f"    - {label}")
        print("\n  These rows kept all their existing Min/Max/DNO values —")
        print("  nothing was reset. They just aren't Token-protected yet.")
    print("="*60)

    return existing_df


def load_matrix_from_sheet(spreadsheet, catalog, create_new=False):
    """Build the working rules_df either from scratch or by reading the
    current matrix directly from the Google Sheet. No local file is
    involved — the Sheet is the source of truth.

    create_new=True skips reading any existing sheet data entirely (for a
    vendor's first-ever sync, or to intentionally rebuild from the catalog).
    """
    if create_new:
        rules_df = catalog.copy()
        rules_df['Order In Quantities'] = DEFAULT_ORDER_QTY
        rules_df['Date Added'] = date.today().isoformat()
        return rules_df

    try:
        records = spreadsheet.sheet1.get_all_records(
            value_render_option='UNFORMATTED_VALUE')
    except Exception as e:
        print(f"Error: Could not read existing sheet data: {e}")
        print("Aborting WITHOUT writing anything, so the sheet isn't overwritten")
        print("with defaults while this is unresolved. Fix the read issue and")
        print("re-run, or re-run and choose 'N' (create new) if this vendor")
        print("genuinely has no sheet data yet.")
        sys.exit(1)

    if not records:
        print("\n" + "="*60)
        print("  No existing data found on the sheet.")
        print("  If this is expected (a genuinely new/empty vendor sheet),")
        print("  that's fine to continue. If it's NOT expected, stop now —")
        print("  continuing will push a fresh matrix with every Min/Max")
        print("  reset to 0, overwriting whatever is really on the sheet.")
        print("="*60)
        confirm = input(
            "Build a fresh matrix from the catalog anyway? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Aborting. Nothing was changed.")
            sys.exit(0)
        rules_df = catalog.copy()
        rules_df['Order In Quantities'] = DEFAULT_ORDER_QTY
        rules_df['Date Added'] = date.today().isoformat()
        return rules_df

    existing_df = pd.DataFrame(records)
    existing_df['SKU'] = existing_df['SKU'].astype(str)

    # Total Max is a formula we regenerate on every push, not state to keep.
    existing_df = existing_df.drop(columns=['Total Max'], errors='ignore')

    # First sync since Token support was added: backfill Token via SKU match.
    if 'Token' not in existing_df.columns:
        existing_df = _migrate_to_token(existing_df, catalog)
    else:
        existing_df['Token'] = existing_df['Token'].astype(str)

    if 'Order_Qty' in existing_df.columns:
        existing_df = existing_df.rename(
            columns={'Order_Qty': 'Order In Quantities'})

    # Item Name / Reporting Category / SKU normally refresh from the catalog
    # every sync — matched on Token now, so a vendor renaming a SKU just
    # updates this row's SKU/name in place instead of orphaning the rules.
    # For a Token with no catalog match (e.g. a discontinued item the user
    # chose to keep, or an unmatched migration row with no Token at all),
    # there's nothing to refresh FROM, so fall back to what's already there.
    merge_cols = ['Token', 'SKU', 'Item Name', 'Reporting Category']
    rules_df = pd.merge(
        existing_df, catalog[merge_cols],
        on='Token', how='left', suffixes=('_existing', '_catalog'))

    for col in ['SKU', 'Item Name', 'Reporting Category']:
        existing_col = f'{col}_existing'
        catalog_col = f'{col}_catalog'
        if catalog_col in rules_df.columns:
            rules_df[col] = rules_df[catalog_col].fillna(
                rules_df.get(existing_col))
        elif existing_col in rules_df.columns:
            rules_df[col] = rules_df[existing_col]

    rules_df = rules_df.drop(
        columns=[c for c in rules_df.columns
                 if c.endswith('_existing') or c.endswith('_catalog')],
        errors='ignore')

    if 'Order In Quantities' not in rules_df.columns:
        rules_df['Order In Quantities'] = DEFAULT_ORDER_QTY
    else:
        rules_df['Order In Quantities'] = pd.to_numeric(
            rules_df['Order In Quantities'], errors='coerce').fillna(DEFAULT_ORDER_QTY)

    today_str = date.today().isoformat()
    if 'Date Added' not in rules_df.columns:
        rules_df['Date Added'] = today_str
    else:
        missing_dates = rules_df['Date Added'].isna() | (
            rules_df['Date Added'] == '')
        if missing_dates.any():
            rules_df.loc[missing_dates, 'Date Added'] = today_str
            print(
                f"Note: {missing_dates.sum()} SKU(s) had no 'Date Added' — backfilled with today's date ({today_str}).")

    # Values come back from Sheets as native types via UNFORMATTED_VALUE, but
    # normalize defensively in case a cell was left blank or typed oddly.
    for code in store_map.values():
        dno_col, min_col, max_col = f'{code}_DNO', f'{code}_Min', f'{code}_Max'
        if dno_col in rules_df.columns:
            rules_df[dno_col] = rules_df[dno_col].apply(
                lambda v: str(v).upper() == 'TRUE' or v is True)
        if min_col in rules_df.columns:
            rules_df[min_col] = pd.to_numeric(
                rules_df[min_col], errors='coerce').fillna(DEFAULT_MIN)
        if max_col in rules_df.columns:
            rules_df[max_col] = pd.to_numeric(
                rules_df[max_col], errors='coerce').fillna(DEFAULT_MAX)

    return rules_df


def remove_excluded_skus(rules_df, excluded_set):
    """Remove any SKUs that are on the exclusion list.

    Stays SKU-keyed on purpose — see load_excluded_skus().
    """
    excluded = rules_df[rules_df['SKU'].isin(excluded_set)]

    if not excluded.empty:
        print(f"\nRemoving {len(excluded)} excluded SKU(s):")
        for sku in excluded['SKU'].tolist():
            print(f"  - {sku}")
        rules_df = rules_df[~rules_df['SKU'].isin(
            excluded_set)].reset_index(drop=True)
    else:
        print("No excluded SKUs found in matrix.")

    return rules_df


def remove_discontinued_skus(rules_df, catalog):
    """Remove rows from the matrix whose Token is no longer in the catalog.

    Token-based now: a plain SKU rename won't show up here anymore, since
    the Token still matches. Only genuinely discontinued items — or
    migration rows that never got a Token — land in this list.
    """
    catalog_tokens = set(catalog['Token'])
    has_token = rules_df['Token'].notna() & (rules_df['Token'] != '')
    discontinued = rules_df[has_token & ~
                            rules_df['Token'].isin(catalog_tokens)]

    if not discontinued.empty:
        print(f"\nRemoving {len(discontinued)} discontinued item(s):")
        for _, row in discontinued.iterrows():
            print(f"  - {row['SKU']} — {row.get('Item Name', '')}")

        confirm = input(
            "\nConfirm removal of discontinued items? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Removal cancelled. Discontinued items will be kept.")
            return rules_df

        keep_mask = ~(has_token & ~rules_df['Token'].isin(catalog_tokens))
        rules_df = rules_df[keep_mask].reset_index(drop=True)
        print("Discontinued items removed.")
    else:
        print("No discontinued items found.")

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
    """Add any catalog items whose Token isn't already in the matrix."""
    existing_tokens = set(rules_df['Token'].dropna())
    new_items = catalog[~catalog['Token'].isin(existing_tokens)].copy()

    if not new_items.empty:
        new_items['Order In Quantities'] = DEFAULT_ORDER_QTY
        new_items['Date Added'] = date.today().isoformat()
        for code in store_map.values():
            new_items[f'{code}_DNO'] = True
            new_items[f'{code}_Min'] = DEFAULT_MIN
            new_items[f'{code}_Max'] = DEFAULT_MAX
        rules_df = pd.concat([rules_df, new_items], ignore_index=True)
        print(f"Added {len(new_items)} new item(s) to the matrix.")
    else:
        print("No new items found.")

    return rules_df


def sync_rules_matrix(vendor, catalog_path, create_new=False):
    """Main logic to sync the rules matrix and push to Google Sheets."""
    try:
        # Load Google client and connect to vendor sheet
        print("\nConnecting to Google Sheets...")
        client = get_google_client()
        spreadsheet = get_or_create_sheet(client, vendor)

        # Load excluded SKUs from the same Rules Matrix spreadsheet
        excluded_skus = load_excluded_skus(spreadsheet)

        catalog = load_catalog(catalog_path)
        # Filter excluded SKUs from catalog
        catalog = catalog[~catalog['SKU'].isin(excluded_skus)]

        if create_new:
            print(
                "\nBuilding a new matrix from the catalog (skipping existing sheet data)...")
        else:
            print("\nReading current matrix from Google Sheets...")
        rules_df = load_matrix_from_sheet(
            spreadsheet, catalog, create_new=create_new)

        # Remove excluded SKUs from matrix
        rules_df = remove_excluded_skus(rules_df, excluded_skus)
        rules_df = remove_discontinued_skus(rules_df, catalog)
        rules_df = ensure_store_columns(rules_df)
        rules_df = append_new_skus(rules_df, catalog)

        matrix_columns = ['Token', 'SKU', 'Item Name',
                          'Reporting Category', 'Order In Quantities']
        for code in store_map.values():
            matrix_columns.extend(
                [f'{code}_DNO', f'{code}_Min', f'{code}_Max'])
        matrix_columns.append('Date Added')

        final_cols = [c for c in matrix_columns if c in rules_df.columns]
        rules_df = rules_df[final_cols]

        # Push to Google Sheets (DNO and Total Max are both inserted as live
        # formulas — DNO follows each store's Min/Max, Total Max sums them)
        push_to_sheets(spreadsheet, rules_df)

        print("\nSuccess! Matrix synced to Google Sheets.")

    except FileNotFoundError as e:
        print(f"Error: File not found — {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    Vendor = select_vendor_interactively()

    # --- MATRIX MODE ---
    print("\n" + "="*60)
    print("  MATRIX MODE")
    print("="*60)
    print("  [N] Create a new matrix")
    print("      (first sync for this vendor — skips reading any")
    print("       existing sheet data and builds fresh from the catalog)")
    print("  [U] Update the current matrix")
    print("      (pulls current values from the sheet, then merges")
    print("       in catalog changes)")
    mode = input("\nCreate new or update current? (N/U): ").strip().lower()
    CREATE_NEW = mode == 'n'

    # --- PROCEED ---
    print("\nPlease select the Catalog Excel file...")
    CATALOG_PATH = get_file_path(title=f"Select Catalog for {Vendor}")
    if not CATALOG_PATH:
        print("No file selected. Exiting.")
        sys.exit(0)

    sync_rules_matrix(Vendor, CATALOG_PATH, create_new=CREATE_NEW)
