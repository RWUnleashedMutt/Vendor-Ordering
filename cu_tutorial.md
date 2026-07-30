# Building the Rules Matrix Sync Script From Scratch

This is a from-the-ground-up build guide — not just a "how it works" reference, but a walkthrough you could follow to rebuild `sync_rules_matrix.py` yourself if you ever needed to, including the reasoning behind each design decision and the real bugs that shaped it.

Build it in the order below. Each stage is a working (if incomplete) version — test after each one rather than writing the whole thing and debugging at the end.

---

## Stage 0 — What you're building, and why it's shaped this way

**The problem:** Each vendor has a Google Sheet — the "Rules Matrix" — that tells the ordering system, per store, whether to order an item (`DNO` = Do Not Order), and if so, the Min/Max quantities to stock. Staff hand-edit Min/Max in that sheet. Separately, vendors periodically send updated product catalogs (new items, discontinued items, renamed SKUs).

**The goal:** a script that reconciles the two — pull the latest catalog, merge it into the existing sheet without clobbering staff's hand-edited Min/Max values, and push the result back.

**Three hard requirements fall out of that:**

1. The Google Sheet is the only source of truth — never a local cached file. Two people running stale local copies is how you get silent data loss.
2. Merging catalog changes must **never** overwrite a Min/Max a human already set, unless the row is genuinely new.
3. Renaming a SKU must not look like "delete old item, add new item" — that would reset all its Min/Max settings to defaults.

Keep these three in your head — nearly every non-obvious line in this script exists to satisfy one of them.

---

## Stage 1 — Skeleton: connect and read a catalog

Start with just enough to prove you can authenticate and open a sheet.

```python
import pandas as pd
import tomllib
import sys
import gspread
from google.oauth2.service_account import Credentials

CREDENTIALS_FILE = './credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

with open('./.streamlit/secrets.toml', 'rb') as f:
    config = tomllib.load(f)
SHEET_IDS = config['sheet_ids']

def get_google_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)

def get_or_create_sheet(client, vendor):
    if vendor not in SHEET_IDS:
        print(f"Error: No Sheet ID found for vendor '{vendor}'.")
        sys.exit(1)
    spreadsheet = client.open_by_key(SHEET_IDS[vendor])
    print(f"Connected to Google Sheet for {vendor}")
    return spreadsheet
```

**Setup you need before this runs:**

- A Google Cloud service account with the Sheets and Drive APIs enabled, its JSON key saved as `credentials.json`.
- The service account's email address added as an Editor on every vendor Sheet (service accounts don't inherit your personal Drive access).
- `.streamlit/secrets.toml` with a `[sheet_ids]` table mapping vendor name → Sheet ID (the long string in the Sheet's URL between `/d/` and `/edit`).

Test: call `get_or_create_sheet(get_google_client(), "SomeVendor")` and confirm it prints the connected message with no traceback.

---

## Stage 2 — Define the store structure

Every store gets a two/three-letter code used as a column prefix. Hardcode the mapping from full store name (as it might appear in other exports) to short code:

```python
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
```

Every store in the matrix will get three columns: `<code>_DNO`, `<code>_Min`, `<code>_Max`. This dict is the single place you'd edit if a store opens, closes, or gets renamed.

Also define your defaults up top, since you'll reference them in several places:

```python
DEFAULT_MIN = 0
DEFAULT_MAX = 0
DEFAULT_ORDER_QTY = 1
REQUIRED_CATALOG_COLS = {'Token', 'SKU', 'Item Name', 'Reporting Category'}
```

---

## Stage 3 — Load and validate the catalog file

```python
def load_catalog(path):
    catalog = pd.read_excel(path, header=1, dtype={
        'SKU': str, 'Token': str, 'Item Name': str})

    missing_cols = REQUIRED_CATALOG_COLS - set(catalog.columns)
    if missing_cols:
        print(f"Error: Catalog is missing required columns: {missing_cols}")
        sys.exit(1)

    catalog = catalog[list(REQUIRED_CATALOG_COLS)]
    catalog = catalog.dropna(subset=['Token']).drop_duplicates(subset=['Token'])
    return catalog
```

Two details worth knowing _why_ they're there:

- **`header=1`**: this vendor export format has a blank/title row above the real headers. If your export's headers are on row 1 instead, this needs to be `header=0`. Check with `pd.read_excel(path, header=None).head()` first if you're unsure.
- **`dtype={'SKU': str, 'Token': str, ...}`**: without this, pandas will guess a numeric type for SKU/Token columns that look number-like, silently corrupting anything with a leading zero or non-numeric character. Force them to strings on read, always.
- **Drop rows with no Token, then de-duplicate on Token.** Token is about to become your primary key for everything downstream — a catalog row without one, or two rows sharing one, would break every merge that follows.

**Test:** load a real catalog export and check `catalog.dtypes` — Token and SKU should both read as `object` (string), not `int64` or `float64`.

---

## Stage 4 — A file picker for interactive use

```python
import tkinter as tk
from tkinter import filedialog

def get_file_path(title="Select File", file_types=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(title=title, filetypes=file_types)
    root.destroy()
    return file_path
```

Nothing subtle here — `withdraw()` hides the empty root Tk window so only the file dialog shows, and `-topmost` keeps it from opening behind other windows.

---

## Stage 5 — Vendor selection prompt

Rather than free-typing a vendor name (error-prone — typos, case mismatches), list what's actually in config and let the user pick by number or exact name:

```python
def select_vendor_interactively():
    vendor_names = list(SHEET_IDS.keys())
    if not vendor_names:
        print("Error: No vendors found in config.")
        sys.exit(1)

    print("\nAvailable vendors:")
    for i, name in enumerate(vendor_names, start=1):
        print(f"  {i}. {name}")

    while True:
        selection = input(f"\nEnter a vendor number (1-{len(vendor_names)}) or name: ").strip()
        if selection.isdigit():
            idx = int(selection)
            if 1 <= idx <= len(vendor_names):
                return vendor_names[idx - 1]
            print(f"  Invalid number. Please enter 1-{len(vendor_names)}.")
            continue
        matches = [name for name in vendor_names if name.lower() == selection.lower()]
        if matches:
            return matches[0]
        print(f"  '{selection}' not found. Try again.")
```

This pattern (numbered list, accept number-or-name, loop until valid) is worth reusing anywhere else you need a vendor picker in a companion script — keeping the UX identical across scripts avoids a "wait, which script wants the number and which wants the name" mixup.

---

## Stage 6 — Handling numpy types before they hit gspread

This one you'll only discover the hard way if you skip it: pandas hands back `numpy.int64`, `numpy.float64`, `numpy.bool_` from things like `.iloc[]` — and gspread/requests cannot JSON-serialize those. You'll get `Object of type int64 is not JSON serializable` the first time you try to push a real DataFrame.

Fix it once, centrally:

```python
def _to_native(val):
    if hasattr(val, 'item'):
        try:
            return val.item()
        except Exception:
            return val
    return val
```

`.item()` is numpy's own escape hatch back to a plain Python scalar. Call this on every value right before it goes into the payload you send to Sheets.

Also grab this small helper now, since the push step needs it for building formula strings:

```python
def col_letter(n):
    """1-indexed column number -> A1 letter (1 -> A, 27 -> AA)."""
    letters = ''
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters
```

---

## Stage 7 — First version of the push (and the bug that breaks it)

A naive first pass looks like this:

```python
def push_to_sheets_v1(spreadsheet, rules_df):
    worksheet = spreadsheet.sheet1
    worksheet.clear()
    worksheet.update([rules_df.columns.tolist()] + rules_df.values.tolist())
```

**Do not ship this.** Here's the failure mode, exactly as it happened during development: if anything in the payload throws partway through building it (say, an unconverted `int64` triggering the JSON serialization error from Stage 6), `worksheet.clear()` has _already run_. The sheet is now empty, the write never completed, and every hand-edited Min/Max on that vendor's matrix is gone.

The fix is two-part, and both parts matter independently:

**Part A — never clear before you're sure the write will succeed.** Write to an explicit range instead:

```python
def push_to_sheets_v2(spreadsheet, rules_df):
    worksheet = spreadsheet.sheet1
    header = rules_df.columns.tolist()
    body = rules_df.values.tolist()
    data = [header] + body

    last_col = col_letter(len(header))
    last_row = len(body) + 1

    worksheet.update(f'A1:{last_col}{last_row}', data, value_input_option='USER_ENTERED')

    # Only trim leftover rows/cols from a previous, larger matrix AFTER a
    # confirmed-successful write — never before.
    try:
        worksheet.resize(rows=last_row, cols=len(header))
    except Exception as e:
        print(f"⚠ Warning: wrote the data, but couldn't trim leftover rows/columns: {e}")
```

Writing to a bounded range means a failed write leaves the existing sheet exactly as it was — nothing is destroyed until the new data has actually landed successfully.

**Part B — treat a failed or empty read as a hard stop, not "start fresh."** This matters on the _read_ side (Stage 9), but it's the same underlying principle: an ambiguous state (did the read fail, or is the sheet genuinely empty?) must never silently resolve to "push defaults over everything."

---

## Stage 8 — Building the real payload: formulas, not values, for two columns

Two columns in the final matrix should never be static values, because they're both fully derivable from other columns in the same row, and staff should be able to edit an input (Min/Max) and see the derived column update instantly without waiting on a script run:

- **`Total Max`** — the sum of every store's `_Max` column.
- **`<store>_DNO`** — `TRUE` when that store's Min and Max are both 0, else `FALSE`.

To write a live formula through gspread, you just put the formula string as the cell value with `value_input_option='USER_ENTERED'` (this is what tells Sheets "parse this like a human typed it," rather than `'RAW'`, which would store the literal string `=SUM(...)` instead of evaluating it).

You need each row's actual A1 cell references, which means you need to know the final column layout _before_ you build the row data — column position drives the formula text.

```python
def build_output_columns(columns):
    """Insert 'Total Max' right before 'Date Added' (or at the end, if
    'Date Added' isn't present)."""
    insert_at = columns.index('Date Added') if 'Date Added' in columns else len(columns)
    return columns[:insert_at] + ['Total Max'] + columns[insert_at:]
```

Then, for each row, build the formula strings:

```python
def push_to_sheets(spreadsheet, rules_df):
    worksheet = spreadsheet.sheet1
    rules_df = rules_df.fillna('')
    columns = rules_df.columns.tolist()

    max_cols = [f'{code}_Max' for code in store_map.values() if f'{code}_Max' in columns]
    output_columns = build_output_columns(columns)
    max_col_letters = [col_letter(output_columns.index(c) + 1) for c in max_cols]

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
        sheet_row = i + 2  # header row + 1-indexing
        row_vals = rules_df.iloc[i]
        row_out = []
        for col in output_columns:
            if col == 'Total Max':
                refs = ','.join(f'{letter}{sheet_row}' for letter in max_col_letters)
                row_out.append(f'=SUM({refs})' if refs else 0)
            elif col in dno_formula_refs:
                min_letter, max_letter = dno_formula_refs[col]
                row_out.append(f'=IF(AND({min_letter}{sheet_row}=0,{max_letter}{sheet_row}=0),TRUE,FALSE)')
            else:
                row_out.append(_to_native(row_vals[col]))
        body.append(row_out)

    data = [header] + body
    last_col = col_letter(len(output_columns))
    last_row = len(body) + 1
    worksheet.update(f'A1:{last_col}{last_row}', data, value_input_option='USER_ENTERED')
```

The reasoning behind `sheet_row = i + 2`: row 1 is the header, and Sheets is 1-indexed, so DataFrame row 0 lands on sheet row 2.

**Test this stage in isolation** on a throwaway test sheet: push a few rows, open the sheet, and confirm Total Max and DNO show up as actual formulas (click the cell, check the formula bar) — not as plain numbers/booleans that merely happen to be correct.

---

## Stage 9 — A bug you'll hit here: formatting doesn't move with the data

`worksheet.update()` only ever writes **values and formulas**. It never touches a cell's number format, font, or fill. That's invisible right up until a column's position shifts between runs — say, you add a new column (like Token) at the front of the matrix, shifting everything one slot right. Whatever cell `Total Max` used to _not_ occupy, it now does — and if that cell previously belonged to a date-formatted column (`Date Added`, which sits right next to it), the leftover date format is still sitting there. The SUM formula that lands in it is correct, but it displays as a date.

Fix: explicitly force the number format on that column, every push, regardless of where it lands:

```python
    total_max_letter = col_letter(output_columns.index('Total Max') + 1)
    try:
        worksheet.format(
            f'{total_max_letter}2:{total_max_letter}{last_row}',
            {'numberFormat': {'type': 'NUMBER', 'pattern': '0'}}
        )
    except Exception as e:
        print(f"⚠ Warning: wrote Total Max values, but couldn't reset its number format: {e}")
```

The general lesson, if you take nothing else from this stage: **any script that overwrites a sheet's contents in place (rather than rebuilding the sheet from nothing) must explicitly own the formatting of any column it writes formulas or computed values into.** Don't assume old formatting is harmless just because it's invisible in your test data.

---

## Stage 10 — Reading the existing matrix back (the merge logic)

This is the most important function in the whole script, because it's where your three Stage-0 requirements actually get enforced.

```python
def load_matrix_from_sheet(spreadsheet, catalog, create_new=False):
    if create_new:
        rules_df = catalog.copy()
        rules_df['Order In Quantities'] = DEFAULT_ORDER_QTY
        rules_df['Date Added'] = date.today().isoformat()
        return rules_df

    try:
        records = spreadsheet.sheet1.get_all_records(value_render_option='UNFORMATTED_VALUE')
    except Exception as e:
        print(f"Error: Could not read existing sheet data: {e}")
        print("Aborting WITHOUT writing anything.")
        sys.exit(1)
```

Two choices here worth calling out:

- **`create_new` skips reading the sheet entirely.** This is for a vendor's very first sync — there's nothing to merge with yet, so don't even try.
- **A failed read is a hard `sys.exit(1)`, never a fallback to "build fresh."** If you can't confirm what's currently on the sheet, the only safe move is to stop. Silently treating "I couldn't read it" the same as "there's nothing there" is exactly how you'd overwrite real data with a blank matrix.

Next, handle the case where the read _succeeded_ but came back empty — genuinely ambiguous, so ask:

```python
    if not records:
        print("No existing data found on the sheet. Confirm before rebuilding from catalog.")
        confirm = input("Build a fresh matrix from the catalog anyway? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Aborting. Nothing was changed.")
            sys.exit(0)
        rules_df = catalog.copy()
        rules_df['Order In Quantities'] = DEFAULT_ORDER_QTY
        rules_df['Date Added'] = date.today().isoformat()
        return rules_df
```

Now the actual merge. Convert to a DataFrame, drop `Total Max` (it's a formula you regenerate every push, never state you carry forward):

```python
    existing_df = pd.DataFrame(records)
    existing_df['SKU'] = existing_df['SKU'].astype(str)
    existing_df = existing_df.drop(columns=['Total Max'], errors='ignore')
```

Handle sheets that predate the Token column (see Stage 11), otherwise make sure Token is a string:

```python
    if 'Token' not in existing_df.columns:
        existing_df = _migrate_to_token(existing_df, catalog)
    else:
        existing_df['Token'] = existing_df['Token'].astype(str)
```

**The core merge — this is what makes SKU renames safe.** Merge on `Token`, not `SKU`:

```python
    merge_cols = ['Token', 'SKU', 'Item Name', 'Reporting Category']
    rules_df = pd.merge(
        existing_df, catalog[merge_cols],
        on='Token', how='left', suffixes=('_existing', '_catalog'))

    for col in ['SKU', 'Item Name', 'Reporting Category']:
        existing_col = f'{col}_existing'
        catalog_col = f'{col}_catalog'
        if catalog_col in rules_df.columns:
            rules_df[col] = rules_df[catalog_col].fillna(rules_df.get(existing_col))
        elif existing_col in rules_df.columns:
            rules_df[col] = rules_df[existing_col]

    rules_df = rules_df.drop(
        columns=[c for c in rules_df.columns if c.endswith('_existing') or c.endswith('_catalog')],
        errors='ignore')
```

Why `how='left'` and `fillna` rather than just taking the catalog value outright: if a Token has no match in the current catalog (a discontinued item you're choosing to keep on the sheet a while longer, or an unmatched migration row), there's nothing to refresh _from_ — you want to fall back to whatever's already on the sheet rather than blanking the field.

Finally, backfill anything still missing and normalize types:

```python
    if 'Order In Quantities' not in rules_df.columns:
        rules_df['Order In Quantities'] = DEFAULT_ORDER_QTY
    else:
        rules_df['Order In Quantities'] = pd.to_numeric(
            rules_df['Order In Quantities'], errors='coerce').fillna(DEFAULT_ORDER_QTY)

    today_str = date.today().isoformat()
    if 'Date Added' not in rules_df.columns:
        rules_df['Date Added'] = today_str
    else:
        missing_dates = rules_df['Date Added'].isna() | (rules_df['Date Added'] == '')
        if missing_dates.any():
            rules_df.loc[missing_dates, 'Date Added'] = today_str

    for code in store_map.values():
        dno_col, min_col, max_col = f'{code}_DNO', f'{code}_Min', f'{code}_Max'
        if dno_col in rules_df.columns:
            rules_df[dno_col] = rules_df[dno_col].apply(lambda v: str(v).upper() == 'TRUE' or v is True)
        if min_col in rules_df.columns:
            rules_df[min_col] = pd.to_numeric(rules_df[min_col], errors='coerce').fillna(DEFAULT_MIN)
        if max_col in rules_df.columns:
            rules_df[max_col] = pd.to_numeric(rules_df[max_col], errors='coerce').fillna(DEFAULT_MAX)

    return rules_df
```

`value_render_option='UNFORMATTED_VALUE'` on the read (back at the top) is what gives you real booleans/numbers instead of display strings in the first place — the normalization above is just a defensive backstop for cells that were hand-typed oddly.

---

## Stage 11 — The one-time Token migration

Any sheet that was created before Token support existed won't have a Token column. You need to backfill it exactly once, by falling back to the old join key (SKU):

```python
def _migrate_to_token(existing_df, catalog):
    print("No Token column found on sheet — migrating from SKU-based matching...")
    existing_df = pd.merge(existing_df, catalog[['SKU', 'Token']], on='SKU', how='left')

    matched_mask = existing_df['Token'].notna() & (existing_df['Token'] != '')
    matched = existing_df[matched_mask]
    unmatched = existing_df[~matched_mask]

    print(f"Rows matched to a Token:   {len(matched)}/{len(existing_df)}")
    print(f"Rows NOT matched:          {len(unmatched)}/{len(existing_df)}")

    if not unmatched.empty:
        name_col = 'Item Name' if 'Item Name' in unmatched.columns else None
        for _, row in unmatched.iterrows():
            label = f"{row['SKU']}"
            if name_col and pd.notna(row.get(name_col)):
                label += f" — {row[name_col]}"
            print(f"    - {label}")
        print("These rows kept all their existing Min/Max/DNO values — nothing was reset.")

    return existing_df
```

The important guarantee: an unmatched row (a SKU typo, a truly orphaned entry) **still keeps its existing data**. It just doesn't get Token-protection until its SKU is confirmed against a future catalog sync. You never delete or reset a row just because migration couldn't match it — print it so a human can look, and move on.

Once every sheet has gone through this once, `'Token' not in existing_df.columns` will be false on all future runs and this function never runs again for that vendor.

---

## Stage 12 — Exclusions, discontinued items, new items

Three straightforward cleanup passes, applied in this order after the merge:

```python
def load_excluded_skus(spreadsheet):
    try:
        worksheet = spreadsheet.worksheet("Excluded SKUs")
        skus = worksheet.col_values(1)
        excluded = {sku.strip() for sku in skus[1:] if sku.strip()}
        print(f"✓ Loaded {len(excluded)} excluded SKUs")
        return excluded
    except Exception as e:
        print(f"⚠ Warning: Could not load 'Excluded SKUs' sheet: {e}")
        return set()

def remove_excluded_skus(rules_df, excluded_set):
    excluded = rules_df[rules_df['SKU'].isin(excluded_set)]
    if not excluded.empty:
        print(f"Removing {len(excluded)} excluded SKU(s)")
        rules_df = rules_df[~rules_df['SKU'].isin(excluded_set)].reset_index(drop=True)
    return rules_df
```

Note exclusions stay keyed on **SKU**, not Token — deliberately. This list is hand-maintained by staff in a separate sheet tab; asking someone to type a 24-character opaque Token into a spreadsheet by hand isn't realistic. Token-based protection is for the automated sync; exclusions are a manual, human-facing concern, so they use the human-facing key.

```python
def remove_discontinued_skus(rules_df, catalog):
    catalog_tokens = set(catalog['Token'])
    has_token = rules_df['Token'].notna() & (rules_df['Token'] != '')
    discontinued = rules_df[has_token & ~rules_df['Token'].isin(catalog_tokens)]

    if not discontinued.empty:
        print(f"Removing {len(discontinued)} discontinued item(s):")
        for _, row in discontinued.iterrows():
            print(f"  - {row['SKU']} — {row.get('Item Name', '')}")
        confirm = input("Confirm removal of discontinued items? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Removal cancelled.")
            return rules_df
        keep_mask = ~(has_token & ~rules_df['Token'].isin(catalog_tokens))
        rules_df = rules_df[keep_mask].reset_index(drop=True)
    return rules_df
```

This is Token-based now, so a plain SKU rename no longer shows up here — only items genuinely missing from the new catalog. It's also the one destructive action in the whole pipeline, which is why it's the one place that stops and asks for explicit confirmation rather than proceeding automatically.

```python
def ensure_store_columns(rules_df):
    for code in store_map.values():
        if f'{code}_DNO' not in rules_df.columns:
            rules_df[f'{code}_DNO'] = False
        if f'{code}_Min' not in rules_df.columns:
            rules_df[f'{code}_Min'] = DEFAULT_MIN
        if f'{code}_Max' not in rules_df.columns:
            rules_df[f'{code}_Max'] = DEFAULT_MAX
    return rules_df

def append_new_skus(rules_df, catalog):
    existing_tokens = set(rules_df['Token'].dropna())
    new_items = catalog[~catalog['Token'].isin(existing_tokens)].copy()
    if not new_items.empty:
        new_items['Order In Quantities'] = DEFAULT_ORDER_QTY
        new_items['Date Added'] = date.today().isoformat()
        for code in store_map.values():
            new_items[f'{code}_DNO'] = True   # default to "don't order" until reviewed
            new_items[f'{code}_Min'] = DEFAULT_MIN
            new_items[f'{code}_Max'] = DEFAULT_MAX
        rules_df = pd.concat([rules_df, new_items], ignore_index=True)
        print(f"Added {len(new_items)} new item(s) to the matrix.")
    return rules_df
```

New items default every store's DNO to `True` deliberately — a brand-new SKU shouldn't silently start getting ordered everywhere until someone's actually reviewed and set real Min/Max values for it.

---

## Stage 13 — Wiring it all together

```python
def sync_rules_matrix(vendor, catalog_path, create_new=False):
    try:
        client = get_google_client()
        spreadsheet = get_or_create_sheet(client, vendor)

        excluded_skus = load_excluded_skus(spreadsheet)
        catalog = load_catalog(catalog_path)
        catalog = catalog[~catalog['SKU'].isin(excluded_skus)]

        rules_df = load_matrix_from_sheet(spreadsheet, catalog, create_new=create_new)
        rules_df = remove_excluded_skus(rules_df, excluded_skus)
        rules_df = remove_discontinued_skus(rules_df, catalog)
        rules_df = ensure_store_columns(rules_df)
        rules_df = append_new_skus(rules_df, catalog)

        matrix_columns = ['Token', 'SKU', 'Item Name', 'Reporting Category', 'Order In Quantities']
        for code in store_map.values():
            matrix_columns.extend([f'{code}_DNO', f'{code}_Min', f'{code}_Max'])
        matrix_columns.append('Date Added')

        final_cols = [c for c in matrix_columns if c in rules_df.columns]
        rules_df = rules_df[final_cols]

        push_to_sheets(spreadsheet, rules_df)
        print("Success! Matrix synced to Google Sheets.")

    except FileNotFoundError as e:
        print(f"Error: File not found — {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)
```

The order here isn't arbitrary: exclude, _then_ check discontinued, _then_ ensure store columns exist, _then_ append new items. Reordering these (say, appending new items before excluding) risks adding a row you're about to filter back out, or evaluating "discontinued" against a catalog that hasn't had exclusions applied yet.

Explicitly rebuilding `matrix_columns` and reindexing at the end guarantees column order in the final sheet is always the same, regardless of what order columns happened to come back from the Sheets read or the merge.

---

## Stage 14 — The CLI entry point

```python
if __name__ == "__main__":
    Vendor = select_vendor_interactively()

    print("[N] Create a new matrix (first sync — skips reading the sheet)")
    print("[U] Update the current matrix (merges catalog changes into existing data)")
    mode = input("Create new or update current? (N/U): ").strip().lower()
    CREATE_NEW = mode == 'n'

    CATALOG_PATH = get_file_path(title=f"Select Catalog for {Vendor}")
    if not CATALOG_PATH:
        print("No file selected. Exiting.")
        sys.exit(0)

    sync_rules_matrix(Vendor, CATALOG_PATH, create_new=CREATE_NEW)
```

---

## Stage 15 — Test plan before trusting it on a real vendor

Work through these on a scratch Google Sheet (a copy of a real vendor sheet, or a small hand-built one) before running against production data:

1. **Fresh vendor, `N` mode** — confirm it builds a matrix from the catalog with sensible defaults and doesn't attempt to read the sheet.
2. **`U` mode, no changes in catalog** — confirm every existing Min/Max value is untouched after the run.
3. **`U` mode, one SKU renamed in the catalog (same Token)** — confirm that row's SKU/name update in place, and its Min/Max are unchanged.
4. **`U` mode, one item removed from the catalog** — confirm it's listed and requires your `yes` before being removed; confirm saying `no` leaves it in place.
5. **`U` mode, one new item in the catalog** — confirm it's appended with DNO=True everywhere and default Min/Max.
6. **A SKU on the Excluded SKUs tab** — confirm it never appears in the pushed matrix, even if it's in the catalog.
7. **Open the pushed sheet directly** — click a Total Max cell and a DNO cell, confirm the formula bar shows the formula (not a static value), and manually edit a Max value to confirm Total Max and DNO update live without rerunning the script.
8. **Pre-Token sheet migration** — run once against a sheet with no Token column, confirm the migration summary's matched/unmatched counts look right, and confirm no existing Min/Max values changed.
9. **Interrupt a run** (e.g. kill it mid-catalog-load) and confirm the sheet is untouched afterward — this is the regression test for the clear-before-write bug from Stage 7.

If all nine hold, you've rebuilt it correctly.
