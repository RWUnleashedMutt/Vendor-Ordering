# Sync Rules Matrix — Complete Tutorial

## Table of Contents

1. [Overview](#overview)
2. [What This Script Does](#what-this-script-does)
3. [Setup & Prerequisites](#setup--prerequisites)
4. [Configuration](#configuration)
5. [Input Files](#input-files)
6. [DNO/Min/Max Logic](#dnominmax-logic)
7. [How It Works: Step-by-Step](#how-it-works-step-by-step)
8. [Code Breakdown](#code-breakdown)
9. [Running the Script](#running-the-script)
10. [The Sync Workflow](#the-sync-workflow)
11. [Troubleshooting](#troubleshooting)
12. [Common Issues & Solutions](#common-issues--solutions)

---

## Overview

The **Sync Rules Matrix** script is a critical tool in Southeast Pet's inventory automation ecosystem. It synchronizes vendor catalog data with a rules matrix that defines ordering behaviors for each store and SKU.

**Core Purpose:** Keep the rules matrix synchronized with the latest catalog while preserving team edits made in Google Sheets, validating data consistency, and managing store-specific ordering rules.

**Key Responsibilities:**

- Load catalog (raw product list) and rules matrix (ordering rules per store)
- Sync SKU changes (additions, removals, exclusions)
- Validate DNO/Min/Max logic (prevent accidental data loss)
- Push synced data to Google Sheets
- Manage bidirectional consistency between local and cloud data

**Integration Points:**

- **Input:** Square catalog export (Excel)
- **Output:** Google Sheets (synced rules matrix)
- **Related Script:** `apply_changes.py` (pulls team edits from Sheets before sync)

---

## What This Script Does

### Problem It Solves

You have two sources of truth that get out of sync:

1. **Catalog** (Source of Products)
   - Updated frequently (new products added, discontinued products removed)
   - Comes from vendor or Square
   - Contains: SKU, Item Name, Category

2. **Rules Matrix** (Ordering Rules)
   - Defines how each store orders each product
   - Contains: SKU + 3 columns per store (DNO, Min, Max)
   - Team members edit directly in Google Sheets
   - Gets stale when catalog changes

**Manual Sync Problems:**

- Easy to miss new products (gaps in ordering rules)
- Easy to forget to remove discontinued products (clutter, failed orders)
- Team edits can be overwritten if you're not careful
- No validation (you might set DNO=True AND Min=100, which contradicts)

### Solution

This script automates the sync by:

1. **Loading both sources:**
   - Reads the latest catalog (from file)
   - Reads the existing rules matrix (from local Excel)

2. **Handling SKU changes:**
   - **New SKUs:** Adds them to the matrix with default settings
   - **Discontinued SKUs:** Removes them (with confirmation)
   - **Excluded SKUs:** Respects the "Excluded SKUs" sheet in Google Sheets

3. **Validating logic:**
   - Checks for inconsistencies (DNO=True but Min/Max are non-zero)
   - Alerts you before data loss occurs
   - Offers a chance to fix issues

4. **Applying bidirectional rules:**
   - **DNO → Zeroing:** If DNO=True, sets Min/Max to 0 (clean state)
   - **Zero → DNO:** If Min=0 AND Max=0, sets DNO=True (consistent)

5. **Pushing to Google Sheets:**
   - Overwrites the remote matrix with the synced local version
   - Team has a fresh starting point for their edits

### Workflow Integration

This script is part of a two-step process:

```
Team edits rules matrix in Google Sheets
         ↓
Run apply_changes.py  ← Pulls team edits locally
         ↓
New catalog arrives
         ↓
Run sync_rules_matrix.py  ← Syncs new SKUs, validates, pushes to Sheets
         ↓
Team reviews updated matrix, makes new edits
```

---

## Setup & Prerequisites

### 1. Python & Dependencies

```bash
pip install pandas openpyxl gspread google-auth-oauthlib google-auth-httplib2 google-cloud-storage
```

Or simpler (usually already installed in your environment):

```bash
pip install pandas openpyxl gspread google-auth
```

### 2. Google Sheets Authentication

This script needs credentials to read/write Google Sheets. Set up a service account:

#### Step 1: Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (e.g., "Southeast Pet Automation")
3. Enable two APIs:
   - Google Sheets API
   - Google Drive API

#### Step 2: Create a Service Account

1. Navigate to **IAM & Admin** → **Service Accounts**
2. Click **Create Service Account**
3. Name it: "southeast-pet-automation"
4. Click **Create and Continue**
5. Grant it **Editor** role (for Sheets and Drive)
6. Click **Continue** and **Done**

#### Step 3: Create a Key

1. Click on the service account you just created
2. Go to the **Keys** tab
3. Click **Add Key** → **Create new key**
4. Choose **JSON**
5. A `credentials.json` file will download
6. Place it in the same directory as the script (or update `CREDENTIALS_FILE` path)

**Security:** Treat `credentials.json` like a password. Never commit it to version control.

#### Step 4: Share Google Sheets with Service Account

1. For each vendor's Rules Matrix sheet in Google Drive:
   - Click **Share**
   - Add the service account email: `southeast-pet-automation@YOUR_PROJECT.iam.gserviceaccount.com`
   - Grant **Editor** access

### 3. Secrets Configuration

The script reads vendor → Sheet ID mappings from `.streamlit/secrets.toml`:

**File:** `.streamlit/secrets.toml`

```toml
[sheet_ids]
Southeast = "1a2b3c4d5e6f7g8h9i0j"
Sysco = "2x3y4z5a6b7c8d9e0f1g"
US Foods = "3p3q4r5s6t7u8v9w0x1y"
# Add more as needed
```

**How to get a Sheet ID:**

1. Open the Google Sheet in your browser
2. The URL is: `https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit`
3. Copy the `SHEET_ID` part
4. Add it to `secrets.toml` under `[sheet_ids]`

### 4. Directory Structure

```
project_root/
├── sync_rules_matrix.py          (the script)
├── apply_changes.py              (companion script)
├── credentials.json              (Google service account key)
├── .streamlit/
│   └── secrets.toml              (vendor Sheet IDs)
│
├── Data/
│   └── Rules/
│       ├── Southeast Rules Matrix.xlsx
│       ├── Sysco Rules Matrix.xlsx
│       └── ...
│
└── (Catalog files selected via file dialog)
```

---

## Configuration

### Vendor Name

When you run the script, it prompts for a vendor name:

```
Input Vendor: (Ex: Southeast):
```

**This name is used to:**

1. Look up the Sheet ID in `secrets.toml`
2. Name the local rules matrix file: `Data/Rules/{Vendor} Rules Matrix.xlsx`

**Must be exact match** (case-sensitive) of the key in `secrets.toml`

### Store Mapping

The `store_map` dictionary maps Square's column names to your store codes:

```python
store_map = {
    'Current Quantity City Market: DTR': 'CM',
    'Current Quantity Crabtree Valley Mall': 'CVM',
    'Current Quantity Crescent Commons': 'CC',
    # ... more stores
}
```

**Why this mapping?**

- Square exports have columns like "Current Quantity City Market: DTR"
- You need short codes (CM, CVM, CC) for the rules matrix
- This dict translates between them

**If a store is missing:**

1. Add it to `store_map`
2. The script will auto-create DNO, Min, Max columns for it

**Syntax:**

- **Key:** Must match the Square export column name exactly (including "Current Quantity" prefix)
- **Value:** 2–3 letter store code

### Column Defaults

These constants define what values are used for new SKUs:

```python
DEFAULT_MIN = 0        # Minimum stock level for new SKUs
DEFAULT_MAX = 0        # Maximum stock level for new SKUs
DEFAULT_ORDER_QTY = 1  # Order in quantities (bulk size)
```

**Example:** When a new SKU is added, each store gets:

- `DNO = True` (Do Not Order, until manually set)
- `Min = 0`
- `Max = 0`
- `Order In Quantities = 1`

**To change defaults** (e.g., new SKUs default to Min=10, Max=50):

```python
DEFAULT_MIN = 10
DEFAULT_MAX = 50
```

### Catalog Column Requirements

The script expects the catalog Excel file to have these columns:

```python
REQUIRED_CATALOG_COLS = {'SKU', 'Item Name', 'Reporting Category'}
```

**Column Details:**

- **SKU:** Product code (unique identifier)
- **Item Name:** Product description
- **Reporting Category:** Brand or product category

**Header Row Note:**

```python
catalog = pd.read_excel(path, header=1, ...)
```

The script assumes there's a title/blank row before the headers. If your catalog headers are on row 1, change to `header=0`.

---

## Input Files

### File 1: Catalog (Excel)

**Format:** Standard Excel file with product list

**Required Columns:**
| SKU | Item Name | Reporting Category | (other columns ignored) |
|-----|-----------|-------------------|------------------------|
| 001234 | Dog Food 25lb | Natures Logic | ... |
| 001235 | Cat Litter 40lb | Precious Cat | ... |
| 001236 | Dog Treat Packs | Wellness CORE | ... |

**Notes:**

- **SKU:** Must be unique (duplicates are auto-removed)
- **Null SKUs:** Rows with blank SKU are removed
- **Extra columns:** Ignored (safe to export)
- **Header row:** Assumes there's a blank row before headers (adjust `header=1` if needed)

**Example sources:**

- Square catalog export
- Vendor product list
- Internal inventory catalog

### File 2: Rules Matrix (Excel, Local)

**Location:** `Data/Rules/{Vendor} Rules Matrix.xlsx`

**Format:** Created automatically if it doesn't exist; synced each run

**Columns:**
| SKU | Item Name | Reporting Category | Order In Quantities | CM_DNO | CM_Min | CM_Max | CVM_DNO | CVM_Min | CVM_Max | ... |
|-----|-----------|-------------------|---------------------|--------|--------|--------|---------|---------|---------|-----|
| 001234 | Dog Food 25lb | Natures Logic | 1 | False | 10 | 50 | True | 0 | 0 | ... |

**Columns per store:** For each store code (CM, CVM, CC, etc.):

- `{CODE}_DNO` – Do Not Order flag (True/False)
- `{CODE}_Min` – Minimum stock level (number)
- `{CODE}_Max` – Maximum stock level (number)

### File 3: Excluded SKUs (Google Sheet)

**Location:** Within the vendor's Rules Matrix workbook on Google Sheets

**Sheet Name:** "Excluded SKUs" (must be exact)

**Format:** Single column (A) with SKUs to exclude

| SKU    |
| ------ |
| 000001 |
| 000002 |
| 000999 |

**Purpose:** SKUs in this list are:

- Removed from the catalog before syncing
- Removed from the rules matrix if they exist
- Never synced again (until removed from this list)

**Example uses:**

- Vendor products you don't stock
- Discontinued items not yet removed from vendor catalog
- Products that cause ordering issues

---

## DNO/Min/Max Logic

This is the most important part to understand. The script enforces bidirectional consistency between three related fields.

### The Three Fields

For each store + SKU combination:

| Field   | Meaning       | Values         |
| ------- | ------------- | -------------- |
| **DNO** | Do Not Order  | True or False  |
| **Min** | Minimum stock | 0–999 (number) |
| **Max** | Maximum stock | 0–999 (number) |

### The Rules

**Rule 1: DNO=True implies Min/Max = 0**

- If `DNO = True`, the store doesn't order this SKU
- Therefore, Min and Max should be 0 (no stocking rules apply)
- Script enforces: `DNO=True` → `Min=0, Max=0`

**Rule 2: Min/Max=0 implies DNO=True**

- If both `Min = 0` and `Max = 0`, the store doesn't need a stocking rule
- Therefore, `DNO` should be True (explicit "do not order")
- Script enforces: `Min=0, Max=0` → `DNO=True`

### Why Both Rules?

They keep the matrix clean and prevent confusion:

**Without Rule 1 (DNO zeroing):**

```
Bad state: DNO=True but Min=10, Max=50
↳ Contradictory: "Don't order" but also "keep 10–50 units"
↳ If someone later forgets DNO is True, ordering would happen incorrectly
```

**Without Rule 2 (zero-to-DNO):**

```
Bad state: DNO=False but Min=0, Max=0
↳ Contradictory: "Do order" but "no stocking level"
↳ Unclear: should the system try to order this or not?
```

### Validation Before Sync

Before applying these rules, the script **validates** for inconsistencies:

```
⚠️  WARNING: DNO/Min/Max Inconsistencies Detected!

The following items have DNO=True BUT non-zero Min/Max values.
This likely means you forgot to set DNO=False after setting Min/Max.
If you proceed, these Min/Max values will be ERASED and set to 0.

  SKU: 001234         Store: CM      DNO=True, Min=10, Max=50
  SKU: 001235         Store: CVM     DNO=True, Min=5, Max=25
```

**At this point, you can:**

- **Option 1:** `yes` – Fix it automatically (set DNO=False, keep Min/Max)
- **Option 2:** `no` – Proceed with sync (erase Min/Max to 0)

### Example: Setting Up a Store Rule

**You want to add ordering for "Dog Food 25lb" (SKU 001234) at the CM store:**

**Step 1: Set Min and Max**

- `CM_Min = 10` (keep at least 10 units)
- `CM_Max = 50` (don't order more than 50 units)

**Step 2: Set DNO**

- `CM_DNO = False` (actively order this product)

**After sync:**

- Validation detects: DNO=False, Min=10, Max=50 ✓ (consistent)
- No changes needed
- Sync proceeds normally

---

## How It Works: Step-by-Step

### High-Level Flow

```
User inputs vendor name
         ↓
Warn: Did you run apply_changes.py first?
         ↓
User selects catalog file
         ↓
Load catalog + rules matrix + excluded SKUs
         ↓
Sync SKUs:
  ├── Add new SKUs from catalog
  ├── Remove discontinued SKUs from matrix
  └── Remove excluded SKUs
         ↓
Validate DNO/Min/Max logic
  └── Alert if inconsistencies found
         ↓
Apply bidirectional rules:
  ├── DNO=True → Min/Max=0
  └── Min=0, Max=0 → DNO=True
         ↓
Save locally + push to Google Sheets
         ↓
Print success message with Sheet URL
```

### Detailed Step-by-Step

#### Step 1: User Input & Validation

```python
Vendor = sanitize_vendor_name(input('Input Vendor: (Ex: Southeast): '))
```

- Prompts for vendor name
- `sanitize_vendor_name()` removes invalid characters: `< > : " / \ | ? *`
- Used later to look up Sheet ID and name files

#### Step 2: Safety Check

```python
print("Have you already run apply_changes.py to pull team edits? (yes/no): ")
confirm = input().strip().lower()
if confirm != 'yes':
    sys.exit(0)
```

**Why this check?**

- If team members have edited the rules matrix in Google Sheets
- And you run sync without pulling those edits first
- Those changes will be overwritten by the sync
- This warning prevents accidental data loss

**Correct workflow:**

1. Team edits Google Sheets → `apply_changes.py` (pull edits locally)
2. New catalog arrives → `sync_rules_matrix.py` (sync + push back to Sheets)

#### Step 3: File Selection

```python
CATALOG_PATH = get_file_path(title=f"Select Catalog for {Vendor}")
```

- Opens file dialog (GUI)
- User browses to the catalog Excel file
- Path is stored in `CATALOG_PATH`

#### Step 4: Load Google Sheets

```python
client = get_google_client()
spreadsheet = get_or_create_sheet(client, vendor)
```

- Authenticates using `credentials.json`
- Looks up vendor's Sheet ID in `secrets.toml`
- Opens the Google Sheet

#### Step 5: Load Excluded SKUs

```python
excluded_skus = load_excluded_skus(spreadsheet)
```

- Reads the "Excluded SKUs" sheet from Google Sheets
- Returns a set of SKU strings to exclude
- If sheet doesn't exist, returns empty set with a warning

#### Step 6: Load Catalog & Filter

```python
catalog = load_catalog(CATALOG_PATH)
catalog = catalog[~catalog['SKU'].isin(excluded_skus)]
```

- Reads the Excel file
- Validates required columns exist
- Removes duplicates (keeps first occurrence)
- Filters out excluded SKUs

#### Step 7: Load Rules Matrix

```python
rules_df = load_or_create_matrix(matrix_path, catalog)
```

**If matrix exists:**

- Load from Excel
- Merge with current catalog (updates Item Name, Reporting Category)
- Add missing store columns

**If matrix doesn't exist:**

- Create from scratch using the catalog
- Initialize all store columns with defaults

#### Step 8: Remove Excluded SKUs from Matrix

```python
rules_df = remove_excluded_skus(rules_df, excluded_skus)
```

- Filters out any SKUs that are on the exclusion list
- Prints which SKUs are being removed

#### Step 9: Remove Discontinued SKUs

```python
rules_df = remove_discontinued_skus(rules_df, catalog)
```

**What it does:**

1. Compare SKUs in matrix vs. catalog
2. If a SKU is in matrix but NOT in catalog, it's discontinued
3. Alert user with a list:
   ```
   Removing 3 discontinued SKU(s):
     - 001234
     - 001235
     - 001236
   ```
4. Prompt for confirmation: "Remove these? (yes/no)"
5. If yes, remove them; if no, keep them

**Why confirm?**

- You might want to keep a discontinued SKU for reference
- Or the catalog might be incomplete/wrong
- User has final say

#### Step 10: Ensure Store Columns

```python
rules_df = ensure_store_columns(rules_df)
```

- For each store in `store_map`:
  - If `{CODE}_DNO` column missing, add it (default: False)
  - If `{CODE}_Min` column missing, add it (default: 0)
  - If `{CODE}_Max` column missing, add it (default: 0)

**When needed:**

- New stores added to `store_map` that weren't in the previous matrix
- Existing stores but missing some columns due to data corruption

#### Step 11: Add New SKUs

```python
rules_df = append_new_skus(rules_df, catalog)
```

**What it does:**

1. Compare SKUs: catalog vs. existing matrix
2. Find SKUs in catalog but NOT in matrix = new items
3. For each new SKU:
   - Add a row to the matrix
   - Set `Order In Quantities = DEFAULT_ORDER_QTY`
   - For each store:
     - Set `DNO = True` (do not order by default)
     - Set `Min = DEFAULT_MIN` (0)
     - Set `Max = DEFAULT_MAX` (0)

**Result:** New products are in the matrix but flagged as "do not order" until managers explicitly set Min/Max and DNO=False

#### Step 12: Validate DNO/Min/Max

```python
rules_df = validate_dno_consistency(rules_df)
```

**This function:**

1. Finds all rows where `DNO=True` BUT `Min ≠ 0` OR `Max ≠ 0`
2. Alerts user with a detailed warning
3. Offers choice:
   - **"yes"** – Fix automatically (set DNO=False, keep Min/Max)
   - **"no"** – Proceed (erase Min/Max to 0)

**Why this check?**

- Prevents accidental data loss
- Catches cases where users set Min/Max but forgot to change DNO

#### Step 13: Apply DNO Zeroing

```python
rules_df = apply_dno_zeroing(rules_df)
```

**Rule: If DNO=True, then Min=0 and Max=0**

- For each store and SKU:
  - If `DNO = True`
  - Then set `Min = 0` and `Max = 0`
- Reports how many values were zeroed

#### Step 14: Apply Zero-to-DNO

```python
rules_df = apply_zero_to_dno(rules_df)
```

**Rule: If Min=0 and Max=0, then DNO=True**

- For each store and SKU:
  - If `Min = 0` AND `Max = 0`
  - AND `DNO ≠ True`
  - Then set `DNO = True`
- Reports how many DNO flags were set

#### Step 15: Select Columns & Order

```python
matrix_columns = ['SKU', 'Item Name', 'Reporting Category', 'Order In Quantities']
for code in store_map.values():
    matrix_columns.extend([f'{code}_DNO', f'{code}_Min', f'{code}_Max'])
final_cols = [c for c in matrix_columns if c in rules_df.columns]
rules_df = rules_df[final_cols]
```

- Reorders columns in a logical sequence
- Ensures consistent output format across runs
- Removes any unexpected columns

#### Step 16: Save Locally

```python
rules_df.to_excel(matrix_path, index=False)
```

- Writes synced matrix to local Excel file
- Location: `Data/Rules/{Vendor} Rules Matrix.xlsx`

#### Step 17: Push to Google Sheets

```python
push_to_sheets(spreadsheet, rules_df)
```

- Clears the Google Sheet
- Uploads the synced matrix
- Prints the Sheet URL for easy access

---

## Code Breakdown

### Imports & Constants

```python
import pandas as pd
import tomllib
import os
import sys
import re
import tkinter as tk
from tkinter import filedialog
import gspread
from google.oauth2.service_account import Credentials
```

**pandas:** Data manipulation
**tomllib:** Read `.toml` config files (Python 3.11+)
**os, sys:** File/system operations
**re:** Regex for sanitizing vendor names
**tkinter:** GUI file dialog
**gspread:** Google Sheets API client
**google.oauth2:** Service account authentication

### Function: `get_file_path()`

```python
def get_file_path(title="Select File", file_types=(...)):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(...)
    root.destroy()
    return file_path
```

- Opens a file browser dialog
- `withdraw()` hides the empty Tk window
- `-topmost` keeps dialog on top
- Returns the selected file path or empty string if cancelled

### Function: `sanitize_vendor_name()`

```python
def sanitize_vendor_name(name):
    return re.sub(r'[<>:"/\\|?*]', '', name).strip()
```

- Removes invalid filename characters
- Reason: Vendor name is used in file paths and Sheet lookups
- Prevents errors from special characters

### Function: `get_google_client()`

```python
def get_google_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)
```

- Loads `credentials.json` (service account key)
- Authorizes with the two scopes (Sheets + Drive)
- Returns a gspread client ready to use

### Function: `get_or_create_sheet()`

```python
def get_or_create_sheet(client, vendor):
    if vendor not in SHEET_IDS:
        print(f"Error: No Sheet ID found for vendor '{vendor}'.")
        sys.exit(1)
    spreadsheet = client.open_by_key(SHEET_IDS[vendor])
    return spreadsheet
```

- Looks up vendor's Sheet ID in config
- Opens the Google Sheet
- Exits if vendor not configured

### Function: `load_excluded_skus()`

```python
def load_excluded_skus(spreadsheet):
    try:
        worksheet = spreadsheet.worksheet("Excluded SKUs")
        skus = worksheet.col_values(1)
        excluded = {sku.strip() for sku in skus[1:] if sku.strip()}
        return excluded
    except Exception as e:
        print(f"⚠ Warning: Could not load 'Excluded SKUs' sheet: {e}")
        return set()
```

- Reads the "Excluded SKUs" sheet from Google Sheets
- Column A, starting from row 2 (skipping header)
- Returns a set of SKUs to exclude
- If sheet doesn't exist, warns and returns empty set

### Function: `push_to_sheets()`

```python
def push_to_sheets(spreadsheet, rules_df):
    worksheet = spreadsheet.sheet1
    worksheet.clear()
    rules_df = rules_df.fillna('')
    data = [rules_df.columns.tolist()] + rules_df.values.tolist()
    worksheet.update(data)
```

- Gets the first sheet
- Clears all existing data
- Converts DataFrame to list of lists (rows)
- Uploads to Google Sheets
- `fillna('')` replaces NaN with empty strings (Sheets compatibility)

### Function: `load_catalog()`

```python
def load_catalog(path):
    catalog = pd.read_excel(path, header=1, usecols=list(REQUIRED_CATALOG_COLS), ...)
    catalog = catalog.dropna(subset=['SKU']).drop_duplicates(subset=['SKU'])
    return catalog
```

- `header=1` – Skip first row (assumes title row exists)
- `usecols` – Only read required columns
- `dtype={'SKU': str, ...}` – Preserve SKU as text
- `dropna(subset=['SKU'])` – Remove rows with missing SKU
- `drop_duplicates(subset=['SKU'])` – Keep first occurrence

### Function: `load_or_create_matrix()`

```python
def load_or_create_matrix(path, catalog):
    if os.path.exists(path):
        rules_df = pd.read_excel(path, dtype={'SKU': str})
        # Merge with catalog to update Item Name, Reporting Category
        rules_df = pd.merge(rules_df, catalog[['SKU', 'Item Name', ...]], on='SKU', how='left')
    else:
        rules_df = catalog.copy()
        rules_df['Order In Quantities'] = DEFAULT_ORDER_QTY
    return rules_df
```

**If matrix exists:**

- Load it
- Handle legacy column names (Order_Qty → Order In Quantities)
- Merge with latest catalog data (updates product info)
- Add missing Order In Quantities column if needed

**If matrix doesn't exist:**

- Create from scratch using catalog
- Initialize Order In Quantities to default

**Why merge with catalog?**

- Catalog might have updated Item Names or Categories
- Ensures rules matrix always has current product info
- Avoids stale data in the matrix

### Function: `remove_excluded_skus()`

```python
def remove_excluded_skus(rules_df, excluded_set):
    excluded = rules_df[rules_df['SKU'].isin(excluded_set)]
    if not excluded.empty:
        print(f"Removing {len(excluded)} excluded SKU(s):")
        for sku in excluded['SKU'].tolist():
            print(f"  - {sku}")
        rules_df = rules_df[~rules_df['SKU'].isin(excluded_set)].reset_index(drop=True)
    return rules_df
```

- Filter matrix to find rows with SKUs in the exclusion list
- Print which ones are being removed
- Remove them from the matrix
- Reset index (keeps row numbers clean)

### Function: `remove_discontinued_skus()`

```python
def remove_discontinued_skus(rules_df, catalog):
    catalog_skus = set(catalog['SKU'].tolist())
    discontinued = rules_df[~rules_df['SKU'].isin(catalog_skus)]
    if not discontinued.empty:
        print(f"Removing {len(discontinued)} discontinued SKU(s):")
        for sku in discontinued['SKU'].tolist():
            print(f"  - {sku}")
        confirm = input("Confirm removal of discontinued SKUs? (yes/no): ")
        if confirm != 'yes':
            return rules_df
        rules_df = rules_df[rules_df['SKU'].isin(catalog_skus)].reset_index(drop=True)
    return rules_df
```

- Compare matrix SKUs to catalog SKUs
- If a SKU is in matrix but not in catalog, it's discontinued
- Alert user with list
- Ask for confirmation before removing
- Remove only if confirmed

**Why confirm?**

- Discontinued removal is permanent (unless you undo manually)
- Catalog might be incomplete/wrong
- User should have final say

### Function: `ensure_store_columns()`

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
```

- For each store code (CM, CVM, etc.):
  - Check if DNO, Min, Max columns exist
  - If not, create them with default values
- Used when new stores are added to `store_map`

### Function: `append_new_skus()`

```python
def append_new_skus(rules_df, catalog):
    existing_skus = set(rules_df['SKU'].tolist())
    new_items = catalog[~catalog['SKU'].isin(existing_skus)].copy()

    if not new_items.empty:
        new_items['Order In Quantities'] = DEFAULT_ORDER_QTY
        for code in store_map.values():
            new_items[f'{code}_DNO'] = True
            new_items[f'{code}_Min'] = DEFAULT_MIN
            new_items[f'{code}_Max'] = DEFAULT_MAX
        rules_df = pd.concat([rules_df, new_items], ignore_index=True)
        print(f"Added {len(new_items)} new SKU(s) to the matrix.")
    return rules_df
```

- Find SKUs in catalog but not in matrix = new products
- For each new SKU:
  - Initialize all store columns
  - Set DNO=True (don't order by default)
  - Set Min=0, Max=0 (no stocking level)
- Concatenate to existing matrix

### Function: `validate_dno_consistency()`

```python
def validate_dno_consistency(rules_df):
    inconsistencies = []
    for code in store_map.values():
        dno_mask = rules_df[dno_col].apply(lambda v: str(v).upper() == 'TRUE' or v is True)
        nonzero_mask = (rules_df[min_col] != 0) | (rules_df[max_col] != 0)
        issue_mask = dno_mask & nonzero_mask

        if issue_mask.any():
            for idx in rules_df[issue_mask].index:
                inconsistencies.append({...})

    if inconsistencies:
        print("⚠️  WARNING: DNO/Min/Max Inconsistencies Detected!")
        for item in inconsistencies:
            print(f"  SKU: {item['SKU']:<15} Store: {item['Store']:<5} ...")

        confirm = input("Do you want to KEEP these Min/Max values and set DNO=False? (yes/no): ")
        if confirm == 'yes':
            # Fix by setting DNO=False
            for item in inconsistencies:
                rules_df.loc[rules_df['SKU'] == sku, dno_col] = False

    return rules_df
```

**Purpose:** Detect cases where DNO=True but Min/Max are non-zero

**What it does:**

1. For each store + SKU combination:
   - Check if DNO=True AND (Min≠0 OR Max≠0)
2. If found, alert user with details
3. Offer choice:
   - **"yes"** – User wants to keep Min/Max, so set DNO=False
   - **"no"** – Proceed with sync (will zero the Min/Max in next step)

**Why this function?**

- Prevents accidental data loss from inconsistent states
- Gives user a chance to fix before irreversible changes

### Function: `apply_dno_zeroing()`

```python
def apply_dno_zeroing(rules_df):
    zeroed_count = 0
    for code in store_map.values():
        dno_mask = rules_df[dno_col].apply(lambda v: str(v).upper() == 'TRUE' or v is True)
        if dno_mask.any():
            if min_col in rules_df.columns:
                zeroed_count += (rules_df.loc[dno_mask, min_col] != 0).sum()
                rules_df.loc[dno_mask, min_col] = 0
            if max_col in rules_df.columns:
                zeroed_count += (rules_df.loc[dno_mask, max_col] != 0).sum()
                rules_df.loc[dno_mask, max_col] = 0
    return rules_df
```

**Rule: DNO=True → Min=0, Max=0**

- For each store + SKU:
  - If DNO=True, set Min=0 and Max=0
- Counts how many values were changed
- Prints summary

### Function: `apply_zero_to_dno()`

```python
def apply_zero_to_dno(rules_df):
    dno_count = 0
    for code in store_map.values():
        zero_mask = (rules_df[min_col] == 0) & (rules_df[max_col] == 0)
        to_set = zero_mask & ~(rules_df[dno_col].apply(...))
        dno_count += to_set.sum()
        rules_df.loc[to_set, dno_col] = True
    return rules_df
```

**Rule: Min=0 AND Max=0 → DNO=True**

- For each store + SKU:
  - If Min=0 AND Max=0 AND DNO≠True
  - Then set DNO=True
- Counts how many DNO flags were set
- Prints summary

### Function: `sync_rules_matrix()`

```python
def sync_rules_matrix(vendor, catalog_path, matrix_path):
    # Connect to Google Sheets
    client = get_google_client()
    spreadsheet = get_or_create_sheet(client, vendor)

    # Load data
    excluded_skus = load_excluded_skus(spreadsheet)
    catalog = load_catalog(catalog_path)
    catalog = catalog[~catalog['SKU'].isin(excluded_skus)]
    rules_df = load_or_create_matrix(matrix_path, catalog)

    # Sync & validate
    rules_df = remove_excluded_skus(rules_df, excluded_skus)
    rules_df = remove_discontinued_skus(rules_df, catalog)
    rules_df = ensure_store_columns(rules_df)
    rules_df = append_new_skus(rules_df, catalog)
    rules_df = validate_dno_consistency(rules_df)

    # Apply rules
    rules_df = apply_dno_zeroing(rules_df)
    rules_df = apply_zero_to_dno(rules_df)

    # Format & save
    rules_df = rules_df[final_cols]
    rules_df.to_excel(matrix_path, index=False)
    push_to_sheets(spreadsheet, rules_df)

    print("Success! Matrix synced locally and to Google Sheets.")
```

The orchestrator that ties everything together. Follows the workflow described in "How It Works."

---

## Running the Script

### Checklist

- [ ] Python 3.7+ installed
- [ ] Dependencies installed: `pip install pandas openpyxl gspread google-auth`
- [ ] `credentials.json` in script directory (Google service account key)
- [ ] `.streamlit/secrets.toml` configured with vendor Sheet IDs
- [ ] Catalog Excel file ready
- [ ] `apply_changes.py` has been run recently (to pull team edits)

### Step 1: Prepare

1. **Have team pull their edits:**

   ```bash
   python apply_changes.py
   ```

   (This brings Google Sheets edits into the local matrix file.)

2. **Obtain new catalog:**
   - Export from vendor or Square
   - Save as Excel file (e.g., `Sysco_Catalog_2024.xlsx`)

3. **Update Excluded SKUs** (if needed):
   - Open the vendor's Rules Matrix sheet in Google Sheets
   - Add/remove SKUs from the "Excluded SKUs" sheet

### Step 2: Run the Script

**From terminal:**

```bash
python sync_rules_matrix.py
```

**From IDE:**

- Open the file
- Press Run button

### Step 3: Follow Prompts

**Prompt 1: Safety check**

```
Have you already run apply_changes.py to pull team edits? (yes/no):
```

- Type `yes` if you've pulled team edits
- Type `no` to exit and run `apply_changes.py` first

**Prompt 2: File selection**

```
Please select the Catalog Excel file...
```

- File browser opens
- Navigate to and select your catalog file
- Click "Open"

**Prompt 3: Confirm discontinued removals** (if any)

```
Removing 2 discontinued SKU(s):
  - 001234
  - 001235

Confirm removal of discontinued SKUs? (yes/no):
```

- Type `yes` to remove them
- Type `no` to keep them

**Prompt 4: Validate DNO/Min/Max** (if inconsistencies found)

```
⚠️  WARNING: DNO/Min/Max Inconsistencies Detected!
...
Do you want to KEEP these Min/Max values and set DNO=False? (yes/no):
```

- Type `yes` to auto-fix (set DNO=False, keep Min/Max)
- Type `no` to proceed with zeroing Min/Max

### Step 4: Monitor Output

Example console output:

```
Connecting to Google Sheets...
Connected to Google Sheet for Sysco

✓ Loaded 15 excluded SKUs from 'Excluded SKUs' sheet

No new SKUs found.
No discontinued SKUs found.
No excluded SKUs found in matrix.

DNO zeroing: no Min/Max values needed adjustment.
Zero-to-DNO: no DNO flags needed adjustment.

Local matrix saved at ./Data/Rules/Sysco Rules Matrix.xlsx
Pushed 245 SKUs to Google Sheets.
Sheet URL: https://docs.google.com/spreadsheets/d/1a2b3c4d5e6f/edit

Success! Matrix synced locally and to Google Sheets.
```

### Step 5: Verify in Google Sheets

1. Open the vendor's Rules Matrix in Google Sheets
2. Verify:
   - New SKUs are present ✓
   - Discontinued SKUs are removed ✓
   - All store columns are populated ✓
   - Data looks correct ✓

---

## The Sync Workflow

This section explains how `sync_rules_matrix.py` fits into the larger team workflow.

### Timeline Example

**Monday Morning:**

- Sysco catalog is updated
- New file: `Sysco_Catalog_2024.xlsx`
- You need to sync this into the system

**Step 1: Alert the team**

```
Hey team, I'm about to sync the new Sysco catalog.
If you have pending edits in Google Sheets, please tell me now.
```

**Step 2: Run apply_changes.py** (Pull team edits locally)

```bash
python apply_changes.py
```

- Reads the current Google Sheet
- Saves to local file: `Data/Rules/Sysco Rules Matrix.xlsx`

**Step 3: Run sync_rules_matrix.py** (Sync + Push)

```bash
python sync_rules_matrix.py
```

- Loads the catalog
- Syncs with the local matrix
- Pushes updated matrix back to Google Sheets

**Step 4: Team reviews & edits**

- Team opens Google Sheets
- Reviews new SKUs, updated product info
- Makes any necessary edits (set Min/Max, DNO flags, etc.)

**Step 5: Repeat cycle** (next time catalog updates)

- Manager runs `apply_changes.py` again (pulls latest team edits)
- Manager runs `sync_rules_matrix.py` again (syncs new catalog, pushes back)

### Why Two Scripts?

**`apply_changes.py`** (Pull):

- Reads from Google Sheets
- Writes to local file
- Purpose: Preserve team edits before pushing new data

**`sync_rules_matrix.py`** (Sync + Push):

- Reads local files (catalog + matrix)
- Validates & syncs
- Writes to Google Sheets
- Purpose: Update Sheets with new catalog data

**Why not combine them?**

- Allows more control: manually trigger pull vs. push
- Safer: you can review local changes before pushing
- Flexible: different sync frequencies (pull often, push carefully)

---

## Troubleshooting

### Issue: "Error: No Sheet ID found for vendor '{vendor}'"

**Cause:** Vendor name not in `secrets.toml`

**Solution:**

1. Open `.streamlit/secrets.toml`
2. Add the vendor:
   ```toml
   [sheet_ids]
   Vendor_Name = "1a2b3c4d5e6f7g8h9i0j"
   ```
3. Get the Sheet ID from the Google Sheets URL
4. Rerun the script with exact vendor name

### Issue: "No module named 'gspread'"

**Cause:** Dependencies not installed

**Solution:**

```bash
pip install gspread google-auth
```

### Issue: "Please select the Catalog Excel file... No file selected."

**Cause:** User cancelled the file dialog

**Solution:**

- Run the script again
- Select a catalog file (don't cancel)

### Issue: "Error: Catalog is missing required columns: {'SKU', 'Item Name'}"

**Cause:** Catalog file doesn't have required columns

**Solution:**

1. Open the catalog in Excel
2. Verify it has columns: SKU, Item Name, Reporting Category
3. If columns are named differently, either:
   - Rename them to match, OR
   - Update `REQUIRED_CATALOG_COLS` in the script

### Issue: "PermissionError: [Errno 13] Permission denied: './Data/Rules/...xlsx'"

**Cause:** File is open in Excel and locked by Windows

**Solution:**

1. Close the matrix file in Excel
2. Rerun the script

### Issue: "⚠ Warning: Could not load 'Excluded SKUs' sheet"

**Cause:** "Excluded SKUs" sheet doesn't exist in Google Sheets

**Solution:**

1. Open the vendor's Rules Matrix in Google Sheets
2. Create a new sheet called "Excluded SKUs" (exactly)
3. Add SKUs to exclude in column A (starting row 2)
4. Rerun the script

### Issue: Large numbers of SKUs are being marked as discontinued

**Cause:** Catalog file path or content is wrong

**Investigation:**

1. Check you selected the correct catalog file
2. Open the catalog and verify it contains expected products
3. Check if header row matches expected columns
4. If the catalog has headers on row 1 (not row 2), change script:
   ```python
   catalog = pd.read_excel(path, header=0, ...)  # Change header=1 to header=0
   ```

### Issue: DNO/Min/Max validation keeps showing inconsistencies

**Cause:** Team forgot to update DNO when setting Min/Max

**Solution:**

- At the prompt, choose one:
  - **"yes"** – Fix automatically (set DNO=False)
  - **"no"** – Accept the zeroing and manually fix in Sheets later

### Issue: Script seems to hang after "Connecting to Google Sheets..."

**Cause:** Network timeout or authentication issue

**Solution:**

1. Check internet connection
2. Verify `credentials.json` is valid and not expired
3. Regenerate credentials if needed:
   - Go to Google Cloud Console
   - Service Account → Keys → Delete old key
   - Create new key (JSON)
   - Replace `credentials.json`

---

## Common Issues & Solutions

### Issue: Data Loss Scenario

**Situation:** You run `sync_rules_matrix.py` but team members made edits in Google Sheets that you didn't pull first.

**Result:** Their edits are overwritten.

**Prevention:**

1. Always run `apply_changes.py` before `sync_rules_matrix.py`
2. The script warns you: "Have you already run apply_changes.py?"
3. Answer "yes" only if you've actually run it

### Issue: Duplicate SKUs

**Situation:** Catalog has duplicate SKU entries.

**Result:** Only the first is kept (others are dropped).

**Prevention:**

- Have the catalog provider clean duplicates
- Or manually remove duplicates before uploading to the script

### Issue: Inconsistent SKU Format

**Situation:** Some SKUs are "001234" (with leading zeros), others are "1234" (without).

**Result:** They're treated as different SKUs.

**Prevention:**

- Standardize SKU format in the catalog
- The script preserves whatever format is in the file
- If needed, add a preprocessing step to normalize SKU format

### Issue: DNO Flag Not Working

**Situation:** You set DNO=True but the store still tries to order the product.

**Result:** Likely, the DNO column isn't being read by the ordering system.

**Prevention:**

- Verify the ordering system checks the DNO column
- Ensure the column name matches exactly
- Test with a single SKU to debug

### Issue: New SKUs Not Appearing

**Situation:** Catalog has new products, but they're not in the rules matrix after sync.

**Cause:** They might be in the Excluded SKUs list.

**Solution:**

1. Check the "Excluded SKUs" sheet in Google Sheets
2. Remove them from the exclusion list if they should be included
3. Rerun the script

### Issue: Min/Max Values Disappear

**Situation:** You set Min=10, Max=50, but they became 0 after sync.

**Cause:** DNO=True was set (maybe by mistake).

**Prevention:**

- Always set DNO=False when setting Min/Max
- The script validates this, so it should warn you
- If you see the warning, choose "yes" to fix automatically

---

## Summary

The **Sync Rules Matrix** script is a critical integration point between vendor catalogs and Southeast Pet's ordering rules. It:

1. **Keeps data synchronized** – Adds new products, removes discontinued ones
2. **Validates logic** – Ensures DNO/Min/Max are consistent
3. **Preserves team work** – Doesn't overwrite manual edits (if you use `apply_changes.py` first)
4. **Pushes to the cloud** – Shares synced data with the team via Google Sheets

**Proper workflow:**

1. Run `apply_changes.py` (pull team edits)
2. Run `sync_rules_matrix.py` (sync catalog, validate, push)
3. Team reviews & edits in Google Sheets
4. Repeat cycle

By following this process and understanding the DNO/Min/Max logic, you avoid data loss and keep the system in a consistent state.
