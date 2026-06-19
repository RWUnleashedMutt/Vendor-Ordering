# Apply Changes — Complete Tutorial

## Table of Contents

1. [Overview](#overview)
2. [What This Script Does](#what-this-script-does)
3. [Setup & Prerequisites](#setup--prerequisites)
4. [Configuration](#configuration)
5. [How It Works: Step-by-Step](#how-it-works-step-by-step)
6. [Code Breakdown](#code-breakdown)
7. [Running the Script](#running-the-script)
8. [Understanding the Change Log](#understanding-the-change-log)
9. [Integration with Sync](#integration-with-sync)
10. [Troubleshooting](#troubleshooting)
11. [Common Workflows](#common-workflows)

---

## Overview

The **Apply Changes** script is the "pull" half of Southeast Pet's two-script synchronization system. While `sync_rules_matrix.py` pushes new catalog data to Google Sheets, `apply_changes.py` pulls team edits from Google Sheets back to your local system.

**Core Purpose:** Preserve team edits made in Google Sheets by pulling them into the local rules matrix before running a sync operation.

**Key Function:**

- Reads the current Google Sheets version (with team edits)
- Compares it to the local version (original state)
- Detects and logs all changes
- Updates the local file
- Exports a detailed change log

**Why It Exists:**
Prevents data loss. If team members edit the rules matrix in Google Sheets and you immediately run `sync_rules_matrix.py` without pulling first, those edits are overwritten. This script ensures you capture team work before pushing new catalog data.

---

## What This Script Does

### Problem It Solves

**Scenario:**

1. You push an updated rules matrix to Google Sheets
2. Team members make edits (set Min/Max, adjust DNO flags) over several days
3. A new catalog arrives; you want to sync it
4. But if you run `sync_rules_matrix.py` without pulling team edits first → **their work is lost**

### Solution

Before running the sync script:

1. Run `apply_changes.py` – Pulls team edits from Google Sheets
2. Detects and logs what changed
3. Asks for confirmation before applying changes
4. Updates the local matrix with team edits
5. Now you're ready to run `sync_rules_matrix.py` with current data

### Input

**Source:** Google Sheets (vendor's Rules Matrix workbook)

- Contains all team edits
- Current version reflects changes made over time
- May have new/removed SKUs, updated Min/Max/DNO values

### Output

**Files Created:**

1. **Local Matrix (Updated)**
   - Path: `Data/Rules/{Vendor} Rules Matrix.xlsx`
   - Contains: Original SKUs + team edits
   - Used as input to `sync_rules_matrix.py`

2. **Change Log (Report)**
   - Path: `Data/log/{Vendor}_ChangeLog_YYYYMMDD_HHMMSS.xlsx`
   - Format: Multi-sheet workbook with Summary, New SKUs, Removed SKUs, Value Changes
   - Purpose: Audit trail of what changed

### Change Detection

The script detects:

- **New SKUs** – Added to Sheets since last sync
- **Removed SKUs** – Deleted from Sheets
- **Value Changes** – Min, Max, or DNO values changed
- **Special Case:** DNO=True triggers Min/Max zeroing (with logging)

---

## Setup & Prerequisites

### 1. Google Sheets Authentication

This script uses the same authentication as `sync_rules_matrix.py`. If you've already set up that script, you're good to go.

**If not, follow these steps:**

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
4. Grant it **Editor** role
5. Click **Create and Continue** → **Done**

#### Step 3: Create a Key

1. Click on the service account
2. Go to the **Keys** tab
3. Click **Add Key** → **Create new key** → **JSON**
4. A `credentials.json` file downloads
5. Place it in the same directory as the script

#### Step 4: Share Google Sheets with Service Account

1. For each vendor's Rules Matrix sheet:
   - Click **Share**
   - Add: `southeast-pet-automation@YOUR_PROJECT.iam.gserviceaccount.com`
   - Grant **Editor** access

### 2. Python & Dependencies

```bash
pip install pandas openpyxl gspread google-auth
```

### 3. Directory Structure

```
project_root/
├── apply_changes.py              (the script)
├── sync_rules_matrix.py          (companion script)
├── credentials.json              (Google service account key)
│
├── Data/
│   ├── Rules/
│   │   ├── Southeast Rules Matrix.xlsx
│   │   ├── Sysco Rules Matrix.xlsx
│   │   └── ...
│   │
│   └── log/
│       ├── Southeast_ChangeLog_20240115_143022.xlsx
│       └── ...
│
└── .streamlit/
    └── secrets.toml              (vendor Sheet IDs – optional for this script)
```

### 4. File Requirements

**Local Rules Matrix must exist:**

- Path: `Data/Rules/{Vendor} Rules Matrix.xlsx`
- The script uses this as the baseline for comparing changes
- If missing, script exits with error

**Google Sheet must exist:**

- Name: `{Vendor} Rules Matrix`
- Shared with the service account
- Contains team edits to pull

---

## Configuration

### Vendor Name

When you run the script, it prompts:

```
Input Vendor: (Ex: Southeast):
```

**This name is used to:**

1. Find the local matrix file: `Data/Rules/{Vendor} Rules Matrix.xlsx`
2. Find the Google Sheet: `{Vendor} Rules Matrix`
3. Name the change log: `{Vendor}_ChangeLog_TIMESTAMP.xlsx`

**Must be exact match** (case-sensitive) of the Sheet name in Google Drive.

### Credentials File

```python
CREDENTIALS_FILE = './credentials.json'
```

Path to your Google service account key. Change this if you store credentials elsewhere.

### Scopes

```python
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
```

These are fixed and needed for reading Google Sheets and creating change logs.

---

## How It Works: Step-by-Step

### High-Level Flow

```
User inputs vendor name
         ↓
Connect to Google Sheets
         ↓
Pull team edits from Google Sheet
         ↓
Load local baseline matrix
         ↓
Compare: Sheets version vs. Local version
  ├── Detect new SKUs
  ├── Detect removed SKUs
  └── Detect value changes (Min/Max/DNO)
         ↓
Display summary of changes
         ↓
Ask: Apply these changes? (yes/no)
         ↓
If yes:
  ├── Update local matrix with team edits
  ├── Export change log Excel file
  └── Print success message
         ↓
If no:
  └── Exit without applying changes
```

### Detailed Step-by-Step

#### Step 1: User Input

```python
Vendor = sanitize_vendor_name(input('Input Vendor: (Ex: Southeast): '))
```

- Prompts for vendor name
- `sanitize_vendor_name()` removes invalid characters
- Example input: "Sysco" → stored as "Sysco"

#### Step 2: Connect to Google Sheets

```python
client = get_google_client()
sheets_df = pull_from_sheets(client, vendor)
```

- Authenticates using `credentials.json`
- Looks for a sheet named `"{Vendor} Rules Matrix"`
- Reads all data (all rows, all columns)
- Converts to a pandas DataFrame

#### Step 3: Load Local Baseline

```python
local_df = load_local_matrix(matrix_path)
```

- Reads `Data/Rules/{Vendor} Rules Matrix.xlsx`
- This is the last known "synced" state
- Used as baseline for detecting changes

#### Step 4: Compare Versions

```python
updated_df, changes_df, added, removed = compare_and_apply(local_df, sheets_df)
```

**What `compare_and_apply()` does:**

**4a. Compare SKU Sets**

```python
local_skus = set(local_df['SKU'].tolist())
sheets_skus = set(sheets_df['SKU'].tolist())

added = sheets_skus - local_skus      # In Sheets but not Local
removed = local_skus - sheets_skus    # In Local but not Sheets
```

**4b. Detect Value Changes**

```python
common_skus = local_skus & sheets_skus  # SKUs in both versions
```

For each common SKU:

- Compare Min, Max, DNO columns
- Log any differences
- Special handling: If DNO became True, zero out Min/Max

**4c. Return Results**

- `updated_df` – Sheets version with DNO zeroing applied
- `changes_df` – DataFrame of all value changes found
- `added` – Set of new SKU codes
- `removed` – Set of deleted SKU codes

#### Step 5: Display Summary

```python
print(f"\n--- Changes Found ---")
print(f"  New SKUs:        {len(added)}")
print(f"  Removed SKUs:    {len(removed)}")
print(f"  Value Changes:   {len(changes_df)}")
```

**Example output:**

```
--- Changes Found ---
  New SKUs:        3
  Removed SKUs:    1
  Value Changes:   12
```

#### Step 6: Ask for Confirmation

```python
confirm = input("Apply these changes to your local matrix? (yes/no): ").strip().lower()
if confirm != 'yes':
    print("Changes not applied.")
    return
```

- User can review the summary
- User decides whether to apply changes
- If "no", script exits without modifying local file
- If "yes", continues to save

#### Step 7: Save Local Matrix

```python
updated_df.to_excel(matrix_path, index=False)
```

- Overwrites the local rules matrix
- Now includes team edits from Google Sheets
- Ready to be used by `sync_rules_matrix.py`

#### Step 8: Export Change Log

```python
export_change_log(changes_df, added, removed, sheets_df, local_df, vendor)
```

Creates a multi-sheet Excel file with:

- **Summary** – Count of new, removed, and changed items
- **New SKUs** – List of added products
- **Removed SKUs** – List of deleted products
- **Value Changes** – Detailed log of Min/Max/DNO edits

**File location:** `Data/log/{Vendor}_ChangeLog_YYYYMMDD_HHMMSS.xlsx`

**Timestamp example:** `Southeast_ChangeLog_20240115_143022.xlsx`

#### Step 9: Success Message

```python
print("\nChange log saved to: ./Data/log/...")
print("\nDone!")
```

---

## Code Breakdown

### Imports

```python
import pandas as pd
import os
import sys
import re
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
```

**pandas** – Data manipulation
**os, sys** – File/system operations
**re** – Sanitize vendor names
**tkinter** – (Not used in this script; legacy)
**datetime** – Generate timestamps for change logs
**gspread** – Google Sheets API
**google.oauth2** – Service account authentication

### Function: `get_google_client()`

```python
def get_google_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)
```

- Loads credentials from `credentials.json`
- Authorizes with Sheets + Drive scopes
- Returns a gspread client

### Function: `pull_from_sheets()`

```python
def pull_from_sheets(client, vendor):
    sheet_name = f'{vendor} Rules Matrix'
    try:
        spreadsheet = client.open(sheet_name)
        worksheet = spreadsheet.sheet1
        data = worksheet.get_all_records(value_render_option='UNFORMATTED_VALUE')
        df = pd.DataFrame(data)
        df['SKU'] = df['SKU'].astype(str)
        return df
    except gspread.SpreadsheetNotFound:
        print(f"Error: Could not find Google Sheet '{sheet_name}'.")
        sys.exit(1)
```

**What it does:**

1. Constructs sheet name: `"{Vendor} Rules Matrix"`
2. Opens the sheet using gspread client
3. Reads all data as records (rows → dicts → DataFrame)
4. Converts SKU column to string (preserves leading zeros)
5. Returns DataFrame

**Error handling:**

- If sheet not found, prints error and exits

**Note:** `value_render_option='UNFORMATTED_VALUE'` ensures you get actual values, not formulas or formatted text.

### Function: `load_local_matrix()`

```python
def load_local_matrix(path):
    if not os.path.exists(path):
        print(f"Error: Local matrix not found at {path}")
        sys.exit(1)
    df = pd.read_excel(path, dtype={'SKU': str})
    return df
```

- Checks if file exists (exit if not)
- Reads Excel file
- Converts SKU to string
- Returns DataFrame

### Function: `compare_and_apply()`

```python
def compare_and_apply(local_df, sheets_df):
    # ... (detailed below)
```

This is the core logic. It:

1. Compares SKU sets (new, removed)
2. Finds value changes
3. Applies DNO zeroing logic
4. Returns updated DataFrame + change log

**Key sections:**

**SKU Set Comparison:**

```python
local_skus = set(local_df['SKU'].tolist())
sheets_skus = set(sheets_df['SKU'].tolist())

added = sheets_skus - local_skus        # Sheets has, Local doesn't
removed = local_skus - sheets_skus      # Local has, Sheets doesn't
```

**Common SKU Processing:**

```python
common_skus = local_skus & sheets_skus
local_common = local_df[...].set_index('SKU')
sheets_common = sheets_df[...].set_index('SKU')
```

**Value Column Identification:**

```python
value_cols = [c for c in local_df.columns if c.endswith('_Min')
              or c.endswith('_Max') or c.endswith('_DNO')]
dno_cols = [c for c in value_cols if c.endswith('_DNO')]
```

**DNO → Zeroing Logic:**

```python
for dno_col in dno_cols:
    old_dno = local_common.loc[sku, dno_col]
    new_dno = sheets_common.loc[sku, dno_col]
    dno_became_true = (str(old_dno).upper() != 'TRUE') and (str(new_dno).upper() == 'TRUE' or new_dno is True)

    if dno_became_true:
        # Zero out the corresponding Min/Max columns
        # Log the change
```

**Why this logic?**

- If DNO changed from False → True
- Then the corresponding Min/Max should be zeroed (for consistency)
- This change should be logged

**Value Change Detection:**

```python
for col in value_cols:
    old_val = local_common.loc[sku, col]
    new_val = updated_df.loc[sku, col]
    if str(old_val) != str(new_val):
        # Log the change
        change_records.append({...})
```

- Compares each value column
- Logs any differences
- Avoids duplicates (checks if already logged by DNO zeroing)

**Returns:**

```python
return updated_df, changes_df, added, removed
```

### Function: `export_change_log()`

```python
def export_change_log(changes_df, added, removed, sheets_df, local_df, vendor):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f'./Data/log/{vendor}_ChangeLog_{timestamp}.xlsx'
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
```

**Creates directory if missing** – `os.makedirs(..., exist_ok=True)`

**Generates filename with timestamp** – Example: `Southeast_ChangeLog_20240115_143022.xlsx`

**Builds multiple DataFrames:**

**Summary Sheet:**

```python
summary_df = pd.DataFrame({
    'Category': ['New SKUs Added', 'SKUs Removed', 'Value Changes (Min/Max/DNO)'],
    'Count': [len(added), len(removed), len(changes_df)]
})
```

**New SKUs Sheet:**

```python
added_df = sheets_df[sheets_df['SKU'].isin(added)][['SKU', 'Item Name']].copy()
added_df.insert(0, 'Change', 'Added')
```

**Removed SKUs Sheet:**

```python
removed_df = local_df[local_df['SKU'].isin(removed)][['SKU', 'Item Name']].copy()
removed_df.insert(0, 'Change', 'Removed')
```

**Value Changes Sheet:**

- Uses the `changes_df` passed in
- Contains: SKU, Item Name, Store, Field, Old Value, New Value

**Writes to Excel:**

```python
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    # ... more sheets

    for sheet_name in writer.sheets:
        ws = writer.sheets[sheet_name]
        # Auto-fit column widths
```

**Auto-fits column widths** – Makes the report readable

### Function: `apply_changes()`

```python
def apply_changes(vendor):
    try:
        matrix_path = f'./Data/Rules/{vendor} Rules Matrix.xlsx'

        print("\nConnecting to Google Sheets...")
        client = get_google_client()
        sheets_df = pull_from_sheets(client, vendor)
        local_df = load_local_matrix(matrix_path)

        updated_df, changes_df, added, removed = compare_and_apply(local_df, sheets_df)

        if len(changes_df) == 0 and not added and not removed:
            print("\nNo changes detected. Local file unchanged.")
            return

        confirm = input("\nApply these changes to your local matrix? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Changes not applied.")
            return

        updated_df.to_excel(matrix_path, index=False)
        export_change_log(changes_df, added, removed, sheets_df, local_df, vendor)
        print("\nDone!")
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)
```

The orchestrator that ties everything together.

---

## Running the Script

### Checklist

- [ ] Python 3.7+ installed
- [ ] Dependencies: `pip install pandas openpyxl gspread google-auth`
- [ ] `credentials.json` in script directory
- [ ] Local matrix exists at `Data/Rules/{Vendor} Rules Matrix.xlsx`
- [ ] Google Sheet named `{Vendor} Rules Matrix` exists and is shared with service account

### Step 1: Prepare

**Ensure local matrix exists:**

```bash
ls Data/Rules/Southeast\ Rules\ Matrix.xlsx
```

If missing, run `sync_rules_matrix.py` first to create it.

### Step 2: Run the Script

**From terminal:**

```bash
python apply_changes.py
```

**From IDE:**

- Open the file
- Press Run button

### Step 3: Follow Prompts

**Prompt 1: Vendor name**

```
Input Vendor: (Ex: Southeast):
```

- Type exact vendor name (case-sensitive)
- Example: `Sysco` (not `sysco` or `SYSCO`)

### Step 4: Monitor Output

**Example output:**

```
Connecting to Google Sheets...
Pulled 245 SKUs from Google Sheets.
Loaded local matrix: 242 SKUs

--- Changes Found ---
  New SKUs:        3
  Removed SKUs:    0
  Value Changes:   12

Apply these changes to your local matrix? (yes/no):
```

**At this prompt:**

- Review the summary
- Type `yes` to apply changes
- Type `no` to cancel

**If you type `yes`:**

```
Local matrix updated at ./Data/Rules/Southeast Rules Matrix.xlsx

Change log saved to: ./Data/log/Southeast_ChangeLog_20240115_143022.xlsx

Done!
```

---

## Understanding the Change Log

The change log is a multi-sheet Excel file that documents all changes pulled from Google Sheets.

### Sheet 1: Summary

Quick overview of what changed.

| Category                    | Count |
| --------------------------- | ----- |
| New SKUs Added              | 3     |
| SKUs Removed                | 0     |
| Value Changes (Min/Max/DNO) | 12    |

**Use case:** Quickly understand the scope of changes.

### Sheet 2: New SKUs

Products added to Google Sheets since last sync.

| Change | SKU    | Item Name       |
| ------ | ------ | --------------- |
| Added  | 001234 | Dog Food 25lb   |
| Added  | 001235 | Cat Litter 40lb |
| Added  | 001236 | Dog Treats      |

**Use case:** See which products the team added.

### Sheet 3: Removed SKUs

Products deleted from Google Sheets since last sync.

| Change  | SKU    | Item Name         |
| ------- | ------ | ----------------- |
| Removed | 999999 | Discontinued Item |

**Use case:** See which products the team removed.

### Sheet 4: Value Changes

Detailed log of Min, Max, and DNO edits.

| SKU    | Item Name     | Store | Field | Old Value | New Value |
| ------ | ------------- | ----- | ----- | --------- | --------- |
| 001234 | Dog Food 25lb | CM    | Min   | 0         | 10        |
| 001234 | Dog Food 25lb | CM    | Max   | 0         | 50        |
| 001234 | Dog Food 25lb | CM    | DNO   | True      | False     |
| 001234 | Dog Food 25lb | CVM   | Min   | 10        | 15        |
| ...    | ...           | ...   | ...   | ...       | ...       |

**Columns:**

- **SKU** – Product code
- **Item Name** – Product description
- **Store** – Store code (CM, CVM, CC, etc.)
- **Field** – Which column changed (Min, Max, DNO)
- **Old Value** – Previous value (local baseline)
- **New Value** – Current value (Google Sheets)

**Use case:**

- Understand exactly what the team edited
- Audit trail of changes
- Keep for compliance/documentation

---

## Integration with Sync

This script is one half of a two-script system. Understand how they work together.

### Workflow Diagram

```
Team edits in Google Sheets
    ↓
    ← apply_changes.py (Pull edits locally)
    ↓
Local matrix updated with team edits
    ↓
New catalog arrives
    ↓
    ← sync_rules_matrix.py (Sync new SKUs, push to Sheets)
    ↓
Google Sheets updated with new catalog + team edits
    ↓
Team reviews, makes new edits
    ↓
(Repeat cycle)
```

### Correct Sequence

**Always:**

1. Run `apply_changes.py` first ← Pulls team edits
2. Run `sync_rules_matrix.py` second ← Syncs catalog

**Never:**

1. Run `sync_rules_matrix.py` alone without pulling team edits first
   - This overwrites team work

### Example: Monday Morning Workflow

**8:00 AM – Team starts working**

- Team opens Google Sheets
- Makes edits to Min/Max/DNO for various SKUs
- Adds notes in comments

**12:00 PM – New catalog arrives**

- Vendor emails updated catalog

**1:00 PM – You need to sync**

- DON'T immediately run `sync_rules_matrix.py`
- Instead:
  ```bash
  python apply_changes.py      # Pull team edits
  ```

  - Change log shows what team edited
  - Local matrix now has their changes
- Now run:
  ```bash
  python sync_rules_matrix.py   # Sync catalog
  ```

  - New SKUs are added
  - Discontinued SKUs are removed
  - Team edits are preserved (because you pulled them first)
- Google Sheets is updated with new catalog + team edits

**3:00 PM – Team continues**

- Team reviews the updated matrix
- Makes more edits if needed
- System is ready for next sync

---

## Troubleshooting

### Issue: "Error: Could not find Google Sheet '{vendor} Rules Matrix'"

**Cause:** Google Sheet doesn't exist or isn't shared with service account

**Solution:**

1. Verify the sheet exists in Google Drive
2. Verify the name is exactly `"{Vendor} Rules Matrix"` (case-sensitive)
3. Share the sheet with the service account:
   - Email: `southeast-pet-automation@YOUR_PROJECT.iam.gserviceaccount.com`
   - Role: Editor
4. Rerun the script

### Issue: "Error: Local matrix not found at ./Data/Rules/{vendor} Rules Matrix.xlsx"

**Cause:** Local matrix file doesn't exist

**Solution:**

1. Run `sync_rules_matrix.py` first to create the file
2. Or manually place an Excel file at that path

### Issue: "Error: Vendor name is invalid or empty."

**Cause:** You pressed Enter without typing a vendor name

**Solution:**

- Run the script again
- Type the vendor name (e.g., "Sysco")

### Issue: "No module named 'gspread'"

**Cause:** Dependencies not installed

**Solution:**

```bash
pip install gspread google-auth pandas openpyxl
```

### Issue: Script says "No changes detected" but I know the team made edits

**Cause:** Local matrix is already up-to-date (no differences to Google Sheets)

**Solution:**

- This is normal
- The script correctly detected no new changes
- Local matrix already matches Google Sheets

**Or:** Team edits might be stored differently (e.g., in comments, not column values)

### Issue: Change log doesn't show a change I made

**Cause:** The script only logs changes to Min, Max, and DNO columns

**What it logs:**

- New SKUs
- Removed SKUs
- Min column changes
- Max column changes
- DNO column changes

**What it doesn't log:**

- Changes to Item Name, Reporting Category (these are updated from the catalog, not tracked as edits)
- Comments
- Sheet formatting
- Other columns

---

## Common Workflows

### Workflow 1: Daily Sync

**Goal:** Pull team edits at the end of each day

**Steps:**

```bash
python apply_changes.py
```

- Check the change log
- Archive it with the date
- Local matrix is current

**Why:** Ensures you have the latest state before next day's work

### Workflow 2: Pre-Sync Verification

**Goal:** Before running sync_rules_matrix.py, confirm what team edited

**Steps:**

```bash
python apply_changes.py
```

- Review the change log
- Verify changes look reasonable
- Look for any errors (e.g., typos in Min/Max)

**If issues found:**

- Manually fix in Google Sheets
- Run `apply_changes.py` again

**If looks good:**

- Proceed to `sync_rules_matrix.py`

### Workflow 3: Compare Before & After

**Goal:** Document the state of the matrix before and after a sync

**Steps:**

1. Run `apply_changes.py` before sync
   - Change log shows: **Before State**
   - Documents team edits

2. Run `sync_rules_matrix.py`
   - New catalog is synced

3. Run `apply_changes.py` again
   - Change log shows: **After State**
   - Documents new SKUs, removals, etc.

**Result:** You have two change logs documenting before + after

### Workflow 4: Audit Trail for Compliance

**Goal:** Keep a record of all changes for audit purposes

**Steps:**

- Each time you run `apply_changes.py`, a timestamped change log is created
- Archive these files: `Data/log/`

**Organization:**

```
Data/log/
├── Sysco_ChangeLog_20240110_083000.xlsx
├── Sysco_ChangeLog_20240115_143022.xlsx
├── Sysco_ChangeLog_20240120_095500.xlsx
└── ...
```

**Use case:** Compliance, auditing, historical reference

### Workflow 5: Recovery After Accidental Changes

**Goal:** Undo recent changes to the local matrix

**Steps:**

1. Check the most recent change log
2. Note what was changed
3. Manually revert in Excel or Google Sheets
4. Run `apply_changes.py` again to pull the corrected version

**Or:** Restore from backup if you have one

---

## Summary

The **Apply Changes** script is essential infrastructure for safe, collaborative editing. It:

1. **Protects team work** – Pulls Google Sheets edits before pushing new data
2. **Provides visibility** – Change logs show exactly what changed
3. **Enables auditing** – Timestamped records of all changes
4. **Integrates seamlessly** – Works with `sync_rules_matrix.py`

**Key Rules:**

1. Always run `apply_changes.py` before `sync_rules_matrix.py`
2. Review the change log before confirming
3. Archive change logs for compliance
4. Use the confirmation prompt ("yes/no") as a final safety check

**Typical Usage:**

```bash
# Morning: Pull team edits from last night
python apply_changes.py

# Review the change log
cat Data/log/Sysco_ChangeLog_*.xlsx

# Afternoon: New catalog arrives, sync it
python sync_rules_matrix.py

# Evening: Team continues editing in Google Sheets
```

By following this workflow, you keep your local rules matrix synchronized with Google Sheets while safely integrating new catalog data.
