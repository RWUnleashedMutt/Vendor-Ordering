import time
import gspread
from google.oauth2.service_account import Credentials
import logging
from datetime import datetime
import tomllib
import streamlit as st

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# with open('./.streamlit/secrets.toml', 'rb') as f:
#     config = tomllib.load(f)

# SHEET_IDS = config['cons_sheet_ids']
SHEET_IDS = st.secrets['cons_sheet_ids']

# SHEET_IDS = config['cons_sheet_ids']

# ID of the master sheet to write to - you'll need to create this first or provide the ID
MASTER_SHEET_ID = '1W-AGqIXwcqL7clDHad43hFmpPrrXzNUDYC4-dVGpngo'

# Vendor name(s) to exclude when the user picks the "all except SE" option
EXCLUDE_KEYWORD = 'SE'


def authenticate():
    """Authenticate with Google Sheets API using service account credentials from secrets.toml."""

    creds = Credentials.from_service_account_info(
        st.secrets['gcp_service_account'],
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    return gspread.authorize(creds)


def select_vendors_interactively(sheet_ids):
    """
    Prompt the user in the console to choose which vendors to consolidate:
      [A] All vendors
      [E] All vendors except SE
      [C] Custom selection (pick specific vendors)
    Returns a filtered dict of {vendor_name: sheet_id}.
    """
    vendor_names = list(sheet_ids.keys())

    print("\n" + "="*60)
    print("VENDOR SELECTION")
    print("="*60)
    print("  [A] All vendors")
    print("  [E] All vendors except SE")
    print("  [C] Custom selection")

    choice = input(
        "\nWhich vendors do you want to consolidate? (A/E/C): ").strip().lower()

    if choice == 'a':
        return sheet_ids

    if choice == 'e':
        filtered = {
            name: sid for name, sid in sheet_ids.items()
            if name.lower() != EXCLUDE_KEYWORD.lower()
        }
        excluded = [name for name in vendor_names if name not in filtered]
        if excluded:
            print(f"Excluding: {', '.join(excluded)}")
        else:
            print(
                f"No vendor named '{EXCLUDE_KEYWORD}' found — running all vendors.")
        return filtered

    if choice == 'c':
        print("\nAvailable vendors:")
        for i, name in enumerate(vendor_names, start=1):
            print(f"  {i}. {name}")

        while True:
            selection = input(
                "\nEnter vendor numbers or names, comma-separated (e.g. 1,3,5 or Acme,Beta): "
            ).strip()

            if not selection:
                print("  Please enter at least one vendor.")
                continue

            tokens = [t.strip() for t in selection.split(',') if t.strip()]
            chosen = {}
            invalid = []

            for token in tokens:
                if token.isdigit():
                    idx = int(token)
                    if 1 <= idx <= len(vendor_names):
                        name = vendor_names[idx - 1]
                        chosen[name] = sheet_ids[name]
                    else:
                        invalid.append(token)
                else:
                    matches = [
                        name for name in vendor_names if name.lower() == token.lower()]
                    if matches:
                        chosen[matches[0]] = sheet_ids[matches[0]]
                    else:
                        invalid.append(token)

            if invalid:
                print(f"  Could not match: {', '.join(invalid)}. Try again.")
                continue

            if not chosen:
                print("  No valid vendors selected. Try again.")
                continue

            return chosen

    # Fallback for anything unrecognized - default to all vendors
    print("Unrecognized option — running all vendors.")
    return sheet_ids


def get_sheet_data(client, sheet_id, single_vendor=False):
    """Get all data from the first worksheet in a spreadsheet."""
    try:
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.sheet1  # Get the first sheet, whatever it's named
        data = worksheet.get_all_values()
        if not single_vendor:
            time.sleep(4)
        logger.info(
            f"Retrieved {len(data)} rows from sheet ID {sheet_id} (sheet: {worksheet.title})")
        return data
    except Exception as e:
        logger.error(f"Error reading sheet {sheet_id}: {e}")
        return []


def consolidate_sheets(client, sheet_ids):
    """Consolidate the selected vendor sheets into one master sheet."""
    all_data = []
    headers = None

    single_vendor = len(sheet_ids) == 1

    for vendor_name, sheet_id in sheet_ids.items():
        logger.info(f"Processing {vendor_name}...")
        data = get_sheet_data(client, sheet_id, single_vendor=single_vendor)

        if not data:
            logger.warning(f"No data found for {vendor_name}")
            continue

        # First vendor - capture headers
        if headers is None and len(data) > 0:
            headers = data[0] + ['Vendor']
            all_data.append(headers)

        # Add rows with vendor name
        for row in data[1:]:  # Skip header row
            row_with_vendor = row + [vendor_name]
            all_data.append(row_with_vendor)

    logger.info(f"Total consolidated rows: {len(all_data)}")
    return all_data


def write_to_master_sheet(client, data):
    """Write consolidated data to master sheet, clearing existing data first."""
    try:
        spreadsheet = client.open_by_key(MASTER_SHEET_ID)
        worksheet = spreadsheet.sheet1  # Writes to first sheet

        # Clear existing data
        worksheet.clear()
        logger.info("Cleared master sheet")

        # Write new data
        worksheet.update(data, value_input_option='RAW')
        logger.info(f"Wrote {len(data)} rows to master sheet")

    except Exception as e:
        logger.error(f"Error writing to master sheet: {e}")
        raise


def main():
    """Main function to run consolidation."""
    logger.info("Starting vendor sheet consolidation...")

    if MASTER_SHEET_ID == 'YOUR_MASTER_SHEET_ID_HERE':
        logger.error("Please set MASTER_SHEET_ID to your master sheet's ID")
        return

    selected_vendors = select_vendors_interactively(SHEET_IDS)
    logger.info(
        f"Selected {len(selected_vendors)} vendor(s): {', '.join(selected_vendors.keys())}")

    client = authenticate()
    consolidated_data = consolidate_sheets(client, selected_vendors)
    write_to_master_sheet(client, consolidated_data)

    logger.info(f"Consolidation completed at {datetime.now()}")


if __name__ == '__main__':
    main()
