import gspread
from google.oauth2.service_account import Credentials
import logging
from datetime import datetime
import tomllib

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

SHEET_IDS = {
    'Canine Caviar': '1TJXe9V_aF1A1wm_O9XK_iWJNU119iH3ZBBorUlX_0ss',
    'Fluff & Tuff': '1nGWM9Lt34e3vpqaETjPeMVsCTKVC9kIEQ3VVx1mEUqY',
    'SE': '1O6HWGeLgtdScnJ0_pQc8asaSj3-L4pP9vjCvvXa26vQ',
    'Front Porch Pets': '1CyW8rNNWzmYH9iqVRgN5iTWCiqgd-cJnrAJGktGS2a0',
    'Butchers Block': '1nDtvvDVu9tAzR2iDB4uMpUG3rcN3Fm_v09WJw3jvbJI',
    'Adored Beast': '1HwOxpAzI_HlntVVfOqxBVAWDy7cznPxxhUqOR5cy6ng',
    'InClover': '1GJX-rqphRYAHM50HKrXhE3qG3ZUeB9kP0njwcuM56co',
    'Bradley Caldwell': '1eqENDXTdDJVKdos-VUXYNYMNM806rNcDrv63Q654nyc'
}

# ID of the master sheet to write to - you'll need to create this first or provide the ID
MASTER_SHEET_ID = '1W-AGqIXwcqL7clDHad43hFmpPrrXzNUDYC4-dVGpngo'


def authenticate():
    """Authenticate with Google Sheets API using service account credentials from secrets.toml."""
    with open('.streamlit/secrets.toml', 'rb') as f:
        secrets = tomllib.load(f)

    creds = Credentials.from_service_account_info(
        secrets['gcp_service_account'],
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    return gspread.authorize(creds)


def get_sheet_data(client, sheet_id):
    """Get all data from the first worksheet in a spreadsheet."""
    try:
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.sheet1  # Get the first sheet, whatever it's named
        data = worksheet.get_all_values()
        logger.info(
            f"Retrieved {len(data)} rows from sheet ID {sheet_id} (sheet: {worksheet.title})")
        return data
    except Exception as e:
        logger.error(f"Error reading sheet {sheet_id}: {e}")
        return []


def consolidate_sheets(client):
    """Consolidate all vendor sheets into one master sheet."""
    all_data = []
    headers = None

    for vendor_name, sheet_id in SHEET_IDS.items():
        logger.info(f"Processing {vendor_name}...")
        data = get_sheet_data(client, sheet_id)

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

    client = authenticate()
    consolidated_data = consolidate_sheets(client)
    write_to_master_sheet(client, consolidated_data)

    logger.info(f"Consolidation completed at {datetime.now()}")


if __name__ == '__main__':
    main()
