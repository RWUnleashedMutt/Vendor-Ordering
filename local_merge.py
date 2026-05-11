import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import filedialog
from datetime import datetime


def get_file_path(title="Select File", file_types=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(title=title, filetypes=file_types)
    root.destroy()
    return file_path


def load_from_excel(path, label):
    """Load a local Excel file."""
    if not os.path.exists(path):
        print(f"Error: File not found at {path}")
        sys.exit(1)
    df = pd.read_excel(path, dtype={'SKU': str})
    print(f"Loaded {label}: {len(df)} SKUs")
    return df


def get_store_columns(df):
    """Detect all store codes in a DataFrame."""
    store_codes = set()
    for col in df.columns:
        if col.endswith('_Min') or col.endswith('_Max') or col.endswith('_DNO'):
            store_codes.add(col.rsplit('_', 1)[0])
    return sorted(store_codes)


def merge_vendor_into_matrix(matrix_df, vendor_df):
    """
    Match SKUs from the vendor sheet to the all-vendors matrix
    and update Min, Max, DNO, and Order In Quantities values where matches are found.
    SKUs in the vendor sheet that don't exist in the matrix are skipped.
    """
    matrix_df = matrix_df.copy()
    matrix_skus = set(matrix_df['SKU'].tolist())
    vendor_skus = set(vendor_df['SKU'].tolist())

    skipped_skus = vendor_skus - matrix_skus
    matched_skus = vendor_skus & matrix_skus

    if skipped_skus:
        print(
            f"\nSkipping {len(skipped_skus)} SKU(s) not found in the matrix:")
        for sku in sorted(skipped_skus):
            print(f"  - {sku}")

    print(f"\nMatched {len(matched_skus)} SKU(s) to update.")

    vendor_store_codes = get_store_columns(vendor_df)
    matrix_store_codes = get_store_columns(matrix_df)
    common_store_codes = [
        c for c in vendor_store_codes if c in matrix_store_codes]

    if not common_store_codes:
        print("Warning: No matching store columns found between the two sheets.")
        return matrix_df, [], skipped_skus

    matrix_df = matrix_df.set_index('SKU')
    vendor_df = vendor_df.set_index('SKU')

    change_records = []

    for sku in matched_skus:
        item_name = matrix_df.loc[sku,
                                  'Item Name'] if 'Item Name' in matrix_df.columns else ''

        # --- Check Order In Quantities ---
        if 'Order In Quantities' in vendor_df.columns and 'Order In Quantities' in matrix_df.columns:
            old_val = matrix_df.loc[sku, 'Order In Quantities']
            new_val = vendor_df.loc[sku, 'Order In Quantities']
            if str(old_val) != str(new_val):
                change_records.append({
                    'SKU': sku,
                    'Item Name': item_name,
                    'Store': 'All',
                    'Field': 'Order In Quantities',
                    'Old Value': old_val,
                    'New Value': new_val
                })
                matrix_df.at[sku, 'Order In Quantities'] = new_val

        # --- Check Min, Max, DNO ---
        for code in common_store_codes:
            for suffix in ['_Min', '_Max', '_DNO']:
                col = f'{code}{suffix}'
                if col in vendor_df.columns and col in matrix_df.columns:
                    old_val = matrix_df.loc[sku, col]
                    new_val = vendor_df.loc[sku, col]
                    if str(old_val) != str(new_val):
                        change_records.append({
                            'SKU': sku,
                            'Item Name': item_name,
                            'Store': code,
                            'Field': suffix.lstrip('_'),
                            'Old Value': old_val,
                            'New Value': new_val
                        })
                        matrix_df.at[sku, col] = new_val

    matrix_df = matrix_df.reset_index()
    changes_df = pd.DataFrame(change_records)
    return matrix_df, changes_df, skipped_skus


def export_change_log(changes_df, skipped_skus, output_dir, label):
    """Export a change log of what was updated."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(
        output_dir, f"{label}_MergeLog_{timestamp}.xlsx")
    os.makedirs(output_dir, exist_ok=True)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        summary_df = pd.DataFrame({
            'Category': ['Values Updated', 'SKUs Skipped (not in matrix)'],
            'Count': [len(changes_df), len(skipped_skus)]
        })
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        if not changes_df.empty:
            changes_df.to_excel(
                writer, sheet_name='Updated Values', index=False)
        else:
            pd.DataFrame({'Result': ['No value changes found.']}).to_excel(
                writer, sheet_name='Updated Values', index=False)

        if skipped_skus:
            skipped_df = pd.DataFrame({'SKU': sorted(skipped_skus)})
            skipped_df.to_excel(writer, sheet_name='Skipped SKUs', index=False)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_length = max((len(str(cell.value))
                                 for cell in col if cell.value), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(
                    max_length + 4, 50)

    print(f"\nChange log saved to: {output_path}")


def run_merge():
    """Main logic to load vendor sheet, merge into matrix, and save locally."""
    try:
        print(
            "\nPlease select the Vendor Sheet (the one with values already filled out)...")
        vendor_path = get_file_path(title="Select Vendor Sheet")
        if not vendor_path:
            print("No file selected. Exiting.")
            sys.exit(0)

        print("\nPlease select the All-Vendors Matrix...")
        matrix_path = get_file_path(title="Select All-Vendors Matrix")
        if not matrix_path:
            print("No file selected. Exiting.")
            sys.exit(0)

        vendor_df = load_from_excel(vendor_path, "Vendor Sheet")
        matrix_df = load_from_excel(matrix_path, "All-Vendors Matrix")

        updated_matrix_df, changes_df, skipped_skus = merge_vendor_into_matrix(
            matrix_df, vendor_df)

        print(f"\n--- Merge Summary ---")
        print(f"  Values Updated:  {len(changes_df)}")
        print(f"  SKUs Skipped:    {len(skipped_skus)}")

        if len(changes_df) == 0:
            print("\nNo changes detected. Matrix unchanged.")
            return

        confirm = input(
            "\nApply these changes to the all-vendors matrix? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Changes not applied. Exiting.")
            return

        updated_matrix_df.to_excel(matrix_path, index=False)
        print(f"\nMatrix updated at {matrix_path}")

        label = os.path.splitext(os.path.basename(vendor_path))[0]
        output_dir = os.path.dirname(matrix_path)
        export_change_log(changes_df, skipped_skus, output_dir, label)

        print("\nDone!")

    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    run_merge()
