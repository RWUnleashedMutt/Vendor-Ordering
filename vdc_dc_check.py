import pandas as pd
import re

# --- Load your data ---
# Replace 'your_file.xlsx' with your actual file path
# or pd.read_csv('your_file.csv')
df = pd.read_excel(
    '../../Downloads/T6GBWNXFRGS1J_catalog-2026-05-14-1544.xlsx', header=1)

# --- Check for VDC or DC at the end of Item Name ---
# Uses a regex: matches ' VDC' or ' DC' at the end of the string (case-insensitive)
pattern = r'\s(VDC|DC)$'

df['Suffix_Match'] = df['Item Name'].str.extract(
    pattern, flags=re.IGNORECASE)[0]
df['Has_VDC_or_DC'] = df['Suffix_Match'].notna()

# --- Summary ---
total = len(df)
matched = df['Has_VDC_or_DC'].sum()
print(f"Total rows:     {total}")
print(f"Matched (VDC/DC): {matched}")
print(f"No match:       {total - matched}")

# --- View matched rows ---
print("\n--- Rows ending in VDC or DC ---")
print(df[df['Has_VDC_or_DC']][['SKU', 'Item Name', 'Suffix_Match']])

# --- View unmatched rows ---
print("\n--- Rows NOT ending in VDC or DC ---")
print(df[~df['Has_VDC_or_DC']][['SKU', 'Item Name']])

# --- Copyable SKU list ---
matched_skus = df[df['Has_VDC_or_DC']]['SKU'].tolist()
print("\n--- Matched SKUs (copyable list) ---")
print('\n'.join(str(sku) for sku in matched_skus))

# --- Optional: export results ---
df.to_excel('item_name_suffix_check.xlsx', index=False)
print("\nResults saved to item_name_suffix_check.xlsx")
