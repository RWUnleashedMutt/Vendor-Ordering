import streamlit as st
import pandas as pd
import io
import numpy as np
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIG ---
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

SHEET_IDS = {
    'All Vendors': '1iX-LpiavNqcyZqe1r068DmziafQbsDuugmdszqS89Tw',
    'Bradley Caldwell': '1eqENDXTdDJVKdos-VUXYNYMNM806rNcDrv63Q654nyc',
    'Fluff & Tuff': '1nGWM9Lt34e3vpqaETjPeMVsCTKVC9kIEQ3VVx1mEUqY'
    # Add more vendors here
}

MASTER_RULES_VENDOR = 'All Vendors'

# Stores that receive overflow/extra units from rounding up case packs
OVERFLOW_STORES = ['CVM', 'CC', 'LB', 'SH', 'LF', 'MF']

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
    'Current Quantity The Streets at Southpoint': 'SS'
}

inv_store_map = {v: k for k, v in store_map.items()}
all_stores = ['CC', 'CM', 'CVM', 'DTD',
              'LB', 'LF', 'MF', 'PP', 'SH', 'SP', 'SS']
priority_stores = ['CC', 'CM', 'CVM', 'LB', 'SH']


# --- HELPERS ---
def clean_id(val):
    if pd.isna(val):
        return ""
    return str(int(val)) if isinstance(val, float) and val.is_integer() else str(val)


@st.cache_data
def load_catalog(file) -> pd.DataFrame:
    df = pd.read_excel(file, header=1)
    df.columns = df.columns.str.strip()
    df['SKU'] = df['SKU'].apply(clean_id)
    if 'GTIN' in df.columns:
        df['GTIN'] = df['GTIN'].apply(clean_id)
    return df


@st.cache_resource
def get_google_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES
    )
    return gspread.authorize(creds)


@st.cache_data
def load_rules_from_sheets(vendor: str) -> pd.DataFrame:
    if vendor not in SHEET_IDS:
        raise ValueError(f"No Sheet ID configured for vendor '{vendor}'.")
    client = get_google_client()
    spreadsheet = client.open_by_key(SHEET_IDS[vendor])
    worksheet = spreadsheet.sheet1
    data = worksheet.get_all_records(value_render_option='UNFORMATTED_VALUE')
    df = pd.DataFrame(data)
    df.columns = df.columns.str.strip()
    df['SKU'] = df['SKU'].apply(clean_id)
    return df


def compute_store_order(df_master, rules_matrix, short_name, current_lt, hq_col):
    """Compute raw units needed per store before combined case pack logic."""
    long_name = inv_store_map[short_name]
    if long_name not in df_master.columns:
        return None

    lookup_cols = ['SKU', 'Order In Quantities',
                   f'{short_name}_DNO', f'{short_name}_Min', f'{short_name}_Max']
    valid_lookup = [c for c in lookup_cols if c in rules_matrix.columns]
    store_rules = rules_matrix[valid_lookup].copy().rename(columns={
        f'{short_name}_DNO': 'DNO',
        f'{short_name}_Min': 'Min',
        f'{short_name}_Max': 'Max'
    })

    store_inv = df_master[[
        'SKU', 'GTIN', 'Item Name', 'Default Unit Cost', long_name, hq_col
    ]].copy().rename(columns={long_name: 'Current_Inv', hq_col: 'HQ_Qty'})

    data = pd.merge(store_inv, store_rules, on='SKU', how='left')
    data = data.fillna({
        'DNO': False, 'Order In Quantities': 1, 'Min': 0,
        'Max': 0, 'Current_Inv': 0, 'HQ_Qty': 0, 'Default Unit Cost': 0
    })

    data['Effective_Min'] = data['Min'] + (current_lt * 0.2)
    data['Needs_Order'] = np.where(
        data['Order In Quantities'] == 1,
        (data['Current_Inv'] < data['Max']),
        (data['Current_Inv'] < data['Effective_Min'])
    )
    data['Needs_Order'] = data['Needs_Order'] & (data['DNO'] == False)

    # Raw units needed to reach max — no per-store case pack rounding yet
    data['Units_Needed'] = np.where(
        data['Needs_Order'],
        np.maximum(data['Max'] - data['Current_Inv'], 0),
        0
    )

    return data


def compute_combined_order(all_store_orders: dict, selected_stores: list) -> pd.DataFrame:
    """
    Combine all stores' raw unit needs into one order.
    1. Sum raw units needed across all stores per SKU
    2. Round up to nearest case pack ONCE across all stores combined
    3. Distribute extras proportionally to OVERFLOW_STORES based on their max
    Returns a per-SKU summary and updates each store's allocated units.
    """
    # Collect all SKU metadata and per-store needs
    sku_info = {}     # SKU -> {GTIN, Item Name, Case Pack, Default Unit Cost}
    store_needs = {}  # SKU -> {store: {units_needed, max, current_inv}}

    for store, data in all_store_orders.items():
        if data is None:
            continue
        for _, row in data.iterrows():
            sku = row['SKU']
            if sku not in sku_info:
                sku_info[sku] = {
                    'GTIN': row['GTIN'],
                    'Item Name': row['Item Name'],
                    'Case Pack': int(row['Order In Quantities']),
                    'Default Unit Cost': row['Default Unit Cost']
                }
            if sku not in store_needs:
                store_needs[sku] = {}
            store_needs[sku][store] = {
                'units_needed': max(int(row['Units_Needed']), 0),
                'max': max(int(row['Max']), 0),
                'current_inv': max(int(row['Current_Inv']), 0)
            }

    order_rows = []
    # store_allocations[store][sku] = final allocated units for that store
    store_allocations = {s: {} for s in selected_stores}

    for sku, stores in store_needs.items():
        case_pack = sku_info[sku]['Case Pack']
        if case_pack <= 0:
            case_pack = 1

        # Total raw units needed across all stores
        total_raw = sum(v['units_needed'] for v in stores.values())

        if total_raw == 0:
            for store in selected_stores:
                store_allocations[store][sku] = 0
            continue

        # Round up to nearest case pack across all stores combined
        total_cases = int(np.ceil(total_raw / case_pack))
        total_units_ordered = total_cases * case_pack
        extra_units = total_units_ordered - total_raw

        # Start with each store's raw need
        allocated = {store: stores.get(store, {}).get('units_needed', 0)
                     for store in selected_stores}

        # Distribute extras proportionally to overflow stores based on their max
        overflow_eligible = [
            s for s in OVERFLOW_STORES
            if s in selected_stores and stores.get(s, {}).get('max', 0) > 0
        ]

        if extra_units > 0 and overflow_eligible:
            total_max = sum(stores.get(s, {}).get('max', 0)
                            for s in overflow_eligible)

            remaining_extra = extra_units
            proportions = []
            for s in overflow_eligible:
                store_max = stores.get(s, {}).get('max', 0)
                proportion = store_max / total_max if total_max > 0 else 0
                proportions.append((s, proportion))

            # Distribute proportionally, giving whole units only
            distributed = {}
            running_total = 0
            for idx, (s, proportion) in enumerate(proportions):
                if idx == len(proportions) - 1:
                    # Last store gets the remainder to avoid rounding loss
                    units = remaining_extra - running_total
                else:
                    units = int(proportion * extra_units)
                    running_total += units
                distributed[s] = units

            for s, extra in distributed.items():
                current_inv = stores.get(s, {}).get('current_inv', 0)
                store_max = stores.get(s, {}).get('max', 0)
                # Cap extras so store doesn't exceed max + 1 case pack
                headroom = max((store_max + case_pack) -
                               current_inv - allocated[s], 0)
                allocated[s] += min(extra, headroom)

        for store in selected_stores:
            store_allocations[store][sku] = allocated.get(store, 0)

        order_rows.append({
            'SKU': sku,
            'GTIN': sku_info[sku]['GTIN'],
            'Item Name': sku_info[sku]['Item Name'],
            'Case Pack': case_pack,
            'Total Units Needed': total_raw,
            'Total Cases to Order': total_cases,
            'Total Units to Order': total_units_ordered,
            'Extra Units': extra_units,
            'Default Unit Cost': sku_info[sku]['Default Unit Cost']
        })

    combined_df = pd.DataFrame(order_rows) if order_rows else pd.DataFrame(
        columns=['SKU', 'GTIN', 'Item Name', 'Case Pack',
                 'Total Units Needed', 'Total Cases to Order',
                 'Total Units to Order', 'Extra Units', 'Default Unit Cost']
    )

    if not combined_df.empty:
        combined_df = combined_df.sort_values(
            'Item Name').reset_index(drop=True)

    return combined_df, store_allocations


def build_allocated_breakout(store_allocations: dict, selected_stores: list,
                             all_store_orders: dict) -> pd.DataFrame:
    """
    Build breakout using final allocated quantities (including extras)
    rather than raw units needed.
    """
    # Collect SKU metadata
    sku_meta = {}
    for store, data in all_store_orders.items():
        if data is None:
            continue
        for _, row in data.iterrows():
            sku = row['SKU']
            if sku not in sku_meta:
                sku_meta[sku] = {
                    'GTIN': row['GTIN'],
                    'Item Name': row['Item Name']
                }

    rows = []
    all_skus = set()
    for store_alloc in store_allocations.values():
        all_skus.update(store_alloc.keys())

    for sku in all_skus:
        total = sum(
            store_allocations[s].get(sku, 0) for s in selected_stores)
        if total == 0:
            continue
        row = {
            'Item Name': sku_meta.get(sku, {}).get('Item Name', ''),
            'GTIN': sku_meta.get(sku, {}).get('GTIN', '')
        }
        for store in selected_stores:
            row[store] = store_allocations[store].get(sku, 0)
        rows.append(row)

    if not rows:
        return pd.DataFrame()

    breakout_df = pd.DataFrame(rows)
    breakout_df = breakout_df.sort_values('Item Name').reset_index(drop=True)
    return breakout_df


def build_push_master_summary(all_store_pushes: dict) -> pd.DataFrame:
    frames = []
    for store, data in all_store_pushes.items():
        if data is None:
            continue
        pushed = data[data['Units_To_Push'] > 0][
            ['SKU', 'GTIN', 'Item Name', 'Units_To_Push', 'Cases_To_Push']
        ].copy()
        pushed['Store'] = store
        frames.append(pushed)

    if not frames:
        return pd.DataFrame(columns=['GTIN', 'Item Name', 'Total Cases', 'Total Units'])

    combined = pd.concat(frames, ignore_index=True)
    summary = (
        combined.groupby(['SKU', 'GTIN', 'Item Name'], as_index=False)
        .agg(Total_Cases=('Cases_To_Push', 'sum'),
             Total_Units=('Units_To_Push', 'sum'))
    )
    summary = summary.drop(columns='SKU')
    summary = summary.rename(columns={
        'Total_Cases': 'Total Cases', 'Total_Units': 'Total Units'})
    summary = summary.sort_values('Item Name').reset_index(drop=True)
    return summary


def build_breakout(all_store_orders: dict, selected_stores: list,
                   units_col: str = 'Total_Units_Needed') -> pd.DataFrame:
    sku_meta = {}
    store_units = {}

    for store, data in all_store_orders.items():
        if data is None:
            continue
        for _, row in data.iterrows():
            sku = row['SKU']
            if sku not in sku_meta:
                sku_meta[sku] = {
                    'GTIN': row['GTIN'], 'Item Name': row['Item Name']}
            if sku not in store_units:
                store_units[sku] = {}
            store_units[sku][store] = max(int(row[units_col]), 0)

    if not store_units:
        return pd.DataFrame()

    rows = []
    for sku, units in store_units.items():
        if sum(units.values()) == 0:
            continue
        row = {
            'Item Name': sku_meta[sku]['Item Name'],
            'GTIN': sku_meta[sku]['GTIN']
        }
        for store in selected_stores:
            row[store] = units.get(store, 0)
        rows.append(row)

    breakout_df = pd.DataFrame(rows)
    breakout_df = breakout_df.sort_values('Item Name').reset_index(drop=True)
    return breakout_df


def build_combined_excel(combined_df: pd.DataFrame, breakout_df: pd.DataFrame,
                         store_allocations: dict, selected_stores: list,
                         selected_vendor: str, date_str: str) -> bytes:
    """Build Excel with Master Order, Breakout, and one tab per store."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        wb = writer.book
        text_fmt = wb.add_format({'num_format': '@'})

        # Master Order tab
        export_master = combined_df[[
            'GTIN', 'Item Name', 'Case Pack',
            'Total Cases to Order', 'Total Units to Order', 'Extra Units'
        ]].copy()
        export_master.to_excel(writer, index=False, sheet_name='Master Order')
        ws = writer.sheets['Master Order']
        ws.set_column('A:A', 20, text_fmt)
        ws.set_column('B:B', 40)
        ws.set_column('C:F', 15)

        # Breakout tab
        if not breakout_df.empty:
            breakout_df.to_excel(writer, index=False, sheet_name='Breakout')
            ws_b = writer.sheets['Breakout']
            ws_b.set_column('A:A', 40)
            ws_b.set_column('B:B', 20, text_fmt)
            for col_idx in range(len(selected_stores)):
                ws_b.set_column(col_idx + 2, col_idx + 2, 8)

        # Per-store tabs
        for store in selected_stores:
            alloc = store_allocations.get(store, {})
            store_rows = [
                {'GTIN': combined_df.loc[
                    combined_df['SKU'] == sku, 'GTIN'].values[0]
                    if sku in combined_df['SKU'].values else '',
                 'Item Name': combined_df.loc[
                    combined_df['SKU'] == sku, 'Item Name'].values[0]
                    if sku in combined_df['SKU'].values else sku,
                 'Units Allocated': units}
                for sku, units in alloc.items() if units > 0
            ]
            if not store_rows:
                continue
            store_df = pd.DataFrame(store_rows).sort_values(
                'Item Name').reset_index(drop=True)
            store_df.to_excel(writer, index=False, sheet_name=store[:31])
            ws_s = writer.sheets[store[:31]]
            ws_s.set_column('A:A', 20, text_fmt)
            ws_s.set_column('B:B', 40)
            ws_s.set_column('C:C', 15)

    return buf.getvalue()


def build_master_excel(all_store_orders: dict, master_summary: pd.DataFrame,
                       breakout_df: pd.DataFrame, selected_stores: list,
                       units_col: str, cases_col: str) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        wb = writer.book
        text_fmt = wb.add_format({'num_format': '@'})

        master_summary.to_excel(writer, index=False,
                                sheet_name='Master Summary')
        ws = writer.sheets['Master Summary']
        ws.set_column('A:A', 20, text_fmt)
        ws.set_column('B:B', 40)
        ws.set_column('C:D', 15)

        if not breakout_df.empty:
            breakout_df.to_excel(writer, index=False, sheet_name='Breakout')
            ws_b = writer.sheets['Breakout']
            ws_b.set_column('A:A', 40)
            ws_b.set_column('B:B', 20, text_fmt)
            for col_idx in range(len(selected_stores)):
                ws_b.set_column(col_idx + 2, col_idx + 2, 8)

        for store, data in all_store_orders.items():
            if data is None:
                continue
            order = data[data[units_col] > 0][[
                'GTIN', 'Item Name', cases_col, units_col
            ]].copy()
            order = order.rename(columns={
                cases_col: 'Cases', units_col: 'Total Units'})
            if order.empty:
                continue
            order.to_excel(writer, index=False, sheet_name=store[:31])
            ws_s = writer.sheets[store[:31]]
            ws_s.set_column('A:A', 20, text_fmt)
            ws_s.set_column('B:B', 40)
            ws_s.set_column('C:D', 15)

    return buf.getvalue()


def file_prefix(date_str: str, vendor_label: str) -> str:
    return f"{date_str}_{vendor_label}" if vendor_label else date_str


def compute_store_push(warehouse_df, rules_matrix, short_name):
    """Compute the warehouse push for a single store."""
    long_name = inv_store_map[short_name]
    if long_name not in warehouse_df.columns:
        return None

    lookup_cols = ['SKU', 'Order In Quantities',
                   f'{short_name}_DNO', f'{short_name}_Min', f'{short_name}_Max']
    valid_lookup = [c for c in lookup_cols if c in rules_matrix.columns]
    store_rules = rules_matrix[valid_lookup].copy().rename(columns={
        f'{short_name}_DNO': 'DNO',
        f'{short_name}_Min': 'Min',
        f'{short_name}_Max': 'Max'
    })

    store_inv = warehouse_df[[
        'SKU', 'GTIN', 'Item Name', long_name, 'HQ_Qty'
    ]].copy().rename(columns={long_name: 'Current_Inv'})

    data = pd.merge(store_inv, store_rules, on='SKU', how='left')
    data = data.fillna({
        'DNO': False, 'Order In Quantities': 1,
        'Min': 0, 'Max': 0, 'Current_Inv': 0, 'HQ_Qty': 0
    })

    data['HQ_Qty'] = data['HQ_Qty'].clip(lower=0)
    data['Needs_Push'] = (
        (data['Current_Inv'] < data['Max']) &
        (data['DNO'] == False)
    )
    data['Units_To_Push'] = np.where(
        data['Needs_Push'],
        np.maximum(data['Max'] - data['Current_Inv'], 0),
        0
    )
    data['Units_To_Push'] = np.where(
        data['Units_To_Push'] > 0,
        np.ceil(data['Units_To_Push'] / data['Order In Quantities']
                ) * data['Order In Quantities'],
        0
    )
    data['Units_To_Push'] = np.minimum(
        data['Units_To_Push'], data['HQ_Qty']).clip(lower=0)
    data['Cases_To_Push'] = (
        data['Units_To_Push'] / data['Order In Quantities']).clip(lower=0)

    return data


# =========================================================
# PAGE: VENDOR ORDERING
# =========================================================
def page_vendor_ordering():
    st.title("🛒 Vendor Ordering System")

    with st.sidebar:
        st.header("1. Select Vendor")
        selected_vendor = st.selectbox(
            "Select a vendor:",
            options=["-- Select a Vendor --"] + [
                k for k in SHEET_IDS.keys() if k != MASTER_RULES_VENDOR
            ],
            key="vendor_select"
        )

        load_rules_btn = st.button(
            "📥 Load Rules from Google Sheets", key="vendor_load_btn")

        st.divider()
        st.header("2. Upload Catalog")
        catalog_file = st.file_uploader(
            "Upload Vendor Catalog (.xlsx)", type=['xlsx'], key="vendor_catalog")

        st.divider()
        st.header("3. Store Selection")
        selected_stores = st.multiselect(
            "Select stores:", options=all_stores,
            default=all_stores, key="vendor_stores"
        )

        st.divider()
        st.header("4. Store Lead Times (Days)")
        store_lead_times = {
            s: st.number_input(
                f"Lead Time: {s}", 0, 30,
                (1 if s in priority_stores else 7),
                key=f"vendor_lt_{s}")
            for s in selected_stores
        }

    rules_matrix = None
    if selected_vendor == "-- Select a Vendor --":
        st.sidebar.info("Please select a vendor to load rules.")
    elif load_rules_btn:
        with st.spinner(f"Loading rules for **{selected_vendor}**..."):
            try:
                rules_matrix = load_rules_from_sheets(selected_vendor)
                st.session_state["vendor_rules_matrix"] = rules_matrix
                st.session_state["vendor_rules_vendor"] = selected_vendor
                st.sidebar.success(f"✅ Rules loaded: {len(rules_matrix)} SKUs")
            except Exception as e:
                st.sidebar.error(f"❌ Failed to load rules: {e}")
    elif ("vendor_rules_matrix" in st.session_state and
          st.session_state.get("vendor_rules_vendor") == selected_vendor):
        rules_matrix = st.session_state["vendor_rules_matrix"]
        st.sidebar.success(f"✅ Rules loaded: {len(rules_matrix)} SKUs")

    if catalog_file and rules_matrix is not None and selected_stores:
        df_master = load_catalog(catalog_file)
        catalog_skus = set(df_master['SKU'].unique())
        rules_matrix = rules_matrix[rules_matrix['SKU'].isin(
            catalog_skus)].copy()

        hq_col = 'Current Quantity HQ'
        date_str = datetime.now().strftime("%Y-%m-%d")

        if hq_col not in df_master.columns:
            st.error(f"❌ Missing column: '{hq_col}'")
            st.stop()

        matched = len(rules_matrix['SKU'].unique())
        total = len(catalog_skus)
        st.caption(
            f"✅ {selected_vendor} — Matched {matched} of {total} catalog SKUs to rules.")

        # Compute raw store needs
        all_store_orders = {
            s: compute_store_order(
                df_master, rules_matrix, s, store_lead_times[s], hq_col)
            for s in selected_stores
        }

        # Combined case pack logic
        combined_df, store_allocations = compute_combined_order(
            all_store_orders, selected_stores)

        # Breakout uses final allocated quantities including extras
        breakout_df = build_allocated_breakout(
            store_allocations, selected_stores, all_store_orders)

        tab_labels = ["📋 Master Order", "📊 Breakout"] + selected_stores
        tabs = st.tabs(tab_labels)

        # --- MASTER ORDER TAB ---
        with tabs[0]:
            st.subheader(f"📋 Combined Master Order — {selected_vendor}")
            st.caption(
                "Case packs are calculated across all stores combined. "
                "Extra units from rounding are distributed proportionally "
                f"to: {', '.join(OVERFLOW_STORES)}."
            )

            if not combined_df.empty:
                total_cases = int(combined_df['Total Cases to Order'].sum())
                total_units = int(combined_df['Total Units to Order'].sum())
                total_cost = (combined_df['Total Units to Order'] *
                              combined_df['Default Unit Cost']).sum()
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Cases to Order", f"{total_cases}")
                col2.metric("Total Units to Order", f"{total_units}")
                col3.metric("Estimated Cost", f"${total_cost:,.2f}")

                display_cols = [
                    'GTIN', 'Item Name', 'Case Pack',
                    'Total Units Needed', 'Total Cases to Order',
                    'Total Units to Order', 'Extra Units'
                ]
                ed_combined = st.data_editor(
                    combined_df[display_cols],
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    key="combined_master_ed"
                )

                combined_excel = build_combined_excel(
                    combined_df, breakout_df, store_allocations,
                    selected_stores, selected_vendor, date_str
                )
                st.download_button(
                    "📥 Download Combined Order (All Stores + Breakout)",
                    combined_excel,
                    file_name=f"{date_str}_{selected_vendor}_Combined_Order.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.success("✅ No orders needed across any store.")

        # --- BREAKOUT TAB ---
        with tabs[1]:
            st.subheader(f"📊 Allocation Breakout — {selected_vendor}")
            st.caption(
                "Final units allocated per store, including extras distributed "
                f"proportionally to overflow stores: {', '.join(OVERFLOW_STORES)}."
            )
            if not breakout_df.empty:
                st.dataframe(
                    breakout_df, use_container_width=True, hide_index=True)
                buf_b = io.BytesIO()
                with pd.ExcelWriter(buf_b, engine='xlsxwriter') as writer:
                    breakout_df.to_excel(
                        writer, index=False, sheet_name='Breakout')
                    text_fmt = writer.book.add_format({'num_format': '@'})
                    ws_b = writer.sheets['Breakout']
                    ws_b.set_column('A:A', 40)
                    ws_b.set_column('B:B', 20, text_fmt)
                    for col_idx in range(len(selected_stores)):
                        ws_b.set_column(col_idx + 2, col_idx + 2, 8)
                st.download_button(
                    "📥 Download Breakout", buf_b.getvalue(),
                    file_name=f"{date_str}_{selected_vendor}_Breakout.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.success("✅ No allocations to break out.")

        # --- PER-STORE TABS ---
        for i, short_name in enumerate(selected_stores):
            with tabs[i + 2]:
                data = all_store_orders[short_name]
                if data is None:
                    st.error(
                        f"Missing column '{inv_store_map[short_name]}' in Catalog.")
                    continue

                st.subheader(f"🛒 Allocation: {short_name}")
                alloc = store_allocations.get(short_name, {})

                # Build display merging allocation back with item info
                store_display_rows = []
                for _, row in data.iterrows():
                    sku = row['SKU']
                    allocated = alloc.get(sku, 0)
                    if allocated == 0:
                        continue
                    store_display_rows.append({
                        'SKU': sku,
                        'GTIN': row['GTIN'],
                        'Item Name': row['Item Name'],
                        'Case Pack': int(row['Order In Quantities']),
                        'Units Needed': int(row['Units_Needed']),
                        'Units Allocated': allocated,
                        'Current Inv': int(row['Current_Inv']),
                        'Max': int(row['Max']),
                        'Default Unit Cost': row['Default Unit Cost']
                    })

                if store_display_rows:
                    store_df = pd.DataFrame(
                        store_display_rows).sort_values('Item Name').reset_index(drop=True)

                    frozen_mask = store_df['Item Name'].str.startswith(
                        'FRZN', na=False)

                    for label, df_type in [
                        ("📦 Dry", store_df[~frozen_mask]),
                        ("❄️ Frozen", store_df[frozen_mask])
                    ]:
                        st.markdown(f"#### {label}")
                        if not df_type.empty:
                            ed_df = st.data_editor(
                                df_type, use_container_width=True,
                                hide_index=True, num_rows="dynamic",
                                key=f"alloc_{label}_{short_name}"
                            )
                            cost = (ed_df['Units Allocated'] *
                                    ed_df['Default Unit Cost']).sum()
                            st.metric(f"{label} Allocated Cost",
                                      f"${cost:,.2f}")

                            export_df = ed_df[[
                                'GTIN', 'Item Name', 'Units Allocated']].copy()
                            buf = io.BytesIO()
                            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                                export_df.to_excel(
                                    writer, index=False, sheet_name='Allocation')
                                text_fmt = writer.book.add_format(
                                    {'num_format': '@'})
                                writer.sheets['Allocation'].set_column(
                                    'A:A', 20, text_fmt)
                                writer.sheets['Allocation'].set_column(
                                    'B:B', 40)
                                writer.sheets['Allocation'].set_column(
                                    'C:C', 15)
                            st.download_button(
                                f"📥 Download {label} Allocation", buf.getvalue(),
                                file_name=f"{date_str}_{selected_vendor}_Allocation_{short_name}.xlsx",
                                key=f"dl_alloc_{label}_{short_name}"
                            )
                        else:
                            st.write("No items in this category.")
                else:
                    st.success(f"✅ No allocation needed for {short_name}.")

    elif not selected_stores:
        st.warning("Please select at least one store in the sidebar to begin.")
    elif not catalog_file:
        st.info(
            "👋 **Welcome! Please select a vendor and upload the catalog to begin.**")
        col_inst, col_img = st.columns([1, 1])
        with col_inst:
            st.subheader("📋 Step-by-Step Export Instructions")
            st.markdown("""
            1. **Login to Square Dashboard.**
            2. **Go to Items → Item Library.**
            3. **Filter by the vendor you are ordering for.**
            4. **Click Actions → Export Library.**
            5. **Select "Export items matching applied filters".**
            6. **Upload the file here.**
            """)
        with col_img:
            st.subheader("📸 Reference Settings")
            try:
                st.image("./Data/Images/Export Example.png",
                         use_container_width=True,
                         caption="Select the 'Filtered' option.")
            except:
                st.warning("Reference image not found.")
    elif rules_matrix is None:
        st.warning(
            "⚠️ Please select a vendor and click 'Load Rules from Google Sheets' to continue.")


# =========================================================
# PAGE: WAREHOUSE PUSH
# =========================================================
def page_warehouse_push():
    st.title("🏭 Warehouse Push")

    with st.sidebar:
        st.header("1. Load Master Rules")
        vendor_label = st.text_input(
            "Vendor Label (for file names):",
            placeholder="e.g. Bradley Caldwell",
            key="push_vendor_label"
        ).strip()

        load_rules_btn = st.button(
            "📥 Load Master Rules from Google Sheets", key="push_load_btn")

        st.divider()
        st.header("2. Upload Warehouse Catalog")
        warehouse_file = st.file_uploader(
            "Upload Warehouse Inventory (.xlsx)", type=['xlsx'],
            key="warehouse_catalog"
        )

        st.divider()
        st.header("3. Store Selection")
        selected_stores = st.multiselect(
            "Select stores:", options=all_stores,
            default=all_stores, key="push_stores"
        )

    rules_matrix = None
    if load_rules_btn:
        with st.spinner("Loading master rules from Google Sheets..."):
            try:
                rules_matrix = load_rules_from_sheets(MASTER_RULES_VENDOR)
                st.session_state["push_rules_matrix"] = rules_matrix
                st.sidebar.success(
                    f"✅ Master rules loaded: {len(rules_matrix)} SKUs")
            except Exception as e:
                st.sidebar.error(f"❌ Failed to load master rules: {e}")
    elif "push_rules_matrix" in st.session_state:
        rules_matrix = st.session_state["push_rules_matrix"]
        st.sidebar.success(
            f"✅ Master rules loaded: {len(rules_matrix)} SKUs")
    else:
        st.sidebar.info("Click 'Load Master Rules' to begin.")

    if warehouse_file and rules_matrix is not None and selected_stores:
        warehouse_df = load_catalog(warehouse_file)

        hq_col_full = 'Current Quantity HQ'
        if hq_col_full not in warehouse_df.columns:
            st.error(
                f"❌ Missing column: '{hq_col_full}' in warehouse catalog.")
            st.stop()

        warehouse_df = warehouse_df.rename(columns={hq_col_full: 'HQ_Qty'})
        catalog_skus = set(warehouse_df['SKU'].unique())
        rules_filtered = rules_matrix[rules_matrix['SKU'].isin(
            catalog_skus)].copy()

        date_str = datetime.now().strftime("%Y-%m-%d")
        prefix = file_prefix(date_str, vendor_label)

        matched = len(rules_filtered['SKU'].unique())
        total = len(catalog_skus)
        st.caption(
            f"✅ Matched {matched} of {total} warehouse SKUs to master rules.")

        all_store_pushes = {
            s: compute_store_push(warehouse_df, rules_filtered, s)
            for s in selected_stores
        }

        push_summary = build_push_master_summary(all_store_pushes)
        breakout_df = build_breakout(
            all_store_pushes, selected_stores, 'Units_To_Push')

        tab_labels = ["📋 Master Push", "📊 Breakout"] + selected_stores
        tabs = st.tabs(tab_labels)

        with tabs[0]:
            st.subheader("📋 Master Push Summary")
            if vendor_label:
                st.caption(
                    f"Vendor: **{vendor_label}** — combined push quantities across all selected stores.")
            else:
                st.caption(
                    "Combined push quantities across all selected stores.")

            if not push_summary.empty:
                col1, col2 = st.columns(2)
                col1.metric("Total Cases to Push",
                            f"{int(push_summary['Total Cases'].sum())}")
                col2.metric("Total Units to Push",
                            f"{int(push_summary['Total Units'].sum())}")
                st.dataframe(
                    push_summary, use_container_width=True, hide_index=True)
                master_excel = build_master_excel(
                    all_store_pushes, push_summary, breakout_df,
                    selected_stores, 'Units_To_Push', 'Cases_To_Push'
                )
                if vendor_label:
                    st.download_button(
                        "📥 Download Master Push (All Stores + Breakout)",
                        master_excel,
                        file_name=f"{prefix}_Warehouse_Push_Master.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning(
                        "⚠️ Enter a Vendor Label in the sidebar to enable downloads.")
            else:
                st.success(
                    "✅ No pushes needed — all stores are at or above max.")

        with tabs[1]:
            st.subheader("📊 Push Breakout")
            st.caption(
                "One row per item being pushed, with units per store as columns.")
            if not breakout_df.empty:
                st.dataframe(
                    breakout_df, use_container_width=True, hide_index=True)
                buf_b = io.BytesIO()
                with pd.ExcelWriter(buf_b, engine='xlsxwriter') as writer:
                    breakout_df.to_excel(
                        writer, index=False, sheet_name='Breakout')
                    text_fmt = writer.book.add_format({'num_format': '@'})
                    ws_b = writer.sheets['Breakout']
                    ws_b.set_column('A:A', 40)
                    ws_b.set_column('B:B', 20, text_fmt)
                    for col_idx in range(len(selected_stores)):
                        ws_b.set_column(col_idx + 2, col_idx + 2, 8)
                if vendor_label:
                    st.download_button(
                        "📥 Download Breakout", buf_b.getvalue(),
                        file_name=f"{prefix}_Warehouse_Push_Breakout.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning(
                        "⚠️ Enter a Vendor Label in the sidebar to enable downloads.")
            else:
                st.success("✅ No items to break out.")

        for i, short_name in enumerate(selected_stores):
            with tabs[i + 2]:
                data = all_store_pushes[short_name]

                if data is None:
                    st.error(
                        f"Missing column '{inv_store_map[short_name]}' in warehouse catalog.")
                    continue

                st.subheader(f"🏭 Warehouse Push: {short_name}")
                push_data = data[data['Units_To_Push'] > 0][[
                    'SKU', 'GTIN', 'Item Name', 'Cases_To_Push',
                    'Order In Quantities', 'Units_To_Push',
                    'Current_Inv', 'Max', 'HQ_Qty'
                ]].copy().reset_index(drop=True)

                push_data.rename(columns={
                    'Cases_To_Push': 'Cases to Push',
                    'Order In Quantities': 'Case Pack',
                    'Units_To_Push': 'Units to Push',
                    'Current_Inv': 'Store On Hand',
                    'HQ_Qty': 'HQ On Hand'
                }, inplace=True)

                if not push_data.empty:
                    st.metric("Total Units to Push",
                              f"{int(push_data['Units to Push'].sum())}")
                    ed_df = st.data_editor(
                        push_data, use_container_width=True,
                        hide_index=True, num_rows="dynamic",
                        key=f"push_ed_{short_name}"
                    )
                    export_df = ed_df[[
                        'GTIN', 'Item Name', 'Cases to Push', 'Units to Push'
                    ]].copy()
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                        export_df.to_excel(
                            writer, index=False, sheet_name='Push_List')
                        text_fmt = writer.book.add_format({'num_format': '@'})
                        writer.sheets['Push_List'].set_column(
                            'A:A', 20, text_fmt)
                        writer.sheets['Push_List'].set_column('B:B', 40)
                        writer.sheets['Push_List'].set_column('C:D', 15)
                    if vendor_label:
                        st.download_button(
                            "📥 Download Push List", buf.getvalue(),
                            file_name=f"{prefix}_Warehouse_Push_{short_name}.xlsx",
                            key=f"dl_push_{short_name}"
                        )
                    else:
                        st.warning(
                            "⚠️ Enter a Vendor Label in the sidebar to enable downloads.",
                            icon="⚠️"
                        )
                else:
                    st.success(
                        f"✅ {short_name} is fully stocked — nothing to push.")

    elif not selected_stores:
        st.warning("Please select at least one store in the sidebar to begin.")
    elif not warehouse_file:
        st.info(
            "👋 **Please load master rules and upload the warehouse inventory catalog to begin.**")
    elif rules_matrix is None:
        st.warning(
            "⚠️ Please click 'Load Master Rules from Google Sheets' to continue.")


# =========================================================
# MAIN — PAGE ROUTING
# =========================================================
st.set_page_config(page_title="Inventory & Ordering System", layout="wide")

page = st.sidebar.selectbox(
    "📂 Navigate",
    options=["🛒 Vendor Ordering", "🏭 Warehouse Push"]
)

if page == "🛒 Vendor Ordering":
    page_vendor_ordering()
else:
    page_warehouse_push()
