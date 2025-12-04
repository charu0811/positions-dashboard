import streamlit as st
import pandas as pd
import numpy as np
import time
import sys
import os
import platform

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="DAP Position Manager",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- PASTE YOUR FULL PATH HERE (keep the r before the quotes) ---
FILE_PATH = r"C:\Users\c.charukant\Downloads\Live_DAP.xlsx"
FILE_NAME_ONLY = "Live_DAP.xlsx"

SHEET_MAIN = "DAP_Main"
SHEET_PROFIT = "Profit"

# -----------------------------------------------------------------------------
# HYBRID DATA LOADER
# -----------------------------------------------------------------------------
def load_data():
    """
    1. Tries to connect to an OPEN Excel workbook via xlwings (Live Mode).
    2. If not found/Linux, tries to read the file from disk using pandas (Static Mode).
    """
    status_log = []
    
    # CASE 1: Try Live Connection (Windows/Mac only)
    try:
        import xlwings as xw
        
        # Method A: Check specifically for the file name in open apps
        try:
            book = xw.books[FILE_NAME_ONLY]
            return book, f"Connected to open file: {FILE_NAME_ONLY}", True
        except:
            status_log.append("Target file not active in xw.books")

        # Method B: Loop through all open apps and check full paths
        # This catches it if the file is open but the name is slightly different in the title bar
        for app in xw.apps:
            for bk in app.books:
                # Check if the open book matches our target path or name
                if FILE_NAME_ONLY.lower() in bk.name.lower():
                    return bk, f"Found open book: {bk.name}", True
                try:
                    if bk.fullname.lower() == FILE_PATH.lower():
                        return bk, f"Found by path: {bk.name}", True
                except:
                    pass
        
        status_log.append("No matching open Excel found.")

    except ImportError:
        status_log.append("xlwings not installed.")
    except Exception as e:
        status_log.append(f"xlwings error: {e}")

    # CASE 2: Static Fallback (Pandas)
    # If we are here, we couldn't connect live. Let's try reading the file from disk.
    if os.path.exists(FILE_PATH):
        return None, f"Static Mode (Reading from Disk): {FILE_PATH}", False
    else:
        return None, f"FILE NOT FOUND at: {FILE_PATH}", False


def fetch_market_data(book, is_live):
    """
    Parses DAP_Main from either the live book object or the file on disk.
    """
    try:
        # --- GET RAW DATA ---
        if is_live and book:
            sheet = book.sheets[SHEET_MAIN]
            # Grab A1:AZ300 live
            df_raw = sheet.range('A1:AZ300').options(pd.DataFrame, header=False, index=False).value
        else:
            # Static read
            df_raw = pd.read_excel(FILE_PATH, sheet_name=SHEET_MAIN, header=None)
            df_raw = df_raw.iloc[:300] # Limit rows

        # --- FIND HEADERS ---
        # Locate the row containing "Outrights" or "Spread"
        header_row_idx = -1
        for i in range(15):
            row_vals = [str(x) for x in df_raw.iloc[i].values]
            if "Outrights" in row_vals or "Spread" in row_vals:
                header_row_idx = i
                break
        
        if header_row_idx == -1:
            return pd.DataFrame(), "Could not find 'Outrights' header row."

        # Get Headers
        headers = [str(x).strip() for x in df_raw.iloc[header_row_idx].values]
        data_rows = df_raw.iloc[header_row_idx+1:].reset_index(drop=True)

        # --- MAP COLUMNS ---
        # We search for column indices dynamically based on header names
        cols = {
            'out_name': -1, 'out_last': -1, 'out_tv': -1,
            'spd_name': -1, 'spd_last': -1,
            'fly_name': -1, 'fly_last': -1
        }

        for idx, h in enumerate(headers):
            h = h.lower()
            # Outrights (Left)
            if idx < 10:
                if 'outright' in h or 'symbol' in h: cols['out_name'] = idx
                if ('last' in h or 'price' in h) and cols['out_last'] == -1: cols['out_last'] = idx
            
            if idx < 15 and 'tick value' in h: cols['out_tv'] = idx
            
            # Spreads (Middle)
            if idx > 10 and idx < 20:
                if 'spread' in h and cols['spd_name'] == -1: cols['spd_name'] = idx
            
            if idx > 10 and idx < 25:
                if ('last' in h or 'ltp' in h) and cols['spd_last'] == -1 and idx > cols['spd_name']: 
                    cols['spd_last'] = idx

            # Flies (Right)
            if idx > 20:
                if 'fly' in h and cols['fly_name'] == -1: cols['fly_name'] = idx
                if ('last' in h or 'ltp' in h) and cols['fly_last'] == -1 and idx > cols['fly_name']:
                    cols['fly_last'] = idx

        # Fallbacks (Hardcoded based on your CSV snippet if dynamic fails)
        if cols['out_name'] == -1: cols['out_name'] = 1  # Column B
        if cols['out_last'] == -1: cols['out_last'] = 3  # Column D
        if cols['out_tv'] == -1: cols['out_tv'] = 14     # Column O (roughly)
        
        if cols['spd_name'] == -1: cols['spd_name'] = 13 # Column N
        if cols['spd_last'] == -1: cols['spd_last'] = 14 # Column O
        
        if cols['fly_name'] == -1: cols['fly_name'] = 25 # Column Z
        if cols['fly_last'] == -1: cols['fly_last'] = 26 # Column AA

        market_list = []

        def clean_float(x):
            try:
                return float(x)
            except:
                return 0.0

        for i, row in data_rows.iterrows():
            # 1. OUTRIGHTS
            if cols['out_name'] != -1:
                name = str(row[cols['out_name']])
                if name and name.lower() not in ['nan', 'none', '']:
                    market_list.append({
                        "Instrument": name, "Type": "Outright",
                        "Price": clean_float(row[cols['out_last']]),
                        "TickValue": clean_float(row[cols['out_tv']]) if cols['out_tv'] != -1 else 100.0
                    })
            
            # 2. SPREADS
            if cols['spd_name'] != -1:
                name = str(row[cols['spd_name']])
                if name and name.lower() not in ['nan', 'none', '']:
                     market_list.append({
                        "Instrument": name, "Type": "Spread",
                        "Price": clean_float(row[cols['spd_last']]),
                        "TickValue": 100.0 # Standard logic
                    })

            # 3. FLIES
            if cols['fly_name'] != -1:
                name = str(row[cols['fly_name']])
                if name and name.lower() not in ['nan', 'none', '']:
                     market_list.append({
                        "Instrument": name, "Type": "Fly",
                        "Price": clean_float(row[cols['fly_last']]),
                        "TickValue": 100.0
                    })

        return pd.DataFrame(market_list), "Success"

    except Exception as e:
        return pd.DataFrame(), str(e)

# -----------------------------------------------------------------------------
# APP UI
# -----------------------------------------------------------------------------
st.title("📊 DAP Live Dashboard")

# 1. Connection
book, msg, is_live = load_data()

with st.sidebar:
    st.header("Status")
    if is_live:
        st.success(f"🟢 {msg}")
        if st.button("🔄 Refresh Live"):
            st.rerun()
    else:
        # If static, allow refresh from disk
        st.warning(f"🟠 {msg}")
        if "FILE NOT FOUND" in msg:
            st.error("Please check the FILE_PATH in the code.")
        else:
            if st.button("📂 Reload File from Disk"):
                st.cache_data.clear()
                st.rerun()

# 2. Data Load
df_market, data_msg = fetch_market_data(book, is_live)

if df_market.empty:
    st.error(f"Data Error: {data_msg}")
    st.stop()

# 3. Positions (Session State)
if 'positions' not in st.session_state:
    st.session_state.positions = []

def add_trade(inst, lots, entry, tv):
    st.session_state.positions.append({
        "id": int(time.time()*10000),
        "Instrument": inst, "Lots": lots, "Entry": entry, "TV": tv
    })

# 4. Tabs
t1, t2, t3 = st.tabs(["Monitor", "Add Trade", "Data View"])

with t1:
    if not st.session_state.positions:
        st.info("No active trades.")
    else:
        total_pnl = 0.0
        pnl_list = []
        
        for p in st.session_state.positions:
            # Find live price
            row = df_market[df_market['Instrument'] == p['Instrument']]
            if not row.empty:
                live_price = row.iloc[0]['Price']
                # Use manual TV if set, otherwise market TV
                tv = p['TV'] if p['TV'] > 0 else row.iloc[0]['TickValue']
                
                pnl = (live_price - p['Entry']) * p['Lots'] * tv
                total_pnl += pnl
                
                pnl_list.append({
                    "ID": p['id'], "Instrument": p['Instrument'],
                    "Lots": p['Lots'], "Entry": p['Entry'],
                    "Live": live_price, "PnL": pnl
                })
            else:
                pnl_list.append({
                    "ID": p['id'], "Instrument": p['Instrument'],
                    "Lots": p['Lots'], "Entry": p['Entry'],
                    "Live": 0.0, "PnL": 0.0
                })

        st.metric("Total Open PnL", f"${total_pnl:,.2f}")
        
        # Display Rows
        for item in pnl_list:
            c1, c2, c3, c4, c5 = st.columns([3, 1, 2, 2, 1])
            c1.markdown(f"**{item['Instrument']}**")
            c2.text(f"{item['Lots']}")
            c3.text(f"@{item['Entry']}")
            
            color = "green" if item['PnL'] >= 0 else "red"
            c4.markdown(f":{color}[${item['PnL']:,.2f}]")
            
            if c5.button("X", key=f"del_{item['ID']}"):
                st.session_state.positions = [x for x in st.session_state.positions if x['id'] != item['ID']]
                st.rerun()
            st.divider()

with t2:
    st.subheader("New Trade")
    
    # Filter list
    type_filter = st.radio("Filter", ["All", "Spread", "Fly", "Outright"], horizontal=True)
    if type_filter != "All":
        opts = sorted(df_market[df_market['Type'] == type_filter]['Instrument'].unique())
    else:
        opts = sorted(df_market['Instrument'].unique())
        
    sel_inst = st.selectbox("Instrument", opts)
    
    # Get Current Details
    if sel_inst:
        curr = df_market[df_market['Instrument'] == sel_inst].iloc[0]
        st.caption(f"Live Price: {curr['Price']} | Type: {curr['Type']}")
        
        c_1, c_2 = st.columns(2)
        lots = c_1.number_input("Lots (+ Long / - Short)", value=1, step=1)
        entry = c_2.number_input("Entry Price", value=curr['Price'], format="%.4f")
        tv = st.number_input("Tick Value (Optional Override)", value=curr['TickValue'])
        
        if st.button("Add Position", type="primary"):
            add_trade(sel_inst, lots, entry, tv)
            st.success("Added!")
            time.sleep(0.5)
            st.rerun()

with t3:
    st.subheader("Parsed Market Data")
    st.dataframe(df_market, use_container_width=True)
    
    st.subheader("Raw Excel View")
    if st.button("Load Profit Sheet"):
        try:
            if is_live:
                data = book.sheets[SHEET_PROFIT].range('A1:Z50').options(pd.DataFrame).value
            else:
                data = pd.read_excel(FILE_PATH, sheet_name=SHEET_PROFIT).iloc[:50]
            st.dataframe(data)
        except Exception as e:
            st.error(f"Error loading profit sheet: {e}")
