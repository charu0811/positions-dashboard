import streamlit as st
import pandas as pd
import numpy as np
import time
import sys
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

FILE_NAME = "Live_DAP.xlsx"
SHEET_MAIN = "DAP_Main"
SHEET_PROFIT = "Profit"

# -----------------------------------------------------------------------------
# HYBRID DATA LOADER (XLWINGS vs PANDAS)
# -----------------------------------------------------------------------------
def load_data():
    """
    Attempts to connect to live Excel. If that fails (Linux/Server), 
    falls back to reading the file from disk.
    """
    system = platform.system()
    
    # CASE 1: Linux / Server Environment (No Live Excel)
    if system == "Linux":
        return None, "Static Mode (Linux)", False

    # CASE 2: Windows/Mac (Try Live Connection)
    try:
        import xlwings as xw
        # Try specific name
        try:
            book = xw.books[FILE_NAME]
            return book, "Connected Live", True
        except:
            # Try finding by partial name
            for app in xw.apps:
                for bk in app.books:
                    if "Live_DAP" in bk.name:
                        return bk, f"Connected: {bk.name}", True
        
        return None, "Excel file not found. Open Live_DAP.xlsx", False
    except ImportError:
        return None, "xlwings not installed", False
    except Exception as e:
        return None, f"Connection Error: {e}", False

def fetch_market_data(book, is_live):
    """
    Fetches data from DAP_Main. 
    If is_live=True, uses xlwings.
    If is_live=False, uses pandas read_excel.
    """
    try:
        # --- STRATEGY: GET RAW DATA ---
        if is_live and book:
            sheet = book.sheets[SHEET_MAIN]
            # Grab A1:AZ300
            df_raw = sheet.range('A1:AZ300').options(pd.DataFrame, header=False, index=False).value
        else:
            # Static Fallback
            df_raw = pd.read_excel(FILE_NAME, sheet_name=SHEET_MAIN, header=None)
            # Limit rows for performance if needed
            df_raw = df_raw.iloc[:300]

        # --- PARSING LOGIC (Same for both) ---
        # 1. Find the Header Row (Look for 'Outrights')
        header_row_idx = -1
        for i in range(10): # Check first 10 rows
            row_vals = [str(x) for x in df_raw.iloc[i].values]
            if "Outrights" in row_vals or "Spread" in row_vals:
                header_row_idx = i
                break
        
        if header_row_idx == -1:
            return pd.DataFrame(), "Header 'Outrights' not found in DAP_Main"

        # Get Headers and Data
        headers = [str(x).strip() for x in df_raw.iloc[header_row_idx].values]
        data_rows = df_raw.iloc[header_row_idx+1:].reset_index(drop=True)

        # 2. Identify Column Indices dynamically
        cols = {
            'out_name': -1, 'out_last': -1, 'out_tv': -1,
            'spd_name': -1, 'spd_last': -1,
            'fly_name': -1, 'fly_last': -1
        }

        # Heuristic search through headers
        for idx, h in enumerate(headers):
            h_lower = h.lower()
            # Outrights (usually Col A/B)
            if idx < 10 and ('outright' in h_lower or 'symbol' in h_lower): cols['out_name'] = idx
            if idx < 10 and ('last' in h_lower or 'price' in h_lower) and cols['out_last'] == -1: cols['out_last'] = idx
            if idx < 15 and 'tick value' in h_lower: cols['out_tv'] = idx
            
            # Spreads (Middle)
            if idx > 10 and idx < 20 and 'spread' in h_lower: cols['spd_name'] = idx
            if idx > 10 and idx < 25 and ('last' in h_lower or 'ltp' in h_lower) and cols['spd_last'] == -1 and idx > cols['spd_name']: cols['spd_last'] = idx

            # Flies (Right)
            if idx > 20 and 'fly' in h_lower: cols['fly_name'] = idx
            if idx > 20 and ('last' in h_lower or 'ltp' in h_lower) and cols['fly_last'] == -1 and idx > cols['fly_name']: cols['fly_last'] = idx

        # Fallbacks if headers are ambiguous (based on your file structure)
        if cols['out_name'] == -1: cols['out_name'] = 1 # Col B usually
        if cols['out_last'] == -1: cols['out_last'] = 3 
        if cols['spd_name'] == -1: cols['spd_name'] = 13
        if cols['spd_last'] == -1: cols['spd_last'] = 14
        if cols['fly_name'] == -1: cols['fly_name'] = 25
        if cols['fly_last'] == -1: cols['fly_last'] = 26

        market_list = []

        def clean_float(x):
            try:
                return float(x)
            except:
                return 0.0

        for i, row in data_rows.iterrows():
            # Outright
            if cols['out_name'] != -1:
                name = str(row[cols['out_name']])
                if name and name != 'nan' and name != 'None':
                    market_list.append({
                        "Instrument": name,
                        "Type": "Outright",
                        "Price": clean_float(row[cols['out_last']]),
                        "TickValue": clean_float(row[cols['out_tv']]) if cols['out_tv'] != -1 else 100.0
                    })
            
            # Spread
            if cols['spd_name'] != -1:
                name = str(row[cols['spd_name']])
                if name and name != 'nan' and name != 'None':
                    market_list.append({
                        "Instrument": name,
                        "Type": "Spread",
                        "Price": clean_float(row[cols['spd_last']]),
                        "TickValue": 100.0 # Default
                    })

            # Fly
            if cols['fly_name'] != -1:
                name = str(row[cols['fly_name']])
                if name and name != 'nan' and name != 'None':
                    market_list.append({
                        "Instrument": name,
                        "Type": "Fly",
                        "Price": clean_float(row[cols['fly_last']]),
                        "TickValue": 100.0 # Default
                    })

        df_final = pd.DataFrame(market_list)
        return df_final, "Success"

    except Exception as e:
        return pd.DataFrame(), str(e)

# -----------------------------------------------------------------------------
# APP LOGIC
# -----------------------------------------------------------------------------
st.title("📊 DAP Manager Dashboard")

# 1. CONNECT
book, status, is_live = load_data()

# Sidebar
with st.sidebar:
    st.header("Connection")
    if is_live:
        st.success(f"🟢 {status}")
        if st.button("🔄 Refresh Live"):
            st.rerun()
    else:
        st.warning(f"🟠 {status}")
        st.info("Using static file mode. (Live requires Windows/Mac + Excel)")
        if st.button("Reload File"):
            st.cache_data.clear()
            st.rerun()

# 2. LOAD DATA
df_market, msg = fetch_market_data(book, is_live)

if df_market.empty:
    st.error(f"Could not load market data: {msg}")
    st.stop()

# 3. SESSION STATE (Positions)
if 'positions' not in st.session_state:
    st.session_state.positions = []

def add_pos(inst, lots, entry, tv):
    st.session_state.positions.append({
        "id": int(time.time()*1000),
        "Instrument": inst, "Lots": lots, "Entry": entry, "TV": tv
    })

# 4. TABS
tab1, tab2, tab3 = st.tabs(["Monitor", "Strategy Builder", "Market Data"])

with tab1:
    st.subheader("Your Positions")
    if not st.session_state.positions:
        st.info("No trades active. Go to 'Strategy Builder'.")
    else:
        # Calculate PnL
        total_pnl = 0.0
        pnl_rows = []
        
        for p in st.session_state.positions:
            # Find current price
            match = df_market[df_market['Instrument'] == p['Instrument']]
            if not match.empty:
                curr = match.iloc[0]['Price']
                tv = p['TV'] if p['TV'] else match.iloc[0]['TickValue']
                
                # PnL Calc
                val = (curr - p['Entry']) * p['Lots'] * tv
                total_pnl += val
                
                pnl_rows.append({
                    "ID": p['id'], "Instrument": p['Instrument'],
                    "Lots": p['Lots'], "Entry": p['Entry'], "Live": curr,
                    "PnL": val
                })
            else:
                pnl_rows.append({
                    "ID": p['id'], "Instrument": p['Instrument'],
                    "Lots": p['Lots'], "Entry": p['Entry'], "Live": 0.0,
                    "PnL": 0.0
                })
                
        st.metric("Total PnL", f"${total_pnl:,.2f}")
        
        # Display
        for row in pnl_rows:
            c1, c2, c3, c4, c5 = st.columns([3, 1, 2, 2, 1])
            c1.markdown(f"**{row['Instrument']}**")
            c2.text(f"{row['Lots']}")
            c3.text(f"@{row['Entry']}")
            
            color = "green" if row['PnL'] >= 0 else "red"
            c4.markdown(f":{color}[${row['PnL']:,.2f}]")
            
            if c5.button("X", key=row['ID']):
                st.session_state.positions = [x for x in st.session_state.positions if x['id'] != row['ID']]
                st.rerun()
            st.divider()

with tab2:
    st.subheader("Add Trade")
    
    # Filter
    ftype = st.radio("Type", ["All", "Outright", "Spread", "Fly"], horizontal=True)
    if ftype != "All":
        opts = df_market[df_market['Type'] == ftype]['Instrument'].unique()
    else:
        opts = df_market['Instrument'].unique()
        
    sel = st.selectbox("Instrument", sorted(opts))
    
    # Details
    det = df_market[df_market['Instrument'] == sel].iloc[0]
    st.caption(f"Last: {det['Price']} | Type: {det['Type']}")
    
    c1, c2 = st.columns(2)
    l = c1.number_input("Lots", value=1)
    e = c2.number_input("Entry Price", value=det['Price'], format="%.4f")
    tv = st.number_input("Tick Value", value=det['TickValue'])
    
    if st.button("Add Position", type="primary"):
        add_pos(sel, l, e, tv)
        st.success("Added!")
        time.sleep(0.5)
        st.rerun()

with tab3:
    st.dataframe(df_market)
