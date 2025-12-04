import streamlit as st
import xlwings as xw
import pandas as pd
import time

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="DAP Position Manager",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

TARGET_FILE_KEYWORD = "Live_DAP" # Will look for any book containing this
SHEET_MAIN = "DAP_Main"
SHEET_PROFIT = "Profit"

# -----------------------------------------------------------------------------
# ROBUST CONNECTION FUNCTION
# -----------------------------------------------------------------------------
def connect_to_excel():
    """
    Tries to find the Excel book using multiple methods.
    Returns: (book_object, status_message, is_success)
    """
    try:
        # 1. Try connecting to the specific name
        try:
            return xw.books['Live_DAP.xlsx'], "Connected to Live_DAP.xlsx", True
        except:
            pass

        # 2. Search through all open books for a partial match
        apps = xw.apps
        found_books = []
        for app in apps:
            for book in app.books:
                found_books.append(book.name)
                if TARGET_FILE_KEYWORD.lower() in book.name.lower():
                    return book, f"Found open file: {book.name}", True
        
        # 3. If we are here, we couldn't find it. Return list of what we found.
        if found_books:
            return None, f"Could not find '{TARGET_FILE_KEYWORD}'. Found: {', '.join(found_books)}", False
        else:
            return None, "No Excel instances found. Is Excel open?", False

    except Exception as e:
        return None, f"Critical Error: {str(e)}", False

# -----------------------------------------------------------------------------
# SMART DATA PARSING
# -----------------------------------------------------------------------------
def get_column_indices(header_row):
    """
    Scans the header row to find where Outrights, Spreads, and Flies start.
    """
    indices = {
        'outright_name': 0, 'outright_last': 1, 'outright_tv': 8,
        'spread_name': -1, 'spread_last': -1,
        'fly_name': -1, 'fly_last': -1
    }
    
    # Convert header row to list of strings
    headers = [str(x).strip() if x else "" for x in header_row]
    
    # 1. Scan for "Spread" and "Fly" headers
    # We assume 'Last' or 'LTP' follows the name
    
    for i, col_name in enumerate(headers):
        if col_name == "Spread":
            indices['spread_name'] = i
            # Usually Last/LTP is next to it
            if i+1 < len(headers): indices['spread_last'] = i+1
            
        if col_name == "Fly":
            indices['fly_name'] = i
            if i+1 < len(headers): indices['fly_last'] = i+1
            
        # Refine Tick Value search (usually near Outrights)
        if "Tick Value" in col_name and i < 15:
            indices['outright_tv'] = i

    return indices

def parse_market_data(book):
    try:
        sheet = book.sheets[SHEET_MAIN]
        
        # Pull large range
        # We assume headers are in Row 1 (index 0) or Row 2 (index 1)
        # We grab A1 to AZ300
        raw_values = sheet.range('A1:AZ300').options(pd.DataFrame, header=False, index=False).value
        
        # Find the header row (contains "Outright" or "Spread")
        header_idx = 0
        found_header = False
        for i in range(5): # Check first 5 rows
            row_str = [str(x) for x in raw_values.iloc[i].tolist()]
            if "Outright" in row_str or "Spread" in row_str:
                header_idx = i
                found_header = True
                break
        
        if not found_header:
            return pd.DataFrame(), "Could not find headers in DAP_Main"

        headers = raw_values.iloc[header_idx].tolist()
        idxs = get_column_indices(headers)
        
        market_data = {}
        
        # Helper
        def clean_val(val):
            try:
                return float(val)
            except:
                return 0.0

        # Iterate rows after header
        for i in range(header_idx + 1, len(raw_values)):
            row = raw_values.iloc[i]
            
            # --- Outrights ---
            name = row[idxs['outright_name']]
            if name and isinstance(name, str) and len(name) < 10:
                market_data[name] = {
                    "Type": "Outright",
                    "Price": clean_val(row[idxs['outright_last']]),
                    "TickValue": clean_val(row[idxs['outright_tv']]) if idxs['outright_tv'] != -1 else 100.0
                }
                
            # --- Spreads ---
            if idxs['spread_name'] != -1:
                s_name = row[idxs['spread_name']]
                if s_name and isinstance(s_name, str):
                    market_data[s_name] = {
                        "Type": "Spread",
                        "Price": clean_val(row[idxs['spread_last']]),
                        "TickValue": 100.0
                    }

            # --- Flies ---
            if idxs['fly_name'] != -1:
                f_name = row[idxs['fly_name']]
                if f_name and isinstance(f_name, str):
                    market_data[f_name] = {
                        "Type": "Fly",
                        "Price": clean_val(row[idxs['fly_last']]),
                        "TickValue": 100.0
                    }

        df = pd.DataFrame.from_dict(market_data, orient='index').reset_index().rename(columns={"index": "Instrument"})
        return df, "Success"

    except Exception as e:
        return pd.DataFrame(), f"Error parsing data: {str(e)}"

# -----------------------------------------------------------------------------
# MAIN APP
# -----------------------------------------------------------------------------
st.title("📊 DAP Live Manager")

# Sidebar Status
with st.sidebar:
    st.header("Connection Status")
    
    book, msg, success = connect_to_excel()
    
    if success:
        st.success(msg)
        if st.button("🔄 Refresh Data"):
            st.rerun()
    else:
        st.error("Connection Failed")
        st.warning(msg)
        if st.button("Retry Connection"):
            st.rerun()
        st.stop()

# Initialize Session State
if 'positions' not in st.session_state:
    st.session_state['positions'] = []

def add_pos(inst, lots, entry, tv):
    st.session_state['positions'].append({
        "id": int(time.time()*1000),
        "Instrument": inst, "Lots": lots, "Entry": entry, "TV": tv
    })

# Load Data
df_market, data_msg = parse_market_data(book)

if df_market.empty:
    st.error(f"Data Error: {data_msg}")
    st.stop()

# TABS
t1, t2, t3 = st.tabs(["Monitor", "Builder", "Raw Data"])

with t1:
    if not st.session_state['positions']:
        st.info("No positions. Add one in the 'Builder' tab.")
    else:
        portfolio = []
        total_pnl = 0.0
        
        for p in st.session_state['positions']:
            row = df_market[df_market['Instrument'] == p['Instrument']]
            if not row.empty:
                curr = row.iloc[0]['Price']
                tv = p['TV'] if p['TV'] else row.iloc[0]['TickValue']
                # PnL logic
                pnl = (curr - p['Entry']) * p['Lots'] * tv
                total_pnl += pnl
                portfolio.append({
                    "ID": p['id'], "Instrument": p['Instrument'], 
                    "Lots": p['Lots'], "Entry": p['Entry'], 
                    "Live": curr, "PnL": pnl
                })
            else:
                 portfolio.append({
                    "ID": p['id'], "Instrument": p['Instrument'], 
                    "Lots": p['Lots'], "Entry": p['Entry'], 
                    "Live": 0.0, "PnL": 0.0
                })
        
        st.metric("Total PnL", f"${total_pnl:,.2f}")
        
        # Table
        for item in portfolio:
            c1, c2, c3, c4, c5 = st.columns([3, 1, 2, 2, 1])
            c1.write(f"**{item['Instrument']}**")
            c2.write(f"{item['Lots']} lots")
            c3.write(f"@{item['Entry']}")
            
            val_color = "green" if item['PnL'] >= 0 else "red"
            c4.markdown(f":{val_color}[${item['PnL']:,.2f}]")
            
            if c5.button("🗑", key=item['ID']):
                st.session_state['positions'] = [x for x in st.session_state['positions'] if x['id'] != item['ID']]
                st.rerun()
            st.divider()

with t2:
    st.subheader("Add Trade")
    # Clean list
    opts = sorted(df_market['Instrument'].unique().tolist())
    sel = st.selectbox("Instrument", opts)
    
    curr_data = df_market[df_market['Instrument'] == sel].iloc[0]
    st.caption(f"Live Price: {curr_data['Price']} | Type: {curr_data['Type']}")
    
    c1, c2 = st.columns(2)
    l = c1.number_input("Lots", value=1)
    e = c2.number_input("Entry", value=curr_data['Price'], format="%.4f")
    tv = st.number_input("Tick Value", value=curr_data['TickValue'])
    
    if st.button("Add"):
        add_pos(sel, l, e, tv)
        st.success("Added!")
        time.sleep(0.5)
        st.rerun()

with t3:
    st.write(df_market)
    if st.button("Show Profit Sheet Raw"):
        try:
            st.write(book.sheets[SHEET_PROFIT].range('A1:Z100').options(pd.DataFrame).value)
        except:
            st.error("Profit sheet not found")
