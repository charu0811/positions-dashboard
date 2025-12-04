import streamlit as st
import xlwings as xw
import pandas as pd
import numpy as np
import time

# -----------------------------------------------------------------------------
# CONFIGURATION & SETUP
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="DAP Position Manager",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

EXCEL_FILE_NAME = "Live_DAP.xlsx"

# -----------------------------------------------------------------------------
# DATA CONNECTION FUNCTIONS
# -----------------------------------------------------------------------------
@st.cache_resource
def get_excel_app():
    """
    Connects to the active Excel instance.
    We use cache_resource to avoid re-connecting on every rerun,
    though xlwings handles this gracefully.
    """
    try:
        # Try to connect to the specific book if open
        book = xw.books[EXCEL_FILE_NAME]
        return book
    except Exception as e:
        return None

def fetch_market_data(book):
    """
    Reads the DAP_Main sheet and parses the stacked tables 
    (Outrights, Spreads, Flies) into a single clean DataFrame.
    """
    try:
        sheet = book.sheets['DAP_Main']
        # Read the used range. We assume a reasonable max size to keep it fast.
        # Adjust 'A1:Z500' if your sheet is massive.
        raw_data = sheet.range('A1:P200').options(pd.DataFrame, header=1, index=False).value
        
        # The sheet has multiple headers (Outrights, Spread, Fly)
        # We need to normalize this into a clean "Instrument -> Price" lookup.
        
        market_data = []
        
        # 1. Identify rows that look like data
        # We look for rows where column 0 (Symbol) is not None and not a Header keyword
        # Based on CSV: Col 0 is Symbol, Col 1 is Last, Col 2 is Settle, etc.
        
        # Clean numeric columns helper
        def clean_float(val):
            try:
                return float(val)
            except:
                return 0.0

        for index, row in raw_data.iterrows():
            symbol = row.iloc[0] # Assuming first column is Instrument Name
            
            # Skip empty rows or header repeaters
            if pd.isna(symbol) or str(symbol).strip() in ['Outrights', 'Spread', 'Fly', 'Symbol', 'Last']:
                continue
                
            # Extract basic data (Adjust indices based on your exact column layout)
            # Based on CSV snippet: 
            # Col 0: Symbol (Z25)
            # Col 1: Last (12.292)
            # Col 2: Settle
            # Col 3: BidQty
            # Col 4: Bid
            # Col 5: Ask
            # Col 8: Tick Value (roughly) - This varies in the csv, we'll try to catch it
            
            try:
                last_price = clean_float(row.iloc[1])
                bid = clean_float(row.iloc[4])
                ask = clean_float(row.iloc[5])
                
                # Tick Value often appears around column 8 or 9 in your sheet
                # We default to 100 if missing/zero, to prevent PnL errors
                tick_value = clean_float(row.iloc[8]) 
                if tick_value == 0: tick_value = 100.0 

                market_data.append({
                    "Instrument": str(symbol).strip(),
                    "Last": last_price,
                    "Bid": bid,
                    "Ask": ask,
                    "TickValue": tick_value
                })
            except Exception as e:
                continue

        return pd.DataFrame(market_data)

    except Exception as e:
        st.error(f"Error reading market data: {e}")
        return pd.DataFrame()

def fetch_excel_positions(book):
    """
    Attempts to read the 'Profit' sheet. 
    Since the format is complex (Strategy blocks), we grab the raw values 
    for display or simple parsing.
    """
    try:
        sheet = book.sheets['Profit']
        # Grabbing a large block to display raw for now
        data = sheet.range('A1:Z50').options(pd.DataFrame).value
        return data
    except:
        return pd.DataFrame()

# -----------------------------------------------------------------------------
# SIDEBAR
# -----------------------------------------------------------------------------
st.sidebar.title("🎮 DAP Manager")
st.sidebar.markdown("---")

# Connection Status
book = get_excel_app()
if book:
    st.sidebar.success(f"Connected: {EXCEL_FILE_NAME}")
    if st.sidebar.button("🔄 Refresh Data"):
        st.rerun()
else:
    st.sidebar.error(f"Not Found: {EXCEL_FILE_NAME}")
    st.sidebar.warning("Please open the Excel file and reload.")
    st.stop() # Stop execution if no excel

# Auto Refresh logic (simple loop)
auto_refresh = st.sidebar.checkbox("Auto-Refresh (5s)", value=False)
if auto_refresh:
    time.sleep(5)
    st.rerun()

# -----------------------------------------------------------------------------
# MAIN LOGIC
# -----------------------------------------------------------------------------

# 1. Load Market Data
df_market = fetch_market_data(book)

if df_market.empty:
    st.warning("Could not parse market data from 'DAP_Main'. Check column layout.")
else:
    # Set index for easy lookup
    df_market.set_index("Instrument", inplace=True)

# -----------------------------------------------------------------------------
# TABS
# -----------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📊 Live Positions & Dashboard", "🛠 Build Strategy", "📋 Raw Data"])

# --- TAB 1: DASHBOARD ---
with tab1:
    st.header("Live Position Monitor")
    
    # Initialize Session State for Positions if not exists
    if 'positions' not in st.session_state:
        st.session_state.positions = []

    # Calculate PnL for stored positions
    if not st.session_state.positions:
        st.info("No active positions tracked in Dashboard. Add one in the 'Build Strategy' tab.")
    else:
        portfolio_data = []
        total_pnl = 0.0

        for pos in st.session_state.positions:
            # pos structure: {'name': 'Fly1', 'legs': [('Q26', 1), ('K27', -2), ('Q28', 1)], 'lots': 10, 'entry': 5.5}
            
            current_structure_price = 0.0
            avg_tick_value = 0.0
            valid_price = True
            
            leg_details = []

            for instrument, ratio in pos['legs']:
                if instrument in df_market.index:
                    price = df_market.loc[instrument, 'Last']
                    tv = df_market.loc[instrument, 'TickValue']
                    
                    current_structure_price += (price * ratio)
                    avg_tick_value = tv # Taking last leg's TV or average
                    
                    leg_details.append(f"{instrument}({price})")
                else:
                    valid_price = False
                    leg_details.append(f"{instrument}(N/A)")

            # PnL Calculation
            # PnL = (Current Price - Entry Price) * Lots * TickValue
            # Note: For spreads/flies, ensure the price scaling matches your TV convention
            
            if valid_price:
                diff = current_structure_price - pos['entry']
                pnl = diff * pos['lots'] * avg_tick_value
                total_pnl += pnl
            else:
                pnl = 0.0

            portfolio_data.append({
                "Strategy": pos['name'],
                "Composition": " | ".join([f"{r}x{i}" for i, r in pos['legs']]),
                "Lots": pos['lots'],
                "Entry Price": pos['entry'],
                "Live Price": round(current_structure_price, 4),
                "Diff": round(current_structure_price - pos['entry'], 4),
                "Tick Val": avg_tick_value,
                "PnL": round(pnl, 2)
            })

        df_portfolio = pd.DataFrame(portfolio_data)
        
        # Display Metrics
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Open PnL", f"${total_pnl:,.2f}")
        col2.metric("Active Strategies", len(st.session_state.positions))
        
        # Display Table with styling
        st.dataframe(
            df_portfolio.style.format({
                "Entry Price": "{:.4f}",
                "Live Price": "{:.4f}",
                "Diff": "{:.4f}",
                "PnL": "{:.2f}"
            }).background_gradient(subset=['PnL'], cmap='RdYlGn', vmin=-1000, vmax=1000),
            use_container_width=True
        )
        
        if st.button("Clear All Positions"):
            st.session_state.positions = []
            st.rerun()

# --- TAB 2: BUILD STRATEGY ---
with tab2:
    st.subheader("Strategy Builder")
    st.markdown("Construct a Spread or Fly using live instruments found in `DAP_Main`.")

    col_build_1, col_build_2 = st.columns([1, 2])

    with col_build_1:
        st.markdown("#### Configuration")
        strat_type = st.radio("Type", ["Spread (2 Legs)", "Fly (3 Legs)", "Custom"])
        strat_name = st.text_input("Strategy Name", value="My Strategy")
        
        lots = st.number_input("Lots", value=1, step=1)
        entry_price = st.number_input("Entry Price", value=0.0, format="%.4f", help="Price at which you entered the trade")

    with col_build_2:
        st.markdown("#### Leg Selection")
        available_instruments = sorted(df_market.index.tolist())
        
        legs = []
        
        if strat_type == "Spread (2 Legs)":
            l1 = st.selectbox("Leg 1 (Long)", available_instruments, index=0)
            l2 = st.selectbox("Leg 2 (Short)", available_instruments, index=1)
            legs = [(l1, 1), (l2, -1)]
            
        elif strat_type == "Fly (3 Legs)":
            l1 = st.selectbox("Leg 1 (Wing)", available_instruments, index=0)
            l2 = st.selectbox("Leg 2 (Body)", available_instruments, index=1)
            l3 = st.selectbox("Leg 3 (Wing)", available_instruments, index=2)
            legs = [(l1, 1), (l2, -2), (l3, 1)]
            
        else: # Custom
            st.info("Add legs manually below")
            num_legs = st.number_input("Number of Legs", 1, 10, 4)
            for i in range(int(num_legs)):
                c1, c2 = st.columns(2)
                inst = c1.selectbox(f"Inst {i+1}", available_instruments, key=f"inst_{i}")
                ratio = c2.number_input(f"Ratio {i+1}", value=1.0, key=f"ratio_{i}")
                legs.append((inst, ratio))

        # Live Preview
        preview_price = 0.0
        details = []
        for inst, ratio in legs:
            if inst in df_market.index:
                p = df_market.loc[inst, 'Last']
                preview_price += (p * ratio)
                details.append(f"{ratio}x {inst} @ {p}")
        
        st.info(f"**Live Structure Price:** {preview_price:.4f}")
        st.caption(f"Calculation: {' + '.join(details)}")
        
        if st.button("Add to Dashboard", type="primary"):
            new_position = {
                'name': strat_name,
                'legs': legs,
                'lots': lots,
                'entry': entry_price
            }
            st.session_state.positions.append(new_position)
            st.success("Position Added! Go to 'Live Positions' tab to monitor.")

# --- TAB 3: RAW DATA ---
with tab3:
    st.subheader("Raw Market Data (DAP_Main)")
    st.dataframe(df_market, height=400)
    
    st.subheader("Excel 'Profit' Sheet View")
    st.caption("Raw dump of the sheet for reference")
    df_excel_profit = fetch_excel_positions(book)
    st.dataframe(df_excel_profit)