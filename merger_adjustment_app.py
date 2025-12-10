import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Merger & Demerger Engine", layout="wide")


# -----------------------------------------------------------
# UTILITY: Create the output Excel file
# -----------------------------------------------------------
def create_output_excel(rows):
    df = pd.DataFrame(rows)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Transactions")
    return output.getvalue()


# -----------------------------------------------------------
# UI START
# -----------------------------------------------------------
st.title("📘 Merger & Demerger Adjustment Engine – Streamlit App")

# MODE SELECTION
mode = st.radio("Select Operation Mode:", ["Merger", "Demerger"])


# -----------------------------------------------------------
# COMMON: UPLOAD BASE FILE
# -----------------------------------------------------------
st.header("Step 1 — Upload Base File (Must Contain Merger & Demerger Sheets)")
base_file = st.file_uploader("Upload Base File", type=["xlsx", "xls"])

if not base_file:
    st.stop()

base_xl = pd.ExcelFile(base_file)


# =====================================================================
# ========================== DEMERGER MODE =============================
# =====================================================================
if mode == "Demerger":

    if "Demerger" not in base_xl.sheet_names:
        st.error("❌ 'Demerger' sheet missing.")
        st.stop()

    st.success("Base file loaded successfully!")

    st.header("Step 2 — Demerger Parameters")

    demerger_df = pd.read_excel(base_xl, sheet_name="Demerger")

    # USER INPUTS
    buy_date = st.date_input("Enter Buy Date for Demerged Shares")
    buy_rate = st.number_input("Enter Buy Rate / Market Rate", min_value=0.0, format="%.4f")

    if st.button("Generate Demerger Transaction File"):

        rows_output = []

        for _, row in demerger_df.iterrows():

            # Column mapping (from screenshot)
            stock_name   = str(row.iloc[22]).strip()   # W
            isin         = str(row.iloc[23]).strip()   # X
            client_code  = str(row.iloc[25]).strip()   # Z
            client_name  = str(row.iloc[26]).strip()   # AA

            # Units must be numeric
            try:
                units = float(str(row.iloc[27]).strip())
            except:
                units = 0.0

            rows_output.append({
                "Name": client_name,
                "Exchange Name": "NSE",
                "Segment Name": "Capital Market",
                "Date": buy_date,
                "Client Code": client_code,
                "F": "",
                "Scrip Name": stock_name,
                "ISIN": isin,
                "I": "", "J": "", "K": "",
                "Trade Type": "B",
                "M": "", "N": "", "O": "", "P": "",
                "BUY Quantity": units,
                "Sell Quantity": "",
                "Net Quantity": "",
                "Buy Market Value": units * buy_rate,
                "Sell Market Value": "",
                "Net Market Value": "",
                "Market Rate": buy_rate
            })

        final_excel = create_output_excel(rows_output)

        st.success("🎉 Demerger Transaction Sheet Generated for ALL Clients!")

        st.download_button(
            label="⬇️ Download Demerger File",
            data=final_excel,
            file_name="Demerger_All_Clients.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.stop()


# =====================================================================
# ============================ MERGER MODE =============================
# =====================================================================
if "Merger" not in base_xl.sheet_names:
    st.error("❌ 'Merger' sheet missing.")
    st.stop()

merger_df = pd.read_excel(base_xl, sheet_name="Merger")

st.header("Step 2 — Select Client (Merger Mode)")

# CLIENT NAME COLUMN IS AT INDEX 4
client_list = sorted(merger_df.iloc[:, 4].dropna().unique())
selected_client = st.selectbox("Choose Client", client_list)

if not selected_client:
    st.stop()

# Upload client file
st.header("Step 3 — Upload Client File (Return Computation Required)")
client_file = st.file_uploader("Upload Client File", type=["xlsx", "xls"])

if not client_file:
    st.stop()

client_xl = pd.ExcelFile(client_file)

if "Return Computation" not in client_xl.sheet_names:
    st.error("❌ 'Return Computation' sheet missing in client file.")
    st.stop()

rc_df = pd.read_excel(client_xl, sheet_name="Return Computation")

st.success("Client file loaded!")


# Extract this client’s merger data
row = merger_df[merger_df.iloc[:, 4] == selected_client].iloc[0]

old_stock = row.iloc[0]
old_isin = row.iloc[1]
client_code = row.iloc[3]
new_stock = row.iloc[9]
new_isin = row.iloc[10]
new_units = int(row.iloc[14])
deficit_units = int(row.iloc[16])

st.info(
    f"""
### Merger Details for **{selected_client}**
- Old Stock: **{old_stock}**
- Old ISIN: **{old_isin}**
- New Stock: **{new_stock}**
- New ISIN: **{new_isin}**
- New Units: **{new_units}**
- Deficit Units: **{deficit_units}**
"""
)

# User inputs
st.header("Step 4 — Enter Merger Inputs")

merger_date = st.date_input("Merger Effective Date")

# NEW INPUT — Sell Rate for 1 deficit unit
deficit_sell_rate = st.number_input("Market Sell Rate for Deficit 1 Unit", min_value=0.0, format="%.4f")

if st.button("Generate Merger Transaction File"):

    # FIFO rows: ISIN column index 2, Sale Date index 7
    fifo_rows = rc_df[(rc_df.iloc[:, 2] == old_isin) & (rc_df.iloc[:, 7].isna())]

    if fifo_rows.empty:
        st.error("❌ No unsold lots found for this ISIN in client file.")
        st.stop()

    fifo_lots = fifo_rows.iloc[:, [4, 5]].values.tolist()  # Quantity, Purchase Rate

    rows_output = []
    total_sell_excluding_deficit = 0  # Used to compute neutral buy price

    # 1) SELL DEFICIT FIRST (using manually entered market rate)
    if deficit_units > 0:
        qty, purchase_rate = fifo_lots[0]

        rows_output.append({
            "Name": selected_client,
            "Exchange Name": "NSE",
            "Segment Name": "Capital Market",
            "Date": merger_date,
            "Client Code": client_code,
            "F": "",
            "Scrip Name": old_stock,
            "ISIN": old_isin,
            "I": "", "J": "", "K": "",
            "Trade Type": "S",
            "M": "", "N": "", "O": "", "P": "",
            "BUY Quantity": "",
            "Sell Quantity": deficit_units,
            "Net Quantity": "",
            "Buy Market Value": "",
            "Sell Market Value": -(deficit_sell_rate * deficit_units),
            "Net Market Value": "",
            "Market Rate": deficit_sell_rate
        })

        # Reduce FIFO lot
        fifo_lots[0][0] -= deficit_units

    # 2) SELL REMAINING FIFO LOTS (using purchase rates)
    for qty, purchase_rate in fifo_lots:
        if qty <= 0:
            continue

        sell_value = -(purchase_rate * qty)
        total_sell_excluding_deficit += sell_value

        rows_output.append({
            "Name": selected_client,
            "Exchange Name": "NSE",
            "Segment Name": "Capital Market",
            "Date": merger_date,
            "Client Code": client_code,
            "F": "",
            "Scrip Name": old_stock,
            "ISIN": old_isin,
            "I": "", "J": "", "K": "",
            "Trade Type": "S",
            "M": "", "N": "", "O": "", "P": "",
            "BUY Quantity": "",
            "Sell Quantity": qty,
            "Net Quantity": "",
            "Buy Market Value": "",
            "Sell Market Value": sell_value,
            "Net Market Value": "",
            "Market Rate": purchase_rate
        })

    # 3) AUTO-CALCULATE BUY RATE TO MAKE NET = 0
    if new_units > 0:
        new_buy_rate = -(total_sell_excluding_deficit) / new_units
    else:
        new_buy_rate = 0

    # 4) BUY NEW STOCK ROW (using auto computed buy rate)
    rows_output.append({
        "Name": selected_client,
        "Exchange Name": "NSE",
        "Segment Name": "Capital Market",
        "Date": merger_date,
        "Client Code": client_code,
        "F": "",
        "Scrip Name": new_stock,
        "ISIN": new_isin,
        "I": "", "J": "", "K": "",
        "Trade Type": "B",
        "M": "", "N": "", "O": "", "P": "",
        "BUY Quantity": new_units,
        "Sell Quantity": "",
        "Net Quantity": "",
        "Buy Market Value": new_units * new_buy_rate,
        "Sell Market Value": "",
        "Net Market Value": "",
        "Market Rate": new_buy_rate
    })

    final_excel = create_output_excel(rows_output)

    st.success("🎉 Merger Transaction File Generated!")

    st.download_button(
        label="⬇️ Download Merger File",
        data=final_excel,
        file_name=f"{selected_client}_Merger.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
