import yfinance as yf
import pandas as pd

# Function to fetch data and save to Excel
def fetch_and_save_to_excel(ticker, sheet_name, file_name):
    data = ticker
    df = data.reset_index()
    df.to_excel(file_name, sheet_name=sheet_name, index=False)

# Ticker for Apple Inc.
aapl = yf.Ticker("AAPL")

# Fetch and save quarterly balance sheet data
fetch_and_save_to_excel(aapl.quarterly_balance_sheet, "Balance_Sheet", "Apple_data_combined.xlsx")

# Fetch and save quarterly cashflow data
fetch_and_save_to_excel(aapl.quarterly_cashflow, "Cashflow_Statement", "Apple_data_combined.xlsx")

# Fetch and save quarterly income statement data
fetch_and_save_to_excel(aapl.quarterly_income_stmt, "Income_Statement", "Apple_data_combined.xlsx")

# Read the combined Excel file
df_combined = pd.read_excel("Apple_data_combined.xlsx", sheet_name=None)

# Print the data from each sheet
for sheet_name, sheet_data in df_combined.items():
    print(f"\nSheet: {sheet_name}")
    print(sheet_data)
