import yfinance as yf
import pandas as pd

""" 
        Install openpyxl if not already installed.
    Can cause error with sending data to excel fill if not installed
    You can install it using: pip install openpyxl
"""

aapl = yf.Ticker("AAPL")

# Fetch quarterly balance sheet data
#balance_sheet_df = aapl.quarterly_balance_sheet
# Save the data to an Excel file
# balance_sheet_df.to_excel("Apple_data.xlsx")

# Fetch dividends data and parse 'Date' column during retrieval
dividends_df = aapl.dividends.reset_index()
# Convert 'Date' column to datetime
dividends_df['Date'] = pd.to_datetime(dividends_df['Date'])
# Convert datetime values to timezone-unaware datetimes
dividends_df['Date'] = dividends_df['Date'].dt.tz_localize(None)
# Save the data to an Excel file
dividends_df.to_excel("Apple_data.xlsx", index=False)

