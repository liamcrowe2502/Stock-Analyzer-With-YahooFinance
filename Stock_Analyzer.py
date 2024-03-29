import yfinance as yf
import pandas as pd
import os

""" 
    Install openpyxl if not already installed.
    Can cause error with sending data to excel fill if not installed
    You can install it using: pip install openpyxl
"""

aapl = yf.Ticker("AAPL")

# Fetch dividends data and parse 'Date' column during retrieval
dividends_df = aapl.dividends.reset_index()
# Convert 'Date' column to datetime
dividends_df['Date'] = pd.to_datetime(dividends_df['Date'])
# Convert datetime values to timezone-unaware datetimes
dividends_df['Date'] = dividends_df['Date'].dt.tz_localize(None)
# Save the data to an Excel file
dividends_df.to_excel("Apple_data_dividends.xlsx", index=False)

# Fetch quarterly cashflow data
cashflow_statement_df = aapl.quarterly_cash_flow
# Save the data to an Excel file
cashflow_statement_df.to_excel("Apple_data_quarterly_cashflow_statement.xlsx")

# Fetch quarterly balance sheet data
income_statement_df = aapl.quarterly_income_stmt
# Save the data to an Excel file
income_statement_df.to_excel("Apple_data_quarterly_income_statement.xlsx")

# Fetch quarterly balance sheet data
balance_sheet_df = aapl.quarterly_balance_sheet
# Save the data to an Excel file
balance_sheet_df.to_excel("Apple_data_quarterly_balance_sheet.xlsx")

#df = pd.read_excel("Apple_data_quarterly_balance_sheet.xlsx")
#print(df)

#Add all the data into a large spreadsheet
folder_path = 'D:\CodeProjects\Python\Stock-Analyzer-With-YahooFinance'
files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

dataFrames = []
for file in files:
    df = pd.read_excel(os.path.join(folder_path, file))
    dataFrames.append(df)

merged_df = pd.concat(dataFrames)

output_path = 'D:\CodeProjects\Python\Stock-Analyzer-With-YahooFinance'
merged_df.to_excel(os.path.join(output_path, 'data_combined.xlsx'), index=False)


"""
Format the data correctly
Widen the columns
Correct the numbers in the sheets
"""

"""
Undertand the how to change each of the data sheets
Look at yfinance library and openpyxl
"""