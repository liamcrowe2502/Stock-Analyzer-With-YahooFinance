import os
import yfinance as yf
from excel_files_creations import (
    fetch_and_save_dividends,
    fetch_and_save_cashflow,
    fetch_and_save_income_statement,
    fetch_and_save_balance_sheet,
    merge_excel_files
)
from format_excel import stretch_columns_and_move_data

# Define the ticker
aapl = yf.Ticker("AAPL")

# Define folder and file paths
folder_path = 'D:\\CodeProjects\\Python\\Stock-Analyzer-With-YahooFinance'
output_path = 'D:\\CodeProjects\\Python\\Stock-Analyzer-With-YahooFinance'

# Fetch and save data
fetch_and_save_dividends(aapl, os.path.join(folder_path, "Apple_data_dividends.xlsx"))
fetch_and_save_cashflow(aapl, os.path.join(folder_path, "Apple_data_quarterly_cashflow_statement.xlsx"))
fetch_and_save_income_statement(aapl, os.path.join(folder_path, "Apple_data_quarterly_income_statement.xlsx"))
fetch_and_save_balance_sheet(aapl, os.path.join(folder_path, "Apple_data_quarterly_balance_sheet.xlsx"))

# Merge all Excel files into one
merge_excel_files(folder_path, output_path)

# Stretch columns and move data
file_path = os.path.join(output_path, 'data_combined.xlsx')
stretch_columns_and_move_data(file_path)
