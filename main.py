import os
import yfinance as yf
from excel_files_creations import (
    fetch_save_all_data,
    merge_excel_files,
    delete_files
)
from format_excel import stretch_columns_and_move_data

# Define the ticker
aapl = yf.Ticker("MSFT")

# Define folder and file paths
folder_path = 'D:\\CodeProjects\\Python\\Stock-Analyzer-With-YahooFinance'
output_path = 'D:\\CodeProjects\\Python\\Stock-Analyzer-With-YahooFinance'

# Fetch and save all data
file_paths = fetch_save_all_data(aapl, folder_path)

# Merge all Excel files into one
merge_excel_files(folder_path, output_path)
# Stretch columns and move data
file_path = os.path.join(output_path, 'data_combined.xlsx')
stretch_columns_and_move_data(file_path)
# Delete the individual Excel files
delete_files(file_paths)
