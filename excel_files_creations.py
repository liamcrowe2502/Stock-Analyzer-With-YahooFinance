import pandas as pd
import os

def fetch_and_save_dividends(ticker, filename):
    dividends_df = ticker.dividends.reset_index()
    dividends_df['Date'] = pd.to_datetime(dividends_df['Date'])
    dividends_df['Date'] = dividends_df['Date'].dt.tz_localize(None)
    dividends_df.to_excel(filename, index=False)

def fetch_and_save_cashflow(ticker, filename):
    cashflow_statement_df = ticker.quarterly_cash_flow
    cashflow_statement_df.to_excel(filename)

def fetch_and_save_income_statement(ticker, filename):
    income_statement_df = ticker.quarterly_income_stmt
    income_statement_df.to_excel(filename)

def fetch_and_save_balance_sheet(ticker, filename):
    balance_sheet_df = ticker.quarterly_balance_sheet
    balance_sheet_df.to_excel(filename)

def merge_excel_files(folder_path, output_path):
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    dataFrames = []
    for file in files:
        df = pd.read_excel(os.path.join(folder_path, file))
        dataFrames.append(df)
    merged_df = pd.concat(dataFrames)
    merged_df.to_excel(os.path.join(output_path, 'data_combined.xlsx'), index=False)

def delete_files(file_paths):
    for file_path in file_paths:
        if os.path.exists(file_path):
            os.remove(file_path)
        else:
            print(f"The file {file_path} does not exist")

def fetch_save_all_data(ticker, folder_path):
    dividends_file = os.path.join(folder_path, "Apple_data_dividends.xlsx")
    cashflow_file = os.path.join(folder_path, "Apple_data_quarterly_cashflow_statement.xlsx")
    income_statement_file = os.path.join(folder_path, "Apple_data_quarterly_income_statement.xlsx")
    balance_sheet_file = os.path.join(folder_path, "Apple_data_quarterly_balance_sheet.xlsx")

    fetch_and_save_dividends(ticker, dividends_file)
    fetch_and_save_cashflow(ticker, cashflow_file)
    fetch_and_save_income_statement(ticker, income_statement_file)
    fetch_and_save_balance_sheet(ticker, balance_sheet_file)

    return [dividends_file, cashflow_file, income_statement_file, balance_sheet_file]
