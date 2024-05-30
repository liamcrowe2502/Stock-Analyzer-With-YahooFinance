from openpyxl import load_workbook

def stretch_columns_and_move_data(file_path):
    wb = load_workbook(file_path)

    # Rename the active sheet to "Dividends"
    ws = wb.active
    ws.title = "Dividends"

    # Set column widths
    column_widths = {'A': 20, 'B': 20, 'C': 20, 'D': 20, 'E': 20, 'F': 20, 'G': 20, 'H': 20}
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # Create and prepare "Extracted_Data" sheet
    new_ws = wb.create_sheet(title='Extracted_Data')
    headers = [cell.value for cell in ws[1][2:8]]
    new_ws.append(headers)

    # Move data from "Dividends" to "Extracted_Data"
    for row in ws.iter_rows(min_row=85, max_row=233, min_col=3, max_col=8):
        new_ws.append([cell.value for cell in row])

    # Clear the original data in "Dividends"
    for cell in ws.iter_cols(min_col=3, max_col=8, min_row=1, max_row=1):
        for c in cell:
            c.value = None

    for row in ws.iter_rows(min_row=85, max_row=233, min_col=3, max_col=8):
        for cell in row:
            cell.value = None

    # Set column widths for "Extracted_Data" sheet
    for col, width in column_widths.items():
        new_ws.column_dimensions[col].width = width

    # Save the workbook
    wb.save(file_path)
