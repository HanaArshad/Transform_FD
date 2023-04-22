from pathlib import Path

input_dir = Path.cwd() / '/Users/hanaarshadahmed/Desktop/VV'
from openpyxl import load_workbook  # pip install openpyxl

for path in list(input_dir.rglob("*.xlsx*")):
    wb = load_workbook(filename=path)
    ws= wb["Victims"]
    del wb["Victims"]
    we = wb["Premises Type"]
    del wb["Premises Type"]
    wr = wb['Summary of offences']
    del wb['Summary of offences']

# Code to delete specific items in a sheet

    # Changes in Sheet 'Offenders'
    sheet = wb['Offenders']
    rows_to_delete = [1, 2, 3, 4, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46]
    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    # Changes in Sheet 'Aborginality'
    sheet = wb['Aboriginality']
    rows_to_delete = [1, 5, 11, 12, 13, 14, 15, 16, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34,
                      35, 36, 37]
    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    # Changes in Sheet 'Alcohol Related'
    sheet = wb['Alcohol Related']
    rows_to_delete = [1, 2, 3, 4, 26, 27, 28, 29, 30, 31, 32]
    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    # Changes in Sheet 'Month'
    sheet = wb['Month']
    rows_to_delete = [1, 2, 3, 4, 41, 42, 43, 44, 45, 46, 47]
    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    # Changes in Sheet 'Time'
    sheet = wb['Time']
    rows_to_delete = [1, 2, 3, 4, 41, 42, 43, 44, 45, 46, 47]
    for row_index in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_index)

    output_dir = Path.cwd() / '/Users/hanaarshadahmed/Desktop/E_LGA'
    output_dir.mkdir(exist_ok=True)
    wb.save(output_dir/path.name)
