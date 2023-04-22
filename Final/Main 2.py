#import openpyxl
import openpyxl
#Open the file - Only need to chnage the file
workbook =openpyxl.load_workbook('/Users/hanaarshadahmed/Desktop/LGA/LGA_NM.xlsx')

#Code to delete rows
sheet = workbook['Local Government Area']
rows_to_delete = [1,2,3,8126,8127,8128,8129,8130,8131,8132,8133]
for row_index in sorted(rows_to_delete, reverse=True):
    sheet.delete_rows(row_index)
sheet.delete_cols(8,3)

#Save the updates in this file - only change file
workbook.save('/Users/hanaarshadahmed/Desktop/LGA/Final_Trendd')

import pandas as pd
'File 1'
location = pd.ExcelFile('/Users/hanaarshadahmed/Desktop/LGA/Final_Trendd')
df = pd.read_excel(location)

# Remove betting and gaming
df = df.loc[~(df['Offence type'] == 'Betting and gaming offences')]

#Update
output = '/Users/hanaarshadahmed/Desktop/LGA/Final_Trend2_.xlsx'
df.to_excel(output, index = False)
data = pd.read_excel(output)