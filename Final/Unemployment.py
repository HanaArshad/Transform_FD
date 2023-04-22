#import openpyxl
import openpyxl
#Open the file - Only need to chnage the file
workbook =openpyxl.load_workbook('/Users/hanaarshadahmed/Desktop/LGA/Unemployment20.xlsx')

#Code to delete sheet
del workbook['Smoothed LGA unemployment']
del workbook['Smoothed LGA labour force']

#Code to delete rows
sheet = workbook['Smoothed LGA unemployment rates']
rows_to_delete = [1, 2, 3]
for row_index in sorted(rows_to_delete, reverse=True):
    sheet.delete_rows(row_index)
#Code to delete columns
sheet.delete_cols(2,22)

#Save the updates in this file - only change file
workbook.save('/Users/hanaarshadahmed/Desktop/LGA/Final_Unemployment_2022')


