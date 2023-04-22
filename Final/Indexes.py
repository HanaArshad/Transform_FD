

import pandas as pd
# import os

# Load the Excel file into a pandas dataframe

sheets = pd.read_excel('/Users/hanaarshadahmed/Desktop/Indexes /LGA_Indexes.xls', sheet_name='Table 1')
rows_to_remove = [0, 1, 2, 3,4,551]

sheets.head()
var = sheets.columns
sheets = sheets.drop(sheets.columns[[0,3,5,7,9,10]], axis=1)

sheets = sheets.rename(columns={'Australian Bureau of Statistics ': '2016 Local Government Area (LGA) Code'})
sheets = sheets.rename(columns={'Unnamed: 1': '2016 Local Government Area (LGA) Name'})
sheets = sheets.rename(columns={'Unnamed: 2': 'Index of Relative Socio-economic Disadvantage'})
sheets = sheets.rename(columns={'Unnamed: 3': 'Decile'})
sheets = sheets.rename(columns={'Unnamed: 4': 'Index of Relative Socio-economic Advantage and Disadvantage'})
sheets = sheets.rename(columns={'Unnamed: 5': 'Decile'})
sheets = sheets.rename(columns={'Unnamed: 6': 'Index of Economic Resources'})
sheets = sheets.rename(columns={'Unnamed: 7': 'Decile'})
sheets = sheets.rename(columns={'Unnamed: 8': 'Index of Education and Occupation'})
sheets = sheets.rename(columns={'Unnamed: 9': 'Decile'})
sheets = sheets.rename(columns={'Unnamed: 10': 'Usual Resident Population'})


# for example, remove rows 0, 2, and 4
sheets = sheets.drop(rows_to_remove)
print(sheets)

sheets.to_excel('/Users/hanaarshadahmed/Desktop/Final_Indexes/Final.xlsx', index=False)