
import pandas as pd
# import os

# Load the Excel file into a pandas dataframe
file_path = '/Users/hanaarshadahmed/Desktop/LGA/Benefits.csv'
df = pd.read_csv(file_path)
df = df.drop(df.columns[[0,1]], axis=1)
# Save the modified dataframe back to a new Excel file
output_file_path = '/Users/hanaarshadahmed/Desktop/LGA/Final_Benfit.csv'
df.to_csv(output_file_path, index=False)
# use pandas read_excel function to read the file
data = pd.read_csv(output_file_path)

