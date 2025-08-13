import pandas as pd

# Load the first sheet
sheet1 = pd.read_excel('230Nagercoil.xlsx')

# Load the second sheet
sheet2 = pd.read_excel('B230Nagercoil.xlsx')

# Merge the two sheets based on the 'Sl.No' column
merged_sheet = pd.merge(sheet1, sheet2, on='pollingstationo-')

# Save the merged sheet to a new Excel file
merged_sheet.to_excel('merged_sheet.xlsx', index=False)

print("Sheets merged successfully!")