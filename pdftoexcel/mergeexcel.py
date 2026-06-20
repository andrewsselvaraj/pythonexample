import pandas as pd

# Load the first sheet
sheet1 = pd.read_excel(r'D:\andrew\election\research\Alandur\ae_AC028_perfect.xlsx')

# Load the second sheet
sheet2 = pd.read_excel(r'D:\andrew\election\research\Alandur\pb_AC028.xlsx')

# Merge the two sheets based on the 'Sl.No' column
merged_sheet = pd.merge(sheet1, sheet2, on='BoothNo')

# Save the merged sheet to a new Excel file
merged_sheet.to_excel('alandur_merged_sheet.xlsx', index=False)

print("Sheets merged successfully!")