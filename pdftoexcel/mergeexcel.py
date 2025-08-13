import pandas as pd

# Load the first sheet
sheet1 = pd.read_excel('193MaduraiCentral_tableau_STA.xlsx')

# Load the second sheet
sheet2 = pd.read_excel('193MaduraiCentral_tableau_PS.xlsx')

# Merge the two sheets based on the 'Sl.No' column
merged_sheet = pd.merge(sheet1, sheet2, on='ps_station')

# Save the merged sheet to a new Excel file
merged_sheet.to_excel('merged_sheet.xlsx', index=False)

print("Sheets merged successfully!")