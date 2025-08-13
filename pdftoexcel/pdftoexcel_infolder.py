import os
import pdfplumber
import pandas as pd

# Folder containing PDF files
folder_path = r"D:\andrew\git\63.142.240.31\html\constituency_display\polling_station_files\excel"

# Ensure the folder exists
if not os.path.exists(folder_path):
    print("Folder does not exist.")
    exit()

# Process all PDF files in the folder
for filename in os.listdir(folder_path):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(folder_path, filename)
        excel_path = os.path.join(folder_path, filename.replace(".pdf", ".xlsx"))

        data = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_table()
                if tables:
                    for table in tables:
                        data.extend(table)
                else:
                    text = page.extract_text()
                    if text:
                        data.append([text])

        # Convert extracted data to a DataFrame and save as Excel
        if data:
            df = pd.DataFrame(data)
            df.to_excel(excel_path, index=False, header=False)
            print(f"Converted: {filename} -> {os.path.basename(excel_path)}")
        else:
            print(f"Skipping (no data found): {filename}")

print("Conversion completed!")
