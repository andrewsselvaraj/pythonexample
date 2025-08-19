import pdfplumber
import pandas as pd
import os

# Function to extract text from PDF and save it as Excel
def pdf_to_excel(pdf_path, excel_path):
    data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                data.extend(table)  # Append extracted table data
    
    if data:
        df = pd.DataFrame(data)
        df.to_excel(excel_path, index=False, header=False)
        print(f"Converted: {pdf_path} → {excel_path}")
    else:
        print(f"No table data found in: {pdf_path}")

# Folder containing PDFs
input_folder = r"D:\andrew\python\pythonexample\pdftoexcel\research\t4"
output_folder = os.path.join(input_folder, "excel_output")

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Loop through all PDF files in the folder
for file_name in os.listdir(input_folder):
    if file_name.lower().endswith(".pdf"):
        pdf_path = os.path.join(input_folder, file_name)
        excel_name = os.path.splitext(file_name)[0] + ".xlsx"
        excel_path = os.path.join(output_folder, excel_name)
        
        pdf_to_excel(pdf_path, excel_path)

print("\n✅ All PDFs processed.")
