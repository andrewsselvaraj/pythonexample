import pdfplumber
import pandas as pd
import os

def pdf_to_excel(pdf_path, excel_path):
    all_tables = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=2):
            print(f"⚠️ PATH {pdf_path}, skipping...")
            table = page.extract_table()
            
            if table:  # ✅ Only process if a table exists
                df = pd.DataFrame(table[1:], columns=table[0])  # first row as header
                df["Page"] = page_num  # optional: keep page number
                all_tables.append(df)
            else:
                print(f"⚠️ No table found on page {page_num}, skipping...")
    
    if all_tables:
        final_df = pd.concat(all_tables, ignore_index=True)
        final_df.to_excel(excel_path, index=False)
        print(f"✅ Tables extracted and saved to {excel_path}")
    else:
        print("❌ No tables found in the entire PDF.")

# Example usage
pdf_folder = r"D:\andrew\python\pythonexample\pdftoexcel\research\t4\complexpdf"
excel_output = r"D:\andrew\python\pythonexample\pdftoexcel\research\t4\complexpdf\output.xlsx"

# Convert all PDFs in folder
for file in os.listdir(pdf_folder):
    if file.endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, file)
        pdf_to_excel(pdf_path, excel_output)
