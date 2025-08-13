import pdfplumber
import pandas as pd

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
        print(f"PDF converted to Excel successfully: {excel_path}")
    else:
        print("No table data found in the PDF.")

# Example usage
pdf_path = "TN_MLA_21_MaduraiCentralPTR.pdf"  # Provide the PDF file path
excel_path = "TN_MLA_21_MaduraiCentralPTR.xlsx"  # Provide the output Excel file path
pdf_to_excel(pdf_path, excel_path)
