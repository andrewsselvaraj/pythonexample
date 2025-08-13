import os
import pdfplumber
import pandas as pd
import traceback

# Folder containing PDF files
folder_path = r"cdD:\andrew\git\63.142.240.31\html\constituency_display\polling_station_files\excel"
error_log_path = os.path.join(folder_path, "error_log.txt")

# Ensure the folder exists
if not os.path.exists(folder_path):
    print("Folder does not exist.")
    exit()

# Open error log file
with open(error_log_path, "w") as error_log:
    error_log.write("PDF to Excel Conversion Error Log\n")
    error_log.write("=" * 50 + "\n")

    # Process all PDF files in the folder
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            excel_path = os.path.join(folder_path, filename.replace(".pdf", ".xlsx"))

            try:
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

                # Convert extracted data to DataFrame and save as Excel
                if data:
                    df = pd.DataFrame(data)
                    df.to_excel(excel_path, index=False, header=False)
                    print(f"✅ Converted: {filename} -> {os.path.basename(excel_path)}")
                else:
                    print(f"⚠️ Skipping (No data found): {filename}")

            except Exception as e:
                error_message = f"❌ Error processing {filename}: {str(e)}"
                print(error_message)
                error_log.write(error_message + "\n")
                error_log.write(traceback.format_exc() + "\n")
                error_log.write("=" * 50 + "\n")

print("\n✅ Conversion completed! Check 'error_log.txt' for any issues.")
