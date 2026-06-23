import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import os

def select_file():
    file_path = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if file_path:
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, file_path)

def convert_pdf_to_excel():
    pdf_path = entry_pdf.get()
    if not pdf_path:
        messagebox.showerror("Error", "Please select a PDF file")
        return

    try:
        all_text = []
        # Read PDF with pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    all_text.append(df)

        if not all_text:
            messagebox.showwarning("No Table Found", "No tables detected in PDF")
            return

        # Merge all tables
        final_df = pd.concat(all_text, ignore_index=True)

        # ---- FIX: replace any comma inside cell values with a semicolon ----
        # This prevents data loss when the CSV is later read by tools that
        # do not honor quoted fields. Only affects values, not column structure.
        final_df = final_df.apply(
            lambda col: col.map(lambda v: v.replace(",", ";") if isinstance(v, str) else v)
        )

        base_path = os.path.splitext(pdf_path)[0]

        # Save as Excel
        excel_path = base_path + ".xlsx"
        final_df.to_excel(excel_path, index=False)

        # Save as CSV (UTF-8 with BOM so Tamil text opens correctly in Excel)
        csv_path = base_path + ".csv"
        final_df.to_csv(csv_path, index=False, encoding="utf-8-sig")

        messagebox.showinfo(
            "Success",
            f"Converted successfully!\nExcel: {excel_path}\nCSV: {csv_path}"
        )

    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert: {str(e)}")


# GUI Setup
root = tk.Tk()
root.title("PDF to Excel Converter")
root.geometry("500x200")

label_pdf = tk.Label(root, text="Select PDF File:")
label_pdf.pack(pady=5)

entry_pdf = tk.Entry(root, width=50)
entry_pdf.pack(pady=5)

btn_browse = tk.Button(root, text="Browse", command=select_file)
btn_browse.pack(pady=5)

btn_convert = tk.Button(root, text="Convert to Excel", command=convert_pdf_to_excel, bg="green", fg="white")
btn_convert.pack(pady=20)

root.mainloop()
