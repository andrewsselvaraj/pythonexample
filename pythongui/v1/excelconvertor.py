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

        # Save as Excel
        excel_path = os.path.splitext(pdf_path)[0] + ".xlsx"
        final_df.to_excel(excel_path, index=False)

        messagebox.showinfo("Success", f"Converted successfully!\nSaved at: {excel_path}")

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
