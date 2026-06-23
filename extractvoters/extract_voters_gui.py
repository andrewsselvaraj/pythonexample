import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import sys

# Import the main extraction logic from extract_voters_v2.py
import importlib.util
spec = importlib.util.spec_from_file_location("extract_voters_v2", os.path.join(os.path.dirname(__file__), "extract_voters_v2.py"))
extract_voters = importlib.util.module_from_spec(spec)
sys.modules["extract_voters_v2"] = extract_voters
spec.loader.exec_module(extract_voters)

class VoterExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Tamil Nadu Voter PDF Extractor")
        self.selected_files = []

        self.frame = tk.Frame(root, padx=16, pady=16)
        self.frame.pack(fill=tk.BOTH, expand=True)

        self.select_btn = tk.Button(self.frame, text="Select PDF Files", command=self.select_files)
        self.select_btn.pack(fill=tk.X, pady=(0, 8))

        self.files_listbox = tk.Listbox(self.frame, selectmode=tk.EXTENDED, height=8)
        self.files_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.run_btn = tk.Button(self.frame, text="Extract Voters", command=self.run_extraction)
        self.run_btn.pack(fill=tk.X)

        self.status_label = tk.Label(self.frame, text="Ready", anchor="w")
        self.status_label.pack(fill=tk.X, pady=(8, 0))

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Electoral Roll PDFs",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if files:
            self.selected_files = list(files)
            self.files_listbox.delete(0, tk.END)
            for f in self.selected_files:
                self.files_listbox.insert(tk.END, f)

    def run_extraction(self):
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select one or more PDF files.")
            return
        self.status_label.config(text="Extracting... Please wait.")
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self._extract_worker, daemon=True).start()

    def _extract_worker(self):
        for pdf_path in self.selected_files:
            try:
                # Set the PDF_FILE and output paths for each file
                extract_voters.PDF_FILE = pdf_path
                base = os.path.splitext(os.path.basename(pdf_path))[0]
                # Try to extract constituency code and name from filename (e.g., 229-Kanniyakumari)
                import re
                m = re.search(r'(\d{2,4})[-_]?([\w\(\)\u0B80-\u0BFF]+)?', base)
                if m:
                    code = m.group(1)
                    name = m.group(2) or ''
                    # Clean up name: replace non-alphanum with underscore, strip underscores
                    name = re.sub(r'[^\w\u0B80-\u0BFF]+', '_', name).strip('_')
                    out_dir = os.path.join(os.path.dirname(pdf_path), f"output_{code}_{name}" if name else f"output_{code}")
                else:
                    out_dir = os.path.join(os.path.dirname(pdf_path), "output")
                os.makedirs(out_dir, exist_ok=True)
                extract_voters.CSV_FILE = os.path.join(out_dir, f"{base}_voters.csv")
                extract_voters.EXCEL_FILE = os.path.join(out_dir, f"{base}_voters.xlsx")
                extract_voters.main()
            except Exception as e:
                print(f"Error processing {pdf_path}: {e}")
        self.status_label.config(text="Done! Check output files in output_<constituency> folders.")
        self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = VoterExtractorGUI(root)
    root.mainloop()
