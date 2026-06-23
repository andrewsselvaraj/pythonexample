# Dump the raw text of the first page of the PDF for analysis
import pdfplumber

PDF_FILE = 'd:/andrew/election/voterlist/2026-EROLLGEN-S22-229-SIR-FinalRoll-Revision1-TAM-7-WI.pdf'
DUMP_FILE = 'd:/andrew/election/voterlist/S22-229-extract/first_page_text.txt'

with pdfplumber.open(PDF_FILE) as pdf:
    first_page = pdf.pages[0]
    text = first_page.extract_text()

with open(DUMP_FILE, 'w', encoding='utf-8') as f:
    f.write(text or '')

print(f"First page text dumped to {DUMP_FILE}")
