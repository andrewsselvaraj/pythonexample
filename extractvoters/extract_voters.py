# This script extracts voter details from the given electoral roll PDF.
# Output: CSV with Serial No, Name, Father's Name, House No, Sex, Voter ID
# Place this script in the same folder as the PDF and run it.

import pdfplumber
import re
import csv


PDF_FILE = 'd:/andrew/election/voterlist/2026-EROLLGEN-S22-229-SIR-FinalRoll-Revision1-TAM-7-WI.pdf'
OUTPUT_CSV = 'd:/andrew/election/voterlist/S22-229-extract/voters.csv'

# Regex patterns (may need adjustment based on actual PDF text layout)
serial_re = re.compile(r'^(\d{1,4})\s')
voterid_re = re.compile(r'([A-Z]{3}[0-9]{7,})')
sex_re = re.compile(r'(Male|Female|Transgender)', re.IGNORECASE)
house_re = re.compile(r'House No\.?\s*:?\s*(\S+)')
partno_re = re.compile(r'Part\s*No\.?\s*:?\s*(\d+)')
const_re = re.compile(r'Constituency\s*:?\s*(.+)', re.IGNORECASE)
prev_en_re = re.compile(r'Prev\.?\s*EN\s*:?\s*(\S+)')
address_re = re.compile(r'Address\s*:?\s*(.+)', re.IGNORECASE)

rows = []

with pdfplumber.open(PDF_FILE) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        if not text:
            continue
        lines = text.split('\n')
        # Try to extract page-level info (constituency, part no, etc.)
        constituency = ''
        part_no = ''
        prev_en = ''
        address = ''
        for l in lines[:20]:
            if not constituency:
                m = const_re.search(l)
                if m:
                    constituency = m.group(1).strip()
            if not part_no:
                m = partno_re.search(l)
                if m:
                    part_no = m.group(1).strip()
            if not prev_en:
                m = prev_en_re.search(l)
                if m:
                    prev_en = m.group(1).strip()
            if not address:
                m = address_re.search(l)
                if m:
                    address = m.group(1).strip()

        i = 0
        while i < len(lines):
            line = lines[i].strip()
            # Try to find a serial number line
            m = serial_re.match(line)
            if m:
                serial = m.group(1)
                # Next lines: Name, Father's Name, House No, Sex, Voter ID
                name, father, house, sex, voterid = '', '', '', '', ''
                # Name (usually next line)
                if i+1 < len(lines):
                    name = lines[i+1].strip()
                # Father's Name (look for 'Father' or 'Husband')
                for j in range(i+2, min(i+6, len(lines))):
                    if 'Father' in lines[j] or 'Husband' in lines[j]:
                        father = lines[j].split(':',1)[-1].strip()
                    if house_re.search(lines[j]):
                        house = house_re.search(lines[j]).group(1)
                    if sex_re.search(lines[j]):
                        sex = sex_re.search(lines[j]).group(1)
                    if voterid_re.search(lines[j]):
                        voterid = voterid_re.search(lines[j]).group(1)
                rows.append([
                    constituency, part_no, prev_en, address,
                    serial, name, father, house, sex, voterid
                ])
            i += 1

with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow([
        'Constituency', 'Part No', 'Prev EN', 'Address',
        'Serial No', 'Name', "Father's Name", 'House No', 'Sex', 'Voter ID'
    ])
    writer.writerows(rows)

print(f"Extracted {len(rows)} voters to {OUTPUT_CSV}")
