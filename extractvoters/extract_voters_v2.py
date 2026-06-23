"""
Extract voter data from Tamil electoral roll PDF using OCR.
PDF: 2026-EROLLGEN-S22-229-SIR-FinalRoll-Revision1-TAM-7-WI.pdf
Output: voters.csv  (and  voters.xlsx)

Fields: Serial No, Voter ID, Name, Relation Type, Relation Name, House No, Age, Sex, Part No, Section
"""
import os
import re
import csv
import fitz          # pymupdf
import pytesseract
from PIL import Image
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

os.environ['TESSDATA_PREFIX'] = r'd:\andrew\election\tessdata'

PDF_FILE   = r'd:\andrew\election\voterlist\2026-EROLLGEN-S22-229-SIR-FinalRoll-Revision1-TAM-7-WI.pdf'
CSV_FILE   = r'd:\andrew\election\voterlist\S22-229-extract\voters_2026.csv'
EXCEL_FILE = r'd:\andrew\election\voterlist\S22-229-extract\voters_2026.xlsx'

DPI = 200   # 200 DPI works best for Tamil OCR

# Tamil field-label keywords (OCR may add ‌ ZWJ after consonants)
# Voter block start patterns to match:
#   - "32 2210978360 | |"  (trailing OCR noise)
#   - "254 2211829837."    (trailing dot)
#   - "765 2 2ZK3705530"   (serial gender voter_id on Page 31 form)
#   - "34 HFX1175306"      (normal)
#   - "211"                (serial only, voter ID on next line)
RE_BLOCK_START = re.compile(
    r'^[#\s]*(\d{1,4})\s+(\d)\s+([A-Z0-9]{9,14})[\s|.]*$|'    # serial gender voter_id
    r'^[#\s]*(\d{1,4})\s+([A-Z0-9]{8,14})[\s|.]*$|'           # serial voter_id (with trailing noise)
    r'^[#\s]*(\d{1,4})[\s.]*$'                                 # serial only
)
RE_VOTERID  = re.compile(r'\b([A-Z]{3}[0-9]{6,8})\b')
RE_VOTERID2 = re.compile(r'\b([A-Z0-9]{9,14})\b')   # fallback for OCR-corrupted IDs
RE_SERIAL   = re.compile(r'^\s*(\d{1,4})\s*$')
RE_NAME     = re.compile(r'பெயர்[\u200c\s]*:[\u200c\s]*(.*)')
RE_FATHER   = re.compile(r'தந்தை(?:யின்)?[\u200c\s]*பெயர்[\u200c\s]*:[\u200c\s]*(.*)')
RE_HUSBAND  = re.compile(r'கணவர்[\u200c\s]*பெயர்[\u200c\s]*:[\u200c\s]*(.*)')
RE_MOTHER   = re.compile(r'தாயின்[\u200c\s]*பெயர்[\u200c\s]*:[\u200c\s]*(.*)')
RE_HOUSE    = re.compile(r'வீட்டு[\u200c\s]*எண்[\u200c\s]*[:\./][\u200c\s]*(\S+)')
RE_AGE      = re.compile(r'வயது[\u200c\s]*[:\./][\u200c\s]*(\d+)')
RE_SEX      = re.compile(r'பாலினம்[\u200c\s]*[:\./][\u200c\s]*(ஆண்|பெண்|மூன்றாம்)', re.UNICODE)
RE_PARTNO   = re.compile(r'பாகம்[\u200c\s]*எண்[\u200c\s]*[:\./][\u200c\s]*(\d+)')
RE_SECTION  = re.compile(r'பிரிவு[\u200c\s]*எண்[\u200c\s]*மற்றும்[\u200c\s]*பெயர்[\u200c\s]*[:\./]?[\u200c\s]*(.+?)(?:\n|$)')
RE_CONST    = re.compile(r'(\d{3,})-(.+?)(?:\s+பாகம்|$)', re.UNICODE)

SEX_MAP = {'ஆண்': 'Male', 'பெண்': 'Female', 'மூன்றாம்': 'Third Gender'}


def render_page(page, dpi=DPI):
    """Render a PDF page to a PIL Image."""
    scale = dpi / 72.0
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img


def ocr_column(img_crop):
    """Run Tesseract OCR on a cropped column image, return plain text."""
    return pytesseract.image_to_string(
        img_crop, lang='tam+eng',
        config='--psm 4 --oem 1'  # psm 4 = single column, works best for electoral roll columns
    )


def split_image_into_columns(img, num_cols=3):
    """Crop a page image into num_cols vertical strips."""
    w, h = img.size
    col_w = w // num_cols
    cols = []
    for i in range(num_cols):
        x0 = i * col_w
        x1 = (i + 1) * col_w if i < num_cols - 1 else w
        cols.append(img.crop((x0, 0, x1, h)))
    return cols


def parse_voter_block(text):
    """
    Parse a block of Tamil OCR text for a single voter record.
    Returns dict with extracted fields.
    """
    # Serial number + voter ID from first line
    serial = ''
    voter_id = ''
    first_line = text.split('\n')[0].strip()
    m = RE_BLOCK_START.match(first_line)
    if m:
        # RE_BLOCK_START has 3 alternatives:
        #   groups (1,2,3): serial, gender, voter_id  (page 31 format)
        #   groups (4,5):   serial, voter_id          (normal format with trailing noise)
        #   group (6):      serial only
        if m.group(1):  # serial + gender + voter_id
            serial = m.group(1)
            voter_id = m.group(3) or ''
        elif m.group(4):  # serial + voter_id
            serial = m.group(4)
            voter_id = m.group(5) or ''
        elif m.group(6):  # serial only
            serial = m.group(6)

    # If voter ID not found in first line, search rest of block for proper ID
    if not voter_id:
        vid_m = RE_VOTERID.search(text)
        if vid_m:
            voter_id = vid_m.group(1)
        else:
            vid_m2 = RE_VOTERID2.search(text)
            if vid_m2:
                voter_id = vid_m2.group(1)

    # Name
    name = ''
    m = RE_NAME.search(text)
    if m:
        name = m.group(1).strip().replace('\u200c', '').strip()

    # Relation
    relation_type, relation_name = '', ''
    for pat, rtype in [(RE_FATHER, 'Father'), (RE_HUSBAND, 'Husband'), (RE_MOTHER, 'Mother')]:
        m = pat.search(text)
        if m:
            relation_type = rtype
            relation_name = m.group(1).strip().replace('\u200c', '').strip()
            break

    # House No
    house = ''
    m = RE_HOUSE.search(text)
    if m:
        house = m.group(1).strip()

    # Age
    age = ''
    m = RE_AGE.search(text)
    if m:
        age = m.group(1)

    # Sex
    sex = ''
    m = RE_SEX.search(text)
    if m:
        sex = SEX_MAP.get(m.group(1).strip().replace('\u200c', ''), m.group(1))

    # Only return a record if we got at least a serial number
    if not serial:
        return None

    return {
        'Serial No': serial,
        'Voter ID': voter_id,
        'Name': name,
        'Relation Type': relation_type,
        'Relation Name': relation_name,
        'House No': house,
        'Age': age,
        'Sex': sex,
    }


def parse_page_header(text):
    """Extract part no and section name from page header text."""
    part_no = ''
    section = ''
    m = RE_PARTNO.search(text)
    if m:
        part_no = m.group(1)
    m = RE_SECTION.search(text)
    if m:
        section = m.group(1).strip()
    return part_no, section


def split_column_into_voter_blocks(col_text):
    """
    Split a single column's OCR text into individual voter blocks.
    Each block starts with a line: [optional #] <serial_no> [<voter_id>]
    """
    lines = col_text.split('\n')
    blocks = []
    current = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if RE_BLOCK_START.match(stripped):
            if current:
                blocks.append('\n'.join(current))
            current = [line]
        else:
            current.append(line)
    if current:
        blocks.append('\n'.join(current))
    return blocks


def main():
    doc = fitz.open(PDF_FILE)
    total_pages = doc.page_count
    print(f"PDF: {total_pages} pages")

    all_voters = []
    current_part = ''
    current_section = ''

    for page_num in range(total_pages):
        print(f"Processing page {page_num + 1}/{total_pages}...", end=' ', flush=True)
        page = doc[page_num]
        img = render_page(page)

        # Split page into 3 column images and OCR each
        cols = split_image_into_columns(img, num_cols=3)
        page_count = 0

        # OCR all columns first
        col_texts = [ocr_column(col_img) for col_img in cols]
        combined_col_text = '\n'.join(col_texts)

        # Update part/section from header BEFORE assigning to voters
        part_no_found, section_found = parse_page_header(combined_col_text)
        if part_no_found:
            current_part = part_no_found
        if section_found:
            current_section = section_found

        # Now parse voter blocks
        for col_text in col_texts:
            blocks = split_column_into_voter_blocks(col_text)
            for block in blocks:
                voter = parse_voter_block(block)
                if voter:
                    voter['Part No'] = current_part
                    voter['Section'] = current_section
                    all_voters.append(voter)
                    page_count += 1

        if page_count == 0:
            print("(no voters)")
        else:
            print(f"{page_count} voters")

    print(f"\nTotal voters extracted: {len(all_voters)}")

    if not all_voters:
        print("No voters extracted. Check OCR output.")
        return


    # Add new columns and values
    constituency_val = '229-கன்னியாகுமரி'
    section_info_val = '1-அருமநல்லூர்(வ.கி), மற்றும் (ஊ), பிளாக் 1,2, வார்டு 3 செக்கடி'
    part_info_val = '7'

    # New column order: Serial No, Section, Constituency, Section Info, Voter ID, ...
    fieldnames = [
        'Serial No', 'Section', 'Constituency', 'Section Info',
        'Voter ID', 'Name', 'Relation Type', 'Relation Name',
        'House No', 'Age', 'Sex', 'Part No', 'Part Info'
    ]
    for voter in all_voters:
        voter['Constituency'] = constituency_val
        voter['Section Info'] = section_info_val
        voter['Part Info'] = part_info_val
        # Move Section to be after Serial No (already present)
        # All other fields are already present

    with open(CSV_FILE, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_voters)
    print(f"CSV saved: {CSV_FILE}")

    # Write Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Voters"

    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left   = Alignment(horizontal='left', vertical='center', wrap_text=True)
    thin   = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    alt_fill = PatternFill(start_color="EBF3FF", end_color="EBF3FF", fill_type="solid")

    ws.merge_cells(f'A1:{chr(64+len(fieldnames))}1')
    t = ws['A1']
    t.value = "Electoral Roll 2026 – Constituency 229 (Kanyakumari) | Voter List"
    t.font = Font(bold=True, size=11, color="FFFFFF")
    t.fill = hdr_fill
    t.alignment = center
    ws.row_dimensions[1].height = 22


    col_widths = [9, 35, 22, 50, 12, 28, 12, 28, 10, 6, 8, 8, 8]
    for i, (h, w) in enumerate(zip(fieldnames, col_widths), start=1):
        cell = ws.cell(row=2, column=i, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = center
        cell.border = thin
        ws.column_dimensions[cell.column_letter].width = w
    ws.row_dimensions[2].height = 30

    for row_i, voter in enumerate(all_voters, start=3):
        fill = alt_fill if (row_i % 2 == 0) else None
        for col_i, key in enumerate(fieldnames, start=1):
            cell = ws.cell(row=row_i, column=col_i, value=voter.get(key, ''))
            cell.border = thin
            if fill:
                cell.fill = fill
            # Left align for Name, Relation Name, Section, Section Info
            cell.alignment = left if col_i in (2, 4, 6, 8) else center
    ws.freeze_panes = 'A3'

    wb.save(EXCEL_FILE)
    print(f"Excel saved: {EXCEL_FILE}")


if __name__ == '__main__':
    main()
