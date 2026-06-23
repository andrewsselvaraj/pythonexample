import os, re
os.environ['TESSDATA_PREFIX'] = r'd:\andrew\election\tessdata'
import fitz, pytesseract
from PIL import Image

RE_BLOCK = re.compile(r'^[#\s]*(\d{1,4})\s+([A-Z0-9]{8,14})\s*$|^[#\s]*(\d{1,4})\s*$')

PDF = r'd:\andrew\election\voterlist\2026-EROLLGEN-S22-229-SIR-FinalRoll-Revision1-TAM-7-WI.pdf'
doc = fitz.open(PDF)

page = doc[4]  # page 5
scale = 200/72.0
mat = fitz.Matrix(scale, scale)
pix = page.get_pixmap(matrix=mat)
img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
w, h = img.width, img.height

all_misses = []
for pg_idx in range(doc.page_count):
    page = doc[pg_idx]
    scale = 200/72.0
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    w, h = img.width, img.height

    for c in range(3):
        x0 = c * w // 3
        x1 = (c + 1) * w // 3 if c < 2 else w
        col = img.crop((x0, 0, x1, h))
        text = pytesseract.image_to_string(col, lang='tam+eng', config='--psm 4 --oem 1')
        for l in text.split('\n'):
            s = l.strip()
            if not s:
                continue
            if re.match(r'^\d', s) and not RE_BLOCK.match(s) and len(s) < 50:
                all_misses.append(f"P{pg_idx+1}C{c+1}: {repr(s)}")

print(f"Total missed digit-lines: {len(all_misses)}")
for m in all_misses:
    print(m)
