from pathlib import Path
import openpyxl
import pdfplumber

# ① Set up file paths
base_dir = Path(__file__).parent
excel_path = base_dir / "files" / "20261105_528733401.xlsx"
pdf_path = base_dir / "files" / "POHF01260700386 revised.pdf"

# ② Extract data from Excel (using openpyxl with cell addresses)
wb = openpyxl.load_workbook(excel_path)
sheet = wb.active

# Standard safety set: convert to string + strip leading/trailing spaces
excel_po = str(sheet["D11"].value).strip()       # P/O No.
excel_item = str(sheet["D20"].value).strip()     # Item / Specification
excel_qty = str(sheet["S20"].value).strip()      # Quantity (Merged cell: S20)
excel_price = str(sheet["D23"].value).strip()    # Unit Price

print("--- Extracted Data from Excel ---")
print(f"P/O No.   : {excel_po}")
print(f"Item      : {excel_item}")
print(f"Quantity  : {excel_qty}")
print(f"Unit Price: {excel_price}")

# ③ Extract full text from PDF
with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[0]
    pdf_text = page.extract_text()

# ④ Cross-check processing
print("\n--- Cross-check Results ---")

# Check P/O No.
if excel_po in pdf_text:
    print(f"⭕ P/O No.   : MATCH ({excel_po})")
else:
    print(f"❌ P/O No.   : MISMATCH ({excel_po})")

# Check Item Name
if excel_item in pdf_text:
    print(f"⭕ Item      : MATCH ({excel_item})")
else:
    print(f"❌ Item      : MISMATCH ({excel_item})")

# Check Quantity
if excel_qty in pdf_text:
    print(f"⭕ Quantity  : MATCH ({excel_qty})")
else:
    print(f"❌ Quantity  : MISMATCH ({excel_qty})")

# Check Unit Price
if excel_price in pdf_text:
    print(f"⭕ Unit Price: MATCH ({excel_price})")
else:
    print(f"❌ Unit Price: MISMATCH ({excel_price})")