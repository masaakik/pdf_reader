import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage

def extract_excel_data(sheet) -> dict:
    """Excelシートから固定・可変データを抽出"""
    data = {
        "po": str(sheet["D11"].value or "").strip(),     # PO Number
        "item": str(sheet["D20"].value or "").strip(),   # 型式
        "qty": str(sheet["S20"].value or "").strip(),    # 数量
        "price": ""                                      # 単価
    }
    
    for row in range(20, 30):
        cell_a = str(sheet[f"A{row}"].value or "").strip()
        if "単価" in cell_a:
            data["price"] = str(sheet[f"D{row}"].value or "").strip()
            break
            
    return data

def apply_stamp_to_excel(excel_file, stamp_img_path: str, cell_position: str = "T7"):
    """Excelファイルにスタンプ画像を貼り付けて保存"""
    wb = None
    try:
        wb = openpyxl.load_workbook(excel_file)
        sheet = wb.active

        img = OpenpyxlImage(stamp_img_path)
        img.width, img.height = 75, 75
        sheet.add_image(img, cell_position)
        wb.save(excel_file)
    finally:
        if wb and hasattr(wb, 'close'):
            wb.close()