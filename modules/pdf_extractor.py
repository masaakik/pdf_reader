import logging
import re
from pathlib import Path
import pdfplumber
from pypdf import PdfReader

# 最終手段用OCRライブラリ（遅延読み込み用フラグ）
EASYOCR_AVAILABLE = False
try:
    import easyocr
    import pypdfium2
    import numpy as np
    EASYOCR_AVAILABLE = True
except ImportError:
    EASYOCR_AVAILABLE = False

# pdfminerの警告メッセージを抑制
logging.getLogger("pdfminer").setLevel(logging.ERROR)

# EasyOCRリーダーのキャッシュ変数
ocr_reader = None

def get_ocr_reader():
    """EasyOCRリーダーのシングルトン取得"""
    global ocr_reader
    if ocr_reader is None and EASYOCR_AVAILABLE:
        ocr_reader = easyocr.Reader(['en'], gpu=False)
    return ocr_reader


def extract_pdf_data(pdf_file: Path) -> dict:
    """PDFファイルから本文テキストおよび注釈（アノテーション）テキストを一括抽出"""
    full_text = ""
    
    # 1. 本文テキスト抽出 (pdfplumber)
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                full_text += (page.extract_text() or "") + "\n"
    except Exception as e:
        print(f"   ⚠️ 本文抽出警告 ({pdf_file.name}): {e}")

    # 2. 注釈（上書き）テキスト抽出 (pypdf: 文字コードエラーを完全回避)
    annot_text = ""
    try:
        reader = PdfReader(pdf_file)
        for pypdf_page in reader.pages:
            if "/Annots" in pypdf_page:
                for annot in pypdf_page["/Annots"]:
                    try:
                        obj = annot.get_object()
                        contents = obj.get("/Contents")
                        if contents and isinstance(contents, str):
                            annot_text += f"\n{contents}"
                    except Exception:
                        continue
    except Exception as e:
        print(f"   ⚠️ 注釈抽出警告 ({pdf_file.name}): {e}")

    combined_text = full_text + "\n" + annot_text + "\n"

    return {
        "path": pdf_file,
        "name": pdf_file.name,
        "text": combined_text,
        "text_raw": combined_text
    }


def extract_left_column_text(pdf_path: Path, ratio: float = 0.18) -> str:
    """指定した幅割合でPDFの左側エリアを垂直抽出する（救済クロップ機能）"""
    extracted_text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                width = page.width
                height = page.height
                crop_box = (0, 0, width * ratio, height)
                
                cropped_page = page.crop(crop_box)
                text = cropped_page.extract_text() or ""
                extracted_text += text + "\n"
    except Exception:
        pass
    return extracted_text


def perform_ocr_rescue(pdf_path: Path, target_item_clean: str) -> tuple[bool, str]:
    """最終手段: PDFを画像化(pypdfium2)してOCRで型式文字を認識・照合する（ファイルロック解除対応）"""
    if not EASYOCR_AVAILABLE:
        print("      └ ⚠️ easyocr / pypdfium2 パッケージが未インストールのためOCRをスキップします。")
        return False, ""

    pdf = None
    all_ocr_text = ""
    is_item_found = False

    try:
        import pypdfium2 as pdfium
        import numpy as np

        reader = get_ocr_reader()
        if not reader:
            return False, ""

        pdf = pdfium.PdfDocument(pdf_path)
        target_norm = re.sub(r'[^A-Za-z0-9]', '', target_item_clean).upper()

        for page in pdf:
            pil_image = page.render(scale=3).to_pil()
            img_np = np.array(pil_image)
            
            results = reader.readtext(img_np, detail=0)
            page_ocr_text = " ".join(results)
            all_ocr_text += page_ocr_text + " "

            ocr_norm = re.sub(r'[^A-Za-z0-9]', '', page_ocr_text).upper()
            if target_norm in ocr_norm:
                is_item_found = True

    except Exception as e:
        print(f"      └ ⚠️ OCR処理例外: {e}")
    finally:
        if pdf is not None:
            try:
                pdf.close()  # 確実にファイルロックを解除
            except Exception:
                pass
    return is_item_found, all_ocr_text