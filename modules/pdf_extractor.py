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
    """最終手段: 1ページ目を軽量化画像(scale=1.5)で高速OCR解析する"""
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

        # --------------------------------------------------
        # 1. 注文書ページ（1ページ目など図面以外）のみ取得
        # --------------------------------------------------
        target_indices = []
        try:
            with pdfplumber.open(pdf_path) as plumb_pdf:
                for idx, page in enumerate(plumb_pdf.pages):
                    page_text = page.extract_text() or ""
                    if "This Drawing is the property of Kyowa" not in page_text:
                        target_indices.append(idx)
        except Exception:
            pass

        if not target_indices:
            target_indices = [0]  # フォールバック

        # --------------------------------------------------
        # 2. 高精度OCR処理 (scale=3.0 で文字認識の精度を最優先)
        # --------------------------------------------------
        for idx in target_indices:
            page = pdf[idx]
            
            # ★ scale=3.0 ＋ グレースケールで微小な文字（4とAなど）の潰れを完全防止
            pil_image = page.render(scale=3.0).to_pil().convert('L')
            img_np = np.array(pil_image)
            
            results = reader.readtext(img_np, detail=0)
            page_ocr_text = " ".join(results)
            all_ocr_text += page_ocr_text + " "

            # 基本的な英数字正規化
            # perform_ocr_rescue 内の文字正規化と照合部分

            ocr_norm = re.sub(r'[^A-Za-z0-9]', '', page_ocr_text).upper()

            # 表記揺れ（1/I/O/0、記号潰れ '?'）を完全に統一
            ocr_norm_replaced = (
                ocr_norm.replace("OX", "10X")      # 'OX' ➔ '10X'
                        .replace("1OX", "10X")     # '1OX' ➔ '10X'
                        .replace("IOX", "10X")     # 'IOX' ➔ '10X'
                        .replace("OX2", "10X2")    # 'OX2' ➔ '10X2'
            )

            target_norm_replaced = (
                target_norm.replace("IOX", "10X")
                        .replace("1OX", "10X")
            )

            if (target_norm in ocr_norm) or (target_norm_replaced in ocr_norm_replaced):
                is_item_found = True
                break

    except Exception as e:
        print(f"      └ ⚠️ OCR処理例外: {e}")
    finally:
        if pdf is not None:
            try:
                pdf.close()
            except Exception:
                pass

    return is_item_found, all_ocr_text