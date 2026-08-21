import logging
import os
from datetime import datetime
from pathlib import Path
import shutil
import re

import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import pdfplumber
from pypdf import PdfReader
from PIL import Image, ImageDraw, ImageFont

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
    global ocr_reader
    if ocr_reader is None and EASYOCR_AVAILABLE:
        ocr_reader = easyocr.Reader(['en'], gpu=False)
    return ocr_reader


# --------------------------------------------------
# 1. 電子印鑑（デーツスタンプ）の自動生成機能
# --------------------------------------------------
def create_inspection_stamp(name_top="河", name_bottom="本", output_path="temp_stamp.png"):
    size = 400
    padding = 10
    color_vermilion = "#FF4500"
    font_path = "C:/Windows/Fonts/msgothic.ttc"

    image = Image.new("RGBA", (size, size), (255, 255, 255, 0))
    draw = ImageDraw.Draw(image)

    left, top = padding, padding
    right, bottom = size - padding, size - padding
    center_x, center_y = size // 2, size // 2

    line_width = 8
    draw.ellipse([left, top, right, bottom], outline=color_vermilion, width=line_width)

    line_y1 = size * 0.32
    line_y2 = size * 0.68
    radius = (size - 2 * padding) / 2

    def get_x_for_y(y):
        y_offset = abs(y - center_y)
        if y_offset >= radius:
            return center_x
        return (radius**2 - y_offset**2)**0.5

    x_offset1 = get_x_for_y(line_y1)
    draw.line([center_x - x_offset1, line_y1, center_x + x_offset1, line_y1], fill=color_vermilion, width=line_width)

    x_offset2 = get_x_for_y(line_y2)
    draw.line([center_x - x_offset2, line_y2, center_x + x_offset2, line_y2], fill=color_vermilion, width=line_width)

    now = datetime.now()
    date_str = now.strftime("%Y.%#m.%#d") if os.name == 'nt' else now.strftime("%Y.%-m.%-d")

    try:
        font_large = ImageFont.truetype(font_path, 80)
        font_date = ImageFont.truetype(font_path, 70)
    except Exception:
        font_large = ImageFont.load_default()
        font_date = ImageFont.load_default()

    draw.text((center_x, line_y1 * 0.6), name_top, fill=color_vermilion, font=font_large, anchor="mm")
    draw.text((center_x, center_y), date_str, fill=color_vermilion, font=font_date, anchor="mm")
    draw.text((center_x, line_y2 + (size - padding - line_y2) * 0.45), name_bottom, fill=color_vermilion, font=font_large, anchor="mm")

    image.save(output_path)
    return output_path


# --------------------------------------------------
# 2. Excelデータ抽出機能
# --------------------------------------------------
def extract_excel_data(sheet):
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


# --------------------------------------------------
# 補助機能: 救済モード用クロップ抽出（幅指定可能）
# --------------------------------------------------
def extract_left_column_text(pdf_path: Path, ratio: float = 0.18) -> str:
    """指定した幅割合でPDFの左側エリアを垂直抽出する"""
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


# --------------------------------------------------
# 最終手段機能: pypdfium2 + easyocr による画像文字認識（型式＋単価抽出対応）
# --------------------------------------------------
def perform_ocr_rescue(pdf_path: Path, target_item_clean: str) -> tuple[bool, str]:
    """最終手段: PDFを画像化(pypdfium2)してOCRで型式および文字全体のテキストを取得する"""
    if not EASYOCR_AVAILABLE:
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
                pdf.close()
            except Exception:
                pass
    return is_item_found, all_ocr_text


# --------------------------------------------------
# 3. 単一フォルダ内の照合 ＆ 押印処理
# --------------------------------------------------
def process_folder(folder_path: Path, stamp_img_path: str) -> bool:
    """指定された1つのフォルダ内の Excel / PDF を照合・押印する"""
    excel_files = [f for f in folder_path.glob("*.xlsx") if not f.name.startswith("~$")]
    pdf_files = [f for f in folder_path.glob("*.pdf") if not f.name.startswith("~$")]

    if not excel_files or not pdf_files:
        print(f"⚠️ スキップ: {folder_path.name} (Excel または PDF が不足しています)")
        return False

    print(f"\n📂 フォルダ処理中: 【 {folder_path.name} 】")
    print(f"   (Excel {len(excel_files)} 件 / PDF {len(pdf_files)} 件)")

    all_stamped = True

    # 全PDFテキストおよび注釈の読み込み
    pdf_list = []
    for pdf_file in pdf_files:
        try:
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

            full_pdf_text = full_text + "\n" + annot_text + "\n"

            pdf_list.append({
                "path": pdf_file,
                "name": pdf_file.name,
                "text": full_pdf_text,
                "text_raw": full_pdf_text
            })
        except Exception as e:
            print(f"❌ PDF読み込みエラー ({pdf_file.name}): {e}")

    # 各Excelの照合処理
    for excel_file in excel_files:
        wb = None
        try:
            try:
                wb = openpyxl.load_workbook(excel_file, data_only=True)
            except Exception:
                wb = openpyxl.load_workbook(excel_file, data_only=True, read_only=True)

            sheet = wb.active
            
            ex_data = extract_excel_data(sheet)
            ex_po = ex_data["po"]
            ex_item = ex_data["item"]
            ex_qty = ex_data["qty"]
            ex_price_clean = ex_data["price"].replace(",", "").strip()

            print(f"   📊 Excelデータ: {excel_file.name}")
            print(f"      └ [PO: '{ex_po}'] | [型式: '{ex_item}'] | [数量: '{ex_qty}'] | [単価: '{ex_price_clean}']")

            # --- PDFの特定ロジック（頭ゼロ除去・アスタリスク除去柔軟マッチング） ---
            matched_pdf = None
            if ex_po:
                ex_po_lstrip = ex_po.lstrip('0')
                
                # 1. ファイル名で検索
                for pdf_data in pdf_list:
                    if ex_po in pdf_data["name"] or (ex_po_lstrip and ex_po_lstrip in pdf_data["name"]):
                        matched_pdf = pdf_data
                        break
                
                # 2. 本文テキストで検索（記号・頭ゼロ除去対応）
                if not matched_pdf:
                    for pdf_data in pdf_list:
                        pdf_text_clean_po = pdf_data["text"].replace("*", "").replace("|", "").replace(" ", "")
                        if (ex_po in pdf_data["text"] or 
                            ex_po in pdf_text_clean_po or 
                            (ex_po_lstrip and ex_po_lstrip in pdf_text_clean_po)):
                            matched_pdf = pdf_data
                            break
            else:
                if pdf_list:
                    matched_pdf = pdf_list[0]
                    print("      ℹ️ 注文番号がないため、フォルダ内の最初のPDFと照合します。")

            if matched_pdf:
                pdf_text_raw = matched_pdf["text_raw"]
                pdf_text_no_comma = pdf_text_raw.replace(",", "")
                ocr_extracted_text = ""  # OCR実行時に取得した全文を保持

                # --------------------------------------------------
                # 1. 型式チェック（4段階判定：標準 ➔ 18%クロップ ➔ 35%クロップ ➔ 最終手段OCR）
                # --------------------------------------------------
                ex_item_clean = ex_item.strip().replace(" ", "").replace(" ", "")
                pdf_text_no_newline = (
                    pdf_text_raw.replace("\n", "")
                                .replace("\r", "")
                                .replace("\t", "")
                                .replace(" ", "")
                                .replace(" ", "")
                )
                
                ex_item_norm = ex_item_clean.replace("-", "").upper()
                pdf_text_norm = pdf_text_no_newline.replace("-", "").upper()
                
                # 英数字のみに完全修飾（記号ブレ吸収）
                ex_item_alphanumeric = re.sub(r'[^A-Za-z0-9]', '', ex_item_clean).upper()
                pdf_text_alphanumeric = re.sub(r'[^A-Za-z0-9]', '', pdf_text_no_newline).upper()

                # 第1段階: 標準チェック
                check_item = bool(
                    ex_item_clean and (
                        ex_item_clean in pdf_text_no_newline or 
                        f"K-{ex_item_clean}" in pdf_text_no_newline or
                        ex_item_clean in pdf_text_no_newline.replace("K-", "") or
                        ex_item_norm in pdf_text_norm or
                        ex_item_alphanumeric in pdf_text_alphanumeric
                    )
                )

                # 第2・3段階: 救済クロップ（18% / 35%）
                if not check_item and ex_item_clean:
                    print("      └ ⚠️ 通常抽出で型式が不一致のため、救済モード（ブロック位置解析）を発動します...")
                    
                    # 2段階目: MCL用（安全な18%幅）
                    left_text_18 = extract_left_column_text(matched_pdf["path"], ratio=0.18)
                    clean_18 = left_text_18.replace("\n", "").replace("\r", "").replace("\t", "").replace(" ", "").replace(" ", "")
                    norm_18 = clean_18.replace("-", "").upper()
                    alpha_18 = re.sub(r'[^A-Za-z0-9]', '', clean_18).upper()

                    if (ex_item_clean in clean_18 or f"K-{ex_item_clean}" in clean_18 or ex_item_clean in clean_18.replace("K-", "") or ex_item_norm in norm_18 or ex_item_alphanumeric in alpha_18):
                        check_item = True
                        print("      └ 🌸 救済モード(18%幅)により型式ブロックを正常検出！")
                    else:
                        # 3段階目: 通常レイアウト用（広めの35%幅）
                        left_text_35 = extract_left_column_text(matched_pdf["path"], ratio=0.35)
                        clean_35 = left_text_35.replace("\n", "").replace("\r", "").replace("\t", "").replace(" ", "").replace(" ", "")
                        norm_35 = clean_35.replace("-", "").upper()
                        alpha_35 = re.sub(r'[^A-Za-z0-9]', '', clean_35).upper()

                        if (ex_item_clean in clean_35 or f"K-{ex_item_clean}" in clean_35 or ex_item_clean in clean_35.replace("K-", "") or ex_item_norm in norm_35 or ex_item_alphanumeric in alpha_35):
                            check_item = True
                            print("      └ 🌸 救済モード(35%幅)により型式ブロックを正常検出！")

                # 第4段階: 【最終手段】OCRスキャン解析（テキスト層が存在しない図面PDF用）
                if not check_item and ex_item_clean:
                    print("      └ 🔍 最終手段: OCR（画像文字認識）スキャン解析を起動中...")
                    is_ocr_ok, ocr_text_res = perform_ocr_rescue(matched_pdf["path"], ex_item_clean)
                    ocr_extracted_text = ocr_text_res
                    if is_ocr_ok:
                        check_item = True
                        print("      └ 🌸 OCRスキャン解析により図面内の型式文字を正常検出！")

                # --------------------------------------------------
                # 2. 数量チェック
                # --------------------------------------------------
                check_qty = bool(ex_qty and (ex_qty in pdf_text_no_comma or ex_qty in ocr_extracted_text))

                # --------------------------------------------------
                # 3. 単価チェック（空欄/特定型式免除 ＋ カンマ・円記号・OCR対応）
                # --------------------------------------------------
                if ex_price_clean == "" or ex_item_clean == "57-04-322":
                    check_price = True
                else:
                    try:
                        price_val = float(ex_price_clean)

                        patterns = [
                            ex_price_clean,                                                           # 例: 10850 / 11550
                            f"{price_val:,.2f}",                                                      # 例: 10,850.00
                            f"{price_val:,.0f}",                                                      # 例: 10,850 / 11,550
                            f"¥{price_val:,.0f}", f"￥{price_val:,.0f}",                              # 例: ¥11,550 / ￥11,550
                            f"JPY{price_val:,.0f}", f"JPY {price_val:,.0f}",                          # 例: JPY 11,550
                            f"{price_val:.0f}",                                                       # 例: 10850
                            f"{price_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), # 例: 1.922,20
                            f"{price_val:.2f}".replace(".", ","),                                     
                            ex_price_clean.replace(".", ",")
                        ]

                        # 生テキスト、カンマなしテキスト、OCRテキストの全方向から照合
                        check_price = (
                            any(p in pdf_text_raw for p in patterns) or 
                            any(p in pdf_text_no_comma for p in patterns) or
                            any(p in ocr_extracted_text for p in patterns)
                        )

                    except ValueError:
                        check_price = (ex_price_clean in pdf_text_raw) or (ex_price_clean in pdf_text_no_comma) or (ex_price_clean in ocr_extracted_text)

                # PDF側データの検出状況ログ出力
                po_text_check = matched_pdf["text"].replace("*", "").replace("|", "").replace(" ", "")
                po_name_check = matched_pdf["name"].replace(" ", "")
                ex_po_lstrip = ex_po.lstrip('0') if ex_po else ""
                
                check_po_ok = bool(
                    ex_po and (
                        ex_po in po_text_check or 
                        ex_po in po_name_check or 
                        (ex_po_lstrip and (ex_po_lstrip in po_text_check or ex_po_lstrip in po_name_check))
                    )
                )
                po_log = f"'{ex_po}'" if check_po_ok else ("なし" if not ex_po else "未検出")
                
                item_log = f"'{ex_item}'" if check_item else "未検出"
                qty_log = f"'{ex_qty}'" if check_qty else "未検出"
                price_log = f"'{ex_price_clean}'" if check_price else ("免除(OK)" if (ex_price_clean == "" or ex_item_clean == "57-04-322") else "未検出")

                print(f"   📄 PDF検出結果: {matched_pdf['name']}")
                print(f"      └ [PO: {po_log}] | [型式: {item_log}] | [数量: {qty_log}] | [単価: {price_log}]")

                results = [
                    f"型式: {'OK' if check_item else 'NG'}",
                    f"数量: {'OK' if check_qty else 'NG'}",
                    f"単価: {'OK' if check_price else 'NG'}"
                ]
                print(f"      🔍 照合結果: { ' | '.join(results) }")

                if check_item and check_qty and check_price:
                    if getattr(wb, 'read_only', False):
                        wb.close()
                        wb = openpyxl.load_workbook(excel_file)
                        sheet = wb.active

                    img = OpenpyxlImage(stamp_img_path)
                    img.width, img.height = 75, 75
                    sheet.add_image(img, "T7")
                    wb.save(excel_file)
                    print(f"      🌸 3項目一致 ➔ スタンプ押印完了 (T7)")
                else:
                    print("      ⚠️ 不一致項目があるため押印スキップ")
                    all_stamped = False

            else:
                print(f"      ❌ 対応する注文書PDFが見つかりませんでした。")
                all_stamped = False

        except Exception as e:
            print(f"❌ Excelエラー ({excel_file.name}): {e}")
            all_stamped = False
        finally:
            if wb and hasattr(wb, 'close'):
                wb.close()

    return all_stamped


# --------------------------------------------------
# 4. 全子フォルダ巡回コントローラー
# --------------------------------------------------
def process_all_subfolders(parent_dir: Path, stamp_top="河", stamp_bottom="本"):
    """親フォルダ内のすべての子フォルダを取得し、順番に処理を実行する"""
    if not parent_dir.exists():
        print(f"❌ エラー: 親フォルダ '{parent_dir}' が存在しません。")
        return

    approved_dir = parent_dir / "approved"
    approved_dir.mkdir(exist_ok=True)

    subfolders = [f for f in parent_dir.iterdir() if f.is_dir() and f.name != "approved"]

    if not subfolders:
        print(f"⚠️ '{parent_dir}' 内に処理対象の子フォルダが見つかりませんでした。")
        return

    print("=" * 60)
    print(f"🚀 一括処理を開始します: 全 {len(subfolders)} 個のフォルダを検出")
    print("=" * 60)

    stamp_img_path = create_inspection_stamp(stamp_top, stamp_bottom)

    for folder in subfolders:
        is_success = process_folder(folder, stamp_img_path)
        
        if is_success:
            target_path = approved_dir / folder.name
            
            if target_path.exists():
                shutil.rmtree(target_path)
                
            shutil.move(str(folder), str(approved_dir))
            print(f"   🚚 押印完了のため '{approved_dir.name}/{folder.name}' へ移動しました。")

    print("\n" + "=" * 60)
    print("✨ すべての子フォルダに対する一括処理が完了しました。")


# --------------------------------------------------
# 実行ブロック
# --------------------------------------------------
if __name__ == "__main__":
    parent_directory = Path("./files")
    process_all_subfolders(parent_directory, stamp_top="河", stamp_bottom="本")