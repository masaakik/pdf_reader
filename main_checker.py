import logging
import os
from datetime import datetime
from pathlib import Path
import shutil

import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import pdfplumber
from PIL import Image, ImageDraw, ImageFont

# pdfminerの警告メッセージを抑制
logging.getLogger("pdfminer").setLevel(logging.ERROR)


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
            with pdfplumber.open(pdf_file) as pdf:
                full_text = ""
                for page in pdf.pages:
                    # 1. レイアウトモードで本文テキスト抽出（列の文字割り込みを防止）
                    text = ""
                    try:
                        text = page.extract_text(layout=True) or ""
                    except Exception as e:
                        print(f"   ⚠️ 本文抽出警告 ({pdf_file.name}): {e}")

                    # 2. 注釈テキスト抽出（破損文字コード例外はスキップ）
                    annot_text = ""
                    try:
                        if hasattr(page, 'annots') and page.annots:
                            for annot in page.annots:
                                try:
                                    content = annot.get("contents") or annot.get("contents_pt") or ""
                                    if content:
                                        annot_text += f"\n{content}"
                                except Exception:
                                    continue
                    except Exception:
                        pass

                    full_text += text + "\n" + annot_text + "\n"

                pdf_list.append({
                    "name": pdf_file.name,
                    "text": full_text,
                    "text_no_comma": full_text.replace(",", "")
                })
        except Exception as e:
            print(f"❌ PDF読み込みエラー ({pdf_file.name}): {e}")

    # 各Excelの照合処理
    for excel_file in excel_files:
        wb = None
        try:
            # 特殊スタイルの読み込みエラー回避
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

            # --- PDFの特定ロジック ---
            matched_pdf = None
            if ex_po:
                for pdf_data in pdf_list:
                    if ex_po in pdf_data["text"]:
                        matched_pdf = pdf_data
                        break
            else:
                if pdf_list:
                    matched_pdf = pdf_list[0]
                    print("      ℹ️ 注文番号がないため、フォルダ内の最初のPDFと照合します。")

            if matched_pdf:
                pdf_text_clean = matched_pdf["text_no_comma"]

                # 1. 型式チェック（レイアウト維持抽出により1つのブロックとして厳密判定）
                ex_item_clean = ex_item.strip().replace(" ", "").replace(" ", "")
                pdf_text_no_newline = (
                    pdf_text_clean.replace("\n", "")
                                  .replace("\r", "")
                                  .replace("\t", "")
                                  .replace(" ", "")
                                  .replace(" ", "")
                )
                
                check_item = bool(
                    ex_item_clean and (
                        ex_item_clean in pdf_text_no_newline or 
                        f"K-{ex_item_clean}" in pdf_text_no_newline or
                        ex_item_clean in pdf_text_no_newline.replace("K-", "")
                    )
                )

                # 2. 数量チェック
                check_qty = bool(ex_qty and ex_qty in pdf_text_clean)

                # 3. 単価チェック
                check_price = bool(ex_price_clean and (
                    ex_price_clean in pdf_text_clean or 
                    ex_price_clean in pdf_text_no_newline or 
                    f"{ex_price_clean}.00" in pdf_text_clean
                ))

                # PDF側データの検出状況ログ出力
                po_log = f"'{ex_po}'" if (ex_po and ex_po in matched_pdf["text"]) else ("なし" if not ex_po else "未検出")
                item_log = f"'{ex_item}'" if check_item else "未検出"
                qty_log = f"'{ex_qty}'" if check_qty else "未検出"
                price_log = f"'{ex_price_clean}'" if check_price else "未検出"

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