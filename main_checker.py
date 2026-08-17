import os
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image, ImageDraw, ImageFont
import pdfplumber
import logging
logging.getLogger("pdfminer").setLevel(logging.ERROR)  # pdfminerの警告メッセージを無視する


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
def process_folder(folder_path: Path, stamp_img_path: str):
    """指定された1つのフォルダ内の Excel / PDF を照合・押印する"""
    excel_files = [f for f in folder_path.glob("*.xlsx") if not f.name.startswith("~$")]
    pdf_files = [f for f in folder_path.glob("*.pdf") if not f.name.startswith("~$")]

    if not excel_files or not pdf_files:
        print(f"⚠️ スキップ: {folder_path.name} (Excel または PDF が不足しています)")
        return

    print(f"\n📂 フォルダ処理中: 【 {folder_path.name} 】")
    print(f"   (Excel {len(excel_files)} 件 / PDF {len(pdf_files)} 件)")

    # 全PDFテキストの読み込み
    pdf_list = []
    for pdf_file in pdf_files:
        try:
            with pdfplumber.open(pdf_file) as pdf:
                full_text = ""
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        full_text += text + "\n"
                
                pdf_list.append({
                    "name": pdf_file.name,
                    "text": full_text,
                    "text_no_comma": full_text.replace(",", "")
                })
        except Exception as e:
            print(f"❌ PDF読み込みエラー ({pdf_file.name}): {e}")

    # 各Excelの照合処理
    for excel_file in excel_files:
        try:
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            sheet = wb.active
            
            ex_data = extract_excel_data(sheet)
            ex_po = ex_data["po"]
            ex_item = ex_data["item"]
            ex_qty = ex_data["qty"]
            ex_price_clean = ex_data["price"].replace(",", "").strip()

            print(f"   📊 Excel: {excel_file.name} (PO: '{ex_po}')")

            if not ex_po:
                print("      ⚠️ Excel内に注文番号 (D11) がありません。スキップします。")
                continue

            matched_pdf = None
            for pdf_data in pdf_list:
                if ex_po in pdf_data["text"]:
                    matched_pdf = pdf_data
                    break

            if matched_pdf:
                # 照合用：カンマを除去したPDFテキスト
                pdf_text_clean = matched_pdf["text_no_comma"]

                # --------------------------------------------------
                # 1. 型式チェック（改行とスペースを除去して包含チェック）
                # --------------------------------------------------
                ex_item_clean = ex_item.strip().replace(" ", "").replace("　", "")
                
                # PDFテキストから改行(\n)・復帰(\r)・スペース(半角・全角)を除去
                pdf_text_no_newline = (
                    pdf_text_clean.replace("\n", "")
                                  .replace("\r", "")
                                  .replace(" ", "")
                                  .replace(" ", "")
                )
                
                # Excelの型式がPDF内にそのまま含まれているか判定
                check_item = bool(ex_item_clean and (ex_item_clean in pdf_text_no_newline))

                # --------------------------------------------------
                # 2. 数量チェック
                # --------------------------------------------------
                check_qty = bool(ex_qty and ex_qty in pdf_text_clean)

                # --------------------------------------------------
                # 3. 単価チェック（整数パターン と .00 パターンの両方を吸収）
                # --------------------------------------------------
                check_price = bool(ex_price_clean and (
                    ex_price_clean in pdf_text_clean or 
                    ex_price_clean in pdf_text_no_newline or 
                    f"{ex_price_clean}.00" in pdf_text_clean
                ))
                

                results = [
                    f"型式: {'OK' if check_item else 'NG'}",
                    f"数量: {'OK' if check_qty else 'NG'}",
                    f"単価: {'OK' if check_price else 'NG'}"
                ]
                print(f"      🔍 結果: { ' | '.join(results) } (PDF: {matched_pdf['name']})")

                # 型式・数量・単価の3項目 OK でスタンプ挿入
                if check_item and check_qty and check_price:
                    img = OpenpyxlImage(stamp_img_path)
                    img.width, img.height = 75, 75
                    sheet.add_image(img, "T7")
                    wb.save(excel_file)
                    print(f"      🌸 3項目一致 ➔ スタンプ押印完了 (T7)")
                else:
                    print("      ⚠️ 不一致項目があるため押印スキップ")

            else:
                print(f"      ❌ 対応する注文書PDFが見つかりませんでした。")

        except Exception as e:
            print(f"❌ Excelエラー ({excel_file.name}): {e}")


# --------------------------------------------------
# 4. 全子フォルダ巡回コントローラー
# --------------------------------------------------
def process_all_subfolders(parent_dir: Path, stamp_top="河", stamp_bottom="本"):
    """親フォルダ内のすべての子フォルダを取得し、順番に処理を実行する"""
    if not parent_dir.exists():
        print(f"❌ エラー: 親フォルダ '{parent_dir}' が存在しません。")
        return

    # 親フォルダ直下にある「ディレクトリ（フォルダ）」のみを抽出
    subfolders = [f for f in parent_dir.iterdir() if f.is_dir()]

    if not subfolders:
        print(f"⚠️ '{parent_dir}' 内に子フォルダが見つかりませんでした。")
        return

    print("=" * 60)
    print(f"🚀 一括処理を開始します: 全 {len(subfolders)} 個のフォルダを検出")
    print("=" * 60)

    # 共通の印鑑画像を1枚だけ事前生成
    stamp_img_path = create_inspection_stamp(stamp_top, stamp_bottom)

    for folder in subfolders:
        process_folder(folder, stamp_img_path)

    print("\n" + "=" * 60)
    print("✨ すべての子フォルダに対する一括処理が完了しました。")


# --------------------------------------------------
# 実行ブロック
# --------------------------------------------------
if __name__ == "__main__":
    # 親フォルダのパスを指定（例: "./files" や "./customer_orders"）
    parent_directory = Path("./files")
    
    # 実行（親フォルダ内の全子フォルダを連続処理）
    process_all_subfolders(parent_directory, stamp_top="河", stamp_bottom="本")