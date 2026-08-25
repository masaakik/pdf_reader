from pathlib import Path
import shutil
import openpyxl

# 自作モジュールのインポート
from modules.stamp_generator import create_inspection_stamp
from modules.excel_handler import extract_excel_data, apply_stamp_to_excel
from modules.pdf_extractor import extract_pdf_data
from modules.match_engine import match_po_number, verify_po_items

print("プログラムの実行を開始しました。")

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

    # 全PDFデータの抽出
    pdf_list = []
    for pdf_file in pdf_files:
        try:
            pdf_data = extract_pdf_data(pdf_file)
            pdf_list.append(pdf_data)
        except Exception as e:
            print(f"❌ PDF読み込みエラー ({pdf_file.name}): {e}")

    # 各Excelファイルの照合処理
    for excel_file in excel_files:
        wb = None
        try:
            # openpyxl による値取得
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

            # 該当するPDFの特定
            matched_pdf = match_po_number(ex_po, pdf_list)

            if matched_pdf:
                # 照合エンジンの実行
                match_res = verify_po_items(ex_data, matched_pdf)

                # ログ用テキスト構築
                po_log = f"'{ex_po}'" if match_res["po"] else ("なし" if not ex_po else "未検出")
                item_log = f"'{ex_item}'" if match_res["item"] else "未検出"
                qty_log = f"'{ex_qty}'" if match_res["qty"] else "未検出"
                
                is_price_exempt = (ex_price_clean == "" or ex_item.strip().replace(" ", "") == "57-04-322")
                price_log = f"'{ex_price_clean}'" if match_res["price"] else ("免除(OK)" if is_price_exempt else "未検出")

                # 1. 📄 PDF検出結果を先に出力
                print(f"   📄 PDF検出結果: {matched_pdf['name']}")
                print(f"      └ [PO: {po_log}] | [型式: {item_log}] | [数量: {qty_log}] | [単価: {price_log}]")

                # 2. ⚠️ 救済モードやOCR等の割り込みログがあればここに出力
                for log_msg in match_res.get("rescue_logs", []):
                    print(log_msg)

                # 3. 🔍 照合結果を出力
                results = [
                    f"型式: {'OK' if match_res['item'] else 'NG'}",
                    f"数量: {'OK' if match_res['qty'] else 'NG'}",
                    f"単価: {'OK' if match_res['price'] else 'NG'}"
                ]
                print(f"      🔍 照合結果: { ' | '.join(results) }")

                # 3項目（型式・数量・単価）一致でスタンプ押印
                if match_res["is_all_ok"]:
                    if wb and hasattr(wb, 'close'):
                        wb.close()  # 保存前に一旦クローズ

                    apply_stamp_to_excel(excel_file, stamp_img_path, cell_position="T7")
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


if __name__ == "__main__":
    parent_directory = Path("./files")
    process_all_subfolders(parent_directory, stamp_top="河", stamp_bottom="本")