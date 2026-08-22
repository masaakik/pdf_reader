import re
from pathlib import Path
from modules.pdf_extractor import extract_left_column_text, perform_ocr_rescue

def match_po_number(ex_po: str, pdf_list: list) -> dict | None:
    """ExcelのPO番号を元に、合致するPDFデータを特定する"""
    if not ex_po:
        return pdf_list[0] if pdf_list else None

    ex_po_lstrip = ex_po.lstrip('0')  # 例: '00022789' -> '22789'

    # 1. PDFファイル名で検索
    for pdf_data in pdf_list:
        if ex_po in pdf_data["name"] or (ex_po_lstrip and ex_po_lstrip in pdf_data["name"]):
            return pdf_data

    # 2. 本文テキストで検索（記号・頭ゼロ除去対応）
    for pdf_data in pdf_list:
        pdf_text_clean_po = pdf_data["text"].replace("*", "").replace("|", "").replace(" ", "")
        if (ex_po in pdf_data["text"] or 
            ex_po in pdf_text_clean_po or 
            (ex_po_lstrip and ex_po_lstrip in pdf_text_clean_po)):
            return pdf_data

    return None


def verify_po_items(ex_data: dict, matched_pdf: dict) -> dict:
    """
    ExcelデータとPDFデータの照合を行い、判定結果（PO, 型式, 数量, 単価）を返す
    """
    pdf_text_raw = matched_pdf["text_raw"]
    pdf_text_no_comma = pdf_text_raw.replace(",", "")
    ocr_extracted_text = ""

    ex_po = ex_data["po"]
    ex_item = ex_data["item"]
    ex_qty = ex_data["qty"]
    ex_price_clean = ex_data["price"].replace(",", "").strip()

    # --------------------------------------------------
    # 1. PO番号チェック
    # --------------------------------------------------
    po_text_check = matched_pdf["text"].replace("*", "").replace("|", "").replace(" ", "")
    po_name_check = matched_pdf["name"].replace(" ", "")
    ex_po_lstrip = ex_po.lstrip('0') if ex_po else ""

    check_po = bool(
        ex_po and (
            ex_po in po_text_check or 
            ex_po in po_name_check or 
            (ex_po_lstrip and (ex_po_lstrip in po_text_check or ex_po_lstrip in po_name_check))
        )
    )

    # --------------------------------------------------
    # 2. 型式チェック（4段階判定）
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

        # 18%クロップ (MCL用)
        left_text_18 = extract_left_column_text(matched_pdf["path"], ratio=0.18)
        clean_18 = left_text_18.replace("\n", "").replace("\r", "").replace("\t", "").replace(" ", "").replace(" ", "")
        norm_18 = clean_18.replace("-", "").upper()
        alpha_18 = re.sub(r'[^A-Za-z0-9]', '', clean_18).upper()

        if (ex_item_clean in clean_18 or f"K-{ex_item_clean}" in clean_18 or ex_item_clean in clean_18.replace("K-", "") or ex_item_norm in norm_18 or ex_item_alphanumeric in alpha_18):
            check_item = True
            print("      └ 🌸 救済モード(18%幅)により型式ブロックを正常検出！")
        else:
            # 35%クロップ (通常用)
            left_text_35 = extract_left_column_text(matched_pdf["path"], ratio=0.35)
            clean_35 = left_text_35.replace("\n", "").replace("\r", "").replace("\t", "").replace(" ", "").replace(" ", "")
            norm_35 = clean_35.replace("-", "").upper()
            alpha_35 = re.sub(r'[^A-Za-z0-9]', '', clean_35).upper()

            if (ex_item_clean in clean_35 or f"K-{ex_item_clean}" in clean_35 or ex_item_clean in clean_35.replace("K-", "") or ex_item_norm in norm_35 or ex_item_alphanumeric in alpha_35):
                check_item = True
                print("      └ 🌸 救済モード(35%幅)により型式ブロックを正常検出！")

    # 第4段階: 【最終手段】OCRスキャン解析
    if not check_item and ex_item_clean:
        print("      └ 🔍 最終手段: OCR（画像文字認識）スキャン解析を起動中...")
        is_ocr_ok, ocr_text_res = perform_ocr_rescue(matched_pdf["path"], ex_item_clean)
        ocr_extracted_text = ocr_text_res
        if is_ocr_ok:
            check_item = True
            print("      └ 🌸 OCRスキャン解析により図面内の型式文字を正常検出！")

    # --------------------------------------------------
    # 3. 数量チェック
    # --------------------------------------------------
    check_qty = bool(ex_qty and (ex_qty in pdf_text_no_comma or ex_qty in ocr_extracted_text))

    # --------------------------------------------------
    # 4. 単価チェック
    # --------------------------------------------------
    if ex_price_clean == "" or ex_item_clean == "57-04-322":
        check_price = True
    else:
        try:
            price_val = float(ex_price_clean)

            patterns = [
                ex_price_clean,
                f"{price_val:,.2f}",
                f"{price_val:,.0f}",
                f"¥{price_val:,.0f}", f"￥{price_val:,.0f}",
                f"JPY{price_val:,.0f}", f"JPY {price_val:,.0f}",
                f"{price_val:.0f}",
                f"{price_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                f"{price_val:.2f}".replace(".", ","),
                ex_price_clean.replace(".", ",")
            ]

            check_price = (
                any(p in pdf_text_raw for p in patterns) or 
                any(p in pdf_text_no_comma for p in patterns) or
                any(p in ocr_extracted_text for p in patterns)
            )

        except ValueError:
            check_price = (ex_price_clean in pdf_text_raw) or (ex_price_clean in pdf_text_no_comma) or (ex_price_clean in ocr_extracted_text)

    return {
        "po": check_po,
        "item": check_item,
        "qty": check_qty,
        "price": check_price,
        "is_all_ok": (check_item and check_qty and check_price)
    }