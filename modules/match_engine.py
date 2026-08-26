import re
from pathlib import Path
from modules.pdf_extractor import extract_left_column_text, perform_ocr_rescue

def normalize_model_text(text: str) -> str:
    """
    型式文字列の表記揺れ（頭ゼロ、余計なスペース、ハイフン、特殊記号等）を強力に補正する
    """
    if not text:
        return ""
    
    # 全角英数字を半角化し大文字に統一
    t = text.upper()
    
    # 1. 「-08」のようなハイフン直後の頭ゼロを補正 (例: "-08" -> "-8")
    t = re.sub(r'-0(\d)', r'-\1', t)
    
    # 2. 数字とアルファベットの間のスペースを削除 (例: "8 ZSAY" -> "8ZSAY")
    t = re.sub(r'(\d)\s+([A-Z])', r'\1\2', t)
    
    # 3. 改行・タブ・スペース全削除
    t = re.sub(r'\s+', '', t)
    
    return t


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
    ExcelデータとPDFデータの照合を行い、判定結果（PO, 型式, 数量, 単価）とログを返す
    """
    pdf_text_raw = matched_pdf["text_raw"]
    pdf_text_no_comma = pdf_text_raw.replace(",", "")
    ocr_extracted_text = ""
    rescue_logs = []  # メッセージ出力用配列

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
    ex_item_clean = ex_item.strip().replace(" ", "")
    pdf_text_no_newline = (
        pdf_text_raw.replace("\n", "")
                    .replace("\r", "")
                    .replace("\t", "")
                    .replace(" ", "")
    )

    # 💡 [強化] 表記揺れ（頭ゼロ・スペース）吸収用の正規化テキストを作成
    ex_item_norm_custom = normalize_model_text(ex_item)
    pdf_text_norm_custom = normalize_model_text(pdf_text_raw)

    ex_item_norm = ex_item_clean.replace("-", "").upper()
    pdf_text_norm = pdf_text_no_newline.replace("-", "").upper()

    ex_item_alphanumeric = re.sub(r'[^A-Za-z0-9]', '', ex_item_clean).upper()
    pdf_text_alphanumeric = re.sub(r'[^A-Za-z0-9]', '', pdf_text_no_newline).upper()

    # O/I/0 の表記揺れを補正した判定テキスト
    pdf_text_zero_o = pdf_text_alphanumeric.replace("1OX", "10X").replace("IOX", "10X").replace("I0X", "10X")
    ex_item_zero_o = ex_item_alphanumeric.replace("1OX", "10X").replace("IOX", "10X").replace("I0X", "10X")

    # 第1段階: 標準チェック
    check_item = bool(
        ex_item_clean and (
            ex_item_clean in pdf_text_no_newline or 
            f"K-{ex_item_clean}" in pdf_text_no_newline or
            ex_item_clean in pdf_text_no_newline.replace("K-", "") or
            ex_item_norm in pdf_text_norm or
            ex_item_alphanumeric in pdf_text_alphanumeric or
            ex_item_zero_o in pdf_text_zero_o or
            # 🌸 新規追加: 「-08」->「-8」やスペース詰めの表記揺れ補正一致
            ex_item_norm_custom in pdf_text_norm_custom or
            ex_item_norm_custom.replace("-", "") in pdf_text_norm_custom.replace("-", "")
        )
    )

    # 第2・3段階: 救済クロップ（18% / 35%）
    if not check_item and ex_item_clean:
        rescue_logs.append("      └ ⚠️ 通常抽出で型式が不一致のため、救済モード（ブロック位置解析）を発動します...")

        # 18%クロップ (MCL用)
        left_text_18 = extract_left_column_text(matched_pdf["path"], ratio=0.18)
        clean_18 = left_text_18.replace("\n", "").replace("\r", "").replace("\t", "").replace(" ", "")
        norm_18 = clean_18.replace("-", "").upper()
        alpha_18 = re.sub(r'[^A-Za-z0-9]', '', clean_18).upper()
        custom_18 = normalize_model_text(left_text_18)

        if (ex_item_clean in clean_18 or 
            f"K-{ex_item_clean}" in clean_18 or 
            ex_item_clean in clean_18.replace("K-", "") or 
            ex_item_norm in norm_18 or 
            ex_item_alphanumeric in alpha_18 or
            ex_item_norm_custom in custom_18):
            check_item = True
            rescue_logs.append("      └ 🌸 救済モード(18%幅)により型式ブロックを正常検出！")
        else:
            # 35%クロップ (通常用)
            left_text_35 = extract_left_column_text(matched_pdf["path"], ratio=0.35)
            clean_35 = left_text_35.replace("\n", "").replace("\r", "").replace("\t", "").replace(" ", "")
            norm_35 = clean_35.replace("-", "").upper()
            alpha_35 = re.sub(r'[^A-Za-z0-9]', '', clean_35).upper()
            custom_35 = normalize_model_text(left_text_35)

            if (ex_item_clean in clean_35 or 
                f"K-{ex_item_clean}" in clean_35 or 
                ex_item_clean in clean_35.replace("K-", "") or 
                ex_item_norm in norm_35 or 
                ex_item_alphanumeric in alpha_35 or
                ex_item_norm_custom in custom_35):
                check_item = True
                rescue_logs.append("      └ 🌸 救済モード(35%幅)により型式ブロックを正常検出！")

    # --------------------------------------------------
    # 3. 事前単価チェック
    # --------------------------------------------------
    is_price_exempt = (ex_price_clean == "" or ex_item_clean == "57-04-322")
    pre_check_price = False

    if is_price_exempt:
        pre_check_price = True
    elif ex_price_clean:
        pdf_text_no_space = pdf_text_raw.replace(" ", "")
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
            pre_check_price = (
                any(p in pdf_text_raw for p in patterns) or 
                any(p in pdf_text_no_comma for p in patterns) or
                (ex_price_clean in pdf_text_no_space)
            )
        except ValueError:
            pre_check_price = (
                (ex_price_clean in pdf_text_raw) or 
                (ex_price_clean in pdf_text_no_comma) or 
                (ex_price_clean in pdf_text_no_space)
            )

    # --------------------------------------------------
    # 4. 第4段階: 【最終手段】OCRスキャン解析（型式NG または 単価NG の場合に起動）
    # --------------------------------------------------
    if (not check_item or not pre_check_price) and ex_item_clean:
        rescue_logs.append("      └ 🔍 最終手段: OCR（画像文字認識）スキャン解析を起動中...")
        is_ocr_ok, ocr_text_res = perform_ocr_rescue(matched_pdf["path"], ex_item_clean)
        ocr_extracted_text = ocr_text_res
        
        if is_ocr_ok:
            check_item = True
            rescue_logs.append("      └ 🌸 OCRスキャン解析により図面内の文字情報を正常検出！")

    # --------------------------------------------------
    # 5. 最終数量 ＆ 最終単価チェック (OCRテキスト含めて確定)
    # --------------------------------------------------
    check_qty = bool(ex_qty and (ex_qty in pdf_text_no_comma or ex_qty in ocr_extracted_text))

    if is_price_exempt:
        check_price = True
    else:
        pdf_text_no_space = pdf_text_raw.replace(" ", "")
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
                any(p in ocr_extracted_text for p in patterns) or
                (ex_price_clean in pdf_text_no_space) or
                (ex_price_clean in ocr_extracted_text.replace(",", "").replace(" ", ""))
            )
        except ValueError:
            check_price = (
                (ex_price_clean in pdf_text_raw) or 
                (ex_price_clean in pdf_text_no_comma) or 
                (ex_price_clean in ocr_extracted_text) or
                (ex_price_clean in pdf_text_no_space) or
                (ex_price_clean in ocr_extracted_text.replace(",", "").replace(" ", ""))
            )

    return {
        "po": check_po,
        "item": check_item,
        "qty": check_qty,
        "price": check_price,
        "is_all_ok": (check_item and check_qty and check_price),
        "rescue_logs": rescue_logs
    }