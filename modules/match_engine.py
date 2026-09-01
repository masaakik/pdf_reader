import re
from pathlib import Path
from modules.pdf_extractor import extract_left_column_text, perform_ocr_rescue

def normalize_model_text(text: str) -> str:
    """
    型式文字列の表記揺れ（ソフトハイフン\xad、識別記号KYO、0/O文字化け、ハイフン・記号等）を補正する
    """
    if not text:
        return ""
    
    t = text.upper()
    
    # 🌸 0. PDF特有の不可視文字(Soft Hyphen \xad) と 識別記号「KYO」を消去
    t = t.replace("\xad", "").replace("KYO", "")
    
    # 1. 先頭の「K-」や「K」を除去 (MCL用)
    if t.startswith("K-"):
        t = t[2:]
    elif t.startswith("K") and len(t) > 1 and not t.startswith("K3"):
        t = t[1:]

    # 2. OCR文字化け補正 (0/O 補正)
    t = re.sub(r'(\d)O([A-Z0-9])', r'\g<1>0\g<2>', t)
    t = re.sub(r'([A-Z0-9])O(\d)', r'\g<1>0\g<2>', t)

    # 3. 画面サイズ・寸法記号の OCR 文字化け補正 ( Rollflex 対策: "W1S" -> "W15" )
    t = re.sub(r'W(\d)S', r'W\g<1>5', t)
    t = t.replace("1S", "15")

    # 4. 「-08」を「-8」へ変換 (ダイフク用)
    t = re.sub(r'-0(\d)', r'-\g<1>', t)
    
    # 5. 数字と英字の間のスペースを削除
    t = re.sub(r'(\d)\s+([A-Z])', r'\g<1>\g<2>', t)
    
    # 🌸 6. 改行・スペース・ハイフン・アンダースコア・カンマ・ピリオド・記号類を完全除去
    t = re.sub(r'[\s\-\_\,\.]+', '', t)
    return t

def match_po_number(ex_po: str, pdf_list: list) -> dict | None:
    """ExcelのPO番号を元に、合致するPDFデータを特定する"""
    if not pdf_list:
        return None

    # [山崎タイ案件対策] フォルダ内にPDFが1つしかない場合は、ファイル名や本文の判定をスキップして自動採用
    if len(pdf_list) == 1:
        return pdf_list[0]

    if not ex_po:
        return pdf_list[0]

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
    ex_item_clean = ex_item.strip().replace(" ", "").replace(" ", "")
    pdf_text_no_newline = (
        pdf_text_raw.replace("\n", "")
                    .replace("\r", "")
                    .replace("\t", "")
                    .replace(" ", "")
                    .replace(" ", "")
    )

    # 表記揺れ補正用のテキストを作成
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
            ex_item_norm_custom in pdf_text_norm_custom or
            ex_item_norm_custom.replace("-", "") in pdf_text_norm_custom.replace("-", "")
        )
    )

    # 第2・3段階: 救済クロップ（18% / 35%）
    if not check_item and ex_item_clean:
        rescue_logs.append("      └ ⚠️ 通常抽出で型式が不一致のため、救済モード（ブロック位置解析）を発動します...")

        # 18%クロップ (MCL用)
        left_text_18 = extract_left_column_text(matched_pdf["path"], ratio=0.18)
        clean_18 = left_text_18.replace("\n", "").replace("\r", "").replace("\t", "").replace(" ", "").replace(" ", "")
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
            clean_35 = left_text_35.replace("\n", "").replace("\r", "").replace("\t", "").replace(" ", "").replace(" ", "")
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

    # 文字間スペース・改行・カンマを削ぎ落としてピリオド変換した単価照合用テキスト
    pdf_text_clean_price = (
        pdf_text_raw.replace(" ", "")
                    .replace(" ", "")
                    .replace("\n", "")
                    .replace("\r", "")
                    .replace(",", ".")
    )
    pdf_text_no_space = pdf_text_raw.replace(" ", "").replace(" ", "")

    if is_price_exempt:
        pre_check_price = True
    elif ex_price_clean:
        try:
            price_val = float(ex_price_clean)
            price_str_dot = f"{price_val:.2f}"   # 例: "1911.60"
            price_str_short = f"{price_val}"     # 例: "1911.6"

            patterns = [
                ex_price_clean,
                f"{price_val:,.2f}",
                f"{price_val:,.0f}",
                f"¥{price_val:,.0f}", f"￥{price_val:,.0f}",
                f"JPY{price_val:,.0f}", f"JPY {price_val:,.0f}",
                f"{price_val:.0f}",
                price_str_dot,
                price_str_short
            ]

            pre_check_price = (
                any(p in pdf_text_raw for p in patterns) or 
                any(p in pdf_text_no_comma for p in patterns) or
                (ex_price_clean in pdf_text_no_space) or
                (price_str_dot in pdf_text_clean_price) or
                (price_str_short in pdf_text_clean_price)
            )
        except ValueError:
            pre_check_price = (
                (ex_price_clean in pdf_text_raw) or 
                (ex_price_clean in pdf_text_no_comma) or 
                (ex_price_clean in pdf_text_no_space) or
                (ex_price_clean in pdf_text_clean_price)
            )

    # --------------------------------------------------
    # 4. 第4段階: 【最終手段】OCRスキャン解析（型式NG または 単価NG の場合に起動）
    # --------------------------------------------------
    if (not check_item or not pre_check_price) and ex_item_clean:
        rescue_logs.append("      └ 🔍 最終手段: OCR（画像文字認識）スキャン解析を起動中...")
        is_ocr_ok, ocr_text_res = perform_ocr_rescue(matched_pdf["path"], ex_item_clean)
        ocr_extracted_text = ocr_text_res
        
        ocr_norm_custom = normalize_model_text(ocr_extracted_text)
        
        if is_ocr_ok or (ex_item_norm_custom and ex_item_norm_custom in ocr_norm_custom):
            check_item = True
            rescue_logs.append("      └ 🌸 OCRスキャン解析により図面内の文字情報を正常検出！")

    # --------------------------------------------------
    # 5. 最終数量 ＆ 最終単価チェック (OCRテキスト含めて確定)
    # --------------------------------------------------
    check_qty = bool(ex_qty and (ex_qty in pdf_text_no_comma or ex_qty in ocr_extracted_text))

    if is_price_exempt:
        check_price = True
    else:
        try:
            price_val = float(ex_price_clean)
            price_str_dot = f"{price_val:.2f}"
            price_str_short = f"{price_val}"

            patterns = [
                ex_price_clean,
                f"{price_val:,.2f}",
                f"{price_val:,.0f}",
                f"¥{price_val:,.0f}", f"￥{price_val:,.0f}",
                f"JPY{price_val:,.0f}", f"JPY {price_val:,.0f}",
                f"{price_val:.0f}",
                price_str_dot,
                price_str_short
            ]

            check_price = (
                any(p in pdf_text_raw for p in patterns) or 
                any(p in pdf_text_no_comma for p in patterns) or
                any(p in ocr_extracted_text for p in patterns) or
                (ex_price_clean in pdf_text_no_space) or
                (price_str_dot in pdf_text_clean_price) or
                (price_str_short in pdf_text_clean_price) or
                (ex_price_clean in ocr_extracted_text.replace(",", "").replace(" ", ""))
            )
        except ValueError:
            check_price = (
                (ex_price_clean in pdf_text_raw) or 
                (ex_price_clean in pdf_text_no_comma) or 
                (ex_price_clean in ocr_extracted_text) or
                (ex_price_clean in pdf_text_no_space) or
                (ex_price_clean in pdf_text_clean_price) or
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