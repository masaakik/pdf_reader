import sys
from pathlib import Path
import pdfplumber
import pypdfium2 as pdfium
import easyocr
import numpy as np

def test_extract_all(pdf_path: Path):
    print("=" * 70)
    print(f"📄 解析対象ファイル: {pdf_path.name}")
    print("=" * 70)

    if not pdf_path.exists():
        print(f"❌ エラー: ファイルが存在しません ➔ {pdf_path}")
        return

    # --------------------------------------------------
    # EasyOCR 抽出 (scale=3.0 / Lモード)
    # --------------------------------------------------
    print("\n--- 🔹 OCR（画像認識）抽出テスト (scale=3.0 / Lモード) ---")
    print("⏳ OCRを初期化して解析中...")
    
    reader = easyocr.Reader(['en'], gpu=False)
    pdf_doc = pdfium.PdfDocument(pdf_path)

    for i, page in enumerate(pdf_doc, start=1):
        # 図面判定（英語表記で判定）
        plumb_text = ""
        with pdfplumber.open(pdf_path) as plumb_pdf:
            plumb_text = plumb_pdf.pages[i-1].extract_text() or ""

        is_drawing = "This Drawing is the property of Kyowa" in plumb_text
        status_label = "【図面ページ (スキップ対象)】" if is_drawing else "【注文書ページ (解析対象)】"

        print(f"\n--- [ Page {i} ] {status_label} ---")

        if is_drawing:
            print("（図面ページのためスキップします）")
            continue

        # scale=3.0 + グレースケールで解析
        pil_image = page.render(scale=3.0).to_pil().convert('L')
        img_np = np.array(pil_image)
        
        results = reader.readtext(img_np, detail=0)
        page_ocr_text = " ".join(results)

        if page_ocr_text.strip():
            print(page_ocr_text)
        else:
            print("(OCRで認識できる文字がありませんでした)")

    pdf_doc.close()
    print("\n" + "=" * 70)
    print("✨ すべての抽出テストが完了しました。")


if __name__ == "__main__":
    target_pdf = Path("./files/山崎 DHL DEXION P2008343 9.28 (NSD利用）/P2008343-1_DEXION 図面付き.pdf")

    if not target_pdf.exists():
        pdf_files = list(Path("./files").rglob("*.pdf"))
        if pdf_files:
            target_pdf = pdf_files[0]
        else:
            print("❌ files フォルダ内にテスト用のPDFが見つかりませんでした。")
            sys.exit(1)

    test_extract_all(target_pdf)