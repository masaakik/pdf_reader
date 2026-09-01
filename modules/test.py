from pathlib import Path
import pdfplumber

# 1. files フォルダから「4502234358」または「KION」を含むフォルダを自動検出
base_dir = Path("files")
target_dirs = [d for d in base_dir.glob("*4502234358*") if d.is_dir()]

if not target_dirs:
    target_dirs = [d for d in base_dir.glob("*KION*") if d.is_dir()]

if not target_dirs:
    print("❌ 指定条件に合うフォルダが見つかりませんでした。")
else:
    target_dir = target_dirs[0]
    pdf_files = list(target_dir.glob("*.pdf"))

    if not pdf_files:
        print(f"❌ フォルダ（{target_dir.name}）内にPDFが見つかりませんでした。")
    else:
        pdf_path = pdf_files[0]
        print(f"📂 検出フォルダ: {target_dir.name}")
        print(f"📄 対象PDF: {pdf_path.name}\n")

        # --- ① pdfplumber テキスト抽出 ---
        print("=== 【1. pdfplumber テキスト抽出】 ===")
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    print(f"--- ページ {i+1} ---")
                    print(repr(text))
        except Exception as e:
            print(f"⚠️ pdfplumber 抽出エラー: {e}")

        # --- ② EasyOCR 解析 (未アクセスの場合は安全にスキップ) ---
        print("\n=== 【2. EasyOCR (scale=3.0) 認識テキスト】 ===")
        try:
            import pypdfium2 as pdfium
            import easyocr
            import numpy as np

            pdf_doc = pdfium.PdfDocument(pdf_path)
            page = pdf_doc[0]

            pil_image = page.render(scale=3.0).to_pil().convert('L')
            img_np = np.array(pil_image)

            reader = easyocr.Reader(['en'], gpu=False)
            results = reader.readtext(img_np, detail=0)

            print(" ".join(results))
            pdf_doc.close()
        except ModuleNotFoundError as e:
            print(f"⚠️ パッケージ未インストールのためのスキップ: {e}")
        except Exception as e:
            print(f"⚠️ OCR処理エラー: {e}")
            