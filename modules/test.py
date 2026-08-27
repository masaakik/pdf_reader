from pathlib import Path
import pypdfium2 as pdfium
import easyocr
import numpy as np

# files フォルダ内の「0091794」を含むフォルダを自動検出
files_dir = Path("files")
target_dirs = [d for d in files_dir.glob("*0091794*") if d.is_dir()]

if not target_dirs:
    print("❌ 対象フォルダが見つかりませんでした。")
else:
    target_dir = target_dirs[0]
    pdf_files = list(target_dir.glob("*.pdf"))

    if not pdf_files:
        print(f"❌ フォルダ（{target_dir.name}）内にPDFが見つかりませんでした。")
    else:
        pdf_path = pdf_files[0]
        print(f"📂 検出フォルダ: {target_dir.name}")
        print(f"📄 読み込みPDF: {pdf_path.name}")

        pdf = pdfium.PdfDocument(pdf_path)
        page = pdf[0]

        # scale=3.0 でレンダリングしてOCR実行
        pil_image = page.render(scale=3.0).to_pil().convert('L')
        img_np = np.array(pil_image)

        reader = easyocr.Reader(['en'], gpu=False)
        results = reader.readtext(img_np, detail=0)

        print("\n=== EasyOCRが認識した全テキスト ===")
        print(" ".join(results))