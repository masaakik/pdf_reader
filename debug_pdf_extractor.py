from pathlib import Path
from pypdf import PdfReader

parent_dir = Path("./files")
target_folders = [f for f in parent_dir.glob("*Damon*Australia*") if f.is_dir()]

if target_folders:
    pdf_path = list(target_folders[0].glob("*.pdf"))[0]
    print(f"Target PDF: {pdf_path.name}\n")

    reader = PdfReader(pdf_path)
    
    print("=== 全アノテーション（注釈）オブジェクト抽出 ===")
    for i, page in enumerate(reader.pages):
        print(f"\n--- Page {i+1} ---")
        if "/Annots" in page:
            for idx, annot in enumerate(page["/Annots"]):
                try:
                    obj = annot.get_object()
                    print(f"\n[Annot {idx+1}] Subtype: {obj.get('/Subtype')}")
                    for k, v in obj.items():
                        val_str = str(v)
                        if any(term in val_str for term in ["PR", "390", "SENERGY", "5290615", "10,850"]):
                            print(f"  ★ {k}: {val_str}")
                        elif k in ["/Contents", "/RC", "/V", "/AP"]:
                            print(f"  {k}: {val_str[:100]}")
                except Exception as e:
                    print(f"  Error reading annot: {e}")
        else:
            print("アノテーションなし")
else:
    print("Folder not found.")