import logging
from pathlib import Path
import pdfplumber

# pdfminerのログレベル制御
logging.getLogger("pdfminer").setLevel(logging.ERROR)

def inspect_pdf_file(pdf_path: Path):
    """1つのPDFファイルから抽出されるすべてのテキスト（本文・注釈）を出力する"""
    print("\n" + "=" * 70)
    print(f"📄 対象PDF: {pdf_path.name}")
    print("=" * 70)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"総ページ数: {len(pdf.pages)}\n")

            for page_idx, page in enumerate(pdf.pages, 1):
                print(f"--- ページ {page_idx} ---")

                # 1. 本文テキスト抽出
                print("【1. 本文テキスト (extract_text)】")
                try:
                    text = page.extract_text()
                    if text:
                        print(text)
                    else:
                        print("(本文テキストなし)")
                except Exception as e:
                    print(f"❌ 本文抽出エラー: {e}")

                # 2. 改行・スペース除去済みテキスト
                if text:
                    text_no_nl = text.replace("\n", "").replace("\r", "").replace(" ", "").replace(" ", "")
                    print("\n【2. 改行・スペース除去後テキスト】")
                    print(text_no_nl)

                # 3. 注釈テキスト抽出（PDF-XChange等で書き加えた上書き文字）
                print("\n【3. 注釈データ (annots)】")
                if hasattr(page, 'annots') and page.annots:
                    annot_found = False
                    for idx, annot in enumerate(page.annots, 1):
                        try:
                            content = annot.get("contents") or annot.get("contents_pt") or ""
                            title = annot.get("title") or ""
                            subtype = annot.get("subtype") or ""
                            
                            if content or title or subtype:
                                print(f"  ・注釈 #{idx} [Subtype: {subtype}, Title: {title}]:")
                                print(f"    内容: {repr(content)}")
                                annot_found = True
                        except Exception as e:
                            print(f"  ・注釈 #{idx} 解析エラー (破損文字コード等): {e}")
                    
                    if not annot_found:
                        print("(解析可能な注釈本文なし)")
                else:
                    print("(注釈なし)")

                print("-" * 50)

    except Exception as e:
        print(f"❌ PDFファイル全体エラー: {e}")

def run_debug_extractor(parent_dir_path="./files"):
    """files フォルダ配下のすべてのPDFを探索して出力"""
    parent_dir = Path(parent_dir_path)

    if not parent_dir.exists():
        print(f"❌ 親フォルダ '{parent_dir}' が見つかりません。")
        return

    # files/ 配下の全PDFを再帰検索
    pdf_files = list(parent_dir.glob("**/*.pdf"))

    if not pdf_files:
        print(f"⚠️ '{parent_dir}' 内にPDFファイルが見つかりませんでした。")
        return

    print(f"🔍 合計 {len(pdf_files)} 件のPDFファイルを検出しました。順次テキストを出力します。")

    for pdf_file in pdf_files:
        inspect_pdf_file(pdf_file)

if __name__ == "__main__":
    run_debug_extractor("./files")