from pathlib import Path
from pypdf import PdfReader


def read_pdf_text(pdf_path: Path) -> str:
    """単一のPDFファイルからテキストを抽出する関数"""
    reader = PdfReader(pdf_path)
    text_list = []

    for i, page in enumerate(reader.pages):
        page_text = page.extract_text() or ""
        text_list.append(f"  [ {i + 1} ページ目 ]\n{page_text}")

    return "\n".join(text_list)


if __name__ == "__main__":
    # スクリプトがあるフォルダを基準にする
    base_dir = Path(__file__).parent
    pdf_dir = base_dir / "pdf_files"

    # フォルダが存在しない場合
    if not pdf_dir.exists():
        print(f"❌ エラー: '{pdf_dir}' フォルダが存在しません。")
    else:
        # pdf_files フォルダ内のすべての .pdf ファイルを取得 (*.pdf)
        pdf_files = list(pdf_dir.glob("*.pdf"))

        if not pdf_files:
            print(
                f"⚠️ '{pdf_dir.name}' フォルダ内にPDFファイルが見つかりませんでした。"
            )
        else:
            print(f"📂 {len(pdf_files)} 個のPDFファイルが見つかりました。\n")

            all_results = []

            # 見つかったPDFファイルを1つずつ処理
            for pdf_file in pdf_files:
                print(f"📄 処理中: {pdf_file.name} ...")
                try:
                    pdf_text = read_pdf_text(pdf_file)

                    # ファイル名ヘッダーをつけて結果をまとめる
                    header = (
                        f"========================================\n"
                        f" ファイル名: {pdf_file.name}\n"
                        f"========================================\n"
                    )
                    all_results.append(f"{header}{pdf_text}")
                except Exception as e:
                    print(
                        f"  ❌ '{pdf_file.name}' の読み込み中にエラーが発生しました: {e}"
                    )

            # すべてのテキストを結合
            combined_text = "\n\n\n".join(all_results)

            # 画面に出力
            print("\n--- 抽出結果 ---")
            print(combined_text)

            # output.txt に結果をひとまとめにして保存
            output_file = base_dir / "output.txt"
            output_file.write_text(combined_text, encoding="utf-8")
            print(
                f"\n✅ すべてのPDFのテキストを '{output_file.name}' に保存しました！"
            )