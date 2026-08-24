@echo off
chcp 65001 > nul

:: バッチファイルがある場所に移動
cd /d "%~dp0"

echo 仮想環境 (.venv) を立ち上げています...
call .venv\Scripts\activate.bat

:: Pythonの出力文字コードを UTF-8 に強制指定
set PYTHONIOENCODING=utf-8

echo main_checker.py を実行中...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$OutputEncoding = [System.Text.Encoding]::UTF8; Set-Location '%~dp0'; python main_checker.py | Tee-Object -FilePath 'execution.log'"

echo.
echo 処理が完了しました。ログは execution.log に保存されました。
pause