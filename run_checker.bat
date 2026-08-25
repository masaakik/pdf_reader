@echo off
chcp 65001 > nul

:: バッチファイルがある場所に移動
cd /d "%~dp0"

echo 仮想環境 (.venv) を立ち上げています...
call .venv\Scripts\activate.bat

:: Pythonの出力文字コードを UTF-8 に強制指定
set PYTHONIOENCODING=utf-8

echo main_checker.py を実行中...

:: 実行時の現在日時をヘッダーとしてログと画面に出力
powershell -NoProfile -ExecutionPolicy Bypass -Command " Get-Date -Format '=== yyyy/MM/dd HH:mm:ss 実行開始 ===' | Tee-Object -FilePath 'execution.log' -Append "

:: メイン処理の実行（-Append を指定して末尾へ追記）
powershell -NoProfile -ExecutionPolicy Bypass -Command " $OutputEncoding = [System.Text.Encoding]::UTF8; Set-Location '%~dp0'; python main_checker.py | Tee-Object -FilePath 'execution.log' -Append "

:: 区切り線を出力
powershell -NoProfile -ExecutionPolicy Bypass -Command " Write-Output '`n' | Tee-Object -FilePath 'execution.log' -Append "

echo.
echo 処理が完了しました。ログは execution.log に追記保存されました。
pause