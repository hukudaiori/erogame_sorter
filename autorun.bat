@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

set "exit_trigger=end"

echo ==============================
echo 複数のURLを入力してください。
echo 入力を終了するには「%exit_trigger%」と入力してください。
echo ==============================

:input_loop
set /p newurl=URL: 
if /i "%newurl%"=="%exit_trigger%" goto after_input

:: 重複チェックして追加
findstr /C:"%newurl%" urls.txt >nul 2>&1
if errorlevel 1 (
    echo %newurl% >> urls.txt
    echo ✅ URLを追加しました。
) else (
    echo ⚠️ URLはすでに存在しています。
)
goto input_loop

:after_input
python erogeme_scraper.py

start "" "erogame_stats.xlsx"

echo.
echo 📁 終了しました。何かキーを押すと閉じます…
pause > nul