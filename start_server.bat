@echo off
REM 簡單的 Python HTTP 伺服器啟動腳本
REM 確保你的電腦已安裝 Python

echo.
echo ====================================
echo Excel 視覺化工具 - 本地伺服器
echo ====================================
echo.

REM 檢查 Python
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo 錯誤: 未找到 Python
    echo 請先安裝 Python: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 獲取目前目錄
cd /d "%~dp0"

echo.
echo 正在啟動伺服器...
echo.
echo 🌐 在瀏覽器中打開: http://localhost:8000
echo.
echo 按 Ctrl+C 停止伺服器
echo.

REM 啟動 Python HTTP 伺服器
python -m http.server 8000

pause
