@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

echo Building BuyerReports.exe...
echo.

py -3 -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo.
    echo Build failed while installing dependencies.
    pause
    exit /b 1
)

if exist build rmdir /s /q build
if exist release\BuyerReports rmdir /s /q release\BuyerReports

py -3 -m PyInstaller ^
  --clean ^
  --onefile ^
  --console ^
  --name BuyerReports ^
  --distpath release\BuyerReports ^
  --workpath build ^
  --specpath build ^
  generate_buyer_reports.py
if errorlevel 1 (
    echo.
    echo Build failed while creating the exe.
    pause
    exit /b 1
)

if not exist release\BuyerReports\intput mkdir release\BuyerReports\intput
if not exist release\BuyerReports\output mkdir release\BuyerReports\output
copy /Y "Windows執行檔(exe)使用說明.txt" "release\BuyerReports\Windows執行檔(exe)使用說明.txt" >nul

echo.
echo Done.
echo Release folder:
echo %cd%\release\BuyerReports
echo.
pause
