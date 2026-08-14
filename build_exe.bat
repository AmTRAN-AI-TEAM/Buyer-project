@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

echo Building BuyerReports.exe...
echo.

set "PYTHON_CMD="
py -3 --version >nul 2>nul
if not errorlevel 1 set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD (
    python --version >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=python"
)

if not defined PYTHON_CMD (
    echo No Python runtime found.
    echo Please install Python 3, or make sure python.exe is available in PATH.
    echo.
    pause
    exit /b 1
)

echo Using Python command: %PYTHON_CMD%
echo.

%PYTHON_CMD% -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo.
    echo Build failed while installing dependencies.
    pause
    exit /b 1
)

if exist build rmdir /s /q build
if exist release\BuyerReports rmdir /s /q release\BuyerReports

%PYTHON_CMD% -m PyInstaller ^
  --clean ^
  --onedir ^
  --console ^
  --name BuyerReports ^
  --distpath release ^
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
if not exist release\BuyerReports\intput\AVTC mkdir release\BuyerReports\intput\AVTC
if not exist release\BuyerReports\intput\RAKEN mkdir release\BuyerReports\intput\RAKEN
if not exist release\BuyerReports\output mkdir release\BuyerReports\output
if not exist release\BuyerReports\output\AVTC mkdir release\BuyerReports\output\AVTC
if not exist release\BuyerReports\output\RAKEN mkdir release\BuyerReports\output\RAKEN
copy /Y "Windows執行檔(exe)使用說明.txt" "release\BuyerReports\Windows執行檔(exe)使用說明.txt" >nul
copy /Y "buyer_reports.ini" "release\BuyerReports\buyer_reports.ini" >nul

echo.
echo Done.
echo Release folder:
echo %cd%\release\BuyerReports
echo.
pause
