@echo off
title Building Asset Configuration Tool...
echo.
echo 📦 Preparing to build standalone EXE...
echo.

:: --- CONFIGURATION ---
set MAIN_SCRIPT=app.py
set EXE_NAME=AssetConfigTool
:: -----------------------

:: Check if main script exists
if not exist "%MAIN_SCRIPT%" (
    echo ❌ ERROR: %MAIN_SCRIPT% not found in this folder!
    echo Please make sure this .bat file is in the SAME folder as your Python script.
    pause
    exit /b 1
)

:: Check required files
if not exist "templates\index.html" (
    echo ❌ WARNING: templates\index.html not found — web UI may fail!
)
if not exist "configuration_form template.xlsx" (
    echo ❌ ERROR: configuration_form template.xlsx missing — cannot generate forms!
    pause
    exit /b 1
)

echo ✅ Required files present. Starting PyInstaller...
echo.

pyinstaller ^
  --onefile ^
  --windowed ^
  --name="%EXE_NAME%" ^
  --add-data="configuration_form template.xlsx;." ^
  --add-data="asset_database.json;." ^
  --add-data="templates;templates" ^
  --add-data="static;static" ^
  --hidden-import=pandas ^
  --hidden-import=openpyxl ^
  --hidden-import=flask.json ^
  --collect-submodules="flask" ^
  "%MAIN_SCRIPT%"

echo.
if %ERRORLEVEL% EQU 0 (
    echo.
    echo 🎉 SUCCESS! EXE created at:
    echo     dist\%EXE_NAME%.exe
    echo.
    echo 💡 Tip: Distribute the entire 'dist' folder + this folder's:
    echo        - configuration_form template.xlsx
    echo        - asset_database.json
    echo        - templates\ (folder)
    echo        - static\ (folder)
    echo.
    pause
    if exist "dist\%EXE_NAME%.exe" (
        echo Launching app...
        start "" "dist\%EXE_NAME%.exe"
    )
) else (
    echo.
    echo ❌ FAILED. Common fixes:
    echo   • Run Command Prompt as Administrator
    echo   • pip install pyinstaller flask openpyxl pandas
    echo   • Check file/folder names (spaces? typos?)
    pause
)