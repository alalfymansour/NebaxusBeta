@echo off
title NebaxusBeta Builder
echo ============================================
echo          NebaxusBeta Builder
echo ============================================
echo.

echo [1/3] Installing Python dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo FAILED: pip install. Check requirements.txt
    pause
    exit /b 1
)
pip install cairosvg pillow nuitka
if %errorlevel% neq 0 (
    echo FAILED: pip install build dependencies
    pause
    exit /b 1
)
echo.

echo [2/3] Generating icon + compiling EXE...
python build.py
if %errorlevel% neq 0 (
    echo FAILED: Build step
    pause
    exit /b 1
)
echo.

echo [3/3] Done!
echo.
if exist dist\NebaxusBeta_Setup.exe (
    echo SUCCESS: dist\NebaxusBeta_Setup.exe
    dir dist\NebaxusBeta_Setup.exe
) else (
    echo WARNING: Installer not found. Open installer.iss in Inno Setup and compile manually.
)
echo.
pause
