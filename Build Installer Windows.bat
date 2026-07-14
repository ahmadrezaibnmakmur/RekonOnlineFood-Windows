@echo off
setlocal
title Build Rekon Online Food Installer

cd /d "%~dp0"

echo.
echo ==========================================
echo   BUILD REKON ONLINE FOOD INSTALLER
echo ==========================================
echo.

if not exist "requirements.txt" (
    echo ERROR: File requirements.txt tidak ditemukan di folder ini.
    echo.
    echo Pastikan kamu sudah extract seluruh folder project, bukan menjalankan file ini
    echo langsung dari ZIP/RAR/WinRAR.
    echo.
    echo Folder yang benar harus berisi:
    echo - requirements.txt
    echo - build_windows.py
    echo - rekon.py
    echo - webapp
    echo - installer
    echo.
    echo Folder saat ini:
    cd
    echo.
    pause
    exit /b 1
)

where py >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python launcher tidak ditemukan.
    echo Install Python dulu dari https://www.python.org/downloads/windows/
    echo Saat install, centang "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

echo [1/2] Install/update build dependency...
py -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo.
    echo ERROR: Install dependency gagal.
    pause
    exit /b 1
)

echo.
echo [2/2] Build app dan installer...
py build_windows.py
if errorlevel 1 (
    echo.
    echo ERROR: Build gagal.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo   SELESAI
echo ==========================================
echo.
echo Installer ada di:
echo dist\installer\RekonOnlineFoodSetup-1.01.exe
echo.
pause
