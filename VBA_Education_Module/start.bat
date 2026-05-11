@echo off
title VBA Macro Education Module
color 07
cls

echo.
echo  ============================================================
echo   VBA Macro Basics — Learning Portal
echo   DX Education 2026  ^|  Yamaha Motor Co., Ltd.
echo  ============================================================
echo.

:: ── Locate Python ────────────────────────────────────────────
set PYCMD=

python --version >nul 2>&1
if %errorlevel% equ 0 ( set PYCMD=python & goto :found_python )

py --version >nul 2>&1
if %errorlevel% equ 0 ( set PYCMD=py & goto :found_python )

python3 --version >nul 2>&1
if %errorlevel% equ 0 ( set PYCMD=python3 & goto :found_python )

echo  [ERROR] Python not found on this system.
echo.
echo  Install Python 3.8+ from https://www.python.org/downloads/
echo  Make sure to tick "Add Python to PATH" during setup.
echo.
pause
exit /b 1

:found_python
for /f "tokens=*" %%i in ('%PYCMD% --version 2^>^&1') do set PYVER=%%i
echo  [OK] %PYVER% found

:: ── Verify files ─────────────────────────────────────────────
if not exist "%~dp0modules\" (
    echo  [ERROR] Modules folder missing. Ensure the full package is present.
    pause
    exit /b 1
)
if not exist "%~dp0server.py" (
    echo  [ERROR] server.py not found in %~dp0
    pause
    exit /b 1
)
echo  [OK] Module files verified

:: ── Start server ─────────────────────────────────────────────
echo.
echo  Starting server... your browser will open automatically.
echo  Share the Network URL shown below with colleagues.
echo  Press Ctrl+C in this window to stop.
echo.

cd /d "%~dp0"
%PYCMD% server.py

if %errorlevel% neq 0 (
    echo.
    echo  [ERROR] Server exited with an error.
    echo  Check that no other program is using port 8000.
    echo.
    pause
    exit /b 1
)

exit /b 0
