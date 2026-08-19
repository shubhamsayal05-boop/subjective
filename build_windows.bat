@echo off
setlocal EnableDelayedExpansion

echo ============================================================
echo   Building Subjective Spreadsheet Tool (.exe)
echo   Requires: Python 3.10+ on Windows
echo ============================================================
echo.

cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo Python not found. Install Python from https://www.python.org/downloads/
    echo Check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo [1/4] Installing dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements-build.txt
if errorlevel 1 goto :error

echo.
echo [2/4] Cleaning old build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo [3/4] Running PyInstaller (may take several minutes)...
python -m PyInstaller build_exe.spec --noconfirm
if errorlevel 1 goto :error

echo.
echo [4/4] Build complete!
echo.
echo   Application folder:
echo   %CD%\dist\SubjectiveSpreadsheetTool\
echo.
echo   Run: dist\SubjectiveSpreadsheetTool\SubjectiveSpreadsheetTool.exe
echo.
echo   You can copy the entire "SubjectiveSpreadsheetTool" folder to any PC.
echo   No Python or Excel installation required on the target machine.
echo.
pause
exit /b 0

:error
echo.
echo BUILD FAILED. See errors above.
pause
exit /b 1
