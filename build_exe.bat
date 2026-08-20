@echo off
setlocal enabledelayedexpansion

echo ============================================
echo  DRB Subjective Drivability Tool - Build
echo ============================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not on PATH.
    echo Install Python 3.10+ from https://www.python.org/downloads/
    exit /b 1
)

echo [1/3] Installing dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt -r requirements-build.txt
if errorlevel 1 exit /b 1

echo.
echo [2/3] Building Windows executable (this may take several minutes)...
python -m PyInstaller --noconfirm --clean drb_tool.spec
if errorlevel 1 exit /b 1

echo.
echo [3/3] Done!
echo.
echo Executable: dist\DRB_Subjective_Tool.exe
echo.
echo Double-click the .exe to launch the tool in your browser.
echo Session files are saved in a "sessions" folder next to the .exe.
echo.
pause
