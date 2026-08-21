@echo off
setlocal enabledelayedexpansion

rem Always run from the folder that contains this script.
cd /d "%~dp0"

echo ============================================
echo  DRB Subjective Drivability Tool - Build
echo ============================================
echo.
echo Working folder: %CD%
echo.

set "LOG=build.log"
echo Build started %DATE% %TIME% > "%LOG%"

where python >> "%LOG%" 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not on PATH.
    echo Install Python 3.10+ from https://www.python.org/downloads/
    echo During install, check "Add python.exe to PATH".
    echo See %LOG% for details.
    pause
    exit /b 1
)

for %%F in (launcher.py app.py config.py storage.py excel_export.py table_image.py drb_tool.spec requirements.txt requirements-build.txt) do (
    if not exist "%%F" (
        echo ERROR: Missing required file: %%F
        echo Make sure you downloaded the full project folder, not just app.py.
        pause
        exit /b 1
    )
)

for %%F in ("Subjective_SprdSheet_072926.xlsm" "BEV_Subjective_SprdSheet_072926.xlsm" "CVT Subjective_SprdSheet_072926.xlsm") do (
    if not exist %%F (
        echo ERROR: Missing Excel template: %%F
        pause
        exit /b 1
    )
)

echo [1/3] Installing dependencies...
python -m pip install --upgrade pip >> "%LOG%" 2>&1
if errorlevel 1 (
    echo ERROR: pip upgrade failed. See %LOG%
    pause
    exit /b 1
)

python -m pip install -r requirements.txt -r requirements-build.txt >> "%LOG%" 2>&1
if errorlevel 1 (
    echo ERROR: dependency install failed. See %LOG%
    pause
    exit /b 1
)

echo.
echo [2/3] Building Windows executable (5-10 minutes, please wait)...
python -m PyInstaller --noconfirm --clean drb_tool.spec >> "%LOG%" 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed.
    echo Open %LOG% in Notepad to see the full error.
    echo.
    type "%LOG%"
    pause
    exit /b 1
)

if not exist "dist\DRB_Subjective_Tool.exe" (
    echo ERROR: Build finished but dist\DRB_Subjective_Tool.exe was not created.
    echo See %LOG%
    pause
    exit /b 1
)

echo.
echo [3/3] Done!
echo.
echo Executable: %CD%\dist\DRB_Subjective_Tool.exe
for %%A in ("dist\DRB_Subjective_Tool.exe") do echo Size: %%~zA bytes
echo.
echo Double-click the .exe to launch the tool in your browser.
echo Session files are saved in a "sessions" folder next to the .exe.
echo Full build log: %CD%\%LOG%
echo.
pause
