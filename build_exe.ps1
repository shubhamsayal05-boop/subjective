# Build DRB_Subjective_Tool.exe on Windows (PowerShell alternative)
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "============================================"
Write-Host " DRB Subjective Drivability Tool - Build"
Write-Host "============================================"
Write-Host "Working folder: $(Get-Location)"

$log = Join-Path $PSScriptRoot "build.log"
"Build started $(Get-Date)" | Out-File $log -Encoding utf8

function Fail($msg) {
    Write-Host "ERROR: $msg" -ForegroundColor Red
    Write-Host "See $log for details."
    Read-Host "Press Enter to close"
    exit 1
}

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Fail "Python is not installed or not on PATH."
}

$required = @(
    "launcher.py", "app.py", "config.py", "storage.py", "excel_export.py",
    "table_image.py", "drb_tool.spec", "requirements.txt", "requirements-build.txt",
    "Subjective_SprdSheet_072926.xlsm", "BEV_Subjective_SprdSheet_072926.xlsm",
    "CVT Subjective_SprdSheet_072926.xlsm"
)
foreach ($f in $required) {
    if (-not (Test-Path $f)) { Fail "Missing required file: $f" }
}

Write-Host "[1/3] Installing dependencies..."
python -m pip install --upgrade pip 2>&1 | Tee-Object -FilePath $log -Append
python -m pip install -r requirements.txt -r requirements-build.txt 2>&1 | Tee-Object -FilePath $log -Append

Write-Host "[2/3] Building Windows executable (5-10 minutes)..."
python -m PyInstaller --noconfirm --clean drb_tool.spec 2>&1 | Tee-Object -FilePath $log -Append
if ($LASTEXITCODE -ne 0) { Fail "PyInstaller build failed." }

$exe = Join-Path $PSScriptRoot "dist\DRB_Subjective_Tool.exe"
if (-not (Test-Path $exe)) { Fail "dist\DRB_Subjective_Tool.exe was not created." }

Write-Host "[3/3] Done!"
Write-Host "Executable: $exe"
Write-Host "Size: $((Get-Item $exe).Length) bytes"
Read-Host "Press Enter to close"
