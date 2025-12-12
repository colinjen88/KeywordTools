<#
build_exe.ps1 - Build a Windows .exe for run_gui.py using PyInstaller

Usage:
  powershell -ExecutionPolicy Bypass -File .\build_exe.ps1 -OneFile -NoConsole -IconPath .\assets\app.ico -ExeName GSC_Keyword_Tool

Requirements:
  - Python 3.8+ installed and available in PATH
  - Recommended: create/activate a venv and install requirements.txt
  - Install PyInstaller: pip install pyinstaller

Notes:
  - Do NOT include service account JSON in the repo or embed it in the .exe.
  - The script will build under the current folder and place the final exe in the dist\ folder.
#>

[CmdletBinding()]
param(
    [switch]$OneFile = $true,
    [switch]$NoConsole = $true,
    [string]$IconPath = '',
    [string]$ExeName = 'GSC_Keyword_Tool',
    [string]$EntryScript = 'run_gui.py'
)

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Err($msg) { Write-Host "[ERROR] $msg" -ForegroundColor Red }

Write-Info "Start building exe: $ExeName from $EntryScript"

if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
    Write-Info "PyInstaller not found. Installing into the current environment..."
    python -m pip install --upgrade pip
    python -m pip install pyinstaller
}

$pyArgs = @()
if ($OneFile) { $pyArgs += '--onefile' }
if ($NoConsole) { $pyArgs += '--noconsole' }
$pyArgs += "--name"; $pyArgs += $ExeName

# 強制收錄所有 Google API 子模組與資源
$pyArgs += "--collect-all"; $pyArgs += "googleapiclient"
$pyArgs += "--collect-all"; $pyArgs += "google.oauth2"
$pyArgs += "--collect-all"; $pyArgs += "google.auth"

# Hidden imports may help include dynamic imports (ttkbootstrap etc.)
$pyArgs += "--hidden-import"; $pyArgs += "ttkbootstrap"
# 'ttk' is part of tkinter; ensure tkinter.ttk is included
$pyArgs += "--hidden-import"; $pyArgs += "tkinter"
$pyArgs += "--hidden-import"; $pyArgs += "tkinter.ttk"
$pyArgs += "--hidden-import"; $pyArgs += "pandas"
$pyArgs += "--hidden-import"; $pyArgs += "openpyxl"
# ensure CLI module is included in analysis so we can import it inside the exe
$pyArgs += "--hidden-import"; $pyArgs += "gsc_keyword_report"
# ensure google apis and oauth libs are included
$pyArgs += "--hidden-import"; $pyArgs += "google.oauth2"
$pyArgs += "--hidden-import"; $pyArgs += "googleapiclient"
$pyArgs += "--hidden-import"; $pyArgs += "google_auth_oauthlib"
$pyArgs += "--hidden-import"; $pyArgs += "google.auth"

# 補充 Google API 子模組，避免 PyInstaller 漏包
$pyArgs += "--hidden-import"; $pyArgs += "googleapiclient.discovery"
$pyArgs += "--hidden-import"; $pyArgs += "googleapiclient.errors"
$pyArgs += "--hidden-import"; $pyArgs += "google.oauth2.service_account"
$pyArgs += "--hidden-import"; $pyArgs += "google.auth.transport.requests"

# 強制收錄所有 Google API 子模組
$pyArgs += "--collect-submodules"; $pyArgs += "googleapiclient"
$pyArgs += "--collect-submodules"; $pyArgs += "google.oauth2"
$pyArgs += "--collect-submodules"; $pyArgs += "google.auth"

if ($IconPath -and (Test-Path $IconPath)) {
    Write-Info "Using icon: $IconPath"
    $pyArgs += "--icon"; $pyArgs += $IconPath
}

# Add any data files that should be packaged (eg CSV templates) - format: src;dest (Windows uses ;)
# Example: --add-data "allKeyWord_normalized.csv;."
# default include of sample CSV and keyword lists (do NOT include service account keys)
$pyArgs += "--add-data"; $pyArgs += "allKeyWord_normalized.csv;."
$pyArgs += "--add-data"; $pyArgs += "gsc_keyword_report_sample.csv;."
# include the CLI script so frozen exe can spawn or import it
$pyArgs += "--add-data"; $pyArgs += "gsc_keyword_report.py;."

# optionally include a folder named 'assets' if it exists
if (Test-Path .\assets) {
  Write-Info "Including assets folder"
  $pyArgs += "--add-data"; $pyArgs += "assets;assets"
}

# build dir cleanup
if (Test-Path .\dist) { Remove-Item -Recurse -Force .\dist }
if (Test-Path .\build) { Remove-Item -Recurse -Force .\build }
if (Test-Path ("$ExeName.spec")) { Remove-Item -Force "$ExeName.spec" }

Write-Info "Running PyInstaller..."
pyinstaller @pyArgs $EntryScript

if ($LASTEXITCODE -ne 0) {
    Write-Err "PyInstaller build failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}

Write-Info "Build complete. Executable is in .\dist\$ExeName.exe"
Write-Host "You should not bundle credentials with this executable. Keep the Service Account key out of repo and specify it in the GUI or via secure store." -ForegroundColor Yellow

Exit 0
