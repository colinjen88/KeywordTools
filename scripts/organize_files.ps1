# 檔案整理腳本
# 用途：將現有檔案移動到新的目錄結構

Write-Host "開始整理檔案..." -ForegroundColor Green

# 移動關鍵字文件
Write-Host "`n移動關鍵字文件..." -ForegroundColor Yellow
if (Test-Path "allKeyWord.csv") {
    Move-Item "allKeyWord.csv" "data\keywords\" -Force
    Write-Host "  ✓ allKeyWord.csv" -ForegroundColor Green
}
if (Test-Path "allKeyWord_normalized.csv") {
    Move-Item "allKeyWord_normalized.csv" "data\keywords\" -Force
    Write-Host "  ✓ allKeyWord_normalized.csv" -ForegroundColor Green
}

# 移動報表文件
Write-Host "`n移動報表文件..." -ForegroundColor Yellow
Get-ChildItem -Filter "gsc_keyword_report_*.csv" | ForEach-Object {
    Move-Item $_.FullName "data\reports\" -Force
    Write-Host "  ✓ $($_.Name)" -ForegroundColor Green
}

# 移動示例文件
Write-Host "`n移動示例文件..." -ForegroundColor Yellow
if (Test-Path "gsc_keyword_report_sample.csv") {
    Move-Item "gsc_keyword_report_sample.csv" "data\samples\" -Force
    Write-Host "  ✓ gsc_keyword_report_sample.csv" -ForegroundColor Green
}

# 移動配置文件
Write-Host "`n移動配置文件..." -ForegroundColor Yellow
if (Test-Path "favorites.json") {
    Move-Item "favorites.json" "config\" -Force
    Write-Host "  ✓ favorites.json" -ForegroundColor Green
}

# 移動腳本文件
Write-Host "`n移動腳本文件..." -ForegroundColor Yellow
if (Test-Path "build_exe.ps1") {
    Move-Item "build_exe.ps1" "scripts\" -Force
    Write-Host "  ✓ build_exe.ps1" -ForegroundColor Green
}
if (Test-Path "GSC_Keyword_Tool.spec") {
    Move-Item "GSC_Keyword_Tool.spec" "scripts\" -Force
    Write-Host "  ✓ GSC_Keyword_Tool.spec" -ForegroundColor Green
}
if (Test-Path "GSC_Keyword_Tool_Debug.spec") {
    Move-Item "GSC_Keyword_Tool_Debug.spec" "scripts\" -Force
    Write-Host "  ✓ GSC_Keyword_Tool_Debug.spec" -ForegroundColor Green
}

# 移動工具腳本
Write-Host "`n移動工具腳本..." -ForegroundColor Yellow
if (Test-Path "tools") {
    Get-ChildItem "tools\*.ps1" | ForEach-Object {
        Move-Item $_.FullName "scripts\" -Force
        Write-Host "  ✓ $($_.Name)" -ForegroundColor Green
    }
    Get-ChildItem "tools\*.py" | ForEach-Object {
        Move-Item $_.FullName "tests\" -Force
        Write-Host "  ✓ $($_.Name)" -ForegroundColor Green
    }
}

# 刪除臨時和測試文件
Write-Host "`n刪除臨時文件..." -ForegroundColor Yellow
$filesToDelete = @(
    "debug_kws.py",
    "debug_tkcalendar.py",
    "test_export.csv",
    "test_export.xlsx",
    "test_mock.csv",
    "test_row_export.csv",
    "gsc_keyword_report.csv",
    "tracked_json_files.txt",
    "KeywordsTool.py"
)

foreach ($file in $filesToDelete) {
    if (Test-Path $file) {
        Remove-Item $file -Force
        Write-Host "  ✗ $file" -ForegroundColor Red
    }
}

# 刪除空的 tools 目錄
if (Test-Path "tools") {
    if ((Get-ChildItem "tools").Count -eq 0) {
        Remove-Item "tools" -Force
        Write-Host "  ✗ tools/" -ForegroundColor Red
    }
}

Write-Host "`n檔案整理完成！" -ForegroundColor Green
Write-Host "`n請檢查以下目錄：" -ForegroundColor Cyan
Write-Host "  - data\keywords\" -ForegroundColor White
Write-Host "  - data\reports\" -ForegroundColor White
Write-Host "  - data\samples\" -ForegroundColor White
Write-Host "  - config\" -ForegroundColor White
Write-Host "  - scripts\" -ForegroundColor White
Write-Host "  - tests\" -ForegroundColor White
