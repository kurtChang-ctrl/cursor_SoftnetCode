# ONNX 模型下載腳本
# 此腳本會自動下載 all-MiniLM-L6-v2 模型的 ONNX 版本

$ErrorActionPreference = "Stop"

# 設定變數
$modelUrl = "https://huggingface.co/sentence-transformers/all-MiniLM-L6-v2/resolve/main/onnx/model.onnx"
$vocabUrl = "https://huggingface.co/sentence-transformers/all-MiniLM-L6-v2/resolve/main/vocab.txt"
$outputDir = Join-Path $PSScriptRoot
$outputFile = Join-Path $outputDir "all-MiniLM-L6-v2.onnx"
$vocabFile = Join-Path $outputDir "vocab.txt"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "ONNX 模型下載腳本" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 檢查輸出目錄
if (-not (Test-Path $outputDir)) {
    Write-Host "建立目錄: $outputDir" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# 檢查檔案是否已存在
if (Test-Path $outputFile) {
    $fileSize = (Get-Item $outputFile).Length / 1MB
    Write-Host "模型檔案已存在: $outputFile" -ForegroundColor Yellow
    Write-Host "檔案大小: $([math]::Round($fileSize, 2)) MB" -ForegroundColor Yellow
    $overwrite = Read-Host "是否要重新下載？(Y/N)"
    if ($overwrite -ne "Y" -and $overwrite -ne "y") {
        Write-Host "已取消下載。" -ForegroundColor Green
        exit 0
    }
    Remove-Item $outputFile -Force
}

Write-Host "開始下載模型..." -ForegroundColor Green
Write-Host "來源: $modelUrl" -ForegroundColor Gray
Write-Host "目標: $outputFile" -ForegroundColor Gray
Write-Host ""

try {
    # 使用 Invoke-WebRequest 下載
    $ProgressPreference = 'SilentlyContinue'
    
    # 下載 ONNX 模型
    Write-Host "下載 ONNX 模型..." -ForegroundColor Green
    Invoke-WebRequest -Uri $modelUrl -OutFile $outputFile -UseBasicParsing
    
    # 下載 vocab.txt
    Write-Host "下載 vocab.txt..." -ForegroundColor Green
    Invoke-WebRequest -Uri $vocabUrl -OutFile $vocabFile -UseBasicParsing
    
    # 驗證檔案
    $allFilesExist = $true
    if (Test-Path $outputFile) {
        $fileSize = (Get-Item $outputFile).Length / 1MB
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "下載完成！" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "ONNX 模型:" -ForegroundColor White
        Write-Host "  檔案位置: $outputFile" -ForegroundColor Gray
        Write-Host "  檔案大小: $([math]::Round($fileSize, 2)) MB" -ForegroundColor Gray
    }
    else {
        Write-Host "警告: ONNX 模型檔案下載失敗" -ForegroundColor Yellow
        $allFilesExist = $false
    }
    
    if (Test-Path $vocabFile) {
        $vocabSize = (Get-Item $vocabFile).Length / 1KB
        Write-Host "vocab.txt:" -ForegroundColor White
        Write-Host "  檔案位置: $vocabFile" -ForegroundColor Gray
        Write-Host "  檔案大小: $([math]::Round($vocabSize, 2)) KB" -ForegroundColor Gray
    }
    else {
        Write-Host "警告: vocab.txt 檔案下載失敗" -ForegroundColor Yellow
        $allFilesExist = $false
    }
    
    Write-Host ""
    if ($allFilesExist) {
        Write-Host "所有檔案已準備就緒，可以開始使用 RAG 功能（包含完整的 BERT tokenizer）。" -ForegroundColor Green
    }
    else {
        Write-Host "部分檔案下載失敗，請檢查網路連線或手動下載。" -ForegroundColor Yellow
    }
}
catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "下載失敗！" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "錯誤訊息: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "請手動下載檔案：" -ForegroundColor Yellow
    Write-Host "1. 前往: https://huggingface.co/sentence-transformers/all-MiniLM-L6-v2" -ForegroundColor Yellow
    Write-Host "2. 下載以下檔案：" -ForegroundColor Yellow
    Write-Host "   - onnx/model.onnx → 重新命名為 all-MiniLM-L6-v2.onnx" -ForegroundColor Yellow
    Write-Host "   - vocab.txt" -ForegroundColor Yellow
    Write-Host "3. 放置到: $outputDir" -ForegroundColor Yellow
    exit 1
}
