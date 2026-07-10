# Node.js 安裝腳本（需要管理員權限）
# 執行方式：以管理員身份執行 PowerShell，然後執行此腳本

Write-Host "=== Node.js 安裝腳本 ===" -ForegroundColor Cyan
Write-Host ""

# 檢查是否以管理員身份執行
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "❌ 此腳本需要管理員權限！" -ForegroundColor Red
    Write-Host ""
    Write-Host "請執行以下步驟：" -ForegroundColor Yellow
    Write-Host "1. 以滑鼠右鍵點擊 PowerShell" -ForegroundColor Yellow
    Write-Host "2. 選擇「以系統管理員身分執行」" -ForegroundColor Yellow
    Write-Host "3. 切換到此腳本所在目錄" -ForegroundColor Yellow
    Write-Host "4. 執行: .\install-nodejs.ps1" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "或直接執行以下命令：" -ForegroundColor Yellow
    Write-Host 'Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\install-nodejs.ps1`""' -ForegroundColor Cyan
    exit 1
}

Write-Host "✓ 已確認管理員權限" -ForegroundColor Green
Write-Host ""

# 檢查 Node.js 是否已安裝
$nodePath = "C:\Program Files\nodejs\node.exe"
if (Test-Path $nodePath) {
    $version = & $nodePath --version
    Write-Host "✓ Node.js 已安裝，版本: $version" -ForegroundColor Green
    Write-Host ""
    Write-Host "請重新啟動終端機後執行以下命令驗證：" -ForegroundColor Yellow
    Write-Host "  node --version" -ForegroundColor Cyan
    Write-Host "  npm --version" -ForegroundColor Cyan
    exit 0
}

# 下載 Node.js 安裝程式
Write-Host "正在下載 Node.js LTS 安裝程式..." -ForegroundColor Cyan
$url = "https://nodejs.org/dist/v20.11.0/node-v20.11.0-x64.msi"
$output = "$env:TEMP\nodejs-installer.msi"

try {
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing
    Write-Host "✓ 下載完成: $output" -ForegroundColor Green
    Write-Host "  檔案大小: $([math]::Round((Get-Item $output).Length / 1MB, 2)) MB" -ForegroundColor Gray
    Write-Host ""
} catch {
    Write-Host "❌ 下載失敗: $_" -ForegroundColor Red
    exit 1
}

# 執行安裝
Write-Host "正在安裝 Node.js（這可能需要幾分鐘）..." -ForegroundColor Cyan
$arguments = "/i `"$output`" /quiet /norestart ADDLOCAL=ALL"
$process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -PassThru

if ($process.ExitCode -eq 0) {
    Write-Host "✓ Node.js 安裝成功！" -ForegroundColor Green
    Write-Host ""
    
    # 更新環境變數
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
    
    # 驗證安裝
    if (Test-Path $nodePath) {
        $version = & $nodePath --version
        $npmVersion = & "C:\Program Files\nodejs\npm.cmd" --version
        Write-Host "✓ Node.js 版本: $version" -ForegroundColor Green
        Write-Host "✓ npm 版本: $npmVersion" -ForegroundColor Green
        Write-Host ""
        Write-Host "安裝完成！請重新啟動終端機或 Cursor 以使用 Node.js。" -ForegroundColor Green
    } else {
        Write-Host "⚠ 安裝完成，但請重新啟動終端機後再驗證" -ForegroundColor Yellow
    }
} else {
    Write-Host "❌ 安裝失敗，退出代碼: $($process.ExitCode)" -ForegroundColor Red
    Write-Host ""
    Write-Host "可能的解決方案：" -ForegroundColor Yellow
    Write-Host "1. 確認有足夠的磁碟空間" -ForegroundColor Yellow
    Write-Host "2. 檢查是否有其他安裝程式正在執行" -ForegroundColor Yellow
    Write-Host "3. 嘗試手動執行安裝程式: $output" -ForegroundColor Yellow
    exit 1
}

# 清理安裝檔案
Write-Host ""
Write-Host "是否要刪除安裝程式檔案？(Y/N): " -NoNewline -ForegroundColor Yellow
$response = Read-Host
if ($response -eq "Y" -or $response -eq "y") {
    Remove-Item $output -Force
    Write-Host "✓ 已刪除安裝程式檔案" -ForegroundColor Green
}
