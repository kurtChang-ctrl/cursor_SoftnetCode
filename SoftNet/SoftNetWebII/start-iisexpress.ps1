# IIS Express 啟動腳本
# 此腳本用於啟動 SoftNetWebII 專案使用 IIS Express

$projectPath = Split-Path -Parent $PSScriptRoot
$configPath = Join-Path $projectPath ".vs\SoftNet\config\applicationhost.config"
$siteName = "SoftNetWebII"
$port = "24353"

# 尋找 IIS Express 可執行文件
$iisExpressPaths = @(
    "${env:ProgramFiles}\IIS Express\iisexpress.exe",
    "${env:ProgramFiles(x86)}\IIS Express\iisexpress.exe",
    "${env:LOCALAPPDATA}\Programs\IIS Express\iisexpress.exe"
)

$iisExpress = $null
foreach ($path in $iisExpressPaths) {
    if (Test-Path $path) {
        $iisExpress = $path
        break
    }
}

# 如果找不到，嘗試在 Visual Studio 目錄中尋找
if (-not $iisExpress) {
    $vsPaths = Get-ChildItem "${env:ProgramFiles}\Microsoft Visual Studio" -Directory -ErrorAction SilentlyContinue
    foreach ($vsPath in $vsPaths) {
        $iisPath = Join-Path $vsPath.FullName "Common7\IDE\iisexpress.exe"
        if (Test-Path $iisPath) {
            $iisExpress = $iisPath
            break
        }
    }
}

if (-not $iisExpress) {
    Write-Host "錯誤: 找不到 IIS Express。請確保已安裝 IIS Express。" -ForegroundColor Red
    Write-Host ""
    Write-Host "安裝方式:" -ForegroundColor Yellow
    Write-Host "1. 通過 Visual Studio Installer 安裝 'IIS Express' 組件" -ForegroundColor Yellow
    Write-Host "2. 或下載並安裝 IIS 10.0 Express:" -ForegroundColor Yellow
    Write-Host "   https://www.microsoft.com/en-us/download/details.aspx?id=48264" -ForegroundColor Cyan
    exit 1
}

if (-not (Test-Path $configPath)) {
    Write-Host "錯誤: 找不到 applicationhost.config 文件: $configPath" -ForegroundColor Red
    Write-Host "請在 Visual Studio 中打開專案一次，以生成配置文件。" -ForegroundColor Yellow
    exit 1
}

Write-Host "啟動 IIS Express..." -ForegroundColor Green
Write-Host "專案路徑: $projectPath" -ForegroundColor Gray
Write-Host "配置路徑: $configPath" -ForegroundColor Gray
Write-Host "站點名稱: $siteName" -ForegroundColor Gray
Write-Host "端口: $port" -ForegroundColor Gray
Write-Host ""
Write-Host "訪問地址: http://localhost:$port" -ForegroundColor Cyan
Write-Host "按 Ctrl+C 停止 IIS Express" -ForegroundColor Yellow
Write-Host ""

# 啟動 IIS Express
& $iisExpress /config:"$configPath" /site:"$siteName" /apppool:"SoftNetWebII AppPool"


