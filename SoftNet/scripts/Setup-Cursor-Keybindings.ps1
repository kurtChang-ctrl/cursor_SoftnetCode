# Cursor 快捷鍵自動設定腳本
# 此腳本會自動將折疊快捷鍵設定加入到 Cursor 的使用者設定檔中

$cursorUserDir = Join-Path $env:APPDATA "Cursor\User"
$keybindingsFile = Join-Path $cursorUserDir "keybindings.json"

Write-Host "正在設定 Cursor 折疊快捷鍵..." -ForegroundColor Green

# 確保目錄存在
if (-not (Test-Path $cursorUserDir)) {
    New-Item -ItemType Directory -Path $cursorUserDir -Force | Out-Null
    Write-Host "已建立 Cursor 使用者目錄: $cursorUserDir" -ForegroundColor Yellow
}

# 定義要加入的快捷鍵設定
$newKeybindings = @(
    @{
        key = "ctrl+k ctrl+0"
        command = "editor.foldAll"
        when = "editorTextFocus"
    },
    @{
        key = "ctrl+k ctrl+j"
        command = "editor.unfoldAll"
        when = "editorTextFocus"
    },
    @{
        key = "ctrl+k ctrl+1"
        command = "editor.foldLevel1"
        when = "editorTextFocus"
    },
    @{
        key = "ctrl+k ctrl+2"
        command = "editor.foldLevel2"
        when = "editorTextFocus"
    },
    @{
        key = "ctrl+k ctrl+3"
        command = "editor.foldLevel3"
        when = "editorTextFocus"
    }
)

# 讀取現有的 keybindings.json（如果存在）
$existingKeybindings = @()
if (Test-Path $keybindingsFile) {
    try {
        $content = Get-Content $keybindingsFile -Raw -Encoding UTF8
        if ($content.Trim()) {
            $existingKeybindings = ConvertFrom-Json $content
            if (-not $existingKeybindings) {
                $existingKeybindings = @()
            }
        }
        Write-Host "已讀取現有的快捷鍵設定檔" -ForegroundColor Cyan
    }
    catch {
        Write-Host "警告: 無法讀取現有的 keybindings.json，將建立新檔案" -ForegroundColor Yellow
        $existingKeybindings = @()
    }
}

# 檢查並移除重複的快捷鍵（避免重複設定）
$commandsToAdd = @("editor.foldAll", "editor.unfoldAll", "editor.foldLevel1", "editor.foldLevel2", "editor.foldLevel3")
$filteredExisting = $existingKeybindings | Where-Object { 
    $_.command -notin $commandsToAdd -or $_.key -notmatch "ctrl\+k ctrl\+[0-3j]"
}

# 合併現有和新的快捷鍵設定
$allKeybindings = @()
if ($filteredExisting) {
    $allKeybindings += $filteredExisting
}
$allKeybindings += $newKeybindings

# 轉換為 JSON 並寫入檔案
try {
    $jsonContent = $allKeybindings | ConvertTo-Json -Depth 10
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($keybindingsFile, $jsonContent, $utf8NoBom)
    
    Write-Host "`n✅ 快捷鍵設定完成！" -ForegroundColor Green
    Write-Host "設定檔位置: $keybindingsFile" -ForegroundColor Cyan
    Write-Host "`n已設定的快捷鍵:" -ForegroundColor Yellow
    Write-Host "  - Ctrl+K, Ctrl+0  → 折疊所有區塊" -ForegroundColor White
    Write-Host "  - Ctrl+K, Ctrl+J  → 展開所有區塊" -ForegroundColor White
    Write-Host "  - Ctrl+K, Ctrl+1  → 折疊到第 1 層級" -ForegroundColor White
    Write-Host "  - Ctrl+K, Ctrl+2  → 折疊到第 2 層級" -ForegroundColor White
    Write-Host "  - Ctrl+K, Ctrl+3  → 折疊到第 3 層級" -ForegroundColor White
    Write-Host "`n請重新載入 Cursor 視窗（Ctrl+Shift+P → Reload Window）以套用設定" -ForegroundColor Yellow
}
catch {
    Write-Host "`n❌ 錯誤: 無法寫入 keybindings.json" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "`n請手動編輯: $keybindingsFile" -ForegroundColor Yellow
    exit 1
}

