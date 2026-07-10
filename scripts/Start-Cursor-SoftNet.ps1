$targetDir = "C:\04_SoftNet\SoftNet"
$cursorExeDefault = Join-Path $env:LOCALAPPDATA "Programs\Cursor\Cursor.exe"
$cursorExe = $null

if (Test-Path $cursorExeDefault) {
    $cursorExe = $cursorExeDefault
} else {
    $cursorCmd = Get-Command cursor -ErrorAction SilentlyContinue
    if ($cursorCmd) {
        $cursorExe = $cursorCmd.Source
    }
}

if (-not $cursorExe) {
    Write-Error "找不到 Cursor 可執行檔，請確認已安裝或將其加入 PATH。常見路徑：$cursorExeDefault"
    exit 1
}

Start-Process -FilePath $cursorExe -ArgumentList $targetDir