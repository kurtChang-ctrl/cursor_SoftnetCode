@echo off
set TARGET_DIR=C:\04_SoftNet\SoftNet
set CURSOR_EXE=%LOCALAPPDATA%\Programs\Cursor\Cursor.exe

if exist "%CURSOR_EXE%" (
  start "" "%CURSOR_EXE%" "%TARGET_DIR%"
  exit /b 0
)

where cursor >nul 2>&1
if %errorlevel%==0 (
  for /f "delims=" %%i in ('where cursor') do (
    start "" "%%i" "%TARGET_DIR%"
    exit /b 0
  )
)

echo 找不到 Cursor 可執行檔，請確認已安裝或加入 PATH。
echo 常見路徑：%LOCALAPPDATA%\Programs\Cursor\Cursor.exe
exit /b 1