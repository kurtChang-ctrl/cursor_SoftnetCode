# Cursor 自動上下文（Auto Context）與 @folders 使用說明

目的：開啟 Cursor 時自動納入工作目錄內容，並快速加入本專案規範檔。

## 一鍵開啟 Cursor 指向本專案
- PowerShell：
  ```powershell
  .\scripts\Start-Cursor-SoftNet.ps1
  ```
- CMD：
  ```bat
  scripts\Start-Cursor-SoftNet.bat
  ```

## 自動上下文設定（已配置）

專案已包含以下配置檔案，可自動提供上下文：

### 1. `.cursorrules` 檔案
- 位置：專案根目錄
- 功能：定義專案規則和上下文提示
- 說明：Cursor 會自動讀取此檔案，AI 助手會知道專案位置和結構

### 2. `.cursor/rules/auto-context.mdc` 檔案
- 位置：`.cursor/rules/` 目錄
- 功能：定義自動包含的檔案模式
- 說明：自動包含所有 `.cs`、`.cshtml`、`.js`、`.json`、`.md`、`.csproj`、`.sln` 檔案

## 建議的 Cursor 設定（依版本介面可能不同）
- 啟用自動上下文：在 Cursor 設定中啟用「Auto Context」或「自動加入工作目錄至對話」。
- 啟用完整資料夾內容：在 Cursor 設定 > Features 中啟用「Full Folder Content」選項。
- 之後在對話可直接使用：
  - `@folders "C:\04_SoftNet\SoftNet"` - 包含整個專案目錄
  - `@ProjectRules.md` - 附加專案規範文件

## 小技巧
- 若未自動加入，可在第一則訊息貼上：
  - `@folders "C:\04_SoftNet\SoftNet"` 並附上 `@ProjectRules.md`
- 若想指定子資料夾（例如僅 Web 專案）：
  - `@folders "C:\04_SoftNet\SoftNet\SoftNetWebII"`

## 配置檔案說明
- `.cursorrules` - 專案規則和上下文提示（舊版方式，仍可用）
- `.cursor/rules/auto-context.mdc` - 自動上下文規則（新版方式）
- `.cursorignore` - 排除不需要索引的檔案和目錄

> 註：不同 Cursor 版本對「自動上下文」支援程度不一，若此功能不可用，請以上述一鍵腳本開啟並手動加入 `@folders` 與 `@ProjectRules.md`。