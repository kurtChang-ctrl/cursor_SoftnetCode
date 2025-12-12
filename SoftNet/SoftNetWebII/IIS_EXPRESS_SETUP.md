# IIS Express 配置說明

## 已完成的配置

1. ✅ `launchSettings.json` 已配置 IIS Express profile
2. ✅ `.vs\SoftNet\config\applicationhost.config` 已存在並配置了 SoftNetWebII 站點
3. ✅ 站點綁定到端口 24353 (localhost 和 192.168.66.3)

## 安裝 IIS Express

### 方法 1: 通過 Visual Studio Installer（推薦）

1. 打開 **Visual Studio Installer**
2. 點擊 **修改** (Modify)
3. 在 **個別元件** (Individual Components) 標籤中
4. 搜尋並勾選 **IIS Express**
5. 點擊 **修改** 完成安裝

### 方法 2: 單獨下載安裝

1. 下載 IIS 10.0 Express:
   - 下載連結: https://www.microsoft.com/en-us/download/details.aspx?id=48264
2. 執行安裝程式並完成安裝

## 安裝 ASP.NET Core Hosting Bundle

IIS Express 需要 ASP.NET Core Module 才能運行 ASP.NET Core 應用程式：

1. 下載 .NET Hosting Bundle:
   - 下載連結: https://dotnet.microsoft.com/download/dotnet
   - 選擇對應的 .NET 10.0 版本
2. 執行安裝程式（需要管理員權限）

## 啟動方式

### 方式 1: 使用 Visual Studio（最簡單）

1. 在 Visual Studio 中打開專案
2. 在工具列選擇 **IIS Express** 作為啟動配置
3. 按 F5 或點擊 **啟動** 按鈕

### 方式 2: 使用 PowerShell 腳本

執行專案根目錄下的 `start-iisexpress.ps1`:

```powershell
cd c:\04_SoftNet\SoftNet\SoftNetWebII
.\start-iisexpress.ps1
```

### 方式 3: 手動啟動 IIS Express

如果 IIS Express 已安裝，可以手動執行：

```powershell
& "C:\Program Files\IIS Express\iisexpress.exe" /config:"C:\04_SoftNet\SoftNet\.vs\SoftNet\config\applicationhost.config" /site:"SoftNetWebII"
```

## 配置說明

### launchSettings.json

- **IIS Express profile**: 配置了 `commandName: "IISExpress"`
- **端口**: 24353
- **環境變數**: Development

### applicationhost.config

- **站點名稱**: SoftNetWebII
- **應用程式池**: SoftNetWebII AppPool
- **綁定**: 
  - `http://localhost:24353`
  - `http://192.168.66.3:24353`
- **ASP.NET Core 模組**: 已配置為使用 AspNetCoreModuleV2

## 驗證安裝

執行以下命令檢查 IIS Express 是否已安裝：

```powershell
Test-Path "C:\Program Files\IIS Express\iisexpress.exe"
```

如果返回 `True`，表示已安裝。

## 疑難排解

### 問題 1: 找不到 IIS Express

**解決方案**: 按照上述安裝步驟安裝 IIS Express

### 問題 2: 502.5 錯誤

**原因**: 缺少 ASP.NET Core Hosting Bundle

**解決方案**: 安裝 .NET Hosting Bundle

### 問題 3: 端口已被占用

**解決方案**: 
1. 檢查端口占用: `netstat -ano | findstr :24353`
2. 停止占用端口的進程
3. 或修改 `launchSettings.json` 中的端口號

### 問題 4: 找不到 applicationhost.config

**解決方案**: 
1. 在 Visual Studio 中打開專案
2. Visual Studio 會自動生成 `.vs` 目錄和配置文件

## 注意事項

- IIS Express 配置檔案位於 `.vs` 目錄，此目錄通常不會被提交到版本控制
- 如果刪除 `.vs` 目錄，Visual Studio 會在下次打開專案時重新生成
- 確保已安裝對應版本的 .NET Hosting Bundle


