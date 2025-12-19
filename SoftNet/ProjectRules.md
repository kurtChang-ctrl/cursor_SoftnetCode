# SoftNet 專案規範（Project Rules）

目的：為本解（Solution）提供一致、可維護的工程規範，讓新舊成員在開發、部署與維運上遵循同一套流程與約定。

---

## 1) 專案架構與啟動
- 主要 Web 專案：`SoftNetWebII`
- 共用函式庫：`Base`（DBADO、Models、Services）、`SoftNetCommLib`
- 啟動流程：`SoftNetWebII/Program.cs` → `SoftNetWebII/Startup.cs`
- 長期背景任務：`SoftNetWebII/Services/RUNTimeServer.cs` 由 `services.AddHostedService<RUNTimeServer>()` 啟用

## 2) 建置與執行（指令）
- 建置整個解：
  ```powershell
  dotnet build SoftNet.sln
  ```
- 發佈：
  ```powershell
  dotnet publish SoftNet.sln
  ```
- 開發熱重載（Web 專案）：
  ```powershell
  $env:ASPNETCORE_ENVIRONMENT = "Development"
  dotnet watch run --project SoftNetWebII
  ```
- 單專案執行（無熱重載）：
  ```powershell
  $env:ASPNETCORE_ENVIRONMENT = "Development"
  dotnet run --project SoftNetWebII
  ```

## 3) 設定與組態（FunConfig）
- 設定來源：「appsettings.json」的 `FunConfig` section 在 `Startup.cs` 綁定到 `_Fun.Config`
- 規範：
  - JSON 命名採 PascalCase（`AddNewtonsoftJson(opts => opts.UseMemberCasing())` 並將 `JsonSerializerOptions.PropertyNamingPolicy = null`）
  - 變更行為前先檢視 `FunConfig` 欄位，避免硬編碼
  - 機敏資訊（例如 `MailCredentialsPWD`）請使用 `dotnet user-secrets` 或環境變數；勿提交至版本控制
- 開發環境覆寫：使用 `appsettings.Development.json` 或 `ASPNETCORE_ENVIRONMENT=Development`

## 4) 資料庫規範（DB）
- 連線：使用 `Microsoft.Data.SqlClient`，透過 DI 註冊 `DbConnection/DbCommand`
- 共用 DB 層：參考 `Base/DBADO.cs` 與 `SoftNetCommLib/DBADO.cs`；必要時同步維護兩處
- 參數化查詢：禁止字串拼接 SQL；一律使用參數化方法
- 既定表操作（強制規範）：
  - `APS_WarningData`：一律呼叫 `SFC_Common.InsertWarningData`
  - `APS_EventLog`：一律呼叫 `SFC_Common.InsertApsEventLog`
  - `APS_Simulation_ErrorData`（ErrorType 15）：一律呼叫 `SFC_Common.InsertSimulationError`
- Schema 變更（建議流程）：
  1. 先於測試環境驗證（含資料遷移腳本）
  2. 更新 Models/DTO 與 DBADO 相關方法
  3. 撰寫升級/回滾腳本並附簡要說明
  4. 送出 PR 並在描述中標示影響範圍與回滾策略

## 5) 背景服務與併發
- 背景任務：建立 `BackgroundService` 類別，於 `Startup.cs` 以 `services.AddHostedService<YourService>()` 註冊
- 併發與同步：
  - 保護共享資源（如 `_MasterRMSUserList`）使用適當鎖或 `ConcurrentDictionary`
  - 避免在鎖內執行 I/O 或長工作；改以背景 Task 排程
- TCP/Socket：`RUNTimeServer` 處理封包可能黏包/分片；一律採累積緩衝與安全解析
- 測試：使用 telnet 或簡易 TCP client 模擬設備連線與封包分片

## 6) 日誌與除錯
- 主要日誌目錄：`SoftNetWebII/_log/`
- 常見檔案：`Socket5431Log.txt`、`sql.txt`、`error.txt`
- 問題排查：先檢查 `_log`，再比對 `_Fun.Config` 與網路設定（port 與防火牆）

## 7) 程式碼風格與品質
- C# 命名：類/屬性/JSON 採 PascalCase；區域變數/參數採 camelCase
- 不使用一字母變數名（除迴圈索引外）；避免內嵌長 SQL 或複雜邏輯
- 盡量維持現有風格與 API；避免大幅重構未經驗證的舊碼
- 警告處理：
  - 平台特定 API（CA1416）：以條件編譯或環境判斷包覆
  - 過時 API（SYSLIB0014：`WebRequest`）：改用 `HttpClient`
  - `CS0162/CS0414`：清理無用或不可達程式碼

## 8) Pull Request（PR）與版本管控
- 分支策略（建議）：
  - `main`：穩定分支；發佈版本
  - `dev`：整合分支；日常開發
  - feature/hotfix 分支：以明確名稱與 Issue 編號命名
- 提交訊息：`<type>(scope): <subject>`，例如：`feat(RUNTimeServer): route APS_EventLog via SFC_Common`
- PR 檢核清單：
  - 是否遵循 DB 規範（三表 insert 走 `SFC_Common`）
  - 是否維持 JSON PascalCase 與 FunConfig 規範
  - 是否提供必要的設定/部署說明與回滾策略
  - 是否通過 `dotnet build`，且新警告可被合理說明

## 9) 安全與機密
- 機敏設定絕不進 repo：改用 `user-secrets` 或安全管理工具
- 日誌避免個資與密碼；必要時做遮蔽或脫敏
- 外部服務憑證使用環境變數或安全存放機制

## 10) 常用參考
- 啟動與配置：`SoftNetWebII/Program.cs`、`SoftNetWebII/Startup.cs`、`SoftNetWebII/SoftNetWebII.csproj`
- 長期任務範例：`SoftNetWebII/Services/RUNTimeServer.cs`
- 共用 DB 層：`Base/DBADO.cs`、`SoftNetCommLib/DBADO.cs`
- 運行紀錄：`SoftNetWebII/_log/`

---

如需更進階的指引（例如 `RUNTimeServer` 的 TCP/Socket 封包流程、`SFC_Common` 的 DB 呼叫最佳實務、或資料庫 schema 變更 SOP），請提出需求，我們可補充一頁快速參考。