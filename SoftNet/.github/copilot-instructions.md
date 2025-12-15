# Copilot 指令摘要（給 AI 編碼助理）

目的：快速讓 AI 編碼助理立刻上手本 repo、理解主要架構、開發流程與專案慣例。

1) 大致架構
- 本解包含多個 project：`SoftNetWebII`（主要的 ASP.NET Core WebApp）、`Base`（共用函式庫、DBADO、Models、Services）、`SoftNetCommLib`（共用 commlib）。
- Web 啟動點：`SoftNetWebII/Program.cs` -> `SoftNetWebII/Startup.cs`（註冊服務、HostedService）。
- 長期背景工作：`SoftNetWebII/Services/RUNTimeServer.cs` 為 `BackgroundService`，以 `services.AddHostedService<RUNTimeServer>()` 啟用，負責 TCP/Socket 與排程工作。

2) 重要開發流程 / 指令（可直接執行）
- 建置整個方案（解）：
  - `dotnet build SoftNet.sln`
- 發佈：
  - `dotnet publish SoftNet.sln`
- 本機開發熱重載（Web 專案）：
  - `dotnet watch run --project SoftNetWebII` 或使用 workspace task `watch`（若在 VS Code 可執行）。
- 執行單一專案（開發）：
  - `dotnet run --project SoftNetWebII`（可先設定 `ASPNETCORE_ENVIRONMENT=Development`）

3) 專案慣例與模式（非通用建議）
- 設定來源：appsettings.json 的 `FunConfig` section 會 bind 到 `_Fun.Config`（在 `Startup.cs` 設定）。調整行為時先看 `FunConfig` 欄位。
- JSON 序列化：專案顯式改為 PascalCase（`AddNewtonsoftJson(opts => opts.UseMemberCasing())` 與 `JsonSerializerOptions.PropertyNamingPolicy = null`），請產生符合這個命名約定的輸出/測試資料。
- 共用 DB 層/工具：資料存取相關邏輯散在 `Base/DBADO.cs` 與 `SoftNetCommLib/DBADO.cs`，新改動應檢查兩處是否需同步。
- BackgroundService 與 Singleton：長期任務放在 `Services` 下並以 `AddHostedService` 或 `AddSingleton` 註冊。範例：`SNWebSocketService` 與 `SFC_Common` 在 `Startup.cs` 以 singleton 建立。

4) 整合點與外部相依
- SQL Server：使用 `Microsoft.Data.SqlClient`，連線透過 DI 註冊 `DbConnection/DbCommand`。
- RabbitMQ 客戶端包會出現在 packages 目錄（如需要搜尋使用情形）。
- WebSocket：`SNWebSocketService` 為內建 WebSocket server，監聽設定來源為 `_Fun.Config.WesocketPort`。
- 日誌：Web 專案會產生 `_log/` 目錄（例：SQL 與 Socket log），調查錯誤時直接查看 `SoftNetWebII/_log`。

5) 代碼產生/修改的具體提示
- 若要新增長期背景任務：新增 `BackgroundService` 為 class，並在 `Startup.cs` 的 `ConfigureServices` 呼叫 `services.AddHostedService<YourService>()`。
- 如果變更 DB 連線或 config key：更新 `appsettings.json` 的 `FunConfig`，並留意 `_Fun.Config` 的使用處。
- 增加 API Controller：放在 `SoftNetWebII/Controllers`，路由慣例使用 MVC default route（Startup 中定義）。

6) 參考檔案（請優先閱讀）
- 啟動與配置：[SoftNetWebII/Program.cs](SoftNetWebII/Program.cs) 、[SoftNetWebII/Startup.cs](SoftNetWebII/Startup.cs) 、[SoftNetWebII/SoftNetWebII.csproj](SoftNetWebII/SoftNetWebII.csproj)
- 長期任務範例：[SoftNetWebII/Services/RUNTimeServer.cs](SoftNetWebII/Services/RUNTimeServer.cs)
- 共用函式庫與 DB：`Base/DBADO.cs`、`SoftNetCommLib/DBADO.cs`、`Base/Services`（查找服務實作）
- 運行紀錄：`SoftNetWebII/_log/`（觀察 socket/sql 日誌）

7) 其他注意事項
- 代碼中有許多歷史或被註解的行，修改時注意保留原本行為或確認落實測試。
- 在變更共用專案（`Base`、`BaseWeb`、`BaseApi`）時，先以 `dotnet build SoftNet.sln` 確認編譯可過，再測試 `SoftNetWebII`。

若你想要我把某段更深入的運作流程（例如 `RUNTimeServer` 的 TCP 流程或 `SFC_Common` 的 DB 呼叫）寫成更詳細指引，我可以接著把該段落展開成 1-page 的快速參考。請告訴我你要先聚焦哪個元件。

---

## RUNTimeServer TCP/Socket 快速參考
- 位置：[SoftNetWebII/Services/RUNTimeServer.cs](SoftNetWebII/Services/RUNTimeServer.cs)
- 目的：一個以 `BackgroundService` 實作的長期任務，管理 TcpListener、Socket 連線、排程與與外部系統（如 5431 設備、RMS）的即時通訊。
- 啟動方式：由 `Startup.cs` 呼叫 `services.AddHostedService<RUNTimeServer>()` 自動啟動。
- 關鍵流程（閱讀 `MasterTcpListenerThread`、`MasterProcessRequest` 等方法）:
  - 建立 `TcpListener` 並 `Start()`，進入 Accept loop（`AcceptSocket()`），為每個 client 建立 background thread 處理 `MasterProcessRequest`。
  - 使用字典 ` _MasterRMSUserList` 追蹤已連線設備（key = ip:port）。
  - 常見同步點：`lock__MasterRMSUserList`、`lock_logID`，請注意鎖的範圍以避免死鎖或性能瓶頸。
- 日誌與除錯：檢查 `SoftNetWebII/_log/`（例：Socket5431Log、sql.txt、error.txt）。若 socket 無法連線，建議：
  - 檢查防火牆與服務監聽 port（`netstat -ano | findstr :<port>`）。
  - 確認 `_Fun.Config` 中的 IP/Port 設定。
  - 使用短時間偵錯日誌（在 `RUNTimeServer` 中新增臨時 `Console.WriteLine` 或 NLog 記錄）。
- 常見錯誤與處置：
  - SocketException：通常為 port 被佔用或網路中斷 -> 檢查其他程式佔用、重啟 Listener。
  - 多執行緒 race：若發現字典 key collision 或 NullReference，請先加鎖或改用 `ConcurrentDictionary`。
- 測試建議：在本機以 telnet 或簡易 Tcp client 模擬設備連線，驗證 Accept 與 Request 處理流程。

### `MasterProcessRequest` 詳解（重點、流程與注意事項）
- 位置：`RUNTimeServer.MasterProcessRequest(object socket)`。
- 流程要點：
  1. 以 `Socket.ReceiveFrom` 讀取原始 bytes，累積到 `receiveBuffer`（List<byte>）。
  2. 每筆封包前 4 bytes 為 int 長度 `len`，第 5 byte 為 `type`，接著是 `len` bytes 的 payload。可能會有黏包或分片，因此以累積緩衝處理多個封包。
  3. 根據 `type` 分派處理：
     - 0: ping (忽略)
     - 1: 字串 -> 解 UTF8 並以 `,` 切割，呼叫 `Master_ResolveData2String`
     - 252: 自訂二進位處理 -> `Master_ResolveData2252`
     - 253: 檔案 -> `Master_ResolveData2253`
     - 254: JSON RMSProtocol -> 反序列化後呼叫 `Master_ResolveData2RMSProtocol`
  4. 處理採 Task 或 `SNThreadScheduler` 執行背景工作，避免在接收執行大量同步計算。
  5. 異常處理：捕捉 `SocketException`（常見 ErrorCode 10053 為連線被遠端關閉），與一般 Exception，並遞增全域計數 `RMSDBErrorCount`。

- 同步與資源：
  - 使用 `lock_receiveBuffer` 保護對 `receiveBuffer` 的修改；使用 `lock__MasterRMSUserList` 保護 `_MasterRMSUserList` 的增刪。
  - 關閉 socket 時會移除 `_MasterRMSUserList` 中對應的 `ip:port` 條目並呼叫 `Dispose()`。

- 常見陷阱與修正建議：
  - 黏包/分片：請不要假設每次 `ReceiveFrom` 回傳一完整封包，測試需包含跨多次 `Receive` 的封包重組。
  - 錯誤代碼處理：`SocketException` 除 10053 外都會記錄為錯誤，必要時增加更詳細 log（包含 `sex.ErrorCode` 與 stacktrace）。
  - 鎖範圍：避免在鎖內呼叫外部同步或長時間工作（例如 DB 呼叫）；應在鎖外派發 Task。
  - 資料競爭：若連線數成長，改用 `ConcurrentDictionary`、`ConcurrentQueue<byte[]>` 等非阻塞結構以提升可擴充性。

### 建議的測試用例（可自動化或手動）
- TC1: 建立 TCP 連線 -> 傳送 type=0 (ping) -> 服務不回應錯誤、連線保持。
- TC2: 傳送 type=1 的字串 `IIS_Login,DeviceA` -> 確認 `_MasterRMSUserList` 新增 key `ip:port` 與 deviceName 為 `DeviceA`。
- TC3: 傳送 type=254 的 JSON（RMSProtocol） -> 確認 `Master_ResolveData2RMSProtocol` 被執行（可在方法加入暫時性 log）
- TC4: 傳送分片封包（先送長度/部分 payload，再送剩餘 payload）-> 驗證完整 payload 被重組並處理。
- TC5: 傳送未定義的 `type`（例如 99）-> `RMSDBErrorCount` 增加且不當掉服務。
- TC6: 大量併發客戶端（50+）連線並傳送封包 -> 觀察 memory/CPU 與 `_MasterRMSUserList` 行為。
- TC7: 模擬 SocketException（關閉 client）-> 確認 server 釋放資源並移除 user list 條目。

---

## `_Fun.Config` 快速參考（appsettings.json -> FunConfig）
- 綁定位置：在 `Startup.cs` 有段程式碼：

  var config = new ConfigDto();
  Configuration.GetSection("FunConfig").Bind(config);
  _Fun.Config = config;

-- 常見欄位（從 `appsettings.json` 實際抽出，按出現順序）：
  - `Db`: 資料庫連線字串（範例含 Integrated Security / UserId 密碼）。供 `SFC_Common`、`DBADO` 使用。
  - `Locale`: 區域設定（例如 `zh-TW`），於啟動時由 `_Locale.SetCulture` 使用。
  - `LogSql`: 是否開啟 SQL 記錄（`true`/`false`）。
  - `LogDebug`: 是否開啟偵錯級別日誌（`true`/`false`）。
  - `ServerId`: 服務識別字串（例：`01`）。
  - `OutPackStationName`, `IsOutPackStationStore`, `DefaultStoreNO`, `DefaultFactoryName`, `DefaultLineName`: 工廠/倉儲與站位預設值。
  - `APS_Simulation_ErrorData_Clear_Day`: 自動清理天數（數字）。
  - `Default_Simulation_AGE01/02/03`, `Default_SimulationDelay`: 排程模擬相關開關與延遲（秒）。
  - `Default_WorkingPaper_*`: 工作底稿相關預設（布林、時間、紙張大小、倉位）。
  - `Default_EStore_ControlURL`, `Default_EStore_MachineToken`: 電子儲物櫃 API 設定（可多台以分號分隔）。
  - `AdministratorEmail`, `SystemEmail`, `MailSmtpServer`, `MailSmtpPort`, `MailCredentialsAccount`, `MailCredentialsPWD`: 郵件設定。**注意**：`MailCredentialsPWD` 已註示移至 user-secrets（使用 `dotnet user-secrets set "FunConfig:MailCredentialsPWD" "your-password"`）。
  - `MailSubjectIsSame`, `Smtp`: 郵件格式或額外 SMTP 配置。
  - `SendMonitorMail00` ~ `SendMonitorMail23`: 多個監控通知收件者清單（範例為多個 email 以逗號分隔）。
  - `MasterServiceIP`, `ElectronicTagsURL`, `LocalWebURL`, `WesocketURL`, `WesocketPort`: 網路與外部服務 URL/Port（WebSocket/電子標籤/Local URL）。
  - `RUNMode`, `DefaultCalendarName`, `AdminKey03`, `AdminKey14`: 系統模式與效能/計算相關開關。
  - `IsAutoDispatch`, `IsAutoDispatch_IsAutoUpdate_WO`, `APS_CT_Custom_Rate`: 派工/APS 相關控制參數。
  - `RunTimeServerLoopTime`: `RUNTimeServer` 的主循環間隔（毫秒），範例為 `60000`。

 以上為現有 `appsettings.json` 中可見的欄位；專案中可能還有程式動態判斷或延伸的 `FunConfig` 欄位（請以 `Configuration.GetSection("FunConfig")` 綁定後的 `ConfigDto` 定義為最終來源）。
- 變更注意事項：
  - 修改 `appsettings.json` 後重新啟動應用程式以套用變更；某些欄位會於啟動時被立即綁定到 `_Fun.Config`。
  - 若要在開發環境覆寫設定，使用 `appsettings.Development.json` 或環境變數（`ASPNETCORE_ENVIRONMENT`）。
  - 安全性：不要把生產 DB 密碼放入 repo；在生產環境使用機密管理或環境變數。
- 範例（示意）：

```json
"FunConfig": {
  "Db": "Server=.;Database=SoftNet;User Id=sa;Password=...;",
  "WesocketPort": 8081,
  "Locale": "zh-TW",
  "MasterIP": "192.168.1.100",
  "MasterPort": 5431
}
```

---

如果你要我把 `RUNTimeServer` 中的某幾個方法（例如 `MasterProcessRequest`）逐行解讀並列出測試用例，我可以繼續展開；或我也可以把 `_Fun.Config` 的所有實際欄位從 `appsettings.json` 抽取並列出。請告訴我下一步要聚焦哪個項目。
