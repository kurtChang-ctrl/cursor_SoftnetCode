# 供應鏈協同（Mail 交換）實作清單

> **對應規格文件**: `docs/PRD/prd_供應鏈協同（mail_交換）.md`  
> **文件版本**: v1.1  
> **最後更新**: 2026-04-30  
> **目的**: 供開發排程、分工與驗收追蹤使用（本文件為實作拆解，不是最終需求來源）

---

## 1. 實作目標

- 依既有規格完成 Mail Bridge（收件、寄件、路由、冪等、重送）。
- 打通 `DOC1` / `DOC4` 觸發 `CreateNeed` 的供應鏈交換流程。
- 完成 `NeedId_A + SimulationId_A ↔ NeedId_B` 關聯落地與查詢。
- 支援 `AskStatus` / `ProdStatus` / `Exception` 基礎交換。

---

## 2. 開發分階段（建議）

### Phase 1（MVP）

- 建立資料表：`SCM_MailOutbox`、`SCM_MailInbox`、`SCM_NeedLink`
- 完成 `CreateNeed -> PromiseDate -> OKNeed/NGNeed/CancelNeed -> NeedCreated`
- 完成主旨路由、冪等、重送、錯誤記錄
- 完成關聯建立與查詢頁最小顯示

### Phase 2

- 建立資料表：`SCM_StatusSnapshot`
- 完成 `AskStatus` / `ProdStatus`（含排程與報工摘要）
- 完成 `Exception` 事件入站與警示落地

### Phase 3

- 完成 `CloseNeed`
- 完成自動前置處理（WarnOnly / Suggest / AutoAdjust）
- 完成稽核報表與運維監控

---

## 3. 資料庫實作清單（Migration）

## 3.1 新增資料表

- `SCM_MailOutbox`
- `SCM_MailInbox`
- `SCM_NeedLink`
- `SCM_StatusSnapshot`

### 3.2 必要索引

- `SCM_MailOutbox`
  - `UX_SCM_MailOutbox_MessageId`（unique）
  - `IX_SCM_MailOutbox_Status_NextRetryTime`
- `SCM_MailInbox`
  - `IX_SCM_MailInbox_MessageId`
  - `IX_SCM_MailInbox_ToServerId_ReceiveTime`
- `SCM_NeedLink`
  - `IX_SCM_NeedLink_NeedIdA_SimulationIdA`
  - `IX_SCM_NeedLink_NeedIdB`
- `SCM_StatusSnapshot`
  - `IX_SCM_StatusSnapshot_LinkId_SnapshotTime`

### 3.3 主檔擴充欄位（可選但建議）

- `MFData` 新增 `SCM_ServerId`（供應商目標識別）
- `CTData` 新增 `SCM_ServerId`（客戶來源識別）

### 3.4 佈署腳本

- 提供 `up` / `down` 腳本（可回滾）
- 提供初始化狀態碼資料（如需代碼表）

---

## 4. 設定檔與參數清單

> 放在 `FunConfig`，由 `appsettings.json` 綁定

- `SCM_IsEnabled`（bool）
- `SCM_WorkingDoWhileTime`（int，分鐘）
- `SCM_MaxRetryCount`（int，建議預設 3）
- `SCM_MailSubjectPrefix`（string，預設 `[SoftNetSCM]`）
- `SCM_EnableSignature`（bool）
- `SCM_SignatureKey`（string，機敏值）

---

## 5. 程式模組實作清單

## 5.1 Service 模組（建議命名）

- `ScmMailBridgeService`
  - `BuildSubject()`
  - `ParseSubject()`
  - `ValidateRoute()`
  - `ValidateSignature()`

- `ScmOutboxService`
  - `EnqueueMessage()`
  - `SendPendingMessages()`
  - `HandleSendFailAndRetry()`

- `ScmInboxService`
  - `FetchMailboxMessages()`
  - `StoreInbox()`
  - `ProcessInboxMessages()`

- `ScmNeedLinkService`
  - `CreateOrUpdateNeedLink()`
  - `GetNeedLinkByA()`
  - `GetNeedLinkByB()`

- `ScmStatusService`
  - `BuildProdStatusSnapshot()`
  - `SaveSnapshot()`

## 5.2 背景工作（HostedService）

- `ScmMailWorker`
  - 循環 1：Outbox 寄件
  - 循環 2：Inbox 收件
  - 循環 3：失敗重試

## 5.3 MessageType Handler（建議）

- `HandleCreateNeed()`
- `HandlePromiseDate()`
- `HandleOKNeed()`
- `HandleNGNeed()`
- `HandleCancelNeed()`
- `HandleNeedCreated()`
- `HandleAskStatus()`
- `HandleProdStatus()`
- `HandleException()`
- `HandleCloseNeed()`
- `HandleAckMsg()`

---

## 6. 與既有流程整合點（開發掛點）

### 6.1 單據產生後送出訊息

- 掛點 A：`Create_DOC1stock` 成功後 enqueue `CreateNeed`
- 掛點 B：`Create_DOC4stock` 成功後 enqueue `CreateNeed`

### 6.2 關聯建立

- 收到 `NeedCreated` 後，寫入 `SCM_NeedLink`
- 同步更新單據/查詢頁面的對應欄位（只讀顯示）

### 6.3 進度與異常

- 收到 `ProdStatus`，寫入 `SCM_StatusSnapshot`
- 收到 `Exception`，寫入中繼與警示（供排程判斷）

---

## 7. 主旨解析與路由檢核清單

- 驗證前綴：`[SoftNetSCM]`
- 驗證必要段：`From/To/Type/Msg/V`
- 檢查 `ToServerId == FunConfig.ServerId`
- 檢查 `MessageId` 是否已處理
- 驗證 `FromServerId` 是否為合法節點
- `ToServerId` 非本機時：直接 `continue`，不寫入 `SCM_MailInbox`

---

## 8. 冪等與重送規則（必做）

- 同一 `MessageId` 僅允許成功處理一次
- 重複訊息回 `AckMsg`，不重複建立資料
- 寄送失敗依規則重送：
  - 第 1~3 次：使用 `FunConfig.SCM_WorkingDoWhileTime` 週期重送
  - 超過 3 次：停止重送並告警

---

## 9. 測試清單

## 9.1 單元測試

- 主旨生成/解析
- 冪等判斷
- 重送時間計算
- MessageType 路由分發

## 9.2 整合測試

- A 發 `CreateNeed` -> B 收到
- B 回 `PromiseDate` -> A 收到
- A 發 `OKNeed` -> B 建立 `NeedId_B`
- A 發 `NGNeed` -> B 重新試算並回新 `PromiseDate`
- A 發 `CancelNeed` -> B 刪除/中止對應 Need
- B 回 `NeedCreated` -> A 寫入 `SCM_NeedLink`
- A `AskStatus` -> B `ProdStatus`

## 9.3 異常測試

- ToServerId 非本機（應略過業務處理，且不落地 Inbox）
- 主旨格式錯誤（應記錄 ParseFail）
- JSON 內容缺欄（應記錄 ProcessFail）
- SMTP/IMAP 暫時失敗（應重送）
- 重覆 MessageId（應 Duplicated）

## 9.4 壓力與穩定性

- 連續 1000 封訊息處理
- 大量重試情境下的 CPU/DB 壓力
- 長時間輪巡（24h）穩定性

---

## 10. 上線前檢核（Go-Live Checklist）

- Migration 已在測試與預備環境驗證
- `FunConfig` 參數與機敏值配置完成
- Mail 帳號權限（SMTP/IMAP）已確認
- 來源白名單與 ServerId 對應已建立
- 監控與告警（Failed/Retry 超限）已開啟
- 回滾方案已演練

---

## 11. 驗收標準（Definition of Done）

- 功能：MVP 流程可完整跑通且可追溯
- 正確性：重複訊息不重複入帳
- 可靠性：失敗可重送，最終可收斂
- 可觀測：每筆收發有完整 log 與狀態
- 安全性：非目標/非法來源不進業務流程

---

## 12. 任務分工建議（可直接建 Issue）

- DB Migration（1 人）
- Mail Bridge Service（1 人）
- Message Handlers（1~2 人）
- 與單據/排程整合（1 人）
- 測試與驗收腳本（1 人）

---

**備註**：本文件為「實作分解清單」，需求準據以 `prd_供應鏈協同（mail_交換）.md` 為主；若兩者衝突，應先更新需求文件再回寫本清單。
