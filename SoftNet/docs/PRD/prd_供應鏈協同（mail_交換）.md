# 供應鏈協同模組 PRD（Mail 交換版，可實作規格）

> **文件版本**: v1.1（可實作）  
> **適用對象**: 產品規劃、工程開發、導入顧問、AI 輔助開發  
> **最後更新**: 2026-04-30

---

## 📋 目錄

1. [目的與範圍](#1-目的與範圍)
2. [業務情境與限制](#2-業務情境與限制)
3. [系統邊界與責任](#3-系統邊界與責任)
4. [Mail 交換總覽](#4-mail-交換總覽)
5. [主旨規格與路由規則](#5-主旨規格與路由規則)
6. [訊息類型定義](#6-訊息類型定義)
7. [核心流程（可直接開發）](#7-核心流程可直接開發)
8. [Need 關聯模型](#8-need-關聯模型)
9. [資料表規格（建議）](#9-資料表規格建議)
10. [狀態機與重送規範](#10-狀態機與重送規範)
11. [異常交換與排程前置處理](#11-異常交換與排程前置處理)
12. [安全與治理](#12-安全與治理)
13. [與既有模組整合點](#13-與既有模組整合點)
14. [MVP 範圍與驗收案例](#14-mvp-範圍與驗收案例)
15. [附錄：JSON 範本](#15-附錄json-範本)

---

## 1. 目的與範圍

本文件定義 SoftNet MES 在供應鏈協同場景下，透過**同一個 Mail 通道**交換上下游的需求、回覆交期、生產進度與異常事件的規格。

### 1.1 目標

- A 端（需求端）於建立 `DOC1Buy/DOC1BuyII` 或 `DOC4Production/DOC4ProductionII` 後，可通知 B 端 (供貨端) 建立排程需求。
- 建立 `NeedId_A + SimulationId_A ↔ NeedId_B` 的可追溯關聯。
- 支援交期回覆、投產狀態、異常交換，供 A 端判斷是否影響排程。
- 不修改既有核心商業邏輯的前提下，以增量模組方式導入。

### 1.2 非範圍

- 本文件不規範 UI 美術與版面細節。
- 本文件不要求即時秒級通訊（允許分鐘級延遲）。

---

## 2. 業務情境與限制

### 2.1 限制條件

- 同時讓多套的 SoftNet MES 可以交換生產資訊
- 各客戶站點 SoftNet MES 皆有獨立 `ServerId` (唯一用戶編號)
- 利用共用 Mail 通道作為訊息交換介面 (使用 FunConfig.SCM_DataExchange_Mail 參數)。

### 2.2 基本規則

- 以主旨解析 `FromServerId` 與 `ToServerId`。
- 非本機 `ToServerId` 讀取主旨後直接略過並 `continue`，不進入業務處理，亦不落地 `SCM_MailInbox`。
- 同一 `MessageId` 必須冪等處理（只入帳一次）。
- 寄件後的 `MessageId` 與收件的 `MessageId` 為同一個, 方便日後追查對應。
- `MessageId` 碼需含有寄件者本身的 `ServerId`

---

## 3. 系統邊界與責任

### 3.1 Mail Bridge 層責任

- 寄件排隊與重送
- 收件擷取與解析
- 路由判斷
- 冪等檢查
- 呼叫業務處理器

### 3.2 業務處理層責任

- 依 `MessageType` 執行對應業務（建立 Need、回覆交期、記錄進度、產生警示）
- 回寫關聯表與處理狀態
- 產生回執訊息

---

## 4. Mail 交換總覽

### 4.1 邏輯元件

- `Outbound Queue`：待寄訊息暫存。
- `Mail Sender`：批次寄件。
- `Mail Receiver`：定時讀信。
- `Inbound Processor`：依類型處理。
- `Ack Processor`：處理回執與狀態遷移。

### 4.2 執行節奏

- 背景服務固定輪巡，使用 `FunConfig.SCM_WorkingDoWhileTime`（分鐘）控制週期。
- 寄件與收件為兩段循環，錯誤互不阻塞。

---

## 5. 主旨規格與路由規則

> 主旨必須可被程式穩定解析；不可使用自由格式。

### 5.1 主旨格式（正式）

`[SoftNetSCM]|From:{FromServerId}|To:{ToServerId}|Type:{MessageType}|Msg:{MessageId}|V:{Version}`

### 5.2 主旨範例

`[SoftNetSCM]|From:01|To:07|Type:CreateNeed|Msg:6f2d2b5c-9a93-4f3a-96f3-cd0ddf6c6f29|V:1`

### 5.3 解析規則

1. 先檢查前綴是否為 `[SoftNetSCM]`。  
2. 解析 `From/To/Type/Msg/V`。  
3. 若 `To != 本機 FunConfig.ServerId`，直接略過並進入下一封。  
4. 若 `Msg` 已存在且已成功處理，標記 `Duplicated`。  
5. 通過後才讀 Body JSON。

---

## 6. 訊息類型定義

### 6.1 `MessageType`（固定枚舉）

- `CreateNeed`：A 發需求給 B（採購/委外）
- `PromiseDate`：B 回覆可承諾交期（可含可承諾量）
- `OKNeed`：A 確認採納承諾，要求 B 建立正式需求
- `NGNeed`：A 不接受承諾交期，回覆可接受日期要求 B 重新評估
- `CancelNeed`：A 取消需求，要求 B 刪除/中止該需求
- `NeedCreated`：B 回覆已建立 `NeedId` A的記錄檔中將此 `NeedId` 列為有效
- `AskStatus`：A 詢問 B 目前生產狀態
- `ProdStatus`：B 回覆排程/投產/報工摘要
- `Exception`：B 回覆異常事件
- `CloseNeed`：B 回覆關聯需求結案
- `AckMsg`：通用技術回執（收訖/解析結果）

### 6.2 必要欄位（Body）

- `messageId`
- `messageType`
- `fromServerId`
- `toServerId`
- `eventTime`
- `docType`（`DOC1`/`DOC4`）
- `docNumberNO`
- `needMapping.needIdA`
- `needMapping.simulationIdA`
- `needMapping.needIdB`（依階段可空）
- `payload`

---

## 7. 核心流程（可直接開發）

### 7.1 流程 A：建立需求關聯（採購/委外）

1. A 建立 `DOC1` 或 `DOC4` (若單據資料表 MFNO 欄位的 MFData表的SCM_ServerId有值，才會觸發SCM整個流程，否則流程不會往下執行)。  
2. A 寫入 `SCM_MailOutbox(CreateNeed)` 並寄送 `CreateNeed`。  
3. B 收到 `CreateNeed`，先建立模擬需求，取得 B 的 NeedId（SFC_Common 的 RunSetSimulation 函式 模擬），回覆 `PromiseDate`。  
4. A 收到 `PromiseDate`，發Mail給 SCM_LocalDecision_Mail 參數確認接收。  
5. A Mail確認同意後送出 `OKNeed`，若不同意送出 `NGNeed` 另加可接受的日期，若取消送出 `CancelNeed`， 將 NeedId 傳給 B。  
6. B 收到 `OKNeed`，用傳來的 NeedId 轉正式排程 SFC_Common 的 RunSetSimulation 函式 轉正式生產）並回覆 `NeedCreated`。  
7. B 收到 `CancelNeed`，用傳來的 NeedId 執行 SimulationService 的 OtherDeleteAsync 刪除 NeedId。 
8. B 收到 `NGNeed`，用傳來的 NeedId 與可接受日期重新執行第 3 步驟模擬試算（可調整急件參數），並回覆新的 `PromiseDate`。
9. A 收到 `NeedCreated`，寫入 `SCM_NeedLink(NeedId_A/SimulationId_A/NeedId_B)`，紀錄 B 的 NeedId 正式成立。

### 7.2 流程 B：狀態追蹤

1. A 定時送 `AskStatus`（問交貨進度）。  
2. B 回覆 `ProdStatus` 或 `Exception`（進度、報工、預估完成日、異常）。  
3. A 更新 `SCM_StatusSnapshot` 並顯示。


### 7.3 流程 C：異常處理

1. B 偵測到異常，送出 `Exception`。  
2. A 收到後寫入警示/中繼表。  
3. A 的排程或干涉流程評估是否調整計畫。  
4. 若有調整，A 可回送確認訊息給 B（可用 `AckMsg`）。

---

## 8. Need 關聯模型

### 8.1 主關聯欄位

- `FromServerId`
- `ToServerId`
- `NeedId_A`
- `SimulationId_A`
- `NeedId_B`
- `DocType`
- `DocNumberNO`
- `LinkStatus`

### 8.2 關聯規則

- `NeedId_A` 可對多個 `SimulationId_A`。
- `SimulationId_A` 可對一個或多個 `NeedId_B`。
- `NeedCreated` 訊息為建立 `NeedId_B` 的唯一確認來源。

---

## 9. 資料表規格（建議）

> 可用新表；若需併入既有表，需保留同等資訊欄位。

### 9.1 `SCM_MailOutbox`

**中文說明（寄件佇列表）**

- 用途：記錄所有「待寄送 / 重送中 / 已寄送 / 取消」的 SCM 郵件訊息。
- 角色：Mail Sender 只讀取此表中 `Status=Pending` 或符合重送條件的資料。
- 重點：以 `MessageId` 做唯一鍵，避免同一訊息重複建立。

| 欄位 | 型別 | 中文說明 |
|---|---|---|
| `Id` | PK, bigint identity | 寄件佇列流水號 |
| `MessageId` | nvarchar(50), unique | 訊息唯一識別碼（冪等鍵） |
| `FromServerId` | nchar(2) | 發送端 ServerId |
| `ToServerId` | nchar(2) | 目標接收端 ServerId |
| `MessageType` | nvarchar(20) | 訊息類型（CreateNeed/PromiseDate...） |
| `Subject` | nvarchar(500) | 郵件主旨（含路由資訊） |
| `BodyJson` | nvarchar(max) | 郵件內容 JSON |
| `Status` | char(1) | 寄送狀態（`0=Pending,1=Sent,2=Failed,3=Cancelled`） |
| `RetryCount` | int | 已重送次數 |
| `NextRetryTime` | datetime | 下次可重送時間 |
| `CreatedTime` | datetime | 佇列建立時間 |
| `SentTime` | datetime null | 實際寄送成功時間 |
| `ErrorMessage` | nvarchar(1000) null | 最後一次失敗錯誤訊息 |

### 9.2 `SCM_MailInbox`

**中文說明（收件暫存表）**

- 用途：記錄從 Mail 收進來的原始訊息與解析/處理結果。
- 角色：Mail Receiver 寫入，Inbound Processor 讀取後更新處理狀態。
- 重點：即使格式錯誤，也要落地記錄以便追查。

| 欄位 | 型別 | 中文說明 |
|---|---|---|
| `Id` | PK, bigint identity | 收件暫存流水號 |
| `MessageId` | nvarchar(50) | 訊息唯一識別碼（對應寄件端） |
| `FromServerId` | nchar(2) | 來源端 ServerId |
| `ToServerId` | nchar(2) | 目標端 ServerId |
| `MessageType` | nvarchar(20) | 訊息類型 |
| `Subject` | nvarchar(500) | 原始郵件主旨 |
| `BodyJson` | nvarchar(max) | 原始郵件內容 JSON |
| `ReceiveTime` | datetime | 收件時間 |
| `ParseStatus` | char(1) | 主旨/內容解析狀態（`0=OK,1=InvalidSubject,2=InvalidJson`；`3=NotTarget` 保留碼，現行流程不落地） |
| `ProcessStatus` | char(1) | 業務處理狀態（`0=Pending,1=Done,2=Failed,3=Duplicated`） |
| `ErrorMessage` | nvarchar(1000) null | 解析或處理錯誤訊息 |

### 9.3 `SCM_NeedLink`

**中文說明（需求關聯主表）**

- 用途：保存 A/B 雙方需求關聯，核心為 `NeedId_A + SimulationId_A ↔ NeedId_B`。
- 角色：收到 `NeedCreated` 後建立或更新此表，供後續查詢與追蹤使用。
- 重點：此表是供應鏈可追溯的關鍵，狀態需完整（Pending/Linked/Closed/Rejected）。

| 欄位 | 型別 | 中文說明 |
|---|---|---|
| `LinkId` | PK, nchar(12) or bigint | 關聯主鍵 |
| `FromServerId` | nchar(2) | 下單端（A 端）ServerId |
| `ToServerId` | nchar(2) | 接單端（B 端）ServerId |
| `NeedId_A` | nchar(12) | A 端需求編號 |
| `SimulationId_A` | nchar(12) null | A 端子排程編號（可空） |
| `NeedId_B` | nchar(12) | B 端需求編號 |
| `DocType` | nchar(4) | A 端單據類型（DOC1/DOC4） |
| `DocNumberNO` | nvarchar(20) | A 端來源單據編號 |
| `LinkStatus` | char(1) | 關聯狀態（`0=Pending,1=Linked,2=Closed,3=Rejected`） |
| `CreatedTime` | datetime | 建立時間 |
| `UpdatedTime` | datetime | 最後更新時間 |

### 9.4 `SCM_StatusSnapshot`

**中文說明（供應鏈進度快照表）**

- 用途：保存對方回覆的交期、排程、生產進度與異常摘要。
- 角色：收到 `ProdStatus` / `Exception` 後寫入，供畫面與排程評估使用。
- 重點：採快照模式（可保留歷史），避免直接覆蓋造成追溯困難。

| 欄位 | 型別 | 中文說明 |
|---|---|---|
| `SnapshotId` | PK, bigint identity | 快照流水號 |
| `LinkId` | FK | 對應 `SCM_NeedLink` 關聯鍵 |
| `NeedId_B` | nchar(12) | B 端需求編號 |
| `PromiseDate` | datetime null | B 端承諾交期 |
| `PlanState` | nvarchar(50) | B 端排程/投產狀態 |
| `ReportProgress` | nvarchar(200) | 生產進度摘要（可含百分比/報工量） |
| `ExceptionCode` | nvarchar(30) null | 異常代碼 |
| `ExceptionRemark` | nvarchar(500) null | 異常說明 |
| `SnapshotTime` | datetime | 快照產生時間 |

---

## 10. 狀態機與重送規範

### 10.1 Outbox 狀態機

- `Pending -> Sent`：SMTP 成功
- `Pending -> Failed`：SMTP 失敗
- `Failed -> Pending`：到達 `NextRetryTime`
- `Any -> Cancelled`：人工取消

### 10.2 重送規範（建議）

- 第 1~3 次：使用 `FunConfig.SCM_WorkingDoWhileTime` 重送
- 超過 3 次：停送並告警

### 10.3 冪等鍵

- 技術冪等：`MessageId`
- 業務冪等：`FromServerId + ToServerId + MessageType + NeedId_A + SimulationId_A + NeedId_B`

---

## 11. 異常交換與排程前置處理

### 11.1 `ExceptionCode` 建議

- `NeedDelayed`
- `MaterialShortage`
- `CapacityIssue`
- `ProductionStopped`
- `NeedPaused`

### 11.2 前置處理規則

- 先寫入中繼/警示表，再由既有排程邏輯評估。
- 不直接跨模組即時改排程，用發Mail 使用 `FunConfig.SCM_LocalDecision_Mail` 讓決策者用Mail按鈕決定如何干涉。
- 干涉策略：依系統既有干涉邏輯處置。

---

## 12. 安全與治理

### 12.1 身分驗證

- 寄件地址需與 `ServerId` 白名單對應。
- `FromServerId`、`ToServerId` 必須是合作夥伴資料中的有效值。

### 12.2 訊息完整性

- 建議在 Body 加入 `signature`（HMAC-SHA256）。
- 驗簽失敗只記錄不處理。

### 12.3 稽核

- 收發、解析、處理、重送、失敗都需可追蹤。

---

## 13. 與既有模組整合點

### 13.1 單據觸發

- `DOC1Buy/DOC1BuyII` 產生後觸發 `CreateNeed`
- `DOC4Production/DOC4ProductionII` 產生後觸發 `CreateNeed`

### 13.2 排程觸發

- 收到 `NeedCreated` 後建立關聯，進入 A 端排程判斷。
- 收到 `PromiseDate/Exception/ProdStatus` 後更新快照與警示，供干涉流程參考。

### 13.3 夥伴主檔對應

- 供應商接收端可由 `MFData` 擴充欄位對應 `SCM_ServerId`。
- 客戶端可由 `CTData` 擴充欄位對應 `SCM_ServerId`。

---

## 14. MVP 範圍與驗收案例

### 14.1 MVP 功能

- `CreateNeed`、`PromiseDate`、`OKNeed`、`NeedCreated`
- `SCM_NeedLink` 建立
- `AskStatus/ProdStatus` 基礎查詢
- Outbox/Inbox 冪等與重送

### 14.2 驗收案例

1. A 發 `DOC1`，B 能收到並回 `PromiseDate`。  
2. A 發 `OKNeed`，B 回 `NeedCreated`，A 成功寫入關聯。  
3. 非本機 `ToServerId` 郵件不進業務處理。  
4. 重覆同 `MessageId` 郵件不會重複建單。  
5. B 回報 `ProdStatus` 後，A 能顯示最新快照。  
6. B 回報 `Exception` 後，A 能落地警示供排程判斷。

---

## 15. 附錄：JSON 範本

### 15.1 `CreateNeed`

```json
{
  "messageId": "6f2d2b5c-9a93-4f3a-96f3-cd0ddf6c6f29",
  "messageType": "CreateNeed",
  "fromServerId": "01",
  "toServerId": "07",
  "eventTime": "2026-04-25T17:30:00+08:00",
  "docType": "DOC4",
  "docNumberNO": "PA202604250001",
  "needMapping": {
    "needIdA": "X011A4T1A13C",
    "simulationIdA": "Y011A4T1AHF8",
    "needIdB": ""
  },
  "payload": {
    "partNO": "P-1001",
    "qty": 1200,
    "arrivalDate": "2026-04-30T00:00:00+08:00",
    "remark": "委外加工需求"
  },
  "version": 1,
  "signature": ""
}
```

### 15.2 `NeedCreated`

```json
{
  "messageId": "6f2d2b5c-9a93-4f3a-96f3-cd0ddf6c6f30",
  "messageType": "NeedCreated",
  "fromServerId": "07",
  "toServerId": "01",
  "eventTime": "2026-04-25T18:00:00+08:00",
  "docType": "DOC4",
  "docNumberNO": "PA202604250001",
  "needMapping": {
    "needIdA": "X011A4T1A13C",
    "simulationIdA": "Y011A4T1AHF8",
    "needIdB": "Y071A4TER2EE"
  },
  "payload": {
    "state": "Created",
    "promiseDate": "2026-05-02T00:00:00+08:00"
  },
  "version": 1,
  "signature": ""
}
```

---

**備註**：本版本已調整為可直接作為實作規格使用；若後續需要可再補「錯誤碼表」與「API/Service 命名映射」附錄。

