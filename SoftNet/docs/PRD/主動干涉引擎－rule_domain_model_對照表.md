# 主動干涉引擎｜Rule × Domain Model（資料表責任對照）

> 本文件目的：
> 將目前 **RUNTimeServer.cs + Background Thread + #region 規則**
> 正式轉譯為「可維護、可重構、可產品化」的 **Rule × Domain Model 責任對照表**。
>
> 本表是後續：
> - Rule Engine 化
> - Cursor 重構
> - 產品版本切分（Standard / Pro）
> 的核心依據。

---

## 一、Domain 分層總覽（結論先行）

| Domain | 說明 | 是否被主動干涉修改 |
|---|---|---|
| APS / 排程 | 計畫與模擬 | ✅ 高 |
| 製造 / 工單 | 現場執行 | ✅ 極高 |
| 倉儲 / 物料 | 庫存與流轉 | ✅ 高 |
| 警示 / 治理 | 通知與修正 | ✅ 中 |

---

## 二、Rule × Domain × 資料表 對照

> ⚠️ Read = 讀取判斷條件
> ✏️ Write = 系統自動修正

---

### R-01 背景服務與系統連線確認

| 項目 | 內容 |
|---|---|
| Domain | 系統治理 |
| Read | SystemConfig, DBStatus |
| Write | — |
| 風險 | 低 |
| 備註 | 防止誤觸後續 Rule |

---

### R-02 APS Simulation 完成未轉正式生產

| 項目 | 內容 |
|---|---|
| Domain | APS / 排程 |
| Read | APS_Simulation |
| Write | APS_Simulation (Delete / Update) |
| 風險 | 中 |
| 備註 | 防止模擬資料污染正式排程 |

---

### R-03 APS Simulation State 自動修正

| 項目 | 內容 |
|---|---|
| Domain | APS / 排程 |
| Read | APS_Simulation.State |
| Write | APS_Simulation.State |
| 風險 | 中 |
| 備註 | State 7 → 9 修正 |

---

### R-04 排程採購 / 外加工 / 安全存量自動開立

| 項目 | 內容 |
|---|---|
| Domain | APS / 倉儲 |
| Read | Stock, Math_TotalStock, Math_Online |
| Write | PurchaseOrder / OutsourceOrder |
| 風險 | 高 |
| 備註 | 關係到金流與供應鏈 |

---

### R-05 工站生產用料單自動開立

| 項目 | 內容 |
|---|---|
| Domain | 製造 / 倉儲 |
| Read | WorkOrder, WorkStation |
| Write | MaterialIssue |
| 風險 | 高 |
| 備註 | 上下站關聯強 |

---

### R-06 工單暫停 / 未完成修正

| 項目 | 內容 |
|---|---|
| Domain | 製造 |
| Read | WorkOrder, WorkOrderDetail |
| Write | WorkOrder.Status |
| 風險 | 高 |
| 備註 | 避免卡單 |

---

### R-07 已派工未領料監控

| 項目 | 內容 |
|---|---|
| Domain | 製造 / 倉儲 |
| Read | WorkOrder, MaterialIssue |
| Write | APS_WarningData |
| 風險 | 中 |
| 備註 | 可設定是否 Auto Fix |

---

### R-08 Wait 單據自動轉 OK

| 項目 | 內容 |
|---|---|
| Domain | 單據 |
| Read | Document.Status |
| Write | Document.IsOK |
| 風險 | 高 |
| 備註 | 需納入權限與稽核 |

---

### R-09 生產達標自動結案

| 項目 | 內容 |
|---|---|
| Domain | 製造 |
| Read | ProductionLog, WorkOrder |
| Write | WorkOrder, WorkStationLoad |
| 風險 | 極高 |
| 備註 | 影響後續排程與結算 |

---

### R-10 CT / UPH 偏離有效值

| 項目 | 內容 |
|---|---|
| Domain | 效率分析 |
| Read | ProductionLog, EffectiveCT |
| Write | APS_WarningData |
| 風險 | 低 |
| 備註 | 不自動修正，只警示 |

---

### R-11 人工報工延遲

| 項目 | 內容 |
|---|---|
| Domain | 製造 |
| Read | ProductionLog |
| Write | MailLog, APS_WarningData |
| 風險 | 中 |
| 備註 | 影響績效評估 |

---

### R-12 工單應開未開

| 項目 | 內容 |
|---|---|
| Domain | 製造 |
| Read | WorkOrder |
| Write | APS_WarningData |
| 風險 | 高 |
| 備註 | 生產延遲的主要來源 |

---

### R-13 料件安全量 / 移轉合理性

| 項目 | 內容 |
|---|---|
| Domain | 倉儲 |
| Read | Stock, TransferLog |
| Write | — |
| 風險 | 中 |
| 備註 | 阻止錯誤移轉 |

---

### R-14 停工未停（跨班別）

| 項目 | 內容 |
|---|---|
| Domain | 製造 |
| Read | WorkStation, Shift |
| Write | APS_WarningData |
| 風險 | 高 |
| 備註 | 需跨班別判斷 |

---

### R-15 排程完成日偏離

| 項目 | 內容 |
|---|---|
| Domain | APS / 排程 |
| Read | APS_WorkingPaper |
| Write | APS_WarningData |
| 風險 | 中 |
| 備註 | 供主管決策 |

---

## 三、關鍵結論（請務必保留）

1. **主動干涉 = 高風險寫入系統**，必須標示 Rule 等級
2. 不同 Rule 應允許：
   - Auto Fix
   - Warning Only
   - Disabled
3. Domain 清楚後，才能：
   - 拆 Service
   - 做設定化
   - 分產品版本

---

> 這份文件，代表：
> **Softnet MES 的「靈魂」，第一次不是寫在程式裡，而是寫在產品文件裡。**

