# 單據管理模組 PRD

> **文件版本**: v1.0  
> **適用對象**: 產品規劃、工程開發、AI 輔助開發  
> **最後更新**: 2025-01-12

---

## 📋 目錄

1. [單據管理概述](#1-單據管理概述)
2. [單據類型](#2-單據類型)
3. [自動產生邏輯](#3-自動產生邏輯)
4. [單據確認流程](#4-單據確認流程)
5. [單據簽核機制](#5-單據簽核機制)
6. [相關資料表](#6-相關資料表)

---

## 1. 單據管理概述

### 1.1 核心概念

單據管理模組負責自動產生、追蹤、確認與簽核各種生產相關單據，確保物料流動與生產流程的完整性。

### 1.2 主要功能

- ✅ **自動產生單據**: 根據排程需求、安全存量、報工情況自動產生單據
- ✅ **單據確認流程**: 單據產生後需經過確認才能生效
- ✅ **單據狀態追蹤**: 追蹤單據從產生到完成的完整生命週期
- ✅ **單據簽核機制**: 重要單據需經過簽核流程

---

## 2. 單據類型

| 單據類型 | 單據代碼 | 說明 | 對應資料表 |
|---------|---------|------|-----------|
| **進貨單** | `AA02` | 採購物料進貨 | `DOC1Buy`, `DOC1BuyII` |
| **領料單** | `AC01` | 生產用領料 | `DOC3stock`, `DOC3stockII` |
| **入庫單** | `BC01` | 生產完成入庫 | `DOC3stock`, `DOC3stockII` |
| **委外加工單** | `PA02` | 委外加工需求 | `DOC4Production`, `DOC4ProductionII` |
| **退料單** | `EB01` | 餘料退回入庫 | `DOC3stock`, `DOC3stockII` |

---

## 3. 自動產生邏輯

### 3.1 排程採購自動開立（RUNTimeServer.cs 1120-1175行）

#### 觸發條件

- `APS_Simulation.IsWPaper = '0'`（尚未建立工作底稿）
- `APS_Simulation.StartDate <= 當前日期`
- `APS_Simulation.DOCNumberNO` 為空或 null
- **庫存不足**: `(NeedQTY + SafeQTY) > (Math_TotalStock_HasUseQTY + Math_Online_SurplusQTY)`
- `Material.IS_WorkingPaper = '1'`（物料需建立工作底稿）
- `APS_NeedData.State = '6'`（排程已確認）

#### 產生邏輯

1. **計算缺料數量**:
   ```csharp
   tmp_i = (NeedQTY + SafeQTY) - (Math_TotalStock_HasUseQTY + Math_Online_SurplusQTY)
   ```

2. **查找適合廠商**:
   - 呼叫 `_SFC_Common.SelectDOC1BuyMFNO()` 選擇適合的廠商（MFNO）
   - 取得價格（Price）

3. **查找適合入庫儲別**:
   - 呼叫 `_SFC_Common.SelectINStore()` 選擇入庫儲位（IN_StoreNO, IN_StoreSpacesNO）

4. **建立進貨單**:
   - 呼叫 `_SFC_Common.Create_DOC1stock()` 建立進貨單（DOC1Buy）
   - 單據代碼: `AA02`
   - 備註: "排程物料需求不足"

5. **建立工作底稿**:
   - 插入 `APS_WorkingPaper` 記錄
   - `WorkType = '1'`（採購）
   - 更新 `APS_Simulation.IsWPaper = '1'`

#### 相關資料表

- `APS_Simulation` - 排程模擬資料
- `APS_NeedData` - 排程需求資料
- `Material` - 物料主檔
- `DOC1Buy`, `DOC1BuyII` - 進貨單主檔與明細
- `APS_WorkingPaper` - 工作底稿

---

### 3.2 排程委外加工自動開立（RUNTimeServer.cs 1177-1237行）

#### 觸發條件

- 同「排程採購自動開立」的條件
- **額外條件**: `Source_StationNO = OutPackStationName`（委外加工站）

#### 產生邏輯

1. **計算委外數量**:
   ```csharp
   tmp_i = NeedQTY + SafeQTY
   ```

2. **查找上下階關係**:
   - 查找上階 `APS_Simulation`（`UP_SimulationId`）
   - 查找下階 `APS_Simulation`（`Down_SimulationId`）

3. **查找適合廠商**:
   - 呼叫 `_SFC_Common.SelectDOC4ProductionMFNO()` 選擇委外廠商

4. **查找適合入庫儲別**:
   - 呼叫 `_SFC_Common.SelectINStore()` 選擇入庫儲位

5. **建立委外加工單**:
   - 呼叫 `_SFC_Common.Create_DOC4stock()` 建立委外加工單（DOC4Production）
   - 單據代碼: `PA02`
   - 備註: "監測到排程委外加工需求"

6. **建立工作底稿**:
   - 插入 `APS_WorkingPaper` 記錄
   - `WorkType = '2'`（委外加工）
   - 更新 `APS_Simulation.IsWPaper = '1'` 和 `DOCNumberNO`

#### 相關資料表

- `DOC4Production`, `DOC4ProductionII` - 委外加工單主檔與明細

---

### 3.3 安全存量自動開立（RUNTimeServer.cs 1240-1695行）

#### 3.3.1 非生產件安全存量（1240-1423行）

##### 觸發條件

- `Material.IS_WorkingPaper = '1'`
- `Material.Class != '4'` 且 `Class != '5'`（非生產件）
- **庫存不足**: `SafeQTY > (在倉可用量 + 在途量 + 已在底稿量)`

##### 產生邏輯

1. **計算可用量**:
   ```csharp
   // 在倉可用量 = TotalStock.QTY - (TotalStockII.KeepQTY + TotalStockII.OverQTY)
   // 在途量 = DOC1BuyII.QTY (IsOK='0')
   // 已在底稿量 = APS_WorkingPaper.NeedQTY (WorkType='3')
   tmp_i = 在倉可用量 + 在途量 + 已在底稿量
   ```

2. **判斷是否需要開單**:
   ```csharp
   if (SafeQTY > tmp_i || (Default_WorkingPaper_NOT_SafeQTY_DOC1 && tmp_i < 0))
   ```

3. **查找適合廠商與儲位**:
   - 呼叫 `_SFC_Common.SelectDOC1BuyMFNO()` 選擇廠商
   - 呼叫 `_SFC_Common.SelectINStore()` 選擇儲位

4. **計算到貨日期**:
   - 若有 `PP_EfficientDetail.CT`（週期時間），使用 CT 計算
   - 否則預設 7 天後

5. **建立或更新工作底稿**:
   - 若已存在相同 `PartNO|MFNO|IN_StoreNO|IN_StoreSpacesNO` 的底稿，更新 `NeedQTY`
   - 否則新增 `APS_WorkingPaper` 記錄
   - `WorkType = '3'`（安全存量非生產件）

##### 相關資料表

- `TotalStock`, `TotalStockII` - 庫存主檔與明細
- `PP_EfficientDetail` - 效率明細（用於計算週期時間）

---

#### 3.3.2 生產件安全存量（1425-1695行）

##### 觸發條件

- `Material.IS_WorkingPaper = '1'`
- `Material.Class = '4'` 或 `Class = '5'`（生產件）
- `Material.SafeQTY != 0`
- **庫存不足**: `SafeQTY > (在倉可用量 + 在線安全計畫量 + 多報工量 + 已存在模擬量 + 已在底稿量)`

##### 產生邏輯

1. **計算可用量**:
   ```csharp
   // 在倉可用量 = TotalStock.QTY - (TotalStockII.KeepQTY + TotalStockII.OverQTY)
   // 在線安全計畫量 = APS_Simulation.SafeQTY (PartSN=0, State='6')
   // 多報工量 = APS_PartNOTimeNote.Detail_QTY - NeedQTY (APS_StationNO!='', NoStation='0')
   // 已存在模擬量 = APS_NeedData.NeedQTY (NeedType='5', State!='6' and State!='9')
   // 已在底稿量 = APS_WorkingPaper.NeedQTY (WorkType='4' or WorkType='1')
   tmp_i = 在倉可用量 + 在線安全計畫量 + 多報工量 + 已存在模擬量 + 已在底稿量
   ```

2. **判斷是否需要開單**:
   ```csharp
   if (SafeQTY > tmp_i || (Default_WorkingPaper_NOT_SafeQTY_Order && tmp_i < 0))
   ```

3. **查找 BOM 與製程**:
   - 優先從 `APS_Simulation` 取得 `Apply_BOMId` 和 `Apply_PP_Name`
   - 若無，從 `BOM` 取得最新的 BOM（`IsConfirm='1'`, `Main_Item='1'`）

4. **查找適合儲位**:
   - 呼叫 `_SFC_Common.SelectINStore()` 選擇儲位

5. **建立或更新工作底稿**:
   - 若已存在相同 `PartNO|MBOMId|Apply_PP_Name|IN_StoreNO|IN_StoreSpacesNO` 的底稿，更新 `NeedQTY`
   - 否則新增 `APS_WorkingPaper` 記錄
   - `WorkType = '4'`（安全存量生產件）
   - 預設到貨日期: 3 天後

##### 相關資料表

- `APS_PartNOTimeNote` - 排程料號時間備註
- `BOM` - BOM 主檔

---

### 3.4 報工時自動產生單據（LabelWorkController.cs 519-1619行）

#### 3.4.1 報工時自動產生領料單

##### 觸發條件

- 工站報工完成（`LabelWorkController.SetAction`）
- `SimulationId != ''`（有排程資料）
- 本站報工數量已足夠（`Detail_QTY + Detail_Fail_QTY >= NeedQTY`）

##### 產生邏輯

1. **尋找相關 BOM 原物料**:
   - 查詢 `APS_Simulation` 中 `Apply_PP_Name` 和 `Apply_StationNO` 相同的子件
   - 計算領料數量: `tmp_int = BOMQTY * outQTY`（報工良品數量）

2. **扣 Keep 量與處理領料單**:
   - **優先順序 1**: 扣 `TotalStockII.KeepQTY`（計畫 Keep 量）
     - 依 `StoreOrder` 順序扣
     - 呼叫 `_WebSocket.Create_DOC3stock()` 建立領料單
     - 單據代碼: `AC01`
   - **優先順序 2**: 扣實體倉庫存（`TotalStock.QTY`）
     - 依 `QTY` 降序扣
   - **優先順序 3**: 扣空倉（無庫存時）
     - 呼叫 `_WebSocket.SelectOUTStore()` 選擇出庫儲位

3. **更新 APS_PartNOTimeNote**:
   - 更新 `Next_APS_StationNO`（下一站）
   - 更新 `Next_StationQTY`（下一站數量）
   - 更新 `DOCNumberNO`（領料單號）

##### 特殊情況處理

- **工單最後一站**: 自動產生入庫單（`BC01`）
  - 扣 `TotalStockII.KeepQTY`（計畫 Keep 量）
  - 若計畫量不夠，扣實體倉或建立空倉入庫單
  - 更新 `APS_PartNOTimeNote.Store_DOCNumberNO` 和 `Next_StoreQTY`

- **非最後一站**: 處理在製移轉
  - 若上一階有入庫單（`BC01`），先扣入庫數量
  - 將入庫數量 `Next_StoreQTY` 扣除回寫 `Next_StationQTY`
  - 若在製移轉不足，再扣 Keep 量或實體倉

##### 相關資料表

- `Manufacture` - 製造主檔
- `APS_PartNOTimeNote` - 排程料號時間備註

---

#### 3.4.2 報工時自動產生入庫單

##### 觸發條件

- 工單最後一站報工完成
- `APS_Simulation.PartSN = '0'`（最後一站）

##### 產生邏輯

1. **查找計畫 Keep 量**:
   - 查詢 `TotalStockII.KeepQTY > 0` 且 `SimulationId` 相符的記錄
   - 依 `StoreOrder` 順序扣

2. **建立入庫單**:
   - 呼叫 `_WebSocket.Create_DOC3stock()` 建立入庫單
   - 單據代碼: `BC01`
   - 備註: "工單:{OrderNO} 入庫"

3. **更新 APS_PartNOTimeNote**:
   - 更新 `Store_DOCNumberNO`（入庫單號）
   - 更新 `Next_StoreQTY`（入庫數量）

---

#### 3.4.3 報工時自動產生委外加工單

##### 觸發條件

- 報工完成後，判斷下一站是否為委外加工站
- `dr_APS_Simulation.Apply_StationNO = OutPackStationName`
- `APS_PartNOTimeNote.Detail_QTY` 已足夠

##### 產生邏輯

1. **查找下一站 APS_Simulation**:
   - 查詢 `Source_StationNO = Apply_StationNO` 且 `Source_StationNO_IndexSN = IndexSN + 1`

2. **計算委外數量**:
   ```csharp
   docQTY = Detail_QTY - 已存在 DOC4ProductionII.QTY
   ```

3. **查找上下階關係**:
   - 查找上階 `APS_Simulation`（`UP_SimulationId`）
   - 查找下階 `APS_Simulation`（`Down_SimulationId`）

4. **查找適合廠商與儲位**:
   - 呼叫 `_WebSocket.SelectDOC4ProductionMFNO()` 選擇委外廠商
   - 呼叫 `_WebSocket.SelectINStore()` 選擇入庫儲位

5. **建立委外加工單**:
   - 呼叫 `_WebSocket.Create_DOC4stock()` 建立委外加工單
   - 單據代碼: `PA02`
   - 備註: "工站報工,開下一站委外加工"

6. **建立工作底稿**:
   - 若 `Material.IS_WorkingPaper = '1'`，建立 `APS_WorkingPaper`
   - `WorkType = '2'`（委外加工）

---

### 3.5 工作底稿自動轉單（RUNTimeServer.cs 3142-3166行）

#### 觸發條件

- `APS_WorkingPaper.IsOK = '0'`（未完成）
- `APS_WorkingPaper.StartTime <= 當前時間`
- `Material.IS_WorkingPaper = '1'`

#### 產生邏輯

1. **採購底稿轉進貨單**:
   - `WorkType = '1'` 或 `WorkType = '3'`
   - 呼叫 `_SFC_Common.Create_DOC1stock()` 建立進貨單
   - 單據代碼: `AA02`
   - 備註: "底稿干涉採購"

2. **委外加工底稿轉委外加工單**:
   - `WorkType = '2'`
   - 呼叫 `_SFC_Common.Create_DOC4stock()` 建立委外加工單
   - 單據代碼: `PA02`
   - 備註: "底稿干涉委外加工"

3. **更新工作底稿**:
   - 更新 `APS_WorkingPaper.DOCNumberNO`
   - 更新 `APS_WorkingPaper.IsOK = '1'`（完成）

---

## 4. 單據確認流程

### 4.1 確認時機

- 單據產生後，`IsOK = '0'`（未確認）
- 需經過確認流程才能生效

### 4.2 確認方式

- **自動確認**: 系統自動產生時，部分單據可設定自動確認
- **手動確認**: 透過 `DOCconfirmController` 或 `DOC3confirmController` 確認

### 4.3 確認後處理

- 更新庫存（`TotalStock`）
- 更新排程狀態（`APS_Simulation`, `APS_PartNOTimeNote`）
- 記錄操作日誌（`OperateLog`）

---

## 5. 單據簽核機制

### 5.1 簽核流程

- 重要單據需經過簽核流程
- 透過 `DOCSSignController` 或 `FlowSignController` 處理簽核

### 5.2 簽核狀態

- `IsOK = '0'` - 未簽核
- `IsOK = '1'` - 已簽核

---

## 6. 相關資料表

### 6.1 單據主檔

| 資料表 | 說明 |
|--------|------|
| `DOC1Buy` | 進貨單主檔 |
| `DOC1BuyII` | 進貨單明細 |
| `DOC3stock` | 領料/入庫單主檔 |
| `DOC3stockII` | 領料/入庫單明細 |
| `DOC4Production` | 委外加工單主檔 |
| `DOC4ProductionII` | 委外加工單明細 |

### 6.2 工作底稿

| 資料表 | 說明 |
|--------|------|
| `APS_WorkingPaper` | 工作底稿主檔 |
| `WorkType` | 1=採購, 2=委外加工, 3=安全存量非生產件, 4=安全存量生產件 |

### 6.3 排程相關

| 資料表 | 說明 |
|--------|------|
| `APS_Simulation` | 排程模擬資料 |
| `APS_NeedData` | 排程需求資料 |
| `APS_PartNOTimeNote` | 排程料號時間備註 |

### 6.4 庫存相關

| 資料表 | 說明 |
|--------|------|
| `TotalStock` | 庫存主檔 |
| `TotalStockII` | 庫存明細（Keep 量、Over 量） |

---

## 7. 相關干涉 Rule

單據管理模組與以下 Rule 相關：

- **R-04**: 缺料補齊（自動開立進貨單）
- **R-05**: 自動領料（自動產生領料單）
- **R-07**: 派工未領料（檢查並自動補齊領料單）

---

**最後更新**: 2025年1月12日
