# SoftNet MES RAG MCP 完整實作說明文件

## 📋 文件資訊

| 項目 | 內容 |
|------|------|
| **文件版本** | 2.0 |
| **建立日期** | 2026-01-01 |
| **目的** | 提供 RAG MCP 系統的完整實作說明，包含架構、程式碼、資料庫、API 等詳細資訊 |
| **適用對象** | 開發人員、系統維護人員 |

---

## 1. 系統架構總覽

### 1.1 整體架構圖

```
┌─────────────────────────────────────────────────────────────┐
│                      MES 前端網頁                            │
│              (Views/RagAI/Index.cshtml)                     │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  用戶輸入問題 → 選擇 AI 模型 → 啟用向量搜尋          │   │
│  └──────────────────┬─────────────────────────────────┘   │
└─────────────────────┼──────────────────────────────────────┘
                      │ HTTP POST /RagAI/Ask
                      ▼
┌─────────────────────────────────────────────────────────────┐
│              RagAIController.cs (Controller)                │
│  ┌──────────────────────────────────────────────────────┐ │
│  │  1. 問題理解 (RagIntentService)                       │ │
│  │  2. 資料檢索 (RagHybridRetrievalService)              │ │
│  │  3. 上下文組裝 (RagContextService)                    │ │
│  │  4. LLM 呼叫 (RagLLMService)                          │ │
│  │  5. 記錄查詢 (RAG_QueryLog)                           │ │
│  └──────────────────────────────────────────────────────┘ │
└─────────────────────┬──────────────────────────────────────┘
                      │
        ┌─────────────┴─────────────┐
        │                           │
        ▼                           ▼
┌──────────────────┐      ┌──────────────────┐
│ SQL 檢索服務      │      │ 向量檢索服務      │
│ RagRetrievalService│      │ RagVectorStoreService│
│                  │      │                  │
│ • 生產資料        │      │ • ONNX 模型      │
│ • 庫存資料        │      │ • OpenAI API    │
│ • 工單資料        │      │ • 向量相似度     │
│ • 物料資料        │      │ • 混合方案       │
└──────────────────┘      └──────────────────┘
        │                           │
        └─────────────┬─────────────┘
                      │
                      ▼
        ┌──────────────────────────┐
        │   資料庫層                │
        │                          │
        │ • SoftNetMainDB          │
        │   - Manufacture (生產)    │
        │   - Store (庫存)         │
        │   - PP_WorkOrder (工單)  │
        │   - Material (物料)       │
        │                          │
        │ • SoftNetLogDB           │
        │   - RAG_VectorStore      │
        │   - RAG_QueryLog         │
        └──────────────────────────┘
```

### 1.2 核心組件說明

| 組件 | 檔案位置 | 主要職責 |
|------|---------|---------|
| **RagAIController** | `SoftNetWebII/Controllers/RagAIController.cs` | 處理 HTTP 請求，協調各服務 |
| **RagIntentService** | `SoftNetWebII/Services/RagIntentService.cs` | 解析用戶問題，識別意圖和實體 |
| **RagRetrievalService** | `SoftNetWebII/Services/RagRetrievalService.cs` | SQL 檢索（生產、庫存、工單、物料） |
| **RagVectorStoreService** | `SoftNetWebII/Services/RagVectorStoreService.cs` | 向量搜尋（ONNX/OpenAI embedding） |
| **RagHybridRetrievalService** | `SoftNetWebII/Services/RagHybridRetrievalService.cs` | 混合檢索（SQL + 向量） |
| **RagContextService** | `SoftNetWebII/Services/RagContextService.cs` | 組裝 MCP 格式上下文 |
| **RagLLMService** | `SoftNetWebII/Services/RagLLMService.cs` | 呼叫 Gemini/OpenAI API |

---

## 2. 詳細資料流程

### 2.1 完整請求流程

```
步驟 1: 用戶提交問題
  ↓
步驟 2: RagAIController.Ask() 接收請求
  ↓
步驟 3: RagIntentService.ParseIntent()
  ├─ 提取時間範圍（今天、昨天、最近N天）
  ├─ 提取實體（料號、工單號、站別）
  └─ 識別資料類型（Production、Inventory、WorkOrder、Material）
  ↓
步驟 4: RagHybridRetrievalService.HybridRetrieve()
  ├─ SQL 檢索（RagRetrievalService）
  │   ├─ RetrieveProductionData()
  │   ├─ RetrieveInventoryData()
  │   ├─ RetrieveWorkOrderData()
  │   └─ RetrieveMaterialData()
  │
  └─ 向量檢索（RagVectorStoreService，如果啟用）
      ├─ GenerateEmbedding() → ONNX 或 OpenAI
      └─ SearchByEmbedding() → SQL Server 或 C# 計算
  ↓
步驟 5: 合併和去重結果
  ├─ 按 DataType_DataId 去重
  ├─ 提升多來源結果的相關性分數
  └─ 按相關性排序，取前 50 筆
  ↓
步驟 6: RagContextService.BuildMCPContext()
  ├─ 組裝用戶問題
  ├─ 按資料類型分組
  └─ 限制上下文長度（預設 4000 字元）
  ↓
步驟 7: RagLLMService.GenerateResponse()
  ├─ 選擇提供者（Gemini 或 OpenAI）
  ├─ 建立提示詞（包含上下文和問題）
  └─ 呼叫 API 並解析回應
  ↓
步驟 8: 記錄查詢（背景執行）
  └─ 寫入 RAG_QueryLog 表
  ↓
步驟 9: 返回結果給前端
```

### 2.2 資料檢索流程詳解

#### A. SQL 檢索流程

```csharp
// RagRetrievalService.RetrieveData()
1. 根據 QueryIntent.DataTypes 決定要檢索的資料類型
2. 對每種資料類型：
   - 建立 SQL WHERE 條件
     • 時間範圍：StartDate, EndDate
     • 實體過濾：PartNO, OrderNO, StationNO
   - 執行 SQL 查詢
   - 轉換為 RetrievedData 物件
3. 合併所有結果
4. 按 RelevanceScore 排序（SQL 檢索預設為 1.0）
5. 返回前 50 筆
```

#### B. 向量檢索流程

```csharp
// RagVectorStoreService.VectorSearch()
1. 生成查詢的 embedding
   ├─ 優先：GenerateEmbeddingFromONNX()（本地模型）
   └─ 備用：GenerateEmbeddingFromOpenAI()（API）
2. 搜尋相關資料
   ├─ 優先：SearchByEmbeddingSQL()（SQL Server 向量功能）
   └─ 備用：SearchByEmbeddingCSharp()（C# 計算）
3. 計算餘弦相似度
4. 按相似度排序
5. 返回前 topK 筆
```

#### C. 混合檢索流程

```csharp
// RagHybridRetrievalService.HybridRetrieve()
1. 執行 SQL 檢索（快速，精確匹配）
2. 執行向量檢索（語義相似，如果啟用）
3. 合併結果
4. 去重（使用 DataType_DataId 作為唯一鍵）
5. 提升多來源結果的分數（×1.2）
6. 按相關性排序
7. 返回前 50 筆
```

---

## 3. 服務詳細說明

### 3.1 RagIntentService（問題理解服務）

**檔案位置**：`SoftNetWebII/Services/RagIntentService.cs`

#### 核心方法

##### `ParseIntent(string question)`

解析用戶問題，返回 `QueryIntent` 物件。

**實作邏輯**：

```csharp
public QueryIntent ParseIntent(string question)
{
    var intent = new QueryIntent { OriginalQuestion = question };
    
    // 1. 提取時間範圍
    intent.TimeRange = ExtractTimeRange(question);
    // 支援：今天、昨天、最近N天、具體日期範圍
    
    // 2. 提取實體
    intent.Entities.AddRange(ExtractPartNumbers(question));      // 料號
    intent.Entities.AddRange(ExtractWorkOrderNumbers(question)); // 工單號
    intent.Entities.AddRange(ExtractStationNumbers(question));   // 站別
    
    // 3. 識別資料類型
    intent.DataTypes.AddRange(ExtractDataTypes(question));
    // 支援：Production, Inventory, WorkOrder, Material
    // 如果未識別，預設包含 Production 和 WorkOrder
    
    return intent;
}
```

**支援的實體提取模式**：

| 實體類型 | 支援的模式 | 範例 |
|---------|-----------|------|
| **料號** | `料號: P001`<br>`PartNO: P001`<br>`P001`<br>`P-001` | `分析料號P001的生產狀況` |
| **工單號** | `工單: WO001`<br>`OrderNO: WO001`<br>`WO001`<br>`WO-001` | `查詢工單WO001的進度` |
| **站別** | `站別: ST001`<br>`StationNO: ST001`<br>`ST001`<br>`ST-001` | `ST001工站的效率如何？` |

**支援的時間範圍**：

| 關鍵字 | 對應的時間範圍 |
|--------|--------------|
| `今天`、`本日`、`今日` | `TimeRange.Today` |
| `昨天`、`昨日` | `TimeRange.Yesterday` |
| `最近7天`、`前30天` | `TimeRange.LastDays(N)` |
| `2025-01-01` | 具體日期 |
| `2025-01-01 至 2025-01-07` | 日期範圍 |

**資料類型識別關鍵字**：

| 資料類型 | 關鍵字 |
|---------|--------|
| **Production** | 生產、產量、製造、報工、工站、站別 |
| **Inventory** | 庫存、存貨、儲位 |
| **WorkOrder** | 工單、訂單、workorder |
| **Material** | 物料、料件、材料、material |

---

### 3.2 RagRetrievalService（SQL 檢索服務）

**檔案位置**：`SoftNetWebII/Services/RagRetrievalService.cs`

#### 核心方法

##### `RetrieveData(QueryIntent intent)`

根據查詢意圖檢索相關資料。

**實作邏輯**：

```csharp
public async Task<List<RetrievedData>> RetrieveData(QueryIntent intent)
{
    var results = new List<RetrievedData>();
    
    // 根據資料類型檢索
    foreach (var dataType in intent.DataTypes)
    {
        List<RetrievedData> data = null;
        
        switch (dataType)
        {
            case "Production":
                data = await RetrieveProductionData(intent);
                break;
            case "Inventory":
                data = await RetrieveInventoryData(intent);
                break;
            case "WorkOrder":
                data = await RetrieveWorkOrderData(intent);
                break;
            case "Material":
                data = await RetrieveMaterialData(intent);
                break;
        }
        
        if (data != null)
            results.AddRange(data);
    }
    
    // 按相關性排序，取前 50 筆
    return results.OrderByDescending(r => r.RelevanceScore).Take(50).ToList();
}
```

#### 生產資料檢索

**方法**：`RetrieveProductionData(QueryIntent intent)`

**資料來源**：`SoftNetMainDB.[dbo].[Manufacture]`

**SQL 查詢範例**：

```sql
SELECT TOP 100
    m.Id,
    m.OrderNO,
    m.PartNO,
    m.StationNO,
    m.StartTime,
    m.EndTime,
    m.OutputQty,
    mat.PartName,
    mat.Specification
FROM SoftNetMainDB.[dbo].[Manufacture] m
LEFT JOIN SoftNetMainDB.[dbo].[Material] mat 
    ON m.PartNO = mat.PartNO AND m.ServerId = mat.ServerId
WHERE m.ServerId = @ServerId
    AND m.StartTime >= @StartDate      -- 時間範圍
    AND m.StartTime <= @EndDate
    AND m.PartNO IN ('P001', 'P002')   -- 料號過濾（如果有）
    AND m.OrderNO IN ('WO001', 'WO002') -- 工單號過濾（如果有）
    AND m.StationNO IN ('ST001', 'ST002') -- 站別過濾（如果有）
ORDER BY m.StartTime DESC
```

**返回資料格式**：

```csharp
new RetrievedData
{
    DataId = "123",  // Manufacture.Id
    DataType = "Production",
    Text = "工單編號: WO001, 料號: P001, 品名: 產品A, 站別: ST001, 開始時間: 2025-01-01 08:00:00, 結束時間: 2025-01-01 12:00:00, 產量: 100",
    RelevanceScore = 1.0,  // SQL 檢索預設為完全相關
    Metadata = new Dictionary<string, object>
    {
        { "OrderNO", "WO001" },
        { "PartNO", "P001" },
        { "StationNO", "ST001" },
        { "StartTime", "2025-01-01 08:00:00" },
        { "EndTime", "2025-01-01 12:00:00" },
        { "OutputQty", 100 }
    }
}
```

#### 庫存資料檢索

**方法**：`RetrieveInventoryData(QueryIntent intent)`

**資料來源**：`SoftNetMainDB.[dbo].[Store]`

**SQL 查詢範例**：

```sql
SELECT TOP 50
    s.PartNO,
    s.StoreNO,
    s.StoreSpacesNO,
    s.QTY,
    s.WeightQty,
    mat.PartName,
    mat.Specification
FROM SoftNetMainDB.[dbo].[Store] s
LEFT JOIN SoftNetMainDB.[dbo].[Material] mat 
    ON s.PartNO = mat.PartNO AND s.ServerId = mat.ServerId
WHERE s.ServerId = @ServerId
    AND s.PartNO IN ('P001', 'P002')  -- 料號過濾（如果有）
ORDER BY s.QTY DESC
```

#### 工單資料檢索

**方法**：`RetrieveWorkOrderData(QueryIntent intent)`

**資料來源**：`SoftNetSYSDB.[dbo].[PP_WorkOrder]`

**SQL 查詢範例**：

```sql
SELECT TOP 50
    wo.Id,
    wo.OrderNO,
    wo.PartNO,
    wo.StartTime,
    wo.EndTime,
    wo.TmpClose,
    mat.PartName,
    mat.Specification
FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] wo
LEFT JOIN SoftNetMainDB.[dbo].[Material] mat 
    ON wo.PartNO = mat.PartNO AND wo.ServerId = mat.ServerId
WHERE wo.ServerId = @ServerId
    AND wo.OrderNO IN ('WO001', 'WO002')  -- 工單號過濾（如果有）
    AND wo.PartNO IN ('P001', 'P002')     -- 料號過濾（如果有）
ORDER BY wo.StartTime DESC
```

#### 物料資料檢索

**方法**：`RetrieveMaterialData(QueryIntent intent)`

**資料來源**：`SoftNetMainDB.[dbo].[Material]`

**SQL 查詢範例**：

```sql
SELECT TOP 50
    m.Id,
    m.PartNO,
    m.PartName,
    m.Specification,
    m.Class
FROM SoftNetMainDB.[dbo].[Material] m
WHERE m.ServerId = @ServerId
    AND m.PartNO IN ('P001', 'P002')  -- 料號過濾（如果有）
ORDER BY m.PartNO
```

---

### 3.3 RagVectorStoreService（向量儲存服務）

**檔案位置**：`SoftNetWebII/Services/RagVectorStoreService.cs`

#### 核心功能

1. **Embedding 生成**：使用 ONNX 模型或 OpenAI API
2. **向量搜尋**：混合方案（SQL Server 或 C# 計算）
3. **預先計算**：背景任務預先計算 embedding

#### Embedding 生成流程

```
GenerateEmbedding(text)
  ↓
1. 嘗試 ONNX 模型（GenerateEmbeddingFromONNX）
   ├─ 檢查模型文件是否存在
   ├─ 載入模型（如果存在）
   ├─ Tokenize 文字（簡化版本）
   ├─ 執行推理
   └─ 正規化向量（L2 正規化）
  ↓
2. 如果 ONNX 失敗，使用 OpenAI API（GenerateEmbeddingFromOpenAI）
   ├─ 呼叫 OpenAI Embeddings API
   ├─ 模型：text-embedding-3-small（1536 維度）
   └─ 返回向量陣列
```

#### 向量搜尋流程（混合方案）

```
SearchByEmbedding(queryEmbedding, topK, dataType)
  ↓
1. 檢查 SQL Server 向量支援（TestSQLVectorSupport）
   ├─ 檢查 EmbeddingVector 欄位是否存在
   ├─ 檢查 VECTOR_DISTANCE 函數是否可用
   └─ 快取支援狀態
  ↓
2. 如果支援 SQL Server 向量功能
   ├─ SearchByEmbeddingSQL()
   │   ├─ 轉換 float[] 為 SQL Server vector 格式
   │   ├─ 呼叫 sp_RAG_VectorSearch 預存程序
   │   └─ 返回 SQL 計算的相似度
   └─ 如果失敗，回退到 C# 計算
  ↓
3. 如果 SQL Server 不支援或失敗
   └─ SearchByEmbeddingCSharp()
       ├─ 呼叫 sp_RAG_VectorSearch_JSON 預存程序
       ├─ 從 JSON 讀取 embedding
       ├─ 計算餘弦相似度（C#）
       └─ 排序並返回結果
```

#### 餘弦相似度計算

```csharp
private double CalculateCosineSimilarity(float[] vector1, float[] vector2)
{
    if (vector1.Length != vector2.Length)
        return 0.0;
    
    double dotProduct = 0.0;
    double norm1 = 0.0;
    double norm2 = 0.0;
    
    for (int i = 0; i < vector1.Length; i++)
    {
        dotProduct += vector1[i] * vector2[i];
        norm1 += vector1[i] * vector1[i];
        norm2 += vector2[i] * vector2[i];
    }
    
    if (norm1 == 0.0 || norm2 == 0.0)
        return 0.0;
    
    return dotProduct / (Math.Sqrt(norm1) * Math.Sqrt(norm2));
}
```

**相似度範圍**：0.0（不相關）到 1.0（完全相關）

#### 預先計算 Embedding

**方法**：`PrecomputeEmbeddings()`

**執行時機**：背景服務或手動觸發

**流程**：

```csharp
1. PrecomputeProductionEmbeddings()
   ├─ 查詢最近 30 天且尚未計算 embedding 的生產資料
   ├─ 對每筆資料生成 embedding
   ├─ 同時儲存到 Embedding（JSON）和 EmbeddingVector（vector 格式）
   └─ 記錄處理數量
   
2. PrecomputeWorkOrderEmbeddings()
   └─ 類似流程
   
3. PrecomputeInventoryEmbeddings()
   └─ 類似流程
```

**SQL 查詢範例（生產資料）**：

```sql
SELECT 
    CONCAT('PROD_', m.Id) AS Id,
    m.ServerId,
    'Production' AS DataType,
    m.Id AS DataId,
    CONCAT(
        '工單編號: ', ISNULL(m.OrderNO, ''), 
        ', 料號: ', ISNULL(m.PartNO, ''),
        ', 站別: ', ISNULL(m.StationNO, ''),
        ', 開始時間: ', FORMAT(m.StartTime, 'yyyy-MM-dd HH:mm:ss'),
        ', 結束時間: ', FORMAT(ISNULL(m.EndTime, GETDATE()), 'yyyy-MM-dd HH:mm:ss'),
        ', 產量: ', ISNULL(CAST(m.OutputQty AS NVARCHAR), '0')
    ) AS Text,
    JSON_OBJECT(
        'OrderNO': ISNULL(m.OrderNO, ''),
        'PartNO': ISNULL(m.PartNO, ''),
        'StationNO': ISNULL(m.StationNO, ''),
        'StartTime': FORMAT(m.StartTime, 'yyyy-MM-dd HH:mm:ss'),
        'EndTime': FORMAT(ISNULL(m.EndTime, GETDATE()), 'yyyy-MM-dd HH:mm:ss'),
        'OutputQty': ISNULL(CAST(m.OutputQty AS NVARCHAR), '0')
    ) AS Metadata
FROM SoftNetMainDB.[dbo].[Manufacture] m
WHERE m.ServerId = @ServerId
    AND m.StartTime >= DATEADD(DAY, -30, GETDATE())
    AND NOT EXISTS (
        SELECT 1 FROM SoftNetLogDB.[dbo].[RAG_VectorStore] v
        WHERE v.DataId = CAST(m.Id AS NVARCHAR) AND v.DataType = 'Production'
    );
```

---

### 3.4 RagHybridRetrievalService（混合檢索服務）

**檔案位置**：`SoftNetWebII/Services/RagHybridRetrievalService.cs`

#### 核心方法

##### `HybridRetrieve(QueryIntent intent, string question, bool enableVectorSearch)`

結合 SQL 精確匹配和向量語義搜尋。

**實作邏輯**：

```csharp
public async Task<List<RetrievedData>> HybridRetrieve(
    QueryIntent intent,
    string question,
    bool enableVectorSearch = true)
{
    var allResults = new List<RetrievedData>();
    
    // 1. SQL 檢索（精確匹配，快速）
    var sqlResults = await _sqlRetrieval.RetrieveData(intent);
    allResults.AddRange(sqlResults);
    
    // 2. 向量檢索（語義相似，如果啟用）
    if (enableVectorSearch)
    {
        try
        {
            var dataTypes = intent.DataTypes.Any()
                ? intent.DataTypes.ToArray()
                : new[] { "Production", "WorkOrder" };
            
            foreach (var dataType in dataTypes)
            {
                var vectorResults = await _vectorRetrieval.VectorSearch(
                    question, topK: 20, dataType: dataType);
                allResults.AddRange(vectorResults);
            }
        }
        catch (Exception ex)
        {
            // 向量搜尋失敗不影響整體流程
            await _Log.ErrorAsync($"向量搜尋失敗: {ex.Message}", true);
        }
    }
    
    // 3. 合併和去重
    var combined = allResults
        .GroupBy(r => $"{r.DataType}_{r.DataId}")  // 唯一鍵
        .Select(g =>
        {
            var best = g.OrderByDescending(r => r.RelevanceScore).First();
            
            // 如果有多個來源（SQL + 向量），提高相關性分數
            if (g.Count() > 1)
            {
                best.RelevanceScore = Math.Min(1.0, best.RelevanceScore * 1.2);
            }
            
            return best;
        })
        .OrderByDescending(r => r.RelevanceScore)
        .Take(50)
        .ToList();
    
    return combined;
}
```

**去重邏輯**：

- 使用 `DataType_DataId` 作為唯一鍵
- 如果同一筆資料同時出現在 SQL 和向量結果中，選擇分數最高的
- 多來源結果的分數會提升 20%（最高不超過 1.0）

---

### 3.5 RagContextService（上下文組裝服務）

**檔案位置**：`SoftNetWebII/Services/RagContextService.cs`

#### 核心方法

##### `BuildMCPContext(List<RetrievedData> retrievedData, string question, int maxLength = 4000)`

將檢索到的資料組裝為 MCP 格式的上下文。

**輸出格式**：

```
## 用戶問題
[用戶的問題]

## 相關資料上下文
共找到 N 筆相關資料：

### 生產資料 (M 筆)
1. 工單編號: WO001, 料號: P001, 品名: 產品A, 站別: ST001, 開始時間: 2025-01-01 08:00:00, 結束時間: 2025-01-01 12:00:00, 產量: 100
2. 工單編號: WO002, 料號: P002, 品名: 產品B, 站別: ST002, 開始時間: 2025-01-01 09:00:00, 結束時間: 2025-01-01 13:00:00, 產量: 150
...

### 庫存資料 (K 筆)
1. 料號: P001, 品名: 產品A, 倉庫: A1, 儲位: S001, 數量: 500, 重量: 1000
...

### 工單資料 (L 筆)
1. 工單編號: WO001, 料號: P001, 開始時間: 2025-01-01 08:00:00, 結束時間: 2025-01-01 12:00:00
...
```

**長度限制**：

- 預設最大長度：4000 字元
- 如果超過，會截斷並加上 "..."
- 每種資料類型最多顯示 20 筆

---

### 3.6 RagLLMService（LLM 呼叫服務）

**檔案位置**：`SoftNetWebII/Services/RagLLMService.cs`

#### 核心方法

##### `GenerateResponse(string context, string question, string provider = "gemini")`

呼叫 LLM API 生成回答。

#### Gemini API 呼叫

**方法**：`CallGeminiAPI(string context, string question)`

**模型優先順序**：

1. `gemini-2.5-flash`（最新，支援 1M tokens）
2. `gemini-2.5-pro`（Pro 版本）
3. `gemini-2.0-flash-001`（穩定版本）
4. `gemini-2.0-flash`（備用）

**API 版本**：`v1`

**提示詞範本**：

```
你是一個 MES 系統的 AI 分析助手。請根據以下上下文資料回答用戶問題。

[上下文資料]

用戶問題：[用戶問題]

請提供詳細且準確的分析，並引用具體的資料。如果資料不足，請明確說明。
```

**API 請求格式**：

```json
{
  "contents": [
    {
      "parts": [
        {
          "text": "[提示詞]"
        }
      ]
    }
  ],
  "generationConfig": {
    "temperature": 0.7,
    "topK": 40,
    "topP": 0.95,
    "maxOutputTokens": 2048
  }
}
```

**API URL**：

```
https://generativelanguage.googleapis.com/v1/models/{modelName}:generateContent?key={apiKey}
```

#### OpenAI API 呼叫

**方法**：`CallOpenAIAPI(string context, string question)`

**模型**：`gpt-4o-mini`（可改為 `gpt-4o`）

**提示詞範本**：

```
你是一個 MES 系統的 AI 分析助手。請根據以下上下文資料回答用戶問題。

[上下文資料]

用戶問題：[用戶問題]

請提供詳細且準確的分析，並引用具體的資料。如果資料不足，請明確說明。
```

**API 請求格式**：

```json
{
  "model": "gpt-4o-mini",
  "messages": [
    {
      "role": "system",
      "content": "你是一個 MES 系統的 AI 分析助手。請根據提供的上下文資料回答用戶問題，並引用具體的資料。"
    },
    {
      "role": "user",
      "content": "[提示詞]"
    }
  ],
  "temperature": 0.7,
  "max_tokens": 2000
}
```

**API URL**：`https://api.openai.com/v1/chat/completions`

---

### 3.7 RagAIController（控制器）

**檔案位置**：`SoftNetWebII/Controllers/RagAIController.cs`

#### API 端點

##### `GET /RagAI/Index`

顯示 RAG AI 問答頁面。

**權限檢查**：需要登入

**返回**：`Views/RagAI/Index.cshtml`

##### `POST /RagAI/Ask`

智能問答 API。

**請求參數**：

| 參數 | 類型 | 必填 | 說明 |
|------|------|------|------|
| `question` | string | 是 | 用戶問題 |
| `provider` | string | 否 | LLM 提供者（`gemini` 或 `openai`），預設 `gemini` |
| `enableVectorSearch` | bool | 否 | 是否啟用向量搜尋，預設 `true` |

**請求範例**：

```javascript
_ajax.getJson('/RagAI/Ask', {
    question: '分析最近一週 ST001 工站的生產效率',
    provider: 'gemini',
    enableVectorSearch: true
}, successCallback, errorCallback);
```

**回應格式**：

```json
{
  "success": true,
  "question": "分析最近一週 ST001 工站的生產效率",
  "answer": "根據檢索到的資料，ST001 工站在最近一週...",
  "retrievedDataCount": 15,
  "retrievedData": [
    {
      "dataType": "Production",
      "text": "工單編號: WO001, 料號: P001, 站別: ST001...",
      "relevanceScore": 0.95
    },
    ...
  ],
  "provider": "gemini",
  "responseTime": 2345
}
```

**錯誤回應**：

```json
{
  "success": false,
  "error": "錯誤訊息"
}
```

##### `POST /RagAI/GetQueryHistory`

取得查詢歷史記錄。

**請求參數**：

| 參數 | 類型 | 必填 | 說明 |
|------|------|------|------|
| `page` | int | 否 | 頁碼，預設 1 |
| `pageSize` | int | 否 | 每頁筆數，預設 20 |

**回應格式**：

```json
{
  "success": true,
  "records": [
    {
      "id": 1,
      "question": "分析最近一週的生產狀況",
      "provider": "gemini",
      "retrievedDataCount": 15,
      "responseTime": 2345,
      "success": true,
      "createTime": "2026-01-01T10:00:00"
    },
    ...
  ],
  "page": 1,
  "pageSize": 20
}
```

---

## 4. 資料庫結構詳解

### 4.1 RAG_VectorStore（向量儲存表）

**資料庫**：`SoftNetLogDB`

**用途**：儲存 MES 資料的向量 embedding，用於語義搜尋

**完整結構**：

```sql
CREATE TABLE [dbo].[RAG_VectorStore] (
    [Id] [nvarchar](50) NOT NULL PRIMARY KEY,              -- 唯一識別碼（例如：PROD_123）
    [ServerId] [nchar](10) NOT NULL,                        -- 伺服器 ID
    [DataType] [nvarchar](50) NOT NULL,                     -- 資料類型：Production, Inventory, WorkOrder, Material
    [DataId] [nvarchar](50) NOT NULL,                       -- 原始資料的 ID
    [Text] [nvarchar](max) NOT NULL,                        -- 原始文字（用於生成 embedding）
    [Embedding] [nvarchar](max) NULL,                       -- Embedding 向量（JSON 格式，供 C# 計算使用）
    [EmbeddingVector] [nvarchar](max) NULL,                 -- Embedding 向量（SQL Server 格式，未來可升級為 vector 類型）
    [Metadata] [nvarchar](max) NULL,                        -- JSON 格式的額外資訊
    [CreateTime] [datetime] NOT NULL DEFAULT GETDATE(),     -- 建立時間
    [UpdateTime] [datetime] NOT NULL DEFAULT GETDATE(),     -- 更新時間
    
    INDEX IX_DataType_ServerId (DataType, ServerId),
    INDEX IX_DataId_DataType (DataId, DataType),
    INDEX IX_Embedding_NonClustered (DataType, ServerId) INCLUDE (Id, DataId, Text, Metadata)
);
```

**欄位說明**：

| 欄位 | 類型 | 說明 | 範例 |
|------|------|------|------|
| `Id` | nvarchar(50) | 唯一識別碼 | `PROD_123`、`INV_P001_A1_S001` |
| `ServerId` | nchar(10) | 伺服器 ID | `01` |
| `DataType` | nvarchar(50) | 資料類型 | `Production`、`Inventory`、`WorkOrder`、`Material` |
| `DataId` | nvarchar(50) | 原始資料 ID | `123`（對應 Manufacture.Id） |
| `Text` | nvarchar(max) | 原始文字 | `工單編號: WO001, 料號: P001, 站別: ST001...` |
| `Embedding` | nvarchar(max) | JSON 格式的 embedding | `[0.1,0.2,0.3,...,0.384]` |
| `EmbeddingVector` | nvarchar(max) | SQL Server 格式的 embedding | `0.1,0.2,0.3,...,0.384` |
| `Metadata` | nvarchar(max) | JSON 格式的額外資訊 | `{"OrderNO":"WO001","PartNO":"P001",...}` |

**資料範例**：

```sql
INSERT INTO RAG_VectorStore (Id, ServerId, DataType, DataId, Text, Embedding, EmbeddingVector, Metadata)
VALUES (
    'PROD_123',
    '01',
    'Production',
    '123',
    '工單編號: WO001, 料號: P001, 品名: 產品A, 站別: ST001, 開始時間: 2025-01-01 08:00:00, 結束時間: 2025-01-01 12:00:00, 產量: 100',
    '[0.1,0.2,0.3,...,0.384]',  -- JSON 格式（384 維度）
    '0.1,0.2,0.3,...,0.384',     -- SQL Server 格式
    '{"OrderNO":"WO001","PartNO":"P001","StationNO":"ST001","StartTime":"2025-01-01 08:00:00","EndTime":"2025-01-01 12:00:00","OutputQty":100}'
);
```

### 4.2 RAG_QueryLog（查詢記錄表）

**資料庫**：`SoftNetLogDB`

**用途**：記錄每次 RAG AI 查詢的詳細資訊

**完整結構**：

```sql
CREATE TABLE [dbo].[RAG_QueryLog] (
    [Id] [bigint] IDENTITY(1,1) PRIMARY KEY,                -- 自動遞增主鍵
    [ServerId] [nchar](10) NOT NULL,                        -- 伺服器 ID
    [UserNO] [nvarchar](50) NOT NULL,                       -- 使用者編號
    [Question] [nvarchar](500) NOT NULL,                    -- 用戶問題
    [Provider] [nvarchar](20) NOT NULL,                     -- LLM 提供者：gemini 或 openai
    [RetrievedDataCount] [int] NOT NULL DEFAULT 0,          -- 檢索到的資料筆數
    [Answer] [nvarchar](max) NULL,                          -- AI 回答內容
    [ResponseTime] [int] NOT NULL,                          -- 回應時間（毫秒）
    [Success] [bit] NOT NULL DEFAULT 1,                     -- 是否成功
    [Error] [nvarchar](max) NULL,                           -- 錯誤訊息（如果失敗）
    [CreateTime] [datetime] NOT NULL DEFAULT GETDATE(),      -- 查詢時間
    
    INDEX IX_ServerId_CreateTime (ServerId, CreateTime),
    INDEX IX_UserNO_CreateTime (UserNO, CreateTime)
);
```

**查詢範例**：

```sql
-- 查詢最近 10 筆記錄
SELECT TOP 10
    Question, Provider, RetrievedDataCount, ResponseTime, Success, CreateTime
FROM RAG_QueryLog
WHERE ServerId = '01'
ORDER BY CreateTime DESC;

-- 查詢特定使用者的記錄
SELECT *
FROM RAG_QueryLog
WHERE ServerId = '01' AND UserNO = 'USER001'
ORDER BY CreateTime DESC;

-- 統計查詢效能
SELECT 
    Provider,
    AVG(ResponseTime) AS AvgResponseTime,
    COUNT(*) AS QueryCount,
    SUM(CASE WHEN Success = 1 THEN 1 ELSE 0 END) AS SuccessCount
FROM RAG_QueryLog
WHERE ServerId = '01' AND CreateTime >= DATEADD(DAY, -7, GETDATE())
GROUP BY Provider;
```

### 4.3 預存程序

#### sp_RAG_VectorSearch

**用途**：使用 SQL Server 2025 向量功能進行搜尋（目前返回空結果，由 C# 處理）

**參數**：

| 參數 | 類型 | 說明 |
|------|------|------|
| `@QueryEmbedding` | NVARCHAR(MAX) | 向量資料（逗號分隔的字串） |
| `@ServerId` | NCHAR(10) | 伺服器 ID |
| `@DataType` | NVARCHAR(50) | 資料類型（NULL 表示所有類型） |
| `@TopK` | INT | 返回前 K 筆結果 |

**目前實作**：返回空結果（SQL Server 2025 Express 不支援向量功能）

**未來實作**（當 SQL Server 支援向量功能時）：

```sql
SELECT TOP (@TopK)
    [Id], [DataType], [DataId], [Text], [Metadata],
    (1.0 - VECTOR_DISTANCE([EmbeddingVector], CAST(@QueryEmbedding AS [vector](384)), COSINE)) AS SimilarityScore
FROM [dbo].[RAG_VectorStore]
WHERE [ServerId] = @ServerId
    AND [EmbeddingVector] IS NOT NULL
    AND (@DataType IS NULL OR [DataType] = @DataType)
ORDER BY VECTOR_DISTANCE([EmbeddingVector], CAST(@QueryEmbedding AS [vector](384)), COSINE) ASC;
```

#### sp_RAG_VectorSearch_JSON

**用途**：返回 JSON 格式的 embedding，供 C# 計算相似度

**參數**：同 `sp_RAG_VectorSearch`

**返回欄位**：

- `Id`
- `DataType`
- `DataId`
- `Text`
- `Embedding`（JSON 格式）
- `Metadata`

---

## 5. 前端介面說明

### 5.1 頁面結構

**檔案位置**：`SoftNetWebII/Views/RagAI/Index.cshtml`

**主要區塊**：

1. **問答設定區**
   - 問題輸入框（多行文字）
   - AI 模型選擇（Gemini 或 OpenAI）
   - 向量搜尋開關
   - 開始問答按鈕

2. **載入狀態區**
   - 顯示「正在分析中...」
   - 顯示處理步驟

3. **AI 回答區**
   - 顯示 AI 回答內容
   - 顯示回應資訊（模型、檢索筆數、回應時間）

4. **檢索資料區**
   - 顯示檢索到的相關資料
   - 顯示資料類型和相關度分數

5. **查詢歷史區**
   - 顯示歷史查詢記錄
   - 可點擊重新查詢

### 5.2 JavaScript 函數

#### `askQuestion()`

發送問答請求。

**流程**：

```javascript
1. 取得問題、提供者、向量搜尋設定
2. 顯示載入狀態
3. 發送 AJAX 請求到 /RagAI/Ask
4. 處理回應：
   - 成功：顯示回答和檢索資料
   - 失敗：顯示錯誤訊息
5. 隱藏載入狀態
```

#### `loadHistory()`

載入查詢歷史記錄。

**流程**：

```javascript
1. 發送 AJAX 請求到 /RagAI/GetQueryHistory
2. 顯示歷史記錄列表
3. 可點擊記錄重新查詢
```

---

## 6. 配置設定

### 6.1 appsettings.json

**檔案位置**：`SoftNetWebII/appsettings.json`

**必要設定**：

```json
{
  "FunConfig": {
    "GeminiApiKey": "your-gemini-api-key",
    "OpenAIApiKey": "your-openai-api-key",
    "ServerId": "01",
    "Db": "Data Source=Kurt_NB1\\SQLEXPRESS;Initial Catalog=SoftNetMainDB;..."
  }
}
```

### 6.2 ConfigDto.cs

**檔案位置**：`Base/Base/Models/ConfigDto.cs`

**已新增屬性**：

```csharp
public string GeminiApiKey { get; set; }
public string OpenAIApiKey { get; set; }
```

### 6.3 User Secrets（推薦）

**設定 Gemini API Key**：

```powershell
dotnet user-secrets set "FunConfig:GeminiApiKey" "AIzaSyCR7qGeU2G6Rv01su2nlZX_U6uCER2T7NE" --project SoftNetWebII
```

**設定 OpenAI API Key**：

```powershell
dotnet user-secrets set "FunConfig:OpenAIApiKey" "sk-..." --project SoftNetWebII
```

---

## 7. 錯誤處理機制

### 7.1 分層錯誤處理

#### Controller 層

```csharp
try
{
    // 處理請求
}
catch (Exception ex)
{
    // 記錄錯誤
    await _Log.ErrorAsync($"RagAIController.Ask 錯誤: {ex.Message} {ex.StackTrace}", true);
    
    // 記錄錯誤查詢
    await LogQuery(..., false, ex.Message);
    
    // 返回錯誤回應
    return JsonToCnt(JObject.FromObject(new { success = false, error = ex.Message }));
}
```

#### Service 層

**向量搜尋錯誤處理**：

```csharp
try
{
    // 嘗試 SQL 向量搜尋
    var sqlResults = await SearchByEmbeddingSQL(...);
    if (sqlResults != null && sqlResults.Count > 0)
        return sqlResults;
}
catch
{
    // 自動回退到 C# 計算
    return await SearchByEmbeddingCSharp(...);
}
```

**LLM API 錯誤處理**：

```csharp
// Gemini API：嘗試多個模型
foreach (var model in models)
{
    try
    {
        return await CallGeminiAPIWithModel(...);
    }
    catch (Exception ex)
    {
        if (ex.Message.Contains("404") || ex.Message.Contains("NOT_FOUND"))
            continue;  // 嘗試下一個模型
        throw;
    }
}
```

### 7.2 常見錯誤與處理

| 錯誤類型 | 原因 | 處理方式 |
|---------|------|---------|
| **API Key 未設定** | `GeminiApiKey` 或 `OpenAIApiKey` 為空 | 返回錯誤訊息，提示設定 API Key |
| **模型不可用** | Gemini 模型返回 404 | 自動嘗試下一個模型 |
| **向量搜尋失敗** | SQL Server 不支援或 ONNX 模型不存在 | 自動回退到 C# 計算或 OpenAI API |
| **資料庫連線失敗** | 資料庫無法連線 | 記錄錯誤，返回空結果 |
| **Token 限制** | 上下文太長 | 自動截斷上下文 |

---

## 8. 效能優化建議

### 8.1 資料檢索優化

#### SQL 檢索

1. **索引優化**：
   - 確保 `Manufacture` 表有 `ServerId, StartTime` 索引
   - 確保 `Store` 表有 `ServerId, PartNO` 索引
   - 確保 `PP_WorkOrder` 表有 `ServerId, OrderNO` 索引

2. **查詢限制**：
   - 使用 `TOP` 限制結果數量
   - 預設時間範圍（最近 7 天）避免全表掃描

#### 向量搜尋

1. **預先計算**：
   - 背景任務預先計算 embedding
   - 避免即時計算造成的延遲

2. **快取機制**：
   - 快取常見查詢的 embedding
   - 使用 Redis 快取查詢結果（可選）

### 8.2 LLM API 優化

1. **模型選擇**：
   - 使用較快的模型（`gemini-2.5-flash`、`gpt-4o-mini`）
   - 根據需求選擇模型（簡單問題用 Flash，複雜分析用 Pro）

2. **上下文長度**：
   - 限制上下文長度（預設 4000 字元）
   - 只包含最相關的資料（前 50 筆）

3. **並行處理**：
   - SQL 檢索和向量檢索可以並行執行（未來優化）

### 8.3 資料庫優化

1. **索引維護**：
   - 定期重建索引
   - 監控索引使用情況

2. **資料清理**：
   - 定期清理舊的向量資料（例如：超過 90 天）
   - 清理舊的查詢記錄（例如：超過 180 天）

---

## 9. 部署指南

### 9.1 前置需求

1. **資料庫**：
   - SQL Server 2019 或更新版本
   - 執行 `RAG_Database_Setup.sql` 建立資料表

2. **API Keys**：
   - Gemini API Key（必須）
   - OpenAI API Key（可選，用於 embedding 備用）

3. **ONNX 模型**（可選）：
   - 下載 `all-MiniLM-L6-v2.onnx` 模型
   - 放置到 `SoftNetWebII/Models/` 目錄

### 9.2 部署步驟

#### Step 1: 執行資料庫腳本

```powershell
sqlcmd -S "Kurt_NB1\SQLEXPRESS" -U sa -P "admin+123" -d SoftNetLogDB -i "docs\PRD\SQL_Scripts\RAG_Database_Setup.sql"
```

#### Step 2: 設定 API Keys

**方法 1：使用 User Secrets（推薦）**

```powershell
cd SoftNetWebII
dotnet user-secrets set "FunConfig:GeminiApiKey" "your-gemini-api-key"
dotnet user-secrets set "FunConfig:OpenAIApiKey" "your-openai-api-key"
```

**方法 2：修改 appsettings.json**

```json
{
  "FunConfig": {
    "GeminiApiKey": "your-gemini-api-key",
    "OpenAIApiKey": "your-openai-api-key"
  }
}
```

#### Step 3: 下載 ONNX 模型（可選）

1. 前往 https://huggingface.co/sentence-transformers/all-MiniLM-L6-v2
2. 下載 ONNX 格式的模型文件
3. 重新命名為 `all-MiniLM-L6-v2.onnx`
4. 放置到 `SoftNetWebII/Models/` 目錄

#### Step 4: 編譯和執行

```powershell
dotnet build SoftNetWebII
dotnet run --project SoftNetWebII
```

#### Step 5: 驗證部署

1. 開啟瀏覽器：`http://localhost:5000/RagAI`
2. 輸入測試問題：`分析最近一週的生產狀況`
3. 檢查是否正常返回結果

---

## 10. 測試指南

### 10.1 單元測試建議

#### 測試 RagIntentService

```csharp
[Test]
public void ParseIntent_應該正確提取時間範圍()
{
    var service = new RagIntentService();
    var intent = service.ParseIntent("分析最近7天的生產狀況");
    
    Assert.IsNotNull(intent.TimeRange);
    Assert.AreEqual(7, intent.TimeRange.Days);
}

[Test]
public void ParseIntent_應該正確提取料號()
{
    var service = new RagIntentService();
    var intent = service.ParseIntent("查詢料號P001的庫存");
    
    Assert.IsTrue(intent.Entities.Any(e => e.Type == "PartNO" && e.Value == "P001"));
}
```

#### 測試 RagRetrievalService

```csharp
[Test]
public async Task RetrieveProductionData_應該返回正確的資料()
{
    var service = new RagRetrievalService();
    var intent = new QueryIntent
    {
        DataTypes = new List<string> { "Production" },
        TimeRange = TimeRange.LastDays(7)
    };
    
    var results = await service.RetrieveData(intent);
    
    Assert.IsTrue(results.Count > 0);
    Assert.IsTrue(results.All(r => r.DataType == "Production"));
}
```

### 10.2 整合測試

#### 測試完整流程

```csharp
[Test]
public async Task Ask_應該返回正確的回答()
{
    var controller = new RagAIController();
    var result = await controller.Ask("分析最近一週的生產狀況", "gemini", true);
    
    var json = result.Content;
    var obj = JObject.Parse(json);
    
    Assert.IsTrue((bool)obj["success"]);
    Assert.IsNotNull(obj["answer"]);
    Assert.IsTrue((int)obj["retrievedDataCount"] > 0);
}
```

### 10.3 手動測試案例

#### 測試案例 1：生產資料查詢

**問題**：`分析最近一週 ST001 工站的生產效率`

**預期結果**：
- 識別資料類型：`Production`
- 識別站別：`ST001`
- 識別時間範圍：最近 7 天
- 返回相關的生產資料
- AI 生成分析報告

#### 測試案例 2：庫存查詢

**問題**：`P001 料號目前的庫存狀況如何？`

**預期結果**：
- 識別資料類型：`Inventory`
- 識別料號：`P001`
- 返回庫存資料
- AI 生成庫存分析

#### 測試案例 3：混合查詢

**問題**：`整體生產狀況如何？有哪些需要立即處理的問題？`

**預期結果**：
- 識別多種資料類型：`Production`、`WorkOrder`
- SQL 檢索 + 向量檢索
- 返回綜合分析

---

## 11. 故障排除

### 11.1 常見問題

#### 問題 1：API Key 錯誤

**症狀**：返回 `Gemini API 錯誤` 或 `OpenAI API 錯誤`

**解決方案**：
1. 檢查 API Key 是否正確設定
2. 檢查 API Key 是否有效
3. 檢查網路連線

#### 問題 2：向量搜尋返回空結果

**症狀**：`retrievedDataCount` 為 0

**可能原因**：
1. 尚未預先計算 embedding
2. ONNX 模型不存在且 OpenAI API Key 未設定
3. 資料庫中沒有相關資料

**解決方案**：
1. 執行 `PrecomputeEmbeddings()` 預先計算 embedding
2. 設定 OpenAI API Key 或下載 ONNX 模型
3. 檢查資料庫是否有資料

#### 問題 3：回應時間過長

**症狀**：`responseTime` 超過 10 秒

**可能原因**：
1. 資料量太大
2. LLM API 回應慢
3. 網路延遲

**解決方案**：
1. 限制檢索結果數量
2. 使用較快的模型（Flash 而非 Pro）
3. 檢查網路連線

#### 問題 4：SQL 查詢失敗

**症狀**：返回錯誤訊息

**可能原因**：
1. 資料表不存在
2. 欄位名稱錯誤
3. 資料庫連線失敗

**解決方案**：
1. 檢查資料表是否存在
2. 檢查 SQL 語法
3. 檢查資料庫連線字串

---

## 12. 未來改進方向

### 12.1 短期改進（1-2 個月）

1. **完整 Tokenizer**：
   - 整合完整的 BERT tokenizer
   - 下載 vocab.txt
   - 提升 embedding 品質

2. **快取機制**：
   - 實作 Redis 快取
   - 快取常見查詢的結果

3. **效能監控**：
   - 記錄 SQL 和向量搜尋的效能
   - 分析瓶頸

### 12.2 中期改進（3-6 個月）

1. **SQL Server 2025 向量功能**：
   - 當 SQL Server 正式支援時，升級到 SQL 方案
   - 建立向量索引

2. **多語言支援**：
   - 支援英文問題
   - 多語言 embedding

3. **自訂提示範本**：
   - 根據資料類型使用不同的提示範本
   - 支援使用者自訂提示

### 12.3 長期改進（6 個月以上）

1. **進階 RAG**：
   - 實作 Re-ranking
   - 實作 Query Expansion
   - 實作 Multi-hop Retrieval

2. **知識圖譜整合**：
   - 建立 MES 知識圖譜
   - 結合圖譜和向量搜尋

3. **自動化學習**：
   - 從查詢記錄中學習
   - 自動優化檢索策略

---

## 13. 參考資料

### 13.1 相關文件

- [MES_RAG_MCP_整合建議方案.md](./MES_RAG_MCP_整合建議方案.md) - 原始建議方案
- [SQL_Server_2025_向量功能整合說明.md](./SQL_Server_2025_向量功能整合說明.md) - 向量功能說明
- [ONNX_模型設定指南.md](./ONNX_模型設定指南.md) - ONNX 模型設定
- [資料庫欄位說明書.md](./資料庫欄位說明書.md) - 資料庫結構

### 13.2 外部資源

- [RAG 概念說明](https://www.pinecone.io/learn/retrieval-augmented-generation/)
- [Gemini API 文件](https://ai.google.dev/docs)
- [OpenAI Embeddings API](https://platform.openai.com/docs/guides/embeddings)
- [Microsoft.ML.OnnxRuntime 文件](https://github.com/microsoft/onnxruntime)

---

## 14. 附錄

### 14.1 完整檔案清單

| 檔案類型 | 檔案路徑 | 說明 |
|---------|---------|------|
| **Controller** | `SoftNetWebII/Controllers/RagAIController.cs` | RAG AI 控制器 |
| **Service** | `SoftNetWebII/Services/RagIntentService.cs` | 問題理解服務 |
| **Service** | `SoftNetWebII/Services/RagRetrievalService.cs` | SQL 檢索服務 |
| **Service** | `SoftNetWebII/Services/RagVectorStoreService.cs` | 向量儲存服務 |
| **Service** | `SoftNetWebII/Services/RagHybridRetrievalService.cs` | 混合檢索服務 |
| **Service** | `SoftNetWebII/Services/RagContextService.cs` | 上下文組裝服務 |
| **Service** | `SoftNetWebII/Services/RagLLMService.cs` | LLM 呼叫服務 |
| **View** | `SoftNetWebII/Views/RagAI/Index.cshtml` | 前端頁面 |
| **SQL** | `docs/PRD/SQL_Scripts/RAG_Database_Setup.sql` | 資料庫設定腳本 |
| **Config** | `Base/Base/Models/ConfigDto.cs` | 設定模型 |

### 14.2 資料流程圖

```
用戶問題
  ↓
RagIntentService.ParseIntent()
  ├─ 提取時間範圍
  ├─ 提取實體（料號、工單號、站別）
  └─ 識別資料類型
  ↓
RagHybridRetrievalService.HybridRetrieve()
  ├─ SQL 檢索（RagRetrievalService）
  │   ├─ RetrieveProductionData()
  │   ├─ RetrieveInventoryData()
  │   ├─ RetrieveWorkOrderData()
  │   └─ RetrieveMaterialData()
  │
  └─ 向量檢索（RagVectorStoreService，如果啟用）
      ├─ GenerateEmbedding() → ONNX 或 OpenAI
      └─ SearchByEmbedding() → SQL Server 或 C# 計算
  ↓
合併和去重結果
  ↓
RagContextService.BuildMCPContext()
  └─ 組裝 MCP 格式上下文
  ↓
RagLLMService.GenerateResponse()
  ├─ Gemini API 或
  └─ OpenAI API
  ↓
記錄查詢（RAG_QueryLog）
  ↓
返回結果
```

---

**文件版本**：2.0  
**最後更新**：2026-01-01  
**維護者**：開發團隊
