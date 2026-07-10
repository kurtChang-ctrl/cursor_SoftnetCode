using Base;
using Base.Models;
using Base.Services;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;
using Newtonsoft.Json;
using SoftNetWebII.Tables;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using static SoftNetWebII.Services.RagIntentService;

namespace SoftNetWebII.Services
{
    /// <summary>
    /// RAG 向量儲存服務
    /// 使用本地 ONNX 模型生成 embedding，支援向量語義搜尋
    /// 支援混合方案：自動偵測 SQL Server 2025 向量功能，優先使用 SQL 計算，否則回退到 C# 計算
    /// </summary>
    public class RagVectorStoreService
    {
        private readonly HttpClient _httpClient;
        private InferenceSession _onnxSession;
        private readonly string _onnxModelPath;
        private readonly string _vocabPath;
        private readonly object _onnxLock = new object();

        // BERT Tokenizer 實例（使用單例模式快取，避免重複載入）
        private BertTokenizer _bertTokenizer = null;
        private readonly object _tokenizerLock = new object();

        // SQL Server 向量功能支援狀態
        private bool? _supportsSQLVector = null;
        private readonly object _sqlVectorTestLock = new object();

        public RagVectorStoreService()
        {
            _httpClient = new HttpClient();

            // ONNX 模型路徑（相對於應用程式根目錄）
            var appRoot = AppDomain.CurrentDomain.BaseDirectory;
            _onnxModelPath = System.IO.Path.Combine(appRoot, "Models", "all-MiniLM-L6-v2.onnx");
            _vocabPath = System.IO.Path.Combine(appRoot, "Models", "vocab.txt");

            // 載入 BERT Tokenizer（使用單例模式快取）
            LoadTokenizer();

            // 嘗試載入 ONNX 模型（如果存在）
            LoadOnnxModel();
        }

        /// <summary>
        /// 載入 BERT Tokenizer（使用單例模式快取，避免重複載入）
        /// </summary>
        private void LoadTokenizer()
        {
            try
            {
                if (File.Exists(_vocabPath))
                {
                    // 使用單例模式取得 tokenizer 實例（快取）
                    _bertTokenizer = BertTokenizer.GetInstance(_vocabPath);
                    _Log.InfoAsync($"BERT Tokenizer 載入成功: {_vocabPath}（詞彙表大小: {_bertTokenizer.VocabSize}）").Wait();
                }
                else
                {
                    _Log.InfoAsync($"vocab.txt 不存在: {_vocabPath}，將使用簡化 tokenizer").Wait();
                }
            }
            catch (Exception ex)
            {
                _Log.ErrorAsync($"載入 BERT Tokenizer 失敗: {ex.Message}，將使用簡化 tokenizer", false).Wait();
            }
        }

        /// <summary>
        /// 載入 ONNX 模型
        /// </summary>
        private void LoadOnnxModel()
        {
            try
            {
                if (File.Exists(_onnxModelPath))
                {
                    var options = new SessionOptions
                    {
                        GraphOptimizationLevel = GraphOptimizationLevel.ORT_ENABLE_ALL
                    };
                    _onnxSession = new InferenceSession(_onnxModelPath, options);
                    _Log.InfoAsync($"ONNX 模型載入成功: {_onnxModelPath}").Wait();
                }
                else
                {
                    _Log.InfoAsync($"ONNX 模型文件不存在: {_onnxModelPath}，將使用 OpenAI API 作為備用方案").Wait();
                }
            }
            catch (Exception ex)
            {
                _Log.ErrorAsync($"載入 ONNX 模型失敗: {ex.Message}", false).Wait();
            }
        }

        /// <summary>
        /// 使用向量相似度搜尋相關資料
        /// 優先使用本地 ONNX 模型生成 embedding，如果不可用則使用 OpenAI API
        /// </summary>
        public async Task<List<RagRetrievalService.RetrievedData>> VectorSearch(string queryText, int topK = 10, string dataType = null)
        {
            try
            {
                // 1. 生成查詢的 embedding
                var queryEmbedding = await GenerateEmbedding(queryText);

                if (queryEmbedding == null || queryEmbedding.Length == 0)
                {
                    return new List<RagRetrievalService.RetrievedData>();
                }

                // 2. 使用向量相似度搜尋相關資料
                return await SearchByEmbedding(queryEmbedding, topK, dataType, null, null);
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"RagVectorStoreService.VectorSearch 錯誤: {ex.Message}", true);
                return new List<RagRetrievalService.RetrievedData>();
            }
        }

        /// <summary>
        /// 向量搜尋（可指定時間範圍，讓向量檢索跟 intent.TimeRange 一致）
        /// </summary>
        public async Task<List<RagRetrievalService.RetrievedData>> VectorSearch(string queryText, int topK, string dataType, DateTime? startDate, DateTime? endDate)
        {
            try
            {
                var queryEmbedding = await GenerateEmbedding(queryText);
                if (queryEmbedding == null || queryEmbedding.Length == 0)
                    return new List<RagRetrievalService.RetrievedData>();
                return await SearchByEmbedding(queryEmbedding, topK, dataType, startDate, endDate);
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"RagVectorStoreService.VectorSearch(WithRange) 錯誤: {ex.Message}", true);
                return new List<RagRetrievalService.RetrievedData>();
            }
        }

        /// <summary>
        /// 一次性向量檢索：依多個 (DataType, topK) 並行搜尋後合併結果（例如 Exception=15、其他=5）
        /// </summary>
        public async Task<List<RagRetrievalService.RetrievedData>> VectorSearchBatch(string queryText, IReadOnlyList<(string dataType, int topK)> dataTypesWithTopK)
        {
            if (dataTypesWithTopK == null || dataTypesWithTopK.Count == 0)
                return new List<RagRetrievalService.RetrievedData>();
            try
            {
                var queryEmbedding = await GenerateEmbedding(queryText);
                if (queryEmbedding == null || queryEmbedding.Length == 0)
                    return new List<RagRetrievalService.RetrievedData>();

                var tasks = dataTypesWithTopK.Select(x => SearchByEmbedding(queryEmbedding, x.topK, x.dataType, null, null)).ToList();
                var resultLists = await Task.WhenAll(tasks).ConfigureAwait(false);
                return resultLists.SelectMany(r => r ?? new List<RagRetrievalService.RetrievedData>()).ToList();
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"RagVectorStoreService.VectorSearchBatch 錯誤: {ex.Message}", true);
                return new List<RagRetrievalService.RetrievedData>();
            }
        }

        /// <summary>
        /// 一次性向量檢索（可指定時間範圍）
        /// </summary>
        public async Task<List<RagRetrievalService.RetrievedData>> VectorSearchBatch(string queryText, IReadOnlyList<(string dataType, int topK)> dataTypesWithTopK, DateTime? startDate, DateTime? endDate)
        {
            if (dataTypesWithTopK == null || dataTypesWithTopK.Count == 0)
                return new List<RagRetrievalService.RetrievedData>();
            try
            {
                var queryEmbedding = await GenerateEmbedding(queryText);
                if (queryEmbedding == null || queryEmbedding.Length == 0)
                    return new List<RagRetrievalService.RetrievedData>();
                var tasks = dataTypesWithTopK.Select(x => SearchByEmbedding(queryEmbedding, x.topK, x.dataType, startDate, endDate)).ToList();
                var resultLists = await Task.WhenAll(tasks).ConfigureAwait(false);
                return resultLists.SelectMany(r => r ?? new List<RagRetrievalService.RetrievedData>()).ToList();
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"RagVectorStoreService.VectorSearchBatch(WithRange) 錯誤: {ex.Message}", true);
                return new List<RagRetrievalService.RetrievedData>();
            }
        }

        /// <summary>
        /// 將超過保留天數的 RAG 向量資料移至歷史表並刪除（依設定 RAG_VectorStoreRetentionDays，預設 90 天）
        /// </summary>
        public async Task<int> PurgeOldDataAsync()
        {
            int moved = 0;
            int retentionDays = _Fun.Config.RAG_VectorStoreRetentionDays > 0 ? _Fun.Config.RAG_VectorStoreRetentionDays : 90;
            try
            {
                using (var db = new DBADO("1", _Fun.Config.Db))
                {
                    string sql = "EXEC SoftNetLogDB.[dbo].[sp_RAG_VectorStore_PurgeOldData] @RetentionDays, @ServerId";
                    var table = db.DB_GetDataByDataTable(sql, new List<object> { "RetentionDays", retentionDays, "ServerId", (object)_Fun.Config.ServerId ?? DBNull.Value });
                    if (table != null && table.Rows.Count > 0 && table.Columns.Contains("RowsMovedToHistory") && table.Rows[0]["RowsMovedToHistory"] != DBNull.Value && table.Rows[0]["RowsMovedToHistory"] != null)
                        moved = Convert.ToInt32(table.Rows[0]["RowsMovedToHistory"]);
                }
                if (moved > 0)
                    await _Log.InfoAsync($"RAG 向量庫已將 {moved} 筆超過 {retentionDays} 天的資料移至歷史表");
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"RAG 向量庫 PurgeOldData 失敗: {ex.Message}", true);
            }
            return moved;
        }

        /// <summary>
        /// 生成 Embedding
        /// 優先使用本地 ONNX 模型，如果不可用則回退到 OpenAI API
        /// </summary>
        private async Task<float[]> GenerateEmbedding(string text)
        {
            // 方法 1：嘗試使用本地 ONNX 模型
            var onnxEmbedding = await GenerateEmbeddingFromONNX(text);
            if (onnxEmbedding != null && onnxEmbedding.Length > 0)
            {
                return onnxEmbedding;
            }

            // 方法 2：如果 ONNX 模型不可用，使用 OpenAI API
            return await GenerateEmbeddingFromOpenAI(text);
        }

        /// <summary>
        /// 使用本地 ONNX 模型生成 Embedding
        /// 模型：all-MiniLM-L6-v2（384 維度）
        /// 注意：需要正確的 tokenizer，這裡使用簡化版本
        /// 實際應用中建議使用完整的 BERT tokenizer（如使用 HuggingFace transformers 或對應的 C# 實作）
        /// </summary>
        private async Task<float[]> GenerateEmbeddingFromONNX(string text)
        {
            if (_onnxSession == null)
            {
                return null;
            }

            try
            {
                // 使用完整的 BERT tokenizer（如果可用）
                long[] tokenIds;
                var maxLength = 512;

                if (_bertTokenizer != null)
                {
                    // 使用完整的 BERT tokenizer
                    tokenIds = _bertTokenizer.Encode(text, maxLength: maxLength, addSpecialTokens: true);
                }
                else
                {
                    // 回退到簡化 tokenizer
                    var tokens = SimpleTokenize(text);
                    tokenIds = new long[maxLength];

                    // 初始化為 0（padding）
                    for (int i = 0; i < maxLength; i++)
                    {
                        tokenIds[i] = 0;
                    }

                    // 填充 token IDs
                    var actualLength = Math.Min(tokens.Length, maxLength);
                    for (int i = 0; i < actualLength; i++)
                    {
                        tokenIds[i] = tokens[i];
                    }
                }

                // 建立輸入張量
                // all-MiniLM-L6-v2 的輸入格式：[batch_size, sequence_length]
                var inputIdsArray = tokenIds;
                var attentionMaskArray = new long[maxLength];
                var tokenTypeIdsArray = new long[maxLength];

                // 建立 attention mask（1 表示實際 token，0 表示 padding）
                // token_type_ids 用於區分句子，單句任務全部設為 0
                for (int i = 0; i < maxLength; i++)
                {
                    // 如果 token ID 不是 PAD (0)，則 attention mask 為 1
                    attentionMaskArray[i] = (inputIdsArray[i] != 0) ? 1 : 0;
                    tokenTypeIdsArray[i] = 0; // 單句任務，全部設為 0
                }

                var inputIdsTensor = new DenseTensor<long>(inputIdsArray, new[] { 1, maxLength });
                var attentionMaskTensor = new DenseTensor<long>(attentionMaskArray, new[] { 1, maxLength });
                var tokenTypeIdsTensor = new DenseTensor<long>(tokenTypeIdsArray, new[] { 1, maxLength });

                var inputs = new List<NamedOnnxValue>
                {
                    NamedOnnxValue.CreateFromTensor("input_ids", inputIdsTensor),
                    NamedOnnxValue.CreateFromTensor("attention_mask", attentionMaskTensor),
                    NamedOnnxValue.CreateFromTensor("token_type_ids", tokenTypeIdsTensor)
                };

                // 執行推理
                float[] embedding;
                lock (_onnxLock)
                {
                    using (var results = _onnxSession.Run(inputs))
                    {
                        // all-MiniLM-L6-v2 輸出格式通常是 [batch_size, 384]
                        var outputTensor = results.First().Value as DenseTensor<float>;
                        if (outputTensor != null)
                        {
                            // 有些匯出模型輸出為 last_hidden_state: [1, seqLen(512), hidden(384)]
                            // 需做 mean pooling（用 attention_mask 排除 padding）才能得到句向量 [384]
                            var dims = outputTensor.Dimensions.ToArray();
                            if (dims.Length == 3 && dims[0] == 1)
                            {
                                int seqLen = dims[1];
                                int hidden = dims[2];
                                var pooled = new float[hidden];
                                double denom = 0;

                                // outputTensor is flattened in row-major: [batch, seq, hidden]
                                for (int t = 0; t < seqLen && t < maxLength; t++)
                                {
                                    if (attentionMaskArray[t] == 0) continue;
                                    denom += 1;
                                    int baseIdx = (t * hidden);
                                    for (int h = 0; h < hidden; h++)
                                    {
                                        pooled[h] += outputTensor.Buffer.Span[baseIdx + h];
                                    }
                                }

                                if (denom > 0)
                                {
                                    for (int h = 0; h < hidden; h++)
                                        pooled[h] = (float)(pooled[h] / denom);
                                }

                                embedding = pooled;
                            }
                            else
                            {
                                embedding = outputTensor.ToArray();
                            }
                        }
                        else
                        {
                            // 如果輸出格式不同，嘗試其他方式
                            var output = results.First().AsEnumerable<float>().ToArray();
                            embedding = output;
                        }
                    }
                }

                // 正規化向量（L2 正規化）
                var norm = Math.Sqrt(embedding.Sum(x => x * x));
                if (norm > 0)
                {
                    for (int i = 0; i < embedding.Length; i++)
                    {
                        embedding[i] = (float)(embedding[i] / norm);
                    }
                }

                await _Log.InfoAsync($"使用 ONNX 模型生成 embedding（維度: {embedding.Length}）");
                return embedding;
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"GenerateEmbeddingFromONNX 錯誤: {ex.Message}", true);
                return null;
            }
        }

        /// <summary>
        /// 簡化的 tokenizer（僅供示範）
        /// 實際應用中應該使用完整的 BERT tokenizer 和 vocab.txt
        /// </summary>
        private long[] SimpleTokenize(string text)
        {
            // 簡化處理：將文字轉換為 token IDs
            // 實際應該使用 BERT tokenizer 和 vocab.txt
            var words = text.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            var tokens = new List<long>();

            // 簡化：使用字元的 hash code 作為 token ID（僅供示範）
            foreach (var word in words)
            {
                var hash = Math.Abs(word.GetHashCode());
                tokens.Add(hash % 30000); // 模擬 vocab 大小
            }

            return tokens.ToArray();
        }

        /// <summary>
        /// 使用 OpenAI API 生成 Embedding（備用方案）
        /// </summary>
        private async Task<float[]> GenerateEmbeddingFromOpenAI(string text)
        {
            try
            {
                var apiKey = _Fun.Config.OpenAIApiKey;
                if (string.IsNullOrEmpty(apiKey))
                {
                    await _Log.ErrorAsync("請設定 OpenAIApiKey 以使用向量搜尋功能", false);
                    return null;
                }

                var apiUrl = "https://api.openai.com/v1/embeddings";

                var payload = new
                {
                    model = "text-embedding-3-small", // 1536 維度
                    input = text
                };

                var json = JsonConvert.SerializeObject(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

                var response = await _httpClient.PostAsync(apiUrl, content);
                var responseContent = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    var result = JsonConvert.DeserializeObject<dynamic>(responseContent);
                    var embeddingArray = result.data[0].embedding;
                    var embedding = new List<float>();

                    foreach (var value in embeddingArray)
                    {
                        embedding.Add((float)value);
                    }

                    await _Log.InfoAsync("使用 OpenAI API 生成 embedding");
                    return embedding.ToArray();
                }
                else
                {
                    await _Log.ErrorAsync($"OpenAI Embedding API 錯誤: {responseContent}", true);
                    return null;
                }
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"GenerateEmbeddingFromOpenAI 錯誤: {ex.Message}", true);
                return null;
            }
        }

        /// <summary>
        /// 使用 embedding 向量搜尋（混合方案：自動選擇最佳方案）
        /// 優先使用 SQL Server 2025 向量功能，如果不可用則回退到 C# 計算
        /// </summary>
        private async Task<List<RagRetrievalService.RetrievedData>> SearchByEmbedding(float[] queryEmbedding, int topK, string dataType, DateTime? startDate, DateTime? endDate)
        {
            // 檢查 SQL Server 是否支援向量功能
            // 優先使用 _Fun 中的系統變數（在啟動時已測試）
            if (!_supportsSQLVector.HasValue)
            {
                // 使用 _Fun 中的系統變數（在啟動時已測試）
                _supportsSQLVector = _Fun.SupportsSQLVector;
            }

            // 如果支援 SQL Server 向量功能，優先使用
            if (_supportsSQLVector.Value)
            {
                var sqlResults = await SearchByEmbeddingSQL(queryEmbedding, topK, dataType);
                if (sqlResults != null && sqlResults.Count > 0)
                {
                    return sqlResults;
                }
                // 如果 SQL 搜尋失敗，回退到 C# 計算
                await _Log.InfoAsync("SQL Server 向量搜尋失敗，回退到 C# 計算");
            }

            // 使用 C# 計算（備用方案）
            return await SearchByEmbeddingCSharp(queryEmbedding, topK, dataType, startDate, endDate);
        }

        /// <summary>
        /// 測試 SQL Server 是否支援向量功能
        /// </summary>
        public async Task<bool> TestSQLVectorSupportAsync()
        {
            bool result = false;
            try
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    // 測試 1：檢查是否有 EmbeddingVector 欄位
                    /* 暫時先拿掉
                    string testSql1 = @"
                        SELECT COUNT(*) 
                        FROM sys.columns 
                        WHERE object_id = OBJECT_ID(N'SoftNetLogDB.[dbo].[RAG_VectorStore]') 
                          AND name = 'EmbeddingVector'
                    ";
                    var hasColumn = db.DB_GetQueryCount(testSql1) > 0;
                    */
                    var hasColumn = false;
                    if (!hasColumn)
                    {
                        await _Log.InfoAsync("SQL Server 向量功能測試：EmbeddingVector 欄位不存在");
                        result = false;
                        _Fun.SupportsSQLVector = result;
                        return result;
                    }

                    // 測試 2：檢查 EmbeddingVector 欄位中是否有實際資料
                    // 因為我們不寫入 EmbeddingVector 欄位，所以即使欄位存在，如果沒有資料也無法使用
                    string testSql2 = @"
                        SELECT COUNT(*) 
                        FROM SoftNetLogDB.[dbo].[RAG_VectorStore]
                        WHERE [EmbeddingVector] IS NOT NULL
                    ";
                    var hasData = db.DB_GetQueryCount(testSql2) > 0;

                    if (!hasData)
                    {
                        await _Log.InfoAsync("SQL Server 向量功能測試：EmbeddingVector 欄位存在但無資料，不使用 SQL Server 向量功能");
                        result = false;
                        _Fun.SupportsSQLVector = result;
                        return result;
                    }

                    // 測試 3：檢查預存程序是否存在
                    string testSql3 = @"
                        SELECT COUNT(*) 
                        FROM sys.procedures 
                        WHERE name = 'sp_RAG_VectorSearch' 
                          AND schema_id = SCHEMA_ID('dbo')
                    ";
                    var hasProcedure = db.DB_GetQueryCount(testSql3) > 0;

                    if (!hasProcedure)
                    {
                        await _Log.InfoAsync("SQL Server 向量功能測試：預存程序 sp_RAG_VectorSearch 不存在");
                        result = false;
                        _Fun.SupportsSQLVector = result;
                        return result;
                    }

                    // 如果所有條件都滿足，才返回 true
                    await _Log.InfoAsync("SQL Server 向量功能測試：支援向量功能（欄位存在且有資料，預存程序存在）");
                    result = true;
                    // 將結果存入 _Fun 變數
                    _Fun.SupportsSQLVector = result;
                    return result;
                }
            }
            catch (Exception ex)
            {
                await _Log.InfoAsync($"SQL Server 向量功能測試失敗: {ex.Message}");
                result = false;
                _Fun.SupportsSQLVector = result;
                return result;
            }
        }

        /// <summary>
        /// 使用 SQL Server 2025 向量功能搜尋
        /// </summary>
        private async Task<List<RagRetrievalService.RetrievedData>> SearchByEmbeddingSQL(float[] queryEmbedding, int topK, string dataType)
        {
            var results = new List<RagRetrievalService.RetrievedData>();

            try
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    // 將 float[] 轉換為 SQL Server vector 格式
                    // SQL Server 2025 vector 格式：逗號分隔的字串，例如 "0.1,0.2,0.3,..."
                    var embeddingString = string.Join(",", queryEmbedding);

                    // 呼叫預存程序
                    string sql = "EXEC SoftNetLogDB.[dbo].[sp_RAG_VectorSearch] @QueryEmbedding, @ServerId, @DataType, @TopK";

                    var args = new List<object>
                    {
                        "QueryEmbedding", embeddingString,
                        "ServerId", _Fun.Config.ServerId,
                        "TopK", topK
                    };

                    if (!string.IsNullOrEmpty(dataType))
                    {
                        args.Add("DataType");
                        args.Add(dataType);
                    }
                    else
                    {
                        args.Add("DataType");
                        args.Add(DBNull.Value);
                    }

                    var table = db.DB_GetDataByDataTable(sql, args);

                    // 如果返回空結果，可能是向量功能不可用
                    if (table == null || table.Rows.Count == 0)
                    {
                        return null; // 觸發回退到 C# 計算
                    }

                    foreach (DataRow row in table.Rows)
                    {
                        var metadata = new Dictionary<string, object>();
                        try
                        {
                            if (row["Metadata"] != DBNull.Value && !string.IsNullOrEmpty(row["Metadata"].ToString()))
                            {
                                metadata = JsonConvert.DeserializeObject<Dictionary<string, object>>(
                                    row["Metadata"].ToString());
                            }
                        }
                        catch { }

                        results.Add(new RagRetrievalService.RetrievedData
                        {
                            DataId = row["DataId"].ToString(),
                            DataType = row["DataType"].ToString(),
                            Text = row["Text"].ToString(),
                            RelevanceScore = Convert.ToDouble(row["SimilarityScore"]), // SQL 計算的相似度
                            Metadata = metadata
                        });
                    }

                    await _Log.InfoAsync($"使用 SQL Server 向量功能搜尋，找到 {results.Count} 筆結果");
                    return results;
                }
            }
            catch (Exception ex)
            {
                // SQL Server 向量功能不可用或發生錯誤
                await _Log.InfoAsync($"SQL Server 向量搜尋失敗: {ex.Message}，將回退到 C# 計算");
                return null; // 觸發回退到 C# 計算
            }
        }

        /// <summary>
        /// 使用 C# 計算相似度（備用方案）
        /// 從 RAG_VectorStore + RAG_VectorStorePayload 讀取 embedding，只取 90 天內資料，計算餘弦相似度後再取 topK 的 Text/Metadata
        /// </summary>
        private async Task<List<RagRetrievalService.RetrievedData>> SearchByEmbeddingCSharp(float[] queryEmbedding, int topK, string dataType, DateTime? startDate, DateTime? endDate)
        {
            var results = new List<RagRetrievalService.RetrievedData>();
            int retentionDays = _Fun.Config.RAG_VectorStoreRetentionDays > 0 ? _Fun.Config.RAG_VectorStoreRetentionDays : 90;

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                try
                {
                    // 第一階段：只取 Id, DataType, DataId, Embedding（主表 JOIN Payload，90 天內）
                    string sql = $@"
                        SELECT v.[Id], v.[DataType], v.[DataId], p.[Embedding]
                        FROM SoftNetLogDB.[dbo].[RAG_VectorStore] v
                        INNER JOIN SoftNetLogDB.[dbo].[RAG_VectorStorePayload] p ON p.[Id] = v.[Id]
                        WHERE v.[ServerId] = @ServerId AND p.[Embedding] IS NOT NULL
                          {(startDate.HasValue ? "AND v.[CreateTime] >= @StartDate" : "AND v.[CreateTime] >= DATEADD(DAY, -@RetentionDays, GETDATE())")}
                          {(endDate.HasValue ? "AND v.[CreateTime] <= @EndDate" : "")}
                          {(string.IsNullOrEmpty(dataType) ? "" : "AND v.[DataType] = @DataType")}
                    ";
                    var args = new List<object> { "ServerId", _Fun.Config.ServerId, "RetentionDays", retentionDays };
                    if (startDate.HasValue) { args.Add("StartDate"); args.Add(startDate.Value); }
                    if (endDate.HasValue) { args.Add("EndDate"); args.Add(endDate.Value); }
                    if (!string.IsNullOrEmpty(dataType)) { args.Add("DataType"); args.Add(dataType); }
                    var table = db.DB_GetDataByDataTable(sql, args);
                    if (table != null && table.Rows.Count == 0 && (startDate.HasValue || endDate.HasValue))
                    {
                        // 指定時間範圍查不到時，退回只用 retentionDays 的查詢，避免因資料落點問題造成全程 0
                        await _Log.InfoAsync($"C# 向量搜尋：指定時間範圍無候選（start/end），改用 retentionDays fallback dataType={(string.IsNullOrEmpty(dataType) ? "(all)" : dataType)}");
                        string sqlNoRange = $@"
                            SELECT v.[Id], v.[DataType], v.[DataId], p.[Embedding]
                            FROM SoftNetLogDB.[dbo].[RAG_VectorStore] v
                            INNER JOIN SoftNetLogDB.[dbo].[RAG_VectorStorePayload] p ON p.[Id] = v.[Id]
                            WHERE v.[ServerId] = @ServerId AND p.[Embedding] IS NOT NULL
                              AND v.[CreateTime] >= DATEADD(DAY, -@RetentionDays, GETDATE())
                              {(string.IsNullOrEmpty(dataType) ? "" : "AND v.[DataType] = @DataType")}
                        ";
                        var argsNoRange = new List<object> { "ServerId", _Fun.Config.ServerId, "RetentionDays", retentionDays };
                        if (!string.IsNullOrEmpty(dataType)) { argsNoRange.Add("DataType"); argsNoRange.Add(dataType); }
                        table = db.DB_GetDataByDataTable(sqlNoRange, argsNoRange);
                    }

                    var scoredResults = new List<(string id, string dataTypeStr, string dataId, double score)>();
                    int fetchedCandidates = table.Rows.Count;
                    int skippedNullEmbedding = 0;
                    int skippedDimMismatch = 0;
                    int usedForSimilarity = 0;

                    foreach (DataRow row in table.Rows)
                    {
                        try
                        {
                            if (row["Embedding"] == DBNull.Value || string.IsNullOrEmpty(row["Embedding"].ToString()))
                            {
                                skippedNullEmbedding++;
                                continue;
                            }
                            var storedEmbedding = JsonConvert.DeserializeObject<float[]>(row["Embedding"].ToString());
                            if (storedEmbedding == null || storedEmbedding.Length == 0)
                            {
                                skippedNullEmbedding++;
                                continue;
                            }
                            if (storedEmbedding.Length != queryEmbedding.Length)
                            {
                                skippedDimMismatch++;
                                continue;
                            }
                            var similarity = CalculateCosineSimilarity(queryEmbedding, storedEmbedding);
                            scoredResults.Add((row["Id"].ToString(), row["DataType"].ToString(), row["DataId"].ToString(), similarity));
                            usedForSimilarity++;
                        }
                        catch (Exception ex)
                        {
                            await _Log.ErrorAsync($"處理向量資料時發生錯誤 (Id: {row["Id"]}): {ex.Message}", true);
                        }
                    }

                    var topKIds = scoredResults.OrderByDescending(x => x.score).Take(topK).ToList();
                    await _Log.InfoAsync(
                        $"C# 向量搜尋統計 dataType={(string.IsNullOrEmpty(dataType) ? "(all)" : dataType)} fetched={fetchedCandidates} used={usedForSimilarity} skipNullOrEmpty={skippedNullEmbedding} skipDimMismatch={skippedDimMismatch} topK={topKIds.Count}");
                    if (topKIds.Count == 0) { return results; }

                    // 第二階段：只對 topK 的 Id 取 Text, Metadata
                    var idList = string.Join(",", topKIds.Select(x => $"'{x.id.Replace("'", "''")}'"));
                    string sql2 = $@"SELECT [Id], [Text], [Metadata] FROM SoftNetLogDB.[dbo].[RAG_VectorStorePayload] WHERE [Id] IN ({idList})";
                    var table2 = db.DB_GetDataByDataTable(sql2, new List<object>(), false);
                    var payloadById = new Dictionary<string, (string text, string metadataJson)>();
                    foreach (DataRow r in table2.Rows)
                    {
                        var id = r["Id"].ToString();
                        payloadById[id] = (r["Text"]?.ToString() ?? "", r["Metadata"]?.ToString() ?? "{}");
                    }

                    foreach (var t in topKIds)
                    {
                        if (!payloadById.TryGetValue(t.id, out var payload)) continue;
                        var metadata = new Dictionary<string, object>();
                        try
                        {
                            if (!string.IsNullOrEmpty(payload.metadataJson))
                                metadata = JsonConvert.DeserializeObject<Dictionary<string, object>>(payload.metadataJson) ?? metadata;
                        }
                        catch { }
                        results.Add(new RagRetrievalService.RetrievedData
                        {
                            DataId = t.dataId,
                            DataType = t.dataTypeStr,
                            Text = payload.text,
                            RelevanceScore = t.score,
                            Metadata = metadata
                        });
                    }

                    await _Log.InfoAsync($"使用 C# 計算相似度，找到 {results.Count} 筆結果");
                }
                catch (Exception ex)
                {
                    await _Log.ErrorAsync($"向量搜尋失敗: {ex.Message}", true);
                }
            }

            return results;
        }

        /// <summary>
        /// 計算餘弦相似度
        /// </summary>
        private double CalculateCosineSimilarity(float[] vector1, float[] vector2)
        {
            if (vector1.Length != vector2.Length)
            {
                return 0.0;
            }

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
            {
                return 0.0;
            }

            return dotProduct / (Math.Sqrt(norm1) * Math.Sqrt(norm2));
        }

        /// <summary>
        /// 預先計算並儲存 embedding（背景任務）
        /// </summary>
        public async Task PrecomputeEmbeddings(DateTime timeRange)
        {
            DBADO db = new DBADO("1", _Fun.Config.Db);
            try
            {
                // 1. 從生產資料表取得需要轉換的資料
                await PrecomputeProductionEmbeddings(db, timeRange);

                // 2. 從工單資料表取得需要轉換的資料
                await PrecomputeWorkOrderEmbeddings(db, timeRange);

                // 3. 從庫存資料表取得需要轉換的資料
                await PrecomputeInventoryEmbeddings(db, timeRange);

                // 3. 從庫存資料表取得需要轉換的資料
                await PrecomputeMaterialEmbeddings(db, timeRange);

                // 3. 從異常資料表取得需要轉換的資料
                await PrecomputeExceptionEmbeddings(db, timeRange);

                // 3. 從干涉資料表取得需要轉換的資料
                await PrecomputeInterventionEmbeddings(db, timeRange);

                // 3. 從APS資料表取得需要轉換的資料
                await PrecomputeAPSEmbeddings(db, timeRange);
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"RAG 向量資料預先計算 PrecomputeEmbeddings 錯誤: {ex.Message}", true);
            }
            db.Dispose();
        }

        /// <summary>
        /// 僅補算 Exception 的 embedding（給 AIReport/測試用，避免因其他 DataType 的 lastCreateTime 影響）
        /// </summary>
        public async Task PrecomputeExceptionEmbeddingsOnly(DateTime timeRange)
        {
            using (var db = new DBADO("1", _Fun.Config.Db))
            {
                await PrecomputeExceptionEmbeddings(db, timeRange);
            }
        }

        /// <summary>
        /// 預先計算生產資料的 embedding
        /// </summary>
        private async Task PrecomputeProductionEmbeddings(DBADO db, DateTime timeRange)
        {
            /*
            // 取得最近 30 天且尚未計算 embedding 的生產資料
            string selectSql = @"
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
                ";

            var sourceData = db.DB_GetDataByDataTable(selectSql,
                new List<object> { "ServerId", _Fun.Config.ServerId });

            int count = 0;
            foreach (DataRow row in sourceData.Rows)
            {
                try
                {
                    var text = row["Text"].ToString();
                    var embedding = await GenerateEmbedding(text);

                    if (embedding != null && embedding.Length > 0)
                    {
                        // 將 embedding 轉換為 JSON 字串儲存（供 C# 計算使用）
                        var embeddingJson = JsonConvert.SerializeObject(embedding);

                        // 將 embedding 轉換為 SQL Server vector 格式（逗號分隔的字串）
                        var embeddingVectorString = string.Join(",", embedding);

                        // 嘗試同時儲存到 Embedding 和 EmbeddingVector 欄位
                        string insertSql = @"
                                INSERT INTO SoftNetLogDB.[dbo].[RAG_VectorStore]
                                ([Id], [ServerId], [DataType], [DataId], [Text], [Embedding], [EmbeddingVector], [Metadata], [CreateTime], [UpdateTime])
                                VALUES (@Id, @ServerId, @DataType, @DataId, @Text, @Embedding, CAST(@EmbeddingVector AS [vector](384)), @Metadata, GETDATE(), GETDATE());
                            ";

                        try
                        {
                            db.DB_SetDataByParams(insertSql, new List<object>
                                {
                                    "Id", row["Id"],
                                    "ServerId", row["ServerId"],
                                    "DataType", row["DataType"],
                                    "DataId", row["DataId"],
                                    "Text", row["Text"],
                                    "Embedding", embeddingJson,
                                    "EmbeddingVector", embeddingVectorString,
                                    "Metadata", row["Metadata"]
                                });
                        }
                        catch
                        {
                            // 如果 EmbeddingVector 欄位不存在或轉換失敗，只儲存 JSON 格式
                            insertSql = @"
                                    INSERT INTO SoftNetLogDB.[dbo].[RAG_VectorStore]
                                    ([Id], [ServerId], [DataType], [DataId], [Text], [Embedding], [Metadata], [CreateTime], [UpdateTime])
                                    VALUES (@Id, @ServerId, @DataType, @DataId, @Text, @Embedding, @Metadata, GETDATE(), GETDATE());
                                ";

                            db.DB_SetDataByParams(insertSql, new List<object>
                                {
                                    "Id", row["Id"],
                                    "ServerId", row["ServerId"],
                                    "DataType", row["DataType"],
                                    "DataId", row["DataId"],
                                    "Text", row["Text"],
                                    "Embedding", embeddingJson,
                                    "Metadata", row["Metadata"]
                                });
                        }

                        count++;
                    }

                    // 避免 API 速率限制
                    await Task.Delay(100);
                }
                catch (Exception ex)
                {
                    await _Log.ErrorAsync($"預先計算 embedding 錯誤 (Id: {row["Id"]}): {ex.Message}", true);
                }
            }

            if (count > 0)
            {
                await _Log.InfoAsync($"已預先計算 {count} 筆生產資料的 embedding");
            }
            */
        }

        /// <summary>
        /// 預先計算工單資料的 embedding
        /// </summary>
        private async Task PrecomputeWorkOrderEmbeddings(DBADO db, DateTime timeRange)
        {
            // 類似 PrecomputeProductionEmbeddings 的實作
            // 省略詳細實作...
        }

        /// <summary>
        /// 預先計算庫存資料的 embedding
        /// </summary>
        private async Task PrecomputeInventoryEmbeddings(DBADO db, DateTime timeRange)
        {
            // 類似 PrecomputeProductionEmbeddings 的實作
            // 省略詳細實作...
        }
        private async Task PrecomputeExceptionEmbeddings(DBADO db, DateTime timeRange)
        {
            SFC_Common _SFC_Common = new SFC_Common("1", _Fun.Config.Db);
            int totalCount = 0;

            #region OperateLog紀錄干涉
            string selectSql = $@"SELECT '1' as WriteType,'' as ErrorType,a.Id AS DataId,a.ServerId,
                        CONCAT(
                            N'工單編號: ', ISNULL(a.OrderNO, N''), 
                            N', 排程編號: ', ISNULL(a.NeedId, N''),
                            N', 子排程編號: ', ISNULL(a.SimulationId, N''),
                            N', 料號: ', ISNULL(a.PartNO, N''),
                            N', 品名: ', ISNULL(b.PartName, N''),
                            N', 規格: ', ISNULL(b.Specification, N''),
                            N', 站別: ', ISNULL(a.StationNO, N''),
                            N', 紀錄時間: ', FORMAT(a.LOGDateTime, N'yyyy-MM-dd HH:mm:ss'),
                            N', 原因: ', ISNULL(a.OperateType, N'')
                        ) AS Text,
                        -- 選取原始欄位，在 C# 端產生 JSON
                        ISNULL(a.OrderNO, N'') AS OrderNO,
                        ISNULL(a.NeedId, N'') AS NeedId,
                        ISNULL(a.SimulationId, N'') AS SimulationId,
                        ISNULL(a.PartNO, N'') AS PartNO,
                        ISNULL(b.PartName, N'') AS PartName,
                        ISNULL(b.Specification, N'') AS Specification,
                        ISNULL(a.StationNO, N'') AS StationNO,
                        FORMAT(a.LOGDateTime, N'yyyy-MM-dd HH:mm:ss') AS LOGDateTime,
                        ISNULL(a.OperateType, N'') AS OperateType
                        from SoftNetLogDB.[dbo].[OperateLog] as a
                            LEFT JOIN SoftNetMainDB.[dbo].[Material] as b ON b.PartNO = a.PartNO AND b.ServerId = a.ServerId
                        WHERE a.ServerId = @ServerId and a.LOGDateTime>@LOGDateTime and a.OperateType like N'%干涉%' 
                            AND NOT EXISTS (
                                SELECT 1
                                FROM SoftNetLogDB.[dbo].[RAG_VectorStore] v
                                INNER JOIN SoftNetLogDB.[dbo].[RAG_VectorStorePayload] p ON p.[Id] = v.[Id]
                                WHERE v.ServerId = a.ServerId
                                  AND v.DataType = 'Exception'
                                  AND v.DataId = CAST(a.Id AS NVARCHAR(50))
                                  AND p.[Embedding] IS NOT NULL
                                  AND (SELECT COUNT(*) FROM OPENJSON(p.[Embedding])) = 384
                            )
                        order by a.LOGDateTime";
            var sourceData = db.DB_GetDataByDataTable(selectSql,
                new List<object>
                    {
                                    "ServerId", _Fun.Config.ServerId,
                                    "LOGDateTime", timeRange
                    }, true);
            #endregion

            #region APS_EventLog 異常,干涉
            selectSql = $@"SELECT '2' as WriteType,'' as ErrorType,a.Id AS DataId,a.ServerId,
                        CONCAT(
                            N'單據編號: ', ISNULL(c.DOCNumberNO, N''), 
                            N', 排程編號: ', ISNULL(a.NeedId, N''),
                            N', 子排程編號: ', ISNULL(a.SimulationId, N''),
                            N', 料號: ', ISNULL(b.PartNO, N''),
                            N', 站別: ', ISNULL(c.Source_StationNO, N''),
                            N', 紀錄時間: ', FORMAT(a.LOGDateTime, N'yyyy-MM-dd HH:mm:ss'),
                            N', 原因: ', ISNULL(a.Event, N'')
                        ) AS Text,
                        -- 選取原始欄位，在 C# 端產生 JSON
                        ISNULL(c.DOCNumberNO, N'') AS DOCNumberNO,
                        ISNULL(a.NeedId, N'') AS NeedId,
                        ISNULL(a.SimulationId, N'') AS SimulationId,
                        ISNULL(b.PartNO, N'') AS PartNO,
                        ISNULL(c.Source_StationNO, N'') AS StationNO,
                        FORMAT(a.LOGDateTime, N'yyyy-MM-dd HH:mm:ss') AS LOGDateTime,
                        ISNULL(a.Event, N'') AS Event
                        from SoftNetLogDB.[dbo].[APS_EventLog] as a
                    LEFT JOIN SoftNetSYSDB.[dbo].[APS_NeedData] as b on a.NeedId=b.Id
                    LEFT JOIN SoftNetSYSDB.[dbo].[APS_Simulation] as c on a.SimulationId=c.SimulationId 
                    WHERE a.ServerId = @ServerId and a.LOGDateTime>@LOGDateTime and (a.Event like N'%異常%' or a.EventType like N'E%')
                        AND NOT EXISTS (
                            SELECT 1
                            FROM SoftNetLogDB.[dbo].[RAG_VectorStore] v
                            INNER JOIN SoftNetLogDB.[dbo].[RAG_VectorStorePayload] p ON p.[Id] = v.[Id]
                            WHERE v.ServerId = a.ServerId
                              AND v.DataType = 'Exception'
                              AND v.DataId = CAST(a.Id AS NVARCHAR(50))
                              AND p.[Embedding] IS NOT NULL
                              AND (SELECT COUNT(*) FROM OPENJSON(p.[Embedding])) = 384
                        )
                    order by a.LOGDateTime";
            var sourceData2 = db.DB_GetDataByDataTable(selectSql,
                new List<object>
                    {
                        "ServerId", _Fun.Config.ServerId,
                        "LOGDateTime", timeRange
                    }, true);
            if (sourceData2 != null && sourceData2.Rows.Count > 0)
            {
                sourceData.Merge(sourceData2);
            }
            #endregion

            #region APS_Simulation_ErrorData 異常
            string tmp_data = "";
            if (!_Fun.Config.DOCInvoice_Expiry_SendError)
            { tmp_data = " and a.ErrorType!='06' and a.ErrorType!='07' and a.ErrorType!='08' and a.ErrorType!='09' and a.ErrorType!='10'"; }
            selectSql = $@"SELECT '3' as WriteType, a.ErrorType as ErrorType, a.Id AS DataId, a.ServerId,
                        CONCAT(
                            N'單據編號: ', ISNULL(a.DOCNumberNO, N''), 
                            N', 排程編號: ', ISNULL(a.NeedId, N''),
                            N', 子排程編號: ', ISNULL(a.SimulationId, N''),
                            N', 料號: ', ISNULL(b.PartNO, N''), 
                            N', 站別: ', ISNULL(a.StationNO, N''),
                            N', 紀錄時間: ', FORMAT(a.LogDate, N'yyyy-MM-dd HH:mm:ss'),
                            N', 原因: ', ''
                        ) AS Text,
                        -- 選取原始欄位，在 C# 端產生 JSON
                        ISNULL(a.DOCNumberNO, N'') AS DOCNumberNO,
                        ISNULL(a.NeedId, N'') AS NeedId,
                        ISNULL(a.SimulationId, N'') AS SimulationId,
                        ISNULL(b.PartNO, N'') AS PartNO,
                        ISNULL(a.StationNO, N'') AS StationNO,
                        FORMAT(a.LogDate, N'yyyy-MM-dd HH:mm:ss') AS LOGDateTime,
                        ISNULL(a.ErrorType, N'') AS ErrorType
                        from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                        LEFT JOIN SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.SimulationId=b.SimulationId
                    WHERE a.ServerId = @ServerId and a.LogDate>@LogDate {tmp_data} 
                        AND NOT EXISTS (
                            SELECT 1
                            FROM SoftNetLogDB.[dbo].[RAG_VectorStore] v
                            INNER JOIN SoftNetLogDB.[dbo].[RAG_VectorStorePayload] p ON p.[Id] = v.[Id]
                            WHERE v.ServerId = a.ServerId
                              AND v.DataType = 'Exception'
                              AND v.DataId = CAST(a.Id AS NVARCHAR(50))
                              AND p.[Embedding] IS NOT NULL
                              AND (SELECT COUNT(*) FROM OPENJSON(p.[Embedding])) = 384
                        )
                    order by a.LogDate";
            sourceData2 = db.DB_GetDataByDataTable(selectSql,
                new List<object>
                    {
                        "ServerId", _Fun.Config.ServerId,
                        "LogDate", timeRange
                    },true);
            if (sourceData2 != null && sourceData2.Rows.Count > 0)
            {
                sourceData.Merge(sourceData2);
            }
            #endregion

            #region APS_WarningData 異常
            selectSql = $@"SELECT '3' as WriteType, a.ErrorType as ErrorType, a.Id AS DataId, a.ServerId,
                        CONCAT(
                            N'單據編號: ', ISNULL(a.DOCNumberNO, N''), 
                            N', 排程編號: ', ISNULL(a.NeedId, N''),
                            N', 子排程編號: ', ISNULL(a.SimulationId, N''),
                            N', 料號: ', ISNULL(b.PartNO, N''), 
                            N', 站別: ', ISNULL(a.StationNO, N''),
                            N', 紀錄時間: ', FORMAT(a.WarningDate, N'yyyy-MM-dd HH:mm:ss'),
                            N', 原因: ', ''
                        ) AS Text,
                        -- 選取原始欄位，在 C# 端產生 JSON
                        ISNULL(a.DOCNumberNO, N'') AS DOCNumberNO,
                        ISNULL(a.NeedId, N'') AS NeedId,
                        ISNULL(a.SimulationId, N'') AS SimulationId,
                        ISNULL(b.PartNO, N'') AS PartNO,
                        ISNULL(a.StationNO, N'') AS StationNO,
                        FORMAT(a.WarningDate, N'yyyy-MM-dd HH:mm:ss') AS LOGDateTime,
                        ISNULL(a.ErrorType, N'') AS ErrorType
                        FROM SoftNetSYSDB.[dbo].[APS_WarningData] as a
                        LEFT JOIN SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.SimulationId=b.SimulationId
                        WHERE a.ErrorType!='00' and a.ErrorType!='33' and a.ServerId = @ServerId and a.WarningDate>@WarningDate
                            AND NOT EXISTS (
                                SELECT 1
                                FROM SoftNetLogDB.[dbo].[RAG_VectorStore] v
                                INNER JOIN SoftNetLogDB.[dbo].[RAG_VectorStorePayload] p ON p.[Id] = v.[Id]
                                WHERE v.ServerId = a.ServerId
                                  AND v.DataType = 'Exception'
                                  AND v.DataId = CAST(a.Id AS NVARCHAR(50))
                                  AND p.[Embedding] IS NOT NULL
                                  AND (SELECT COUNT(*) FROM OPENJSON(p.[Embedding])) = 384
                            )
                        order by a.WarningDate";
            sourceData2 = db.DB_GetDataByDataTable(selectSql,
                new List<object>
                    {
                        "ServerId", _Fun.Config.ServerId,
                        "WarningDate", timeRange
                    }, true);
            if (sourceData2 != null && sourceData2.Rows.Count > 0)
            {
                sourceData.Merge(sourceData2);

            }
            #endregion
            totalCount += await ProcessExceptionDataRows(db, sourceData, _SFC_Common);
            _SFC_Common.Dispose();
            if (totalCount > 0)
            {
                await _Log.InfoAsync($"已預先計算 Exception {totalCount} 筆資料的 embedding");
            }
        }

        /// <summary>
        /// 處理異常資料的 DataRow，產生 embedding 並儲存
        /// </summary>
        private async Task<int> ProcessExceptionDataRows(DBADO db, System.Data.DataTable sourceData, SFC_Common sfcCommon)
        {
            int count = 0;
            string tmp_data = "";

            // 檢查 sourceData 是否為 null 或空
            if (sourceData == null || sourceData.Rows.Count == 0)
            {
                await _Log.InfoAsync("[ProcessExceptionDataRows] sourceData 為空，無需處理");
                return 0;
            }

            foreach (DataRow row in sourceData.Rows)
            {
                try
                {
                    string text = row["Text"].ToString();
                    string metadataJson = "";
                    // 根據 WriteType 產生對應的 JSON Metadata
                    switch (row["WriteType"].ToString())
                    {
                        case "1":
                            // OperateLog: 使用 OrderNO, NeedId, SimulationId, PartNO, PartName, Specification, StationNO, LOGDateTime, OperateType
                            var metadata1 = new
                            {
                                OrderNO = row["OrderNO"]?.ToString() ?? "",
                                NeedId = row["NeedId"]?.ToString() ?? "",
                                SimulationId = row["SimulationId"]?.ToString() ?? "",
                                PartNO = row["PartNO"]?.ToString() ?? "",
                                PartName = row["PartName"]?.ToString() ?? "",
                                Specification = row["Specification"]?.ToString() ?? "",
                                StationNO = row["StationNO"]?.ToString() ?? "",
                                LOGDateTime = row["LOGDateTime"]?.ToString() ?? "",
                                Causes = row["OperateType"]?.ToString() ?? ""
                            };
                            metadataJson = JsonConvert.SerializeObject(metadata1);
                            break;
                        case "2":
                            // APS_EventLog: 使用 DOCNumberNO, NeedId, SimulationId, PartNO, StationNO, LOGDateTime, Event
                            var metadata2 = new
                            {
                                DOCNumberNO = row["DOCNumberNO"]?.ToString() ?? "",
                                NeedId = row["NeedId"]?.ToString() ?? "",
                                SimulationId = row["SimulationId"]?.ToString() ?? "",
                                PartNO = row["PartNO"]?.ToString() ?? "",
                                StationNO = row["StationNO"]?.ToString() ?? "",
                                LOGDateTime = row["LOGDateTime"]?.ToString() ?? "",
                                Causes = row["Event"]?.ToString() ?? ""
                            };
                            metadataJson = JsonConvert.SerializeObject(metadata2);
                            break;
                        default:
                            {
                                tmp_data = sfcCommon.Get_APS_Simulation_ErrorData_TErrorType2String(row["ErrorType"]?.ToString() ?? "");
                                text = text.Replace(", 原因: ", $", 原因: {tmp_data}");
                                // 判斷是 LogDate 還是 WarningDate
                                //string logDateTime = "";
                                //if (row.Table.Columns.Contains("LogDate") && row["LogDate"] != DBNull.Value)
                                //{
                                //    logDateTime = row["LogDate"]?.ToString() ?? "";
                                //}
                                //else if (row.Table.Columns.Contains("WarningDate") && row["WarningDate"] != DBNull.Value)
                                //{
                                //    logDateTime = row["WarningDate"]?.ToString() ?? "";
                                //}

                                var metadata = new
                                {
                                    DOCNumberNO = row["DOCNumberNO"]?.ToString() ?? "",
                                    NeedId = row["NeedId"]?.ToString() ?? "",
                                    SimulationId = row["SimulationId"]?.ToString() ?? "",
                                    PartNO = row["PartNO"]?.ToString() ?? "",
                                    StationNO = row["StationNO"]?.ToString() ?? "",
                                    LOGDateTime = row["LOGDateTime"]?.ToString() ?? "",
                                    Causes = tmp_data
                                };
                                metadataJson = JsonConvert.SerializeObject(metadata);
                                break;
                            }
                    }

                    var embedding = await GenerateEmbedding(text);

                    if (embedding != null && embedding.Length > 0)
                    {
                        // 將 embedding 轉換為 JSON 字串儲存（供 C# 計算使用）
                        var embeddingJson = JsonConvert.SerializeObject(embedding);

                        // 將 embedding 轉換為 SQL Server vector 格式（逗號分隔的字串）
                        var embeddingVectorString = string.Join(",", embedding);

                        // 驗證參數值
                        var dataId = row["DataId"]?.ToString() ?? "";
                        var serverId = row["ServerId"]?.ToString() ?? _Fun.Config.ServerId;

                        if (string.IsNullOrEmpty(dataId))
                        {
                            await _Log.ErrorAsync($"ProcessExceptionDataRows: DataId 為空，跳過此筆資料", false);
                            continue;
                        }

                        // 先嘗試只儲存 JSON 格式（Embedding 欄位）
                        // 注意：SQL Server 的 vector 類型在參數化查詢中可能無法正確處理 CAST
                        // 因此暫時跳過 EmbeddingVector 欄位，只使用 Embedding JSON 欄位
                        // 如果資料表中有 EmbeddingVector 欄位且不允許 NULL，可能需要設定預設值或修改資料表結構

                        bool insertSuccess = false;
                        string errorMessage = "";

                        try
                        {
                            // 驗證 Metadata JSON 是否有效
                            if (!string.IsNullOrEmpty(metadataJson))
                            {
                                try
                                {
                                    JsonConvert.DeserializeObject(metadataJson);
                                }
                                catch (JsonException jsonEx)
                                {
                                    await _Log.ErrorAsync($"Metadata JSON 格式無效 (DataId: {dataId}): {jsonEx.Message}，將 Metadata 設為空字串", false);
                                    metadataJson = "{}";
                                }
                            }
                            else
                            {
                                metadataJson = "{}";
                            }

                            // 主表（精簡欄位）+ Payload 表（Text, Metadata, Embedding）
                            var newId = _Str.NewId('I');
                            string insertMainSql = @"
                                    INSERT INTO SoftNetLogDB.[dbo].[RAG_VectorStore]
                                    ([Id], [ServerId], [DataType], [DataId], [CreateTime], [UpdateTime])
                                    VALUES (@Id, @ServerId, @DataType, @DataId, @CreateTime, GETDATE());
                                ";
                            var mainParams = new List<object>
                            {
                                "Id", newId,
                                "ServerId", serverId,
                                "DataType", "Exception",
                                "DataId", dataId,
                                "CreateTime", row["LOGDateTime"]?.ToString() ?? ""
                            };
                            insertSuccess = db.DB_SetDataByParams(insertMainSql, mainParams);
                            if (insertSuccess)
                            {
                                string insertPayloadSql = @"
                                    INSERT INTO SoftNetLogDB.[dbo].[RAG_VectorStorePayload] ([Id], [Text], [Metadata], [Embedding])
                                    VALUES (@Id, @Text, @Metadata, @Embedding);
                                ";
                                insertSuccess = db.DB_SetDataByParams(insertPayloadSql, new List<object>
                                {
                                    "Id", newId,
                                    "Text", (text ?? "").Substring(0, Math.Min((text ?? "").Length, 4000)),
                                    "Metadata", metadataJson ?? "{}",
                                    "Embedding", embeddingJson ?? ""
                                });
                            }
                            if (insertSuccess) count++;
                        }
                        catch (Exception ex)
                        {
                            errorMessage = ex.Message;

                            // 記錄完整的錯誤資訊，包括 InnerException 和 StackTrace
                            var fullErrorMsg = $"插入資料失敗 (DataId: {dataId}): {errorMessage}";
                            if (ex.InnerException != null)
                            {
                                fullErrorMsg += $"\nInnerException: {ex.InnerException.Message}";
                                if (ex.InnerException.StackTrace != null)
                                {
                                    fullErrorMsg += $"\nInnerStackTrace: {ex.InnerException.StackTrace}";
                                }
                            }
                            if (ex.StackTrace != null)
                            {
                                fullErrorMsg += $"\nStackTrace: {ex.StackTrace}";
                            }

                            await _Log.ErrorAsync(fullErrorMsg, true);

                            // 如果錯誤訊息包含 vector 相關內容，可能是資料表結構問題
                            if (errorMessage.Contains("向量") || errorMessage.Contains("vector") || errorMessage.Contains("1998") ||
                                (ex.InnerException != null && (ex.InnerException.Message.Contains("向量") || ex.InnerException.Message.Contains("vector") || ex.InnerException.Message.Contains("1998"))))
                            {
                                await _Log.ErrorAsync($"檢測到 vector 類型錯誤，可能是資料表結構問題。建議檢查 RAG_VectorStore 表的 EmbeddingVector 欄位設定，或執行 SQL_Fix_RAG_VectorStore_EmbeddingVector.sql", true);
                            }

                            // 簡化方案：主表 + Payload（Metadata 用空 JSON）
                            try
                            {
                                var newId2 = _Str.NewId('I');
                                string insertMain2 = @"
                                    INSERT INTO SoftNetLogDB.[dbo].[RAG_VectorStore]
                                    ([Id], [ServerId], [DataType], [DataId], [CreateTime], [UpdateTime])
                                    VALUES (@Id, @ServerId, @DataType, @DataId, @CreateTime, GETDATE());
                                ";
                                insertSuccess = db.DB_SetDataByParams(insertMain2, new List<object>
                                {
                                    "Id", newId2, "ServerId", serverId, "DataType", "Exception", "DataId", dataId,
                                    "CreateTime", row["LOGDateTime"]?.ToString() ?? ""
                                });
                                if (insertSuccess)
                                {
                                    string insertPayload2 = @"
                                        INSERT INTO SoftNetLogDB.[dbo].[RAG_VectorStorePayload] ([Id], [Text], [Metadata], [Embedding])
                                        VALUES (@Id, @Text, '{}', @Embedding);
                                    ";
                                    insertSuccess = db.DB_SetDataByParams(insertPayload2, new List<object>
                                    {
                                        "Id", newId2,
                                        "Text", (text ?? "").Substring(0, Math.Min((text ?? "").Length, 4000)),
                                        "Embedding", embeddingJson ?? ""
                                    });
                                    if (insertSuccess) { count++; await _Log.InfoAsync($"成功使用簡化方案插入（不含 Metadata）(DataId: {dataId})"); }
                                }
                            }
                            catch (Exception ex2)
                            {
                                await _Log.ErrorAsync($"簡化插入方案發生異常 (DataId: {dataId}): {ex2.Message}，原始錯誤: {errorMessage}", true);
                            }
                        }
                    }

                    // 避免 API 速率限制
                    await Task.Delay(100);
                }
                catch (Exception ex)
                {
                    await _Log.ErrorAsync($"預先計算 Exception embedding 錯誤 (Id: {row["DataId"]}): {ex.Message}", true);
                }
            }

            return count;
        }
        private async Task PrecomputeMaterialEmbeddings(DBADO db, DateTime timeRange)
        {
            // 類似 PrecomputeProductionEmbeddings 的實作
            // 省略詳細實作...
        }
        private async Task PrecomputeInterventionEmbeddings(DBADO db, DateTime timeRange)
        {
            // 類似 PrecomputeProductionEmbeddings 的實作
            // 省略詳細實作...
        }
        private async Task PrecomputeAPSEmbeddings(DBADO db, DateTime timeRange)
        {
            // 類似 PrecomputeProductionEmbeddings 的實作
            // 省略詳細實作...
        }
    }
}
