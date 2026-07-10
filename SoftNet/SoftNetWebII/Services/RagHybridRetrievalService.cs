using SoftNetWebII.Services;
using System;
using Base.Services;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    /// <summary>
    /// RAG 混合檢索服務
    /// 結合 SQL 精確匹配和向量語義搜尋
    /// </summary>
    public class RagHybridRetrievalService
    {
        private readonly RagRetrievalService _sqlRetrieval;
        private readonly RagVectorStoreService _vectorRetrieval;
        private readonly RagAnomalyDetectionService _anomalyService;

        public RagHybridRetrievalService()
        {
            _sqlRetrieval = new RagRetrievalService();
            _vectorRetrieval = new RagVectorStoreService();
            _anomalyService = new RagAnomalyDetectionService();
        }

        /// <summary>
        /// 混合檢索：SQL 精確匹配 + 向量語義搜尋
        /// </summary>
        public async Task<List<RagRetrievalService.RetrievedData>> HybridRetrieve(
            RagIntentService.QueryIntent intent,
            string question,
            bool enableVectorSearch = true)
        {
            var allResults = new List<RagRetrievalService.RetrievedData>();

            #region 1. SQL 檢索（精確匹配，快速）
            var sqlResults = await _sqlRetrieval.RetrieveData(intent);
            allResults.AddRange(sqlResults);
            #endregion

            #region 2. 向量檢索（語義相似，如果啟用）— 一次性依 DataType 設不同 topK：Exception=15，其他=5
            if (enableVectorSearch)
            {
                try
                {
                    var dataTypes = intent.DataTypes.Any()
                        ? intent.DataTypes.ToArray()
                        : new[] { "Exception", "Intervention" };
                    var dataTypesWithTopK = dataTypes
                        .Select(dt => (dataType: dt, topK: string.Equals(dt, "Exception", StringComparison.OrdinalIgnoreCase) ? 15 : 5))
                        .ToList();
                    var start = intent.TimeRange?.StartDate;
                    var end = intent.TimeRange?.EndDate;
                    var vectorResults = await _vectorRetrieval.VectorSearchBatch(question, dataTypesWithTopK, start, end);
                    allResults.AddRange(vectorResults);
                }
                catch (Exception ex)
                {
                    await _Log.ErrorAsync($"向量搜尋失敗: {ex.Message}", true);
                }
            }
            #endregion

            #region 3. 合併和去重
            var combined = allResults
                .GroupBy(r => $"{r.DataType}_{r.DataId}") // 使用 DataType 和 DataId 作為唯一鍵
                .Select(g =>
                {
                    // 如果有多個結果，選擇相關性分數最高的
                    var best = g.OrderByDescending(r => r.RelevanceScore).First();

                    // 如果有多個來源（SQL + 向量），提高相關性分數
                    if (g.Count() > 1)
                    {
                        best.RelevanceScore = Math.Min(1.0, best.RelevanceScore * 1.2);
                    }

                    return best;
                })
                .ToList();
            #endregion

            #region 4. 異常分類和風險評分
            foreach (var result in combined)
            {
                // 只對異常資料進行分類
                if (result.DataType == "Exception" ||
                    result.DataType == "Intervention" ||
                    (result.Metadata.ContainsKey("ErrorType") ||
                     result.Metadata.ContainsKey("EventType")))
                {
                    try
                    {
                        var classification = await _anomalyService.ClassifyAnomaly(
                            result.Metadata,
                            result.Text);

                        // 將分類結果加入 Metadata
                        result.Metadata["AnomalyType"] = classification.AnomalyType.ToString();
                        result.Metadata["RiskScore"] = classification.RiskScore;
                        result.Metadata["Confidence"] = classification.Confidence;
                        result.Metadata["IsRuleBased"] = classification.IsRuleBased;

                        // 根據風險分數調整相關性分數（高風險優先）
                        result.RelevanceScore = Math.Min(1.0,
                            result.RelevanceScore * (1.0 + classification.RiskScore * 0.5));
                    }
                    catch (Exception ex)
                    {
                        // 分類失敗不影響整體流程
                        await _Log.ErrorAsync($"異常分類失敗: {ex.Message}", false);
                    }
                }
            }

            // 按風險分數和相關性重新排序
            combined = combined
                .OrderByDescending(r =>
                {
                    // 優先排序：風險分數 * 0.7 + 相關性分數 * 0.3
                    if (r.Metadata.ContainsKey("RiskScore"))
                    {
                        var riskScore = Convert.ToDouble(r.Metadata["RiskScore"]);
                        return riskScore * 0.7 + r.RelevanceScore * 0.3;
                    }
                    return r.RelevanceScore;
                })
                .Take(3000) // 最多返回 3000 筆
                .ToList();
            #endregion

            return combined;
        }

        /// <summary>
        /// 僅使用 SQL 檢索（用於向量搜尋不可用的情況）
        /// </summary>
        public async Task<List<RagRetrievalService.RetrievedData>> SqlOnlyRetrieve(
            RagIntentService.QueryIntent intent)
        {
            return await _sqlRetrieval.RetrieveData(intent);
        }

        /// <summary>
        /// 僅使用向量檢索（用於純語義搜尋）
        /// </summary>
        public async Task<List<RagRetrievalService.RetrievedData>> VectorOnlyRetrieve(
            string question,
            string dataType = null,
            int topK = 50)
        {
            return await _vectorRetrieval.VectorSearch(question, topK, dataType);
        }
    }
}
