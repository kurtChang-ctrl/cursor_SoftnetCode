using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Base.Services;

namespace SoftNetWebII.Services
{
    /// <summary>
    /// RAG 問答流程共用服務
    /// 封裝：理解意圖 → 檢索 → 組裝上下文 → LLM 生成回答，供 Controller 與 RUNTimeServer 共用
    /// </summary>
    public class RagQueryService
    {
        private readonly RagIntentService _intentService;
        private readonly RagHybridRetrievalService _retrievalService;
        private readonly RagContextService _contextService;
        private readonly RagLLMService _llmService;

        public RagQueryService()
        {
            _intentService = new RagIntentService();
            _retrievalService = new RagHybridRetrievalService();
            _contextService = new RagContextService();
            _llmService = new RagLLMService();
        }

        /// <summary>
        /// 執行 RAG 問答流程：意圖解析（可略過）→ 混合檢索 → 組裝上下文 → LLM 生成回答
        /// </summary>
        /// <param name="question">用戶問題</param>
        /// <param name="intent">若傳入則直接使用（例如 RUNTimeServer 已組好 TimeRange）；若為 null 則由 ParseIntent(question) 解析</param>
        /// <param name="provider">LLM 提供者（gemini 或 openai）</param>
        /// <param name="enableVectorSearch">是否啟用向量搜尋</param>
        /// <returns>(回答內容, 檢索結果列表)</returns>
        public async Task<(string answer, List<RagRetrievalService.RetrievedData> retrievedData)> AskAsync(
            string question,
            RagIntentService.QueryIntent intent = null,
            string provider = "gemini",
            bool enableVectorSearch = true)
        {
            if (string.IsNullOrWhiteSpace(question))
                return (null, new List<RagRetrievalService.RetrievedData>());

            var swTotal = Stopwatch.StartNew();

            if (intent == null)
                intent = _intentService.ParseIntent(question);
            else if (string.IsNullOrEmpty(intent.OriginalQuestion))
                intent.OriginalQuestion = question;

            var swRetrieve = Stopwatch.StartNew();
            var retrievedData = await _retrievalService.HybridRetrieve(intent, question, enableVectorSearch);
            swRetrieve.Stop();

            // 避免 prompt 過大導致 LLM 超慢：保守裁切（不影響既有流程，只限制送入 LLM 的上下文大小）
            // retrievedData 仍完整回傳給呼叫端
            const int maxItemsForLLM = 8;
            const int maxTextCharsPerItem = 800;
            var retrievedForLLM = (retrievedData ?? new List<RagRetrievalService.RetrievedData>())
                .OrderByDescending(r => r?.RelevanceScore ?? 0)
                .Take(maxItemsForLLM)
                .Select(r =>
                {
                    if (r == null) return null;
                    var text = r.Text ?? "";
                    if (text.Length > maxTextCharsPerItem) text = text.Substring(0, maxTextCharsPerItem);
                    return new RagRetrievalService.RetrievedData
                    {
                        DataId = r.DataId,
                        DataType = r.DataType,
                        RelevanceScore = r.RelevanceScore,
                        Metadata = r.Metadata,
                        Text = text
                    };
                })
                .Where(r => r != null)
                .ToList();

            var swContext = Stopwatch.StartNew();
            // 為了效能與穩定 token：送 LLM 時使用精簡版上下文（避免 BuildMCPContext 的額外排序/分組成本）
            var context = _contextService.BuildSimplifiedContext(retrievedForLLM, question, maxItems: 8);
            swContext.Stop();

            var swLLM = Stopwatch.StartNew();
            var answer = await _llmService.GenerateResponse(context, question, provider);
            swLLM.Stop();

            swTotal.Stop();
            // 避免 log IO 影響上層計時（例如 RUNTimeServer 的 askMs）
            _ = _Log.InfoAsync($"RAG AskAsync provider={provider} vector={enableVectorSearch} retrieved={retrievedData?.Count ?? 0} usedForLLM={retrievedForLLM.Count} ms(retrieve/context/llm/total)={swRetrieve.ElapsedMilliseconds}/{swContext.ElapsedMilliseconds}/{swLLM.ElapsedMilliseconds}/{swTotal.ElapsedMilliseconds}");

            return (answer, retrievedData ?? new List<RagRetrievalService.RetrievedData>());
        }
    }
}
