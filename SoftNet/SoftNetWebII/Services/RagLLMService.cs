using Base;
using Base.Models;
using Base.Services;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    /// <summary>
    /// RAG LLM 服務
    /// 支援 Gemini 和 OpenAI API
    /// </summary>
    public class RagLLMService
    {
        private readonly HttpClient _httpClient;

        public RagLLMService()
        {
            _httpClient = new HttpClient();
            // 避免外部 API 長時間卡住拖慢整體流程
            _httpClient.Timeout = TimeSpan.FromSeconds(30);
        }

        /// <summary>
        /// 生成回應（支援 Gemini 和 OpenAI）
        /// </summary>
        public async Task<string> GenerateResponse(string context, string question, string provider = "gemini")
        {
            if (provider.ToLower() == "gemini")
            {
                return await CallGeminiAPI(context, question);
            }
            else if (provider.ToLower() == "openai")
            {
                return await CallOpenAIAPI(context, question);
            }
            else
            {
                throw new ArgumentException($"不支援的 LLM 提供者: {provider}。支援: gemini, openai");
            }
        }

        /// <summary>
        /// 呼叫 Gemini API
        /// </summary>
        private async Task<string> CallGeminiAPI(string context, string question)
        {
            var models = (_Fun.Config.GeminiModels != null && _Fun.Config.GeminiModels.Length > 0)
                ? _Fun.Config.GeminiModels
                : new[] { "gemini-2.5-flash" };

            var apiVersions = new[] { "v1" };

            Exception lastException = null;

            foreach (var apiVersion in apiVersions)
            {
                foreach (var modelName in models)
                {
                    try
                    {
                        var result = await CallGeminiAPIWithModel(context, question, modelName, apiVersion);
                        if (result != null)
                        {
                            await _Log.InfoAsync($"成功使用 Gemini 模型 {modelName} (API {apiVersion})");
                            return result;
                        }
                    }
                    catch (Exception ex)
                    {
                        lastException = ex;
                        if (ex.Message.Contains("404") || ex.Message.Contains("NOT_FOUND") || ex.Message.Contains("not found"))
                        {
                            await _Log.ErrorAsync($"模型 {modelName} (API {apiVersion}) 不可用，嘗試下一個...", false);
                            continue;
                        }
                        throw;
                    }
                }
            }

            throw new Exception($"所有 Gemini 模型都不可用。最後錯誤: {lastException?.Message}");
        }

        /// <summary>
        /// 使用指定模型呼叫 Gemini API
        /// </summary>
        private async Task<string> CallGeminiAPIWithModel(string context, string question, string modelName, string apiVersion = "v1")
        {
            try
            {
                var apiKey = _Fun.Config.GeminiApiKey;
                if (string.IsNullOrEmpty(apiKey))
                {
                    throw new Exception("請在 appsettings.json 或 user-secrets 中設定 GeminiApiKey");
                }

                var apiUrl = $"https://generativelanguage.googleapis.com/{apiVersion}/models/{modelName}:generateContent?key={apiKey}";

                var prompt = $@"你是一個 MES 系統的 AI 分析助手。請根據以下上下文資料回答用戶問題。

{context}

用戶問題：{question}

## 重要提示：
- 如果資料標記了風險等級（🔴嚴重、🟠高風險、🟡中風險、🟢低風險），請優先分析高風險項目
- 對於嚴重和高風險異常，請提供具體的處理建議和優先順序
- 請引用具體的資料編號、時間、異常類型和風險分數
- 如果資料不足，請明確說明

## 輸出格式（請務必遵守，避免冗長）
- 用條列輸出（最多 10 點）
- 每點盡量 1~2 句
- 總字數控制在 1500 字內

若資料不足，請列出「需要補哪些欄位/報表」(最多 5 點)。";

                var payload = new
                {
                    contents = new[]
                    {
                        new
                        {
                            parts = new[]
                            {
                                new { text = prompt }
                            }
                        }
                    },
                    generationConfig = new
                    {
                        temperature = 0.2,
                        topK = 40,
                        topP = 0.95,
                        // 提高上限以避免回覆被截斷（仍由輸出格式控制長度）
                        maxOutputTokens = 2048
                    }
                };

                var json = JsonConvert.SerializeObject(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(apiUrl, content);
                var responseContent = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    var result = JObject.Parse(responseContent);
                    var candidates = result["candidates"];
                    if (candidates != null && candidates.HasValues)
                    {
                        var candidate = candidates[0];
                        var finishReason = candidate?["finishReason"]?.ToString();

                        // 檢查回應是否被截斷
                        if (finishReason == "MAX_TOKENS" || finishReason == "OTHER")
                        {
                            await _Log.InfoAsync($"Gemini API 回應可能被截斷 (finishReason: {finishReason})。考慮增加 maxOutputTokens 或縮短上下文。");
                        }

                        var answer = candidate?["content"]?["parts"]?[0]?["text"]?.ToString();
                        if (string.IsNullOrEmpty(answer))
                        {
                            throw new Exception("Gemini API 回傳內容結構異常：text 為空");
                        }

                        // 如果回應被截斷，在結尾加入提示（但只在確實被截斷時才顯示）
                        if (finishReason == "MAX_TOKENS")
                        {
                            // 只在回答長度超過一定值時才顯示提示，避免誤判
                            if (answer.Length > 1000)
                            {
                                answer += "\n\n[注意：回應可能因達到最大 token 限制而被截斷。如需完整回答，請將問題拆分為多個較小的問題，或指定更明確的查詢範圍。]";
                            }
                        }

                        return answer;
                    }
                    else
                    {
                        throw new Exception("Gemini API 未返回有效回應");
                    }
                }
                else
                {
                    var errorMsg = responseContent;
                    try
                    {
                        var errorObj = JObject.Parse(responseContent);
                        if (errorObj["error"] != null)
                        {
                            errorMsg = errorObj["error"]["message"]?.ToString() ?? responseContent;
                        }
                    }
                    catch { }

                    throw new Exception($"Gemini API 錯誤 ({response.StatusCode}): {errorMsg}");
                }
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"CallGeminiAPIWithModel ({modelName}) 錯誤: {ex.Message}", false);
                throw;
            }
        }

        /// <summary>
        /// 呼叫 OpenAI API
        /// </summary>
        private async Task<string> CallOpenAIAPI(string context, string question)
        {
            try
            {
                var apiKey = _Fun.Config.OpenAIApiKey;
                if (string.IsNullOrEmpty(apiKey))
                {
                    throw new Exception("請在 appsettings.json 或 user-secrets 中設定 OpenAIApiKey");
                }

                var apiUrl = "https://api.openai.com/v1/chat/completions";

                var prompt = $@"你是一個 MES 系統的 AI 分析助手。請根據以下上下文資料回答用戶問題。

{context}

用戶問題：{question}

## 重要提示：
- 如果資料標記了風險等級（🔴嚴重、🟠高風險、🟡中風險、🟢低風險），請優先分析高風險項目
- 對於嚴重和高風險異常，請提供具體的處理建議和優先順序
- 請引用具體的資料編號、時間、異常類型和風險分數
- 如果資料不足，請明確說明

請務必在上述輸出格式限制內完成輸出；若內容不足也請簡短說明，不要新增額外段落。";

                var payload = new
                {
                    model = "gpt-4o-mini", // 或 gpt-4o
                    messages = new[]
                    {
                        new { role = "system", content = "你是一個 MES 系統的 AI 分析助手。請根據提供的上下文資料回答用戶問題，並引用具體的資料。" },
                        new { role = "user", content = prompt }
                    },
                    temperature = 0.7,
                    // 提高上限以避免回覆被截斷（仍由輸出格式控制長度）
                    max_tokens = 2048
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
                    return result.choices[0].message.content.ToString();
                }
                else
                {
                    var errorMsg = responseContent;
                    try
                    {
                        var errorObj = JObject.Parse(responseContent);
                        if (errorObj["error"] != null)
                        {
                            errorMsg = errorObj["error"]["message"]?.ToString() ?? responseContent;
                        }
                    }
                    catch { }

                    throw new Exception($"OpenAI API 錯誤 ({response.StatusCode}): {errorMsg}");
                }
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"CallOpenAIAPI 錯誤: {ex.Message}", true);
                throw;
            }
        }
    }
}
