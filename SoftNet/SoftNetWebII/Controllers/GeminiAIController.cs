using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    /// <summary>
    /// Gemini AI 分析 Controller
    /// 提供前端呼叫，查詢生產資料並使用 Gemini AI 進行分析
    /// </summary>
    public class GeminiAIController : ApiCtrl
    {
        private readonly HttpClient _httpClient;

        public GeminiAIController()
        {
            _httpClient = new HttpClient();
        }

        /// <summary>
        /// 顯示 AI 分析頁面
        /// </summary>
        public ActionResult Index()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }

            return View();
        }

        /// <summary>
        /// 取得前 7 天的生產資料並進行 AI 分析
        /// </summary>
        /// <param name="question">分析問題（可選）</param>
        /// <returns>分析結果</returns>
        [HttpPost]
        public async Task<ContentResult> AnalyzeProduction(string question = null)
        {
            try
            {
                // 1. 查詢前 7 天的生產資料
                var productionData = await GetProductionDataLast7Days();
                if (productionData == null || productionData.Count == 0)
                {
                    return JsonToCnt(JObject.FromObject(new
                    {
                        success = false,
                        error = "沒有找到前 7 天的生產資料"
                    }));
                }

                // 2. 格式化資料供 Gemini 分析
                var dataSummary = FormatDataForGemini(productionData);

                // 3. 建立提示文字
                var prompt = string.IsNullOrEmpty(question)
                    ? GetDefaultPrompt(dataSummary)
                    : GetCustomPrompt(question, dataSummary);

                // 4. 呼叫 Gemini API
                var analysisResult = await CallGeminiAPI(prompt);

                // 5. 回傳結果
                return JsonToCnt(JObject.FromObject(new
                {
                    success = true,
                    data = productionData,
                    dataCount = productionData.Count,
                    analysis = analysisResult,
                    question = question ?? "預設分析：前 7 天生產資料分析"
                }));
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"GeminiAIController.AnalyzeProduction 錯誤: {ex.Message} {ex.StackTrace}", true);
                return JsonToCnt(JObject.FromObject(new
                {
                    success = false,
                    error = ex.Message
                }));
            }
        }

        /// <summary>
        /// 取得前 7 天的生產資料
        /// </summary>
        private Task<List<Dictionary<string, object>>> GetProductionDataLast7Days()
        {
            var result = new List<Dictionary<string, object>>();

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                var startDate = DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd HH:mm:ss");

                string sql = $@"
                    SELECT a.ServerId as 用戶編號
                      ,a.LOGDateTimeID as 紀錄時間
                      ,a.StationNO as 工站編號
                      ,a.PartNO as 料件編號
                      ,b.PartName as 料件名稱
                      ,b.Specification as 料件規格
                      ,a.EditFinishedQty as 本次良品報工量
                      ,a.EditFailedQty as 本次不良品報工量
                      ,a.OP_NO as 作業員編號
                      ,a.ReportTime as 本次報工花費秒數
                      ,a.SimulationId as 子排程編號
                      ,a.CycleTime as 實際生產CT
                      ,a.ECT as 目標CT
                      ,a.LowerCT as 最佳CT
                      ,a.UpperCT as 最差CT
                      ,a.PP_Name as 製程名稱
                      ,a.IndexSN as 製程順序
                  FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] as a 
                  LEFT JOIN SoftNetMainDB.[dbo].[Material] as b ON b.PartNO = a.PartNO AND b.ServerId = a.ServerId
                  WHERE a.ServerId = '{_Fun.Config.ServerId} AND a.LOGDateTimeID>='{startDate}' order by a.ServerId,a.LOGDateTime";
                DataTable dt = db.DB_GetData(sql);

                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        var dict = new Dictionary<string, object>();
                        foreach (DataColumn col in dt.Columns)
                        {
                            var value = row[col.ColumnName];
                            if (value == DBNull.Value)
                                dict[col.ColumnName] = null;
                            else if (value is DateTime)
                                dict[col.ColumnName] = ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
                            else
                                dict[col.ColumnName] = value;
                        }
                        result.Add(dict);
                    }
                }
            }

            return Task.FromResult(result);
        }

        /// <summary>
        /// 格式化資料供 Gemini 分析
        /// </summary>
        private string FormatDataForGemini(List<Dictionary<string, object>> data)
        {
            if (data == null || data.Count == 0)
                return "無數據";

            var sb = new StringBuilder();
            sb.AppendLine("以下是要分析的生產數據：");
            sb.AppendLine();

            // 建立表格格式
            if (data.Count > 0)
            {
                var headers = data[0].Keys;
                sb.AppendLine(string.Join(" | ", headers));
                sb.AppendLine(string.Join(" | ", headers.Select(_ => "---")));

                // 只取前 100 筆（避免提示過長）
                //var limitedData = data.Take(100).ToList();
                var limitedData = data.ToList();
                foreach (var row in limitedData)
                {
                    var values = headers.Select(header =>
                    {
                        var value = row.ContainsKey(header) ? row[header] : null;
                        return value != null ? value.ToString() : "";
                    });
                    sb.AppendLine(string.Join(" | ", values));
                }
                //if (data.Count > 100)
                //{
                //    sb.AppendLine();
                //    sb.AppendLine($"（共 {data.Count} 筆數據，僅顯示前 100 筆）");
                //}
            }

            return sb.ToString();
        }

        /// <summary>
        /// 取得預設提示
        /// </summary>
        private string GetDefaultPrompt(string dataSummary)
        {
            return $@"你是一個 MES 製造執行系統的數據分析專家。請分析以下前 7 天的生產數據。

{dataSummary}

請以繁體中文提供：
1. 生產數據摘要（總筆數、工站數量、工單數量等）
2. 關鍵發現與趨勢
3. 各工站的生產效率分析
4. 潛在問題與風險
5. 具體改善建議";
        }

        /// <summary>
        /// 取得自訂提示
        /// </summary>
        private string GetCustomPrompt(string question, string dataSummary)
        {
            return $@"你是一個 MES 製造執行系統的數據分析專家。請根據以下前 7 天的生產數據回答問題。

問題：{question}

{dataSummary}

請以繁體中文回答，提供：
1. 針對問題的具體分析
2. 相關數據支持
3. 改善建議";
        }

        /// <summary>
        /// 呼叫 Gemini API（支援自動重試備用模型）
        /// </summary>
        private async Task<string> CallGeminiAPI(string prompt)
        {
            // 嘗試的模型列表（按優先順序）：從 FunConfig:GeminiModels 讀取
            var models = (_Fun.Config.GeminiModels != null && _Fun.Config.GeminiModels.Length > 0)
                ? _Fun.Config.GeminiModels
                : new[] { "gemini-2.5-flash" };

            // 只使用 v1 API（v1beta 已不支援這些模型）
            var apiVersions = new[] { "v1" };

            Exception lastException = null;

            // 依序嘗試每個模型（使用 v1 API）
            foreach (var apiVersion in apiVersions)
            {
                foreach (var modelName in models)
                {
                    try
                    {
                        var result = await CallGeminiAPIWithModel(prompt, modelName, apiVersion);
                        if (result != null)
                        {
                            await _Log.InfoAsync($"成功使用模型 {modelName} (API {apiVersion})");
                            return result;
                        }
                    }
                    catch (Exception ex)
                    {
                        lastException = ex;
                        // 如果是 404 錯誤（模型不存在），嘗試下一個模型
                        if (ex.Message.Contains("404") || ex.Message.Contains("NOT_FOUND") || ex.Message.Contains("not found"))
                        {
                            await _Log.ErrorAsync($"模型 {modelName} (API {apiVersion}) 不可用，嘗試下一個...", false);
                            continue;
                        }
                        // 其他錯誤直接拋出
                        throw;
                    }
                }
            }

            // 所有模型和 API 版本都失敗
            throw new Exception($"所有 Gemini 模型都不可用。最後錯誤: {lastException?.Message}");
        }

        /// <summary>
        /// 使用指定模型和 API 版本呼叫 Gemini API
        /// </summary>
        private async Task<string> CallGeminiAPIWithModel(string prompt, string modelName, string apiVersion = "v1")
        {
            try
            {
                // 從 appsettings.json 取得 Gemini API Key
                var apiKey = GetGeminiApiKey();
                if (string.IsNullOrEmpty(apiKey))
                {
                    throw new Exception("請在 appsettings.json 或 user-secrets 中設定 GeminiApiKey");
                }

                // 使用指定的 API 版本（v1 或 v1beta）
                var apiUrl = $"https://generativelanguage.googleapis.com/{apiVersion}/models/{modelName}:generateContent?key={apiKey}";

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
                        temperature = 0.7,
                        topK = 40,
                        topP = 0.95,
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
                    if (result["candidates"] != null && result["candidates"].Count() > 0)
                    {
                        return result["candidates"][0]["content"]["parts"][0]["text"].ToString();
                    }
                    else
                    {
                        throw new Exception("Gemini API 未返回有效回應");
                    }
                }
                else
                {
                    // 解析錯誤訊息
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
        /// 取得 Gemini API Key（從 appsettings.json 或 user-secrets）
        /// </summary>
        private string GetGeminiApiKey()
        {
            // 從 FunConfig 讀取（支援 appsettings.json 和 user-secrets）
            // user-secrets 設定方式：
            // dotnet user-secrets set "FunConfig:GeminiApiKey" "your-api-key"
            var apiKey = _Fun.Config.GeminiApiKey;

            // 如果沒有，可以從環境變數讀取（備用方案）
            if (string.IsNullOrEmpty(apiKey))
            {
                apiKey = Environment.GetEnvironmentVariable("GEMINI_API_KEY");
            }

            return apiKey;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _httpClient?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
