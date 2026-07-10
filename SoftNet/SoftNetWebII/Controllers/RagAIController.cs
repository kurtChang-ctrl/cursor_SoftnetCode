using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    /// <summary>
    /// RAG AI Controller
    /// 提供智能問答功能，支援 Gemini 和 OpenAI API
    /// </summary>
    public class RagAIController : ApiCtrl
    {
        private readonly RagQueryService _queryService;

        public RagAIController()
        {
            _queryService = new RagQueryService();
        }

        /// <summary>
        /// 顯示 RAG AI 問答頁面
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
        /// 智能問答 API
        /// </summary>
        /// <param name="question">用戶問題</param>
        /// <param name="provider">LLM 提供者（gemini 或 openai）</param>
        /// <param name="enableVectorSearch">是否啟用向量搜尋</param>
        /// <returns>回答結果</returns>
        [HttpPost]
        public async Task<ContentResult> Ask(string question, string provider = "gemini", bool enableVectorSearch = true)
        {
            var stopwatch = Stopwatch.StartNew();
            var user = _Fun.GetBaseUser();

            try
            {
                if (string.IsNullOrWhiteSpace(question))
                {
                    return JsonToCnt(JObject.FromObject(new
                    {
                        success = false,
                        error = "請輸入問題"
                    }));
                }

                // 共用 RAG 問答流程：意圖 → 檢索 → 上下文 → LLM
                var (answer, retrievedData) = await _queryService.AskAsync(question, null, provider, enableVectorSearch);

                stopwatch.Stop();

                // 5. 記錄查詢（背景執行，不阻塞回應）
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await LogQuery(user?.UserNO ?? "", question, provider, retrievedData.Count, answer,
                            (int)stopwatch.ElapsedMilliseconds, true, null);
                    }
                    catch { }
                });

                // 6. 計算風險統計
                var highRiskData = retrievedData
                    .Where(d => d.Metadata.ContainsKey("RiskScore") &&
                                Convert.ToDouble(d.Metadata["RiskScore"]) >= 0.7)
                    .Take(10)
                    .Select(d => new
                    {
                        dataType = d.DataType,
                        text = d.Text,
                        relevanceScore = d.RelevanceScore,
                        riskScore = d.Metadata.ContainsKey("RiskScore") ? d.Metadata["RiskScore"] : 0.0,
                        anomalyType = d.Metadata.ContainsKey("AnomalyType") ? d.Metadata["AnomalyType"] : ""
                    })
                    .ToList();

                // 7. 回傳結果
                return JsonToCnt(JObject.FromObject(new
                {
                    success = true,
                    question = question,
                    answer = answer,
                    retrievedDataCount = retrievedData.Count,
                    highRiskCount = highRiskData.Count,
                    retrievedData = retrievedData.Take(10).Select(d => new
                    {
                        dataType = d.DataType,
                        text = d.Text,
                        relevanceScore = d.RelevanceScore,
                        riskScore = d.Metadata.ContainsKey("RiskScore") ? d.Metadata["RiskScore"] : (double?)null,
                        anomalyType = d.Metadata.ContainsKey("AnomalyType") ? d.Metadata["AnomalyType"] : null
                    }).ToList(),
                    highRiskData = highRiskData,
                    provider = provider,
                    responseTime = stopwatch.ElapsedMilliseconds
                }));
            }
            catch (Exception ex)
            {
                stopwatch.Stop();

                await _Log.ErrorAsync($"RagAIController.Ask 錯誤: {ex.Message} {ex.StackTrace}", true);

                // 記錄錯誤查詢
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await LogQuery(user?.UserNO ?? "", question, provider, 0, null,
                            (int)stopwatch.ElapsedMilliseconds, false, ex.Message);
                    }
                    catch { }
                });

                return JsonToCnt(JObject.FromObject(new
                {
                    success = false,
                    error = ex.Message
                }));
            }
        }

        /// <summary>
        /// 記錄查詢到資料庫
        /// </summary>
        private async Task LogQuery(string userNO, string question, string provider, int retrievedDataCount,
            string answer, int responseTime, bool success, string error)
        {
            try
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    string sql = @"
                        INSERT INTO SoftNetLogDB.[dbo].[RAG_QueryLog]
                        ([ServerId], [UserNO], [Question], [Provider], [RetrievedDataCount], 
                         [Answer], [ResponseTime], [Success], [Error], [CreateTime])
                        VALUES (@ServerId, @UserNO, @Question, @Provider, @RetrievedDataCount,
                                @Answer, @ResponseTime, @Success, @Error, GETDATE())
                    ";

                    var args = new List<object>
                    {
                        "ServerId", _Fun.Config.ServerId,
                        "UserNO", userNO,
                        "Question", question,
                        "Provider", provider,
                        "RetrievedDataCount", retrievedDataCount,
                        "Answer", answer ?? "",
                        "ResponseTime", responseTime,
                        "Success", success ? 1 : 0,
                        "Error", error ?? ""
                    };

                    db.DB_SetDataByParams(sql, args);
                }
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"LogQuery 錯誤: {ex.Message}", true);
            }
        }

        /// <summary>
        /// 取得查詢歷史記錄
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetQueryHistory(int page = 1, int pageSize = 20)
        {
            try
            {
                var user = _Fun.GetBaseUser();
                if (user == null || !user.IsLogin)
                {
                    return JsonToCnt(JObject.FromObject(new
                    {
                        success = false,
                        error = "請先登入"
                    }));
                }

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    string sql = @"
                        SELECT 
                            [Id],
                            [Question],
                            [Provider],
                            [RetrievedDataCount],
                            [ResponseTime],
                            [Success],
                            [CreateTime]
                        FROM SoftNetLogDB.[dbo].[RAG_QueryLog]
                        WHERE [ServerId] = @ServerId
                            AND [UserNO] = @UserNO
                        ORDER BY [CreateTime] DESC
                        OFFSET @Offset ROWS
                        FETCH NEXT @PageSize ROWS ONLY
                    ";

                    var offset = (page - 1) * pageSize;
                    var args = new List<object>
                    {
                        "ServerId", _Fun.Config.ServerId,
                        "UserNO", user.UserNO,
                        "Offset", offset,
                        "PageSize", pageSize
                    };

                    var table = db.DB_GetDataByDataTable(sql, args);
                    var records = new List<object>();

                    foreach (System.Data.DataRow row in table.Rows)
                    {
                        records.Add(new
                        {
                            id = row["Id"],
                            question = row["Question"],
                            provider = row["Provider"],
                            retrievedDataCount = row["RetrievedDataCount"],
                            responseTime = row["ResponseTime"],
                            success = row["Success"],
                            createTime = row["CreateTime"]
                        });
                    }

                    return JsonToCnt(JObject.FromObject(new
                    {
                        success = true,
                        records = records,
                        page = page,
                        pageSize = pageSize
                    }));
                }
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"GetQueryHistory 錯誤: {ex.Message}", true);
                return JsonToCnt(JObject.FromObject(new
                {
                    success = false,
                    error = ex.Message
                }));
            }
        }
    }
}
