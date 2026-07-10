using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SoftNetWebII.Services
{
    /// <summary>
    /// RAG 上下文組裝服務
    /// 將檢索到的資料組裝為 MCP 格式的上下文
    /// </summary>
    public class RagContextService
    {
        /// <summary>
        /// 將檢索到的資料組裝為 MCP 格式的上下文
        /// </summary>
        public string BuildMCPContext(List<RagRetrievalService.RetrievedData> retrievedData, string question, int maxLength = 20000)
        {
            var context = new StringBuilder();

            // 1. 用戶問題
            context.AppendLine("## 用戶問題");
            context.AppendLine(question);
            context.AppendLine();

            // 2. 相關資料上下文
            if (retrievedData == null || retrievedData.Count == 0)
            {
                context.AppendLine("## 相關資料上下文");
                context.AppendLine("未找到相關資料。");
                return context.ToString();
            }

            context.AppendLine("## 相關資料上下文");
            context.AppendLine($"共找到 {retrievedData.Count} 筆相關資料：");
            context.AppendLine();

            // 按資料類型分組（優先排序：風險分數高的在前）
            var grouped = retrievedData
                .OrderByDescending(d =>
                {
                    // 優先排序：風險分數 * 0.7 + 相關性分數 * 0.3
                    if (d.Metadata.ContainsKey("RiskScore"))
                    {
                        var riskScore = Convert.ToDouble(d.Metadata["RiskScore"]);
                        return riskScore * 0.7 + d.RelevanceScore * 0.3;
                    }
                    return d.RelevanceScore;
                })
                .GroupBy(d => d.DataType);

            foreach (var group in grouped)
            {
                context.AppendLine($"### {GetDataTypeName(group.Key)} ({group.Count()} 筆)");

                int index = 1;
                // 優先顯示高風險項目，每種類型最多顯示 20 筆
                var sortedGroup = group
                    .OrderByDescending(d =>
                    {
                        if (d.Metadata.ContainsKey("RiskScore"))
                        {
                            return Convert.ToDouble(d.Metadata["RiskScore"]);
                        }
                        return d.RelevanceScore;
                    })
                    .Take(20)
                    .ToList();

                foreach (var data in sortedGroup)
                {
                    // 加入風險等級標記
                    var riskMark = "";
                    if (data.Metadata.ContainsKey("AnomalyType"))
                    {
                        var anomalyType = data.Metadata["AnomalyType"]?.ToString() ?? "";
                        var riskScore = data.Metadata.ContainsKey("RiskScore")
                            ? Convert.ToDouble(data.Metadata["RiskScore"])
                            : 0.0;

                        riskMark = GetRiskMark(anomalyType, riskScore);
                    }

                    // 截斷過長的文字（每筆最多 200 字元）
                    var displayText = data.Text;
                    if (displayText.Length > 200)
                    {
                        displayText = displayText.Substring(0, 197) + "...";
                    }

                    context.AppendLine($"{index}. {riskMark}{displayText}");
                    index++;
                }

                context.AppendLine();
            }

            // 3. 風險摘要
            var highRiskCount = retrievedData.Count(r =>
                r.Metadata.ContainsKey("RiskScore") &&
                Convert.ToDouble(r.Metadata["RiskScore"]) >= 0.7);

            if (highRiskCount > 0)
            {
                context.AppendLine("## ⚠️ 風險摘要");
                context.AppendLine($"發現 {highRiskCount} 筆高風險異常（風險分數 ≥ 0.7），建議優先處理。");
                context.AppendLine();
            }

            // 4. 限制總長度
            var result = context.ToString();
            if (result.Length > maxLength)
            {
                result = result.Substring(0, maxLength - 3) + "...";
            }

            return result;
        }

        /// <summary>
        /// 取得資料類型的中文名稱
        /// </summary>
        private string GetDataTypeName(string dataType)
        {
            return dataType switch
            {
                "Production" => "生產資料",
                "Inventory" => "庫存資料",
                "WorkOrder" => "工單資料",
                "Material" => "物料資料",
                "Exception" => "異常資料",
                "Intervention" => "干涉資料",
                "APS" => "排程資料",
                _ => dataType
            };
        }

        /// <summary>
        /// 取得風險等級標記
        /// </summary>
        private string GetRiskMark(string anomalyType, double riskScore)
        {
            if (riskScore >= 0.8)
                return "🔴 [嚴重] ";
            else if (riskScore >= 0.6)
                return "🟠 [高風險] ";
            else if (riskScore >= 0.4)
                return "🟡 [中風險] ";
            else if (riskScore >= 0.2)
                return "🟢 [低風險] ";
            else
                return "";
        }

        /// <summary>
        /// 建立簡化的上下文（用於 token 限制較嚴格的情況）
        /// </summary>
        public string BuildSimplifiedContext(List<RagRetrievalService.RetrievedData> retrievedData, string question, int maxItems = 10)
        {
            var context = new StringBuilder();

            context.AppendLine("問題：" + question);
            context.AppendLine();
            context.AppendLine("相關資料：");

            var topData = retrievedData
                .OrderByDescending(d => d.RelevanceScore)
                .Take(maxItems)
                .ToList();

            int index = 1;
            foreach (var data in topData)
            {
                context.AppendLine($"{index}. [{data.DataType}] {data.Text}");
                index++;
            }

            return context.ToString();
        }
    }
}
