using Base.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    /// <summary>
    /// RAG 異常檢測服務
    /// 使用規則基礎的分類方法，根據 ErrorType、EventType 和文字內容進行風險評分
    /// </summary>
    public class RagAnomalyDetectionService
    {
        /// <summary>
        /// 異常類型分類
        /// </summary>
        public enum AnomalyType
        {
            Normal = 0,           // 正常
            LowRisk = 1,          // 低風險
            MediumRisk = 2,       // 中風險
            HighRisk = 3,         // 高風險
            Critical = 4          // 嚴重異常
        }

        /// <summary>
        /// 異常分類結果
        /// </summary>
        public class AnomalyClassificationResult
        {
            public AnomalyType AnomalyType { get; set; }
            public float Confidence { get; set; }
            public float RiskScore { get; set; } // 0-1，越高越危險
            public bool IsRuleBased { get; set; } = false;
            public string ErrorType { get; set; }
            public string EventType { get; set; }
        }

        /// <summary>
        /// 根據 MES 資料進行異常分類
        /// </summary>
        public async Task<AnomalyClassificationResult> ClassifyAnomaly(
            Dictionary<string, object> metadata,
            string text)
        {
            var result = new AnomalyClassificationResult();

            // 1. 優先使用 ErrorType（最準確）
            if (metadata.ContainsKey("ErrorType"))
            {
                var errorType = metadata["ErrorType"]?.ToString() ?? "";
                if (!string.IsNullOrEmpty(errorType))
                {
                    result = ClassifyByErrorType(errorType, metadata);
                    if (result.Confidence > 0.7)
                    {
                        return result;
                    }
                }
            }

            // 2. 使用 EventType
            if (metadata.ContainsKey("EventType"))
            {
                var eventType = metadata["EventType"]?.ToString() ?? "";
                if (!string.IsNullOrEmpty(eventType))
                {
                    result = ClassifyByEventType(eventType, metadata);
                    if (result.Confidence > 0.7)
                    {
                        return result;
                    }
                }
            }

            // 3. 使用文字內容（關鍵字匹配）
            result = ClassifyByText(text);

            return result;
        }

        /// <summary>
        /// 根據 ErrorType 分類（基於現有編碼規則）
        /// </summary>
        private AnomalyClassificationResult ClassifyByErrorType(
            string errorType,
            Dictionary<string, object> metadata)
        {
            // 基於現有的 ErrorType 編碼進行風險評分
            var riskMapping = new Dictionary<string, (AnomalyType type, float risk)>
            {
                // Critical（嚴重）
                
                { "18", (AnomalyType.Critical, 0.90f) },  // 工站應停止未停止
                { "22", (AnomalyType.Critical, 0.90f) },  // 疑似重複報工
                { "23", (AnomalyType.Critical, 0.91f) },  // 疑似委外加工數量多出
                { "24", (AnomalyType.Critical, 0.91f) },  // 工單關閉後才補開委外加工單
                { "25", (AnomalyType.Critical, 0.92f) },  // 系統干涉補報工量
                { "15", (AnomalyType.MediumRisk, 0.93f) },  // 報工延遲
                { "14", (AnomalyType.MediumRisk, 0.94f) },  // 實際CT遠離目標CT


                // High Risk（高風險）
                { "04", (AnomalyType.HighRisk, 0.70f) },  // 應領用而未領用
                { "05", (AnomalyType.HighRisk, 0.70f) },  // 無單據發起
                { "19", (AnomalyType.HighRisk, 0.85f) },  // 排程計畫已偏離需求日
                { "30", (AnomalyType.MediumRisk, 0.85f) },  // 報工時間遠超計畫工時


                // Medium Risk（中風險）
                { "01", (AnomalyType.HighRisk, 0.50f) },  // 庫存不足
                { "02", (AnomalyType.HighRisk, 0.50f) },  // 庫存不足
                { "12", (AnomalyType.Critical, 0.60f) },  // 工單應關未關
                { "03", (AnomalyType.MediumRisk, 0.60f) },  // 單據未如預期開工
                { "17", (AnomalyType.MediumRisk, 0.60f) },  // 移轉量不合理
                { "28", (AnomalyType.MediumRisk, 0.60f) },  // 計畫日已過報工明顯不足

                // Low Risk（低風險）
                { "21", (AnomalyType.LowRisk, 0.30f) },  // 工單強制關閉檢查
                { "29", (AnomalyType.MediumRisk, 0.30f) },  // 工站移轉或在制數量為負


                // Low Risk（很低風險）
                { "31", (AnomalyType.LowRisk, 0.10f) },  // 領料或入庫數量為負
                { "26", (AnomalyType.LowRisk, 0.10f) },  // 領入庫單據時間晚於計畫逾7天
                { "27", (AnomalyType.LowRisk, 0.10f) },  // 領入庫單據時間早於計畫逾7天
            };

            if (riskMapping.ContainsKey(errorType))
            {
                var (type, risk) = riskMapping[errorType];
                return new AnomalyClassificationResult
                {
                    AnomalyType = type,
                    RiskScore = risk,
                    Confidence = 0.95f, // 規則基礎的置信度很高
                    IsRuleBased = true,
                    ErrorType = errorType
                };
            }

            // 未知 ErrorType，使用預設值
            return new AnomalyClassificationResult
            {
                AnomalyType = AnomalyType.MediumRisk,
                RiskScore = 0.50f,
                Confidence = 0.60f,
                IsRuleBased = true,
                ErrorType = errorType
            };
        }

        /// <summary>
        /// 根據 EventType 分類
        /// </summary>
        private AnomalyClassificationResult ClassifyByEventType(
            string eventType,
            Dictionary<string, object> metadata)
        {
            // EventType: 99=系統事件紀錄, WW=警告, EE=異常, II=干涉
            var result = new AnomalyClassificationResult { IsRuleBased = true };

            if (eventType.StartsWith("EE", StringComparison.OrdinalIgnoreCase)) // Ex=異常
            {
                result.AnomalyType = AnomalyType.HighRisk;
                result.RiskScore = 0.75f;
                result.Confidence = 0.85f;
            }
            else if (eventType.StartsWith("WW", StringComparison.OrdinalIgnoreCase)) // Wx=警告
            {
                result.AnomalyType = AnomalyType.MediumRisk;
                result.RiskScore = 0.50f;
                result.Confidence = 0.80f;
            }
            else if (eventType.Equals("II", StringComparison.OrdinalIgnoreCase)) // 干涉
            {
                result.AnomalyType = AnomalyType.LowRisk;
                result.RiskScore = 0.30f;
                result.Confidence = 0.75f;
            }
            //else // 99=系統事件 或其他
            //{
            //    result.AnomalyType = AnomalyType.Normal;
            //    result.RiskScore = 0.10f;
            //    result.Confidence = 0.70f;
            //}

            result.EventType = eventType;
            return result;
        }

        /// <summary>
        /// 根據文字內容分類（關鍵字匹配）
        /// </summary>
        private AnomalyClassificationResult ClassifyByText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return new AnomalyClassificationResult
                {
                    AnomalyType = AnomalyType.Normal,
                    RiskScore = 0.10f,
                    Confidence = 0.50f,
                    IsRuleBased = true
                };
            }

            var lowerText = text.ToLower();
            var result = new AnomalyClassificationResult { IsRuleBased = true };

            // 嚴重關鍵字
            if (lowerText.Contains("嚴重") || lowerText.Contains("critical") ||
                lowerText.Contains("停機") || lowerText.Contains("故障") ||
                lowerText.Contains("應關未關") || lowerText.Contains("應停止未停止") ||
                lowerText.Contains("應關而未關"))
            {
                result.AnomalyType = AnomalyType.Critical;
                result.RiskScore = 0.90f;
                result.Confidence = 0.70f;
            }
            // 高風險關鍵字
            else if (lowerText.Contains("庫存不足") || lowerText.Contains("應領用而未領用") ||
                     lowerText.Contains("無單據") || lowerText.Contains("偏離需求") ||
                     lowerText.Contains("異常") || lowerText.Contains("error"))
            {
                result.AnomalyType = AnomalyType.HighRisk;
                result.RiskScore = 0.75f;
                result.Confidence = 0.65f;
            }
            // 中風險關鍵字
            else if (lowerText.Contains("延遲") || lowerText.Contains("未如預期") ||
                     lowerText.Contains("遠離目標") || lowerText.Contains("不合理") ||
                     lowerText.Contains("警告") || lowerText.Contains("warning"))
            {
                result.AnomalyType = AnomalyType.MediumRisk;
                result.RiskScore = 0.50f;
                result.Confidence = 0.60f;
            }
            // 低風險關鍵字
            else if (lowerText.Contains("注意") || lowerText.Contains("注意事項") ||
                     lowerText.Contains("干涉"))
            {
                result.AnomalyType = AnomalyType.LowRisk;
                result.RiskScore = 0.30f;
                result.Confidence = 0.55f;
            }
            else
            {
                result.AnomalyType = AnomalyType.Normal;
                result.RiskScore = 0.10f;
                result.Confidence = 0.50f;
            }

            return result;
        }
    }
}
