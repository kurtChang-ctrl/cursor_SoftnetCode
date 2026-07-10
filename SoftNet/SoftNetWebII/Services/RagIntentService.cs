using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SoftNetWebII.Services
{
    /// <summary>
    /// RAG 問題理解服務
    /// 解析用戶問題，識別查詢意圖、實體和資料類型
    /// </summary>
    public class RagIntentService
    {
        /// <summary>
        /// 查詢意圖
        /// </summary>
        public class QueryIntent
        {
            public string OriginalQuestion { get; set; }
            public List<Entity> Entities { get; set; } = new List<Entity>();
            public List<string> DataTypes { get; set; } = new List<string>();
            public TimeRange? TimeRange { get; set; }
        }

        /// <summary>
        /// 實體（料號、工單號等）
        /// </summary>
        public class Entity
        {
            public string Type { get; set; } // PartNO, OrderNO, StationNO 等
            public string Value { get; set; }
        }

        /// <summary>
        /// 時間範圍
        /// </summary>
        public class TimeRange
        {
            public DateTime? StartDate { get; set; }
            public DateTime? EndDate { get; set; }
            public int? Days { get; set; } // 最近 N 天

            public static TimeRange Today => new TimeRange
            {
                StartDate = DateTime.Today,
                EndDate = DateTime.Today.AddDays(1).AddSeconds(-1)
            };

            public static TimeRange Yesterday => new TimeRange
            {
                StartDate = DateTime.Today.AddDays(-1),
                EndDate = DateTime.Today.AddSeconds(-1)
            };

            public static TimeRange LastDays(int days) => new TimeRange
            {
                StartDate = DateTime.Now.AddDays(-days),
                EndDate = DateTime.Now,
                Days = days
            };
        }

        /// <summary>
        /// 解析用戶問題，識別查詢意圖
        /// </summary>
        public QueryIntent ParseIntent(string question)
        {
            if (string.IsNullOrWhiteSpace(question))
            {
                return new QueryIntent { OriginalQuestion = question };
            }

            var intent = new QueryIntent { OriginalQuestion = question };

            // 1. 提取時間範圍
            intent.TimeRange = ExtractTimeRange(question);

            // 2. 提取實體（料號、工單號、站別、NeedId、SimulationId等）
            intent.Entities.AddRange(ExtractPartNumbers(question));
            intent.Entities.AddRange(ExtractWorkOrderNumbers(question));
            intent.Entities.AddRange(ExtractStationNumbers(question));
            intent.Entities.AddRange(ExtractNeedIds(question));
            intent.Entities.AddRange(ExtractSimulationIds(question));

            // 3. 識別資料類型
            intent.DataTypes.AddRange(ExtractDataTypes(question));

            return intent;
        }

        /// <summary>
        /// 提取時間範圍
        /// </summary>
        private TimeRange ExtractTimeRange(string question)
        {
            var lowerQuestion = question.ToLower();

            if (lowerQuestion.Contains("今天") || lowerQuestion.Contains("本日") || lowerQuestion.Contains("今日"))
            {
                return TimeRange.Today;
            }

            if (lowerQuestion.Contains("昨天") || lowerQuestion.Contains("昨日"))
            {
                return TimeRange.Yesterday;
            }

            if (lowerQuestion.Contains("最近") || lowerQuestion.Contains("前"))
            {
                // 提取天數：最近7天、前30天等
                var dayMatch = Regex.Match(question, @"(最近|前)\s*(\d+)\s*天");
                if (dayMatch.Success && int.TryParse(dayMatch.Groups[2].Value, out int days))
                {
                    return TimeRange.LastDays(days);
                }
                // 預設最近7天
                return TimeRange.LastDays(7);
            }

            // 提取具體日期範圍
            var datePattern = @"(\d{4}[-/]\d{1,2}[-/]\d{1,2})";
            var dates = Regex.Matches(question, datePattern);
            if (dates.Count >= 2)
            {
                if (DateTime.TryParse(dates[0].Value, out DateTime startDate) &&
                    DateTime.TryParse(dates[1].Value, out DateTime endDate))
                {
                    return new TimeRange { StartDate = startDate, EndDate = endDate };
                }
            }
            else if (dates.Count == 1)
            {
                if (DateTime.TryParse(dates[0].Value, out DateTime date))
                {
                    return new TimeRange
                    {
                        StartDate = date.Date,
                        EndDate = date.Date.AddDays(1).AddSeconds(-1)
                    };
                }
            }

            return null;
        }

        /// <summary>
        /// 提取料號
        /// </summary>
        private List<Entity> ExtractPartNumbers(string question)
        {
            var entities = new List<Entity>();

            // 模式：P001、P-001、料號P001、PartNO: P001 等
            var patterns = new[]
            {
                @"(?:料號|品名|規格|PartNO|Part\s*NO)[:：]?\s*([A-Z0-9\-]+)",
                @"\b([A-Z]\d{3,})\b", // P001, A123 等
                @"\b([A-Z]-\d{3,})\b" // P-001 等
            };

            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(question, pattern, RegexOptions.IgnoreCase);
                foreach (Match match in matches)
                {
                    entities.Add(new Entity
                    {
                        Type = "PartNO",
                        Value = match.Groups[1].Value.ToUpper()
                    });
                }
            }

            return entities.Distinct().ToList();
        }

        /// <summary>
        /// 提取工單號
        /// </summary>
        private List<Entity> ExtractWorkOrderNumbers(string question)
        {
            var entities = new List<Entity>();

            // 模式：WO001、WO-001、工單WO001、OrderNO: WO001 等
            var patterns = new[]
            {
                @"(?:工單|WorkOrder|OrderNO|Order\s*NO|WorkOrder)[:：]?\s*([A-Z0-9\-]+)",
                @"\b(WO\d{3,})\b", // WO001 等
                @"\b(XX\d{2,})\b", // XX01 等
                @"\b(WO-\d{3,})\b" // WO-001 等
            };

            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(question, pattern, RegexOptions.IgnoreCase);
                foreach (Match match in matches)
                {
                    entities.Add(new Entity
                    {
                        Type = "OrderNO",
                        Value = match.Groups[1].Value.ToUpper()
                    });
                }
            }

            return entities.Distinct().ToList();
        }

        /// <summary>
        /// 提取站別編號
        /// </summary>
        private List<Entity> ExtractStationNumbers(string question)
        {
            var entities = new List<Entity>();

            // 模式：ST001、ST-001、站別ST001、StationNO: ST001 等
            var patterns = new[]
            {
                @"(?:站別|StationNO|Station\s*NO)[:：]?\s*([A-Z0-9\-]+)",
                @"\b(ST\d{3,})\b", // ST001 等
                @"\b(ST-\d{3,})\b" // ST-001 等
            };

            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(question, pattern, RegexOptions.IgnoreCase);
                foreach (Match match in matches)
                {
                    entities.Add(new Entity
                    {
                        Type = "StationNO",
                        Value = match.Groups[1].Value.ToUpper()
                    });
                }
            }

            return entities.Distinct().ToList();
        }
        private List<Entity> ExtractNeedIds(string question)
        {
            var entities = new List<Entity>();

            // 模式：ST001、ST-001、站別ST001、StationNO: ST001 等
            var patterns = new[]
            {
                @"(?:APS|排程|模擬|計畫|Station\s*NO)[:：]?\s*([A-Z0-9\-]+)",
                @"\b(ST\d{3,})\b", // ST001 等
                @"\b(ST-\d{3,})\b" // ST-001 等
            };

            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(question, pattern, RegexOptions.IgnoreCase);
                foreach (Match match in matches)
                {
                    entities.Add(new Entity
                    {
                        Type = "NeedId",
                        Value = match.Groups[1].Value.ToUpper()
                    });
                }
            }

            return entities.Distinct().ToList();
        }
        private List<Entity> ExtractSimulationIds(string question)
        {
            var entities = new List<Entity>();

            // 模式：ST001、ST-001、站別ST001、StationNO: ST001 等
            var patterns = new[]
            {
                @"(?:子排程|模擬\s*NO)[:：]?\s*([A-Z0-9\-]+)",
                @"\b(ST\d{3,})\b", // ST001 等
                @"\b(ST-\d{3,})\b" // ST-001 等
            };

            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(question, pattern, RegexOptions.IgnoreCase);
                foreach (Match match in matches)
                {
                    entities.Add(new Entity
                    {
                        Type = "SimulationId",
                        Value = match.Groups[1].Value.ToUpper()
                    });
                }
            }

            return entities.Distinct().ToList();
        }

        /// <summary>
        /// 識別資料類型
        /// </summary>
        private List<string> ExtractDataTypes(string question)
        {
            var dataTypes = new List<string>();
            var lowerQuestion = question.ToLower();

            // 生產相關關鍵字
            if (lowerQuestion.Contains("生產") || lowerQuestion.Contains("產量") ||
                lowerQuestion.Contains("製造") || lowerQuestion.Contains("報工") ||
                lowerQuestion.Contains("製程") || lowerQuestion.Contains("瓶頸") ||
                lowerQuestion.Contains("稼動") || lowerQuestion.Contains("流程") ||
                lowerQuestion.Contains("CT") || lowerQuestion.Contains("UPH") || lowerQuestion.Contains("UPH") ||
                lowerQuestion.Contains("工站") || lowerQuestion.Contains("站別"))
            {
                dataTypes.Add("Production");
            }

            // APS相關關鍵字
            if (lowerQuestion.Contains("APS") || lowerQuestion.Contains("時程") || lowerQuestion.Contains("出貨") ||
                lowerQuestion.Contains("延") || lowerQuestion.Contains("期") || lowerQuestion.Contains("負載") ||
                lowerQuestion.Contains("模擬") || lowerQuestion.Contains("插單") || lowerQuestion.Contains("重排") ||
                lowerQuestion.Contains("進貨") || lowerQuestion.Contains("交貨") ||
                lowerQuestion.Contains("計畫") || lowerQuestion.Contains("排程"))
            {
                dataTypes.Add("APS");
            }

            // 庫存相關關鍵字
            if (lowerQuestion.Contains("庫") || lowerQuestion.Contains("存貨") ||
                lowerQuestion.Contains("倉") || lowerQuestion.Contains("儲位"))
            {
                dataTypes.Add("Inventory");
            }

            // 工單相關關鍵字
            if (lowerQuestion.Contains("工單") || lowerQuestion.Contains("訂單") ||
                lowerQuestion.Contains("拆單") || lowerQuestion.Contains("插單") ||
                lowerQuestion.Contains("workorder") || lowerQuestion.Contains("order"))
            {
                dataTypes.Add("WorkOrder");
            }

            // 物料相關關鍵字
            if (lowerQuestion.Contains("料") || lowerQuestion.Contains("單據") ||
                lowerQuestion.Contains("品名") || lowerQuestion.Contains("規格") ||
                lowerQuestion.Contains("類型") || lowerQuestion.Contains("material"))
            {
                dataTypes.Add("Material");
            }

            // 如果沒有明確的資料類型，預設包含生產和工單
            //if (dataTypes.Count == 0)
            //{
                dataTypes.Add("Exception");
                dataTypes.Add("Intervention");
            //}

            return dataTypes.Distinct().ToList();
        }
    }
}
