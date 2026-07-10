using Base;
using Base.Models;
using Base.Services;
using DocumentFormat.OpenXml.Drawing.Charts;
using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Diagnostics;

namespace SoftNetWebII.Services
{
    /// <summary>
    /// RAG SQL 檢索服務
    /// 根據查詢意圖，使用 SQL 查詢檢索相關資料
    /// </summary>
    public class RagRetrievalService
    {
        /// <summary>
        /// 檢索到的資料
        /// </summary>
        public class RetrievedData
        {
            public string DataId { get; set; }
            public string DataType { get; set; }
            public string Text { get; set; }
            public double RelevanceScore { get; set; } // 0-1，越高越相關
            public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();
        }

        /// <summary>
        /// 根據查詢意圖檢索相關資料
        /// </summary>
        public async Task<List<RetrievedData>> RetrieveData(RagIntentService.QueryIntent intent)
        {
            var results = new List<RetrievedData>();
            DBADO db = new DBADO("1", _Fun.Config.Db);
            // 根據資料類型檢索
            foreach (var dataType in intent.DataTypes)
            {
                List<RetrievedData> data = null;

                switch (dataType)
                {
                    case "Production":
                        data = await RetrieveProductionData(intent,db);
                        break;
                    case "Inventory":
                        data = await RetrieveInventoryData(intent, db);
                        break;
                    case "WorkOrder":
                        data = await RetrieveWorkOrderData(intent, db);
                        break;
                    case "Material":
                        data = await RetrieveMaterialData(intent, db);
                        break;
                    case "APS":
                        data = await RetrieveAPSData(intent, db);
                        break;
                    case "Intervention":
                        data = await RetrieveInterventionData(intent, db);
                        break;
                    case "Exception":
                        data = await RetrieveExceptionData(intent, db);
                        break;
                }

                if (data != null && data.Count > 0)
                {
                    results.AddRange(data);
                }
            }
            db.Dispose();
            // 按相關性排序
            return results.OrderByDescending(r => r.RelevanceScore).Take(3000).ToList();
        }

        /// <summary>
        /// Production 檢索生產資料
        /// </summary>
        private async Task<List<RetrievedData>> RetrieveProductionData(RagIntentService.QueryIntent intent, DBADO db)
        {
            var results = new List<RetrievedData>();

            var whereConditions = new List<string>();
            var args = new List<object> { "ServerId", _Fun.Config.ServerId };
            var sql = "";
            var tmp = "";



            // 時間範圍
            if (intent.TimeRange != null)
            {
                if (intent.TimeRange.StartDate.HasValue)
                {
                    whereConditions.Add("m.StartTime >= @StartDate");
                    args.Add("StartDate");
                    args.Add(intent.TimeRange.StartDate.Value);
                }
                if (intent.TimeRange.EndDate.HasValue)
                {
                    whereConditions.Add("m.StartTime <= @EndDate");
                    args.Add("EndDate");
                    args.Add(intent.TimeRange.EndDate.Value);
                }
            }
            else
            {
                // 預設最近7天
                whereConditions.Add("m.StartTime >= @StartDate");
                args.Add("StartDate");
                args.Add(DateTime.Now.AddDays(-7));
            }


            #region 製造,生產
            tmp = "";
            if (intent.TimeRange != null && intent.TimeRange.StartDate.HasValue)
            { tmp = $" and a.LOGDateTime>={intent.TimeRange.StartDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}"; }
            if (intent.TimeRange != null && intent.TimeRange.EndDate.HasValue && tmp != "")
            { tmp = $" and a.LOGDateTime<={intent.TimeRange.EndDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}"; }

            sql = $@"SELECT TOP 1000 a.ServerId,a.StationNO,a.LOGDateTime,a.NeedId,a.SimulationId,a.OperateType,a.PartNO,b.PartName,b.Specification,a.OrderNO,a.OP_NO from SoftNetLogDB.[dbo].[OperateLog] as a
                            LEFT JOIN SoftNetMainDB.[dbo].[Material] as b ON b.PartNO = a.PartNO AND b.ServerId = a.ServerId
                        WHERE a.ServerId = @ServerId {tmp} order by a.ServerId,a.LOGDateTime";
            var table = db.DB_GetDataByDataTable(sql, args,true);
            foreach (DataRow row in table.Rows)
            {
                var text = $"工單編號: {row["OrderNO"]}, 料件編號: {row["PartNO"]}, 品名: {row["PartName"]}, 規格: {row["Specification"]}, 工站編號: {row["StationNO"]}, 紀錄時間: {row["LOGDateTime"]}, 動作: {row["OperateType"]}, 作業員: {row["OP_NO"]}, 排程編號: {row["NeedId"]}, 子排程編號: {row["SimulationId"]}";

                results.Add(new RetrievedData
                {
                    DataId = row["Id"].ToString(),
                    DataType = "Production",
                    Text = text,
                    RelevanceScore = 1.0, // SQL 檢索預設為完全相關
                    Metadata = new Dictionary<string, object>
                        {
                            { "OrderNO", row["OrderNO"] },
                            { "PartNO", row["PartNO"] },
                            { "StationNO", row["StationNO"] },
                            { "PartName", row["PartName"] },
                            { "Specification", row["Specification"] },
                            { "OperateType", row["OperateType"] },
                            { "OP_NO", row["OP_NO"] },
                            { "NeedId", row["NeedId"] },
                            { "SimulationId", row["SimulationId"] },
                            { "LOGDateTime", row["LOGDateTime"] }
                        }
                });
            }
            #endregion

            // 料號
            var partNOEntities = intent.Entities.Where(e => e.Type == "PartNO").ToList();
            if (partNOEntities.Any())
            {
                var partNOs = string.Join(",", partNOEntities.Select(e => $"'{e.Value}'"));
                whereConditions.Add($"m.PartNO IN ({partNOs})");
            }

            // 工單號
            var orderNOEntities = intent.Entities.Where(e => e.Type == "OrderNO").ToList();
            if (orderNOEntities.Any())
            {
                var orderNOs = string.Join(",", orderNOEntities.Select(e => $"'{e.Value}'"));
                whereConditions.Add($"m.OrderNO IN ({orderNOs})");
            }

            // 站別
            var stationNOEntities = intent.Entities.Where(e => e.Type == "StationNO").ToList();
            if (stationNOEntities.Any())
            {
                var stationNOs = string.Join(",", stationNOEntities.Select(e => $"'{e.Value}'"));
                whereConditions.Add($"m.StationNO IN ({stationNOs})");
            }

            var whereClause = whereConditions.Any() ? "AND " + string.Join(" AND ", whereConditions) : "";

            sql = $@"
                    SELECT TOP 100
                        m.OrderNO,
                        m.PartNO,
                        m.StationNO,
                        m.StartTime,
                        m.EndTime,
                        mat.PartName,
                        mat.Specification
                    FROM SoftNetMainDB.[dbo].[Manufacture] m
                    LEFT JOIN SoftNetMainDB.[dbo].[Material] mat ON m.PartNO = mat.PartNO AND m.ServerId = mat.ServerId
                    WHERE m.ServerId = @ServerId
                        {whereClause}
                    ORDER BY m.StartTime DESC
                ";

            table = db.DB_GetDataByDataTable(sql, args);

            foreach (DataRow row in table.Rows)
            {
                var text = $"工單編號: {row["OrderNO"]}, 料件編號: {row["PartNO"]}, 品名: {row["PartName"]}, " +
                           $"工站編號: {row["StationNO"]}, 開始時間: {row["StartTime"]}, " +
                           $"結束時間: {row["EndTime"]}, 產量: {row["OutputQty"]}";

                results.Add(new RetrievedData
                {
                    DataId = row["Id"].ToString(),
                    DataType = "Production",
                    Text = text,
                    RelevanceScore = 1.0, // SQL 檢索預設為完全相關
                    Metadata = new Dictionary<string, object>
                        {
                            { "OrderNO", row["OrderNO"] },
                            { "PartNO", row["PartNO"] },
                            { "StationNO", row["StationNO"] },
                            { "StartTime", row["StartTime"] },
                            { "EndTime", row["EndTime"] },
                            { "OutputQty", row["OutputQty"] }
                        }
                });
            }


            return results;
        }

        /// <summary>
        /// 檢索庫存資料
        /// </summary>
        private async Task<List<RetrievedData>> RetrieveInventoryData(RagIntentService.QueryIntent intent, DBADO db)
        {
            var results = new List<RetrievedData>();


            var whereConditions = new List<string>();
            var args = new List<object> { "ServerId", _Fun.Config.ServerId };

            // 料號
            var partNOEntities = intent.Entities.Where(e => e.Type == "PartNO").ToList();
            if (partNOEntities.Any())
            {
                var partNOs = string.Join(",", partNOEntities.Select(e => $"'{e.Value}'"));
                whereConditions.Add($"s.PartNO IN ({partNOs})");
            }

            var whereClause = whereConditions.Any() ? "AND " + string.Join(" AND ", whereConditions) : "";

            string sql = $@"
                    SELECT TOP 50
                        s.PartNO,
                        s.StoreNO,
                        s.StoreSpacesNO,
                        s.QTY,
                        s.WeightQty,
                        mat.PartName,
                        mat.Specification
                    FROM SoftNetMainDB.[dbo].[Store] s
                    LEFT JOIN SoftNetMainDB.[dbo].[Material] mat ON s.PartNO = mat.PartNO AND s.ServerId = mat.ServerId
                    WHERE s.ServerId = @ServerId
                        {whereClause}
                    ORDER BY s.QTY DESC
                ";

            var table = db.DB_GetDataByDataTable(sql, args);

            foreach (DataRow row in table.Rows)
            {
                var text = $"料件編號: {row["PartNO"]}, 品名: {row["PartName"]}, " +
                           $"倉庫: {row["StoreNO"]}, 儲位: {row["StoreSpacesNO"]}, " +
                           $"數量: {row["QTY"]}, 重量: {row["WeightQty"]}";

                results.Add(new RetrievedData
                {
                    DataId = $"{row["PartNO"]}_{row["StoreNO"]}_{row["StoreSpacesNO"]}",
                    DataType = "Inventory",
                    Text = text,
                    RelevanceScore = 1.0,
                    Metadata = new Dictionary<string, object>
                        {
                            { "PartNO", row["PartNO"] },
                            { "StoreNO", row["StoreNO"] },
                            { "StoreSpacesNO", row["StoreSpacesNO"] },
                            { "QTY", row["QTY"] },
                            { "WeightQty", row["WeightQty"] }
                        }
                });
            }


            return results;
        }

        /// <summary>
        /// 檢索工單資料
        /// </summary>
        private async Task<List<RetrievedData>> RetrieveWorkOrderData(RagIntentService.QueryIntent intent, DBADO db)
        {
            var results = new List<RetrievedData>();

            var whereConditions = new List<string>();
            var args = new List<object> { "ServerId", _Fun.Config.ServerId };

            // 工單號
            var orderNOEntities = intent.Entities.Where(e => e.Type == "OrderNO").ToList();
            if (orderNOEntities.Any())
            {
                var orderNOs = string.Join(",", orderNOEntities.Select(e => $"'{e.Value}'"));
                whereConditions.Add($"wo.OrderNO IN ({orderNOs})");
            }

            // 料號
            var partNOEntities = intent.Entities.Where(e => e.Type == "PartNO").ToList();
            if (partNOEntities.Any())
            {
                var partNOs = string.Join(",", partNOEntities.Select(e => $"'{e.Value}'"));
                whereConditions.Add($"wo.PartNO IN ({partNOs})");
            }

            // 時間範圍
            if (intent.TimeRange != null)
            {
                if (intent.TimeRange.StartDate.HasValue)
                {
                    whereConditions.Add("wo.StartTime >= @StartDate");
                    args.Add("StartDate");
                    args.Add(intent.TimeRange.StartDate.Value);
                }
                if (intent.TimeRange.EndDate.HasValue)
                {
                    whereConditions.Add("wo.StartTime <= @EndDate");
                    args.Add("EndDate");
                    args.Add(intent.TimeRange.EndDate.Value);
                }
            }

            var whereClause = whereConditions.Any() ? "AND " + string.Join(" AND ", whereConditions) : "";

            string sql = $@"
                    SELECT TOP 50
                        wo.OrderNO,
                        wo.PartNO,
                        wo.StartTime,
                        wo.EndTime,
                        wo.TmpClose,
                        mat.PartName
                    FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] wo
                    LEFT JOIN SoftNetMainDB.[dbo].[Material] mat ON wo.PartNO = mat.PartNO AND wo.ServerId = mat.ServerId
                    WHERE wo.ServerId = @ServerId
                        {whereClause}
                    ORDER BY wo.StartTime DESC
                ";

            var table = db.DB_GetDataByDataTable(sql, args,true);

            foreach (DataRow row in table.Rows)
            {
                var text = $"工單編號: {row["OrderNO"]}, 料件編號: {row["PartNO"]}, 品名: {row["PartName"]}, " +
                           $"開始時間: {row["StartTime"]}, 結束時間: {row["EndTime"]}, " +
                           $"狀態: {(row["TmpClose"] == DBNull.Value ? "進行中" : "已關閉")}";

                results.Add(new RetrievedData
                {
                    DataId = "",
                    DataType = "WorkOrder",
                    Text = text,
                    RelevanceScore = 1.0,
                    Metadata = new Dictionary<string, object>
                        {
                            { "OrderNO", row["OrderNO"] },
                            { "PartNO", row["PartNO"] },
                            { "StartTime", row["StartTime"] },
                            { "EndTime", row["EndTime"] },
                            { "TmpClose", row["TmpClose"] }
                        }
                });
            }


            return results;
        }

        private async Task<List<RetrievedData>> RetrieveAPSData(RagIntentService.QueryIntent intent, DBADO db)
        {
            var results = new List<RetrievedData>();

            return results;
        }
        private async Task<List<RetrievedData>> RetrieveInterventionData(RagIntentService.QueryIntent intent, DBADO db)
        {
            var results = new List<RetrievedData>();

            return results;
        }
        private async Task<List<RetrievedData>> RetrieveExceptionData(RagIntentService.QueryIntent intent, DBADO db)
        {
            SFC_Common _SFC_Common = new SFC_Common("1", _Fun.Config.Db);
            var results = new List<RetrievedData>();

            var whereConditions = new List<string>();
            var args = new List<object> { "ServerId", _Fun.Config.ServerId };
            var tmp = "";
            var sql="";
            var text="";
            var tmp_data = "";
            var start = intent.TimeRange?.StartDate;
            var end = intent.TimeRange?.EndDate;
            if (start.HasValue) { args.Add("StartDate"); args.Add(start.Value); }
            if (end.HasValue) { args.Add("EndDate"); args.Add(end.Value); }


            #region APS_Simulation_ErrorData 異常,干涉
            tmp = "";
            if (start.HasValue) tmp += " AND a.LogDate >= @StartDate";
            if (end.HasValue) tmp += " AND a.LogDate <= @EndDate";
            if (!_Fun.Config.DOCInvoice_Expiry_SendError)
            { tmp = $"{tmp} and a.ErrorType!='06' and a.ErrorType!='07' and a.ErrorType!='08' and a.ErrorType!='09' and a.ErrorType!='10'"; }
            sql = $@"select TOP 100 a.*,b.PartNO from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                    LEFT JOIN SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.SimulationId=b.SimulationId
                    WHERE a.ServerId = @ServerId {tmp} order by a.LogDate desc, ActionType";
            var sw1 = Stopwatch.StartNew();
            var table = db.DB_GetDataByDataTable(sql, args);
            sw1.Stop();
            await _Log.InfoAsync($"RAG RetrieveExceptionData APS_Simulation_ErrorData rows={table?.Rows?.Count ?? 0} ms={sw1.ElapsedMilliseconds}");
            foreach (DataRow row in table.Rows)
            {
                tmp_data = _SFC_Common.Get_APS_Simulation_ErrorData_TErrorType2String(row["ErrorType"].ToString());
                text = $"紀錄時間: {row["LogDate"]}, 原因: {tmp_data}";
                if (row["DOCNumberNO"] != DBNull.Value) { text = $"{text}, 單據編號: {row["DOCNumberNO"]}"; }
                if (row["NeedId"] != DBNull.Value) { text = $"{text}, 排程編號: {row["NeedId"]}"; }
                if (row["SimulationId"] != DBNull.Value) { text = $"{text}, 子排程編號: {row["SimulationId"]}"; }
                if (row["StationNO"] != DBNull.Value) { text = $"{text}, 工站編號: {row["StationNO"]}"; }
                if (row["PartNO"] != DBNull.Value) { text = $"{text}, 料件編號: {row["PartNO"]}"; }

                results.Add(new RetrievedData
                {
                    DataId = row["Id"].ToString(),
                    //DataId = _Str.NewId('0'),
                    DataType = "Exception",
                    Text = text,
                    RelevanceScore = 1.0, // SQL 檢索預設為完全相關
                    Metadata = new Dictionary<string, object>
                        {
                            { "LOGDateTime", row["LogDate"] },
                            { "Causes", tmp_data },
                            { "NeedId", row["NeedId"]== DBNull.Value ? "" : row["NeedId"] },
                            { "SimulationId", row["SimulationId"]== DBNull.Value ? "" : row["SimulationId"] },
                            { "StationNO", row["StationNO"]== DBNull.Value ? "" : row["StationNO"] },
                            { "DOCNumberNO", row["DOCNumberNO"]== DBNull.Value ? "" : row["DOCNumberNO"] },
                            { "PartNO", row["PartNO"]== DBNull.Value ? "" : row["PartNO"] }
                         }
                });
            }
            #endregion

            #region APS_WarningData 異常,干涉
            tmp = "";
            if (start.HasValue) tmp += " AND a.WarningDate >= @StartDate";
            if (end.HasValue) tmp += " AND a.WarningDate <= @EndDate";

            sql = $@"select TOP 100 a.*,b.PartNO from SoftNetSYSDB.[dbo].[APS_WarningData] as a
                    LEFT JOIN SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.SimulationId=b.SimulationId 
                    WHERE a.ErrorType!='00' and a.ErrorType!='33' and a.ServerId = @ServerId {tmp} order by a.WarningDate desc";
            var sw2 = Stopwatch.StartNew();
            table = db.DB_GetDataByDataTable(sql, args);
            sw2.Stop();
            await _Log.InfoAsync($"RAG RetrieveExceptionData APS_WarningData rows={table?.Rows?.Count ?? 0} ms={sw2.ElapsedMilliseconds}");
            foreach (DataRow row in table.Rows)
            {
                tmp_data = _SFC_Common.Get_APS_Simulation_ErrorData_TErrorType2String(row["ErrorType"].ToString());
                text = $"紀錄時間: {row["WarningDate"]}, 原因: {tmp_data}";
                if (row["DOCNumberNO"] != DBNull.Value) { text = $"{text}, 單據編號: {row["DOCNumberNO"]}"; }
                if (row["NeedId"] != DBNull.Value) { text = $"{text}, 排程編號: {row["NeedId"]}"; }
                if (row["SimulationId"] != DBNull.Value) { text = $"{text}, 子排程編號: {row["SimulationId"]}"; }
                if (row["StationNO"] != DBNull.Value) { text = $"{text}, 工站編號: {row["StationNO"]}"; }
                if (row["PartNO"] != DBNull.Value) { text = $"{text}, 料件編號: {row["PartNO"]}"; }

                results.Add(new RetrievedData
                {
                    DataId = row["Id"].ToString(),
                    DataType = "Exception",
                    Text = text,
                    RelevanceScore = 1.0, // SQL 檢索預設為完全相關
                    Metadata = new Dictionary<string, object>
                        {
                            { "LOGDateTime", row["WarningDate"] },
                            { "Causes", tmp_data },
                            { "NeedId", row["NeedId"]== DBNull.Value ? "" : row["NeedId"] },
                            { "SimulationId", row["SimulationId"]== DBNull.Value ? "" : row["SimulationId"] },
                            { "StationNO", row["StationNO"]== DBNull.Value ? "" : row["StationNO"] },
                            { "DOCNumberNO", row["DOCNumberNO"]== DBNull.Value ? "" : row["DOCNumberNO"] },
                            { "PartNO", row["PartNO"]== DBNull.Value ? "" : row["PartNO"] }
                         }
                });
            }
            #endregion

            #region APS_EventLog 異常
            tmp = "";
            if (start.HasValue) tmp += " AND a.LOGDateTime >= @StartDate";
            if (end.HasValue) tmp += " AND a.LOGDateTime <= @EndDate";

            sql = $@"select TOP 100 a.*,b.PartNO,c.DOCNumberNO,c.Source_StationNO from SoftNetLogDB.[dbo].[APS_EventLog] as a
                    LEFT JOIN SoftNetSYSDB.[dbo].[APS_NeedData] as b on a.NeedId=b.Id
                    LEFT JOIN SoftNetSYSDB.[dbo].[APS_Simulation] as c on a.SimulationId=c.SimulationId 
                    WHERE a.ServerId = @ServerId and (a.Event like '%異常%' or a.EventType like 'E%') {tmp} order by a.LOGDateTime desc";
            var sw3 = Stopwatch.StartNew();
            table = db.DB_GetDataByDataTable(sql, args);
            sw3.Stop();
            await _Log.InfoAsync($"RAG RetrieveExceptionData APS_EventLog rows={table?.Rows?.Count ?? 0} ms={sw3.ElapsedMilliseconds}");
            foreach (DataRow row in table.Rows)
            {
                text = $"紀錄時間: {row["LOGDateTime"]}, 原因: {row["Event"]}";
                if (row["DOCNumberNO"] != DBNull.Value) { text = $"{text}, 單據編號: {row["DOCNumberNO"]}"; }
                if (row["NeedId"] != DBNull.Value) { text = $"{text}, 排程編號: {row["NeedId"]}"; }
                if (row["SimulationId"] != DBNull.Value) { text = $"{text}, 子排程編號: {row["SimulationId"]}"; }
                if (row["Source_StationNO"] != DBNull.Value) { text = $"{text}, 工站編號: {row["Source_StationNO"]}"; }
                if (row["PartNO"] != DBNull.Value) { text = $"{text}, 料件編號: {row["PartNO"]}"; }

                results.Add(new RetrievedData
                {
                    DataId = row["Id"].ToString(),
                    //DataId = _Str.NewId('0'),
                    DataType = "Exception",
                    Text = text,
                    RelevanceScore = 1.0, // SQL 檢索預設為完全相關
                    Metadata = new Dictionary<string, object>
                        {
                            { "LOGDateTime", row["LOGDateTime"] },
                            { "Causes", row["Event"] },
                            { "NeedId", row["NeedId"]== DBNull.Value ? "" : row["NeedId"] },
                            { "SimulationId", row["SimulationId"]== DBNull.Value ? "" : row["SimulationId"] },
                            { "StationNO", row["Source_StationNO"]== DBNull.Value ? "" : row["Source_StationNO"] },
                            { "DOCNumberNO", row["DOCNumberNO"]== DBNull.Value ? "" : row["DOCNumberNO"] },
                            { "PartNO", row["PartNO"]== DBNull.Value ? "" : row["PartNO"] }
                         }
                });
            }
            #endregion

            #region OperateLog紀錄干涉
            tmp = "";
            if (start.HasValue) tmp += " AND a.LOGDateTime >= @StartDate";
            if (end.HasValue) tmp += " AND a.LOGDateTime <= @EndDate";

            sql = $@"SELECT TOP 100 a.Id,a.ServerId,a.StationNO,a.LOGDateTime,a.NeedId,a.SimulationId,a.OperateType,a.PartNO,b.PartName,b.Specification,a.OrderNO,a.OP_NO from SoftNetLogDB.[dbo].[OperateLog] as a
                            LEFT JOIN SoftNetMainDB.[dbo].[Material] as b ON b.PartNO = a.PartNO AND b.ServerId = a.ServerId
                        WHERE a.ServerId = @ServerId and OperateType like '%干涉%' {tmp} order by a.LOGDateTime desc";
            var sw4 = Stopwatch.StartNew();
            table = db.DB_GetDataByDataTable(sql, args);
            sw4.Stop();
            await _Log.InfoAsync($"RAG RetrieveExceptionData OperateLog(干涉) rows={table?.Rows?.Count ?? 0} ms={sw4.ElapsedMilliseconds}");
            foreach (DataRow row in table.Rows)
            {
                text = $"工單編號: {row["OrderNO"]}, 料件編號: {row["PartNO"]}, 品名: {row["PartName"]}, 規格: {row["Specification"]}, 工站編號: {row["StationNO"]}, 紀錄時間: {row["LOGDateTime"]}, 動作: {row["OperateType"]}, 作業員: {row["OP_NO"]}, 排程編號: {row["NeedId"]}, 子排程編號: {row["SimulationId"]}";

                results.Add(new RetrievedData
                {
                    DataId = row["Id"].ToString(),
                    //DataId = _Str.NewId('0'),
                    DataType = "Exception",
                    Text = text,
                    RelevanceScore = 1.0, // SQL 檢索預設為完全相關
                    Metadata = new Dictionary<string, object>
                        {
                            { "OrderNO", row["OrderNO"] },
                            { "PartNO", row["PartNO"] },
                            { "StationNO", row["StationNO"] },
                            { "PartName", row["PartName"] },
                            { "Specification", row["Specification"] },
                            { "Remark", row["OperateType"] },
                            { "OP_NO", row["OP_NO"] },
                            { "NeedId", row["NeedId"] },
                            { "SimulationId", row["SimulationId"] },
                            { "LOGDateTime", row["LOGDateTime"] }
                        }
                });
            }
            #endregion



            _SFC_Common.Dispose();
            return results;
        }
        /// <summary>
        /// 檢索物料資料
        /// </summary>
        private async Task<List<RetrievedData>> RetrieveMaterialData(RagIntentService.QueryIntent intent, DBADO db)
        {
            var results = new List<RetrievedData>();

            var whereConditions = new List<string>();
            var args = new List<object> { "ServerId", _Fun.Config.ServerId };

            // 料號
            var partNOEntities = intent.Entities.Where(e => e.Type == "PartNO").ToList();
            if (partNOEntities.Any())
            {
                var partNOs = string.Join(",", partNOEntities.Select(e => $"'{e.Value}'"));
                whereConditions.Add($"mat.PartNO IN ({partNOs})");
            }

            var whereClause = whereConditions.Any() ? "AND " + string.Join(" AND ", whereConditions) : "";

            string sql = $@"
                    SELECT TOP 50
                        mat.PartNO,
                        mat.PartName,
                        mat.Specification,
                        mat.Model,
                        mat.Class,
                        mat.Unit,
                        mat.SafeQTY
                    FROM SoftNetMainDB.[dbo].[Material] mat
                    WHERE mat.ServerId = @ServerId
                        {whereClause}
                    ORDER BY mat.PartNO
                ";

            var table = db.DB_GetDataByDataTable(sql, args);

            foreach (DataRow row in table.Rows)
            {
                var text = $"料件編號: {row["PartNO"]}, 品名: {row["PartName"]}, " +
                           $"規格: {row["Specification"]}, 型號: {row["Model"]}, " +
                           $"類別: {row["Class"]}, 單位: {row["Unit"]}, 安全存量: {row["SafeQTY"]}";

                results.Add(new RetrievedData
                {
                    DataId = row["PartNO"].ToString(),
                    DataType = "Material",
                    Text = text,
                    RelevanceScore = 1.0,
                    Metadata = new Dictionary<string, object>
                        {
                            { "PartNO", row["PartNO"] },
                            { "PartName", row["PartName"] },
                            { "Specification", row["Specification"] },
                            { "Model", row["Model"] },
                            { "Class", row["Class"] },
                            { "Unit", row["Unit"] },
                            { "SafeQTY", row["SafeQTY"] }
                        }
                });
            }


            return results;
        }
    }
}
