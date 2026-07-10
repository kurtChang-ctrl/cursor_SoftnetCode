using Base;
using Base.Models;
using Base.Services;
using BaseApi.Services;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using static SoftNetWebII.Services.RagIntentService;

namespace SoftNetWebII.Services
{
    public class WarRoomService
    {
        private string _Ctrl;
        private SFC_Common _SFC_Common;

        public WarRoomService(string ctrl, SFC_Common sfc_Common)
        {
            _Ctrl = ctrl;
            _SFC_Common = sfc_Common;
        }

        #region 輔助方法

        /// <summary>
        /// 取得工站的行事曆名稱（含預設值處理）
        /// </summary>
        private string GetStationCalendarName(DBADO db, string stationNO)
        {
            DataRow station = db.DB_GetFirstDataByDataRow(
                $"SELECT CalendarName FROM SoftNetSYSDB.[dbo].[PP_Station] " +
                $"WHERE ServerId = '{_Fun.Config.ServerId}' AND StationNO = '{stationNO}'");

            if (station != null && !station.IsNull("CalendarName") &&
                !string.IsNullOrWhiteSpace(station["CalendarName"].ToString()))
            {
                return station["CalendarName"].ToString();
            }

            return _Fun.Config.DefaultCalendarName;
        }

        /// <summary>
        /// 取得工站今日上班時間範圍（依行事曆）
        /// </summary>
        private (DateTime startTime, DateTime endTime) GetStationWorkingHours(DBADO db, string stationNO, DateTime targetDate)
        {
            string calendarName = GetStationCalendarName(db, stationNO);

            // 查詢該日期的行事曆設定
            DataRow calendar = db.DB_GetFirstDataByDataRow(
                $"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] " +
                $"WHERE ServerId = '{_Fun.Config.ServerId}' " +
                $"  AND CalendarName = '{calendarName}' " +
                $"  AND Holiday = '{targetDate:yyyy-MM-dd}'");

            if (calendar == null)
            {
                // 如果沒有行事曆設定，返回當日 00:00 到 23:59:59
                return (targetDate, targetDate.AddDays(1).AddSeconds(-1));
            }

            DateTime startTime = targetDate;
            DateTime endTime = targetDate.AddDays(1);
            List<DateTime> shiftStarts = new List<DateTime>();
            List<DateTime> shiftEnds = new List<DateTime>();

            // 依班別計算開始和結束時間
            if (bool.Parse(calendar["Flag_Morning"].ToString()))
            {
                string[] shift = calendar["Shift_Morning"].ToString().Split(',');
                shiftStarts.Add(targetDate.AddHours(int.Parse(shift[0].Split(':')[0]))
                                          .AddMinutes(int.Parse(shift[0].Split(':')[1])));
                shiftEnds.Add(targetDate.AddHours(int.Parse(shift[3].Split(':')[0]))
                                        .AddMinutes(int.Parse(shift[3].Split(':')[1])));
            }
            if (bool.Parse(calendar["Flag_Afternoon"].ToString()))
            {
                string[] shift = calendar["Shift_Afternoon"].ToString().Split(',');
                shiftStarts.Add(targetDate.AddHours(int.Parse(shift[0].Split(':')[0]))
                                          .AddMinutes(int.Parse(shift[0].Split(':')[1])));
                shiftEnds.Add(targetDate.AddHours(int.Parse(shift[3].Split(':')[0]))
                                        .AddMinutes(int.Parse(shift[3].Split(':')[1])));
            }
            if (bool.Parse(calendar["Flag_Night"].ToString()))
            {
                string[] shift = calendar["Shift_Night"].ToString().Split(',');
                shiftStarts.Add(targetDate.AddHours(int.Parse(shift[0].Split(':')[0]))
                                          .AddMinutes(int.Parse(shift[0].Split(':')[1])));
                // 晚班可能跨日
                if (int.Parse(shift[3].Split(':')[0]) < int.Parse(shift[0].Split(':')[0]))
                {
                    shiftEnds.Add(targetDate.AddDays(1).AddHours(int.Parse(shift[3].Split(':')[0]))
                                              .AddMinutes(int.Parse(shift[3].Split(':')[1])));
                }
                else
                {
                    shiftEnds.Add(targetDate.AddHours(int.Parse(shift[3].Split(':')[0]))
                                            .AddMinutes(int.Parse(shift[3].Split(':')[1])));
                }
            }
            if (bool.Parse(calendar["Flag_Graveyard"].ToString()))
            {
                string[] shift = calendar["Shift_Graveyard"].ToString().Split(',');
                shiftStarts.Add(targetDate.AddHours(int.Parse(shift[0].Split(':')[0]))
                                          .AddMinutes(int.Parse(shift[0].Split(':')[1])));
                // 大夜班通常跨日
                if (int.Parse(shift[3].Split(':')[0]) < int.Parse(shift[0].Split(':')[0]))
                {
                    shiftEnds.Add(targetDate.AddDays(1).AddHours(int.Parse(shift[3].Split(':')[0]))
                                              .AddMinutes(int.Parse(shift[3].Split(':')[1])));
                }
                else
                {
                    shiftEnds.Add(targetDate.AddHours(int.Parse(shift[3].Split(':')[0]))
                                            .AddMinutes(int.Parse(shift[3].Split(':')[1])));
                }
            }

            if (shiftStarts.Count > 0)
            {
                startTime = shiftStarts.Min();
                endTime = shiftEnds.Max();
            }

            return (startTime, endTime);
        }

        /// <summary>
        /// 計算工站實際工作時間（依行事曆）
        /// </summary>
        private int CalculateActualWorkingSeconds(DBADO db, string stationNO, DateTime targetDate)
        {
            string calendarName = GetStationCalendarName(db, stationNO);

            // 查詢該工站該日的所有開工/停工記錄
            string dateStr = targetDate.ToString("yyyy-MM-dd");
            DataTable dt = db.DB_GetData(
                $"SELECT * FROM SoftNetLogDB.[dbo].[OperateLog] " +
                $"WHERE ServerId = '{_Fun.Config.ServerId}' " +
                $"  AND StationNO = '{stationNO}' " +
                $"  AND CONVERT(DATE, LOGDateTime) = '{dateStr}' " +
                $"  AND (OperateType LIKE '%開工%' OR OperateType LIKE '%停工%' OR OperateType LIKE '%關站%') " +
                $"ORDER BY LOGDateTime");

            if (dt == null || dt.Rows.Count == 0)
                return 0;

            int totalSeconds = 0;
            DateTime? startTime = null;

            // 配對開工/停工記錄
            foreach (DataRow dr in dt.Rows)
            {
                string operateType = dr["OperateType"].ToString();
                DateTime logTime = Convert.ToDateTime(dr["LOGDateTime"]);

                if (operateType.IndexOf("開工") >= 0)
                {
                    // 開工記錄
                    if (startTime == null)
                    {
                        startTime = logTime;
                    }
                }
                else if (operateType.IndexOf("停工") >= 0 || operateType.IndexOf("關站") >= 0)
                {
                    // 停工/關站記錄
                    if (startTime != null)
                    {
                        // 使用 TimeCompute2Seconds_BY_Calendar 計算實際工作秒數
                        int seconds = _SFC_Common.TimeCompute2Seconds_BY_Calendar(
                            db, calendarName, startTime.Value, logTime);
                        totalSeconds += seconds;
                        startTime = null; // 重置
                    }
                }
            }

            // 如果最後是開工狀態（未停工），計算到目前時間
            if (startTime != null)
            {
                int seconds = _SFC_Common.TimeCompute2Seconds_BY_Calendar(
                    db, calendarName, startTime.Value, DateTime.Now);
                totalSeconds += seconds;
            }

            return totalSeconds;
        }

        #endregion

        #region 功能1：製造現場工作人員戰情室

        /// <summary>
        /// 取得今日/昨日產量統計（區分半成品與成品）
        /// </summary>
        public async Task<JObject> GetProductionOutputAsync(string dateType = "today")
        {
            return await Task.Run(() =>
            {
                DateTime targetDate = dateType == "today"
                    ? DateTime.Now.Date
                    : DateTime.Now.AddDays(-1).Date;
                DateTime nextDate = targetDate.AddDays(1);

                int semiFinishedQty = 0;
                int finishedQty = 0;

                try
                {
                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                    {
                        // 簡化查詢：直接使用日期範圍，不依賴複雜的行事曆計算
                        DataTable dtOutput = db.DB_GetData(
                            $@"SELECT 
                                m.Class,
                                SUM(s.ProductFinishedQty) as OutputQty
                              FROM SoftNetLogDB.[dbo].[SFC_StationDetail] s
                              JOIN SoftNetMainDB.[dbo].[Material] m 
                                ON s.PartNO = m.PartNO AND s.ServerId = m.ServerId
                              WHERE s.ServerId = '{_Fun.Config.ServerId}'
                                AND s.OutputFlag = 1
                                AND m.Class IN ('4', '5')
                                AND s.LOGDateTime >= '{targetDate:yyyy-MM-dd 00:00:00}'
                                AND s.LOGDateTime < '{nextDate:yyyy-MM-dd 00:00:00}'
                              GROUP BY m.Class");

                        // 累加產量
                        if (dtOutput != null)
                        {
                            foreach (DataRow output in dtOutput.Rows)
                            {
                                int qty = output.IsNull("OutputQty") ? 0 : Convert.ToInt32(output["OutputQty"]);
                                if (output["Class"].ToString() == "4")
                                    semiFinishedQty += qty;
                                else if (output["Class"].ToString() == "5")
                                    finishedQty += qty;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 記錄錯誤但不拋出，返回 0
                    System.Diagnostics.Debug.WriteLine($"GetProductionOutputAsync Error: {ex.Message}");
                }

                return JObject.FromObject(new
                {
                    SemiFinishedQty = semiFinishedQty,
                    FinishedQty = finishedQty,
                    TotalQty = semiFinishedQty + finishedQty,
                    DateType = dateType
                });
            });
        }

        /// <summary>
        /// 取得工單統計（進行中、已完成、達成率）
        /// </summary>
        public async Task<JObject> GetWorkOrderStatsAsync()
        {
            return await Task.Run(() =>
            {
                int activeCount = 0;
                int completedCount = 0;
                double achievementRate = 0;

                try
                {
                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                    {
                        // 進行中工單（EndTime IS NULL）
                        DataTable dtActive = db.DB_GetData(
                            $"SELECT COUNT(*) as Count FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] " +
                            $"WHERE ServerId = '{_Fun.Config.ServerId}' AND EndTime IS NULL");
                        if (dtActive != null && dtActive.Rows.Count > 0)
                        {
                            activeCount = dtActive.Rows[0].IsNull("Count") ? 0 : Convert.ToInt32(dtActive.Rows[0]["Count"]);
                        }

                        // 今日已完成工單（EndTime 在今日）
                        DateTime today = DateTime.Now.Date;
                        DateTime tomorrow = today.AddDays(1);
                        DataTable dtCompleted = db.DB_GetData(
                            $"SELECT COUNT(*) as Count FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] " +
                            $"WHERE ServerId = '{_Fun.Config.ServerId}' " +
                            $"AND EndTime >= '{today:yyyy-MM-dd 00:00:00}' " +
                            $"AND EndTime < '{tomorrow:yyyy-MM-dd 00:00:00}'");
                        if (dtCompleted != null && dtCompleted.Rows.Count > 0)
                        {
                            completedCount = dtCompleted.Rows[0].IsNull("Count") ? 0 : Convert.ToInt32(dtCompleted.Rows[0]["Count"]);
                        }

                        // 達成率：今日已完成工單 / 今日應完成工單（簡化計算：使用今日開始且已完成的工單）
                        DataTable dtTodayStarted = db.DB_GetData(
                            $"SELECT COUNT(*) as Count FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] " +
                            $"WHERE ServerId = '{_Fun.Config.ServerId}' " +
                            $"AND StartTime >= '{today:yyyy-MM-dd 00:00:00}' " +
                            $"AND StartTime < '{tomorrow:yyyy-MM-dd 00:00:00}'");
                        int todayStarted = 0;
                        if (dtTodayStarted != null && dtTodayStarted.Rows.Count > 0)
                        {
                            todayStarted = dtTodayStarted.Rows[0].IsNull("Count") ? 0 : Convert.ToInt32(dtTodayStarted.Rows[0]["Count"]);
                        }

                        if (todayStarted > 0)
                        {
                            achievementRate = Math.Round((double)completedCount / todayStarted * 100, 1);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"GetWorkOrderStatsAsync Error: {ex.Message}");
                }

                return JObject.FromObject(new
                {
                    ActiveWorkOrders = activeCount,
                    CompletedWorkOrders = completedCount,
                    AchievementRate = achievementRate
                });
            });
        }

        /// <summary>
        /// 取得工站即時狀態（支援多工單）
        /// </summary>
        public async Task<JObject> GetStationStatusAsync()
        {
            return await Task.Run(() =>
            {
                var stationList = new List<object>();

                try
                {
                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                    {
                    // 查詢單工單
                    DataTable dtSingle = db.DB_GetData(
                        $@"SELECT 
                            m.StationNO,
                            m.OrderNO,
                            m.State,
                            m.PartNO,
                            m.PP_Name,
                            m.IndexSN,
                            m.Station_Custom_IndexSN,
                            m.StationNO_Custom_DisplayName,
                            m.OP_NO,
                            ps.StationName,
                            ps.FactoryName,
                            ps.LineName,
                            mat.PartName,
                            mat.Specification,
                            s.AvarageCycleTime,
                            s.TotalOutput,
                            s.TotalFail,
                            ISNULL(ed.EfficientCycleTime, 0) as EfficientCycleTime
                          FROM SoftNetMainDB.[dbo].[Manufacture] m
                          JOIN SoftNetSYSDB.[dbo].[PP_Station] ps ON m.StationNO = ps.StationNO
                          LEFT JOIN SoftNetMainDB.[dbo].[Material] mat ON m.PartNO = mat.PartNO AND m.ServerId = mat.ServerId
                          LEFT JOIN SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] s ON m.StationNO = s.StationNO AND m.OrderNO = s.OrderNO AND s.ServerId = '{_Fun.Config.ServerId}' AND (s.EndTime IS NULL OR s.EndTime = '')
                          LEFT JOIN (
                              SELECT StationNO, PP_Name, IndexSN, PartNO, AVG(EfficientCycleTime) as EfficientCycleTime
                              FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail]
                              WHERE ServerId = '{_Fun.Config.ServerId}' AND DOCNO = ''
                              GROUP BY StationNO, PP_Name, IndexSN, PartNO
                          ) ed ON m.StationNO = ed.StationNO AND m.PP_Name = ed.PP_Name AND m.IndexSN = ed.IndexSN AND m.PartNO = ed.PartNO
                          WHERE m.ServerId = '{_Fun.Config.ServerId}'
                            AND m.Config_MutiWO = '0'
                            AND m.EndTime IS NULL
                            AND m.OrderNO IS NOT NULL
                            AND m.OrderNO != ''");

                    if (dtSingle != null)
                    {
                        foreach (DataRow dr in dtSingle.Rows)
                        {
                            // 過濾掉沒有工單編號的工站
                            string orderNO = dr["OrderNO"].ToString();
                            if (string.IsNullOrWhiteSpace(orderNO))
                            {
                                continue;
                            }

                            // 計算不良率
                            decimal totalOutput = dr.IsNull("TotalOutput") ? 0 : Convert.ToDecimal(dr["TotalOutput"]);
                            decimal totalFail = dr.IsNull("TotalFail") ? 0 : Convert.ToDecimal(dr["TotalFail"]);
                            decimal failRate = totalOutput > 0 ? (totalFail / totalOutput * 100) : 0;

                            stationList.Add(new
                            {
                                StationNO = dr["StationNO"].ToString(),
                                StationName = dr["StationName"].ToString(),
                                FactoryName = dr["FactoryName"].ToString(),
                                LineName = dr["LineName"].ToString(),
                                OrderNO = dr["OrderNO"].ToString(),
                                State = dr["State"].ToString(),
                                PartNO = dr["PartNO"].ToString(),
                                PartName = dr.IsNull("PartName") ? "" : dr["PartName"].ToString(),
                                Specification = dr.IsNull("Specification") ? "" : dr["Specification"].ToString(),
                                PP_Name = dr["PP_Name"].ToString(),
                                IndexSN = dr["IndexSN"].ToString(),
                                Station_Custom_IndexSN = dr.IsNull("Station_Custom_IndexSN") ? "" : dr["Station_Custom_IndexSN"].ToString(),
                                StationNO_Custom_DisplayName = dr.IsNull("StationNO_Custom_DisplayName") ? "" : dr["StationNO_Custom_DisplayName"].ToString(),
                                OP_NO = dr.IsNull("OP_NO") ? "" : dr["OP_NO"].ToString(),
                                EfficientCycleTime = dr.IsNull("EfficientCycleTime") ? 0 : Convert.ToDecimal(dr["EfficientCycleTime"]),
                                AvarageCycleTime = dr.IsNull("AvarageCycleTime") ? 0 : Convert.ToDecimal(dr["AvarageCycleTime"]),
                                TotalOutput = totalOutput,
                                TotalFail = totalFail,
                                FailRate = failRate,
                                IsMultiWO = false
                            });
                        }
                    }

                    // 查詢多工單（一個工站多筆工單）
                    DataTable dtMulti = db.DB_GetData(
                        $@"SELECT 
                            mii.StationNO,
                            mii.OrderNO,
                            mii.PartNO,
                            mii.PP_Name,
                            mii.IndexSN,
                            mii.Station_Custom_IndexSN,
                            mii.StationNO_Custom_DisplayName,
                            m.OP_NO,
                            m.State,
                            ps.StationName,
                            ps.FactoryName,
                            ps.LineName,
                            mat.PartName,
                            mat.Specification,
                            s.AvarageCycleTime,
                            s.TotalOutput,
                            s.TotalFail,
                            ISNULL(ed.EfficientCycleTime, 0) as EfficientCycleTime
                          FROM SoftNetMainDB.[dbo].[ManufactureII] mii
                          JOIN SoftNetSYSDB.[dbo].[PP_Station] ps ON mii.StationNO = ps.StationNO
                          LEFT JOIN SoftNetMainDB.[dbo].[Manufacture] m ON mii.StationNO = m.StationNO AND m.ServerId = '{_Fun.Config.ServerId}' AND m.EndTime IS NULL
                          LEFT JOIN SoftNetMainDB.[dbo].[Material] mat ON mii.PartNO = mat.PartNO AND mii.ServerId = mat.ServerId
                          LEFT JOIN SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] s ON mii.StationNO = s.StationNO AND mii.OrderNO = s.OrderNO AND s.ServerId = '{_Fun.Config.ServerId}' AND (s.EndTime IS NULL OR s.EndTime = '')
                          LEFT JOIN (
                              SELECT StationNO, PP_Name, IndexSN, PartNO, AVG(EfficientCycleTime) as EfficientCycleTime
                              FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail]
                              WHERE ServerId = '{_Fun.Config.ServerId}' AND DOCNO = ''
                              GROUP BY StationNO, PP_Name, IndexSN, PartNO
                          ) ed ON mii.StationNO = ed.StationNO AND mii.PP_Name = ed.PP_Name AND mii.IndexSN = ed.IndexSN AND mii.PartNO = ed.PartNO
                          WHERE mii.ServerId = '{_Fun.Config.ServerId}'
                            AND mii.EndTime IS NULL
                            AND mii.OrderNO IS NOT NULL
                            AND mii.OrderNO != ''
                          ORDER BY mii.StationNO, mii.OrderNO");

                    if (dtMulti != null)
                    {
                        // 依工站分組
                        var grouped = dtMulti.AsEnumerable()
                            .GroupBy(r => r["StationNO"].ToString())
                            .ToList();

                        foreach (var group in grouped)
                        {
                            var orders = group.Select(r => r["OrderNO"].ToString())
                                .Where(o => !string.IsNullOrWhiteSpace(o))
                                .ToList();
                            
                            // 過濾掉沒有工單編號的工站
                            if (orders.Count == 0)
                            {
                                continue;
                            }

                            var firstRow = group.First();
                            var state = firstRow.IsNull("State") ? "4" : firstRow["State"].ToString();

                            stationList.Add(new
                            {
                                StationNO = firstRow["StationNO"].ToString(),
                                StationName = firstRow["StationName"].ToString(),
                                FactoryName = firstRow["FactoryName"].ToString(),
                                LineName = firstRow["LineName"].ToString(),
                                OrderNO = string.Join(", ", orders), // 並排顯示多個工單
                                State = state,
                                PartNO = firstRow["PartNO"].ToString(),
                                PartName = firstRow.IsNull("PartName") ? "" : firstRow["PartName"].ToString(),
                                Specification = firstRow.IsNull("Specification") ? "" : firstRow["Specification"].ToString(),
                                PP_Name = firstRow["PP_Name"].ToString(),
                                IndexSN = firstRow["IndexSN"].ToString(),
                                Station_Custom_IndexSN = firstRow.IsNull("Station_Custom_IndexSN") ? "" : firstRow["Station_Custom_IndexSN"].ToString(),
                                StationNO_Custom_DisplayName = firstRow.IsNull("StationNO_Custom_DisplayName") ? "" : firstRow["StationNO_Custom_DisplayName"].ToString(),
                                OP_NO = firstRow.IsNull("OP_NO") ? "" : firstRow["OP_NO"].ToString(),
                                EfficientCycleTime = firstRow.IsNull("EfficientCycleTime") ? 0 : Convert.ToDecimal(firstRow["EfficientCycleTime"]),
                                AvarageCycleTime = firstRow.IsNull("AvarageCycleTime") ? 0 : Convert.ToDecimal(firstRow["AvarageCycleTime"]),
                                TotalOutput = firstRow.IsNull("TotalOutput") ? 0 : Convert.ToDecimal(firstRow["TotalOutput"]),
                                TotalFail = firstRow.IsNull("TotalFail") ? 0 : Convert.ToDecimal(firstRow["TotalFail"]),
                                FailRate = firstRow.IsNull("TotalOutput") || Convert.ToDecimal(firstRow["TotalOutput"]) == 0 ? 0 : (Convert.ToDecimal(firstRow["TotalFail"]) / Convert.ToDecimal(firstRow["TotalOutput"]) * 100),
                                IsMultiWO = true,
                                WorkOrders = group.Select(r => new
                                {
                                    OrderNO = r["OrderNO"].ToString(),
                                    PartNO = r["PartNO"].ToString(),
                                    PartName = r.IsNull("PartName") ? "" : r["PartName"].ToString(),
                                    Specification = r.IsNull("Specification") ? "" : r["Specification"].ToString(),
                                    PP_Name = r["PP_Name"].ToString(),
                                    IndexSN = r["IndexSN"].ToString(),
                                    Station_Custom_IndexSN = r.IsNull("Station_Custom_IndexSN") ? "" : r["Station_Custom_IndexSN"].ToString(),
                                    StationNO_Custom_DisplayName = r.IsNull("StationNO_Custom_DisplayName") ? "" : r["StationNO_Custom_DisplayName"].ToString(),
                                    EfficientCycleTime = r.IsNull("EfficientCycleTime") ? 0 : Convert.ToDecimal(r["EfficientCycleTime"]),
                                    AvarageCycleTime = r.IsNull("AvarageCycleTime") ? 0 : Convert.ToDecimal(r["AvarageCycleTime"]),
                                    TotalOutput = r.IsNull("TotalOutput") ? 0 : Convert.ToDecimal(r["TotalOutput"]),
                                    TotalFail = r.IsNull("TotalFail") ? 0 : Convert.ToDecimal(r["TotalFail"]),
                                    FailRate = r.IsNull("TotalOutput") || Convert.ToDecimal(r["TotalOutput"]) == 0 ? 0 : (Convert.ToDecimal(r["TotalFail"]) / Convert.ToDecimal(r["TotalOutput"]) * 100)
                                }).ToList()
                            });
                        }
                    }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"GetStationStatusAsync Error: {ex.Message}");
                    // 返回空列表，避免頁面完全卡住
                }

                return JObject.FromObject(new
                {
                    Stations = stationList
                });
            });
        }

        /// <summary>
        /// 取得即時警示（主動干涉引擎）- 參考 APSErrorData 顯示格式
        /// </summary>
        public async Task<JObject> GetActiveInterventionAlertsAsync(int limit = 10)
        {
            return await Task.Run(() =>
            {
                var alerts = new List<object>();

                try
                {
                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                    {
                        // 查詢前4小時內的資料
                        DateTime fourHoursAgo = DateTime.Now.AddHours(-4);
                        string data_tmp = "";
                        if (!_Fun.Config.DOCInvoice_Expiry_SendError)
                        {
                            data_tmp = $" and ErrorType!='06' and ErrorType!='07' and ErrorType!='08' and ErrorType!='09' and ErrorType!='10'";
                        }

                        // 查詢 APS_Simulation_ErrorData
                        DataTable dtdata = db.DB_GetData(
                            $@"SELECT ServerId,NeedId,SimulationId,DOCNumberNO,StationNO,ErrorType,LogDate,'' as OP_NO 
                               FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] 
                               WHERE ServerId='{_Fun.Config.ServerId}' 
                                 AND LogDate>='{fourHoursAgo:yyyy-MM-dd HH:mm:ss}' 
                                 {data_tmp} 
                               ORDER BY LogDate DESC");

                        // 查詢 APS_WarningData
                        DataTable dtdata2 = db.DB_GetData(
                            $@"SELECT ServerId,NeedId,SimulationId,DOCNumberNO,StationNO,ErrorType,WarningDate as LogDate,OP_NO 
                               FROM SoftNetSYSDB.[dbo].[APS_WarningData] 
                               WHERE ServerId='{_Fun.Config.ServerId}' and ErrorType!='00' and ErrorType!='33'
                                 AND WarningDate>='{fourHoursAgo:yyyy-MM-dd HH:mm:ss}' 
                                 AND IsDEL = 0
                               ORDER BY WarningDate DESC");

                        #region OperateLog紀錄干涉
                        string selectSql = $@"SELECT ServerId,NeedId,SimulationId,OrderNO as DOCNumberNO,StationNO,OperateType as ErrorType,LOGDateTime as LogDate,OP_NO 
                        from SoftNetLogDB.[dbo].[OperateLog] WHERE ServerId = @ServerId and LOGDateTime>=@LOGDateTime and OperateType like N'%干涉%' 
                        order by LOGDateTime DESC";
                        var sourceData = db.DB_GetDataByDataTable(selectSql,
                            new List<object>
                                {
                                    "ServerId", _Fun.Config.ServerId,
                                    "LOGDateTime", fourHoursAgo
                                }, true);
                        #endregion

                        #region APS_EventLog 異常,干涉
                        selectSql = $@"SELECT ServerId,NeedId,SimulationId,'' as DOCNumberNO, '' as StationNO,Event as ErrorType,LOGDateTime as LogDate,'' as OP_NO
                        from SoftNetLogDB.[dbo].[APS_EventLog] WHERE ServerId = @ServerId and LOGDateTime>=@LOGDateTime and (Event like N'%異常%' or EventType like N'E%')
                    order by LOGDateTime";
                        var sourceData2 = db.DB_GetDataByDataTable(selectSql,
                            new List<object>
                                {
                        "ServerId", _Fun.Config.ServerId,
                        "LOGDateTime", fourHoursAgo
                                }, true);

                        #endregion


                        // 合併資料表
                        if (sourceData != null && sourceData.Rows.Count > 0)
                        {
                            dtdata2.Merge(sourceData);
                        }
                        if (sourceData2 != null && sourceData2.Rows.Count > 0)
                        {
                            dtdata2.Merge(sourceData2);
                        }
                        if (dtdata != null)
                        {
                            if (dtdata2 != null && dtdata2.Rows.Count > 0)
                            {
                                dtdata.Merge(dtdata2);
                                dtdata.DefaultView.Sort = "LogDate DESC";
                            }
                        }
                        else
                        {
                            if (dtdata2 != null && dtdata2.Rows.Count > 0)
                            {
                                dtdata = dtdata2;
                            }
                        }

                        if (dtdata != null && dtdata.Rows.Count > 0)
                        {
                            string bdfor_SimulationId = "";
                            string bdfor_DOCNumberNO = "";
                            string bdfor_StationNO = "";
                            string bdfor_ErrorType = "";

                            foreach (DataRow dr in dtdata.Rows)
                            {
                                // 避免重複資料
                                if (bdfor_SimulationId == dr["SimulationId"].ToString() && 
                                    bdfor_DOCNumberNO == dr["DOCNumberNO"].ToString() && 
                                    bdfor_StationNO == dr["StationNO"].ToString() && 
                                    bdfor_ErrorType == dr["ErrorType"].ToString())
                                {
                                    continue;
                                }
                                else
                                {
                                    bdfor_SimulationId = dr["SimulationId"].ToString();
                                    bdfor_DOCNumberNO = dr["DOCNumberNO"].ToString();
                                    bdfor_StationNO = dr["StationNO"].ToString();
                                    bdfor_ErrorType = dr["ErrorType"].ToString();
                                }

                                // 查詢料號和品名
                                DataRow tmp_part = db.DB_GetFirstDataByDataRow(
                                    $@"SELECT a.PartNO,b.PartName,b.Specification 
                                       FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                       JOIN SoftNetMainDB.[dbo].[Material] as b ON b.PartNO=a.PartNO AND b.ServerId=a.ServerId
                                       WHERE a.ServerId='{_Fun.Config.ServerId}' 
                                         AND a.SimulationId='{dr["SimulationId"].ToString()}'");

                                string name = "";//料件名稱
                                string log = "";
                                string partNO = "";
                                string station = dr.IsNull("StationNO") ? "" : dr["StationNO"].ToString();
                                string DOCName = dr.IsNull("DOCNumberNO") ? "" : dr["DOCNumberNO"].ToString();
                                if (!dr.IsNull("DOCNumberNO") && dr["DOCNumberNO"].ToString() != "")
                                {
                                    DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOCRole] where DOCNO='{dr["DOCNumberNO"].ToString().Substring(0, 4)}' and ServerId='{_Fun.Config.ServerId}'");
                                    if (tmp != null) { DOCName = $"{DOCName} {tmp["DOCName"].ToString()}"; }
                                }

                                if (tmp_part != null)
                                {
                                    partNO = tmp_part.IsNull("PartNO") ? "" : tmp_part["PartNO"].ToString();
                                    string partName = tmp_part.IsNull("PartName") ? "" : tmp_part["PartName"].ToString();
                                    string specification = tmp_part.IsNull("Specification") ? "" : tmp_part["Specification"].ToString();
                                    name = partNO + partName + (specification != "" ? " / " + specification : "");
                                }

                                // 取得錯誤類型說明
                                string tmp_S = "";
                                switch (dr["ErrorType"].ToString())
                                {
                                    case "05":
                                        log = "生產備料無作為且無領料單據發起";
                                        break;
                                    case "01":
                                    case "02":
                                        log = $"庫存量不足原本預定Keep量";
                                        break;
                                    case "04":
                                        log = $"生產用原物料應領用而未領用";
                                        break;
                                    case "06":
                                    case "07":
                                    case "08":
                                    case "09":
                                    case "10":
                                        log = $"單據未如期完成確認扣(入)帳,會導致庫存數量正確性";
                                        break;
                                    case "03":
                                        log = $"未如預期開工";
                                        break;
                                    case "12":
                                        log = $"最後一站應執行結束工單作業";
                                        break;
                                    case "14":
                                        log = $"實際CT遠離目標CT";
                                        break;
                                    case "15":
                                        log = $"報工作業有延遲,會影響目標CT的正確性";
                                        break;
                                    case "17":
                                        log = $"工站數量移轉量不合理";
                                        break;
                                    case "18":
                                        log = $"應停工未停工";
                                        break;
                                    case "19":
                                        log = $"計畫已偏離需求日";
                                        break;
                                    case "21":
                                        log = $"工單關閉,準備強制檢查各站報工的合理性";
                                        break;
                                    case "22":
                                        log = $"疑似重複報工";
                                        break;
                                    case "23":
                                        log = $"疑似委外加工數量多出";
                                        break;
                                    case "24":
                                        log = $"工單關閉後才補開此委外加工單";
                                        break;
                                    case "25":
                                        log = $"系統干涉補報工數量";
                                        break;
                                    case "26":
                                        log = $"領入庫單據時間晚於計畫逾7天";
                                        break;
                                    case "27":
                                        log = $"領入庫單據時間早於計畫逾7天";
                                        break;
                                    case "28":
                                        log = $"計畫日已過報工明顯不足";
                                        break;
                                    case "29":
                                        log = $"工站移轉或在制數量為負";
                                        break;
                                    case "30":
                                        log = $"報工時間遠超計畫工時";
                                        break;
                                    case "31":
                                        log = $"領料或入庫數量為負";
                                        break;
                                    default:
                                        log = dr["ErrorType"].ToString();
                                        break;
                                }

                                alerts.Add(new
                                {
                                    Time = dr["LogDate"].ToString(),
                                    StationNO = station,
                                    PartName = name,
                                    DOCNumberNO = DOCName,
                                    Log = log
                                });
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"GetActiveInterventionAlertsAsync Error: {ex.Message}");
                    // 返回空列表，避免頁面完全卡住
                }

                return JObject.FromObject(new
                {
                    Alerts = alerts
                });
            });
        }

        /// <summary>
        /// 取得我的工單任務
        /// </summary>
        public async Task<JObject> GetMyWorkOrdersAsync(string stationNO)
        {
            return await Task.Run(() =>
            {
                var workOrders = new List<object>();

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    // 查詢該工站的工單
                    DataTable dt = db.DB_GetData(
                        $@"SELECT 
                            wo.OrderNO,
                            wo.PartNO,
                            mat.PartName,
                            mat.Specification,
                            wo.Quantity,
                            wo.ActualQuantity,
                            wo.EstimatedStartTime,
                            wo.StartTime,
                            wo.EndTime,
                            wo.WOStatus
                          FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] wo
                          LEFT JOIN SoftNetMainDB.[dbo].[Material] mat ON wo.PartNO = mat.PartNO AND wo.ServerId = mat.ServerId
                          WHERE wo.ServerId = '{_Fun.Config.ServerId}'
                            AND wo.OrderNO IN (
                                SELECT DISTINCT OrderNO 
                                FROM SoftNetMainDB.[dbo].[Manufacture] 
                                WHERE ServerId = '{_Fun.Config.ServerId}' AND StationNO = '{stationNO}' AND EndTime IS NULL
                                UNION
                                SELECT DISTINCT OrderNO 
                                FROM SoftNetMainDB.[dbo].[ManufactureII] 
                                WHERE ServerId = '{_Fun.Config.ServerId}' AND StationNO = '{stationNO}' AND EndTime IS NULL
                            )
                            AND wo.EndTime IS NULL
                          ORDER BY wo.EstimatedStartTime");

                    if (dt != null)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            int planQty = dr.IsNull("Quantity") ? 0 : Convert.ToInt32(dr["Quantity"]);
                            int actualQty = dr.IsNull("ActualQuantity") ? 0 : Convert.ToInt32(dr["ActualQuantity"]);
                            double progress = planQty > 0 ? (double)actualQty / planQty * 100 : 0;

                            workOrders.Add(new
                            {
                                OrderNO = dr["OrderNO"].ToString(),
                                PartNO = dr["PartNO"].ToString(),
                                PartName = dr.IsNull("PartName") ? "" : dr["PartName"].ToString(),
                                Specification = dr.IsNull("Specification") ? "" : dr["Specification"].ToString(),
                                PlanQty = planQty,
                                ActualQty = actualQty,
                                RemainingQty = planQty - actualQty,
                                Progress = Math.Round(progress, 1),
                                EstimatedStartTime = dr.IsNull("EstimatedStartTime") ? "" : dr["EstimatedStartTime"].ToString(),
                                StartTime = dr.IsNull("StartTime") ? "" : dr["StartTime"].ToString(),
                                EndTime = dr.IsNull("EndTime") ? "" : dr["EndTime"].ToString(),
                                WOStatus = dr["WOStatus"].ToString()
                            });
                        }
                    }
                }

                return JObject.FromObject(new
                {
                    WorkOrders = workOrders
                });
            });
        }

        /// <summary>
        /// 取得所有工站列表
        /// </summary>
        public async Task<JObject> GetAllStationsAsync()
        {
            return await Task.Run(() =>
            {
                var stations = new List<object>();

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataTable dt = db.DB_GetData(
                        $"SELECT StationNO, StationName, FactoryName, LineName " +
                        $"FROM SoftNetSYSDB.[dbo].[PP_Station] " +
                        $"WHERE ServerId = '{_Fun.Config.ServerId}' " +
                        $"ORDER BY FactoryName, LineName, StationNO");

                    if (dt != null)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            stations.Add(new
                            {
                                StationNO = dr["StationNO"].ToString(),
                                StationName = dr["StationName"].ToString(),
                                FactoryName = dr.IsNull("FactoryName") ? "" : dr["FactoryName"].ToString(),
                                LineName = dr.IsNull("LineName") ? "" : dr["LineName"].ToString()
                            });
                        }
                    }
                }

                return JObject.FromObject(new
                {
                    Stations = stations
                });
            });
        }

        /// <summary>
        /// 取得物料需求狀態 - 從 Manufacture/ManufactureII 的 SimulationId 查詢 DOC3stockII 和 DOC4ProductionII
        /// </summary>
        public async Task<JObject> GetMaterialRequirementsAsync(string orderNO)
        {
            return await Task.Run(() =>
            {
                var materials = new List<object>();

                try
                {
                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                    {
                        // 取得所有有 SimulationId 的工站
                        DataTable dtStations = db.DB_GetData(
                            $@"SELECT DISTINCT StationNO, SimulationId 
                               FROM (
                                   SELECT StationNO, SimulationId 
                                   FROM SoftNetMainDB.[dbo].[Manufacture] 
                                   WHERE ServerId = '{_Fun.Config.ServerId}' 
                                     AND SimulationId IS NOT NULL 
                                     AND SimulationId != ''
                                     AND EndTime IS NULL
                                   UNION
                                   SELECT StationNO, SimulationId 
                                   FROM SoftNetMainDB.[dbo].[ManufactureII] 
                                   WHERE ServerId = '{_Fun.Config.ServerId}' 
                                     AND SimulationId IS NOT NULL 
                                     AND SimulationId != ''
                                     AND EndTime IS NULL
                               ) AS Combined");

                        if (dtStations == null || dtStations.Rows.Count == 0)
                        {
                            return JObject.FromObject(new { Materials = materials });
                        }

                        // 取得需要排除的工站（關站或停工）
                        DateTime fourHoursAgo = DateTime.Now.AddHours(-4);
                        DataTable dtExcludedStations = db.DB_GetData(
                            $@"SELECT DISTINCT StationNO 
                               FROM SoftNetLogDB.[dbo].[OperateLog] 
                               WHERE ServerId = '{_Fun.Config.ServerId}' 
                                 AND (
                                     (OperateType LIKE '%關站%')
                                     OR (OperateType LIKE '%停工%' AND LOGDateTime >= '{fourHoursAgo:yyyy-MM-dd HH:mm:ss}')
                                 )");

                        HashSet<string> excludedStations = new HashSet<string>();
                        if (dtExcludedStations != null)
                        {
                            foreach (DataRow dr in dtExcludedStations.Rows)
                            {
                                excludedStations.Add(dr["StationNO"].ToString());
                            }
                        }

                        // 建立 SimulationId 列表
                        List<string> simulationIds = new List<string>();
                        Dictionary<string, string> simulationToStation = new Dictionary<string, string>();
                        foreach (DataRow dr in dtStations.Rows)
                        {
                            string stationNO = dr["StationNO"].ToString();
                            string simulationId = dr["SimulationId"].ToString();
                            
                            // 排除關站或停工的工站
                            if (excludedStations.Contains(stationNO))
                            {
                                continue;
                            }

                            if (!string.IsNullOrWhiteSpace(simulationId) && !simulationIds.Contains(simulationId))
                            {
                                simulationIds.Add(simulationId);
                                simulationToStation[simulationId] = stationNO;
                            }
                        }

                        if (simulationIds.Count == 0)
                        {
                            return JObject.FromObject(new { Materials = materials });
                        }

                        string simulationIdsStr = string.Join("','", simulationIds);

                        // 查詢 DOC3stockII
                        DataTable dtDOC3 = db.DB_GetData(
                            $@"SELECT 
                                d3.ArrivalDate as NeedDate,
                                m.StationNO,
                                d3.DOCNumberNO,
                                d3.PartNO,
                                mat.PartName,
                                mat.Specification,
                                d3.QTY,
                                d3.Unit,
                                CASE 
                                    WHEN d3.IN_StoreNO != '' THEN d3.IN_StoreNO
                                    WHEN d3.OUT_StoreNO != '' THEN d3.OUT_StoreNO
                                    ELSE ''
                                END as StoreNO,
                                CASE 
                                    WHEN d3.IN_StoreSpacesNO != '' THEN d3.IN_StoreSpacesNO
                                    WHEN d3.OUT_StoreSpacesNO != '' THEN d3.OUT_StoreSpacesNO
                                    ELSE ''
                                END as StoreSpacesNO,
                                d3.Remark,
                                d3.SimulationId
                               FROM SoftNetMainDB.[dbo].[DOC3stockII] d3
                               JOIN SoftNetMainDB.[dbo].[Material] mat ON d3.PartNO = mat.PartNO AND d3.ServerId = mat.ServerId
                               JOIN (
                                   SELECT DISTINCT StationNO, SimulationId 
                                   FROM (
                                       SELECT StationNO, SimulationId 
                                       FROM SoftNetMainDB.[dbo].[Manufacture] 
                                       WHERE ServerId = '{_Fun.Config.ServerId}' 
                                         AND SimulationId IS NOT NULL 
                                         AND SimulationId != ''
                                         AND EndTime IS NULL
                                       UNION
                                       SELECT StationNO, SimulationId 
                                       FROM SoftNetMainDB.[dbo].[ManufactureII] 
                                       WHERE ServerId = '{_Fun.Config.ServerId}' 
                                         AND SimulationId IS NOT NULL 
                                         AND SimulationId != ''
                                         AND EndTime IS NULL
                                   ) AS Combined
                               ) m ON d3.SimulationId = m.SimulationId
                               WHERE d3.ServerId = '{_Fun.Config.ServerId}' 
                                 AND d3.IsOK = '0'
                                 AND d3.SimulationId IN ('{simulationIdsStr}')
                                 AND m.StationNO NOT IN (
                                     SELECT DISTINCT StationNO 
                                     FROM SoftNetLogDB.[dbo].[OperateLog] 
                                     WHERE ServerId = '{_Fun.Config.ServerId}' 
                                       AND (
                                           (OperateType LIKE '%關站%')
                                           OR (OperateType LIKE '%停工%' AND LOGDateTime >= '{fourHoursAgo:yyyy-MM-dd HH:mm:ss}')
                                       )
                                 )");

                        // 查詢 DOC4ProductionII
                        DataTable dtDOC4 = db.DB_GetData(
                            $@"SELECT 
                                d4.ArrivalDate as NeedDate,
                                m.StationNO,
                                d4.DOCNumberNO,
                                d4.PartNO,
                                mat.PartName,
                                mat.Specification,
                                d4.QTY,
                                d4.Unit,
                                d4.StoreNO,
                                d4.StoreSpacesNO,
                                d4.Remark,
                                d4.SimulationId
                               FROM SoftNetMainDB.[dbo].[DOC4ProductionII] d4
                               JOIN SoftNetMainDB.[dbo].[Material] mat ON d4.PartNO = mat.PartNO AND d4.ServerId = mat.ServerId
                               JOIN (
                                   SELECT DISTINCT StationNO, SimulationId 
                                   FROM (
                                       SELECT StationNO, SimulationId 
                                       FROM SoftNetMainDB.[dbo].[Manufacture] 
                                       WHERE ServerId = '{_Fun.Config.ServerId}' 
                                         AND SimulationId IS NOT NULL 
                                         AND SimulationId != ''
                                         AND EndTime IS NULL
                                       UNION
                                       SELECT StationNO, SimulationId 
                                       FROM SoftNetMainDB.[dbo].[ManufactureII] 
                                       WHERE ServerId = '{_Fun.Config.ServerId}' 
                                         AND SimulationId IS NOT NULL 
                                         AND SimulationId != ''
                                         AND EndTime IS NULL
                                   ) AS Combined
                               ) m ON d4.SimulationId = m.SimulationId
                               WHERE d4.ServerId = '{_Fun.Config.ServerId}' 
                                 AND d4.IsOK = '0'
                                 AND d4.SimulationId IN ('{simulationIdsStr}')
                                 AND m.StationNO NOT IN (
                                     SELECT DISTINCT StationNO 
                                     FROM SoftNetLogDB.[dbo].[OperateLog] 
                                     WHERE ServerId = '{_Fun.Config.ServerId}' 
                                       AND (
                                           (OperateType LIKE '%關站%')
                                           OR (OperateType LIKE '%停工%' AND LOGDateTime >= '{fourHoursAgo:yyyy-MM-dd HH:mm:ss}')
                                       )
                                 )");

                        // 處理 DOC3stockII 資料
                        if (dtDOC3 != null)
                        {
                            foreach (DataRow dr in dtDOC3.Rows)
                            {
                                materials.Add(new
                                {
                                    NeedDate = dr.IsNull("NeedDate") ? "" : Convert.ToDateTime(dr["NeedDate"]).ToString("yyyy-MM-dd HH:mm:ss"),
                                    StationNO = dr["StationNO"].ToString(),
                                    DOCNumberNO = dr["DOCNumberNO"].ToString(),
                                    PartNO = dr["PartNO"].ToString(),
                                    PartName = dr.IsNull("PartName") ? "" : dr["PartName"].ToString(),
                                    Specification = dr.IsNull("Specification") ? "" : dr["Specification"].ToString(),
                                    QTY = dr.IsNull("QTY") ? 0 : Convert.ToDecimal(dr["QTY"]),
                                    Unit = dr.IsNull("Unit") ? "" : dr["Unit"].ToString(),
                                    StoreNO = dr.IsNull("StoreNO") ? "" : dr["StoreNO"].ToString(),
                                    StoreSpacesNO = dr.IsNull("StoreSpacesNO") ? "" : dr["StoreSpacesNO"].ToString(),
                                    Remark = dr.IsNull("Remark") ? "" : dr["Remark"].ToString()
                                });
                            }
                        }

                        // 處理 DOC4ProductionII 資料
                        if (dtDOC4 != null)
                        {
                            foreach (DataRow dr in dtDOC4.Rows)
                            {
                                materials.Add(new
                                {
                                    NeedDate = dr.IsNull("NeedDate") ? "" : Convert.ToDateTime(dr["NeedDate"]).ToString("yyyy-MM-dd HH:mm:ss"),
                                    StationNO = dr["StationNO"].ToString(),
                                    DOCNumberNO = dr["DOCNumberNO"].ToString(),
                                    PartNO = dr["PartNO"].ToString(),
                                    PartName = dr.IsNull("PartName") ? "" : dr["PartName"].ToString(),
                                    Specification = dr.IsNull("Specification") ? "" : dr["Specification"].ToString(),
                                    QTY = dr.IsNull("QTY") ? 0 : Convert.ToDecimal(dr["QTY"]),
                                    Unit = dr.IsNull("Unit") ? "" : dr["Unit"].ToString(),
                                    StoreNO = dr.IsNull("StoreNO") ? "" : dr["StoreNO"].ToString(),
                                    StoreSpacesNO = dr.IsNull("StoreSpacesNO") ? "" : dr["StoreSpacesNO"].ToString(),
                                    Remark = dr.IsNull("Remark") ? "" : dr["Remark"].ToString()
                                });
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"GetMaterialRequirementsAsync Error: {ex.Message}");
                }

                return JObject.FromObject(new
                {
                    Materials = materials
                });
            });
        }

        #endregion

        #region 功能2：生管人員戰情室

        /// <summary>
        /// 取得生產現況總覽
        /// </summary>
        public async Task<JObject> GetProductionOverviewAsync()
        {
            return await Task.Run(() =>
            {
                var todayOutput = GetProductionOutputAsync("today").Result;
                var yesterdayOutput = GetProductionOutputAsync("yesterday").Result;

                int todayTotal = todayOutput["TotalQty"].Value<int>();
                int yesterdayTotal = yesterdayOutput["TotalQty"].Value<int>();

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    // 查詢進行中工單數
                    DataTable dtActive = db.DB_GetData(
                        $"SELECT COUNT(*) as Count FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] " +
                        $"WHERE ServerId = '{_Fun.Config.ServerId}' AND EndTime IS NULL");

                    // 查詢待派工工單數
                    DataTable dtPending = db.DB_GetData(
                        $"SELECT COUNT(*) as Count FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] " +
                        $"WHERE ServerId = '{_Fun.Config.ServerId}' AND StartTime IS NULL");

                    // 查詢異常工單數（WOStatus != 0）
                    DataTable dtError = db.DB_GetData(
                        $"SELECT COUNT(*) as Count FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] " +
                        $"WHERE ServerId = '{_Fun.Config.ServerId}' AND EndTime IS NULL AND WOStatus != 0");

                    // 查詢今日完成工單數
                    DataTable dtCompleted = db.DB_GetData(
                        $"SELECT COUNT(*) as Count FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] " +
                        $"WHERE ServerId = '{_Fun.Config.ServerId}' " +
                        $"  AND CONVERT(DATE, EndTime) = CONVERT(DATE, GETDATE())");

                    int activeCount = 0;
                    if (dtActive != null && dtActive.Rows.Count > 0 && !dtActive.Rows[0].IsNull("Count"))
                    {
                        activeCount = Convert.ToInt32(dtActive.Rows[0]["Count"]);
                    }
                    int pendingCount = 0;
                    if (dtPending != null && dtPending.Rows.Count > 0 && !dtPending.Rows[0].IsNull("Count"))
                    {
                        pendingCount = Convert.ToInt32(dtPending.Rows[0]["Count"]);
                    }
                    int errorCount = 0;
                    if (dtError != null && dtError.Rows.Count > 0 && !dtError.Rows[0].IsNull("Count"))
                    {
                        errorCount = Convert.ToInt32(dtError.Rows[0]["Count"]);
                    }
                    int completedCount = 0;
                    if (dtCompleted != null && dtCompleted.Rows.Count > 0 && !dtCompleted.Rows[0].IsNull("Count"))
                    {
                        completedCount = Convert.ToInt32(dtCompleted.Rows[0]["Count"]);
                    }

                    return JObject.FromObject(new
                    {
                        TodayOutput = new
                        {
                            SemiFinished = todayOutput["SemiFinishedQty"].Value<int>(),
                            Finished = todayOutput["FinishedQty"].Value<int>(),
                            Total = todayTotal
                        },
                        YesterdayOutput = new
                        {
                            SemiFinished = yesterdayOutput["SemiFinishedQty"].Value<int>(),
                            Finished = yesterdayOutput["FinishedQty"].Value<int>(),
                            Total = yesterdayTotal
                        },
                        ActiveWorkOrders = activeCount,
                        PendingWorkOrders = pendingCount,
                        ErrorWorkOrders = errorCount,
                        CompletedWorkOrders = completedCount
                    });
                }
            });
        }

        /// <summary>
        /// 取得未來N天排程預警
        /// </summary>
        public async Task<JObject> GetScheduleWarningsAsync(int days = 7)
        {
            return await Task.Run(() =>
            {
                var warnings = new List<object>();

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DateTime endDate = DateTime.Now.AddDays(days);

                    // 查詢未來排程中的物料短缺
                    DataTable dt = db.DB_GetData(
                        $@"SELECT 
                            FORMAT(sim.StartDate, 'yyyy-MM-dd') as Date,
                            sim.SimulationId,
                            sim.NeedId,
                            sim.PartNO,
                            mat.PartName,
                            sim.NeedQTY,
                            sim.SafeQTY,
                            ISNULL(SUM(ts.QTY), 0) as StockQty
                          FROM SoftNetSYSDB.[dbo].[APS_Simulation] sim
                          LEFT JOIN SoftNetMainDB.[dbo].[Material] mat ON sim.PartNO = mat.PartNO AND sim.ServerId = mat.ServerId
                          LEFT JOIN SoftNetMainDB.[dbo].[TotalStock] ts ON sim.PartNO = ts.PartNO AND sim.ServerId = ts.ServerId
                          WHERE sim.ServerId = '{_Fun.Config.ServerId}'
                            AND sim.StartDate >= '{DateTime.Now:yyyy-MM-dd}'
                            AND sim.StartDate <= '{endDate:yyyy-MM-dd}'
                            AND sim.IsOK = 0
                          GROUP BY sim.StartDate, sim.SimulationId, sim.NeedId, sim.PartNO, mat.PartName, sim.NeedQTY, sim.SafeQTY
                          ORDER BY sim.StartDate");

                    var warningItemRows = new List<(string Date, string PartNO, string PartName, int NeedQTY, int SafeQTY, int StockQty, string RiskLevel)>();
                    if (dt != null)
                    {
                        foreach (DataRow r in dt.Rows)
                        {
                            int effectiveNeed = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, r["NeedId"].ToString(), r["SimulationId"].ToString(), Convert.ToInt32(r["NeedQTY"]));
                            int demandQTY = effectiveNeed + Convert.ToInt32(r["SafeQTY"]);
                            int stockQty = Convert.ToInt32(r["StockQty"]);
                            if (stockQty >= demandQTY) { continue; }
                            string riskLevel = stockQty < demandQTY ? "高" : (stockQty < (int)(demandQTY * 1.2) ? "中" : "正常");
                            warningItemRows.Add((r["Date"].ToString(), r["PartNO"].ToString(), r.IsNull("PartName") ? "" : r["PartName"].ToString(), effectiveNeed, Convert.ToInt32(r["SafeQTY"]), stockQty, riskLevel));
                        }
                    }

                    if (warningItemRows.Count > 0)
                    {
                        var grouped = warningItemRows.GroupBy(r => r.Date).ToList();

                        foreach (var group in grouped)
                        {
                            var items = group.ToList();
                            string riskLevel = items.Any(r => r.RiskLevel == "高") ? "高" :
                                             items.Any(r => r.RiskLevel == "中") ? "中" : "正常";

                            warnings.Add(new
                            {
                                Date = group.Key,
                                PlanWorkOrders = items.Count,
                                RiskLevel = riskLevel,
                                WarningItems = items.Select(r => new
                                {
                                    PartNO = r.PartNO,
                                    PartName = r.PartName,
                                    NeedQTY = r.NeedQTY,
                                    SafeQTY = r.SafeQTY,
                                    StockQty = r.StockQty,
                                    RiskLevel = r.RiskLevel
                                }).ToList()
                            });
                        }
                    }
                }

                return JObject.FromObject(new
                {
                    Warnings = warnings
                });
            });
        }

        /// <summary>
        /// 取得工站負荷分析
        /// </summary>
        public async Task<JObject> GetStationLoadAnalysisAsync(string dateType = "today")
        {
            return await Task.Run(() =>
            {
                var stations = new List<object>();

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DateTime targetDate = dateType == "today" 
                        ? DateTime.Now.Date 
                        : DateTime.Now.AddDays(-1).Date;

                    // 查詢所有工站
                    DataTable dtStations = db.DB_GetData(
                        $"SELECT StationNO, StationName, FactoryName, LineName " +
                        $"FROM SoftNetSYSDB.[dbo].[PP_Station] " +
                        $"WHERE ServerId = '{_Fun.Config.ServerId}'");

                    if (dtStations != null)
                    {
                        foreach (DataRow station in dtStations.Rows)
                        {
                            string stationNO = station["StationNO"].ToString();

                            // 查詢排程計畫時間（Time_TOT）
                            DataTable dtPlanned = db.DB_GetData(
                                $@"SELECT SUM(Time_TOT) as PlannedSeconds
                                  FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote]
                                  WHERE ServerId = '{_Fun.Config.ServerId}'
                                    AND StationNO = '{stationNO}'
                                    AND CONVERT(DATE, CalendarDate) = '{targetDate:yyyy-MM-dd}'");

                            int plannedSeconds = 0;
                            if (dtPlanned != null && dtPlanned.Rows.Count > 0)
                            {
                                var row = dtPlanned.Rows[0];
                                if (!row.IsNull("PlannedSeconds"))
                                {
                                    plannedSeconds = Convert.ToInt32(row["PlannedSeconds"]);
                                }
                            }

                            // 計算實際工作時間
                            int actualSeconds = CalculateActualWorkingSeconds(db, stationNO, targetDate);

                            double utilizationRate = plannedSeconds > 0
                                ? (double)actualSeconds / plannedSeconds * 100
                                : 0;

                            string status = utilizationRate >= 90 ? "過載" :
                                          utilizationRate >= 70 ? "警告" : "正常";

                            stations.Add(new
                            {
                                StationNO = stationNO,
                                StationName = station["StationName"].ToString(),
                                FactoryName = station["FactoryName"].ToString(),
                                LineName = station["LineName"].ToString(),
                                PlannedSeconds = plannedSeconds,
                                ActualSeconds = actualSeconds,
                                UtilizationRate = Math.Round(utilizationRate, 1),
                                Status = status
                            });
                        }
                    }
                }

                return JObject.FromObject(new
                {
                    Stations = stations
                });
            });
        }

        /// <summary>
        /// 取得物料庫存預警
        /// </summary>
        public async Task<JObject> GetMaterialStockWarningsAsync()
        {
            return await Task.Run(() =>
            {
                var materials = new List<object>();

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    // 查詢庫存低於安全存量的物料
                    DataTable dt = db.DB_GetData(
                        $@"SELECT 
                            mat.PartNO,
                            mat.PartName,
                            mat.SafeQTY,
                            ISNULL(SUM(ts.QTY), 0) as StockQty,
                            CASE 
                                WHEN ISNULL(SUM(ts.QTY), 0) = 0 THEN '缺料'
                                WHEN ISNULL(SUM(ts.QTY), 0) < mat.SafeQTY THEN '低於'
                                WHEN ISNULL(SUM(ts.QTY), 0) < mat.SafeQTY * 1.2 THEN '接近'
                                ELSE '正常'
                            END as WarningStatus
                          FROM SoftNetMainDB.[dbo].[Material] mat
                          LEFT JOIN SoftNetMainDB.[dbo].[TotalStock] ts ON mat.PartNO = ts.PartNO AND mat.ServerId = ts.ServerId
                          WHERE mat.ServerId = '{_Fun.Config.ServerId}'
                            AND mat.SafeQTY > 0
                          GROUP BY mat.PartNO, mat.PartName, mat.SafeQTY
                          HAVING ISNULL(SUM(ts.QTY), 0) < mat.SafeQTY * 1.2
                          ORDER BY WarningStatus DESC, mat.PartNO");

                    if (dt != null)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            string status = dr["WarningStatus"].ToString();
                            string action = status == "缺料" ? "緊急採購" :
                                          status == "低於" ? "已自動補單" :
                                          status == "接近" ? "監控中" : "-";

                            materials.Add(new
                            {
                                PartNO = dr["PartNO"].ToString(),
                                PartName = dr["PartName"].ToString(),
                                SafeQTY = Convert.ToInt32(dr["SafeQTY"]),
                                StockQty = Convert.ToInt32(dr["StockQty"]),
                                WarningStatus = status,
                                Action = action
                            });
                        }
                    }
                }

                return JObject.FromObject(new
                {
                    Materials = materials
                });
            });
        }

        /// <summary>
        /// 取得工單執行狀態
        /// </summary>
        public async Task<JObject> GetWorkOrderStatusAsync(string status = "all")
        {
            return await Task.Run(() =>
            {
                var workOrders = new List<object>();

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    string whereClause = "";
                    if (status == "active")
                        whereClause = "AND wo.EndTime IS NULL";
                    else if (status == "completed")
                        whereClause = "AND wo.EndTime IS NOT NULL";
                    else if (status == "error")
                        whereClause = "AND wo.WOStatus != 0";

                    DataTable dt = db.DB_GetData(
                        $@"SELECT 
                            wo.OrderNO,
                            wo.PartNO,
                            mat.PartName,
                            wo.Quantity,
                            wo.ActualQuantity,
                            wo.EstimatedStartTime,
                            wo.StartTime,
                            wo.EndTime,
                            CASE 
                                WHEN wo.EndTime IS NOT NULL THEN '完成'
                                WHEN wo.ActualQuantity >= wo.Quantity THEN '達標'
                                WHEN wo.StartTime IS NOT NULL AND wo.EndTime IS NULL THEN '進行中'
                                WHEN wo.EstimatedStartTime < GETDATE() AND wo.StartTime IS NULL THEN '延遲'
                                ELSE '風險'
                            END as Status
                          FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] wo
                          LEFT JOIN SoftNetMainDB.[dbo].[Material] mat ON wo.PartNO = mat.PartNO AND wo.ServerId = mat.ServerId
                          WHERE wo.ServerId = '{_Fun.Config.ServerId}'
                            {whereClause}
                          ORDER BY wo.EstimatedStartTime DESC");

                    if (dt != null)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            int planQty = dr.IsNull("Quantity") ? 0 : Convert.ToInt32(dr["Quantity"]);
                            int actualQty = dr.IsNull("ActualQuantity") ? 0 : Convert.ToInt32(dr["ActualQuantity"]);
                            double progress = planQty > 0 ? (double)actualQty / planQty * 100 : 0;

                            workOrders.Add(new
                            {
                                OrderNO = dr["OrderNO"].ToString(),
                                PartNO = dr["PartNO"].ToString(),
                                PartName = dr.IsNull("PartName") ? "" : dr["PartName"].ToString(),
                                PlanQty = planQty,
                                ActualQty = actualQty,
                                Progress = Math.Round(progress, 1),
                                Status = dr["Status"].ToString(),
                                EstimatedStartTime = dr.IsNull("EstimatedStartTime") ? "" : dr["EstimatedStartTime"].ToString(),
                                StartTime = dr.IsNull("StartTime") ? "" : dr["StartTime"].ToString(),
                                EndTime = dr.IsNull("EndTime") ? "" : dr["EndTime"].ToString()
                            });
                        }
                    }
                }

                return JObject.FromObject(new
                {
                    WorkOrders = workOrders
                });
            });
        }

        /// <summary>
        /// 取得主動干涉事件統計
        /// </summary>
        public async Task<JObject> GetInterventionEventStatsAsync(string dateType = "today")
        {
            return await Task.Run(() =>
            {
                var stats = new List<object>();

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DateTime startDate = dateType == "today" 
                        ? DateTime.Now.Date 
                        : DateTime.Now.AddDays(-1).Date;
                    DateTime endDate = startDate.AddDays(1);

                    // 統計各規則類型的觸發次數
                    DataTable dt = db.DB_GetData(
                        $@"SELECT 
                            ErrorType,
                            COUNT(*) as TotalCount,
                            SUM(CASE WHEN IsDEL = 0 THEN 1 ELSE 0 END) as PendingCount,
                            SUM(CASE WHEN IsDEL = 1 THEN 1 ELSE 0 END) as ProcessedCount
                          FROM SoftNetSYSDB.[dbo].[APS_WarningData]
                          WHERE ServerId = '{_Fun.Config.ServerId}' and ErrorType!='00' and ErrorType!='33'
                            AND WarningDate >= '{startDate:yyyy-MM-dd HH:mm:ss}'
                            AND WarningDate < '{endDate:yyyy-MM-dd HH:mm:ss}'
                          GROUP BY ErrorType");

                    if (dt != null)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            stats.Add(new
                            {
                                RuleType = dr["ErrorType"].ToString(),
                                TotalCount = Convert.ToInt32(dr["TotalCount"]),
                                ProcessedCount = Convert.ToInt32(dr["ProcessedCount"]),
                                PendingCount = Convert.ToInt32(dr["PendingCount"])
                            });
                        }
                    }
                }

                return JObject.FromObject(new
                {
                    Stats = stats
                });
            });
        }

        #endregion

        #region 功能3：高階管理戰情室

        /// <summary>
        /// 取得核心KPI總覽
        /// </summary>
        public async Task<JObject> GetCoreKPIsAsync(string period = "month")
        {
            return await Task.Run(() =>
            {
                DateTime startDate = period == "month" 
                    ? new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
                    : DateTime.Now.AddMonths(-1);
                DateTime endDate = DateTime.Now;

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    // 計算總產量（半成品+成品）
                    var todayOutput = GetProductionOutputAsync("today").Result;
                    int todayTotal = todayOutput["TotalQty"].Value<int>();

                    // 計算本月累計產量
                    int monthTotal = 0;
                    int monthSemiFinished = 0;
                    int monthFinished = 0;

                    DataTable dtStations = db.DB_GetData(
                        $"SELECT DISTINCT StationNO FROM SoftNetSYSDB.[dbo].[PP_Station] " +
                        $"WHERE ServerId = '{_Fun.Config.ServerId}'");

                    foreach (DataRow station in dtStations.Rows)
                    {
                        string stationNO = station["StationNO"].ToString();
                        var (startTime, endTime) = GetStationWorkingHours(db, stationNO, startDate);

                        DataTable dtOutput = db.DB_GetData(
                            $@"SELECT 
                                m.Class,
                                SUM(s.ProductFinishedQty) as OutputQty
                              FROM SoftNetLogDB.[dbo].[SFC_StationDetail] s
                              JOIN SoftNetMainDB.[dbo].[Material] m 
                                ON s.PartNO = m.PartNO AND s.ServerId = m.ServerId
                              WHERE s.ServerId = '{_Fun.Config.ServerId}'
                                AND s.StationNO = '{stationNO}'
                                AND s.OutputFlag = 1
                                AND m.Class IN ('4', '5')
                                AND s.LOGDateTime >= '{startDate:yyyy-MM-dd HH:mm:ss}'
                                AND s.LOGDateTime < '{endDate:yyyy-MM-dd HH:mm:ss}'
                              GROUP BY m.Class");

                        if (dtOutput != null)
                        {
                            foreach (DataRow output in dtOutput.Rows)
                            {
                                int qty = Convert.ToInt32(output["OutputQty"]);
                                if (output["Class"].ToString() == "4")
                                    monthSemiFinished += qty;
                                else if (output["Class"].ToString() == "5")
                                    monthFinished += qty;
                            }
                        }
                    }
                    monthTotal = monthSemiFinished + monthFinished;

                    // 計算達成率、產能利用率等（簡化版，實際需更複雜計算）
                    return JObject.FromObject(new
                    {
                        TotalOutput = monthTotal,
                        SemiFinishedOutput = monthSemiFinished,
                        FinishedOutput = monthFinished,
                        TodayOutput = todayTotal,
                        PlanAchievementRate = 87.0, // 需從實際資料計算
                        CapacityUtilization = 82.0, // 需從實際資料計算
                        YieldRate = 95.2, // 需從實際資料計算
                        DeliveryAchievementRate = 92.0, // 需從實際資料計算
                        OEE = 78.0, // 需從實際資料計算
                        Period = period
                    });
                }
            });
        }

        /// <summary>
        /// 取得生產效能趨勢分析
        /// </summary>
        public async Task<JObject> GetProductionTrendAsync(string granularity = "month", int months = 6)
        {
            return await Task.Run(() =>
            {
                var trends = new List<object>();
                DateTime startDate = DateTime.Now.AddMonths(-months);

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    string dateFormat = granularity == "month" ? "yyyy-MM" :
                                       granularity == "week" ? "yyyy-MM-dd" : "yyyy-MM-dd";

                    // 查詢歷史產量
                    DataTable dt = db.DB_GetData(
                        $@"SELECT 
                            FORMAT(s.LOGDateTime, '{dateFormat}') as Period,
                            m.Class,
                            SUM(s.ProductFinishedQty) as OutputQty
                          FROM SoftNetLogDB.[dbo].[SFC_StationDetail] s
                          JOIN SoftNetMainDB.[dbo].[Material] m 
                            ON s.PartNO = m.PartNO AND s.ServerId = m.ServerId
                          WHERE s.ServerId = '{_Fun.Config.ServerId}'
                            AND s.LOGDateTime >= '{startDate:yyyy-MM-dd}'
                            AND s.OutputFlag = 1
                            AND m.Class IN ('4', '5')
                          GROUP BY FORMAT(s.LOGDateTime, '{dateFormat}'), m.Class
                          ORDER BY Period");

                    if (dt != null)
                    {
                        var grouped = dt.AsEnumerable()
                            .GroupBy(r => r["Period"].ToString())
                            .ToList();

                        foreach (var group in grouped)
                        {
                            int semiFinished = group.Where(r => r["Class"].ToString() == "4")
                                                   .Sum(r => Convert.ToInt32(r["OutputQty"]));
                            int finished = group.Where(r => r["Class"].ToString() == "5")
                                                .Sum(r => Convert.ToInt32(r["OutputQty"]));

                            trends.Add(new
                            {
                                Period = group.Key,
                                SemiFinishedQty = semiFinished,
                                FinishedQty = finished,
                                TotalQty = semiFinished + finished
                            });
                        }
                    }
                }

                return JObject.FromObject(new
                {
                    Trends = trends
                });
            });
        }

        /// <summary>
        /// 取得產能利用率趨勢
        /// </summary>
        public async Task<JObject> GetCapacityUtilizationTrendAsync(string granularity = "month", int months = 6)
        {
            return await Task.Run(() =>
            {
                var trends = new List<object>();
                DateTime startDate = DateTime.Now.AddMonths(-months);

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    string dateFormat = granularity == "month" ? "yyyy-MM" :
                                       granularity == "week" ? "yyyy-MM-dd" : "yyyy-MM-dd";

                    // 查詢排程計畫時間
                    DataTable dtPlanned = db.DB_GetData(
                        $@"SELECT 
                            FORMAT(CalendarDate, '{dateFormat}') as Period,
                            StationNO,
                            SUM(Time_TOT) as PlannedSeconds
                          FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote]
                          WHERE ServerId = '{_Fun.Config.ServerId}'
                            AND CalendarDate >= '{startDate:yyyy-MM-dd}'
                          GROUP BY FORMAT(CalendarDate, '{dateFormat}'), StationNO");

                    // 查詢各工站的行事曆
                    DataTable dtStations = db.DB_GetData(
                        $"SELECT StationNO FROM SoftNetSYSDB.[dbo].[PP_Station] " +
                        $"WHERE ServerId = '{_Fun.Config.ServerId}'");

                    Dictionary<string, Dictionary<string, int>> actualSeconds = 
                        new Dictionary<string, Dictionary<string, int>>();

                    // 計算實際工作時間
                    foreach (DataRow station in dtStations.Rows)
                    {
                        string stationNO = station["StationNO"].ToString();
                        for (int i = 0; i < months; i++)
                        {
                            DateTime targetDate = startDate.AddMonths(i);
                            int seconds = CalculateActualWorkingSeconds(db, stationNO, targetDate);
                            string period = targetDate.ToString(dateFormat);

                            if (!actualSeconds.ContainsKey(period))
                                actualSeconds[period] = new Dictionary<string, int>();

                            if (!actualSeconds[period].ContainsKey(stationNO))
                                actualSeconds[period][stationNO] = 0;

                            actualSeconds[period][stationNO] += seconds;
                        }
                    }

                    // 彙總產能利用率
                    if (dtPlanned != null)
                    {
                        var grouped = dtPlanned.AsEnumerable()
                            .GroupBy(r => r["Period"].ToString())
                            .ToList();

                        foreach (var group in grouped)
                        {
                            int plannedTotal = (int)group.Sum(r => Convert.ToDecimal(r["PlannedSeconds"]));
                            int actualTotal = group.Sum(r =>
                            {
                                string period = r["Period"].ToString();
                                string stationNO = r["StationNO"].ToString();
                                return actualSeconds.ContainsKey(period) &&
                                       actualSeconds[period].ContainsKey(stationNO)
                                    ? actualSeconds[period][stationNO] : 0;
                            });

                            double utilizationRate = plannedTotal > 0
                                ? (double)actualTotal / plannedTotal * 100
                                : 0;

                            trends.Add(new
                            {
                                Period = group.Key,
                                PlannedSeconds = plannedTotal,
                                ActualSeconds = actualTotal,
                                UtilizationRate = Math.Round(utilizationRate, 1)
                            });
                        }
                    }
                }

                return JObject.FromObject(new
                {
                    Trends = trends
                });
            });
        }

        /// <summary>
        /// 取得進步/退步分析
        /// </summary>
        public async Task<JObject> GetProgressAnalysisAsync(string period = "month")
        {
            return await Task.Run(() =>
            {
                // 簡化版，實際需比較上月同期資料
                return JObject.FromObject(new
                {
                    Improvements = new[]
                    {
                        new { Item = "計畫達成率", Current = 87.0, Previous = 84.0, Change = 3.1, Trend = "上升" },
                        new { Item = "產能利用率", Current = 82.0, Previous = 79.5, Change = 2.5, Trend = "上升" },
                        new { Item = "良率", Current = 95.2, Previous = 94.4, Change = 0.8, Trend = "上升" }
                    },
                    Concerns = new[]
                    {
                        new { Item = "交期達成率", Current = 92.0, Previous = 89.7, Change = 2.3, Trend = "上升", Note = "但仍低於目標95%" },
                        new { Item = "OEE", Current = 78.0, Previous = 76.5, Change = 1.5, Trend = "上升", Note = "但仍低於目標85%" }
                    }
                });
            });
        }

        /// <summary>
        /// 取得成本效益分析
        /// </summary>
        public async Task<JObject> GetCostBenefitAnalysisAsync(string period = "month")
        {
            return await Task.Run(() =>
            {
                // 簡化版，實際需從財務資料計算
                return JObject.FromObject(new
                {
                    TotalOutput = 28500,
                    UnitCost = 12.5,
                    TotalCost = 356250,
                    EfficiencyImprovement = 3.1,
                    Period = period
                });
            });
        }

        /// <summary>
        /// 取得目標達成狀況
        /// </summary>
        public async Task<JObject> GetTargetAchievementAsync(string period = "month")
        {
            return await Task.Run(() =>
            {
                var kpis = new[]
                {
                    new { Item = "計畫達成率", Target = 90.0, Actual = 87.0, AchievementRate = 96.7, Status = "接近" },
                    new { Item = "產能利用率", Target = 85.0, Actual = 82.0, AchievementRate = 96.5, Status = "接近" },
                    new { Item = "良率", Target = 96.0, Actual = 95.2, AchievementRate = 99.2, Status = "接近" },
                    new { Item = "交期達成率", Target = 95.0, Actual = 92.0, AchievementRate = 96.8, Status = "接近" },
                    new { Item = "OEE", Target = 85.0, Actual = 78.0, AchievementRate = 91.8, Status = "未達" },
                    new { Item = "異常事件", Target = 10.0, Actual = 15.0, AchievementRate = 66.7, Status = "未達" }
                };
                
                return JObject.FromObject(new
                {
                    KPIs = kpis
                });
            });
        }

        /// <summary>
        /// 取得主動干涉次數（前N天）
        /// </summary>
        public async Task<JObject> GetInterventionCountAsync(int days = 30)
        {
            return await Task.Run(() =>
            {
                int totalCount = 0;

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DateTime startDate = DateTime.Now.AddDays(-days);

                    // 統計 APS_WarningData
                    DataTable dtWarning = db.DB_GetData(
                        $"SELECT COUNT(*) as Count FROM SoftNetSYSDB.[dbo].[APS_WarningData] " +
                        $"WHERE ServerId = '{_Fun.Config.ServerId}'  and ErrorType!='00' and ErrorType!='33'" +
                        $"  AND WarningDate >= '{startDate:yyyy-MM-dd HH:mm:ss}' " +
                        $"  AND IsDEL = 0");

                    // 統計 APS_Simulation_ErrorData (ActionType=0)
                    string tmp = "";
                    if (!_Fun.Config.DOCInvoice_Expiry_SendError)
                    { tmp = $"{tmp} and ErrorType!='06' and ErrorType!='07' and ErrorType!='08' and ErrorType!='09' and ErrorType!='10'"; }

                    DataTable dtError = db.DB_GetData(
                        $"SELECT COUNT(*) as Count FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] " +
                        $"WHERE ServerId = '{_Fun.Config.ServerId}' " +
                        $"  AND ActionLogDate >= '{startDate:yyyy-MM-dd HH:mm:ss}' {tmp}");

                    int warningCount = 0;
                    if (dtWarning != null && dtWarning.Rows.Count > 0 && !dtWarning.Rows[0].IsNull("Count"))
                    {
                        warningCount = Convert.ToInt32(dtWarning.Rows[0]["Count"]);
                    }
                    int errorCount = 0;
                    if (dtError != null && dtError.Rows.Count > 0 && !dtError.Rows[0].IsNull("Count"))
                    {
                        errorCount = Convert.ToInt32(dtError.Rows[0]["Count"]);
                    }

                    totalCount = warningCount + errorCount;
                }

                return JObject.FromObject(new
                {
                    TotalCount = totalCount,
                    Days = days
                });
            });
        }

        /// <summary>
        /// 取得主動干涉明細
        /// </summary>
        public async Task<JObject> GetInterventionDetailAsync(int days = 30)
        {
            return await Task.Run(() =>
            {
                var details = new List<object>();

                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DateTime startDate = DateTime.Now.AddDays(-days);
                    string tmp = "";
                    if (!_Fun.Config.DOCInvoice_Expiry_SendError)
                    { tmp = $"{tmp} and ErrorType!='06' and ErrorType!='07' and ErrorType!='08' and ErrorType!='09' and ErrorType!='10'"; }
                    // 查詢 APS_WarningData
                    DataTable dtWarning = db.DB_GetData(
                        $@"SELECT 
                            WarningDate as EventDate,
                            ErrorType,
                            StationNO,
                            DOCNumberNO,
                            PartNO,
                            DATA_Remark
                          FROM SoftNetSYSDB.[dbo].[APS_WarningData]
                          WHERE ServerId = '{_Fun.Config.ServerId}' and ErrorType!='00' and ErrorType!='33'
                            AND WarningDate >= '{startDate:yyyy-MM-dd HH:mm:ss}'
                            AND IsDEL = 0
                          
                          UNION ALL
                          
                          SELECT 
                            ActionLogDate as EventDate,
                            ErrorType,
                            '' as StationNO,
                            '' as DOCNumberNO,
                            '' as PartNO,
                            '' as DATA_Remark
                          FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData]
                          WHERE ServerId = '{_Fun.Config.ServerId}'
                            AND ActionType = '0'
                            AND ActionLogDate >= '{startDate:yyyy-MM-dd HH:mm:ss}' {tmp}
                          
                          ORDER BY EventDate DESC");

                    if (dtWarning != null)
                    {
                        foreach (DataRow dr in dtWarning.Rows)
                        {
                            details.Add(new
                            {
                                EventDate = dr["EventDate"].ToString(),
                                ErrorType = dr["ErrorType"].ToString(),
                                StationNO = dr.IsNull("StationNO") ? "" : dr["StationNO"].ToString(),
                                DOCNumberNO = dr.IsNull("DOCNumberNO") ? "" : dr["DOCNumberNO"].ToString(),
                                PartNO = dr.IsNull("PartNO") ? "" : dr["PartNO"].ToString(),
                                Remark = dr.IsNull("DATA_Remark") ? "" : dr["DATA_Remark"].ToString()
                            });
                        }
                    }
                }

                return JObject.FromObject(new
                {
                    Details = details
                });
            });
        }

        #endregion
    }
}
