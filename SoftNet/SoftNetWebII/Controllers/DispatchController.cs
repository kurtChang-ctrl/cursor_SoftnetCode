using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Models;
using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class DispatchController : ApiCtrl
    {
        public ActionResult Read()
        {
            return View();
        }
        public ActionResult Read2()
        {

            return View();
        }
        public void UpdateEvent(int id, string NewEventStart, string NewEventEnd)
        {
            DiaryEvent.UpdateDiaryEvent(id, NewEventStart, NewEventEnd);
        }


        public bool SaveEvent(string Title, string NewEventDate, string NewEventTime, string NewEventDuration)
        {
            return DiaryEvent.CreateNewEvent(Title, NewEventDate, NewEventTime, NewEventDuration);
        }

        public JsonResult GetDiarySummary(double start, double end)
        {
            var ApptListForDate = DiaryEvent.LoadAppointmentSummaryInDateRange(start, end);
            var eventList = from e in ApptListForDate
                            select new
                            {
                                id = e.id,
                                name = e.name,
                                content = e.content,
                                url = e.url,
                                imgUrl = e.imgUrl,
                                startDate = e.startDate,
                                endDate = e.endDate,
                                textColor = e.textColor,
                            };
            var rows = eventList.ToArray();
            return Json(rows);
        }

        public JsonResult GetDiaryEvents(DateTime start,DateTime end)
        {

            var ApptListForDate = DiaryEvent.GetDateRange(start, end);
            var eventList = from e in ApptListForDate
                            select new
                            {
                                id = e.id,
                                name = e.name,
                                content = e.content,
                                url = e.url,
                                imgUrl = e.imgUrl,
                                startDate = e.startDate,
                                endDate = e.endDate,
                                textColor = e.textColor,
                            };
            var rows = eventList.ToArray();
            return Json(rows);
        }

        [HttpPost]
        public JsonResult SaveNewEvent(string StationNO, string WO, string STime, string CalendarNO)
        {
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow dr = null;
                char type = '0';//1=無工單
                DateTime sTime = Convert.ToDateTime(STime);
                string partNO = "";
                int needTime = 1800;
                if (WO.Trim() != "")
                {
                    dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{WO}'");
                    if (dr != null)
                    {
                        CalendarNO = dr["CalendarName"].ToString();

                        //@要找wo在該站bom實際料與需求數量

                        partNO= dr["PartNO"].ToString();
                        int qty = int.Parse(dr["Quantity"].ToString());
                        dr = db.DB_GetFirstDataByDataRow($"select PartNO,StationNO,ROUND(AVG(AverageCycleTime),0) as ACT,ROUND(AVG(EfficientCycleTime),0) as ECT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}' and StationNO='{StationNO}' group by PartNO,StationNO");
                        if (dr != null)
                        {
                            int ect = 1;
                            if (!dr.IsNull("ECT") && dr["ECT"].ToString() != "") { ect = int.Parse(dr["ECT"].ToString()); }
                            needTime = qty * ect;
                        }
                    }
                    else { type = '1'; }
                }
                else { type = '1'; }
                if (type == '1')
                {
                    #region
                    if (CalendarNO.Trim() != "")
                    {
                        dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and CalendarName='{CalendarNO}' and Holiday='{sTime.ToString("yyyy-MM-dd")}'");
                        if (dr != null)
                        {
                            bool stDefalut = false;
                            string tmp = "Shift_Morning";
                            DateTime stime = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                            if (bool.Parse(dr["Flag_Morning"].ToString()))
                            {
                                #region
                                DateTime t = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                if (stime > t) { t.AddDays(1); }
                                if (t >= sTime)
                                {
                                    DateTime et = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                    DateTime st = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                    int typeTotalTime = TimeCompute2Seconds(st, et);
                                    st = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                    typeTotalTime += TimeCompute2Seconds(st, t);
                                    stDefalut = true;
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] VALUES ('{StationNO}','{sTime.ToString("yyyy-MM-dd HH:mm:ss")}','','00000000-0000-0000-0000-000000000000','1','0','0','0',{typeTotalTime.ToString()},0,0,0,'')");
                                }
                                #endregion
                            }
                            if (bool.Parse(dr["Flag_Afternoon"].ToString()))
                            {
                                #region
                                tmp = "Shift_Afternoon";
                                DateTime t = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                if (stime > t) { t.AddDays(1); }
                                if (t >= sTime)
                                {
                                    DateTime et = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                    DateTime st = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                    int typeTotalTime = TimeCompute2Seconds(st, et);
                                    st = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                    typeTotalTime += TimeCompute2Seconds(st, t);
                                    if (stDefalut)
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] VALUES ('{StationNO}','{st.ToString("yyyy-MM-dd HH:mm:ss")}','','00000000-0000-0000-0000-000000000000','0','1','0','0',0,{typeTotalTime.ToString()},0,0,'')");
                                    }
                                    else
                                    {
                                        stDefalut = true;
                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] VALUES ('{StationNO}','{sTime.ToString("yyyy-MM-dd HH:mm:ss")}','','00000000-0000-0000-0000-000000000000','0','1','0','0',0,{typeTotalTime.ToString()},0,0,'')");
                                    }
                                }
                                #endregion
                            }
                            if (bool.Parse(dr["Flag_Night"].ToString()))
                            {
                                #region
                                tmp = "Shift_Night";
                                DateTime t = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                if (stime > t) { t.AddDays(1); }
                                if (t >= sTime)
                                {
                                    DateTime et = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                    DateTime st = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                    int typeTotalTime = TimeCompute2Seconds(st, et);
                                    st = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                    typeTotalTime += TimeCompute2Seconds(st, t);
                                    if (stDefalut)
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] VALUES ('{StationNO}','{st.ToString("yyyy-MM-dd HH:mm:ss")}','','00000000-0000-0000-0000-000000000000','0','0','1','0',0,0,{typeTotalTime.ToString()},0,'')");
                                    }
                                    else
                                    {
                                        stDefalut = true;
                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] VALUES ('{StationNO}','{sTime.ToString("yyyy-MM-dd HH:mm:ss")}','','00000000-0000-0000-0000-000000000000','0','0','1','0',0,0,{typeTotalTime.ToString()},0,'')");
                                    }
                                }
                                #endregion
                            }
                            if (bool.Parse(dr["Flag_Graveyard"].ToString()))
                            {
                                #region
                                tmp = "Shift_Graveyard";
                                DateTime t = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                if (stime > t) { t.AddDays(1); }
                                if (t >= sTime)
                                {
                                    DateTime et = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                    DateTime st = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                    int typeTotalTime = TimeCompute2Seconds(st, et);
                                    st = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr[tmp].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr[tmp].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                    typeTotalTime += TimeCompute2Seconds(st, t);
                                    if (stDefalut)
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] VALUES ('{StationNO}','{st.ToString("yyyy-MM-dd HH:mm:ss")}','','00000000-0000-0000-0000-000000000000','0','0','0','1',0,0,0,{typeTotalTime.ToString()},'')");
                                    }
                                    else
                                    {
                                        stDefalut = true;
                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] VALUES ('{StationNO}','{sTime.ToString("yyyy-MM-dd HH:mm:ss")}','','00000000-0000-0000-0000-000000000000','0','0','0','1',0,0,0,{typeTotalTime.ToString()},'')");
                                    }
                                }
                                #endregion
                            }
                        }
                    }
                    #endregion
                }
                else
                {
                    GetFinallyDate(db, "", "00000000-0000-0000-0000-000000000000", "站", StationNO, partNO, "", "", "", CalendarNO, sTime, needTime);
                }
            }
            return Json("true");
        }
        private DateTime GetFinallyDate(DBADO db, string needId, string simulationId, string mathType, string staionNO, string partNO, string partClass, string needQTY, string safeQTY, string calendarName, DateTime intime, int times)
        {
            bool fristRUN = true;
            DateTime stime = DateTime.Now;
            if (times <= 0) { return stime.AddSeconds(60); }
            DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{calendarName}' and Holiday>='{intime.ToString("yyyy-MM-dd")}' order by Holiday");
            DateTime etime = DateTime.Now;
            DateTime stime2 = DateTime.Now;
            int reTime = times;
            string sql = "";
            foreach (DataRow dr in dt.Rows)
            {
                if (Convert.ToDateTime(dr["Holiday"]).ToString("yyyy-MM-dd") != intime.ToString("yyyy-MM-dd"))
                {
                    if (Convert.ToDateTime(dr["Holiday"]) < intime) { break; }
                    else
                    {
                        stime2 = Convert.ToDateTime(dr["Holiday"]);
                        intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, 0, 0, 0);
                    }
                }

                DataRow dr_WorkTimeNote = null;

                #region Flag_Morning
                if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                {
                    #region 取得工作時間
                    string[] Shife_Morning = dr["Shift_Morning"].ToString().Trim().Split(',');
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                    stime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                    if (fristRUN && stime >= intime) { fristRUN = false; }
                    int typeTotalTime = TimeCompute2Seconds(stime, etime);
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                    typeTotalTime += TimeCompute2Seconds(stime2, etime);
                    #endregion
                    if (!fristRUN)
                    {
                        if (mathType == "站")
                        {
                            //###???要計算合併站的問題
                            dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time1_C) as TOT from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where CONVERT(varchar(100), CalendarDate, 23)='{intime.ToString("yyyy-MM-dd")}' and StationNO='{staionNO}'");
                            if (dr_WorkTimeNote != null && dr_WorkTimeNote["TOT"].ToString() != "")
                            {
                                #region 已有其他排程, 與其他合計
                                if (times >= (typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString())))
                                {
                                    times -= (typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString()));
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type1,Time1_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime2.Hour}:{stime2.Minute}:59','{needId}','{simulationId}','1',{(typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString()))})"))
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY,staionNO); }
                                }
                                else
                                {
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type1,Time1_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime2.Hour}:{stime2.Minute}:59','{needId}','{simulationId}','1',{times.ToString()})"))
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                    return stime.AddSeconds((Math.Abs(times) + 60));
                                }
                                #endregion
                            }
                            else
                            {
                                #region INSERT
                                if (times >= typeTotalTime)
                                {
                                    times -= typeTotalTime;
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type1,Time1_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{typeTotalTime.ToString()})")) 
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                }
                                else
                                {
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type1,Time1_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{times.ToString()})")) 
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                    return stime.AddSeconds((Math.Abs(times) + 60));
                                }
                                #endregion
                            }
                        }
                        else
                        {
                            #region 料
                            if (times > typeTotalTime) { times -= typeTotalTime; }
                            else
                            {
                                GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY,""); return stime.AddSeconds((Math.Abs(times) + 60));
                            }
                            #endregion
                        }
                    }
                }
                #endregion

                #region Flag_Afternoon
                if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                {
                    #region 取得工作時間
                    string[] Shife_Morning = dr["Shift_Afternoon"].ToString().Trim().Split(',');
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                    stime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                    if (fristRUN && stime >= intime) { fristRUN = false; }
                    int typeTotalTime = TimeCompute2Seconds(stime, etime);
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                    typeTotalTime += TimeCompute2Seconds(stime2, etime);
                    #endregion
                    if (!fristRUN)
                    {
                        if (mathType == "站")
                        {
                            dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time2_C) as TOT from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{staionNO}' and CONVERT(varchar(100), CalendarDate, 23)='{intime.ToString("yyyy-MM-dd")}'");
                            if (dr_WorkTimeNote != null && dr_WorkTimeNote["TOT"].ToString() != "")
                            {
                                #region 已有其他排程, 與其他合計
                                if (times >= (typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString())))
                                {
                                    times -= (typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString()));
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type2,Time2_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{(typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString()))})"))
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                }
                                else
                                {
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type2,Time2_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{times.ToString()})"))
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                    return stime.AddSeconds((Math.Abs(times) + 60));
                                }
                                #endregion
                            }
                            else
                            {
                                #region INSERT
                                if (times >= typeTotalTime)
                                {
                                    times -= typeTotalTime;
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type2,Time2_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{typeTotalTime.ToString()})"))
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                }
                                else
                                {
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type2,Time2_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{times.ToString()})"))
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                    return stime.AddSeconds((Math.Abs(times) + 60));
                                }
                                #endregion
                            }
                        }
                        else
                        {
                            #region 料
                            if (times > typeTotalTime) { times -= typeTotalTime; }
                            else
                            {
                                GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY,"");
                                return stime.AddSeconds((Math.Abs(times) + 60));
                            }
                            #endregion
                        }
                    }
                }
                #endregion

                #region Flag_Night
                if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                {
                    #region 取得工作時間
                    string[] Shife_Morning = dr["Shift_Night"].ToString().Trim().Split(',');
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                    stime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                    if (fristRUN && stime >= intime) { fristRUN = false; }
                    int typeTotalTime = TimeCompute2Seconds(stime, etime);
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                    typeTotalTime += TimeCompute2Seconds(stime2, etime);
                    #endregion
                    if (!fristRUN)
                    {
                        if (mathType == "站")
                        {
                            dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time3_C) as TOT from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{staionNO}' and CONVERT(varchar(100), CalendarDate, 23)='{intime.ToString("yyyy-MM-dd")}'");
                            if (dr_WorkTimeNote != null && dr_WorkTimeNote["TOT"].ToString() != "")
                            {
                                #region 已有其他排程, 與其他合計
                                dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time3_C) as TOT from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{staionNO}' and CONVERT(varchar(100), CalendarDate, 23)='{intime.ToString("yyyy-MM-dd")}'");
                                if (times >= (typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString())))
                                {
                                    times -= (typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString()));
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type3,Time3_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{(typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString()))})")) 
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                }
                                else
                                {
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type3,Time3_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{times.ToString()})"))
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                    return stime.AddSeconds((Math.Abs(times) + 60));
                                }
                                #endregion
                            }
                            else
                            {
                                #region INSERT
                                if (times >= typeTotalTime)
                                {
                                    times -= typeTotalTime;
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type3,Time3_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{typeTotalTime.ToString()})")) 
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                }
                                else
                                {
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type3,Time3_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{times.ToString()})")) 
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                    return stime.AddSeconds((Math.Abs(times) + 60));
                                }
                                #endregion
                            }
                        }
                        else
                        {
                            #region 料
                            if (times > typeTotalTime) { times -= typeTotalTime; }
                            else
                            {
                                GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY,"");
                                return stime.AddSeconds((Math.Abs(times) + 60));
                            }
                            #endregion
                        }
                    }
                }
                #endregion

                #region Flag_Graveyard
                if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                {
                    #region 取得工作時間
                    string[] Shife_Morning = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                    stime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                    if (fristRUN && stime >= intime) { fristRUN = false; }

                    int typeTotalTime = TimeCompute2Seconds(stime, etime);
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                    typeTotalTime += TimeCompute2Seconds(stime2, etime);
                    #endregion
                    if (!fristRUN)
                    {
                        if (mathType == "站")
                        {
                            dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time4_C) as TOT from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{staionNO}' and CONVERT(varchar(100), CalendarDate, 23)='{intime.ToString("yyyy-MM-dd")}'");
                            if (dr_WorkTimeNote != null && dr_WorkTimeNote["TOT"].ToString() != "")
                            {
                                #region 已有其他排程, 與其他合計
                                dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time4_C) as TOT from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{staionNO}' and CONVERT(varchar(100), CalendarDate, 23)='{intime.ToString("yyyy-MM-dd")}'");
                                if (times >= (typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString())))
                                {
                                    times -= (typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString()));
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type4,Time4_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{(typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString()))})"))
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                }
                                else
                                {
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type4,Time4_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{times.ToString()})")) 
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                    return stime.AddSeconds((Math.Abs(times) + 60));
                                }
                                #endregion
                            }
                            else
                            {
                                #region INSERT
                                if (times >= typeTotalTime)
                                {
                                    times -= typeTotalTime;
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type4,Time4_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{typeTotalTime.ToString()})")) 
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                }
                                else
                                {
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO,CalendarDate,NeedId,SimulationId,Type4,Time4_C) VALUES ('{_Fun.Config.ServerId}','{staionNO}','{intime.ToString("yyyy-MM-dd")} {stime.Hour}:{stime.Minute}:59','{needId}','{simulationId}','1',{times.ToString()})"))
                                    { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY, staionNO); }
                                    return stime.AddSeconds((Math.Abs(times) + 60));
                                }
                                #endregion
                            }
                        }
                        else
                        {
                            #region 料
                            if (times > typeTotalTime) { times -= typeTotalTime; }
                            else
                            {
                                GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, stime, partClass, needId, simulationId, needQTY, safeQTY,"");
                                return stime.AddSeconds((Math.Abs(times) + 60));
                            }
                            #endregion
                        }
                    }
                }
                #endregion
            }
            return stime.AddSeconds((Math.Abs(times) + 120));
        }
        private bool GetFinallyDate_insert_APS_PartNOTimeNote(DBADO db,string partNO, DateTime intime, DateTime stime, string partClass, string needId, string simulationId, string needQTY, string safeQTY,string stationNO)
        {
            /*
            //###??? 暫時寫死提前10分, 此處將來改在config檔
            int store_OpenDOC_AdvanceTime = -10;

            DateTime tmp_date = new DateTime(intime.Year, intime.Month, intime.Day, stime.Hour, stime.Minute, 59);
            tmp_date.AddMinutes(store_OpenDOC_AdvanceTime);

            //string mfData = "NULL";
            //DataRow dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"SELECT APS_Default_MFNO FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and APS_Default_MFNO is not null and APS_Default_MFNO!=''");
            //if (dr_WorkTimeNote != null) { mfData = $"'{dr_WorkTimeNote["APS_Default_MFNO"].ToString()}'"; }
            string macID = "";
            if (stationNO != "")
            {
                DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT Config_macID FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}'");
                if (tmp != null) { macID = tmp["Config_macID"].ToString(); }
            }
            //db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] (PartNO,CalendarDate,Class,NeedId,SimulationId,NeedQTY,macID) VALUES ('{partNO}','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss")}','{partClass}','{needId}','{simulationId}',{needQTY},'{macID}')");
            */
            return true;
        }


        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private DOCService EditService()
        {
            return new DOCService(Ctrl);
        }

        [HttpPost]
        public async Task<ContentResult> GetUpdJson(string key)
        {
            return JsonToCnt(await EditService().GetUpdJsonAsync(key));
        }

        [HttpPost]
        public async Task<ContentResult> GetViewJson(string key)
        {
            return JsonToCnt(await EditService().GetViewJsonAsync(key));
        }

        public async Task<JsonResult> Create(string json)
        {
            return Json(await EditService().CreateAsync(_Str.ToJson(json)));
        }

        public async Task<JsonResult> Update(string key, string json)
        {
            return Json(await EditService().UpdateAsync(key, _Str.ToJson(json)));
        }

        public async Task<JsonResult> Delete(string key)
        {
            return Json(await EditService().DeleteAsync(key));
        }

        private int TimeCompute2Seconds(DateTime start, DateTime end)
        {
            int cycleTime = 0;
            if (start > end)
            {
                end = end.AddDays(1);
            }
            TimeSpan ts = new TimeSpan(end.Ticks - start.Ticks);
            if (ts.TotalSeconds > 0)
            { cycleTime = (int)ts.TotalSeconds; }
            return cycleTime;
        }

    }
}
