using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using Microsoft.AspNetCore.Mvc;
using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using System;
using System.Linq;
using Newtonsoft.Json.Linq;
using BaseWeb.Models;
using DocumentFormat.OpenXml.ExtendedProperties;
using Base;

namespace SoftNetWebII.Controllers
{
    public class CIRController : ApiCtrl
    {
        //private SocketClientService _SNsocket = null;
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public CIRController(SNWebSocketService websocket, SFC_Common sfc_Common)
        {
            if (_WebSocket == null)
            {
                _WebSocket = websocket;
            }
            if (_SFC_Common == null)
            {
                _SFC_Common = sfc_Common;
            }
        }

        public ActionResult APSEventLog()
        {
            return View();
        }
        public ActionResult APSErrorData()
        {
            return View();
        }

        public ActionResult Read1()
        {
            List<IdStrDto> stationNO = new List<IdStrDto>();

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        stationNO.Add(new IdStrDto(dr["StationNO"].ToString(), dr["StationName"].ToString()));
                    }
                }
            }
            ViewBag.StationNOData = stationNO;
            return View();
        }
        public ActionResult Read2()
        {
            List<IdStrDto> stationNO = new List<IdStrDto>();

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        stationNO.Add(new IdStrDto(dr["StationNO"].ToString(), dr["StationName"].ToString()));
                    }
                }
            }
            ViewBag.StationNOData = stationNO;

            return View();
        }
        public ActionResult Read3()
        {

            List<IdStrDto> stationNO = new List<IdStrDto>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                /*
                DataTable tmp_dt0 = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='00' and CalendarName='2021行事曆' and Holiday>='2024-01-01'");
                if (tmp_dt0 != null && tmp_dt0.Rows.Count > 0)
                {
                    foreach (DataRow dr in tmp_dt0.Rows)
                    {
                        db.DB_SetData(@$"INSERT INTO SoftNetSYSDB.[dbo].[PP_HolidayCalendar] ([ServerId],[CalendarName],[Holiday],[Flag_Morning],[Flag_Afternoon],[Flag_Night],[Flag_Graveyard],[Shift_Morning],[Shift_Afternoon],[Shift_Night],[Shift_Graveyard])
                                         VALUES ('03','2021行事曆','{Convert.ToDateTime(dr["Holiday"]).ToString("yyyy-MM-dd")}','1','1','0','0','{dr["Shift_Morning"].ToString()}','{dr["Shift_Afternoon"].ToString()}','{dr["Shift_Night"].ToString()}','{dr["Shift_Graveyard"].ToString()}')");
                    }
                }
                */

                string fisrtStation = "";
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    fisrtStation = tmp_dt.Rows[0]["CalendarName"].ToString();
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        stationNO.Add(new IdStrDto(dr["StationNO"].ToString(), dr["StationName"].ToString()));
                    }
                }

            }
            ViewBag.StationNOData = stationNO;
            return View();
        }
        public ActionResult Read4()
        {
            List<IdStrDto> stationNO = new List<IdStrDto>();
            float dis_totAAA = 0;//總成長率
            float dis_totBBB = 0;//總可提升率

            float totAAA = 0;
            float totBBB = 0;

            float aaa_log = 0;
            float C1Time = 0.0f;
            float A1Time = 0.0f;
            float C2Time = 0.0f;
            float A2Time = 0.0f;
            float C3Time = 0.0f;
            float A3Time = 0.0f;
            float C4Time = 0.0f;
            float A4Time = 0.0f;

            int count = 0;
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' order by FactoryName,LineName,StationNO");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    DataRow rmp_dr2 = null;
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        stationNO.Add(new IdStrDto(dr["StationNO"].ToString(), dr["StationName"].ToString()));
                        totBBB = 0;

                        int tmp_week = (int)DateTime.Now.DayOfWeek;
                        DateTime dtime = DateTime.Now.AddDays(-tmp_week - 1);
                        DateTime stime;
                        for (int i = 1; i <= 5; ++i)
                        {
                            stime = dtime.AddDays(-6);

                            #region 取得週 aaa生產能量=有效CT/實際CT bbb可提升率=最佳CT/實際CT
                            rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select (sum((EditFinishedQty+EditFailedQty)*CycleTime)/sum((EditFinishedQty+EditFailedQty))) as ACT,(sum((EditFinishedQty+EditFailedQty)*ECT)/sum((EditFinishedQty+EditFailedQty))) as BCT,(sum((EditFinishedQty+EditFailedQty)*LowerCT)/sum((EditFinishedQty+EditFailedQty))) as CCT from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] 
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTimeID, 111)>='{stime.ToString("yyyy/MM/dd")}' and CONVERT(varchar(100), LOGDateTimeID, 111)<='{dtime.ToString("yyyy/MM/dd")}' and CycleTime!=0 and ECT!=0 and LowerCT!=0 and (LowerCT*0.5)< CycleTime");
                            if (rmp_dr2 != null)
                            {
                                if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("BCT")) { totAAA += (float.Parse(rmp_dr2["BCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100; }
                                if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("CCT")) 
                                {
                                    if ((float.Parse(rmp_dr2["CCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) > 0) { count += 1; }
                                    totBBB += (float.Parse(rmp_dr2["CCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100;
                                }
                            }
                            #endregion


                            switch (i)
                            {
                                case 1:
                                    A1Time = aaa_log;//生產能量
                                    break;
                                case 2:
                                    A2Time = aaa_log;//生產能量
                                    if (aaa_log != 0 && A1Time != 0)
                                    { C1Time = aaa_log / A1Time; }
                                    break;
                                case 3:
                                    A3Time = aaa_log;//生產能量
                                    if (aaa_log != 0 && A2Time != 0)
                                    { C2Time = aaa_log / A2Time; }
                                    break;
                                case 4:
                                    A4Time = aaa_log;//生產能量
                                    if (aaa_log != 0 && A3Time != 0)
                                    { C3Time = aaa_log / A3Time; }
                                    break;
                                case 5:
                                    if (aaa_log != 0 && A4Time != 0)
                                    { C4Time = aaa_log / A4Time; }
                                    break;
                            }
                            dtime = stime.AddDays(-1);
                        }
                        dis_totAAA += (C1Time + C2Time + C3Time + C4Time) / 4;
                        dis_totBBB += (totBBB / 4);
                    }
                    if (count > 0) { dis_totBBB = dis_totBBB/count; }
                }
            }
            ViewBag.Dis_totAAA = dis_totAAA;
            //###???暫時寫死
            ViewBag.Dis_totAAA = 21.94;
            ViewBag.StationNOData = stationNO;

            return View();
        }
        public ActionResult Read5()
        {
            List<IdStrDto> stationNO = new List<IdStrDto>();

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        stationNO.Add(new IdStrDto(dr["StationNO"].ToString(), dr["StationName"].ToString()));
                    }
                }
            }
            ViewBag.StationNOData = stationNO;

            return View();
        }
        public ActionResult Read6()
        {
            List<IdStrDto> stationNO = new List<IdStrDto>();

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        stationNO.Add(new IdStrDto(dr["StationNO"].ToString(), dr["StationName"].ToString()));
                    }
                }
            }
            ViewBag.StationNOData = stationNO;

            return View();
        }
        public ActionResult Read7()
        {
            List<IdStrDto> stationNO = new List<IdStrDto>();

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        stationNO.Add(new IdStrDto(dr["StationNO"].ToString(), dr["StationName"].ToString()));
                    }
                }
            }
            ViewBag.StationNOData = stationNO;

            return View();
        }


        [HttpPost]
        public string ShowStation1WeekDetail_7(string key)
        {
            string re = "無資料.";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DateTime dtime = DateTime.Now;
                DateTime sTime = dtime.AddDays(-1);

                List<string> returnViewBag = new List<string>();
                DataTable tmp_dt2 = db.DB_GetData($"SELECT top 30 * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday<='{sTime.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{_Fun.Config.DefaultCalendarName}' order by CalendarName,Holiday desc");
                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                {
                    foreach (DataRow d in tmp_dt2.Rows)
                    {
                        returnViewBag.Add(Convert.ToDateTime(d["Holiday"]).ToString("yyyy/MM/dd"));
                    }
                }
                tmp_dt2 = db.DB_GetData($@"SELECT  a.*,(select DOCNumberNO from SoftNetSYSDB.[dbo].[APS_Simulation] as b where a.SimulationId=b.SimulationId) as DOCNumberNO FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] as a
                            where a.ServerId='{_Fun.Config.ServerId}' and a.OP_NO like '%{key}%' and CONVERT(varchar(100), a.LOGDateTimeID, 111)>='{returnViewBag[29]}' and CONVERT(varchar(100), a.LOGDateTimeID, 111)<='{returnViewBag[0]}' order by a.OP_NO,a.LOGDateTime,a.SimulationId,a.LOGDateTimeID");
                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                {
                    re = $"<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                    re = $"{re}<tr><td>紀錄日期</td><td>報工時差</td><td>工單編號</td><td>料件資料</td><td>報工人員</td><td>OK數量</td><td>Fail數量</td><td>當下實際CT</td><td>當下有效CT</td><td>當下最佳CT</td><td>當下最差CT</td></tr>";
                    DataRow dr_tmp = null;
                    string ReportTime = "";
                    string opNO = "";
                    string CT = "";
                    string ECT = "";
                    string LowerCT = "";
                    string UpperCT = "";
                    string material = "";
                    foreach (DataRow d in tmp_dt2.Rows)
                    {
                        ReportTime = "0:0";
                        opNO = "";
                        CT = "";
                        ECT = "";
                        LowerCT = "";
                        UpperCT = "";
                        material = "";
                        if (int.Parse(d["ReportTime"].ToString()) > 0)
                        {
                            TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(d["ReportTime"].ToString()));
                            ReportTime = $"{(int)standardTime_DIS.TotalHours}時{standardTime_DIS.Minutes}分";
                        }
                        if (d["PartNO"].ToString().Trim() != "")
                        {
                            dr_tmp = db.DB_GetFirstDataByDataRow($"select PartNO,PartName,Specification FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}'");
                            if (dr_tmp != null) { material = $"{dr_tmp["PartNO"].ToString()}&nbsp;&nbsp;{dr_tmp["PartName"].ToString()}&nbsp;&nbsp;{dr_tmp["Specification"].ToString()}"; }
                        }
                        if (d["OP_NO"].ToString().Trim() != "")
                        {
                            string[] opNOs = d["OP_NO"].ToString().Trim().Split(';');
                            foreach (var s in opNOs)
                            {
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetMainDB.[dbo].[User] where UserNO='{s}'");
                                if (dr_tmp != null)
                                {
                                    if (opNO == "") { opNO = $"{s}{dr_tmp["Name"].ToString()}"; }
                                    else { opNO = $"<br />{opNO}{s}{dr_tmp["Name"].ToString()}"; }
                                }
                            }
                        }
                        if (int.Parse(d["CycleTime"].ToString()) > 0)
                        {
                            TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(d["CycleTime"].ToString()));
                            CT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                        }
                        if (int.Parse(d["ECT"].ToString()) > 0)
                        {
                            TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(d["ECT"].ToString()));
                            ECT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                        }
                        if (int.Parse(d["LowerCT"].ToString()) > 0)
                        {
                            TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(d["LowerCT"].ToString()));
                            LowerCT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                        }
                        if (int.Parse(d["UpperCT"].ToString()) > 0)
                        {
                            TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(d["UpperCT"].ToString()));
                            UpperCT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                        }
                        re = $"{re}<tr><td>{DateTime.Parse(d["LOGDateTimeID"].ToString()).ToString("yyyy/MM/dd HH:mm")}</td><td>{ReportTime}</td><td>{d["DOCNumberNO"].ToString()}</td><td>{material}</td><td>{opNO}</td><td>{d["EditFinishedQty"].ToString()}</td><td>{d["EditFailedQty"].ToString()}</td><td>{CT}</td><td>{ECT}</td><td>{LowerCT}</td><td>{UpperCT}</td></tr>";

                    }
                    re = $"{re}</table>";
                }
            }

            return re;

        }
        [HttpPost]
        public string ShowStation1WeekDetail_3(string key)//工站角度看效益(1週)
        {
            string re = "無資料.";


            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {

                DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{key}'");
                DateTime dtime = DateTime.Now;
                DateTime sTime = dtime.AddDays(-1);


                List<string> returnViewBag = new List<string>();
                DataTable tmp_dt2 = db.DB_GetData($"SELECT top 6 * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday<='{sTime.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_PP_Station["CalendarName"].ToString()}' order by CalendarName,Holiday desc");
                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                {
                    foreach (DataRow d in tmp_dt2.Rows)
                    {
                        returnViewBag.Add(Convert.ToDateTime(d["Holiday"]).ToString("yyyy/MM/dd"));
                    }
                }
                tmp_dt2 = db.DB_GetData($@"SELECT  a.*,(select DOCNumberNO from SoftNetSYSDB.[dbo].[APS_Simulation] as b where a.SimulationId=b.SimulationId) as DOCNumberNO FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] as a
                            where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{key}' and CONVERT(varchar(100), a.LOGDateTimeID, 111)>='{returnViewBag[5]}' and CONVERT(varchar(100), a.LOGDateTimeID, 111)<='{returnViewBag[0]}' order by a.LOGDateTime,a.SimulationId,a.LOGDateTimeID,a.OP_NO");
                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                {
                    re = $"<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                    re = $"{re}<tr><td>紀錄日期</td><td>報工時差</td><td>工單編號</td><td>料件資料</td><td>報工人員</td><td>OK數量</td><td>Fail數量</td><td>當下實際CT</td><td>當下有效CT</td><td>當下最佳CT</td><td>當下最差CT</td></tr>";
                    DataRow dr_tmp = null;
                    string ReportTime = "";
                    string opNO = "";
                    string CT = "";
                    string ECT = "";
                    string LowerCT = "";
                    string UpperCT = "";
                    string material = "";
                    foreach (DataRow d in tmp_dt2.Rows)
                    {
                        ReportTime = "0:0";
                        opNO = "";
                        CT = "";
                        ECT = "";
                        LowerCT = "";
                        UpperCT = "";
                        material = "";
                        if (int.Parse(d["ReportTime"].ToString()) > 0)
                        {
                            TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(d["ReportTime"].ToString()));
                            ReportTime = $"{(int)standardTime_DIS.TotalHours}時{standardTime_DIS.Minutes}分";
                        }
                        if (d["PartNO"].ToString().Trim() != "")
                        {
                            dr_tmp = db.DB_GetFirstDataByDataRow($"select PartNO,PartName,Specification FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}'");
                            if (dr_tmp != null) { material=$"{dr_tmp["PartNO"].ToString()}&nbsp;&nbsp;{dr_tmp["PartName"].ToString()}&nbsp;&nbsp;{dr_tmp["Specification"].ToString()}"; }
                        }
                        if (d["OP_NO"].ToString().Trim() != "")
                        {
                            string[] opNOs = d["OP_NO"].ToString().Trim().Split(';');
                            foreach (var s in opNOs)
                            {
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetMainDB.[dbo].[User] where UserNO='{s}'");
                                if (dr_tmp != null)
                                {
                                    if (opNO == "") { opNO = $"{s}{dr_tmp["Name"].ToString()}"; }
                                    else { opNO = $"<br />{opNO}{s}{dr_tmp["Name"].ToString()}"; }
                                }
                            }
                        }
                        if (int.Parse(d["CycleTime"].ToString()) > 0)
                        {
                            TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(d["CycleTime"].ToString()));
                            CT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                        }
                        if (int.Parse(d["ECT"].ToString()) > 0)
                        {
                            TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(d["ECT"].ToString()));
                            ECT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                        }
                        if (int.Parse(d["LowerCT"].ToString()) > 0)
                        {
                            TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(d["LowerCT"].ToString()));
                            LowerCT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                        }
                        if (int.Parse(d["UpperCT"].ToString()) > 0)
                        {
                            TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(d["UpperCT"].ToString()));
                            UpperCT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                        }
                        re = $"{re}<tr><td>{DateTime.Parse(d["LOGDateTimeID"].ToString()).ToString("yyyy/MM/dd HH:mm")}</td><td>{ReportTime}</td><td>{d["DOCNumberNO"].ToString()}</td><td>{material}</td><td>{opNO}</td><td>{d["EditFinishedQty"].ToString()}</td><td>{d["EditFailedQty"].ToString()}</td><td>{CT}</td><td>{ECT}</td><td>{LowerCT}</td><td>{UpperCT}</td></tr>";

                    }
                    re = $"{re}</table>";
                }
            }

            return re;

        }

        [HttpPost]
        public JObject ShowStation1WeekDetail(string key)
        {
            //string re = "<label>sdfsdfsdfsdf</label>";
            //return JsonToCnt(re);
            JArray x_labels = new JArray();
            JArray data_ST = new JArray();//行事曆工時
            JArray data_OT = new JArray();//實際工時
            JArray data_LT = new JArray();//負荷工時

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                
                DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{key}'");
                DateTime dtime = DateTime.Now;
                DateTime stime = DateTime.Now.AddDays(-31);

                #region 變數
                DateTime logTime = DateTime.Now;
                DateTime tmp_sDate = DateTime.Now;
                DateTime tmp_eDate = DateTime.Now;
                DataTable tmp_dt3 = null;
                float standardTime = 0;
                float totTTime = 0;
                float totTTime2 = 0;
                #endregion

                DataTable tmp_dt2 = db.DB_GetData($"SELECT  * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday>='{stime.ToString("yyyy/MM/dd")}' and Holiday<'{dtime.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr["CalendarName"].ToString()}' order by CalendarName,Holiday");
                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                {
                    foreach (DataRow dr_StandardTime in tmp_dt2.Rows)
                    {
                        logTime = Convert.ToDateTime(dr_StandardTime["Holiday"]);
                        x_labels.Add($"{logTime.Day.ToString()}日");
                        standardTime = 0;
                        totTTime = 0;
                        totTTime2 = 0;
                        #region 取得標準工作時間 data_ST
                        if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                        {
                            #region 取得工作時間
                            tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                            tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                            standardTime += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                            tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                            tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                            standardTime += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                            #endregion
                        }
                        if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                        {
                            #region 取得工作時間
                            tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                            tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                            standardTime += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                            tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                            tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                            standardTime += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                            #endregion
                        }
                        if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                        {
                            #region 取得工作時間
                            tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                            tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                            standardTime += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                            tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                            tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                            standardTime += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                            #endregion
                        }
                        if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) == true)
                        {
                            #region 取得工作時間
                            tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                            tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                            standardTime += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                            tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                            tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                            standardTime += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                            #endregion
                        }
                        data_ST.Add(standardTime);
                        #endregion

                        #region 取得 實際工作時間 data_OT
                        //###??? 有跨日按start的問題
                        tmp_dt3 = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTime, 111)='{logTime.ToString("yyyy/MM/dd")}' and (OperateType like '%開工%' or OperateType like '%停工%' or OperateType like '%關站%') order by LOGDateTime");
                        if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                        {
                            char run = '0';
                            foreach (DataRow dr3 in tmp_dt3.Rows)
                            {
                                if (dr3["OperateType"].ToString().IndexOf("開工") > 0)
                                {
                                    if (run == '0')
                                    {
                                        tmp_sDate = Convert.ToDateTime(dr3["LOGDateTime"]);
                                        run = '1';
                                    }
                                }
                                else
                                {
                                    if (run == '1')
                                    {
                                        if (dr3["OperateType"].ToString().IndexOf("停工") > 0 || dr3["OperateType"].ToString().IndexOf("關站") > 0)
                                        {
                                            tmp_eDate = Convert.ToDateTime(dr3["LOGDateTime"]);
                                            totTTime += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), tmp_sDate, tmp_eDate);
                                            run = '0';
                                        }
                                    }
                                }
                            }
                            if (run == '1' && dr_StandardTime != null)
                            {
                                tmp_eDate = tmp_sDate;
                                if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) == true)
                                {
                                    run = '2';
                                    string[] comp_Night = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                    string[] comp = dr_StandardTime["Shift_Graveyard"].ToString().Trim().Split(',');
                                    if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                                    { tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                    else { tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                }
                                else if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                                {
                                    run = '2';
                                    string[] comp = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                    if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                    { tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                    else
                                    { tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                }
                                else if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                                {
                                    run = '2';
                                    string[] comp = dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',');
                                    tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                }
                                else if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                                {
                                    run = '2';
                                    string[] comp = dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',');
                                    tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                }
                                if (run == '2' && tmp_sDate > tmp_eDate)
                                {
                                    totTTime += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), tmp_sDate, tmp_eDate);
                                }
                            }
                            
                        }
                        data_OT.Add(totTTime);
                        #endregion

                        #region 取得 最大負荷總工時 data_LT
                        tmp_dt3 = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), CalendarDate, 111)='{logTime.ToString("yyyy/MM/dd")}'");
                        if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                        {
                            char run = '0';
                            foreach (DataRow dr3 in tmp_dt3.Rows)
                            {
                                totTTime2 += (int.Parse(dr3["Time_TOT"].ToString()));
                            }
                        }
                        data_LT.Add(totTTime2);
                        #endregion

                    }
                }
            }

            //return '{"stationNO":"A01","data":["A","B"]}'
            return JObject.FromObject(new
            {
                stationNO = key,
                x_labels = x_labels,
                d_ST = data_ST,
                d_OT = data_OT,
                d_LT = data_LT
            });

            //JObject jo1 = new JObject();
            //JProperty stationNO = new JProperty("stationNO", key);
            //jo1.Add(stationNO);
            //JArray ja = new JArray();
            //ja.Add(new string[] { "A", "B" });
            //jo1.Add(ja);
            //return jo1;
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(Ctrl, dt));
        }

        private CIR01Service EditService()
        {
            return new CIR01Service(Ctrl);
        }

        //讀取要修改的資料(Get Updated Json)
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

        //新增(DB)
        public async Task<JsonResult> Create(string json)
        {
            return Json(await EditService().CreateAsync(_Str.ToJson(json)));
        }
        //修改(DB)
        public async Task<JsonResult> Update(string key, string json)
        {
            //return Json(await EditService().UpdateAsync(key, _Str.ToJson(json)));
            ResultDto row = await EditService().UpdateAsync(key, _Str.ToJson(json));
            if (row.ErrorMsg == "")
            {
                //JObject json2 = _Str.ToJson(json);
                //var rows = json2["_rows"] as JArray;
                string[] data = key.Split(';');//['LOGDateTime', 'Id', 'ServerId', 'StationNO', 'OP_NO'];
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataRow dr_log = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{data[2]}' and StationNO='{data[3]}' and Id='{data[1]}' and OP_NO='{data[4]}' and LOGDateTime='{data[0]}'");
                    DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{data[2]}' and StationNO='{data[3]}' and Id='{data[1]}' and OP_NO='{data[4]}'");
                    if (dr != null && dr_log != null)
                    {
                        int okqty = int.Parse(dr_log["OKQTY"].ToString()) - int.Parse(dr["ProductFinishedQty"].ToString());
                        int failqty = int.Parse(dr_log["FailQTY"].ToString()) - int.Parse(dr["ProductFailedQty"].ToString());
                        string tmp = "";
                        if (!dr.IsNull("RemarkTimeS") && !dr.IsNull("RemarkTimeE"))
                        {
                            int ctQTY = int.Parse(dr["ProductFinishedQty"].ToString()) + int.Parse(dr["ProductFailedQty"].ToString()) + okqty + failqty;
                            dr_log = db.DB_GetFirstDataByDataRow($"SELECT CalendarName FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{data[2]}' and StationNO='{data[3]}'");
                            decimal ct = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_log["CalendarName"].ToString(), Convert.ToDateTime(dr["RemarkTimeS"].ToString()), Convert.ToDateTime(dr["RemarkTimeE"].ToString())) / ctQTY;
                            decimal ct_log = ct < 1 ? 0 : ct;
                            if (int.Parse(dr["CycleTime"].ToString()) != 0) { ct = (ct + int.Parse(dr["CycleTime"].ToString())) > 0 ? Math.Round((ct + int.Parse(dr["CycleTime"].ToString())) / 2) : ct; }
                            if (ct < 1) { ct = 0; }
                            tmp = $",CycleTime={ct.ToString()}";
                        }
                        if (db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SFC_StationProjectDetail] SET ProductFinishedQty+={okqty.ToString()},ProductFailedQty+={failqty.ToString()}{tmp} where ServerId='{dr["ServerId"].ToString()}' and Id='{dr["Id"].ToString()}' and StationNO='{dr["StationNO"].ToString()}' and OP_NO='{dr["OP_NO"].ToString()}'"))
                        {
                        }

                    }
                }
            }

            return Json(row);
        }

        //刪除(DB)
        public async Task<JsonResult> Delete(string key)
        {
            //return Json(await EditService().DeleteAsync(key));
            ResultDto row = await EditService().DeleteAsync(key);
            if (row.ErrorMsg == "")
            {
                //JObject json2 = _Str.ToJson(json);
                //var rows = json2["_rows"] as JArray;
                string[] data = key.Split(';');//['LOGDateTime', 'Id', 'ServerId', 'StationNO', 'OP_NO'];
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    db.DB_SetData($"Delete FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{data[2]}' and StationNO='{data[3]}' and Id='{data[1]}' and OP_NO='{data[4]}'");
                }
            }

            return Json(row);

        }
    }

}
