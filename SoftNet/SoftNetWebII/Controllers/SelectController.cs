using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class SelectController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;

        public class Class1
        {
            public string[] labels { get; set; }
            public Class2[] datasets { get; set; }
        }
        public class Class2
        {
            public string type { get; set; }
            public string label { get; set; }
            public float[] data { get; set; }
            public bool fill { get; set; } = false;
            public string borderColor { get; set; }
            public string backgroundColor { get; set; }
        }
        public class Class3
        {
            public bool responsive { get; set; }
            public bool maintainAspectRatio { get; set; }
            public Object scales { get; set; }
        }

        public class ReturnClassJson
        {
            public ReturnClassJson(string type, Class1 data, Class3 options)
            {
                this.type = type;
                this.data = data;
                this.options = options;
            }
            public string type { get; set; }
            public Class1 data { get; set; }
            public Class3 options { get; set; }
        }

        public SelectController(SNWebSocketService websocket, SFC_Common sfc_Common)
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
        public ActionResult Read()
        {
            List<IdStrDto> docName = new List<IdStrDto>();
            List<IdStrDto> inName = new List<IdStrDto>();
            List<IdStrDto> outName = new List<IdStrDto>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOCRole] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        docName.Add(new IdStrDto(dr["DOCNO"].ToString(), dr["DOCName"].ToString()));
                    }
                }
                dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        inName.Add(new IdStrDto(dr["StoreNO"].ToString(), dr["StoreName"].ToString()));
                        outName.Add(new IdStrDto(dr["StoreNO"].ToString(), dr["StoreName"].ToString()));
                    }
                }
            }
            ViewBag.OutName = outName;
            ViewBag.InName = inName;
            ViewBag.DocName = docName;
            return View();
        }

        public ActionResult Efficient()
        {
            return View();
        }
        public JsonResult Efficient_Request(string data1, string data2)//統計範圍,分析類型
        {
            float[] MarriageRate = null;
            float[] LoadRate = null;
            float[] TOTGrowing = null;
            float[] aaa = null;
            float[] bbb = null;
            int[] station_Count = null;
            List<string> station_list = new List<string>();
            string[] labels = null;

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                List<string> date_list = new List<string>();
                string sql = $"SELECT top 30 * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday<'{DateTime.Now.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{_Fun.Config.DefaultCalendarName}' order by Holiday desc";
                DataTable dtdata = db.DB_GetData(sql);
                if (dtdata != null && dtdata.Rows.Count > 0)
                {

                    foreach (DataRow d in dtdata.Rows)
                    {
                        date_list.Add(Convert.ToDateTime(d["Holiday"]).ToString("yyyy/MM/dd"));
                        station_list.Add("");
                    }
                }
                MarriageRate = new float[(date_list.Count - 1)];
                LoadRate = new float[(date_list.Count - 1)];
                TOTGrowing = new float[(date_list.Count - 1)];
                aaa = new float[(date_list.Count - 1)];
                bbb = new float[(date_list.Count - 1)];
                station_Count = new int[(date_list.Count - 1)];
                labels = new string[(date_list.Count - 1)];



                DataTable tmp_dt2 = null;
                DataRow rmp_dr2 = null;

                dtdata = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' order by FactoryName,LineName,StationNO");
                if (dtdata != null && dtdata.Rows.Count > 0)
                {
                    #region 變數
                    DateTime logTime = DateTime.Now;
                    DateTime tmp_sDate = DateTime.Now;
                    DateTime tmp_eDate = DateTime.Now;
                    DateTime stime = DateTime.Now.AddDays(-30);
                    DateTime dtime = DateTime.Now;


                    DataTable tmp_dt3 = null;


                    int[] standardTime = new int[(date_list.Count - 1)];
                    float[] totTTime = new float[(date_list.Count - 1)];
                    int[] totTTime2 = new int[(date_list.Count - 1)];

                    int[] aaa_log = new int[(date_list.Count - 1)];

                    float standardTime_log = 0;
                    float t1Time_log = 0.0f;

                    #endregion

                    int iCount = 0;
                    bool isDisplay = false;
                    float tmp = 0;
                    string beforStation = "";
                    foreach (DataRow dr in dtdata.Rows)
                    {
                        for (int i = 0; i < (date_list.Count - 1); ++i)
                        {
                            labels[i] = date_list[i].Split('/')[2];
                            DataRow dr_lastTime = null;
                            #region 取得標準工作時間 standardTime
                            DataRow dr_StandardTime = db.DB_GetFirstDataByDataRow($"SELECT  * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday='{date_list[i]}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr["CalendarName"].ToString()}' order by Holiday");
                            if (dr_StandardTime != null)
                            {
                                logTime = Convert.ToDateTime(dr_StandardTime["Holiday"]);

                                if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                                {
                                    #region 取得工作時間
                                    tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                    tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                    standardTime[i] += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                    tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                    tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                    standardTime[i] += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                    #endregion
                                }
                                if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                                {
                                    #region 取得工作時間
                                    tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                    tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                    standardTime[i] += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                    tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                    tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                    standardTime[i] += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                    #endregion
                                }
                                if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                                {
                                    #region 取得工作時間
                                    tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                    tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                    standardTime[i] += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                    tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                    tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                    standardTime[i] += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                    #endregion
                                }
                                if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) == true)
                                {
                                    #region 取得工作時間
                                    tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                    tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                    standardTime[i] += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                    tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                    tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                    standardTime[i] += _SFC_Common.TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                    #endregion
                                }
                            }
                            #endregion

                            #region 取得實際工作時間 totTTime
                            //###??? 有跨日按start的問題
                            tmp_dt3 = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTime, 111)='{date_list[i]}' and (OperateType like '%開工%' or OperateType like '%停工%' or OperateType like '%關站%') order by LOGDateTime");
                            if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                            {
                                char run = '0';
                                foreach (DataRow dr3 in tmp_dt3.Rows)
                                {
                                    if (dr3["OperateType"].ToString().IndexOf("開工") > 0)
                                    {
                                        if (beforStation != dr["StationNO"].ToString()) 
                                        {
                                            beforStation = dr["StationNO"].ToString();
                                            station_Count[i] += 1;
                                        }
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
                                                totTTime[i] += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), tmp_sDate, tmp_eDate);
                                                run = '0';
                                            }
                                        }
                                    }
                                }
                                if (run == '1' && dr_lastTime != null)
                                {
                                    tmp_eDate = tmp_sDate;
                                    if (bool.Parse(dr_lastTime["Flag_Graveyard"].ToString()) == true)
                                    {
                                        run = '2';
                                        string[] comp_Night = dr_lastTime["Shift_Night"].ToString().Trim().Split(',');
                                        string[] comp = dr_lastTime["Shift_Graveyard"].ToString().Trim().Split(',');
                                        if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                                        { tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                        else { tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                    }
                                    else if (bool.Parse(dr_lastTime["Flag_Night"].ToString()) == true)
                                    {
                                        run = '2';
                                        string[] comp = dr_lastTime["Shift_Night"].ToString().Trim().Split(',');
                                        if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                        { tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                        else
                                        { tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                    }
                                    else if (bool.Parse(dr_lastTime["Flag_Afternoon"].ToString()) == true)
                                    {
                                        run = '2';
                                        string[] comp = dr_lastTime["Shift_Afternoon"].ToString().Trim().Split(',');
                                        tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                    }
                                    else if (bool.Parse(dr_lastTime["Flag_Morning"].ToString()) == true)
                                    {
                                        run = '2';
                                        string[] comp = dr_lastTime["Shift_Morning"].ToString().Trim().Split(',');
                                        tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                    }
                                    if (run == '2' && tmp_sDate > tmp_eDate)
                                    {
                                        totTTime[i] += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), tmp_sDate, tmp_eDate);
                                    }
                                }
                            }

                            #endregion

                            #region 取得 最大負荷總工時 totTTime2
                            tmp_dt3 = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), CalendarDate, 111)='{date_list[i]}'");
                            if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                            {
                                char run = '0';
                                foreach (DataRow dr3 in tmp_dt3.Rows)
                                {
                                    totTTime2[i] += int.Parse(dr3["Time_TOT"].ToString());
                                }
                            }
                            #endregion

                            #region 取得週 aaa生產能量=有效CT/實際CT bbb可提升率=最佳CT/實際CT
                            rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select (sum((EditFinishedQty+EditFailedQty)*CycleTime)/sum((EditFinishedQty+EditFailedQty))) as ACT,(sum((EditFinishedQty+EditFailedQty)*ECT)/sum((EditFinishedQty+EditFailedQty))) as BCT,(sum((EditFinishedQty+EditFailedQty)*LowerCT)/sum((EditFinishedQty+EditFailedQty))) as CCT from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] 
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTimeID, 111)='{date_list[i]}' and CycleTime!=0 and ECT!=0 and LowerCT!=0 and (LowerCT*0.5)< CycleTime");
                            if (rmp_dr2 != null)
                            {
                                if (station_list[i].IndexOf($"{dr["StationNO"].ToString()},") < 0) { station_list[i] = $"{station_list[i]}{dr["StationNO"].ToString()},"; }
                                if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("BCT")) { aaa[i] += (float.Parse(rmp_dr2["BCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100; }
                                if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("CCT")) { bbb[i] += (float.Parse(rmp_dr2["CCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100; }
                            }
                            #endregion
                        }

                    }

                    for (int i = 0; i < (date_list.Count - 1); ++i)
                    {
                        if (totTTime[i] > 0 && standardTime[i] > 0)
                        {
                            MarriageRate[i] = ((totTTime[i] / standardTime[i]) * 100);
                            if (MarriageRate[i] > 100) { MarriageRate[i] = 100; }
                            else if (MarriageRate[i] < 0) { MarriageRate[i] = 0; }
                        }

                        if (totTTime[i] > 0 && totTTime2[i] > 0)
                        {
                            LoadRate[i] = (totTTime[i] / totTTime2[i]) * 100;
                            if (LoadRate[i] > 100) { LoadRate[i] = 100; }
                            else if (LoadRate[i] < 0) { LoadRate[i] = 0; }
                        }

                        if (aaa[i] != 0)
                        {
                            if (station_list[i] != "") { aaa[i] /= (station_list[i].Split(',').Length-1); }
                        }
                        if (bbb[i] != 0) 
                        {
                            if (station_list[i] != "") { bbb[i] /= (station_list[i].Split(',').Length - 1); }
                        }
                        if (aaa[i] > 100) { aaa[i] = 100; }
                        else if (aaa[i] < 0) { aaa[i] = 0; }
                        if (bbb[i] > 100) { bbb[i] = 100; }
                        else if (bbb[i] < 0) { bbb[i] = 0; }
                    }
                    for (int i = 0; i < (date_list.Count - 1); ++i)
                    {
                        if (i < (date_list.Count - 2))
                        {
                            TOTGrowing[i] = aaa[(i + 1)] - aaa[i];
                            if (TOTGrowing[i] > 100) { TOTGrowing[i] = 100; }
                            else if (TOTGrowing[i] < 0) { TOTGrowing[i] = 0; }
                        }
                    }
                }
            }

            Class3 re_options = new Class3();
            re_options.responsive = true;
            re_options.maintainAspectRatio = false;
            re_options.scales = new { y = new { beginAtZero = true } };

            Class2 re_datasets1 = new Class2();
            re_datasets1.type = "bar";
            re_datasets1.label = "嫁動率";
            re_datasets1.data = MarriageRate;
            re_datasets1.fill = true;
            re_datasets1.borderColor = "rgb(15, 45, 13)";
            re_datasets1.backgroundColor = "rgba(15, 45, 13, 0.2)";

            Class2 re_datasets1_1 = new Class2();
            re_datasets1_1.type = "bar";
            re_datasets1_1.label = "負荷率";
            re_datasets1_1.data = LoadRate;
            re_datasets1_1.fill = true;
            re_datasets1_1.borderColor = "rgb(215, 99, 132)";
            re_datasets1_1.backgroundColor = "rgba(215, 99, 132, 0.2)";

            Class2 re_datasets2 = new Class2();
            re_datasets2.type = "line";
            re_datasets2.label = "成長率";
            re_datasets2.data = TOTGrowing;
            re_datasets2.fill = false;
            re_datasets2.borderColor = "rgb(54, 162, 235)";
            re_datasets2.backgroundColor = "rgb(54, 162, 235)";

            Class2 re_datasets2_2 = new Class2();
            re_datasets2_2.type = "line";
            re_datasets2_2.label = "可提升率";
            re_datasets2_2.data = bbb;
            re_datasets2_2.fill = false;
            re_datasets2_2.borderColor = "rgb(34, 62, 235)";
            re_datasets2_2.backgroundColor = "rgb(34, 62, 235)";

            Class1 re_data = new Class1();
            re_data.labels = labels;
            re_data.datasets=new Class2[] { re_datasets1, re_datasets1_1, re_datasets2_2, re_datasets2 };

            ReturnClassJson config = new ReturnClassJson("scatter", re_data, re_options);



            return Json(config);
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private SelectService EditService()
        {
            return new SelectService(Ctrl);
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
    }
}
