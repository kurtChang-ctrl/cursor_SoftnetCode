using Base.Services;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Models;
using System.Data;
using System;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System.Collections.Generic;
using SoftNetWebII.Services;
using BaseApi.Controllers;
using Base.Models;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Linq;
using Newtonsoft.Json.Linq;
using Base;

namespace SoftNetWebII.Controllers
{
    public class DisplayController : Controller
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public DisplayController(SNWebSocketService websocket, SFC_Common sfc_Common)
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
        //
        public ActionResult DefineStandardsCT()//建立人為標準CT
        {
            List<string[]> HasStationNO_List = new List<string[]>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {

                DataTable dt_tmp = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}'");
                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt_tmp.Rows)
                    {
                        HasStationNO_List.Add(new string[] { dr["StationNO"].ToString(), dr["StationName"].ToString() });
                    }
                }
            }
            ViewBag.HasStationNO_List = HasStationNO_List;
            return View();
        }
        public ActionResult OddCT()//疑似乖離CT異動
        {
            List<string[]> HasStationNO_List = new List<string[]>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {

                DataTable dt_tmp = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}'");
                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt_tmp.Rows)
                    {
                        HasStationNO_List.Add(new string[] { dr["StationNO"].ToString(), dr["StationName"].ToString() });
                    }
                }
            }
            ViewBag.HasStationNO_List = HasStationNO_List;
            return View();
        }
        public string OddCT_GetData(OddCT_GetData data)//疑似乖離CT異動
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string inStation = "";
                if (data.Value != "ALL_Station")
                {
                    foreach(string s in data.Value.Split(';'))
                    {
                        if (inStation == "") { inStation = $"'{s}'"; }
                        else { inStation = $"{inStation},'{s}'"; }
                    }
                    if (inStation != "") { inStation = $" and StationNO in ({inStation})"; }
                }
                DataTable dt_tmp2 = null;
                DataTable dt_tmp = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' {inStation} order by StationNO");
                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                {
                    DataRow dr_tmp = null;
                    re = "<div><button id='button_setSTII' type='button' class='btn xg-btn' onclick='_me.onSetII'>寫入修正值</button></div>";
                    re = $"{re}<div><table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'><thead><tr>";
                    re = $"{re}<th>工站編號</th><th>生產製程名稱</th><th>生產工序</th><th>料號</th><th>品名/規格</th><th>統計筆數</th><th>平均CT</th><th>最佳CT</th><th>最差CT</th><th>有效CT</th><th>修正合理CT</th></tr></thead><tbody>";
                    foreach (DataRow dr in dt_tmp.Rows)
                    {
                        dt_tmp2 = db.DB_GetData($@"select a.*,b.PartName,b.Specification from SoftNetSYSDB.[dbo].[PP_EfficientDetail] as a
                                  join SoftNetMainDB.[dbo].[Material] as b on a.Sub_PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                                  where a.ServerId='{_Fun.Config.ServerId}' and a.DOCNO='' and a.StationNO='{dr["StationNO"].ToString()}' and (a.AverageCycleTime+a.EfficientCycleTime)/2>a.EfficientCycleTime order by a.Apply_PP_Name,a.PartNO");
                        if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                        {
                            string indexDisplay = "";
                            string avg = "";
                            string lct = "";
                            string uct = "";
                            string ect = "";
                            foreach (DataRow dr2 in dt_tmp2.Rows)
                            {
                                indexDisplay = dr2["IndexSN"].ToString();
                                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where PP_Name='{dr2["PP_Name"].ToString()}' and PartNO='{dr2["PartNO"].ToString()}' and StationNO='{dr["StationNO"].ToString()}' and IndexSN={indexDisplay}");
                                if (dr_tmp != null && (dr_tmp["Station_Custom_IndexSN"].ToString().Trim()!="" || dr_tmp["DisplayName"].ToString().Trim() != "")) 
                                { indexDisplay = $"{dr_tmp["Station_Custom_IndexSN"].ToString()}{dr_tmp["DisplayName"].ToString()}"; }
                                TimeSpan standardTime_DIS = new TimeSpan(0, 0, Convert.ToInt32(dr2["AverageCycleTime"].ToString()));
                                avg = $"{(int)standardTime_DIS.TotalMinutes}:{standardTime_DIS.Seconds}";
                                standardTime_DIS = new TimeSpan(0, 0, Convert.ToInt32(dr2["SD_LowerLimit"].ToString()));
                                lct = $"{(int)standardTime_DIS.TotalMinutes}:{standardTime_DIS.Seconds}";
                                standardTime_DIS = new TimeSpan(0, 0, Convert.ToInt32(dr2["SD_UpperLimit"].ToString()));
                                uct = $"{(int)standardTime_DIS.TotalMinutes}:{standardTime_DIS.Seconds}";
                                standardTime_DIS = new TimeSpan(0, 0, Convert.ToInt32(dr2["EfficientCycleTime"].ToString()));
                                ect = $"{(int)standardTime_DIS.TotalMinutes}:{standardTime_DIS.Seconds}";
                                re = $@"{re}<tr><td>{dr["StationNO"].ToString()}</td><td>{dr2["PP_Name"].ToString()}</td><td>{indexDisplay}</td><td>{dr2["PartNO"].ToString()}</td><td>{dr2["PartName"].ToString()}&nbsp;{dr2["Specification"].ToString()}</td><td>{dr2["CountQTY"].ToString()}</td><td>{avg}</td><td>{lct}</td><td>{uct}</td><td>{ect}</td>
                                        <td><input type='number' id='M_{dr2["Id"].ToString()}' name='SetCT_M' value='0' />分<input type='number' id='S_{dr2["Id"].ToString()}' name='SetCT_S' value='0' />秒</td></tr>";
                            }
                        }
                    }
                    re = $"{re}</tbody></table></div>";
                }
            }
            return re;
        }
        public IActionResult SelectOLDSID(string id)//查詢計畫已轉生產的歷史紀錄
        {
            string re = "<p>查詢近一年內已轉生產的歷史紀錄</p><hr>";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt_APS_NeedData = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and State !='0' and State !='1' and State !='2' and State !='3' order by UpdateTime desc");
                if (dt_APS_NeedData != null && dt_APS_NeedData.Rows.Count > 0)
                {
                    DataRow dr_tmp = null;
                    string tmp = "";
                    string state = "";
                    string tmp_date = "";
                    re = $"{re}<div class='xp-prog'>";
                    re = $@"{re}<table id='tableRead' class='table table-bordered xg-table' cellspacing='0'> <thead><tr>
                        <th>排程編號</th><th>更新時間</th><th>類型</th><th>類型名稱</th><th>計畫狀態</th><th>計畫入庫日</th><th>料號/品名/規格</th><th>計畫量</th><th>投產量</th><th>生產製程</th><th>BOM編號</th></tr></thead><tbody>";
                    foreach (DataRow dr in dt_APS_NeedData.Rows)
                    {
                        switch (dr["State"].ToString())
                        {
                            case "0": state = "初建"; break;
                            case "1": state = "模擬中"; break;
                            case "2": state = "完成模擬"; break;
                            case "3": state = "模擬逾時取消"; break;
                            case "4": state = "已轉倉儲計畫"; break;
                            case "5": state = "失敗"; break;
                            case "6": state = "已轉生產"; break;
                            case "9": state = "已入庫"; break;
                            default: state = "未定義"; break;
                        }
                        if (dr["NeedType"].ToString() == "1") { tmp = "訂單"; }
                        else if (dr["NeedType"].ToString() == "2") { tmp = "客戶"; }
                        else if (dr["NeedType"].ToString() == "5") { tmp = "底稿"; }
                        else { tmp = "廠內"; }
                        tmp_date = "";
                        if (!dr.IsNull("UpdateTime")) { tmp_date = Convert.ToDateTime(dr["UpdateTime"]).ToString("yy/MM/dd HH:mm"); }
                        re = $"{re}<tr><td>{dr["Id"].ToString()}</td><td>{tmp_date}</td><td>{tmp}</td>";
                        tmp = "";
                        if (!dr.IsNull("NeedSource") && dr["NeedSource"].ToString() != "") { tmp = dr["NeedSource"].ToString(); }
                        else if (!dr.IsNull("CTName") && dr["CTName"].ToString() != "") { tmp = dr["CTName"].ToString(); } else if (!dr.IsNull("CTNO") && dr["CTNO"].ToString() != "") { tmp = dr["CTNO"].ToString(); }
                        re = $"{re}<td>{tmp}</td><td>{state}</td>";
                        if (dr.IsNull("NeedSimulationDate")) { re = $"{re}<td></td>"; }
                        else { re = $"{re}<td>{Convert.ToDateTime(dr["NeedSimulationDate"]).ToString("yy/MM/dd HH:mm")}</td>"; }
                        dr_tmp = db.DB_GetFirstDataByDataRow($"select PartName,Specification from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                        if (dr_tmp != null)
                        { re = $"{re}<td>{dr["PartNO"].ToString()}&nbsp;{dr_tmp["PartName"].ToString()}&nbsp;{dr_tmp["Specification"].ToString()}</td><td>{dr["NeedQTY"].ToString()}</td>"; }
                        else
                        { re = $"{re}<td></td><td>{dr["NeedQTY"].ToString()}</td>"; }
                        dr_tmp = db.DB_GetFirstDataByDataRow($"select *,(NeedQTY+SafeQTY) as SQTY from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr["Id"].ToString()}' and PartSN=0");
                        if (dr_tmp != null)
                        { re = $"{re}<td>{dr_tmp["SQTY"].ToString()}</td>"; }
                        else
                        { re = $"{re}<td> </td>"; }
                        re = $"{re}<td>{dr["Apply_PP_Name"].ToString()}</td><td>{dr["BOMId"].ToString()}</td></tr>";
                    }
                    re = $"{re}</tbody></table></div>";
                }
                else
                {
                    re = $"{re}<p> 目前無轉生產的紀錄</p>";
                }

            }
            ViewBag.HtmlOutput = re;
            return View();
        }

        public IActionResult SelectStationState(string id)//查詢工站效益總表
        {
            ViewBag.HtmlOutput = SelectStationStateII();
            return View();
        }
        [HttpPost]
        private string SelectStationStateII()
        {
            string re = "";
            Dictionary<string, string[]> stationNOList = new Dictionary<string, string[]>();
            DataTable dt = null;
            DataRow tmp = null;
            DataRow tmp2 = null;
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                #region 生產量無法達標
                #endregion

                #region Simulation_ErrorData
                dt = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and StationNO!='' and StationNO!='{_Fun.Config.OutPackStationName}' and ActionType='' and (ErrorType='05' or ErrorType='01' or ErrorType='02' or ErrorType='04' or ErrorType='03' or ErrorType='15') order by LogDate desc");
                if (dt != null && dt.Rows.Count > 0)
                {
                    string log = "";
                    foreach (DataRow dr in dt.Rows)
                    {
                        log = "";
                        switch (dr["ErrorType"].ToString())
                        {
                            case "05":
                                log = $"生產備料無作為且無領料單據發起";
                                break;
                            case "01":
                            case "02":
                                string tmp_S = "";
                                tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                                if (tmp != null)
                                {
                                    tmp_S = tmp["PartNO"].ToString();
                                }
                                log = $"料號:{tmp_S},庫存量不足原本預定Keep量";
                                break;
                            case "04":
                                log = $"生產用原物料應領用而未領用";
                                break;
                            case "03":
                                log = $"單據:{dr["DOCNumberNO"].ToString()}&nbsp;未如預期開工";
                                break;
                            case "12":
                                log = $"工單:{dr["DOCNumberNO"].ToString()}&nbsp;應執行關閉結束未關";
                                break;
                            case "15":
                                log = $"工單:{dr["DOCNumberNO"].ToString()}&nbsp;報工作業可能有延遲";
                                break;
                            default:
                                log = "未知的錯誤";
                                break;
                        }
                        if (!stationNOList.ContainsKey(dr["StationNO"].ToString())) { stationNOList.Add(dr["StationNO"].ToString(), new string[] { "E", log, dr["LogDate"].ToString() }); }
                    }
                }
                #endregion

                #region 開工的工站
                dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' and State='1'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (!stationNOList.ContainsKey(dr["StationNO"].ToString())) { stationNOList.Add(dr["StationNO"].ToString(), new string[] { "S", "", "" }); }
                    }
                }
                dt = db.DB_GetData($"SELECT StationNO FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' and RemarkTimeS is not NULL and RemarkTimeE is NULL group by StationNO");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (!stationNOList.ContainsKey(dr["StationNO"].ToString())) { stationNOList.Add(dr["StationNO"].ToString(), new string[] { "S", "", "" }); }
                    }
                }
                #endregion

                #region 剩下的工站
                dt = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (!stationNOList.ContainsKey(dr["StationNO"].ToString())) { stationNOList.Add(dr["StationNO"].ToString(), new string[] { "", "", "" }); }
                    }
                }
                #endregion

                DataRow dr_PP_Station = null;
                DataRow APS_PartNOTimeNote = null;
                DataRow APS_WorkTimeNote = null;
                DataRow dr_tmp = null;
                DataTable tmp_dt = null;
                bool is_TOT = false;
                Dictionary<string, string[]> sidList = new Dictionary<string, string[]>();
                string errLOG = "";
                string errTime = "";
                string beforKey = "";
                foreach (KeyValuePair<string, string[]> obj in stationNOList)
                {
                    is_TOT = false;
                    sidList.Clear();
                    if (obj.Key!= beforKey)
                    {
                        errLOG = "";
                        errTime = "";
                        beforKey = obj.Key;
                    }

                    dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{obj.Key}'");
                    if (obj.Value[0] == "E") 
                    { 
                        re = $"{re}<table style='background-color:tomato;border:1px solid black;width:100%' rules='all'>";
                        if (errLOG != "") { errLOG =$"<br />{obj.Value[1]}"; }
                        else { errLOG = obj.Value[1]; }
                        if (errTime != "") { errTime = $"<br />{obj.Value[2]}"; }
                        else { errTime = obj.Value[2]; }
                    }
                    else if (obj.Value[0] == "W") { re = $"{re}<table style='background-color:gold;border:1px solid black;width:100%' rules='all'>"; }
                    else if (obj.Value[0] == "S") { re = $"{re}<table style='background-color:springgreen;border:1px solid black;width:100%' rules='all'>"; }
                    else
                    {
                        re = $"{re}<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                    }
                    string opNO = "";
                    string partNO = "";
                    string partName = "";
                    string muti = "";
                    if (dr_PP_Station["Station_Type"].ToString() == "1")
                    {
                        dr_tmp = db.DB_GetFirstDataByDataRow($"select State from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{obj.Key}'");
                        if (dr_tmp != null && dr_tmp["State"].ToString() == "1")
                        { re = $"{re}<tr><td>單工工站:{obj.Key}&nbsp;&nbsp;{dr_PP_Station["StationName"].ToString()}&nbsp;&nbsp;狀態:開工中</td><td>生產線:{dr_PP_Station["LineName"].ToString()}</td>"; }
                        else
                        { re = $"{re}<tr><td>單工工站:{obj.Key}&nbsp;&nbsp;{dr_PP_Station["StationName"].ToString()}</td><td>生產線:{dr_PP_Station["LineName"].ToString()}</td>"; }
                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{obj.Key}'");
                        if (tmp != null && tmp["OrderNO"].ToString() != "")
                        {
                            is_TOT = true;
                            sidList.Add(tmp["SimulationId"].ToString(), new string[] { tmp["PartNO"].ToString(), tmp["PP_Name"].ToString() });
                            opNO = tmp["OP_NO"].ToString();
                            partNO = tmp["PartNO"].ToString();
                            tmp2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{tmp["PartNO"].ToString()}'");
                            partName = $"{tmp2["PartName"].ToString()}&nbsp;&nbsp;{tmp2["Specification"].ToString()}&nbsp;&nbsp;{tmp2["Model"].ToString()}";
                            APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{tmp["SimulationId"].ToString()}'");
                            if (APS_PartNOTimeNote != null)
                            { re = $"{re}<td>工單:{tmp["OrderNO"].ToString()}</td><td>工單量:{APS_PartNOTimeNote["NeedQTY"].ToString()}</td><td>已報工量:{APS_PartNOTimeNote["Detail_QTY"].ToString()}</td></tr>"; }
                            else { re = $"{re}<td>工單:{tmp["OrderNO"].ToString()}</td><td>工單量:0</td><td>已報工量:0</td></tr>"; }
                        }
                        else { re = $"{re}<td>工單:</td><td>工單量:0</td><td>已報工量:0</td></tr>"; }
                    }
                    else
                    {
                        re = $"{re}<tr><td>多工工站:{obj.Key}&nbsp;&nbsp;{dr_PP_Station["StationName"].ToString()}</td><td>生產線:{dr_PP_Station["LineName"].ToString()}</td>"; 
                        tmp_dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{obj.Key}' and RemarkTimeS is not NULL and RemarkTimeE is NULL");
                        if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                        {
                            is_TOT = true;
                            string wo = "";
                            string qty = "";
                            string okqty = "";
                            if (tmp_dt.Rows.Count > 1) { muti = "<br />"; }
                            int i = 0;
                            foreach (DataRow dr in tmp_dt.Rows)
                            {
                                if (!sidList.ContainsKey(dr["SimulationId"].ToString()))
                                { sidList.Add(dr["SimulationId"].ToString(), new string[] { dr["PartNO"].ToString(), dr["PP_Name"].ToString() }); }
                                if (partNO == "") { partNO = $"{(++i).ToString()}.{dr["PartNO"].ToString()}"; } else { partNO = $"{partNO}<br />{(++i).ToString()}.{dr["PartNO"].ToString()}"; }
                                tmp2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                                if (partName == "") { partName = $"{tmp2["PartName"].ToString()}&nbsp;&nbsp;{tmp2["Specification"].ToString()}&nbsp;&nbsp;{tmp2["Model"].ToString()}"; } else { partName = $"{partName}<br />{tmp2["PartName"].ToString()}&nbsp;&nbsp;{tmp2["Specification"].ToString()}&nbsp;&nbsp;{tmp2["Model"].ToString()}"; }

                                if (wo == "") { wo = dr["OrderNO"].ToString(); } else { wo = $"{wo}<br />{dr["OrderNO"].ToString()}"; }
                                APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr["SimulationId"].ToString()}'");
                                if (APS_PartNOTimeNote != null)
                                {
                                    if (qty == "") { qty = APS_PartNOTimeNote["NeedQTY"].ToString(); } else { qty = $"{qty}<br />{APS_PartNOTimeNote["NeedQTY"].ToString()}"; }
                                    if (okqty == "") { okqty = APS_PartNOTimeNote["Detail_QTY"].ToString(); } else { okqty = $"{okqty}<br />{APS_PartNOTimeNote["Detail_QTY"].ToString()}"; }
                                }
                            }
                            re = $"{re}<td>工單:{muti}{wo}&nbsp;&nbsp;狀態:開工中</td><td>工單量:{muti}{qty}</td><td>已報工量:{muti}{okqty}</td></tr>";
                        }
                        else { re = $"{re}<td>工單:</td><td>工單量:0</td><td>已報工量:0</td></tr>"; }
                    }
                    re = $"{re}<tr><td>生產料號:{muti}{partNO}</td><td colspan='4'>{muti}{partName}</td></tr>";

                    string b01 = "";
                    string b02 = "";
                    string c01 = "0:0";
                    string c02 = "0:0";
                    if (is_TOT)
                    {
                        #region 效能,工時 計算 c01
                        APS_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time1_C+Time2_C+Time3_C+Time4_C) as C01 from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{obj.Key}' and CONVERT(varchar(100), CalendarDate, 111)='{DateTime.Now.ToString("yyyy/MM/dd")}'");
                        if (APS_WorkTimeNote != null && !APS_WorkTimeNote.IsNull("C01") && APS_WorkTimeNote["C01"].ToString() != "")
                        {
                            TimeSpan c01_DIS = new TimeSpan(0, 0, int.Parse(APS_WorkTimeNote["C01"].ToString()));
                            c01 = $"{(int)c01_DIS.TotalHours}:{c01_DIS.Minutes}";
                        }
                        #region 取得每日 實際工作時間 c02
                        DateTime stime = DateTime.Now;
                        DateTime etime = DateTime.Now;
                        DateTime logTime = stime;
                        int t1Time = 0;
                        //###??? 有跨日按start的問題, 智慧模式少IndexSN條件
                        DataRow dr_StandardTime = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday='{DateTime.Now.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_PP_Station["CalendarName"].ToString()}' order by CalendarName,Holiday desc");
                        if (dr_StandardTime != null)
                        {
                            tmp_dt = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{obj.Key}' and CONVERT(varchar(100), LOGDateTime, 111)='{DateTime.Now.ToString("yyyy/MM/dd")}' and (OperateType like '%開工%' or OperateType like '%停工%' or OperateType like '%關站%') order by LOGDateTime");
                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                            {
                                char run = '0';
                                foreach (DataRow dr3 in tmp_dt.Rows)
                                {
                                    if (dr3["OperateType"].ToString().IndexOf("開工") > 0)
                                    {
                                        if (run == '0')
                                        {
                                            stime = Convert.ToDateTime(dr3["LOGDateTime"]);
                                            run = '1';
                                        }
                                    }
                                    else
                                    {
                                        if (run == '1')
                                        {
                                            if (dr3["OperateType"].ToString().IndexOf("停工") > 0 || dr3["OperateType"].ToString().IndexOf("關站") > 0)
                                            {
                                                etime = Convert.ToDateTime(dr3["LOGDateTime"]);
                                                t1Time += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), stime, etime);
                                                run = '0';
                                            }
                                        }
                                    }
                                }
                                if (run == '1')
                                {
                                    etime = stime;
                                    if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) == true)
                                    {
                                        run = '2';
                                        string[] comp_Night = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                        string[] comp = dr_StandardTime["Shift_Graveyard"].ToString().Trim().Split(',');
                                        if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                        else { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                    }
                                    else if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                                    {
                                        run = '2';
                                        string[] comp = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                        if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                        else
                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                    }
                                    else if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                                    {
                                        run = '2';
                                        string[] comp = dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',');
                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                    }
                                    else if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                                    {
                                        run = '2';
                                        string[] comp = dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',');
                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                    }
                                    if (run == '2' && stime > etime)
                                    {
                                        t1Time += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), stime, etime);
                                    }
                                }
                            }
                            if (t1Time != 0)
                            {
                                TimeSpan c01_DIS = new TimeSpan(0, 0, t1Time);
                                c02 = $"{(int)c01_DIS.TotalHours}:{c01_DIS.Minutes}";
                            }
                        }
                        #endregion

                        #endregion
                    }
                    re = $"{re}<tr><td>歷史效益差:{b01}</td><td>在線效益差:{b02}</td><td>本日最大負荷:{c01}分</td><td>本日累計負荷:{c02}分</td><td>作業人員:{opNO}</td></tr>";

                    if (sidList.Count > 0)
                    {
                        string t01 = "";
                        string t02 = "";
                        string t03 = "";
                        string t04 = "";
                        string t05 = "";
                        int i = 0;
                        foreach (KeyValuePair<string, string[]> sid in sidList)
                        {
                            t01 = "";
                            t02 = "";
                            t03 = "";
                            t04 = "";
                            t05 = "";
                            if (dr_PP_Station["Station_Type"].ToString() == "1")
                            {
                                tmp = db.DB_GetFirstDataByDataRow($"SELECT (sum(CycleTime)/count(*)) as T01 FROM SoftNetLogDB.[dbo].[SFC_StationDetail] where SimulationId='{sid.Key}'");
                                if (tmp != null && !tmp.IsNull("T01") && tmp["T01"].ToString() != "") { t01 = tmp["T01"].ToString(); }
                                tmp = db.DB_GetFirstDataByDataRow($"select AVG(EfficientCycleTime) as ECT,AVG(SD_UpperLimit) as UpperCT,AVG(Custom_SD_LowerLimit) as Custom_LowerCT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and PartNO='{sid.Value[0]}' and StationNO='{obj.Key}' and PP_Name='{sid.Value[1]}' and DOCNO=''");
                                if (tmp != null)
                                {
                                    if (!tmp.IsNull("ECT") && tmp["ECT"].ToString() != "") { t02 = tmp["ECT"].ToString(); }
                                    if (!tmp.IsNull("UpperCT") && tmp["UpperCT"].ToString() != "") { t04 = tmp["UpperCT"].ToString(); }
                                    if (!tmp.IsNull("Custom_LowerCT") && tmp["Custom_LowerCT"].ToString() != "") { t03 = tmp["Custom_LowerCT"].ToString(); }
                                    if (t01 != "" && t03 != "" && t01 != "0" && t03 != "0") { t05 = $"{(float.Parse(t03) / float.Parse(t01) * 100).ToString("0.00")}%"; }
                                }
                                if (t01 != "")
                                {
                                    TimeSpan web_DIS = new TimeSpan(0, 0, int.Parse(t01));
                                    t01 = $"{(int)web_DIS.TotalHours}:{web_DIS.Minutes}:{web_DIS.Seconds}";
                                }
                                if (t02 != "")
                                {
                                    TimeSpan web_DIS = new TimeSpan(0, 0, int.Parse(t02));
                                    t02 = $"{(int)web_DIS.TotalHours}:{web_DIS.Minutes}:{web_DIS.Seconds}";
                                }
                                if (t03 != "")
                                {
                                    TimeSpan web_DIS = new TimeSpan(0, 0, int.Parse(t03));
                                    t03 = $"{(int)web_DIS.TotalHours}:{web_DIS.Minutes}:{web_DIS.Seconds}";
                                }
                                if (t04 != "")
                                {
                                    TimeSpan web_DIS = new TimeSpan(0, 0, int.Parse(t04));
                                    t04 = $"{(int)web_DIS.TotalHours}:{web_DIS.Minutes}:{web_DIS.Seconds}";
                                }
                                re = $"{re}<tr><td>目前Cycle:{t01}</td><td>有效Cycle:{t02}</td><td>最佳Cycle:{t03}</td><td>目標Cycle:{t04}</td><td>達成率:{t05}</td></tr>";
                                break;
                            }
                            else
                            {
                                tmp = db.DB_GetFirstDataByDataRow($"SELECT (sum(CycleTime)/count(*)) as T01 FROM SoftNetLogDB.[dbo].[SFC_StationDetail] where SimulationId='{sid.Key}'");
                                if (tmp != null && !tmp.IsNull("T01") && tmp["T01"].ToString() != "") { t01 = tmp["T01"].ToString(); }
                                tmp = db.DB_GetFirstDataByDataRow($"select AVG(EfficientCycleTime) as ECT,AVG(SD_UpperLimit) as UpperCT,AVG(Custom_SD_LowerLimit) as Custom_LowerCT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and PartNO='{sid.Value[0]}' and StationNO='{obj.Key}' and PP_Name='{sid.Value[1]}' and DOCNO=''");
                                if (tmp != null)
                                {
                                    if (!tmp.IsNull("ECT") && tmp["ECT"].ToString() != "") { t02 = tmp["ECT"].ToString(); }
                                    if (!tmp.IsNull("UpperCT") && tmp["UpperCT"].ToString() != "") { t04 = tmp["UpperCT"].ToString(); }
                                    if (!tmp.IsNull("Custom_LowerCT") && tmp["Custom_LowerCT"].ToString() != "") { t03 = tmp["Custom_LowerCT"].ToString(); }
                                    if (t01 != "" && t03 != "" && t01 != "0" && t03 != "0") { t05 = $"{(float.Parse(t03) / float.Parse(t01)*100).ToString("0.00")}%"; }
                                }
                                if (t01 != "")
                                {
                                    TimeSpan web_DIS = new TimeSpan(0, 0, int.Parse(t01));
                                    t01 = $"{(int)web_DIS.TotalHours}:{web_DIS.Minutes}:{web_DIS.Seconds}";
                                }
                                if (t02 != "")
                                {
                                    TimeSpan web_DIS = new TimeSpan(0, 0, int.Parse(t02));
                                    t02 = $"{(int)web_DIS.TotalHours}:{web_DIS.Minutes}:{web_DIS.Seconds}";
                                }
                                if (t03 != "")
                                {
                                    TimeSpan web_DIS = new TimeSpan(0, 0, int.Parse(t03));
                                    t03 = $"{(int)web_DIS.TotalHours}:{web_DIS.Minutes}:{web_DIS.Seconds}";
                                }
                                if (t04 != "")
                                {
                                    TimeSpan web_DIS = new TimeSpan(0, 0, int.Parse(t04));
                                    t04 = $"{(int)web_DIS.TotalHours}:{web_DIS.Minutes}:{web_DIS.Seconds}";
                                }
                                re = $"{re}<tr><td>{(++i).ToString()}.目前Cycle:{t01}</td><td>有效Cycle:{t02}</td><td>最佳Cycle:{t03}</td><td>目標Cycle:{t04}</td><td>達成率:{t05}</td></tr>";
                            }
                        }
                    }

                    if (errTime != "")
                    { re = $"{re}<tr><td>最新警示干涉:{errTime}</td><td colspan='4'>{errLOG}</td></tr>"; }



                    re = $"{re}</table><br />";
                }
            }
            return re;
        }
        public string ChangeDisplayDIV() //採購確認  0=ipport,1.需求碼1,需求碼2....
        {
            //###???暫時寫死單別與廠商 AA02,單價
            string meg = SelectStationStateII();
            return meg;
        }
        public IActionResult ALLStationDetail()
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow tmp = null;
                string rdate = "";
                string rdateII = "";
                string rTime = "";
                string rStationNO = "";
                string rPartNO = "";
                string rOP_NO = "";
                int ct = 0;
                int ect = 0;
                string CT = "";
                string ECT = "";
                string color = "";
                string indexNO = "";

                re = "最近14天報工紀錄";
                re = $"{re}<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                re = $"{re}<tr><td>日期</td><td>時間</td><td>工站</td><td>料號/品名/規格</td><td>工序</td><td>OK數量</td><td>NG數量</td><td>OK前期</td><td>NG前期</td><td>作業人員</td><td>實際CT</td><td>目標CT</td></tr>";
                DataTable dt = db.DB_GetData($@"SELECT b.PartNO,b.Source_StationNO_IndexSN,b.Source_StationNO_Custom_IndexSN,b.Source_StationNO_Custom_DisplayName,a.* FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] as a
                                            join SoftNetSYSDB.[dbo].[APS_Simulation] as b on b.SimulationId=a.SimulationId and b.ServerId='{_Fun.Config.ServerId}'
                                            join SoftNetMainDB.[dbo].[Material] as c on c.PartNO=b.PartNO and c.ServerId='{_Fun.Config.ServerId}'
                                            where a.ServerId='{_Fun.Config.ServerId}' and a.LOGDateTimeID>='{DateTime.Now.AddDays(-14).ToString("yyyy/MM/dd HH:mm:ss")}' and a.LOGDateTimeID<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}' order by a.LOGDateTimeID desc,a.StationNO,a.OP_NO,b.PartNO");
                foreach (DataRow dr in dt.Rows)
                {
                    if (rdate != Convert.ToDateTime(dr["LOGDateTimeID"]).ToString("yyyy-MM-dd"))
                    {
                        rdate = Convert.ToDateTime(dr["LOGDateTimeID"]).ToString("yyyy-MM-dd");
                        rdateII = rdate;
                    }
                    else { rdateII = ""; }
                    rTime = Convert.ToDateTime(dr["LOGDateTimeID"]).ToString("HH:mm");
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT StationNO,StationName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                    if (tmp != null) { rStationNO = $"{tmp["StationNO"].ToString()}&nbsp;{tmp["StationName"].ToString()}"; } else { rStationNO = ""; }
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO,PartName,Specification from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                    if (tmp != null) { rPartNO = $"{tmp["PartNO"].ToString()}&nbsp;{tmp["PartName"].ToString()}&nbsp;{tmp["Specification"].ToString()}"; } else { rPartNO = ""; }
                    if (!dr.IsNull("OP_NO") && dr["OP_NO"].ToString() != "")
                    {
                        if (dr["OP_NO"].ToString().IndexOf(";") >= 0) { rOP_NO = "多人報工"; }
                        else
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT Name from SoftNetMainDB.[dbo].[User] where ServerId='{_Fun.Config.ServerId}' and UserNO='{dr["OP_NO"].ToString()}'");
                            if (tmp != null) { rOP_NO = $"{tmp["Name"].ToString()}"; } else { rPartNO = ""; }
                        }
                    }
                    else { rOP_NO = ""; }
                    if (!dr.IsNull("Source_StationNO_Custom_DisplayName") && dr["Source_StationNO_Custom_DisplayName"].ToString() != "")
                    { indexNO = dr["Source_StationNO_Custom_DisplayName"].ToString(); }
                    else if (!dr.IsNull("Source_StationNO_Custom_IndexSN") && dr["Source_StationNO_Custom_IndexSN"].ToString() != "")
                    { indexNO = dr["Source_StationNO_Custom_IndexSN"].ToString(); }
                    else { indexNO = dr["Source_StationNO_IndexSN"].ToString(); }
                    ct = int.Parse(dr["CycleTime"].ToString());
                    TimeSpan standardTime_DIS = new TimeSpan(0, 0, ct);
                    CT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                    ect = int.Parse(dr["ECT"].ToString());
                    standardTime_DIS = new TimeSpan(0, 0, ect);
                    ECT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                    if (ct > ect && ect!=0) { color = " style='background-color:gold;'"; } else { color = ""; }
                    re = $"{re}<tr><td>{rdateII}</td><td>{rTime}</td><td>{rStationNO}</td><td>{rPartNO}</td><td>{indexNO}</td><td>{dr["EditFinishedQty"].ToString()}</td><td>{dr["EditFailedQty"].ToString()}</td><td>{dr["OLD_ProductFinishedQty"].ToString()}</td><td>{dr["OLD_ProductFailedQty"].ToString()}</td><td>{rOP_NO}</td><td{color}>{CT}</td><td>{ECT}</td></tr>";
                }
                re = $"{re}</table>";
            }
            ViewBag.HtmlOutput = re;
            return View();
        }
        public IActionResult ALLOperateLog()
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow tmp = null;
                string rStationNO = "";
                string rPartNO = "";
                string rOP_NO = "";

                re = "最近14天操作紀錄";
                re = $"{re}<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                re = $"{re}<tr><td>日期</td><td>操作</td><td>工站</td><td>工單編號</td><td>料號/品名/規格</td><td>操作人員</td><td>備註</td></tr>";
                DataTable dt = db.DB_GetData($@"SELECT b.PartName,b.Specification,a.* FROM SoftNetLogDB.[dbo].[OperateLog] as a
                                            join SoftNetMainDB.[dbo].[Material] as b on b.PartNO=a.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                                            where a.ServerId='{_Fun.Config.ServerId}' and a.LOGDateTime>='{DateTime.Now.AddDays(-14).ToString("yyyy/MM/dd HH:mm:ss.fff")}' and a.LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}' order by a.LOGDateTime desc,a.StationNO,a.OP_NO,a.PartNO");
                foreach (DataRow dr in dt.Rows)
                {
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT StationNO,StationName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                    if (tmp != null) { rStationNO = $"{tmp["StationNO"].ToString()}&nbsp;{tmp["StationName"].ToString()}"; } else { rStationNO = ""; }
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO,PartName,Specification from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                    if (tmp != null) { rPartNO = $"{tmp["PartNO"].ToString()}&nbsp;{tmp["PartName"].ToString()}&nbsp;{tmp["Specification"].ToString()}"; } else { rPartNO = ""; }
                    if (!dr.IsNull("OP_NO") && dr["OP_NO"].ToString() != "")
                    {
                        if (dr["OP_NO"].ToString().IndexOf(";") >= 0) { rOP_NO = "多人報工"; }
                        else
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT Name from SoftNetMainDB.[dbo].[User] where ServerId='{_Fun.Config.ServerId}' and UserNO='{dr["OP_NO"].ToString()}'");
                            if (tmp != null) { rOP_NO = $"{tmp["Name"].ToString()}"; } else { rPartNO = ""; }
                        }
                    }
                    else { rOP_NO = ""; }

                    re = $"{re}<tr><td>{dr["LOGDateTime"].ToString()}</td><td>{dr["OperateType"].ToString()}</td><td>{rStationNO}</td><td>{dr["OrderNO"].ToString()}</td><td>{rPartNO}</td><td>{rOP_NO}</td><td></td></tr>";
                }
                re = $"{re}</table>";
            }
            ViewBag.HtmlOutput = re;
            return View();
        }
        public IActionResult ALLLabelLog(APSViewData key)
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                if (key == null || key.SelectFun1 == "")
                {
                    key = new APSViewData();
                    key.SelectFun2 = "-7";
                    key.SelectFun3 = "4";
                    key.SelectFun4 = "1";
                }
                DataRow tmp = null;
                string rStationNO = "";
                string rPartNO = "";
                string rOP_NO = "";

                int betime = -1;
                int.TryParse(key.SelectFun2, out betime);
                string wheresql = "";
                if (key.SelectFun3 == "1") { wheresql = $" and (a.ActionType like '%刪除計畫%' or a.ActionType like '%開工%' or a.ActionType like '%設定工單%' or a.ActionType like '%停工%' or a.ActionType like '%關站%' or a.ActionType like '%干涉派工%' or a.ActionType like '%初始化%')"; }
                else if (key.SelectFun3 == "2") { wheresql = $" and a.ActionType like '%倉庫%'"; }
                else if (key.SelectFun3 == "3") { wheresql = $" and a.ActionType like '%儲位%'"; }
                else if (key.SelectFun3 == "4") { wheresql = $" and (a.ActionType like '%Fail%' or a.ActionType like '%失敗%')"; }
                else { wheresql = ""; }

                re = $"{re}<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                re = $"{re}<tr><td>日期</td><td>電子編號</td><td>編號型態</td><td>動作</td><td>接收</td><td>回饋訊號</td><td>內容</td></tr>";
                DataTable dt = db.DB_GetData($@"SELECT a.*,b.Type,b.StationNO,b.StoreNO,b.StoreSpacesNO FROM [SoftNetLogDB].[dbo].[LabelStateLog] as a
                                                join SoftNetMainDB.[dbo].[LabelStateINFO] as b on a.macID=b.macID
                                                where a.LOGDateTime>='{DateTime.Now.AddDays(betime).ToString("yyyy/MM/dd HH:mm:ss.fff")}' {wheresql} order by a.LOGDateTime desc");
                if (dt != null && dt.Rows.Count > 0)
                {
                    string disType = "";
                    foreach (DataRow dr in dt.Rows)
                    {
                        
                        if (dr["Type"].ToString() == "1" || dr["Type"].ToString() == "4") { disType = dr["StationNO"].ToString(); }
                        else if (dr["Type"].ToString() == "2") { disType = $"倉庫:{dr["StoreNO"].ToString()}"; }
                        else if (dr["Type"].ToString() == "3") { disType = $"倉庫:{dr["StoreNO"].ToString()} 儲位:{dr["StoreSpacesNO"].ToString()}"; }
                        else { disType = ""; }
                        if (key.SelectFun4 == "1")
                        { re = $"{re}<tr><td>{dr["LOGDateTime"].ToString()}</td><td>{dr["macID"].ToString()}</td><td>{disType}</td><td>{dr["ActionType"].ToString()}</td><td>{dr["ReceiveType"].ToString()}</td><td>{dr["INFO"].ToString()}</td><td></td></tr>"; }
                        else { re = $"{re}<tr><td>{dr["LOGDateTime"].ToString()}</td><td>{dr["macID"].ToString()}</td><td>{disType}</td><td>{dr["ActionType"].ToString()}</td><td>{dr["ReceiveType"].ToString()}</td><td>{dr["INFO"].ToString()}</td><td>{dr["JSON"].ToString()}</td></tr>"; }
                    }
                }
                re = $"{re}</table>";
            }
            ViewBag.HtmlOutput = re;
            return View();
        }
        public IActionResult ALLMFData()
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow tmp = null;
                string storeName = "";
                re = $"<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                re = $"{re}<tr><td>廠商編號</td><td>廠商名稱</td><td>統一編號</td><td>聯絡人</td><td>聯絡TEL</td><td>單據收發EMail</td><td>入庫倉別</td><td>排程權重</td><td>地址</td><td>備註</td></tr>";
                DataTable dt = db.DB_GetData($"select * FROM SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' order by MFNO,MFName,TEL,ContactTEL,CTDataWeights desc");
                foreach (DataRow dr in dt.Rows)
                {
                    storeName = "";
                    if (!dr.IsNull("StoreNO"))
                    {
                        tmp = db.DB_GetFirstDataByDataRow($"SELECT StoreName from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr["StoreNO"].ToString()}'");
                        if (tmp != null) { storeName = tmp["StoreName"].ToString(); }
                    }
                    re = $"{re}<tr><td>{dr["MFNO"].ToString()}</td><td>{dr["MFName"].ToString()}</td><td>{dr["UniFormNO"].ToString()}</td><td>{dr["ContactMan"].ToString()}</td><td>{dr["ContactTEL"].ToString()}</td><td>{dr["EMail"].ToString()}</td><td>{dr["StoreNO"].ToString()}{storeName}</td><td>{dr["CTDataWeights"].ToString()}</td><td>{dr["Address"].ToString()}</td><td>{dr["Remark"].ToString()}</td></tr>";
                }
                re = $"{re}</table>";
            }
            ViewBag.HtmlOutput = re;
            return View();
        }
        public IActionResult ALLStationOrderNO()
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = new DataTable();
                tmp_dt.Columns.Add("StationNO");
                tmp_dt.Columns.Add("StationNOName");
                tmp_dt.Columns.Add("OrderNO");
                tmp_dt.Columns.Add("State");
                tmp_dt.Columns.Add("PartNO");
                tmp_dt.Columns.Add("PartNOName");
                tmp_dt.Columns.Add("QTY");
                tmp_dt.Columns.Add("Unit");
                tmp_dt.Columns.Add("FinishedQty");
                tmp_dt.Columns.Add("FailedQty");
                tmp_dt.Columns.Add("BeforQTY");
                tmp_dt.Columns.Add("NetStationQTY");
                tmp_dt.Columns.Add("INStoreQTY");

                DataRow tmp = null;
                DataRow tmp2 = null;
                DataRow tmp3 = null;
                DataRow tmp4 = null;
                string state = "";
                string beforStation = "";
                string disStaion = "";
                string disbeforStaion = "";
                string disbeforStaionQTY = "";
                re = $"<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                re = $"{re}<tr><td>工站編號</td><td>工站名稱</td><td>在線工單編號</td><td>狀態</td><td>料件編號</td><td>品名/規格</td><td>需求數量</td><td>單位</td><td>上站工站</td><td>移轉入量</td><td>本站報工量</td><td>下站工站</td><td>移出下站量</td><td>入庫量</td></tr>";
                DataTable dt = db.DB_GetData($"select * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and OrderNO!='' order by StationNO");
                string fontClass = "";
                foreach (DataRow dr in dt.Rows)
                {
                    fontClass = " style='background-color:whitesmoke;'";
                    switch (dr["State"].ToString())
                    {
                        case "1": state = "開工中"; fontClass = " style='background-color:springgreen;'"; break;
                        case "2": state = "停工"; fontClass = " style='background-color:tomato;'"; break;
                        default: state = "閒置"; break;
                    }
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                    if (tmp != null)
                    {
                        disbeforStaion = "";
                        disbeforStaionQTY = "";
                        #region 計算前站完成量
                        tmp4 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                        if (tmp4 != null)
                        {
                            tmp4 = db.DB_GetFirstDataByDataRow($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation] where Source_StationNO is not NULL and NeedId='{tmp4["NeedId"].ToString()}' and Apply_PP_Name='{tmp4["Apply_PP_Name"].ToString()}' and IndexSN={(int.Parse(tmp4["IndexSN"].ToString()) - 1).ToString()} and Apply_StationNO='{tmp4["Source_StationNO"].ToString()}'");
                            if (tmp4 != null)
                            {
                                tmp4 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].APS_PartNOTimeNote where SimulationId='{tmp4["SimulationId"].ToString()}' and NoStation='0'");
                                if (tmp4 != null)
                                {
                                    disbeforStaion= tmp4["APS_StationNO"].ToString();
                                    disbeforStaionQTY = tmp4["Detail_QTY"].ToString();
                                }
                            }
                        }
                        #endregion

                        tmp2 = db.DB_GetFirstDataByDataRow($"SELECT StationName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                        tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr["SimulationId"].ToString()}'");
                        if (tmp3 != null)
                        {
                            re = $"{re}<tr><td>{dr["StationNO"].ToString()}</td><td>{tmp2["StationName"].ToString()}</td><td>{dr["OrderNO"].ToString()}</td><td{fontClass}>{state}</td><td>{dr["PartNO"].ToString()}</td><td>{tmp["PartName"].ToString()}&nbsp;{tmp["Specification"].ToString()}</td><td>{tmp3["NeedQTY"].ToString()}</td><td>{tmp["Unit"].ToString()}</td><td>{disbeforStaion}</td><td>{disbeforStaionQTY}</td><td>{tmp3["Detail_QTY"].ToString()}/{tmp3["Detail_Fail_QTY"].ToString()}</td><td>{tmp3["Next_APS_StationNO"].ToString()}</td><td>{tmp3["Next_StationQTY"].ToString()}</td><td>{tmp3["Next_StoreQTY"].ToString()}</td></tr>";
                        }
                    }
                }

                dt = db.DB_GetData($"select * FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and EndTime is NULL order by StationNO");
                foreach (DataRow dr in dt.Rows)
                {
                    fontClass = " style='background-color:whitesmoke;'";
                    if (!dr.IsNull("RemarkTimeS") && dr.IsNull("RemarkTimeE")) { state = "開工中"; fontClass = " style='background-color:springgreen;'"; }
                    else if (!dr.IsNull("RemarkTimeE")) { state = "停工"; fontClass = " style='background-color:tomato;'"; }
                    else { state = "閒置"; }
                    if (beforStation != dr["StationNO"].ToString())
                    {
                        beforStation = dr["StationNO"].ToString();
                        disStaion = beforStation;
                    }
                    else { disStaion = ""; }

                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                    if (tmp != null)
                    {
                        disbeforStaion = "";
                        disbeforStaionQTY = "";
                        #region 計算前站完成量
                        tmp4 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                        if (tmp4 != null)
                        {
                            tmp4 = db.DB_GetFirstDataByDataRow($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation] where Source_StationNO is not NULL and NeedId='{tmp4["NeedId"].ToString()}' and Apply_PP_Name='{tmp4["Apply_PP_Name"].ToString()}' and IndexSN={(int.Parse(tmp4["IndexSN"].ToString()) - 1).ToString()} and Apply_StationNO='{tmp4["Source_StationNO"].ToString()}'");
                            if (tmp4 != null)
                            {
                                tmp4 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].APS_PartNOTimeNote where SimulationId='{tmp4["SimulationId"].ToString()}' and NoStation='0'");
                                if (tmp4 != null)
                                {
                                    disbeforStaion = tmp4["APS_StationNO"].ToString();
                                    disbeforStaionQTY = tmp4["Detail_QTY"].ToString();
                                }
                            }
                        }
                        #endregion

                        tmp2 = db.DB_GetFirstDataByDataRow($"SELECT StationName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                        tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr["SimulationId"].ToString()}'");
                        if (tmp3 != null)
                        {
                            re = $"{re}<tr><td>{disStaion}</td><td>{tmp2["StationName"].ToString()}</td><td>{dr["OrderNO"].ToString()}</td><td{fontClass}>{state}</td><td>{dr["PartNO"].ToString()}</td><td>{tmp["PartName"].ToString()}&nbsp;{tmp["Specification"].ToString()}</td><td>{tmp3["NeedQTY"].ToString()}</td><td>{tmp["Unit"].ToString()}</td><td>{disbeforStaion}</td><td>{disbeforStaionQTY}</td><td>{tmp3["Detail_QTY"].ToString()}/{tmp3["Detail_Fail_QTY"].ToString()}</td><td>{tmp3["Next_APS_StationNO"].ToString()}</td><td>{tmp3["Next_StationQTY"].ToString()}</td><td>{tmp3["Next_StoreQTY"].ToString()}</td></tr>";
                        }
                    }
                }

                re = $"{re}</table>";
            }
            ViewBag.HtmlOutput = re;
            return View();
        }
        public IActionResult AllMaterial()
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow tmp = null;
                string storeName = "";
                string className = "";
                string typeName = "";
                string iS_Store_Test = "";
                string aPS_Default_MFNO = "";
                re = $"<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                re = $"{re}<tr><td>料件編號</td><td>品名</td><td>規格</td><td>機種</td><td>類型</td><td>安全量</td><td>單位</td><td>備料工時</td><td>檢驗</td><td>指定供應商</td><td>指定倉庫</td><td>定義</td></tr>";
                DataTable dt = db.DB_GetData($"select * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' order by Class,PartNO,APS_Default_StoreNO,APS_Default_StoreSpacesNO,PartType");
                foreach (DataRow dr in dt.Rows)
                {
                    storeName = "";
                    className = "";
                    typeName = "";
                    iS_Store_Test = "";
                    aPS_Default_MFNO = "";
                    switch (dr["Class"].ToString())
                    {
                        case "1": className = "原物料"; break;
                        case "2": className = "採購件"; break;
                        case "3": className = "委外件"; break;
                        case "4": className = "製造半成品"; break;
                        case "5": className = "製造成品"; break;
                        case "6": className = "刀具"; break;
                        case "7": className = "工具製具"; break;
                    }
                    if (dr["PartType"].ToString()=="0")
                    { typeName = "實料"; }
                    else { typeName = "虛料"; }
                    if (bool.Parse(dr["IS_Store_Test"].ToString()))
                    { iS_Store_Test = "是"; }
                    else { iS_Store_Test = "否"; }
                    if (!dr.IsNull("APS_Default_StoreNO"))
                    {
                        tmp = db.DB_GetFirstDataByDataRow($"SELECT StoreName from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr["APS_Default_StoreNO"].ToString()}'");
                        if (tmp != null) {  storeName = $"{tmp["StoreName"].ToString()}&nbsp;{dr["APS_Default_StoreSpacesNO"].ToString()}"; }
                    }
                    if (!dr.IsNull("APS_Default_MFNO"))
                    {
                        tmp = db.DB_GetFirstDataByDataRow($"SELECT SName from SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' and MFNO='{dr["APS_Default_MFNO"].ToString()}'");
                        if (tmp != null) { aPS_Default_MFNO = tmp["SName"].ToString(); }
                    }
                    re = $"{re}<tr><td>{dr["PartNO"].ToString()}</td><td>{dr["PartName"].ToString()}</td><td>{dr["Specification"].ToString()}</td><td>{dr["Model"].ToString()}</td><td>{className}</td><td>{dr["SafeQTY"].ToString()}</td><td>{dr["Unit"].ToString()}</td><td>{dr["StoreSTime"].ToString()}</td><td>{iS_Store_Test}</td><td>{dr["APS_Default_MFNO"].ToString()}</td><td>{storeName}</td><td>{typeName}</td></tr>";
                }
                re = $"{re}</table>";
            }
            ViewBag.HtmlOutput = re;
            return View();
        }

        public IActionResult ALLStore(APSViewData key)
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                if (key == null || key.SelectFun1 == "")
                {
                    key = new APSViewData();
                    key.SelectFun2 = "4";
                    key.SelectFun3 = "";
                }
                DataRow tmp = null;
                DataTable dt_tmp = null;
                string partNOName = "";
                string className = "";
                int sQTY = 0;
                int keepQTY = 0;
                int pQTY = 0;
                int gQTY = 0;
                int useQTY = 0;
                string wheresql = "";
                foreach(string s in key.SelectFun2.Split(';'))
                {
                    if (s == "") { continue; }
                    if (wheresql == "") { wheresql = $" and (Class='{s}'"; }
                    else { wheresql = $"{wheresql} or Class='{s}'"; }
                }
                if (wheresql != "") { wheresql = $"{wheresql})"; }
                re = $"<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                re = $"{re}<tr><td>料件編號</td><td>品名/規格</td><td>類型</td><td>安全量</td><td>單位</td><td>再倉量</td><td>再製量</td><td>再途量</td><td>保留量</td><td>未來量</td></tr>";
                DataTable dt = db.DB_GetData($"select * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' {wheresql} order by Class,PartNO,APS_Default_StoreNO,APS_Default_StoreSpacesNO,PartType");
                foreach (DataRow dr in dt.Rows)
                {
                    sQTY = 0; keepQTY = 0; pQTY = 0; gQTY = 0;
                    partNOName = $"{dr["PartName"].ToString()}&nbsp;{dr["Specification"].ToString()}";
                    switch (dr["Class"].ToString())
                    {
                        case "1": className = "原物料"; break;
                        case "2": className = "採購件"; break;
                        case "3": className = "委外件"; break;
                        case "4": className = "製造半成品"; break;
                        case "5": className = "製造成品"; break;
                        case "6": className = "刀具"; break;
                        case "7": className = "工具製具"; break;
                        default: className = ""; break;
                    }
                    #region 再倉量 sQTY
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(QTY) as S_QTY from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                    if (tmp != null && !tmp.IsNull("S_QTY") && tmp["S_QTY"].ToString() != "") { sQTY = int.Parse(tmp["S_QTY"].ToString()); }
                    #endregion

                    #region 保留量 keepQTY
                    tmp = db.DB_GetFirstDataByDataRow($@"SELECT sum(b.KeepQTY+b.OverQTY) as Keep_QTY from SoftNetMainDB.[dbo].[TotalStock] as a
                                                        join SoftNetMainDB.[dbo].[TotalStockII] as b on a.Id=b.Id where a.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{dr["PartNO"].ToString()}'");
                    if (tmp != null && !tmp.IsNull("Keep_QTY") && tmp["Keep_QTY"].ToString() != "") { keepQTY = int.Parse(tmp["Keep_QTY"].ToString()); }
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(QTY) as DOC3_Out_QTY from SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}' and OUT_StoreNO!='' and IsOK='0'");
                    if (tmp != null && !tmp.IsNull("DOC3_Out_QTY") && tmp["DOC3_Out_QTY"].ToString() != "") { keepQTY += int.Parse(tmp["DOC3_Out_QTY"].ToString()); }
                    #endregion

                    #region 再製量 pQTY
                    dt_tmp = db.DB_GetData($@"select b.NeedId,b.NeedQTY,a.* from SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                            join SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as b on a.SimulationId=b.SimulationId
                                            where a.IsOK='0' and (a.Class='4' or a.Class='5') and a.PartNO='{dr["PartNO"].ToString()}' and a.Source_StationNO!='{_Fun.Config.OutPackStationName}' and a.Source_StationNO is not null order by a.NeedId,a.PartSN desc");
                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                    {
                        string bpartNO = dt_tmp.Rows[0]["PartNO"].ToString();
                        string bMpartNO = dt_tmp.Rows[0]["Master_PartNO"].ToString();
                        pQTY = int.Parse(dt_tmp.Rows[0]["NeedQTY"].ToString());
                        foreach (DataRow dr2 in dt_tmp.Rows)
                        {
                            if (bpartNO == dr2["PartNO"].ToString() && bMpartNO == dr2["Master_PartNO"].ToString()) { continue; }
                            else
                            {
                                pQTY += int.Parse(dr2["NeedQTY"].ToString());
                            }
                        }
                    }
                    #endregion

                    #region 在途量 gQTY
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(QTY) as G_QTY from SoftNetMainDB.[dbo].[DOC1BuyII] where ServerId='{_Fun.Config.ServerId}' and IsOK='0' and PartNO='{dr["PartNO"].ToString()}'");
                    if (tmp != null && !tmp.IsNull("G_QTY") && tmp["G_QTY"].ToString() != "") { gQTY = int.Parse(tmp["G_QTY"].ToString()); }
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(QTY) as G_QTY from SoftNetMainDB.[dbo].[DOC4ProductionII] where ServerId='{_Fun.Config.ServerId}' and IsOK='0' and PartNO='{dr["PartNO"].ToString()}'");
                    if (tmp != null && !tmp.IsNull("G_QTY") && tmp["G_QTY"].ToString() != "") { gQTY += int.Parse(tmp["G_QTY"].ToString()); }
                    #endregion

                    useQTY = sQTY - keepQTY + pQTY + gQTY;
                    if (sQTY != 0 && pQTY != 0 && gQTY != 0 && keepQTY != 0 && useQTY != 0)
                    {
                        re = $"{re}<tr><td>{dr["PartNO"].ToString()}</td><td>{partNOName}</td><td>{className}</td><td>{dr["SafeQTY"].ToString()}</td><td>{dr["Unit"].ToString()}</td><td>{sQTY.ToString()}</td><td>{pQTY.ToString()}</td><td>{gQTY.ToString()}</td><td>{keepQTY.ToString()}</td><td>{useQTY.ToString()}</td></tr>";
                    }
                }
                re = $"{re}</table>";
            }
            ViewBag.HtmlOutput = re;
            return View();
        }
        public IActionResult ALLQTYMaterial(APSViewData key)
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DateTime now = DateTime.Now;
                if (key == null || key.SelectFun1 == "")
                {
                    key = new APSViewData();
                    key.SelectFun1 = now.AddMonths(-1).ToString("yyyy-MM-dd");
                    key.SelectFun2 = now.ToString("yyyy-MM-dd");
                    key.SelectFun5 = now.ToString("yyyy-MM-dd");
                    return View(key);
                }
                if (key.SelectFun3 == null || key.SelectFun3 == "") { ViewBag.ERRMsg = "沒有選擇 顯示料件類型"; return View(key); }

                ViewBag.ERRMsg = "";
                DataRow tmp = null;
                DataRow tmpII = null;
                string className = "";

                string wheresql = "";
                foreach (string s in key.SelectFun3.Split(';'))
                {
                    if (s == "") { continue; }
                    if (wheresql == "") { wheresql = $" and (b.Class='{s}'"; }
                    else { wheresql = $"{wheresql} or b.Class='{s}'"; }
                }
                if (wheresql != "") { wheresql = $"{wheresql})"; }
                if (key.SelectFun4!=null && key.SelectFun4 != "") { wheresql = $"{wheresql} and (a.PartNO like '%{key.SelectFun4}%' or b.PartName like '%{key.SelectFun4}%' or b.Specification like '%{key.SelectFun4}%')"; }
                re = $"<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                re = $"{re}<tr><td>日期</td><td>單據編號</td><td>料件編號</td><td>品名/規格</td><td>類型</td><td>安全量</td><td>單位</td><td>狀態</td><td>倉庫/儲位</td><td>進倉量</td><td>出倉量</td><td>再途</td><td>帳面數量</td></tr>";
                
                #region 收集資料
                string sql = $@"SELECT '-' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],OUT_StoreNO as StoreNO,OUT_StoreSpacesNO as StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and SUBSTRING(a.DOCNumberNO,1,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{key.SelectFun1}' and StartTime<='{key.SelectFun2}' and OUT_StoreNO!='' and a.StartTime is not NULL {wheresql} order by a.PartNO,a.StartTime desc";
                DataTable dt = db.DB_GetData(sql);
                sql = $@"SELECT '+' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],IN_StoreNO as StoreNO,IN_StoreSpacesNO as StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and SUBSTRING(a.DOCNumberNO,1,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{key.SelectFun1}' and StartTime<='{key.SelectFun2}' and IN_StoreNO!='' and a.StartTime is not NULL {wheresql} order by a.PartNO,a.StartTime desc";
                DataTable dtII = db.DB_GetData(sql);
                if (dtII != null && dtII.Rows.Count > 0)
                {
                    if (dt != null) { dt.Merge(dtII); } else { dt = dtII; }
                }
                sql = $@"SELECT '+' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],StoreNO,StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC1BuyII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and SUBSTRING(a.DOCNumberNO,1,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{key.SelectFun1}' and StartTime<='{key.SelectFun2}' and StoreNO!='' and a.StartTime is not NULL {wheresql} order by a.PartNO,a.StartTime desc";
                dtII = db.DB_GetData(sql);
                if (dtII != null && dtII.Rows.Count > 0)
                {
                    if (dt != null) { dt.Merge(dtII); } else { dt = dtII; }
                }
                sql = $@"SELECT '-' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],StoreNO,StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC2SalesII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and SUBSTRING(a.DOCNumberNO,1,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{key.SelectFun1}' and StartTime<='{key.SelectFun2}' and StoreNO!='' and a.StartTime is not NULL {wheresql} order by a.PartNO,a.StartTime desc";
                dtII = db.DB_GetData(sql);
                if (dtII != null && dtII.Rows.Count > 0)
                {
                    if (dt != null) { dt.Merge(dtII); } else { dt = dtII; }
                }
                sql = $@"SELECT '@' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],StoreNO,StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC4ProductionII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and SUBSTRING(a.DOCNumberNO,1,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{key.SelectFun1}' and StartTime<='{key.SelectFun2}' and StoreNO!='' and a.StartTime is not NULL {wheresql} order by a.PartNO,a.StartTime desc";
                dtII = db.DB_GetData(sql);
                if (dtII != null && dtII.Rows.Count > 0)
                {
                    if (dt != null) { dt.Merge(dtII); } else { dt = dtII; }
                }
                sql = $@"SELECT '-' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],StoreNO,StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC5OUTII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and SUBSTRING(a.DOCNumberNO,1,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{key.SelectFun1}' and StartTime<='{key.SelectFun2}' and StoreNO!='' and a.StartTime is not NULL {wheresql} order by a.PartNO,a.StartTime desc";
                dtII = db.DB_GetData(sql);
                if (dtII != null && dtII.Rows.Count > 0)
                {
                    if (dt != null) { dt.Merge(dtII); } else { dt = dtII; }
                }
                #endregion
                string dis_tmp = "";
                if (dt != null && dt.Rows.Count > 0)
                {
                    dt.DefaultView.Sort = "PartNO,StartTime desc";
                    string partNO = "";
                    int fishQTY = 0;
                    
                    #region first
                    partNO = dt.Rows[0]["PartNO"].ToString();
                    #region 計算期末數量
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(QTY) as FishQTY FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dt.Rows[0]["PartNO"].ToString()}'");
                    if (tmp != null && !tmp.IsNull("FishQTY") && tmp["FishQTY"].ToString() != "")
                    {
                        if (Convert.ToDateTime(key.SelectFun2).ToString("yyyy-MM-dd") == now.ToString("yyyy-MM-dd"))
                        { fishQTY = int.Parse(tmp["FishQTY"].ToString()); }
                        else
                        {
                            #region 反算期末數量
                            sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and OUT_StoreNO!='' and PartNO='{dt.Rows[0]["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                            tmpII = db.DB_GetFirstDataByDataRow(sql);
                            if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { fishQTY += int.Parse(tmpII["QTY"].ToString()); }
                            sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and IN_StoreNO!='' and PartNO='{dt.Rows[0]["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                            tmpII = db.DB_GetFirstDataByDataRow(sql);
                            if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { fishQTY -= int.Parse(tmpII["QTY"].ToString()); }
                            sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC1BuyII] where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and PartNO='{dt.Rows[0]["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                            tmpII = db.DB_GetFirstDataByDataRow(sql);
                            if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { fishQTY -= int.Parse(tmpII["QTY"].ToString()); }
                            sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC2SalesII] where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and PartNO='{dt.Rows[0]["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                            tmpII = db.DB_GetFirstDataByDataRow(sql);
                            if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { fishQTY += int.Parse(tmpII["QTY"].ToString()); }
                            //sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC4ProductionII] where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and PartNO='{dt.Rows[0]["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                            //tmpII = db.DB_GetFirstDataByDataRow(sql);
                            //if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { tmo_qty += int.Parse(tmpII["QTY"].ToString()); }
                            sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC5OUTII] where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and PartNO='{dt.Rows[0]["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                            tmpII = db.DB_GetFirstDataByDataRow(sql);
                            if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { fishQTY += int.Parse(tmpII["QTY"].ToString()); }
                            #endregion
                        }
                    }
                    #endregion
                    #endregion
                    dis_tmp = $"<tr><td colspan='12' align='right'>結餘再倉量</td><td>{fishQTY.ToString()}</td></tr>";

                    DataRow dr = null;
                    string showQTY_a = "";
                    string showQTY_b = "";
                    int notOKQTY = 0;
                    int qty = 0;
                    string state = "";
                    for (int i = 1;i<dt.Rows.Count;i++)
                    {
                        dr = dt.Rows[i];
                        if (partNO != dr["PartNO"].ToString())
                        {
                            tmp = dt.Rows[(i - 1)];
                            switch (tmp["Class"].ToString())
                            {
                                case "1": className = "原物料"; break;
                                case "2": className = "採購件"; break;
                                case "3": className = "委外件"; break;
                                case "4": className = "製造半成品"; break;
                                case "5": className = "製造成品"; break;
                                case "6": className = "刀具"; break;
                                case "7": className = "工具製具"; break;
                                default: className = ""; break;
                            }
                            qty = int.Parse(tmp["QTY"].ToString());
                            if (tmp["MathType"].ToString() == "-")
                            {
                                showQTY_a = "";
                                showQTY_b = tmp["QTY"].ToString();
                                if (bool.Parse(tmp["IsOK"].ToString()))
                                { fishQTY += qty; }
                                else { notOKQTY += qty; }
                            }
                            else
                            {
                                showQTY_b = "";
                                showQTY_a = tmp["QTY"].ToString();
                                if (bool.Parse(tmp["IsOK"].ToString()))
                                { fishQTY -= qty; }
                                else { notOKQTY -= qty; }
                            }
                            if (bool.Parse(tmp["IsOK"].ToString())) { state = "入帳"; } else { state = "再途"; }
                            re = $"{re}<tr><td>{Convert.ToDateTime(tmp["StartTime"]).ToString("yy-MM-dd HH:mm:ss")}</td><td>{tmp["DOCNumberNO"].ToString()}<br />{tmp["DOCName"].ToString()}</td><td>{tmp["PartNO"].ToString()}</td><td>{tmp["PartName"].ToString()}{tmp["Specification"].ToString()}</td><td>{className}</td><td>{tmp["SafeQTY"].ToString()}</td><td>{tmp["Unit"].ToString()}</td><td>{state}</td><td>{tmp["StoreNO"].ToString()}{tmp["StoreSpacesNO"].ToString()}</td><td>{showQTY_a}</td><td>{showQTY_b}</td><td>{notOKQTY.ToString()}</td><td>{fishQTY.ToString()}</td></tr>{dis_tmp}";
                            partNO = dr["PartNO"].ToString();

                            fishQTY = 0;
                            notOKQTY = 0;
                            #region 計算期末數量
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(QTY) as FishQTY FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                            if (tmp != null && !tmp.IsNull("FishQTY") && tmp["FishQTY"].ToString() != "")
                            {
                                if (Convert.ToDateTime(key.SelectFun2).ToString("yyyy-MM-dd") == now.ToString("yyyy-MM-dd"))
                                { fishQTY = int.Parse(tmp["FishQTY"].ToString()); }
                                else
                                {
                                    #region 反算期末數量
                                    sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and OUT_StoreNO!='' and PartNO='{dr["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                                    tmpII = db.DB_GetFirstDataByDataRow(sql);
                                    if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { fishQTY += int.Parse(tmpII["QTY"].ToString()); }
                                    sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and IN_StoreNO!='' and PartNO='{dr["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                                    tmpII = db.DB_GetFirstDataByDataRow(sql);
                                    if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { fishQTY -= int.Parse(tmpII["QTY"].ToString()); }
                                    sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC1BuyII] where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and PartNO='{dr["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                                    tmpII = db.DB_GetFirstDataByDataRow(sql);
                                    if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { fishQTY -= int.Parse(tmpII["QTY"].ToString()); }
                                    sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC2SalesII] where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and PartNO='{dr["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                                    tmpII = db.DB_GetFirstDataByDataRow(sql);
                                    if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { fishQTY += int.Parse(tmpII["QTY"].ToString()); }
                                    //sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC4ProductionII] where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and PartNO='{dr["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                                    //tmpII = db.DB_GetFirstDataByDataRow(sql);
                                    //if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { tmo_qty += int.Parse(tmpII["QTY"].ToString()); }
                                    sql = $@"SELECT sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC5OUTII] where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and PartNO='{dr["PartNO"].ToString()}' and StartTime is not NULL and StartTime>'{key.SelectFun2}' and StartTime<={now.ToString("yyyy-MM-dd")}";
                                    tmpII = db.DB_GetFirstDataByDataRow(sql);
                                    if (tmpII != null && !tmpII.IsNull("QTY") && tmpII["QTY"].ToString() != "") { fishQTY += int.Parse(tmpII["QTY"].ToString()); }
                                    #endregion
                                }
                            }
                            #endregion
                            dis_tmp = $"<tr><td colspan='12' align='right'>結餘再倉量</td><td>{fishQTY.ToString()}</td></tr>";
                        }
                        else
                        {
                            qty = int.Parse(dr["QTY"].ToString());
                            if (dr["MathType"].ToString() == "-")
                            {
                                showQTY_a = "";
                                showQTY_b = dr["QTY"].ToString();
                                if (bool.Parse(dr["IsOK"].ToString()))
                                { fishQTY += qty; }
                                else { notOKQTY += qty; }
                            }
                            else
                            {
                                showQTY_b = "";
                                showQTY_a = dr["QTY"].ToString();
                                if (bool.Parse(dr["IsOK"].ToString()))
                                { fishQTY -= qty; }
                                else { notOKQTY -= qty; }
                            }
                            if (bool.Parse(dr["IsOK"].ToString())) { state = "入帳"; } else { state = "再途"; }
                            dis_tmp = $"<tr><td>{Convert.ToDateTime(dr["StartTime"]).ToString("yy-MM-dd HH:mm:ss")}</td><td>{dr["DOCNumberNO"].ToString()}</td><td>{dr["DOCName"].ToString()}</td><td colspan='4'>{dr["Remark"].ToString()}</td><td>{state}</td><td>{dr["StoreNO"].ToString()}{dr["StoreSpacesNO"].ToString()}</td><td>{showQTY_a}</td><td>{showQTY_b}</td><td>{notOKQTY.ToString()}</td><td>{fishQTY.ToString()}</td></tr>{dis_tmp}";
                        }
                    }
                }
                re = $"{re}</table>";
            }
            ViewBag.HtmlOutput = re;
            return View(key);
        }


        public IActionResult BOMProductProcess()//BOM與製程總表
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt_BOM = null;
                #region 獨立製程 IsShare_PP_Name='0'
                dt_BOM = db.DB_GetData($@"select a.*,b.PartName,b.Specification,b.Class FROM SoftNetMainDB.[dbo].[BOM] as a 
                                    join SoftNetMainDB.[dbo].[Material] as b on b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO 
                                    where a.ServerId='{_Fun.Config.ServerId}' and a.IsShare_PP_Name='0' order by a.Apply_PP_Name,a.Apply_PartNO,a.EffectiveDate,a.IndexSN desc");
                if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                {
                    DataRow dr_tmp = null;
                    string mbomID_PartNO = "";
                    string pp_Name = "";
                    string apply_PartNO = "";
                    string isQTY = "是";
                    string isOK = "是";
                    string isConfirm = "";
                    string station = "";
                    string stationNO_Merge = "";
                    string color = "";
                    foreach (DataRow dr in dt_BOM.Rows)
                    {
                        if (pp_Name != dr["Apply_PP_Name"].ToString() || apply_PartNO != dr["Apply_PartNO"].ToString())
                        {
                            if (pp_Name != "")
                            {
                                re = $"{re}</table>";
                                if (mbomID_PartNO != dr["PartNO"].ToString()) { re = $"{re}<br />"; }
                            }
                            re = $"{re}<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                            pp_Name = dr["Apply_PP_Name"].ToString();
                            apply_PartNO = dr["Apply_PartNO"].ToString();
                            color = "";
                            if (bool.Parse(dr["IsConfirm"].ToString())) { isConfirm = "發行"; }
                            else { isConfirm = "NG"; color = " style='background-color:tomato;'"; }
                            mbomID_PartNO = dr["PartNO"].ToString();
                            re = $"{re}<tr><td>階</td><td>生產母件料號:{dr["PartNO"].ToString()}&emsp;{dr["PartName"].ToString()}&emsp;{dr["Specification"].ToString()}</td><td>製程名稱:{dr["Apply_PP_Name"].ToString()}</td><td>版本:{dr["Version"].ToString()}</td><td>有效期限:{dr["EffectiveDate"].ToString()}~{dr["ExpiryDate"].ToString()}</td><td{color}>{isConfirm}</td></tr>";
                        }

                        isQTY = bool.Parse(dr["IsChackQTY"].ToString()) ? "是" : "否";
                        isOK = bool.Parse(dr["IsChackIsOK"].ToString()) ? "是" : "否";
                        color = "";
                        station = $"主站:{dr["Apply_StationNO"].ToString()}&emsp;";
                        DataTable dt_PP_ProductProcess_Item = db.DB_GetData($@"select * FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where BOMId='{dr["Id"].ToString()}'");
                        if (dt_PP_ProductProcess_Item != null && dt_PP_ProductProcess_Item.Rows.Count > 0)
                        {
                            stationNO_Merge = "";
                            foreach (DataRow drII in dt_PP_ProductProcess_Item.Rows)
                            {
                                if (drII["StationNO"].ToString() != dr["Apply_StationNO"].ToString())
                                {
                                    if (stationNO_Merge == "") { stationNO_Merge = $"合併站:{drII["StationNO"].ToString()}"; }
                                    else { stationNO_Merge = $"{stationNO_Merge};{drII["StationNO"].ToString()}"; }
                                }
                            }
                            station = $"{station}{stationNO_Merge}";
                        }
                        else
                        {
                            station = $"{station}製程設定異常&emsp;"; color = " style='background-color:tomato;'";
                        }
                        re = $"{re}<tr><td>{dr["IndexSN"].ToString()}</td><td colspan='3'{color}>{station}{dr["Station_Custom_IndexSN"].ToString()}&emsp;{dr["StationNO_Custom_DisplayName"].ToString()}</td><td colspan='2'>關站條件 足量:{isQTY} 單據確認:{isOK}</td></tr>";
                        DataTable dt_BOMII = db.DB_GetData($@"select a.*,b.PartName,b.Specification,b.Class FROM SoftNetMainDB.[dbo].[BOMII] as a
                                                join SoftNetMainDB.[dbo].[Material] as b on b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO 
                                                where BOMId='{dr["Id"].ToString()}' order by sn");
                        if (dt_BOMII != null && dt_BOMII.Rows.Count > 0)
                        {
                            foreach (DataRow drII in dt_BOMII.Rows)
                            {
                                if (drII["PartNO"].ToString() == dr["PartNO"].ToString()) { continue; }
                                re = $"{re}<tr><td></td><td>用料:{drII["PartNO"].ToString()}&emsp;{drII["PartName"].ToString()}&emsp;{drII["Specification"].ToString()}</td><td>用量:{drII["BOMQTY"].ToString()}</td><td>損耗率:{drII["AttritionRate"].ToString()}</td></tr>";
                            }
                        }
                    }
                    re = $"{re}</table>";
                }
                else
                {
                    re = $"{re}<p> 目前無獨立製程BOM紀錄</p>";
                }
                #endregion

                re = $"{re}<br />";

                #region 共用製程 IsShare_PP_Name='1'
                dt_BOM = db.DB_GetData($@"select a.*,b.PartName,b.Specification,b.Class FROM SoftNetMainDB.[dbo].[BOM] as a 
                                    join SoftNetMainDB.[dbo].[Material] as b on b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO 
                                    where a.ServerId='{_Fun.Config.ServerId}' and a.IsShare_PP_Name='1' and Main_Item='1' order by a.Apply_PP_Name,a.Apply_PartNO,a.EffectiveDate,a.IndexSN desc");
                if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                {
                    DataRow dr_tmp = null;
                    string mbomID_PartNO = "";
                    string pp_Name = "";
                    string apply_PartNO = "";
                    string isQTY = "是";
                    string isOK = "是";
                    string isConfirm = "";
                    string station = "";
                    string stationNO_Merge = "";
                    string color = "";
                    int count = 0;
                    foreach (DataRow dr in dt_BOM.Rows)
                    {
                        #region 取得第1階
                        re = $"{re}<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                        pp_Name = dr["Apply_PP_Name"].ToString();
                        apply_PartNO = dr["Apply_PartNO"].ToString();
                        color = "";
                        if (bool.Parse(dr["IsConfirm"].ToString())) { isConfirm = "發行"; }
                        else { isConfirm = "NG"; color = " style='background-color:tomato;'"; }
                        mbomID_PartNO = dr["PartNO"].ToString();
                        re = $"{re}<tr><td>階</td><td>生產母件料號:{dr["PartNO"].ToString()}&emsp;{dr["PartName"].ToString()}&emsp;{dr["Specification"].ToString()}</td><td>製程名稱:{dr["Apply_PP_Name"].ToString()}</td><td>版本:{dr["Version"].ToString()}</td><td>有效期限:{dr["EffectiveDate"].ToString()}~{dr["ExpiryDate"].ToString()}</td><td{color}>{isConfirm}</td></tr>";
                        isQTY = bool.Parse(dr["IsChackQTY"].ToString()) ? "是" : "否";
                        isOK = bool.Parse(dr["IsChackIsOK"].ToString()) ? "是" : "否";
                        color = "";
                        station = $"主站:{dr["Apply_StationNO"].ToString()}&emsp;";
                        DataTable dt_PP_ProductProcess_Item = db.DB_GetData($@"select * FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{dr["Apply_PP_Name"].ToString()}' and PartNO='{dr["PartNO"].ToString()}' and IndexSN={dr["IndexSN"].ToString()}");
                        if (dt_PP_ProductProcess_Item != null && dt_PP_ProductProcess_Item.Rows.Count > 0)
                        {
                            stationNO_Merge = "";
                            foreach (DataRow drII in dt_PP_ProductProcess_Item.Rows)
                            {
                                if (drII["StationNO"].ToString() != dr["Apply_StationNO"].ToString())
                                {
                                    if (stationNO_Merge == "") { stationNO_Merge = $"合併站:{drII["StationNO"].ToString()}"; }
                                    else { stationNO_Merge = $"{stationNO_Merge};{drII["StationNO"].ToString()}"; }
                                }
                            }
                            station = $"{station}{stationNO_Merge}";
                        }
                        else
                        {
                            station = $"{station}製程設定異常&emsp;"; color = " style='background-color:tomato;'";
                        }
                        re = $"{re}<tr><td>{dr["IndexSN"].ToString()}</td><td colspan='3'{color}>{station}{dr["Station_Custom_IndexSN"].ToString()}&emsp;{dr["StationNO_Custom_DisplayName"].ToString()}</td><td colspan='2'>關站條件 足量:{isQTY} 單據確認:{isOK}</td></tr>";
                        DataTable dt_BOMII = db.DB_GetData($@"select a.*,b.PartName,b.Specification,b.Class FROM SoftNetMainDB.[dbo].[BOMII] as a
                                                join SoftNetMainDB.[dbo].[Material] as b on b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO 
                                                where BOMId='{dr["Id"].ToString()}' order by sn");
                        if (dt_BOMII != null && dt_BOMII.Rows.Count > 0)
                        {
                            foreach (DataRow drII in dt_BOMII.Rows)
                            {
                                if (drII["PartNO"].ToString() == dr["PartNO"].ToString()) { continue; }
                                re = $"{re}<tr><td></td><td>用料:{drII["PartNO"].ToString()}&emsp;{drII["PartName"].ToString()}&emsp;{drII["Specification"].ToString()}</td><td>用量:{drII["BOMQTY"].ToString()}</td><td>損耗率:{drII["AttritionRate"].ToString()}</td></tr>";
                            }
                        }
                        #endregion


                        dt_BOMII = db.DB_GetData($"select a.BOMId,a.PartNO,a.BOMQTY,a.Class,b.PartType,b.StoreSTime,b.SafeQTY from SoftNetMainDB.[dbo].[BOMII] as a,dbo.[Material] as b where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO and BOMId='{dr["Id"].ToString()}' order by sn");
                        if (dt_BOMII != null)
                        {
                            foreach (DataRow dr0 in dt_BOMII.Rows)
                            {
                                if (dr0["Class"].ToString() == "4" || dr0["Class"].ToString() == "5")
                                { RecursiveBOMII(db, dr["Apply_PP_Name"].ToString(), dr["Apply_PartNO"].ToString(), dr0, int.Parse(dr["IndexSN"].ToString()), ref re); }
                            }
                        }
                        re = $"{re}</table>";
                    }
                }
                else
                {
                    re = $"{re}<p> 目前無共用製程BOM紀錄</p>";
                }
                #endregion
            }
            ViewBag.HtmlOutput = re;
            return View();
        }
        private void RecursiveBOMII(DBADO db, string Apply_PP_Name, string apply_PartNO, DataRow dr0, int indexSN, ref string re)
        {
            string sql = "";
            DataRow dr = db.DB_GetFirstDataByDataRow($"select a.Id,a.PartNO,a.Apply_PartNO,a.Apply_PP_Name,a.Apply_StationNO,a.IsEnd,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName,a.OutPackType,b.Class,IsChackQTY,IsChackIsOK from SoftNetMainDB.[dbo].[BOM] as a,dbo.[Material] as b where a.Main_Item='0' and a.IndexSN={(indexSN - 1)} and b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{dr0["PartNO"].ToString()}' and a.Apply_PP_Name='{Apply_PP_Name}'  and a.Apply_PartNO='{apply_PartNO}' and a.PartNO=b.PartNO order by IndexSN desc");
            if (dr != null)
            {
                string isQTY = bool.Parse(dr["IsChackQTY"].ToString()) ? "是" : "否";
                string isOK = bool.Parse(dr["IsChackIsOK"].ToString()) ? "是" : "否";
                string source_StationNO = dr["Apply_StationNO"].ToString();
                indexSN -= 1;
                string stationNO_Merge = "";
                string station = $"主站:{dr["Apply_StationNO"].ToString()}&emsp;";
                #region 寫合併站
                DataTable dt_S_Merge = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where IndexSN_Merge='1' and PP_Name='{dr["Apply_PP_Name"].ToString()}' and Apply_PartNO='{apply_PartNO}' and PartNO='{dr["PartNO"].ToString()}' and IndexSN={dr["IndexSN"].ToString()} order by DisplaySN");
                if (dt_S_Merge != null && dt_S_Merge.Rows.Count > 0)
                {
                    foreach (DataRow drII in dt_S_Merge.Rows)
                    {
                        if (drII["StationNO"].ToString() != dr["Apply_StationNO"].ToString())
                        {
                            if (stationNO_Merge == "") { stationNO_Merge = $"合併站:{drII["StationNO"].ToString()}"; }
                            else { stationNO_Merge = $"{stationNO_Merge};{drII["StationNO"].ToString()}"; }
                        }
                    }
                    station = $"{station}{stationNO_Merge}";
                }
                #endregion

                re = $"{re}<tr><td>{dr["IndexSN"].ToString()}</td><td colspan='3'>{station}{dr["Station_Custom_IndexSN"].ToString()}&emsp;{dr["StationNO_Custom_DisplayName"].ToString()}</td><td colspan='2'>關站條件 足量:{isQTY} 單據確認:{isOK}</td></tr>";
                DataTable dt_BOMII = db.DB_GetData($@"select a.*,b.PartName,b.Specification,b.Class FROM SoftNetMainDB.[dbo].[BOMII] as a
                                                join SoftNetMainDB.[dbo].[Material] as b on b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO 
                                                where BOMId='{dr["Id"].ToString()}' order by b.Class,sn");
                if (dt_BOMII != null && dt_BOMII.Rows.Count > 0)
                {
                    foreach (DataRow drII in dt_BOMII.Rows)
                    {
                        if (drII["PartNO"].ToString() == dr["PartNO"].ToString() && (drII["Class"].ToString() == "4" || drII["Class"].ToString() == "5")) { RecursiveBOMII(db, dr["Apply_PP_Name"].ToString(), dr["Apply_PartNO"].ToString(), dr0, indexSN, ref re); continue; }
                        re = $"{re}<tr><td></td><td>用料:{drII["PartNO"].ToString()}&emsp;{drII["PartName"].ToString()}&emsp;{drII["Specification"].ToString()}</td><td>用量:{drII["BOMQTY"].ToString()}</td><td>損耗率:{drII["AttritionRate"].ToString()}</td></tr>";
                        if (drII["Class"].ToString() == "4" || drII["Class"].ToString() == "5")
                        { RecursiveBOMII(db, dr["Apply_PP_Name"].ToString(), dr["Apply_PartNO"].ToString(), dr0, indexSN, ref re); }
                    }
                }
            }
        }


        public IActionResult StationEfficient()
        {
            ViewBag.DisplayHTML = "";
            ViewBag.StationNO = "";
            ViewBag.StationName = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' order by StationNO");
                if (dt != null && dt.Rows.Count > 0)
                {
                    string[] data = DisplayHTMLDIV(db, dt);
                    ViewBag.StationNO = data[0];
                    ViewBag.DisplayHTML = data[1];
                    ViewBag.StationName = data[2];

                }
            }
            return View();
        }

        [HttpPost]
        //public JObject OnActionDisplay() 
        public IActionResult OnActionDisplay()
        {
            return RedirectToAction("StationEfficient", "/Display");
        }


        private string[] DisplayHTMLDIV(DBADO db, DataTable dt)
        {
            string[] data = new string[] { "", "", "" };
            int i = 1;
            string re = "<div id='displayHTMLDIV' class='xp-prog'>";
            string stNO = "";
            string stName = "";
            float value = 0;
            DataRow dr_tmp = null;
            float ct = 0;
            int ect = 0;
            int lct = 0;
            foreach (DataRow dr in dt.Rows)
            {
                value = 0.01f;
                #region 計算值
                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 7 AVG(CycleTime+WaitTime) as CT,AVG(ECT) as ECT,AVG(LowerCT) as LCT,AVG(UpperCT) as UCT FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CycleTime!=0");
                if (dr_tmp != null && !dr_tmp.IsNull("CT") && dr_tmp["CT"].ToString() != "" && int.Parse(dr_tmp["CT"].ToString()) > 0 && int.Parse(dr_tmp["ECT"].ToString()) > 0)
                {
                    ct = int.Parse(dr_tmp["CT"].ToString());
                    if (dr_tmp["LCT"].ToString() != "" && int.Parse(dr_tmp["LCT"].ToString()) >= ct)
                    {
                        value = 100;
                    }
                    else if (dr_tmp["UCT"].ToString() != "" && int.Parse(dr_tmp["UCT"].ToString()) >= ct)
                    {
                        value = 0.01f;
                    }
                    else if (dr_tmp["ECT"].ToString() != "" && dr_tmp["LCT"].ToString() != "" && int.Parse(dr_tmp["ECT"].ToString()) >= ct)
                    {
                        lct = int.Parse(dr_tmp["LCT"].ToString()) <= 0 ? int.Parse(dr_tmp["ECT"].ToString()) : int.Parse(dr_tmp["LCT"].ToString());
                        value = ((ct - lct) / ct) * 100;
                        value = (100 - value) * (60 / 100) + 40;
                    }
                    else if (dr_tmp["ECT"].ToString() != "" && ct > int.Parse(dr_tmp["ECT"].ToString()))
                    {
                        ect = int.Parse(dr_tmp["ECT"].ToString());
                        float _c = ct - ect;
                        float _c01 = _c / ct;
                        float _c02 = _c01 * 100;
                        value = ((ct - ect) / ct) * 100;
                        value = (100 - value) * 0.4f;
                    }
                    else { value = 0.01f; }
                }
                if (value <= 0) { value = 0.01f; }
                if (value > 100) { value = 100; }
                #endregion

                if (stNO == "") { stNO = $"{dr["StationNO"].ToString()},{value.ToString("0.00")}"; }
                else { stNO = $"{stNO};{dr["StationNO"].ToString()},{value.ToString("0.00")}"; }
                if (stName == "") { stName = $"{dr["StationName"].ToString()}"; }
                else { stName = $"{stName};{dr["StationName"].ToString()}"; }

                string sState = "閒置";
                string okQTY = "0";
                string failQTY = "0";
                if (dr["Station_Type"].ToString()=="1")
                {
                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                    if (dr_tmp !=null && (dr_tmp["State"].ToString()=="1" || dr_tmp["State"].ToString() == "2"))
                    {
                        if (dr_tmp["State"].ToString() == "1") { sState = "開工"; }
                        else { sState = "停工"; }
                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetLogDB.[dbo].[SFC_StationDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and PP_Name='{dr_tmp["PP_Name"].ToString()}' and OrderNO='{dr_tmp["OrderNO"].ToString()}' and IndexSN={dr_tmp["IndexSN"].ToString()}");
                        if (dr_tmp != null)
                        {
                            okQTY = dr_tmp["ProductFinishedQty"].ToString();
                            failQTY = dr_tmp["ProductFailedQty"].ToString();
                        }
                    }
                }
                else
                {
                    //多工
                }

                #region display StationNO
                if (i == 1) { re = $"{re}<div class='row'>"; }
                re = $"{re}<div id='{dr["StationNO"].ToString()}' class='gaugeSVGVeiw'></div>";
                if (i >= 4) { re = $"{re}</div>"; i = 1; } else { ++i; }
                #endregion
            }
            re = $"{re}</div>";
            data[0] = stNO;
            data[1] = re;
            data[2] = stName;
            return data;
        }

    }
}
