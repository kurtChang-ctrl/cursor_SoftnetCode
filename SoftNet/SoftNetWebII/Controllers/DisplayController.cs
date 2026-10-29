using Base.Services;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Models;
using System.Data;
using System;
using System.Text;
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
            var sb = new StringBuilder();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string inStation = "";
                if (data.Value != "ALL_Station")
                {
                    foreach (string s in data.Value.Split(';'))
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
                    sb.Append("<div><button id='button_setSTII' type='button' class='btn xg-btn' onclick='_me.onSetII'>寫入修正值</button></div>");
                    sb.Append("<div><table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'><thead><tr>");
                    sb.Append("<th>工站編號</th><th>生產製程名稱</th><th>生產工序</th><th>料號</th><th>品名/規格</th><th>統計筆數</th><th>平均CT</th><th>最佳CT</th><th>最差CT</th><th>有效CT</th><th>修正合理CT</th></tr></thead><tbody>");
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
                                if (dr_tmp != null && (dr_tmp["Station_Custom_IndexSN"].ToString().Trim() != "" || dr_tmp["DisplayName"].ToString().Trim() != ""))
                                { indexDisplay = $"{dr_tmp["Station_Custom_IndexSN"].ToString()}{dr_tmp["DisplayName"].ToString()}"; }
                                TimeSpan standardTime_DIS = new TimeSpan(0, 0, Convert.ToInt32(dr2["AverageCycleTime"].ToString()));
                                avg = $"{(int)standardTime_DIS.TotalMinutes}:{standardTime_DIS.Seconds}";
                                standardTime_DIS = new TimeSpan(0, 0, Convert.ToInt32(dr2["SD_LowerLimit"].ToString()));
                                lct = $"{(int)standardTime_DIS.TotalMinutes}:{standardTime_DIS.Seconds}";
                                standardTime_DIS = new TimeSpan(0, 0, Convert.ToInt32(dr2["SD_UpperLimit"].ToString()));
                                uct = $"{(int)standardTime_DIS.TotalMinutes}:{standardTime_DIS.Seconds}";
                                standardTime_DIS = new TimeSpan(0, 0, Convert.ToInt32(dr2["EfficientCycleTime"].ToString()));
                                ect = $"{(int)standardTime_DIS.TotalMinutes}:{standardTime_DIS.Seconds}";
                                sb.Append($"<tr><td>{dr["StationNO"].ToString()}</td><td>{dr2["PP_Name"].ToString()}</td><td>{indexDisplay}</td><td>{dr2["PartNO"].ToString()}</td><td>{dr2["PartName"].ToString()}&nbsp;{dr2["Specification"].ToString()}</td><td>{dr2["CountQTY"].ToString()}</td><td>{avg}</td><td>{lct}</td><td>{uct}</td><td>{ect}</td>");
                                sb.Append($"<td><input type='number' id='M_{dr2["Id"].ToString()}' name='SetCT_M' value='0' />分<input type='number' id='S_{dr2["Id"].ToString()}' name='SetCT_S' value='0' />秒</td></tr>");
                            }
                        }
                    }
                    sb.Append("</tbody></table></div>");
                }
            }
            return sb.ToString();
        }
        public IActionResult SelectOLDSID(string id)//查詢計畫已轉生產的歷史紀錄
        {
            var sb = new StringBuilder();
            sb.Append("<p>查詢近一年內已轉生產的歷史紀錄</p><hr>");
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt_APS_NeedData = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and State !='0' and State !='1' and State !='2' and State !='3' order by UpdateTime desc");
                if (dt_APS_NeedData != null && dt_APS_NeedData.Rows.Count > 0)
                {
                    DataRow dr_tmp = null;
                    string tmp = "";
                    string state = "";
                    string tmp_date = "";
                    sb.Append("<div class='xp-prog'>");
                    sb.Append(@"<table id='tableRead' class='table table-bordered xg-table' cellspacing='0'> <thead><tr>
                        <th>排程編號</th><th>更新時間</th><th>類型</th><th>類型名稱</th><th>計畫狀態</th><th>計畫入庫日</th><th>料號/品名/規格</th><th>計畫量</th><th>投產量</th><th>生產製程</th><th>BOM編號</th></tr></thead><tbody>");
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
                            case "P": state = "暫停計畫"; break;
                            default: state = "未定義"; break;
                        }
                        if (dr["NeedType"].ToString() == "1") { tmp = "訂單"; }
                        else if (dr["NeedType"].ToString() == "2") { tmp = "客戶"; }
                        else if (dr["NeedType"].ToString() == "5") { tmp = "底稿"; }
                        else { tmp = "廠內"; }
                        tmp_date = "";
                        if (!dr.IsNull("UpdateTime")) { tmp_date = Convert.ToDateTime(dr["UpdateTime"]).ToString("yy/MM/dd HH:mm"); }
                        sb.Append($"<tr><td>{dr["Id"]}</td><td>{tmp_date}</td><td>{tmp}</td>");
                        tmp = "";
                        if (!dr.IsNull("NeedSource") && dr["NeedSource"].ToString() != "") { tmp = dr["NeedSource"].ToString(); }
                        else if (!dr.IsNull("CTName") && dr["CTName"].ToString() != "") { tmp = dr["CTName"].ToString(); } else if (!dr.IsNull("CTNO") && dr["CTNO"].ToString() != "") { tmp = dr["CTNO"].ToString(); }
                        sb.Append($"<td>{tmp}</td><td>{state}</td>");
                        if (dr.IsNull("NeedSimulationDate")) { sb.Append("<td></td>"); }
                        else { sb.Append($"<td>{Convert.ToDateTime(dr["NeedSimulationDate"]).ToString("yy/MM/dd HH:mm")}</td>"); }
                        dr_tmp = db.DB_GetFirstDataByDataRow($"select PartName,Specification from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                        string displayNeedQTY = dr["NeedQTY"].ToString();
                        DataRow sim0ForQty = db.DB_GetFirstDataByDataRow($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr["Id"].ToString()}' and PartSN=0");
                        if (sim0ForQty != null)
                        {
                            displayNeedQTY = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, dr["Id"].ToString(), sim0ForQty["SimulationId"].ToString(), int.Parse(displayNeedQTY)).ToString();
                        }
                        if (dr_tmp != null)
                        { sb.Append($"<td>{dr["PartNO"].ToString()}&nbsp;{dr_tmp["PartName"].ToString()}&nbsp;{dr_tmp["Specification"].ToString()}</td><td>{displayNeedQTY}</td>"); }
                        else
                        { sb.Append($"<td></td><td>{displayNeedQTY}</td>"); }
                        dr_tmp = db.DB_GetFirstDataByDataRow($"select *,(NeedQTY+SafeQTY) as SQTY from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr["Id"].ToString()}' and PartSN=0");
                        if (dr_tmp != null)
                        {
                            int effectiveSqty = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, dr["Id"].ToString(), dr_tmp["SimulationId"].ToString(), int.Parse(dr_tmp["NeedQTY"].ToString())) + int.Parse(dr_tmp["SafeQTY"].ToString());
                            sb.Append($"<td>{effectiveSqty.ToString()}</td>");
                        }
                        else
                        { sb.Append("<td> </td>"); }
                        sb.Append($"<td>{dr["Apply_PP_Name"].ToString()}</td><td>{dr["BOMId"].ToString()}</td></tr>");
                    }
                    sb.Append("</tbody></table></div>");
                }
                else
                {
                    sb.Append("<p> 目前無轉生產的紀錄</p>");
                }

            }
            ViewBag.HtmlOutput = sb.ToString();
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
            var sb = new StringBuilder();
            Dictionary<string, string[]> stationNOList = new Dictionary<string, string[]>();
            DataTable dt = null;
            DataRow tmp = null;
            DataRow tmp2 = null;
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                #region 生產量無法達標
                #endregion

                #region Simulation_ErrorData
                // 效能優化：簡化 SQL，只選擇需要的欄位，並使用 JOIN 取得 PartNO
                dt = db.DB_GetData($@"SELECT a.StationNO, a.ErrorType, a.SimulationId, a.DOCNumberNO, a.LogDate, b.PartNO 
                                    FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                                    LEFT JOIN SoftNetSYSDB.[dbo].[APS_Simulation] as b ON a.SimulationId = b.SimulationId
                                    where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO!='' and a.StationNO!='{_Fun.Config.OutPackStationName}' 
                                    and a.ActionType='' and (a.ErrorType='05' or a.ErrorType='01' or a.ErrorType='02' or a.ErrorType='04' or a.ErrorType='03' or a.ErrorType='15') 
                                    order by a.LogDate desc");
                if (dt != null && dt.Rows.Count > 0)
                {
                    string log = "";
                    foreach (DataRow dr in dt.Rows)
                    {
                        //###??? 其他錯誤不期
                        log = "";
                        string errorType = dr.IsNull("ErrorType") ? "" : dr["ErrorType"].ToString();
                        switch (errorType)
                        {
                            case "05":
                                log = $"生產備料無作為且無領料單據發起";
                                break;
                            case "01":
                            case "02":
                                // 效能優化：PartNO 已經在 JOIN 中取得，不需要額外查詢
                                string tmp_S = dr.IsNull("PartNO") ? "" : dr["PartNO"].ToString();
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
                            case "14":
                                log = $"實際CT遠離目標CT";
                                break;
                            case "15":
                                log = $"工單:{dr["DOCNumberNO"].ToString()}&nbsp;報工作業可能有延遲";
                                break;
                            case "17":
                                log = $"數量移轉量不合理";
                                break;
                            case "18":
                                log = $"應停工未停工";
                                break;
                            case "19":
                                log = $"計畫已偏離需求日";
                                break;
                            case "21":
                                log = $"強制檢查各站報工的合理性";
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
                                log = "未知的錯誤";
                                break;
                        }
                        string stationNO = dr.IsNull("StationNO") ? "" : dr["StationNO"].ToString();
                        if (!string.IsNullOrEmpty(stationNO) && !stationNOList.ContainsKey(stationNO))
                        {
                            stationNOList.Add(stationNO, new string[] { "E", log, dr.IsNull("LogDate") ? "" : dr["LogDate"].ToString() });
                        }
                    }
                }
                #endregion

                #region 開工的工站
                // 效能優化：簡化 SQL，只選擇需要的欄位
                dt = db.DB_GetData($"SELECT StationNO FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' and State='1'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string stationNO = dr.IsNull("StationNO") ? "" : dr["StationNO"].ToString();
                        if (!string.IsNullOrEmpty(stationNO) && !stationNOList.ContainsKey(stationNO))
                        {
                            stationNOList.Add(stationNO, new string[] { "S", "", "" });
                        }
                    }
                }
                // 效能優化：簡化 SQL，只選擇需要的欄位
                dt = db.DB_GetData($"SELECT StationNO FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' and RemarkTimeS is not NULL and RemarkTimeE is NULL group by StationNO");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string stationNO = dr.IsNull("StationNO") ? "" : dr["StationNO"].ToString();
                        if (!string.IsNullOrEmpty(stationNO) && !stationNOList.ContainsKey(stationNO))
                        {
                            stationNOList.Add(stationNO, new string[] { "S", "", "" });
                        }
                    }
                }
                #endregion

                #region 剩下的工站
                // 效能優化：簡化 SQL，只選擇需要的欄位
                dt = db.DB_GetData($"SELECT StationNO FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string stationNO = dr.IsNull("StationNO") ? "" : dr["StationNO"].ToString();
                        if (!string.IsNullOrEmpty(stationNO) && !stationNOList.ContainsKey(stationNO))
                        {
                            stationNOList.Add(stationNO, new string[] { "", "", "" });
                        }
                    }
                }
                #endregion

                // 效能優化：批次載入所有需要的 PP_Station 資料
                Dictionary<string, DataRow> ppStationCache = new Dictionary<string, DataRow>();
                var stationNOsForPP = stationNOList.Keys.ToList();
                if (stationNOsForPP.Count > 0)
                {
                    string stationNOsStr = string.Join("','", stationNOsForPP);
                    DataTable dt_PP_Stations = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO IN ('{stationNOsStr}')");
                    if (dt_PP_Stations != null)
                    {
                        foreach (DataRow row in dt_PP_Stations.Rows)
                        {
                            if (!row.IsNull("StationNO"))
                            {
                                ppStationCache[row["StationNO"].ToString()] = row;
                            }
                        }
                    }
                }

                // 效能優化：預先計算今天的日期字串，避免在迴圈中重複計算
                DateTime now = DateTime.Now;
                string todayStr = now.ToString("yyyy-MM-dd");
                string todayStrForSQL = now.ToString("yyyy-MM-dd");

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
                    if (obj.Key != beforKey)
                    {
                        errLOG = "";
                        errTime = "";
                        beforKey = obj.Key;
                    }

                    // 使用快取的 PP_Station 資料
                    if (!ppStationCache.TryGetValue(obj.Key, out dr_PP_Station))
                    {
                        dr_PP_Station = null;
                    }
                    if (obj.Value[0] == "E")
                    {
                        sb.Append("<table style='background-color:tomato;border:1px solid black;width:100%' rules='all'>");
                        if (errLOG != "") { errLOG = $"<br />{obj.Value[1]}"; }
                        else { errLOG = obj.Value[1]; }
                        if (errTime != "") { errTime = $"<br />{obj.Value[2]}"; }
                        else { errTime = obj.Value[2]; }
                    }
                    else if (obj.Value[0] == "W") { sb.Append("<table style='background-color:gold;border:1px solid black;width=100%' rules='all'>"); }
                    else if (obj.Value[0] == "S") { sb.Append("<table style='background-color:springgreen;border:1px solid black;width=100%' rules='all'>"); }
                    else
                    {
                        sb.Append("<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>");
                    }
                    string opNO = "";
                    string partNO = "";
                    string partName = "";
                    string muti = "";
                    if (dr_PP_Station["Station_Type"].ToString() == "1")
                    {
                        dr_tmp = db.DB_GetFirstDataByDataRow($"select State from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{obj.Key}'");
                        if (dr_tmp != null && dr_tmp["State"].ToString() == "1")
                        { sb.Append($"<tr><td>單工工站:{obj.Key}&nbsp;&nbsp;{dr_PP_Station["StationName"].ToString()}&nbsp;&nbsp;狀態:開工中</td><td>生產線:{dr_PP_Station["LineName"].ToString()}</td>"); }
                        else
                        { sb.Append($"<tr><td>單工工站:{obj.Key}&nbsp;&nbsp;{dr_PP_Station["StationName"].ToString()}</td><td>生產線:{dr_PP_Station["LineName"].ToString()}</td>"); }
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
                            {
                                string woQty = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, APS_PartNOTimeNote["NeedId"].ToString(), tmp["SimulationId"].ToString(), int.Parse(APS_PartNOTimeNote["NeedQTY"].ToString())).ToString();
                                sb.Append($"<td>工單:{tmp["OrderNO"].ToString()}</td><td>工單量:{woQty}</td><td>已報工量:{APS_PartNOTimeNote["Detail_QTY"].ToString()}</td></tr>");
                            }
                            else { sb.Append($"<td>工單:{tmp["OrderNO"].ToString()}</td><td>工單量:0</td><td>已報工量:0</td></tr>"); }
                        }
                        else { sb.Append("<td>工單:</td><td>工單量:0</td><td>已報工量:0</td></tr>"); }
                    }
                    else
                    {
                        sb.Append($"<tr><td>多工工站:{obj.Key}&nbsp;&nbsp;{dr_PP_Station["StationName"].ToString()}</td><td>生產線:{dr_PP_Station["LineName"].ToString()}</td>");
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
                                    string rowWoQty = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, APS_PartNOTimeNote["NeedId"].ToString(), dr["SimulationId"].ToString(), int.Parse(APS_PartNOTimeNote["NeedQTY"].ToString())).ToString();
                                    if (qty == "") { qty = rowWoQty; } else { qty = $"{qty}<br />{rowWoQty}"; }
                                    if (okqty == "") { okqty = APS_PartNOTimeNote["Detail_QTY"].ToString(); } else { okqty = $"{okqty}<br />{APS_PartNOTimeNote["Detail_QTY"].ToString()}"; }
                                }
                            }
                            sb.Append($"<td>工單:{muti}{wo}&nbsp;&nbsp;狀態:開工中</td><td>工單量:{muti}{qty}</td><td>已報工量:{muti}{okqty}</td></tr>");
                        }
                        else { sb.Append("<td>工單:</td><td>工單量:0</td><td>已報工量:0</td></tr>"); }
                    }
                    sb.Append($"<tr><td>生產料號:{muti}{partNO}</td><td colspan='4'>{muti}{partName}</td></tr>");

                    string b01 = "";
                    string b02 = "";
                    string c01 = "0:0";
                    string c02 = "0:0";
                    if (is_TOT)
                    {
                        #region 效能,工時 計算 c01
                        // 效能優化：使用預先計算的日期字串
                        APS_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time1_C+Time2_C+Time3_C+Time4_C) as C01 from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{obj.Key}' and CONVERT(varchar(10), CalendarDate, 111)='{todayStrForSQL}'");
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
                        // 效能優化：使用預先計算的日期字串
                        string calendarName = dr_PP_Station != null && !dr_PP_Station.IsNull("CalendarName") ? dr_PP_Station["CalendarName"].ToString() : "";
                        DataRow dr_StandardTime = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday='{todayStrForSQL}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{calendarName}' order by CalendarName,Holiday desc");
                        if (dr_StandardTime != null)
                        {
                            // 效能優化：簡化 SQL，只選擇需要的欄位
                            tmp_dt = db.DB_GetData($"select LOGDateTime, OperateType from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{obj.Key}' and CONVERT(varchar(10), LOGDateTime, 111)='{todayStrForSQL}' and (OperateType like '%開工%' or OperateType like '%停工%' or OperateType like '%關站%') order by LOGDateTime");
                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                            {
                                char run = '0';
                                foreach (DataRow dr3 in tmp_dt.Rows)
                                {
                                    if (dr3["OperateType"].ToString().Contains("開工"))
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
                                            if (dr3["OperateType"].ToString().Contains("停工") || dr3["OperateType"].ToString().Contains("關站"))
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
                    sb.Append($"<tr><td>歷史效益差:{b01}</td><td>在線效益差:{b02}</td><td>本日最大負荷:{c01}分</td><td>本日累計負荷:{c02}分</td><td>作業人員:{opNO}</td></tr>");

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
                                sb.Append($"<tr><td>目前Cycle:{t01}</td><td>有效Cycle:{t02}</td><td>最佳Cycle:{t03}</td><td>目標Cycle:{t04}</td><td>達成率:{t05}</td></tr>");
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
                                sb.Append($"<tr><td>{(++i).ToString()}.目前Cycle:{t01}</td><td>有效Cycle:{t02}</td><td>最佳Cycle:{t03}</td><td>目標Cycle:{t04}</td><td>達成率:{t05}</td></tr>");
                            }
                        }
                    }

                    if (errTime != "")
                    { sb.Append($"<tr><td>最新警示干涉:{errTime}</td><td colspan='4'>{errLOG}</td></tr>"); }



                    sb.Append("</table><br />");
                }
            }
            return sb.ToString();
        }
        public string ChangeDisplayDIV() //採購確認  0=ipport,1.需求碼1,需求碼2....
        {
            //###???暫時寫死單別與廠商 AA02,單價
            string meg = SelectStationStateII();
            return meg;
        }
        public IActionResult ALLStationDetail()
        {
            var sb = new StringBuilder();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                // 效能優化：預先計算時間，避免在 SQL 中重複計算
                DateTime now = DateTime.Now;
                DateTime startDate = now.AddDays(-14);
                string startDateStr = startDate.ToString("yyyy-MM-dd HH:mm:ss");
                string endDateStr = now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                sb.Append("最近14天報工紀錄");
                sb.Append("<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>");
                sb.Append("<tr><td>日期</td><td>時間</td><td>工站</td><td>料號/品名/規格</td><td>工序</td><td>OK數量</td><td>NG數量</td><td>OK前期</td><td>NG前期</td><td>作業人員</td><td>實際CT</td><td>目標CT</td></tr>");

                // 效能優化：一次載入所有需要的資料
                DataTable dt = db.DB_GetData($@"SELECT b.PartNO,b.Source_StationNO_IndexSN,b.Source_StationNO_Custom_IndexSN,b.Source_StationNO_Custom_DisplayName,a.* FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] as a
                                            join SoftNetSYSDB.[dbo].[APS_Simulation] as b on b.SimulationId=a.SimulationId and b.ServerId='{_Fun.Config.ServerId}'
                                            join SoftNetMainDB.[dbo].[Material] as c on c.PartNO=b.PartNO and c.ServerId='{_Fun.Config.ServerId}'
                                            where a.ServerId='{_Fun.Config.ServerId}' and a.LOGDateTimeID>='{startDateStr}' and
                                            a.LOGDateTimeID<='{endDateStr}' order by a.LOGDateTimeID desc,a.StationNO,a.OP_NO,b.PartNO");

                if (dt == null || dt.Rows.Count == 0)
                {
                    sb.Append("</table>");
                    ViewBag.HtmlOutput = sb.ToString();
                    return View();
                }

                // 效能優化：批次載入所有需要的工站、料號、使用者資料，避免在迴圈中重複查詢
                var stationNOs = new HashSet<string>();
                var partNOs = new HashSet<string>();
                var userNOs = new HashSet<string>();

                foreach (DataRow dr in dt.Rows)
                {
                    if (!dr.IsNull("StationNO")) { stationNOs.Add(dr["StationNO"].ToString()); }
                    if (!dr.IsNull("PartNO")) { partNOs.Add(dr["PartNO"].ToString()); }
                    if (!dr.IsNull("OP_NO") && dr["OP_NO"].ToString() != "" && !dr["OP_NO"].ToString().Contains(';'))
                    {
                        userNOs.Add(dr["OP_NO"].ToString());
                    }
                }

                // 批次載入工站資料
                Dictionary<string, string> stationCache = new Dictionary<string, string>();
                if (stationNOs.Count > 0)
                {
                    string stationNOsStr = string.Join("','", stationNOs);
                    DataTable dt_Stations = db.DB_GetData($"SELECT StationNO,StationName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO IN ('{stationNOsStr}')");
                    if (dt_Stations != null)
                    {
                        foreach (DataRow row in dt_Stations.Rows)
                        {
                            if (!row.IsNull("StationNO"))
                            {
                                string stationNO = row["StationNO"].ToString();
                                string stationName = row.IsNull("StationName") ? "" : row["StationName"].ToString();
                                stationCache[stationNO] = $"{stationNO}&nbsp;{stationName}";
                            }
                        }
                    }
                }

                // 批次載入料號資料
                Dictionary<string, string> materialCache = new Dictionary<string, string>();
                if (partNOs.Count > 0)
                {
                    string partNOsStr = string.Join("','", partNOs);
                    DataTable dt_Materials = db.DB_GetData($"SELECT PartNO,PartName,Specification from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO IN ('{partNOsStr}')");
                    if (dt_Materials != null)
                    {
                        foreach (DataRow row in dt_Materials.Rows)
                        {
                            if (!row.IsNull("PartNO"))
                            {
                                string partNO = row["PartNO"].ToString();
                                string partName = row.IsNull("PartName") ? "" : row["PartName"].ToString();
                                string specification = row.IsNull("Specification") ? "" : row["Specification"].ToString();
                                materialCache[partNO] = $"{partNO}&nbsp;{partName}&nbsp;{specification}";
                            }
                        }
                    }
                }

                // 批次載入使用者資料
                Dictionary<string, string> userCache = new Dictionary<string, string>();
                if (userNOs.Count > 0)
                {
                    string userNOsStr = string.Join("','", userNOs);
                    DataTable dt_Users = db.DB_GetData($"SELECT UserNO,Name from SoftNetMainDB.[dbo].[User] where ServerId='{_Fun.Config.ServerId}' and UserNO IN ('{userNOsStr}')");
                    if (dt_Users != null)
                    {
                        foreach (DataRow row in dt_Users.Rows)
                        {
                            if (!row.IsNull("UserNO"))
                            {
                                string userNO = row["UserNO"].ToString();
                                string userName = row.IsNull("Name") ? "" : row["Name"].ToString();
                                userCache[userNO] = userName;
                            }
                        }
                    }
                }

                // 處理資料列
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

                foreach (DataRow dr in dt.Rows)
                {
                    // 效能優化：預先轉換日期，避免重複轉換
                    DateTime logDateTime = Convert.ToDateTime(dr["LOGDateTimeID"]);
                    string currentDate = logDateTime.ToString("yyyy-MM-dd");

                    if (rdate != currentDate)
                    {
                        rdate = currentDate;
                        rdateII = rdate;
                    }
                    else { rdateII = ""; }
                    rTime = logDateTime.ToString("HH:mm");

                    // 使用快取的工站資料
                    string stationNO = dr.IsNull("StationNO") ? "" : dr["StationNO"].ToString();
                    if (!string.IsNullOrEmpty(stationNO) && stationCache.TryGetValue(stationNO, out rStationNO)) { }
                    else { rStationNO = ""; }

                    // 使用快取的料號資料
                    string partNO = dr.IsNull("PartNO") ? "" : dr["PartNO"].ToString();
                    if (!string.IsNullOrEmpty(partNO) && materialCache.TryGetValue(partNO, out rPartNO)) { }
                    else { rPartNO = ""; }

                    // 使用快取的使用者資料
                    if (!dr.IsNull("OP_NO") && dr["OP_NO"].ToString() != "")
                    {
                        string opNo = dr["OP_NO"].ToString();
                        if (opNo.Contains(';')) { rOP_NO = "多人報工"; }
                        else
                        {
                            if (userCache.TryGetValue(opNo, out rOP_NO)) { }
                            else { rOP_NO = ""; }
                        }
                    }
                    else { rOP_NO = ""; }

                    // 工序顯示名稱
                    if (!dr.IsNull("Source_StationNO_Custom_DisplayName") && dr["Source_StationNO_Custom_DisplayName"].ToString() != "")
                    { indexNO = dr["Source_StationNO_Custom_DisplayName"].ToString(); }
                    else if (!dr.IsNull("Source_StationNO_Custom_IndexSN") && dr["Source_StationNO_Custom_IndexSN"].ToString() != "")
                    { indexNO = dr["Source_StationNO_Custom_IndexSN"].ToString(); }
                    else { indexNO = dr.IsNull("Source_StationNO_IndexSN") ? "" : dr["Source_StationNO_IndexSN"].ToString(); }

                    // 效能優化：預先轉換數值，避免重複轉換
                    ct = dr.IsNull("CycleTime") ? 0 : int.Parse(dr["CycleTime"].ToString());
                    TimeSpan standardTime_DIS = new TimeSpan(0, 0, ct);
                    CT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                    ect = dr.IsNull("ECT") ? 0 : int.Parse(dr["ECT"].ToString());
                    standardTime_DIS = new TimeSpan(0, 0, ect);
                    ECT = $"{(int)standardTime_DIS.TotalMinutes}分{standardTime_DIS.Seconds}秒";
                    if (ct > ect && ect != 0) { color = " style='background-color:gold;'"; } else { color = ""; }

                    sb.Append($"<tr><td>{rdateII}</td><td>{rTime}</td><td>{rStationNO}</td><td>{rPartNO}</td><td>{indexNO}</td><td>{dr["EditFinishedQty"]}</td><td>{dr["EditFailedQty"]}</td>");
                    sb.Append($"<td>{dr["OLD_ProductFinishedQty"]}</td><td>{dr["OLD_ProductFailedQty"]}</td><td>{rOP_NO}</td><td{color}>{CT}</td><td>{ECT}</td></tr>");
                }
                sb.Append("</table>");
            }
            ViewBag.HtmlOutput = sb.ToString();
            return View();
        }
        private List<string[]> LoadOperateLogStationList(DBADO db)
        {
            var list = new List<string[]>();
            DataTable dt = db.DB_GetData($"SELECT StationNO,StationName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' order by StationNO");
            if (dt != null)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (!dr.IsNull("StationNO"))
                    {
                        list.Add(new string[] { dr["StationNO"].ToString(), dr.IsNull("StationName") ? "" : dr["StationName"].ToString() });
                    }
                }
            }
            return list;
        }

        private List<string[]> LoadOperateLogUserList(DBADO db)
        {
            var list = new List<string[]>();
            DataTable dt = db.DB_GetData($"SELECT UserNO,Name from SoftNetMainDB.[dbo].[User] where ServerId='{_Fun.Config.ServerId}' order by UserNO");
            if (dt != null)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (!dr.IsNull("UserNO"))
                    {
                        list.Add(new string[] { dr["UserNO"].ToString(), dr.IsNull("Name") ? "" : dr["Name"].ToString() });
                    }
                }
            }
            return list;
        }

        private string BuildOperateLogWhereSql(APSViewData key)
        {
            string wheresql = "";
            if (key.SelectFun3 != "ALL" && !string.IsNullOrEmpty(key.SelectFun3))
            {
                var stationList = new List<string>();
                foreach (string s in key.SelectFun3.Split(';'))
                {
                    if (!string.IsNullOrEmpty(s)) { stationList.Add(s); }
                }
                if (stationList.Count > 0) { wheresql += $" and a.StationNO IN ('{string.Join("','", stationList)}')"; }
            }
            if (key.SelectFun4 != "ALL" && !string.IsNullOrEmpty(key.SelectFun4))
            {
                var userList = new List<string>();
                foreach (string s in key.SelectFun4.Split(';'))
                {
                    if (!string.IsNullOrEmpty(s)) { userList.Add(s); }
                }
                if (userList.Count > 0) { wheresql += $" and a.OP_NO IN ('{string.Join("','", userList)}')"; }
            }
            if (!string.IsNullOrEmpty(key.MBOMId)) { wheresql += $" and a.NeedId like '%{key.MBOMId}%'"; }
            if (!string.IsNullOrEmpty(key.SelectNeedId)) { wheresql += $" and a.OrderNO like '%{key.SelectNeedId}%'"; }
            if (key.SelectFun5 != "ALL" && !string.IsNullOrEmpty(key.SelectFun5))
            {
                var operateList = new List<string>();
                foreach (string s in key.SelectFun5.Split(';'))
                {
                    if (!string.IsNullOrEmpty(s)) { operateList.Add(s); }
                }
                if (operateList.Count > 0) { wheresql += $" and a.OperateType IN (N'{string.Join("',N'", operateList)}')"; }
            }
            return wheresql;
        }

        private string BuildALLOperateLogHtml(DBADO db, APSViewData key)
        {
            var sb = new StringBuilder();
            string wheresql = BuildOperateLogWhereSql(key);

            // 有輸入 計畫編號 或 工單編號 時，不以 LOGDateTime 作為查詢條件（全部日期都在範圍內）
            bool hasKeyNO = !string.IsNullOrEmpty(key.MBOMId) || !string.IsNullOrEmpty(key.SelectNeedId);
            string dateWhereSql = "";
            if (!hasKeyNO)
            {
                string startDateStr = key.SelectFun1 + " 00:00:00.000";
                string endDateStr = key.SelectFun2 + " 23:59:59.999";
                dateWhereSql = $" and a.LOGDateTime>='{startDateStr}' and a.LOGDateTime<='{endDateStr}'";
                sb.Append($"{key.SelectFun1} ~ {key.SelectFun2} 操作紀錄");
            }
            else
            {
                sb.Append("操作紀錄(依 計畫編號/工單編號 查詢, 不限日期)");
            }
            sb.Append("<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>");
            sb.Append("<tr><td>日期</td><td>排程編號</td><td>操作</td><td>工站</td><td>工單編號</td><td>料號/品名/規格</td><td>操作人員</td><td>備註</td></tr>");

            DataTable dt = db.DB_GetData($@"SELECT b.PartName,b.Specification,a.* FROM SoftNetLogDB.[dbo].[OperateLog] as a
                                            join SoftNetMainDB.[dbo].[Material] as b on b.PartNO=a.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                                            where a.ServerId='{_Fun.Config.ServerId}'{dateWhereSql} {wheresql} order by a.LOGDateTime desc,a.StationNO,a.OP_NO,a.PartNO");

            if (dt == null || dt.Rows.Count == 0)
            {
                sb.Append("</table>");
                return sb.ToString();
            }

            var stationNOs = new HashSet<string>();
            var partNOs = new HashSet<string>();
            var userNOs = new HashSet<string>();

            foreach (DataRow dr in dt.Rows)
            {
                if (!dr.IsNull("StationNO")) { stationNOs.Add(dr["StationNO"].ToString()); }
                if (!dr.IsNull("PartNO")) { partNOs.Add(dr["PartNO"].ToString()); }
                if (!dr.IsNull("OP_NO") && dr["OP_NO"].ToString() != "" && !dr["OP_NO"].ToString().Contains(';'))
                {
                    userNOs.Add(dr["OP_NO"].ToString());
                }
            }

            Dictionary<string, string> stationCache = new Dictionary<string, string>();
            if (stationNOs.Count > 0)
            {
                string stationNOsStr = string.Join("','", stationNOs);
                DataTable dt_Stations = db.DB_GetData($"SELECT StationNO,StationName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO IN ('{stationNOsStr}')");
                if (dt_Stations != null)
                {
                    foreach (DataRow row in dt_Stations.Rows)
                    {
                        if (!row.IsNull("StationNO"))
                        {
                            string stationNO = row["StationNO"].ToString();
                            string stationName = row.IsNull("StationName") ? "" : row["StationName"].ToString();
                            stationCache[stationNO] = $"{stationNO}&nbsp;{stationName}";
                        }
                    }
                }
            }

            Dictionary<string, string> materialCache = new Dictionary<string, string>();
            if (partNOs.Count > 0)
            {
                string partNOsStr = string.Join("','", partNOs);
                DataTable dt_Materials = db.DB_GetData($"SELECT PartNO,PartName,Specification from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO IN ('{partNOsStr}')");
                if (dt_Materials != null)
                {
                    foreach (DataRow row in dt_Materials.Rows)
                    {
                        if (!row.IsNull("PartNO"))
                        {
                            string partNO = row["PartNO"].ToString();
                            string partName = row.IsNull("PartName") ? "" : row["PartName"].ToString();
                            string specification = row.IsNull("Specification") ? "" : row["Specification"].ToString();
                            materialCache[partNO] = $"{partNO}&nbsp;{partName}&nbsp;{specification}";
                        }
                    }
                }
            }

            Dictionary<string, string> userCache = new Dictionary<string, string>();
            if (userNOs.Count > 0)
            {
                string userNOsStr = string.Join("','", userNOs);
                DataTable dt_Users = db.DB_GetData($"SELECT UserNO,Name from SoftNetMainDB.[dbo].[User] where ServerId='{_Fun.Config.ServerId}' and UserNO IN ('{userNOsStr}')");
                if (dt_Users != null)
                {
                    foreach (DataRow row in dt_Users.Rows)
                    {
                        if (!row.IsNull("UserNO"))
                        {
                            string userNO = row["UserNO"].ToString();
                            string userName = row.IsNull("Name") ? "" : row["Name"].ToString();
                            userCache[userNO] = userName;
                        }
                    }
                }
            }

            string rStationNO = "";
            string rPartNO = "";
            string rOP_NO = "";

            foreach (DataRow dr in dt.Rows)
            {
                string stationNO = dr.IsNull("StationNO") ? "" : dr["StationNO"].ToString();
                if (!string.IsNullOrEmpty(stationNO) && stationCache.TryGetValue(stationNO, out rStationNO)) { }
                else { rStationNO = ""; }

                string partNO = dr.IsNull("PartNO") ? "" : dr["PartNO"].ToString();
                if (!string.IsNullOrEmpty(partNO))
                {
                    if (materialCache.TryGetValue(partNO, out rPartNO)) { }
                    else
                    {
                        string partName = dr.IsNull("PartName") ? "" : dr["PartName"].ToString();
                        string specification = dr.IsNull("Specification") ? "" : dr["Specification"].ToString();
                        rPartNO = $"{partNO}&nbsp;{partName}&nbsp;{specification}";
                    }
                }
                else { rPartNO = ""; }

                if (!dr.IsNull("OP_NO") && dr["OP_NO"].ToString() != "")
                {
                    string opNo = dr["OP_NO"].ToString();
                    if (opNo.Contains(';')) { rOP_NO = "多人報工"; }
                    else
                    {
                        if (userCache.TryGetValue(opNo, out rOP_NO)) { }
                        else { rOP_NO = ""; }
                    }
                }
                else { rOP_NO = ""; }

                string logDateTime = dr.IsNull("LOGDateTime") ? "" : Convert.ToDateTime(dr["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss");
                string needId = dr.IsNull("NeedId") ? "" : dr["NeedId"].ToString();
                string operateType = dr.IsNull("OperateType") ? "" : dr["OperateType"].ToString();
                string orderNO = dr.IsNull("OrderNO") ? "" : dr["OrderNO"].ToString();
                string remark = dr.IsNull("Remark") ? "" : dr["Remark"].ToString();

                sb.Append($"<tr><td>{logDateTime}</td><td>{needId}</td><td>{operateType}</td><td>{rStationNO}</td><td>{orderNO}</td><td>{rPartNO}</td><td>{rOP_NO}</td><td>{remark}</td></tr>");
            }
            sb.Append("</table>");
            return sb.ToString();
        }

        public IActionResult ALLOperateLog(APSViewData key)
        {
            DateTime now = DateTime.Now;
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                if (key == null || string.IsNullOrEmpty(key.SelectFun1))
                {
                    key = new APSViewData();
                    key.SelectFun1 = now.AddDays(-30).ToString("yyyy-MM-dd");
                    key.SelectFun2 = now.ToString("yyyy-MM-dd");
                    key.SelectFun3 = "ALL";
                    key.SelectFun4 = "ALL";
                    key.SelectFun5 = "ALL";
                    key.MBOMId = "";
                    key.SelectNeedId = "";
                }
                key.SelectFun6 = now.ToString("yyyy-MM-dd");
                key.StationList = LoadOperateLogStationList(db);
                key.UserList = LoadOperateLogUserList(db);

                if (string.IsNullOrEmpty(key.MBOMId) && string.IsNullOrEmpty(key.SelectNeedId)
                    && string.Compare(key.SelectFun1, key.SelectFun2, StringComparison.Ordinal) > 0)
                {
                    key.ERRMsg = "起始日期不可大於終止日期";
                    ViewBag.HtmlOutput = "";
                    return View(key);
                }

                ViewBag.HtmlOutput = BuildALLOperateLogHtml(db, key);
            }
            return View(key);
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

                // 效能優化：預先計算時間，避免在 SQL 中重複計算
                int betime = -1;
                int.TryParse(key.SelectFun2, out betime);
                DateTime now = DateTime.Now;
                DateTime startDate = now.AddDays(betime);
                string startDateStr = startDate.ToString("yyyy-MM-dd HH:mm:ss.fff");

                // 效能優化：預先建立 WHERE 條件字串
                string wheresql = "";
                if (key.SelectFun3 == "1") { wheresql = $" and (a.ActionType like '%刪除計畫%' or a.ActionType like '%開工%' or a.ActionType like '%設定工單%' or a.ActionType like '%停工%' or a.ActionType like '%關站%' or a.ActionType like '%干涉派工%' or a.ActionType like '%初始化%')"; }
                else if (key.SelectFun3 == "2") { wheresql = $" and a.ActionType like '%倉庫%'"; }
                else if (key.SelectFun3 == "3") { wheresql = $" and a.ActionType like '%儲位%'"; }
                else if (key.SelectFun3 == "4") { wheresql = $" and (a.ActionType like '%Fail%' or a.ActionType like '%失敗%')"; }
                else { wheresql = ""; }

                var sb = new StringBuilder();
                sb.Append("<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>");
                sb.Append("<tr><td>日期</td><td>電子編號</td><td>編號型態</td><td>動作</td><td>接收</td><td>回饋訊號</td><td>內容</td></tr>");

                // 效能優化：使用預先計算的時間字串
                DataTable dt = db.DB_GetData($@"SELECT a.*,b.Type,b.StationNO,b.StoreNO,b.StoreSpacesNO FROM [SoftNetLogDB].[dbo].[LabelStateLog] as a
                                                join SoftNetMainDB.[dbo].[LabelStateINFO] as b on a.macID=b.macID
                                                where a.LOGDateTime>='{startDateStr}' {wheresql} order by a.LOGDateTime desc");

                if (dt != null && dt.Rows.Count > 0)
                {
                    string disType = "";
                    bool showJSON = (key.SelectFun4 != "1");

                    foreach (DataRow dr in dt.Rows)
                    {
                        // 效能優化：預先轉換 Type，避免重複轉換
                        string type = dr.IsNull("Type") ? "" : dr["Type"].ToString();

                        if (type == "1" || type == "4")
                        {
                            disType = dr.IsNull("StationNO") ? "" : dr["StationNO"].ToString();
                        }
                        else if (type == "2")
                        {
                            string storeNO = dr.IsNull("StoreNO") ? "" : dr["StoreNO"].ToString();
                            disType = $"倉庫:{storeNO}";
                        }
                        else if (type == "3")
                        {
                            string storeNO = dr.IsNull("StoreNO") ? "" : dr["StoreNO"].ToString();
                            string storeSpacesNO = dr.IsNull("StoreSpacesNO") ? "" : dr["StoreSpacesNO"].ToString();
                            disType = $"倉庫:{storeNO} 儲位:{storeSpacesNO}";
                        }
                        else
                        {
                            disType = "";
                        }

                        // 效能優化：預先轉換所有欄位，避免重複轉換
                        string logDateTime = dr.IsNull("LOGDateTime") ? "" : Convert.ToDateTime(dr["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss");
                        string macID = dr.IsNull("macID") ? "" : dr["macID"].ToString();
                        string actionType = dr.IsNull("ActionType") ? "" : dr["ActionType"].ToString();
                        string receiveType = dr.IsNull("ReceiveType") ? "" : dr["ReceiveType"].ToString();
                        string info = dr.IsNull("INFO") ? "" : dr["INFO"].ToString();
                        string json = showJSON ? (dr.IsNull("JSON") ? "" : dr["JSON"].ToString()) : "";

                        if (showJSON)
                        {
                            sb.Append($"<tr><td>{logDateTime}</td><td>{macID}</td><td>{disType}</td><td>{actionType}</td><td>{receiveType}</td><td>{info}</td><td>{json}</td></tr>");
                        }
                        else
                        {
                            sb.Append($"<tr><td>{logDateTime}</td><td>{macID}</td><td>{disType}</td><td>{actionType}</td><td>{receiveType}</td><td>{info}</td><td></td></tr>");
                        }
                    }
                }
                sb.Append("</table>");
                re = sb.ToString();
            }
            ViewBag.HtmlOutput = re;
            return View();
        }
        public IActionResult ALLMFData()
        {
            var sb = new StringBuilder();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                sb.Append("<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>");
                sb.Append("<tr><td>廠商編號</td><td>廠商名稱</td><td>統一編號</td><td>聯絡人</td><td>聯絡TEL</td><td>單據收發EMail</td><td>入庫倉別</td><td>排程權重</td><td>地址</td><td>備註</td></tr>");

                DataTable dt = db.DB_GetData($"select * FROM SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' order by MFNO,MFName,TEL,ContactTEL,CTDataWeights desc");

                // 效能優化：批次載入所有需要的 Store 資料，避免 N+1 查詢問題
                Dictionary<string, string> storeCache = new Dictionary<string, string>();
                if (dt != null && dt.Rows.Count > 0)
                {
                    var storeNOs = new HashSet<string>();
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (!dr.IsNull("StoreNO"))
                        {
                            string storeNO = dr["StoreNO"].ToString();
                            if (!string.IsNullOrEmpty(storeNO) && !storeNOs.Contains(storeNO))
                            {
                                storeNOs.Add(storeNO);
                            }
                        }
                    }

                    // 批次查詢所有 Store 資料
                    if (storeNOs.Count > 0)
                    {
                        string storeNOsStr = string.Join("','", storeNOs);
                        DataTable dt_Store = db.DB_GetData($"SELECT StoreNO, StoreName FROM SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO IN ('{storeNOsStr}')");
                        if (dt_Store != null)
                        {
                            foreach (DataRow row in dt_Store.Rows)
                            {
                                if (!row.IsNull("StoreNO") && !row.IsNull("StoreName"))
                                {
                                    storeCache[row["StoreNO"].ToString()] = row["StoreName"].ToString();
                                }
                            }
                        }
                    }

                    // 處理資料列
                    foreach (DataRow dr in dt.Rows)
                    {
                        string storeName = "";
                        if (!dr.IsNull("StoreNO"))
                        {
                            string storeNO = dr["StoreNO"].ToString();
                            if (storeCache.TryGetValue(storeNO, out string cachedStoreName))
                            {
                                storeName = cachedStoreName;
                            }
                        }

                        // 效能優化：預先轉換所有欄位，避免重複轉換和 null 檢查
                        string mfno = dr.IsNull("MFNO") ? "" : dr["MFNO"].ToString();
                        string mfname = dr.IsNull("MFName") ? "" : dr["MFName"].ToString();
                        string uniFormNO = dr.IsNull("UniFormNO") ? "" : dr["UniFormNO"].ToString();
                        string contactMan = dr.IsNull("ContactMan") ? "" : dr["ContactMan"].ToString();
                        string contactTEL = dr.IsNull("ContactTEL") ? "" : dr["ContactTEL"].ToString();
                        string email = dr.IsNull("EMail") ? "" : dr["EMail"].ToString();
                        string storeNO_display = dr.IsNull("StoreNO") ? "" : dr["StoreNO"].ToString();
                        string ctDataWeights = dr.IsNull("CTDataWeights") ? "" : dr["CTDataWeights"].ToString();
                        string address = dr.IsNull("Address") ? "" : dr["Address"].ToString();
                        string remark = dr.IsNull("Remark") ? "" : dr["Remark"].ToString();

                        sb.Append($"<tr><td>{mfno}</td><td>{mfname}</td><td>{uniFormNO}</td><td>{contactMan}</td><td>{contactTEL}</td><td>{email}</td><td>{storeNO_display}{storeName}</td><td>{ctDataWeights}</td><td>{address}</td><td>{remark}</td></tr>");
                    }
                }
                sb.Append("</table>");
            }
            ViewBag.HtmlOutput = sb.ToString();
            return View();
        }
        public IActionResult ALLStationOrderNO()
        {
            var sb = new StringBuilder();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                sb.Append("<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>");
                sb.Append("<tr><td>工站編號</td><td>工站名稱</td><td>在線工單編號</td><td>狀態</td><td>料件編號</td><td>品名/規格</td><td>需求數量</td><td>單位</td><td>上站工站</td><td>移轉入量</td><td>本站報工量</td><td>下站工站</td><td>移出下站量</td><td>入庫量</td></tr>");

                // 查詢 Manufacture 資料
                DataTable dt = db.DB_GetData($"select * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and OrderNO!='' order by StationNO");

                // 效能優化：批次載入所有需要的資料，避免 N+1 查詢問題
                Dictionary<string, DataRow> materialCache = new Dictionary<string, DataRow>();
                Dictionary<string, string> stationCache = new Dictionary<string, string>();
                Dictionary<string, DataRow> partNOTimeNoteCache = new Dictionary<string, DataRow>();
                Dictionary<string, DataRow> simulationCache = new Dictionary<string, DataRow>();
                Dictionary<string, Dictionary<string, DataRow>> beforeStationCache = new Dictionary<string, Dictionary<string, DataRow>>();

                if (dt != null && dt.Rows.Count > 0)
                {
                    var partNOs = new HashSet<string>();
                    var stationNOs = new HashSet<string>();
                    var simulationIds = new HashSet<string>();

                    // 收集所有需要的 ID
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (!dr.IsNull("PartNO")) { partNOs.Add(dr["PartNO"].ToString()); }
                        if (!dr.IsNull("StationNO")) { stationNOs.Add(dr["StationNO"].ToString()); }
                        if (!dr.IsNull("SimulationId")) { simulationIds.Add(dr["SimulationId"].ToString()); }
                    }

                    // 批次載入 Material 資料
                    if (partNOs.Count > 0)
                    {
                        string partNOsStr = string.Join("','", partNOs);
                        DataTable dt_Material = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO IN ('{partNOsStr}')");
                        if (dt_Material != null)
                        {
                            foreach (DataRow row in dt_Material.Rows)
                            {
                                if (!row.IsNull("PartNO"))
                                {
                                    materialCache[row["PartNO"].ToString()] = row;
                                }
                            }
                        }
                    }

                    // 批次載入 PP_Station 資料
                    if (stationNOs.Count > 0)
                    {
                        string stationNOsStr = string.Join("','", stationNOs);
                        DataTable dt_Station = db.DB_GetData($"SELECT StationNO, StationName FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO IN ('{stationNOsStr}')");
                        if (dt_Station != null)
                        {
                            foreach (DataRow row in dt_Station.Rows)
                            {
                                if (!row.IsNull("StationNO") && !row.IsNull("StationName"))
                                {
                                    stationCache[row["StationNO"].ToString()] = row["StationName"].ToString();
                                }
                            }
                        }
                    }

                    // 批次載入 APS_PartNOTimeNote 資料
                    if (simulationIds.Count > 0)
                    {
                        string simulationIdsStr = string.Join("','", simulationIds);
                        DataTable dt_PartNOTimeNote = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId IN ('{simulationIdsStr}')");
                        if (dt_PartNOTimeNote != null)
                        {
                            foreach (DataRow row in dt_PartNOTimeNote.Rows)
                            {
                                if (!row.IsNull("SimulationId"))
                                {
                                    partNOTimeNoteCache[row["SimulationId"].ToString()] = row;
                                }
                            }
                        }
                    }

                    // 批次載入 APS_Simulation 資料
                    if (simulationIds.Count > 0)
                    {
                        string simulationIdsStr = string.Join("','", simulationIds);
                        DataTable dt_Simulation = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId IN ('{simulationIdsStr}')");
                        if (dt_Simulation != null)
                        {
                            foreach (DataRow row in dt_Simulation.Rows)
                            {
                                if (!row.IsNull("SimulationId"))
                                {
                                    simulationCache[row["SimulationId"].ToString()] = row;
                                }
                            }
                        }
                    }

                    // 效能優化：批次載入前站 SimulationId 和 APS_PartNOTimeNote 資料
                    Dictionary<string, string> beforeSimulationIdCache = new Dictionary<string, string>();
                    Dictionary<string, DataRow> beforePartNOTimeNoteCache = new Dictionary<string, DataRow>();
                    
                    // 收集所有需要查詢前站的條件
                    var beforeStationKeys = new HashSet<string>();
                    var beforeStationConditions = new List<string>();
                    foreach (DataRow dr in dt.Rows)
                    {
                        string simulationId = dr.IsNull("SimulationId") ? "" : dr["SimulationId"].ToString();
                        if (simulationCache.TryGetValue(simulationId, out DataRow simRow))
                        {
                            string needId = simRow.IsNull("NeedId") ? "" : simRow["NeedId"].ToString();
                            string applyPPName = simRow.IsNull("Apply_PP_Name") ? "" : simRow["Apply_PP_Name"].ToString();
                            string sourceStationNO = simRow.IsNull("Source_StationNO") ? "" : simRow["Source_StationNO"].ToString();
                            string indexSN = simRow.IsNull("IndexSN") ? "" : simRow["IndexSN"].ToString();
                            
                            if (!string.IsNullOrEmpty(needId) && !string.IsNullOrEmpty(applyPPName) &&
                                !string.IsNullOrEmpty(sourceStationNO) && !string.IsNullOrEmpty(indexSN) &&
                                int.TryParse(indexSN, out int indexSNInt) && indexSNInt > 0)
                            {
                                string key = $"{needId}|{applyPPName}|{sourceStationNO}|{indexSNInt - 1}";
                                if (!beforeStationKeys.Contains(key))
                                {
                                    beforeStationKeys.Add(key);
                                    beforeStationConditions.Add($"(NeedId='{needId}' and Apply_PP_Name='{applyPPName}' and IndexSN={indexSNInt - 1} and Apply_StationNO='{sourceStationNO}')");
                                }
                            }
                        }
                    }
                    
                    // 批次載入前站 SimulationId
                    if (beforeStationConditions.Count > 0)
                    {
                        string whereClause = string.Join(" or ", beforeStationConditions);
                        DataTable dt_BeforeSim = db.DB_GetData($"SELECT SimulationId, NeedId, Apply_PP_Name, Apply_StationNO, IndexSN FROM SoftNetSYSDB.[dbo].[APS_Simulation] where Source_StationNO is not NULL and ({whereClause})");
                        if (dt_BeforeSim != null)
                        {
                            var beforeSimulationIds = new HashSet<string>();
                            foreach (DataRow row in dt_BeforeSim.Rows)
                            {
                                if (!row.IsNull("SimulationId"))
                                {
                                    string needId = row["NeedId"].ToString();
                                    string applyPPName = row["Apply_PP_Name"].ToString();
                                    string sourceStationNO = row["Apply_StationNO"].ToString();
                                    string indexSN = row["IndexSN"].ToString();
                                    string key = $"{needId}|{applyPPName}|{sourceStationNO}|{indexSN}";
                                    string simulationId = row["SimulationId"].ToString();
                                    beforeSimulationIdCache[key] = simulationId;
                                    beforeSimulationIds.Add(simulationId);
                                }
                            }
                            
                            // 批次載入前站 APS_PartNOTimeNote 資料
                            if (beforeSimulationIds.Count > 0)
                            {
                                string beforeSimulationIdsStr = string.Join("','", beforeSimulationIds);
                                DataTable dt_BeforePartNOTimeNote = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId IN ('{beforeSimulationIdsStr}') and NoStation='0'");
                                if (dt_BeforePartNOTimeNote != null)
                                {
                                    foreach (DataRow row in dt_BeforePartNOTimeNote.Rows)
                                    {
                                        if (!row.IsNull("SimulationId"))
                                        {
                                            beforePartNOTimeNoteCache[row["SimulationId"].ToString()] = row;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // 處理資料列
                    foreach (DataRow dr in dt.Rows)
                    {
                        string fontClass = " style='background-color:whitesmoke;'";
                        string state = "";
                        switch (dr.IsNull("State") ? "" : dr["State"].ToString())
                        {
                            case "1": state = "開工中"; fontClass = " style='background-color:springgreen;'"; break;
                            case "2": state = "停工"; fontClass = " style='background-color:tomato;'"; break;
                            default: state = "閒置"; break;
                        }

                        string partNO = dr.IsNull("PartNO") ? "" : dr["PartNO"].ToString();
                        string simulationId = dr.IsNull("SimulationId") ? "" : dr["SimulationId"].ToString();
                        string stationNO = dr.IsNull("StationNO") ? "" : dr["StationNO"].ToString();

                        if (materialCache.TryGetValue(partNO, out DataRow materialRow) &&
                            partNOTimeNoteCache.TryGetValue(simulationId, out DataRow partNOTimeNoteRow))
                        {
                            string disbeforStaion = "";
                            string disbeforStaionQTY = "";

                            // 計算前站完成量（使用快取）
                            if (simulationCache.TryGetValue(simulationId, out DataRow simRow))
                            {
                                string needId = simRow.IsNull("NeedId") ? "" : simRow["NeedId"].ToString();
                                string applyPPName = simRow.IsNull("Apply_PP_Name") ? "" : simRow["Apply_PP_Name"].ToString();
                                string sourceStationNO = simRow.IsNull("Source_StationNO") ? "" : simRow["Source_StationNO"].ToString();
                                string indexSN = simRow.IsNull("IndexSN") ? "" : simRow["IndexSN"].ToString();

                                if (!string.IsNullOrEmpty(needId) && !string.IsNullOrEmpty(applyPPName) &&
                                    !string.IsNullOrEmpty(sourceStationNO) && !string.IsNullOrEmpty(indexSN) &&
                                    int.TryParse(indexSN, out int indexSNInt) && indexSNInt > 0)
                                {
                                    // 使用快取的前站 SimulationId
                                    string key = $"{needId}|{applyPPName}|{sourceStationNO}|{indexSNInt - 1}";
                                    if (beforeSimulationIdCache.TryGetValue(key, out string beforeSimulationId))
                                    {
                                        // 使用快取的前站 APS_PartNOTimeNote 資料
                                        if (beforePartNOTimeNoteCache.TryGetValue(beforeSimulationId, out DataRow beforePartNOTimeNoteRow))
                                        {
                                            disbeforStaion = beforePartNOTimeNoteRow.IsNull("APS_StationNO") ? "" : beforePartNOTimeNoteRow["APS_StationNO"].ToString();
                                            disbeforStaionQTY = beforePartNOTimeNoteRow.IsNull("Detail_QTY") ? "" : beforePartNOTimeNoteRow["Detail_QTY"].ToString();
                                        }
                                    }
                                }
                            }

                            string stationName = stationCache.TryGetValue(stationNO, out string cachedStationName) ? cachedStationName : "";
                            string partName = materialRow.IsNull("PartName") ? "" : materialRow["PartName"].ToString();
                            string specification = materialRow.IsNull("Specification") ? "" : materialRow["Specification"].ToString();
                            string unit = materialRow.IsNull("Unit") ? "" : materialRow["Unit"].ToString();
                            string orderNO = dr.IsNull("OrderNO") ? "" : dr["OrderNO"].ToString();
                            string needQTY = partNOTimeNoteRow.IsNull("NeedQTY") ? "" : partNOTimeNoteRow["NeedQTY"].ToString();
                            if (!string.IsNullOrEmpty(needQTY) && simulationCache.TryGetValue(simulationId, out DataRow simRowForQty))
                            {
                                needQTY = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, simRowForQty["NeedId"].ToString(), simulationId, int.Parse(needQTY)).ToString();
                            }
                            string detailQTY = partNOTimeNoteRow.IsNull("Detail_QTY") ? "" : partNOTimeNoteRow["Detail_QTY"].ToString();
                            string detailFailQTY = partNOTimeNoteRow.IsNull("Detail_Fail_QTY") ? "" : partNOTimeNoteRow["Detail_Fail_QTY"].ToString();
                            string nextAPSStationNO = partNOTimeNoteRow.IsNull("Next_APS_StationNO") ? "" : partNOTimeNoteRow["Next_APS_StationNO"].ToString();
                            string nextStationQTY = partNOTimeNoteRow.IsNull("Next_StationQTY") ? "" : partNOTimeNoteRow["Next_StationQTY"].ToString();
                            string nextStoreQTY = partNOTimeNoteRow.IsNull("Next_StoreQTY") ? "" : partNOTimeNoteRow["Next_StoreQTY"].ToString();

                            sb.Append($"<tr><td>{stationNO}</td><td>{stationName}</td><td>{orderNO}</td><td{fontClass}>{state}</td><td>{partNO}</td><td>{partName}&nbsp;{specification}</td><td>{needQTY}</td><td>{unit}</td><td>{disbeforStaion}</td><td>{disbeforStaionQTY}</td><td>{detailQTY}/{detailFailQTY}</td><td>{nextAPSStationNO}</td><td>{nextStationQTY}</td><td>{nextStoreQTY}</td></tr>");
                        }
                    }
                }

                // 處理 ManufactureII 資料
                dt = db.DB_GetData($"select * FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and EndTime is NULL order by StationNO");
                string beforStation = "";
                string disStaion = "";

                if (dt != null && dt.Rows.Count > 0)
                {
                    // 重新收集需要的 ID（因為是新的資料表）
                    var partNOs2 = new HashSet<string>();
                    var stationNOs2 = new HashSet<string>();
                    var simulationIds2 = new HashSet<string>();

                    foreach (DataRow dr in dt.Rows)
                    {
                        if (!dr.IsNull("PartNO")) { partNOs2.Add(dr["PartNO"].ToString()); }
                        if (!dr.IsNull("StationNO")) { stationNOs2.Add(dr["StationNO"].ToString()); }
                        if (!dr.IsNull("SimulationId")) { simulationIds2.Add(dr["SimulationId"].ToString()); }
                    }

                    // 批次載入 Material 資料（如果尚未載入）
                    foreach (string partNO in partNOs2)
                    {
                        if (!materialCache.ContainsKey(partNO))
                        {
                            DataRow materialRow = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}'");
                            if (materialRow != null)
                            {
                                materialCache[partNO] = materialRow;
                            }
                        }
                    }

                    // 批次載入 PP_Station 資料（如果尚未載入）
                    foreach (string stationNO in stationNOs2)
                    {
                        if (!stationCache.ContainsKey(stationNO))
                        {
                            DataRow stationRow = db.DB_GetFirstDataByDataRow($"SELECT StationName FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}'");
                            if (stationRow != null && !stationRow.IsNull("StationName"))
                            {
                                stationCache[stationNO] = stationRow["StationName"].ToString();
                            }
                        }
                    }

                    // 批次載入 APS_PartNOTimeNote 資料（如果尚未載入）
                    foreach (string simulationId in simulationIds2)
                    {
                        if (!partNOTimeNoteCache.ContainsKey(simulationId))
                        {
                            DataRow partNOTimeNoteRow = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{simulationId}'");
                            if (partNOTimeNoteRow != null)
                            {
                                partNOTimeNoteCache[simulationId] = partNOTimeNoteRow;
                            }
                        }
                    }

                    // 批次載入 APS_Simulation 資料（如果尚未載入）
                    foreach (string simulationId in simulationIds2)
                    {
                        if (!simulationCache.ContainsKey(simulationId))
                        {
                            DataRow simRow = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{simulationId}'");
                            if (simRow != null)
                            {
                                simulationCache[simulationId] = simRow;
                            }
                        }
                    }

                    // 處理資料列
                    foreach (DataRow dr in dt.Rows)
                    {
                        string fontClass = " style='background-color:whitesmoke;'";
                        string state = "";
                        if (!dr.IsNull("RemarkTimeS") && dr.IsNull("RemarkTimeE"))
                        {
                            state = "開工中";
                            fontClass = " style='background-color:springgreen;'";
                        }
                        else if (!dr.IsNull("RemarkTimeE"))
                        {
                            state = "停工";
                            fontClass = " style='background-color:tomato;'";
                        }
                        else
                        {
                            state = "閒置";
                        }

                        string currentStationNO = dr.IsNull("StationNO") ? "" : dr["StationNO"].ToString();
                        if (beforStation != currentStationNO)
                        {
                            beforStation = currentStationNO;
                            disStaion = beforStation;
                        }
                        else
                        {
                            disStaion = "";
                        }

                        string partNO = dr.IsNull("PartNO") ? "" : dr["PartNO"].ToString();
                        string simulationId = dr.IsNull("SimulationId") ? "" : dr["SimulationId"].ToString();

                        if (materialCache.TryGetValue(partNO, out DataRow materialRow) &&
                            partNOTimeNoteCache.TryGetValue(simulationId, out DataRow partNOTimeNoteRow))
                        {
                            string disbeforStaion = "";
                            string disbeforStaionQTY = "";

                            // 計算前站完成量
                            if (simulationCache.TryGetValue(simulationId, out DataRow simRow))
                            {
                                string needId = simRow.IsNull("NeedId") ? "" : simRow["NeedId"].ToString();
                                string applyPPName = simRow.IsNull("Apply_PP_Name") ? "" : simRow["Apply_PP_Name"].ToString();
                                string sourceStationNO = simRow.IsNull("Source_StationNO") ? "" : simRow["Source_StationNO"].ToString();
                                string indexSN = simRow.IsNull("IndexSN") ? "" : simRow["IndexSN"].ToString();

                                if (!string.IsNullOrEmpty(needId) && !string.IsNullOrEmpty(applyPPName) &&
                                    !string.IsNullOrEmpty(sourceStationNO) && !string.IsNullOrEmpty(indexSN) &&
                                    int.TryParse(indexSN, out int indexSNInt) && indexSNInt > 0)
                                {
                                    DataRow beforeSimRow = db.DB_GetFirstDataByDataRow($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation] where Source_StationNO is not NULL and NeedId='{needId}' and Apply_PP_Name='{applyPPName}' and IndexSN={(indexSNInt - 1).ToString()} and Apply_StationNO='{sourceStationNO}'");
                                    if (beforeSimRow != null && !beforeSimRow.IsNull("SimulationId"))
                                    {
                                        string beforeSimulationId = beforeSimRow["SimulationId"].ToString();
                                        DataRow beforePartNOTimeNoteRow = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].APS_PartNOTimeNote where SimulationId='{beforeSimulationId}' and NoStation='0'");
                                        if (beforePartNOTimeNoteRow != null)
                                        {
                                            disbeforStaion = beforePartNOTimeNoteRow.IsNull("APS_StationNO") ? "" : beforePartNOTimeNoteRow["APS_StationNO"].ToString();
                                            disbeforStaionQTY = beforePartNOTimeNoteRow.IsNull("Detail_QTY") ? "" : beforePartNOTimeNoteRow["Detail_QTY"].ToString();
                                        }
                                    }
                                }
                            }

                            string stationName = stationCache.TryGetValue(currentStationNO, out string cachedStationName) ? cachedStationName : "";
                            string partName = materialRow.IsNull("PartName") ? "" : materialRow["PartName"].ToString();
                            string specification = materialRow.IsNull("Specification") ? "" : materialRow["Specification"].ToString();
                            string unit = materialRow.IsNull("Unit") ? "" : materialRow["Unit"].ToString();
                            string orderNO = dr.IsNull("OrderNO") ? "" : dr["OrderNO"].ToString();
                            string needQTY = partNOTimeNoteRow.IsNull("NeedQTY") ? "" : partNOTimeNoteRow["NeedQTY"].ToString();
                            if (!string.IsNullOrEmpty(needQTY) && simulationCache.TryGetValue(simulationId, out DataRow simRowForQty))
                            {
                                needQTY = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, simRowForQty["NeedId"].ToString(), simulationId, int.Parse(needQTY)).ToString();
                            }
                            string detailQTY = partNOTimeNoteRow.IsNull("Detail_QTY") ? "" : partNOTimeNoteRow["Detail_QTY"].ToString();
                            string detailFailQTY = partNOTimeNoteRow.IsNull("Detail_Fail_QTY") ? "" : partNOTimeNoteRow["Detail_Fail_QTY"].ToString();
                            string nextAPSStationNO = partNOTimeNoteRow.IsNull("Next_APS_StationNO") ? "" : partNOTimeNoteRow["Next_APS_StationNO"].ToString();
                            string nextStationQTY = partNOTimeNoteRow.IsNull("Next_StationQTY") ? "" : partNOTimeNoteRow["Next_StationQTY"].ToString();
                            string nextStoreQTY = partNOTimeNoteRow.IsNull("Next_StoreQTY") ? "" : partNOTimeNoteRow["Next_StoreQTY"].ToString();

                            sb.Append($"<tr><td>{disStaion}</td><td>{stationName}</td><td>{orderNO}</td><td{fontClass}>{state}</td><td>{partNO}</td><td>{partName}&nbsp;{specification}</td><td>{needQTY}</td><td>{unit}</td><td>{disbeforStaion}</td><td>{disbeforStaionQTY}</td><td>{detailQTY}/{detailFailQTY}</td><td>{nextAPSStationNO}</td><td>{nextStationQTY}</td><td>{nextStoreQTY}</td></tr>");
                        }
                    }
                }

                sb.Append("</table>");
            }
            ViewBag.HtmlOutput = sb.ToString();
            return View();
        }
        public IActionResult AllMaterial()
        {
            var sb = new StringBuilder();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string storeName = "";
                string className = "";
                string typeName = "";
                string iS_Store_Test = "";
                string aPS_Default_MFNO = "";
                sb.Append("<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>");
                sb.Append("<tr><td>料件編號</td><td>品名</td><td>規格</td><td>機種</td><td>類型</td><td>安全量</td><td>單位</td><td>備料工時</td><td>檢驗</td><td>指定供應商</td><td>指定倉庫</td><td>定義</td></tr>");
                DataTable dt = db.DB_GetData($"select * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' order by Class,PartNO,APS_Default_StoreNO,APS_Default_StoreSpacesNO,PartType");
                
                // 效能優化：批次載入所有需要的 Store 和 MFData 資料，避免 N+1 查詢問題
                var storeNOs = new HashSet<string>();
                var mfNOs = new HashSet<string>();
                
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (!dr.IsNull("APS_Default_StoreNO"))
                        {
                            string storeNO = dr["APS_Default_StoreNO"].ToString();
                            if (!string.IsNullOrEmpty(storeNO) && !storeNOs.Contains(storeNO))
                            {
                                storeNOs.Add(storeNO);
                            }
                        }
                        if (!dr.IsNull("APS_Default_MFNO"))
                        {
                            string mfNO = dr["APS_Default_MFNO"].ToString();
                            if (!string.IsNullOrEmpty(mfNO) && !mfNOs.Contains(mfNO))
                            {
                                mfNOs.Add(mfNO);
                            }
                        }
                    }
                    
                    // 批次載入 Store 資料
                    Dictionary<string, string> storeCache = new Dictionary<string, string>();
                    if (storeNOs.Count > 0)
                    {
                        string storeNOsStr = string.Join("','", storeNOs);
                        DataTable dt_Store = db.DB_GetData($"SELECT StoreNO, StoreName FROM SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO IN ('{storeNOsStr}')");
                        if (dt_Store != null)
                        {
                            foreach (DataRow row in dt_Store.Rows)
                            {
                                if (!row.IsNull("StoreNO") && !row.IsNull("StoreName"))
                                {
                                    storeCache[row["StoreNO"].ToString()] = row["StoreName"].ToString();
                                }
                            }
                        }
                    }
                    
                    // 批次載入 MFData 資料
                    Dictionary<string, string> mfCache = new Dictionary<string, string>();
                    if (mfNOs.Count > 0)
                    {
                        string mfNOsStr = string.Join("','", mfNOs);
                        DataTable dt_MF = db.DB_GetData($"SELECT MFNO, SName FROM SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' and MFNO IN ('{mfNOsStr}')");
                        if (dt_MF != null)
                        {
                            foreach (DataRow row in dt_MF.Rows)
                            {
                                if (!row.IsNull("MFNO") && !row.IsNull("SName"))
                                {
                                    mfCache[row["MFNO"].ToString()] = row["SName"].ToString();
                                }
                            }
                        }
                    }
                    
                    // 處理資料列
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
                        if (dr["PartType"].ToString() == "0")
                        { typeName = "實料"; }
                        else { typeName = "虛料"; }
                        if (bool.Parse(dr["IS_Store_Test"].ToString()))
                        { iS_Store_Test = "是"; }
                        else { iS_Store_Test = "否"; }
                        
                        // 使用快取的 Store 資料
                        if (!dr.IsNull("APS_Default_StoreNO"))
                        {
                            string storeNO = dr["APS_Default_StoreNO"].ToString();
                            if (storeCache.TryGetValue(storeNO, out string cachedStoreName))
                            {
                                string storeSpacesNO = dr.IsNull("APS_Default_StoreSpacesNO") ? "" : dr["APS_Default_StoreSpacesNO"].ToString();
                                storeName = $"{cachedStoreName}&nbsp;{storeSpacesNO}";
                            }
                        }
                        
                        // 使用快取的 MFData 資料
                        if (!dr.IsNull("APS_Default_MFNO"))
                        {
                            string mfNO = dr["APS_Default_MFNO"].ToString();
                            if (mfCache.TryGetValue(mfNO, out string cachedMFName))
                            {
                                aPS_Default_MFNO = cachedMFName;
                            }
                        }
                        
                        sb.Append($"<tr><td>{dr["PartNO"].ToString()}</td><td>{dr["PartName"].ToString()}</td><td>{dr["Specification"].ToString()}</td><td>{dr["Model"].ToString()}</td><td>{className}</td><td>{dr["SafeQTY"].ToString()}</td><td>{dr["Unit"].ToString()}</td><td>{dr["StoreSTime"].ToString()}</td><td>{iS_Store_Test}</td><td>{dr["APS_Default_MFNO"].ToString()}</td><td>{storeName}</td><td>{typeName}</td></tr>");
                    }
                }
                sb.Append("</table>");
            }
            ViewBag.HtmlOutput = sb.ToString();
            return View();
        }

        public IActionResult ALLStore(APSViewData keyParam)
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                if (keyParam == null || keyParam.SelectFun1 == "")
                {
                    keyParam = new APSViewData();
                    keyParam.SelectFun2 = "4";
                    keyParam.SelectFun3 = "";
                }

                // 效能優化：建立 WHERE 條件
                string wheresql = "";
                if (!string.IsNullOrEmpty(keyParam.SelectFun2))
                {
                    var classList = new List<string>();
                    foreach (string s in keyParam.SelectFun2.Split(';'))
                    {
                        if (!string.IsNullOrEmpty(s)) { classList.Add($"Class='{s}'"); }
                    }
                    if (classList.Count > 0) { wheresql = $" and ({string.Join(" or ", classList)})"; }
                }

                var sb = new StringBuilder();
                sb.Append("<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>");
                sb.Append("<tr><td>料件編號</td><td>品名/規格</td><td>類型</td><td>安全量</td><td>單位</td><td>再倉量</td><td>再製量</td><td>再途量</td><td>保留量</td><td>未來量</td></tr>");

                // 效能優化：簡化 SQL，只選擇需要的欄位
                DataTable dt = db.DB_GetData($"select PartNO,PartName,Specification,Class,SafeQTY,Unit FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' {wheresql} order by Class,PartNO,APS_Default_StoreNO,APS_Default_StoreSpacesNO,PartType");

                if (dt == null || dt.Rows.Count == 0)
                {
                    sb.Append("</table>");
                    ViewBag.HtmlOutput = sb.ToString();
                    return View();
                }

                // 效能優化：批次載入所有需要的資料，避免在迴圈中重複查詢
                var partNOs = new HashSet<string>();
                foreach (DataRow dr in dt.Rows)
                {
                    if (!dr.IsNull("PartNO")) { partNOs.Add(dr["PartNO"].ToString()); }
                }

                // 批次載入再倉量（TotalStock）
                Dictionary<string, int> sQTYCache = new Dictionary<string, int>();
                if (partNOs.Count > 0)
                {
                    string partNOsStr = string.Join("','", partNOs);
                    DataTable dt_TotalStock = db.DB_GetData($"SELECT PartNO, sum(QTY) as S_QTY from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO IN ('{partNOsStr}') group by PartNO");
                    if (dt_TotalStock != null)
                    {
                        foreach (DataRow row in dt_TotalStock.Rows)
                        {
                            if (!row.IsNull("PartNO"))
                            {
                                string partNO = row["PartNO"].ToString();
                                int qty = row.IsNull("S_QTY") ? 0 : int.Parse(row["S_QTY"].ToString());
                                sQTYCache[partNO] = qty;
                            }
                        }
                    }
                }

                // 批次載入保留量（TotalStockII + DOC3stockII）
                Dictionary<string, int> keepQTYCache = new Dictionary<string, int>();
                if (partNOs.Count > 0)
                {
                    string partNOsStr = string.Join("','", partNOs);
                    // TotalStockII
                    DataTable dt_TotalStockII = db.DB_GetData($@"SELECT a.PartNO, sum(b.KeepQTY+b.OverQTY) as Keep_QTY from SoftNetMainDB.[dbo].[TotalStock] as a
                                                        join SoftNetMainDB.[dbo].[TotalStockII] as b on a.Id=b.Id 
                                                        where a.ServerId='{_Fun.Config.ServerId}' and a.PartNO IN ('{partNOsStr}') group by a.PartNO");
                    if (dt_TotalStockII != null)
                    {
                        foreach (DataRow row in dt_TotalStockII.Rows)
                        {
                            if (!row.IsNull("PartNO"))
                            {
                                string partNO = row["PartNO"].ToString();
                                int qty = row.IsNull("Keep_QTY") ? 0 : int.Parse(row["Keep_QTY"].ToString());
                                keepQTYCache[partNO] = qty;
                            }
                        }
                    }
                    // DOC3stockII
                    DataTable dt_DOC3 = db.DB_GetData($"SELECT PartNO, sum(QTY) as DOC3_Out_QTY from SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and PartNO IN ('{partNOsStr}') and OUT_StoreNO!='' and IsOK='0' group by PartNO");
                    if (dt_DOC3 != null)
                    {
                        foreach (DataRow row in dt_DOC3.Rows)
                        {
                            if (!row.IsNull("PartNO"))
                            {
                                string partNO = row["PartNO"].ToString();
                                int qty = row.IsNull("DOC3_Out_QTY") ? 0 : int.Parse(row["DOC3_Out_QTY"].ToString());
                                if (keepQTYCache.ContainsKey(partNO))
                                {
                                    keepQTYCache[partNO] += qty;
                                }
                                else
                                {
                                    keepQTYCache[partNO] = qty;
                                }
                            }
                        }
                    }
                }

                // 批次載入再製量（APS_Simulation + APS_PartNOTimeNote）
                Dictionary<string, int> pQTYCache = new Dictionary<string, int>();
                if (partNOs.Count > 0)
                {
                    string partNOsStr = string.Join("','", partNOs);
                    DataTable dt_APS = db.DB_GetData($@"select a.PartNO, a.Master_PartNO, a.NeedId, a.SimulationId, a.Apply_PP_Name, a.IndexSN, b.NeedQTY from SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                            join SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as b on a.SimulationId=b.SimulationId
                                            where a.IsOK='0' and (a.Class='4' or a.Class='5') and a.PartNO IN ('{partNOsStr}') and a.Source_StationNO!='{_Fun.Config.OutPackStationName}' and a.Source_StationNO is not null order by a.PartNO, a.NeedId, a.PartSN desc");
                    if (dt_APS != null)
                    {
                        string lastPartNO = "";
                        string lastMasterPartNO = "";
                        string lastNeedId = "";
                        foreach (DataRow row in dt_APS.Rows)
                        {
                            string partNO = row.IsNull("PartNO") ? "" : row["PartNO"].ToString();
                            string masterPartNO = row.IsNull("Master_PartNO") ? "" : row["Master_PartNO"].ToString();
                            string needId = row.IsNull("NeedId") ? "" : row["NeedId"].ToString();
                            int needQTY = row.IsNull("NeedQTY") ? 0 : int.Parse(row["NeedQTY"].ToString());
                            needQTY = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, needId, row["SimulationId"].ToString(), needQTY);

                            // 效能優化：只計算每個 PartNO+Master_PartNO+NeedId 組合一次
                            string cacheKey = $"{partNO}|{masterPartNO}|{needId}";
                            string lastCacheKey = $"{lastPartNO}|{lastMasterPartNO}|{lastNeedId}";

                            if (cacheKey != lastCacheKey)
                            {
                                if (pQTYCache.ContainsKey(partNO))
                                {
                                    pQTYCache[partNO] += needQTY;
                                }
                                else
                                {
                                    pQTYCache[partNO] = needQTY;
                                }
                                lastPartNO = partNO;
                                lastMasterPartNO = masterPartNO;
                                lastNeedId = needId;
                            }
                        }
                    }
                }

                // 批次載入在途量（DOC1BuyII + DOC4ProductionII）
                Dictionary<string, int> gQTYCache = new Dictionary<string, int>();
                if (partNOs.Count > 0)
                {
                    string partNOsStr = string.Join("','", partNOs);
                    // DOC1BuyII
                    DataTable dt_DOC1 = db.DB_GetData($"SELECT PartNO, sum(QTY) as G_QTY from SoftNetMainDB.[dbo].[DOC1BuyII] where ServerId='{_Fun.Config.ServerId}' and IsOK='0' and PartNO IN ('{partNOsStr}') group by PartNO");
                    if (dt_DOC1 != null)
                    {
                        foreach (DataRow row in dt_DOC1.Rows)
                        {
                            if (!row.IsNull("PartNO"))
                            {
                                string partNO = row["PartNO"].ToString();
                                int qty = row.IsNull("G_QTY") ? 0 : int.Parse(row["G_QTY"].ToString());
                                gQTYCache[partNO] = qty;
                            }
                        }
                    }
                    // DOC4ProductionII
                    DataTable dt_DOC4 = db.DB_GetData($"SELECT PartNO, sum(QTY) as G_QTY from SoftNetMainDB.[dbo].[DOC4ProductionII] where ServerId='{_Fun.Config.ServerId}' and IsOK='0' and PartNO IN ('{partNOsStr}') group by PartNO");
                    if (dt_DOC4 != null)
                    {
                        foreach (DataRow row in dt_DOC4.Rows)
                        {
                            if (!row.IsNull("PartNO"))
                            {
                                string partNO = row["PartNO"].ToString();
                                int qty = row.IsNull("G_QTY") ? 0 : int.Parse(row["G_QTY"].ToString());
                                if (gQTYCache.ContainsKey(partNO))
                                {
                                    gQTYCache[partNO] += qty;
                                }
                                else
                                {
                                    gQTYCache[partNO] = qty;
                                }
                            }
                        }
                    }
                }

                // 處理資料列
                string partNOName = "";
                string className = "";
                int sQTY = 0;
                int keepQTY = 0;
                int pQTY = 0;
                int gQTY = 0;
                int useQTY = 0;

                // 效能優化：建立 Class 對應表，避免重複 switch
                Dictionary<string, string> classMap = new Dictionary<string, string>
                {
                    { "1", "原物料" },
                    { "2", "採購件" },
                    { "3", "委外件" },
                    { "4", "製造半成品" },
                    { "5", "製造成品" },
                    { "6", "刀具" },
                    { "7", "工具製具" }
                };

                foreach (DataRow dr in dt.Rows)
                {
                    sQTY = 0; keepQTY = 0; pQTY = 0; gQTY = 0;
                    string partNO = dr.IsNull("PartNO") ? "" : dr["PartNO"].ToString();
                    string partName = dr.IsNull("PartName") ? "" : dr["PartName"].ToString();
                    string specification = dr.IsNull("Specification") ? "" : dr["Specification"].ToString();
                    partNOName = $"{partName}&nbsp;{specification}";

                    string classValue = dr.IsNull("Class") ? "" : dr["Class"].ToString();
                    classMap.TryGetValue(classValue, out className);
                    if (string.IsNullOrEmpty(className)) { className = ""; }

                    // 使用快取的資料
                    if (!string.IsNullOrEmpty(partNO))
                    {
                        sQTYCache.TryGetValue(partNO, out sQTY);
                        keepQTYCache.TryGetValue(partNO, out keepQTY);
                        pQTYCache.TryGetValue(partNO, out pQTY);
                        gQTYCache.TryGetValue(partNO, out gQTY);
                    }

                    useQTY = sQTY - keepQTY + pQTY + gQTY;

                    // 效能優化：根據 SelectFun3 過濾顯示條件
                    bool shouldDisplay = true;
                    if (!string.IsNullOrEmpty(keyParam.SelectFun3))
                    {
                        var filterList = keyParam.SelectFun3.Split(';').Where(s => !string.IsNullOrEmpty(s)).ToList();
                        if (filterList.Count > 0)
                        {
                            shouldDisplay = false;
                            if (filterList.Contains("1") && keepQTY != 0) { shouldDisplay = true; }
                            if (filterList.Contains("2") && sQTY != 0) { shouldDisplay = true; }
                            if (filterList.Contains("3") && pQTY != 0) { shouldDisplay = true; }
                            if (filterList.Contains("4") && gQTY != 0) { shouldDisplay = true; }
                            if (filterList.Contains("5") && keepQTY != 0) { shouldDisplay = true; }
                            if (filterList.Contains("6") && useQTY != 0) { shouldDisplay = true; }
                        }
                    }
                    // 如果沒有過濾條件，顯示所有資料（shouldDisplay 預設為 true）

                    if (shouldDisplay)
                    {
                        string safeQTY = dr.IsNull("SafeQTY") ? "0" : dr["SafeQTY"].ToString();
                        string unit = dr.IsNull("Unit") ? "" : dr["Unit"].ToString();
                        sb.Append($"<tr><td>{partNO}</td><td>{partNOName}</td><td>{className}</td><td>{safeQTY}</td><td>{unit}</td><td>{sQTY}</td><td>{pQTY}</td><td>{gQTY}</td><td>{keepQTY}</td><td>{useQTY}</td></tr>");
                    }
                }
                sb.Append("</table>");
                re = sb.ToString();
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

                // 效能優化：建立 WHERE 條件
                string wheresql = "";
                if (!string.IsNullOrEmpty(key.SelectFun3))
                {
                    var classList = new List<string>();
                    foreach (string s in key.SelectFun3.Split(';'))
                    {
                        if (!string.IsNullOrEmpty(s)) { classList.Add($"b.Class='{s}'"); }
                    }
                    if (classList.Count > 0) { wheresql = $" and ({string.Join(" or ", classList)})"; }
                }
                if (!string.IsNullOrEmpty(key.SelectFun4)) { wheresql = $"{wheresql} and (a.PartNO like '%{key.SelectFun4}%' or b.PartName like '%{key.SelectFun4}%' or b.Specification like '%{key.SelectFun4}%')"; }

                // 效能優化：預先計算日期字串，避免重複格式化
                string startDateStr = key.SelectFun1;
                string endDateStr = key.SelectFun2;
                string endDateForSQL = endDateStr;
                string nowStr = now.ToString("yyyy-MM-dd");
                bool isEndDateToday = (endDateStr == nowStr);
                var sb = new StringBuilder();
                sb.Append("<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>");
                sb.Append("<tr><td>日期</td><td>單據編號</td><td>料件編號</td><td>品名/規格</td><td>類型</td><td>安全量</td><td>單位</td><td>狀態</td><td>倉庫/儲位</td><td>進倉量</td><td>出倉量</td><td>再途</td><td>帳面數量</td></tr>");

                #region 收集資料
                // 效能優化：使用 UNION ALL 合併所有查詢，減少資料庫往返次數
                // 效能優化：簡化 JOIN 條件，將 SUBSTRING 移到 WHERE 子句以改善索引使用
                string sql = $@"SELECT '-' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],OUT_StoreNO as StoreNO,OUT_StoreSpacesNO as StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and LEFT(a.DOCNumberNO,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{startDateStr}' and StartTime<='{endDateStr}' and OUT_StoreNO!='' and a.StartTime is not NULL {wheresql}
                            UNION ALL
                            SELECT '+' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],IN_StoreNO as StoreNO,IN_StoreSpacesNO as StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and LEFT(a.DOCNumberNO,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{startDateStr}' and StartTime<='{endDateStr}' and IN_StoreNO!='' and a.StartTime is not NULL {wheresql}
                            UNION ALL
                            SELECT '+' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],StoreNO,StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC1BuyII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and LEFT(a.DOCNumberNO,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{startDateStr}' and StartTime<='{endDateStr}' and StoreNO!='' and a.StartTime is not NULL {wheresql}
                            UNION ALL
                            SELECT '-' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],StoreNO,StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC2SalesII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and LEFT(a.DOCNumberNO,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{startDateStr}' and StartTime<='{endDateStr}' and StoreNO!='' and a.StartTime is not NULL {wheresql}
                            UNION ALL
                            SELECT '@' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],StoreNO,StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC4ProductionII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and LEFT(a.DOCNumberNO,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{startDateStr}' and StartTime<='{endDateStr}' and StoreNO!='' and a.StartTime is not NULL {wheresql}
                            UNION ALL
                            SELECT '-' as MathType,c.DOCName,b.PartName,b.Specification,b.Class,b.Unit,b.SafeQTY,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[Remark],a.[SimulationId],a.[IsOK],StoreNO,StoreSpacesNO,a.[EndTime],a.[StartTime] 
                            FROM SoftNetMainDB.[dbo].[DOC5OUTII] as a
                            join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}' 
                            join SoftNetMainDB.[dbo].[DOCRole] as c on c.ServerId='{_Fun.Config.ServerId}' and LEFT(a.DOCNumberNO,4)=c.DOCNO
                            where a.ServerId='{_Fun.Config.ServerId}' and StartTime>='{startDateStr}' and StartTime<='{endDateStr}' and StoreNO!='' and a.StartTime is not NULL {wheresql}";
                DataTable dt = db.DB_GetData(sql);
                #endregion
                string dis_tmp = "";
                if (dt != null && dt.Rows.Count > 0)
                {
                    dt.DefaultView.Sort = "PartNO,StartTime desc";

                    // 效能優化：批次載入所有需要的期末數量資料
                    var partNOs = new HashSet<string>();
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (!dr.IsNull("PartNO")) { partNOs.Add(dr["PartNO"].ToString()); }
                    }

                    // 批次載入 TotalStock 資料
                    Dictionary<string, int> totalStockCache = new Dictionary<string, int>();
                    if (partNOs.Count > 0)
                    {
                        string partNOsStr = string.Join("','", partNOs);
                        DataTable dt_TotalStock = db.DB_GetData($"SELECT PartNO, sum(QTY) as FishQTY FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO IN ('{partNOsStr}') group by PartNO");
                        if (dt_TotalStock != null)
                        {
                            foreach (DataRow row in dt_TotalStock.Rows)
                            {
                                if (!row.IsNull("PartNO"))
                                {
                                    string rowPartNO = row["PartNO"].ToString();
                                    int rowQty = row.IsNull("FishQTY") ? 0 : int.Parse(row["FishQTY"].ToString());
                                    totalStockCache[rowPartNO] = rowQty;
                                }
                            }
                        }
                    }

                    // 批次載入反算期末數量所需的資料（僅當結束日期不是今天時）
                    Dictionary<string, int> fishQTYAdjustCache = new Dictionary<string, int>();
                    if (!isEndDateToday && partNOs.Count > 0)
                    {
                        string partNOsStr = string.Join("','", partNOs);
                        // DOC3stockII OUT
                        DataTable dt_DOC3_OUT = db.DB_GetData($@"SELECT PartNO, sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC3stockII] 
                            where ServerId='{_Fun.Config.ServerId}' and OUT_StoreNO!='' and PartNO IN ('{partNOsStr}') 
                            and StartTime is not NULL and StartTime>'{endDateForSQL}' and StartTime<='{nowStr}' group by PartNO");
                        if (dt_DOC3_OUT != null)
                        {
                            foreach (DataRow row in dt_DOC3_OUT.Rows)
                            {
                                if (!row.IsNull("PartNO"))
                                {
                                    string rowPartNO = row["PartNO"].ToString();
                                    int rowQty = row.IsNull("QTY") ? 0 : int.Parse(row["QTY"].ToString());
                                    fishQTYAdjustCache[rowPartNO] = rowQty;
                                }
                            }
                        }
                        // DOC3stockII IN
                        DataTable dt_DOC3_IN = db.DB_GetData($@"SELECT PartNO, sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC3stockII] 
                            where ServerId='{_Fun.Config.ServerId}' and IN_StoreNO!='' and PartNO IN ('{partNOsStr}') 
                            and StartTime is not NULL and StartTime>'{endDateForSQL}' and StartTime<='{nowStr}' group by PartNO");
                        if (dt_DOC3_IN != null)
                        {
                            foreach (DataRow row in dt_DOC3_IN.Rows)
                            {
                                if (!row.IsNull("PartNO"))
                                {
                                    string rowPartNO = row["PartNO"].ToString();
                                    int rowQty = row.IsNull("QTY") ? 0 : int.Parse(row["QTY"].ToString());
                                    if (fishQTYAdjustCache.ContainsKey(rowPartNO))
                                    {
                                        fishQTYAdjustCache[rowPartNO] -= rowQty;
                                    }
                                    else
                                    {
                                        fishQTYAdjustCache[rowPartNO] = -rowQty;
                                    }
                                }
                            }
                        }
                        // DOC1BuyII
                        DataTable dt_DOC1 = db.DB_GetData($@"SELECT PartNO, sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC1BuyII] 
                            where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and PartNO IN ('{partNOsStr}') 
                            and StartTime is not NULL and StartTime>'{endDateForSQL}' and StartTime<='{nowStr}' group by PartNO");
                        if (dt_DOC1 != null)
                        {
                            foreach (DataRow row in dt_DOC1.Rows)
                            {
                                if (!row.IsNull("PartNO"))
                                {
                                    string rowPartNO = row["PartNO"].ToString();
                                    int rowQty = row.IsNull("QTY") ? 0 : int.Parse(row["QTY"].ToString());
                                    if (fishQTYAdjustCache.ContainsKey(rowPartNO))
                                    {
                                        fishQTYAdjustCache[rowPartNO] -= rowQty;
                                    }
                                    else
                                    {
                                        fishQTYAdjustCache[rowPartNO] = -rowQty;
                                    }
                                }
                            }
                        }
                        // DOC2SalesII
                        DataTable dt_DOC2 = db.DB_GetData($@"SELECT PartNO, sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC2SalesII] 
                            where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and PartNO IN ('{partNOsStr}') 
                            and StartTime is not NULL and StartTime>'{endDateForSQL}' and StartTime<='{nowStr}' group by PartNO");
                        if (dt_DOC2 != null)
                        {
                            foreach (DataRow row in dt_DOC2.Rows)
                            {
                                if (!row.IsNull("PartNO"))
                                {
                                    string rowPartNO = row["PartNO"].ToString();
                                    int rowQty = row.IsNull("QTY") ? 0 : int.Parse(row["QTY"].ToString());
                                    if (fishQTYAdjustCache.ContainsKey(rowPartNO))
                                    {
                                        fishQTYAdjustCache[rowPartNO] += rowQty;
                                    }
                                    else
                                    {
                                        fishQTYAdjustCache[rowPartNO] = rowQty;
                                    }
                                }
                            }
                        }
                        // DOC5OUTII
                        DataTable dt_DOC5 = db.DB_GetData($@"SELECT PartNO, sum(QTY) as QTY FROM SoftNetMainDB.[dbo].[DOC5OUTII] 
                            where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and PartNO IN ('{partNOsStr}') 
                            and StartTime is not NULL and StartTime>'{endDateForSQL}' and StartTime<='{nowStr}' group by PartNO");
                        if (dt_DOC5 != null)
                        {
                            foreach (DataRow row in dt_DOC5.Rows)
                            {
                                if (!row.IsNull("PartNO"))
                                {
                                    string rowPartNO = row["PartNO"].ToString();
                                    int rowQty = row.IsNull("QTY") ? 0 : int.Parse(row["QTY"].ToString());
                                    if (fishQTYAdjustCache.ContainsKey(rowPartNO))
                                    {
                                        fishQTYAdjustCache[rowPartNO] += rowQty;
                                    }
                                    else
                                    {
                                        fishQTYAdjustCache[rowPartNO] = rowQty;
                                    }
                                }
                            }
                        }
                    }

                    // 效能優化：建立 Class 對應表，避免重複 switch
                    Dictionary<string, string> classMap = new Dictionary<string, string>
                    {
                        { "1", "原物料" },
                        { "2", "採購件" },
                        { "3", "委外件" },
                        { "4", "製造半成品" },
                        { "5", "製造成品" },
                        { "6", "刀具" },
                        { "7", "工具製具" }
                    };

                    string partNO = "";
                    int fishQTY = 0;

                    #region first
                    partNO = dt.Rows[0].IsNull("PartNO") ? "" : dt.Rows[0]["PartNO"].ToString();
                    #region 計算期末數量
                    // 使用快取的 TotalStock 資料
                    if (!string.IsNullOrEmpty(partNO) && totalStockCache.TryGetValue(partNO, out fishQTY))
                    {
                        if (!isEndDateToday && fishQTYAdjustCache.TryGetValue(partNO, out int adjustQTY))
                        {
                            fishQTY += adjustQTY;
                        }
                    }
                    #endregion
                    #endregion
                    dis_tmp = $"<tr><td colspan='12' align='right'>結餘再倉量</td><td>{fishQTY}</td></tr>";

                    DataRow currentRow = null;
                    string showQTY_a = "";
                    string showQTY_b = "";
                    int notOKQTY = 0;
                    int qty = 0;
                    string state = "";
                    for (int i = 1; i < dt.Rows.Count; i++)
                    {
                        currentRow = dt.Rows[i];
                        string drPartNO = currentRow.IsNull("PartNO") ? "" : currentRow["PartNO"].ToString();
                        if (partNO != drPartNO)
                        {
                            tmp = dt.Rows[(i - 1)];
                            string classValue = tmp.IsNull("Class") ? "" : tmp["Class"].ToString();
                            classMap.TryGetValue(classValue, out className);
                            if (string.IsNullOrEmpty(className)) { className = ""; }

                            qty = tmp.IsNull("QTY") ? 0 : int.Parse(tmp["QTY"].ToString());
                            string mathType = tmp.IsNull("MathType") ? "" : tmp["MathType"].ToString();
                            bool isOK = tmp.IsNull("IsOK") ? false : bool.Parse(tmp["IsOK"].ToString());
                            string tmpQTYStr = tmp.IsNull("QTY") ? "0" : tmp["QTY"].ToString();

                            if (mathType == "-")
                            {
                                showQTY_a = "";
                                showQTY_b = tmpQTYStr;
                                if (isOK)
                                { fishQTY += qty; }
                                else { notOKQTY += qty; }
                            }
                            else
                            {
                                showQTY_b = "";
                                showQTY_a = tmpQTYStr;
                                if (isOK)
                                { fishQTY -= qty; }
                                else { notOKQTY -= qty; }
                            }
                            state = isOK ? "入帳" : "再途";

                            string startTime = "";
                            if (!tmp.IsNull("StartTime"))
                            {
                                try { startTime = Convert.ToDateTime(tmp["StartTime"]).ToString("yy-MM-dd HH:mm:ss"); }
                                catch { startTime = ""; }
                            }
                            string docNumberNO = tmp.IsNull("DOCNumberNO") ? "" : tmp["DOCNumberNO"].ToString();
                            string docName = tmp.IsNull("DOCName") ? "" : tmp["DOCName"].ToString();
                            string partNO_prev = tmp.IsNull("PartNO") ? "" : tmp["PartNO"].ToString();
                            string partName = tmp.IsNull("PartName") ? "" : tmp["PartName"].ToString();
                            string specification = tmp.IsNull("Specification") ? "" : tmp["Specification"].ToString();
                            string safeQTY = tmp.IsNull("SafeQTY") ? "" : tmp["SafeQTY"].ToString();
                            string unit = tmp.IsNull("Unit") ? "" : tmp["Unit"].ToString();
                            string storeNO = tmp.IsNull("StoreNO") ? "" : tmp["StoreNO"].ToString();
                            string storeSpacesNO = tmp.IsNull("StoreSpacesNO") ? "" : tmp["StoreSpacesNO"].ToString();

                            sb.Append($"<tr><td>{startTime}</td><td>{docNumberNO}<br />{docName}</td><td>{partNO_prev}</td><td>{partName}{specification}</td><td>{className}</td><td>{safeQTY}</td><td>{unit}</td><td>{state}</td><td>{storeNO}{storeSpacesNO}</td><td>{showQTY_a}</td><td>{showQTY_b}</td><td>{notOKQTY}</td><td>{fishQTY}</td></tr>{dis_tmp}");
                            partNO = currentRow.IsNull("PartNO") ? "" : currentRow["PartNO"].ToString();

                            fishQTY = 0;
                            notOKQTY = 0;
                            #region 計算期末數量
                            // 使用快取的 TotalStock 資料
                            if (!string.IsNullOrEmpty(partNO) && totalStockCache.TryGetValue(partNO, out fishQTY))
                            {
                                if (!isEndDateToday && fishQTYAdjustCache.TryGetValue(partNO, out int adjustQTY))
                                {
                                    fishQTY += adjustQTY;
                                }
                            }
                            #endregion
                            dis_tmp = $"<tr><td colspan='12' align='right'>結餘再倉量</td><td>{fishQTY}</td></tr>";
                        }
                        else
                        {
                            qty = currentRow.IsNull("QTY") ? 0 : int.Parse(currentRow["QTY"].ToString());
                            string mathType = currentRow.IsNull("MathType") ? "" : currentRow["MathType"].ToString();
                            bool isOK = currentRow.IsNull("IsOK") ? false : bool.Parse(currentRow["IsOK"].ToString());
                            string drQTYStr = currentRow.IsNull("QTY") ? "0" : currentRow["QTY"].ToString();

                            if (mathType == "-")
                            {
                                showQTY_a = "";
                                showQTY_b = drQTYStr;
                                if (isOK)
                                { fishQTY += qty; }
                                else { notOKQTY += qty; }
                            }
                            else
                            {
                                showQTY_b = "";
                                showQTY_a = drQTYStr;
                                if (isOK)
                                { fishQTY -= qty; }
                                else { notOKQTY -= qty; }
                            }
                            state = isOK ? "入帳" : "再途";

                            string startTime = "";
                            if (!currentRow.IsNull("StartTime"))
                            {
                                try { startTime = Convert.ToDateTime(currentRow["StartTime"]).ToString("yy-MM-dd HH:mm:ss"); }
                                catch { startTime = ""; }
                            }
                            string docNumberNO = currentRow.IsNull("DOCNumberNO") ? "" : currentRow["DOCNumberNO"].ToString();
                            string docName = currentRow.IsNull("DOCName") ? "" : currentRow["DOCName"].ToString();
                            string remark = currentRow.IsNull("Remark") ? "" : currentRow["Remark"].ToString();
                            string storeNO = currentRow.IsNull("StoreNO") ? "" : currentRow["StoreNO"].ToString();
                            string storeSpacesNO = currentRow.IsNull("StoreSpacesNO") ? "" : currentRow["StoreSpacesNO"].ToString();

                            dis_tmp = $"<tr><td>{startTime}</td><td>{docNumberNO}</td><td>{docName}</td><td colspan='4'>{remark}</td><td>{state}</td><td>{storeNO}{storeSpacesNO}</td><td>{showQTY_a}</td><td>{showQTY_b}</td><td>{notOKQTY}</td><td>{fishQTY}</td></tr>{dis_tmp}";
                        }
                    }
                }
                sb.Append("</table>");
                re = sb.ToString();
            }
            ViewBag.HtmlOutput = re;
            return View(key);
        }


        public IActionResult BOMProductProcess(string filterPartNO, string filterPartName, string filterSpec, string filterClass, string isQuery)//BOM與製程總表
        {
            ViewBag.FilterPartNO = filterPartNO ?? "";
            ViewBag.FilterPartName = filterPartName ?? "";
            ViewBag.FilterSpec = filterSpec ?? "";
            ViewBag.FilterClass = filterClass ?? "45";

            if (string.IsNullOrEmpty(isQuery))
            {
                ViewBag.HtmlOutput = "";
                return View();
            }

            string classFilter = "";
            if (filterClass == "4") { classFilter = " and b.Class='4'"; }
            else if (filterClass == "5") { classFilter = " and b.Class='5'"; }
            else { classFilter = " and (b.Class='4' or b.Class='5')"; }

            string partNOFilter = "";
            if (!string.IsNullOrWhiteSpace(filterPartNO)) { partNOFilter = $" and a.PartNO LIKE '%{filterPartNO.Trim().Replace("'", "''")}%'"; }

            string partNameFilter = "";
            if (!string.IsNullOrWhiteSpace(filterPartName)) { partNameFilter = $" and b.PartName LIKE N'%{filterPartName.Trim().Replace("'", "''")}%'"; }

            string specFilter = "";
            if (!string.IsNullOrWhiteSpace(filterSpec)) { specFilter = $" and b.Specification LIKE N'%{filterSpec.Trim().Replace("'", "''")}%'"; }

            var sb = new StringBuilder();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt_BOM = null;
                #region 獨立製程 IsShare_PP_Name='0'
                // 效能優化：簡化 SQL，只選擇需要的欄位
                dt_BOM = db.DB_GetData($@"select a.Id,a.PartNO,a.Apply_PP_Name,a.Apply_PartNO,a.Apply_StationNO,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName,a.IsConfirm,a.IsChackQTY,a.IsChackIsOK,a.Version,a.EffectiveDate,a.ExpiryDate,b.PartName,b.Specification,b.Class FROM SoftNetMainDB.[dbo].[BOM] as a 
                                    join SoftNetMainDB.[dbo].[Material] as b on b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO 
                                    where a.ServerId='{_Fun.Config.ServerId}' and a.IsShare_PP_Name='0'{classFilter}{partNOFilter}{partNameFilter}{specFilter} order by a.Apply_PP_Name,a.Apply_PartNO,a.EffectiveDate,a.IndexSN desc");
                if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                {
                    // 效能優化：批次載入所有需要的 PP_ProductProcess_Item 和 BOMII 資料
                    var bomIds = new HashSet<string>();
                    foreach (DataRow dr in dt_BOM.Rows)
                    {
                        if (!dr.IsNull("Id")) { bomIds.Add(dr["Id"].ToString()); }
                    }

                    // 批次載入 PP_ProductProcess_Item
                    Dictionary<string, List<DataRow>> ppItemCache = new Dictionary<string, List<DataRow>>();
                    if (bomIds.Count > 0)
                    {
                        string bomIdsStr = string.Join("','", bomIds);
                        DataTable dt_PP_Items = db.DB_GetData($"SELECT BOMId, StationNO FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where BOMId IN ('{bomIdsStr}')");
                        if (dt_PP_Items != null)
                        {
                            foreach (DataRow row in dt_PP_Items.Rows)
                            {
                                if (!row.IsNull("BOMId"))
                                {
                                    string bomId = row["BOMId"].ToString();
                                    if (!ppItemCache.ContainsKey(bomId))
                                    {
                                        ppItemCache[bomId] = new List<DataRow>();
                                    }
                                    ppItemCache[bomId].Add(row);
                                }
                            }
                        }
                    }

                    // 批次載入 BOMII
                    Dictionary<string, List<DataRow>> bomIICache = new Dictionary<string, List<DataRow>>();
                    if (bomIds.Count > 0)
                    {
                        string bomIdsStr = string.Join("','", bomIds);
                        DataTable dt_BOMII_All = db.DB_GetData($@"select a.BOMId,a.PartNO,a.BOMQTY,a.AttritionRate,b.PartName,b.Specification,b.Class FROM SoftNetMainDB.[dbo].[BOMII] as a
                                                join SoftNetMainDB.[dbo].[Material] as b on b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO 
                                                where a.BOMId IN ('{bomIdsStr}') order by a.BOMId, a.sn");
                        if (dt_BOMII_All != null)
                        {
                            foreach (DataRow row in dt_BOMII_All.Rows)
                            {
                                if (!row.IsNull("BOMId"))
                                {
                                    string bomId = row["BOMId"].ToString();
                                    if (!bomIICache.ContainsKey(bomId))
                                    {
                                        bomIICache[bomId] = new List<DataRow>();
                                    }
                                    bomIICache[bomId].Add(row);
                                }
                            }
                        }
                    }

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
                        string currentPP_Name = dr.IsNull("Apply_PP_Name") ? "" : dr["Apply_PP_Name"].ToString();
                        string currentApply_PartNO = dr.IsNull("Apply_PartNO") ? "" : dr["Apply_PartNO"].ToString();
                        if (pp_Name != currentPP_Name || apply_PartNO != currentApply_PartNO)
                        {
                            if (pp_Name != "")
                            {
                                sb.Append("</table>");
                                string currentPartNO = dr.IsNull("PartNO") ? "" : dr["PartNO"].ToString();
                                if (mbomID_PartNO != currentPartNO) { sb.Append("<br />"); }
                            }
                            sb.Append("<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>");
                            pp_Name = currentPP_Name;
                            apply_PartNO = currentApply_PartNO;
                            color = "";
                            bool isConfirm_bool = dr.IsNull("IsConfirm") ? false : bool.Parse(dr["IsConfirm"].ToString());
                            if (isConfirm_bool) { isConfirm = "發行"; }
                            else { isConfirm = "NG"; color = " style='background-color:tomato;'"; }
                            mbomID_PartNO = dr.IsNull("PartNO") ? "" : dr["PartNO"].ToString();
                            string partName = dr.IsNull("PartName") ? "" : dr["PartName"].ToString();
                            string specification = dr.IsNull("Specification") ? "" : dr["Specification"].ToString();
                            string version = dr.IsNull("Version") ? "" : dr["Version"].ToString();
                            string effectiveDate = dr.IsNull("EffectiveDate") ? "" : dr["EffectiveDate"].ToString();
                            string expiryDate = dr.IsNull("ExpiryDate") ? "" : dr["ExpiryDate"].ToString();
                            sb.Append($"<tr><td>階</td><td>生產母件料號:{mbomID_PartNO}&emsp;{partName}&emsp;{specification}</td><td>製程名稱:{pp_Name}</td><td>版本:{version}</td><td>有效期限:{effectiveDate}~{expiryDate}</td><td{color}>{isConfirm}</td></tr>");
                        }

                        bool isChackQTY = dr.IsNull("IsChackQTY") ? false : bool.Parse(dr["IsChackQTY"].ToString());
                        bool isChackIsOK = dr.IsNull("IsChackIsOK") ? false : bool.Parse(dr["IsChackIsOK"].ToString());
                        isQTY = isChackQTY ? "是" : "否";
                        isOK = isChackIsOK ? "是" : "否";
                        color = "";
                        string apply_StationNO = dr.IsNull("Apply_StationNO") ? "" : dr["Apply_StationNO"].ToString();
                        station = $"主站:{apply_StationNO}&emsp;";

                        // 使用快取的 PP_ProductProcess_Item 資料
                        string bomId = dr.IsNull("Id") ? "" : dr["Id"].ToString();
                        if (!string.IsNullOrEmpty(bomId) && ppItemCache.TryGetValue(bomId, out List<DataRow> ppItems))
                        {
                            stationNO_Merge = "";
                            foreach (DataRow drII in ppItems)
                            {
                                string stationNO = drII.IsNull("StationNO") ? "" : drII["StationNO"].ToString();
                                if (stationNO != apply_StationNO)
                                {
                                    if (stationNO_Merge == "") { stationNO_Merge = $"合併站:{stationNO}"; }
                                    else { stationNO_Merge = $"{stationNO_Merge};{stationNO}"; }
                                }
                            }
                            station = $"{station}{stationNO_Merge}";
                        }
                        else
                        {
                            station = $"{station}製程設定異常&emsp;"; color = " style='background-color:tomato;'";
                        }
                        string indexSN = dr.IsNull("IndexSN") ? "" : dr["IndexSN"].ToString();
                        string station_Custom_IndexSN = dr.IsNull("Station_Custom_IndexSN") ? "" : dr["Station_Custom_IndexSN"].ToString();
                        string stationNO_Custom_DisplayName = dr.IsNull("StationNO_Custom_DisplayName") ? "" : dr["StationNO_Custom_DisplayName"].ToString();
                        sb.Append($"<tr><td>{indexSN}</td><td colspan='3'{color}>{station}{station_Custom_IndexSN}&emsp;{stationNO_Custom_DisplayName}</td><td colspan='2'>關站條件 足量:{isQTY} 單據確認:{isOK}</td></tr>");

                        // 使用快取的 BOMII 資料
                        if (!string.IsNullOrEmpty(bomId) && bomIICache.TryGetValue(bomId, out List<DataRow> bomIIItems))
                        {
                            foreach (DataRow drII in bomIIItems)
                            {
                                string partNO_II = drII.IsNull("PartNO") ? "" : drII["PartNO"].ToString();
                                if (partNO_II == mbomID_PartNO) { continue; }
                                string partName_II = drII.IsNull("PartName") ? "" : drII["PartName"].ToString();
                                string specification_II = drII.IsNull("Specification") ? "" : drII["Specification"].ToString();
                                string bomQTY = drII.IsNull("BOMQTY") ? "" : drII["BOMQTY"].ToString();
                                string attritionRate = drII.IsNull("AttritionRate") ? "" : drII["AttritionRate"].ToString();
                                sb.Append($"<tr><td></td><td>用料:{partNO_II}&emsp;{partName_II}&emsp;{specification_II}</td><td>用量:{bomQTY}</td><td>損耗率:{attritionRate}</td></tr>");
                            }
                        }
                    }
                    sb.Append("</table>");
                }
                else
                {
                    sb.Append("<p> 目前無獨立製程BOM紀錄</p>");
                }
                #endregion

                sb.Append("<br />");

                #region 共用製程 IsShare_PP_Name='1'
                // 效能優化：簡化 SQL，只選擇需要的欄位
                dt_BOM = db.DB_GetData($@"select a.Id,a.PartNO,a.Apply_PP_Name,a.Apply_PartNO,a.Apply_StationNO,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName,a.IsConfirm,a.IsChackQTY,a.IsChackIsOK,a.Version,a.EffectiveDate,a.ExpiryDate,b.PartName,b.Specification,b.Class FROM SoftNetMainDB.[dbo].[BOM] as a 
                                    join SoftNetMainDB.[dbo].[Material] as b on b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO 
                                    where a.ServerId='{_Fun.Config.ServerId}' and a.IsShare_PP_Name='1' and Main_Item='1'{classFilter}{partNOFilter}{partNameFilter}{specFilter} order by a.Apply_PP_Name,a.Apply_PartNO,a.EffectiveDate,a.IndexSN desc");
                if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                {
                    // 效能優化：批次載入所有需要的 PP_ProductProcess_Item 和 BOMII 資料
                    var bomIds_Shared = new HashSet<string>();
                    var ppNamePartNOIndexSN = new HashSet<string>();
                    foreach (DataRow dr in dt_BOM.Rows)
                    {
                        if (!dr.IsNull("Id")) { bomIds_Shared.Add(dr["Id"].ToString()); }
                        string ppName = dr.IsNull("Apply_PP_Name") ? "" : dr["Apply_PP_Name"].ToString();
                        string partNO = dr.IsNull("PartNO") ? "" : dr["PartNO"].ToString();
                        string indexSN = dr.IsNull("IndexSN") ? "" : dr["IndexSN"].ToString();
                        if (!string.IsNullOrEmpty(ppName) && !string.IsNullOrEmpty(partNO) && !string.IsNullOrEmpty(indexSN))
                        {
                            ppNamePartNOIndexSN.Add($"{ppName}|{partNO}|{indexSN}");
                        }
                    }

                    // 批次載入 PP_ProductProcess_Item（共用製程使用不同的查詢條件）
                    Dictionary<string, List<DataRow>> ppItemCache_Shared = new Dictionary<string, List<DataRow>>();
                    if (ppNamePartNOIndexSN.Count > 0)
                    {
                        var conditions = new List<string>();
                        foreach (string key in ppNamePartNOIndexSN)
                        {
                            string[] parts = key.Split('|');
                            if (parts.Length == 3)
                            {
                                conditions.Add($"(PP_Name='{parts[0]}' and PartNO='{parts[1]}' and IndexSN={parts[2]})");
                            }
                        }
                        if (conditions.Count > 0)
                        {
                            string whereClause = string.Join(" or ", conditions);
                            DataTable dt_PP_Items_Shared = db.DB_GetData($"SELECT PP_Name, PartNO, IndexSN, StationNO FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and ({whereClause})");
                            if (dt_PP_Items_Shared != null)
                            {
                                foreach (DataRow row in dt_PP_Items_Shared.Rows)
                                {
                                    string ppName = row.IsNull("PP_Name") ? "" : row["PP_Name"].ToString();
                                    string partNO = row.IsNull("PartNO") ? "" : row["PartNO"].ToString();
                                    string indexSN = row.IsNull("IndexSN") ? "" : row["IndexSN"].ToString();
                                    if (!string.IsNullOrEmpty(ppName) && !string.IsNullOrEmpty(partNO) && !string.IsNullOrEmpty(indexSN))
                                    {
                                        string cacheKey = $"{ppName}|{partNO}|{indexSN}";
                                        if (!ppItemCache_Shared.ContainsKey(cacheKey))
                                        {
                                            ppItemCache_Shared[cacheKey] = new List<DataRow>();
                                        }
                                        ppItemCache_Shared[cacheKey].Add(row);
                                    }
                                }
                            }
                        }
                    }

                    // 批次載入 BOMII（共用製程）
                    Dictionary<string, List<DataRow>> bomIICache_Shared = new Dictionary<string, List<DataRow>>();
                    if (bomIds_Shared.Count > 0)
                    {
                        string bomIdsStr = string.Join("','", bomIds_Shared);
                        DataTable dt_BOMII_All_Shared = db.DB_GetData($@"select a.BOMId,a.PartNO,a.BOMQTY,a.AttritionRate,a.Class,b.PartName,b.Specification,b.PartType,b.StoreSTime,b.SafeQTY FROM SoftNetMainDB.[dbo].[BOMII] as a
                                                join SoftNetMainDB.[dbo].[Material] as b on b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO 
                                                where a.BOMId IN ('{bomIdsStr}') order by a.BOMId, a.sn");
                        if (dt_BOMII_All_Shared != null)
                        {
                            foreach (DataRow row in dt_BOMII_All_Shared.Rows)
                            {
                                if (!row.IsNull("BOMId"))
                                {
                                    string bomId = row["BOMId"].ToString();
                                    if (!bomIICache_Shared.ContainsKey(bomId))
                                    {
                                        bomIICache_Shared[bomId] = new List<DataRow>();
                                    }
                                    bomIICache_Shared[bomId].Add(row);
                                }
                            }
                        }
                    }

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
                        #region 取得第1階
                        sb.Append("<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>");
                        pp_Name = dr.IsNull("Apply_PP_Name") ? "" : dr["Apply_PP_Name"].ToString();
                        apply_PartNO = dr.IsNull("Apply_PartNO") ? "" : dr["Apply_PartNO"].ToString();
                        color = "";
                        bool isConfirm_bool = dr.IsNull("IsConfirm") ? false : bool.Parse(dr["IsConfirm"].ToString());
                        if (isConfirm_bool) { isConfirm = "發行"; }
                        else { isConfirm = "NG"; color = " style='background-color:tomato;'"; }
                        mbomID_PartNO = dr.IsNull("PartNO") ? "" : dr["PartNO"].ToString();
                        string partName = dr.IsNull("PartName") ? "" : dr["PartName"].ToString();
                        string specification = dr.IsNull("Specification") ? "" : dr["Specification"].ToString();
                        string version = dr.IsNull("Version") ? "" : dr["Version"].ToString();
                        string effectiveDate = dr.IsNull("EffectiveDate") ? "" : dr["EffectiveDate"].ToString();
                        string expiryDate = dr.IsNull("ExpiryDate") ? "" : dr["ExpiryDate"].ToString();
                        sb.Append($"<tr><td>階</td><td>生產母件料號:{mbomID_PartNO}&emsp;{partName}&emsp;{specification}</td><td>製程名稱:{pp_Name}</td><td>版本:{version}</td><td>有效期限:{effectiveDate}~{expiryDate}</td><td{color}>{isConfirm}</td></tr>");
                        bool isChackQTY = dr.IsNull("IsChackQTY") ? false : bool.Parse(dr["IsChackQTY"].ToString());
                        bool isChackIsOK = dr.IsNull("IsChackIsOK") ? false : bool.Parse(dr["IsChackIsOK"].ToString());
                        isQTY = isChackQTY ? "是" : "否";
                        isOK = isChackIsOK ? "是" : "否";
                        color = "";
                        string apply_StationNO = dr.IsNull("Apply_StationNO") ? "" : dr["Apply_StationNO"].ToString();
                        station = $"主站:{apply_StationNO}&emsp;";

                        // 使用快取的 PP_ProductProcess_Item 資料（共用製程）
                        string indexSN = dr.IsNull("IndexSN") ? "" : dr["IndexSN"].ToString();
                        string cacheKey = $"{pp_Name}|{mbomID_PartNO}|{indexSN}";
                        if (ppItemCache_Shared.TryGetValue(cacheKey, out List<DataRow> ppItems_Shared))
                        {
                            stationNO_Merge = "";
                            foreach (DataRow drII in ppItems_Shared)
                            {
                                string stationNO = drII.IsNull("StationNO") ? "" : drII["StationNO"].ToString();
                                if (stationNO != apply_StationNO)
                                {
                                    if (stationNO_Merge == "") { stationNO_Merge = $"合併站:{stationNO}"; }
                                    else { stationNO_Merge = $"{stationNO_Merge};{stationNO}"; }
                                }
                            }
                            station = $"{station}{stationNO_Merge}";
                        }
                        else
                        {
                            station = $"{station}製程設定異常&emsp;"; color = " style='background-color:tomato;'";
                        }
                        string station_Custom_IndexSN = dr.IsNull("Station_Custom_IndexSN") ? "" : dr["Station_Custom_IndexSN"].ToString();
                        string stationNO_Custom_DisplayName = dr.IsNull("StationNO_Custom_DisplayName") ? "" : dr["StationNO_Custom_DisplayName"].ToString();
                        sb.Append($"<tr><td>{indexSN}</td><td colspan='3'{color}>{station}{station_Custom_IndexSN}&emsp;{stationNO_Custom_DisplayName}</td><td colspan='2'>關站條件 足量:{isQTY} 單據確認:{isOK}</td></tr>");

                        // 使用快取的 BOMII 資料（共用製程）
                        string bomId = dr.IsNull("Id") ? "" : dr["Id"].ToString();
                        if (!string.IsNullOrEmpty(bomId) && bomIICache_Shared.TryGetValue(bomId, out List<DataRow> bomIIItems_Shared))
                        {
                            foreach (DataRow drII in bomIIItems_Shared)
                            {
                                string partNO_II = drII.IsNull("PartNO") ? "" : drII["PartNO"].ToString();
                                if (partNO_II == mbomID_PartNO) { continue; }
                                string partName_II = drII.IsNull("PartName") ? "" : drII["PartName"].ToString();
                                string specification_II = drII.IsNull("Specification") ? "" : drII["Specification"].ToString();
                                string bomQTY = drII.IsNull("BOMQTY") ? "" : drII["BOMQTY"].ToString();
                                string attritionRate = drII.IsNull("AttritionRate") ? "" : drII["AttritionRate"].ToString();
                                sb.Append($"<tr><td></td><td>用料:{partNO_II}&emsp;{partName_II}&emsp;{specification_II}</td><td>用量:{bomQTY}</td><td>損耗率:{attritionRate}</td></tr>");
                            }
                        }
                        #endregion

                        // 效能優化：使用已快取的 BOMII 資料，避免重複查詢
                        if (!string.IsNullOrEmpty(bomId) && bomIICache_Shared.TryGetValue(bomId, out List<DataRow> bomIIItems_ForRecursive))
                        {
                            string re = sb.ToString();
                            foreach (DataRow dr0 in bomIIItems_ForRecursive)
                            {
                                string classValue = dr0.IsNull("Class") ? "" : dr0["Class"].ToString();
                                if (classValue == "4" || classValue == "5")
                                {
                                    int indexSN_int = dr.IsNull("IndexSN") ? 0 : int.Parse(dr["IndexSN"].ToString());
                                    RecursiveBOMII(db, pp_Name, apply_PartNO, dr0, indexSN_int, ref re);
                                }
                            }
                            sb.Clear();
                            sb.Append(re);
                        }
                        sb.Append("</table>");
                    }
                }
                else
                {
                    sb.Append("<p> 目前無共用製程BOM紀錄</p>");
                }
                #endregion
            }
            ViewBag.HtmlOutput = sb.ToString();
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
            
            // 效能優化：批次載入所有需要的資料，避免 N+1 查詢問題
            var stationNOs = new HashSet<string>();
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (!dr.IsNull("StationNO")) { stationNOs.Add(dr["StationNO"].ToString()); }
                }
                
                // 批次載入 SFC_StationDetail_ChangeLOG 資料（使用子查詢取得每個工站最近7筆的平均值）
                Dictionary<string, DataRow> stationEfficientCache = new Dictionary<string, DataRow>();
                if (stationNOs.Count > 0)
                {
                    string stationNOsStr = string.Join("','", stationNOs);
                    // 使用子查詢取得每個工站最近7筆記錄的平均值
                    DataTable dt_Efficient = db.DB_GetData($@"SELECT StationNO, 
                        AVG(CT) as CT, AVG(ECT) as ECT, AVG(LCT) as LCT, AVG(UCT) as UCT
                        FROM (
                            SELECT StationNO, CycleTime+WaitTime as CT, ECT, LowerCT as LCT, UpperCT as UCT,
                            ROW_NUMBER() OVER (PARTITION BY StationNO ORDER BY LOGDateTimeID DESC) as rn
                            FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG]
                            where ServerId='{_Fun.Config.ServerId}' 
                            and StationNO IN ('{stationNOsStr}') 
                            and CycleTime!=0
                        ) as sub
                        WHERE rn <= 7
                        GROUP BY StationNO");
                    if (dt_Efficient != null)
                    {
                        foreach (DataRow row in dt_Efficient.Rows)
                        {
                            if (!row.IsNull("StationNO"))
                            {
                                stationEfficientCache[row["StationNO"].ToString()] = row;
                            }
                        }
                    }
                }
                
                // 批次載入 Manufacture 資料
                Dictionary<string, DataRow> manufactureCache = new Dictionary<string, DataRow>();
                if (stationNOs.Count > 0)
                {
                    string stationNOsStr = string.Join("','", stationNOs);
                    DataTable dt_Manufacture = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO IN ('{stationNOsStr}') and (State='1' or State='2')");
                    if (dt_Manufacture != null)
                    {
                        foreach (DataRow row in dt_Manufacture.Rows)
                        {
                            if (!row.IsNull("StationNO"))
                            {
                                string stationNO = row["StationNO"].ToString();
                                // 只保留最新的或符合條件的記錄（如果有多筆，保留第一筆）
                                if (!manufactureCache.ContainsKey(stationNO))
                                {
                                    manufactureCache[stationNO] = row;
                                }
                            }
                        }
                    }
                }
                
                // 批次載入 SFC_StationDetail 資料（需要根據 Manufacture 的資料來查詢）
                Dictionary<string, DataRow> stationDetailCache = new Dictionary<string, DataRow>();
                if (manufactureCache.Count > 0)
                {
                    var detailKeys = new List<string>();
                    foreach (var kvp in manufactureCache)
                    {
                        DataRow mRow = kvp.Value;
                        string key = $"{kvp.Key}|{mRow["PP_Name"].ToString()}|{mRow["OrderNO"].ToString()}|{mRow["IndexSN"].ToString()}";
                        detailKeys.Add($"StationNO='{kvp.Key}' and PP_Name='{mRow["PP_Name"].ToString()}' and OrderNO='{mRow["OrderNO"].ToString()}' and IndexSN={mRow["IndexSN"].ToString()}");
                    }
                    
                    if (detailKeys.Count > 0)
                    {
                        string whereClause = string.Join(" or ", detailKeys);
                        DataTable dt_StationDetail = db.DB_GetData($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail] where ServerId='{_Fun.Config.ServerId}' and ({whereClause})");
                        if (dt_StationDetail != null)
                        {
                            foreach (DataRow row in dt_StationDetail.Rows)
                            {
                                if (!row.IsNull("StationNO"))
                                {
                                    string key = $"{row["StationNO"].ToString()}|{row["PP_Name"].ToString()}|{row["OrderNO"].ToString()}|{row["IndexSN"].ToString()}";
                                    stationDetailCache[key] = row;
                                }
                            }
                        }
                    }
                }
                
                // 處理資料列
                foreach (DataRow dr in dt.Rows)
                {
                    value = 0.01f;
                    string stationNO = dr["StationNO"].ToString();
                    
                    #region 計算值
                    // 使用快取的 SFC_StationDetail_ChangeLOG 資料
                    if (stationEfficientCache.TryGetValue(stationNO, out DataRow efficientRow))
                    {
                        if (!efficientRow.IsNull("CT") && efficientRow["CT"].ToString() != "" && 
                            !efficientRow.IsNull("ECT") && efficientRow["ECT"].ToString() != "" &&
                            int.TryParse(efficientRow["CT"].ToString(), out int ctInt) && ctInt > 0 &&
                            int.TryParse(efficientRow["ECT"].ToString(), out int ectInt) && ectInt > 0)
                        {
                            ct = ctInt;
                            if (!efficientRow.IsNull("LCT") && efficientRow["LCT"].ToString() != "" && 
                                int.TryParse(efficientRow["LCT"].ToString(), out int lctInt) && lctInt >= ct)
                            {
                                value = 100;
                            }
                            else if (!efficientRow.IsNull("UCT") && efficientRow["UCT"].ToString() != "" && 
                                     int.TryParse(efficientRow["UCT"].ToString(), out int uctInt) && uctInt >= ct)
                            {
                                value = 0.01f;
                            }
                            else if (!efficientRow.IsNull("ECT") && !efficientRow.IsNull("LCT") && 
                                     efficientRow["ECT"].ToString() != "" && efficientRow["LCT"].ToString() != "" &&
                                     int.TryParse(efficientRow["ECT"].ToString(), out int ectVal) && ectVal >= ct)
                            {
                                lct = int.TryParse(efficientRow["LCT"].ToString(), out int lctVal) && lctVal > 0 ? lctVal : ectVal;
                                value = ((ct - lct) / ct) * 100;
                                value = (100 - value) * (60 / 100) + 40;
                            }
                            else if (!efficientRow.IsNull("ECT") && efficientRow["ECT"].ToString() != "" &&
                                     int.TryParse(efficientRow["ECT"].ToString(), out int ectValue) && ct > ectValue)
                            {
                                ect = ectValue;
                                value = ((ct - ect) / ct) * 100;
                                value = (100 - value) * 0.4f;
                            }
                            else { value = 0.01f; }
                        }
                    }
                    if (value <= 0) { value = 0.01f; }
                    if (value > 100) { value = 100; }
                    #endregion

                    if (stNO == "") { stNO = $"{stationNO},{value.ToString("0.00")}"; }
                    else { stNO = $"{stNO};{stationNO},{value.ToString("0.00")}"; }
                    if (stName == "") { stName = $"{dr["StationName"].ToString()}"; }
                    else { stName = $"{stName};{dr["StationName"].ToString()}"; }

                    string sState = "閒置";
                    string okQTY = "0";
                    string failQTY = "0";
                    if (dr["Station_Type"].ToString() == "1")
                    {
                        // 使用快取的 Manufacture 資料
                        if (manufactureCache.TryGetValue(stationNO, out DataRow manufactureRow))
                        {
                            if (manufactureRow["State"].ToString() == "1") { sState = "開工"; }
                            else { sState = "停工"; }
                            
                            // 使用快取的 SFC_StationDetail 資料
                            string detailKey = $"{stationNO}|{manufactureRow["PP_Name"].ToString()}|{manufactureRow["OrderNO"].ToString()}|{manufactureRow["IndexSN"].ToString()}";
                            if (stationDetailCache.TryGetValue(detailKey, out DataRow detailRow))
                            {
                                okQTY = detailRow["ProductFinishedQty"].ToString();
                                failQTY = detailRow["ProductFailedQty"].ToString();
                            }
                        }
                    }
                    else
                    {
                        //多工
                    }

                    #region display StationNO
                    if (i == 1) { re = $"{re}<div class='row'>"; }
                    re = $"{re}<div id='{stationNO}' class='gaugeSVGVeiw'></div>";
                    if (i >= 4) { re = $"{re}</div>"; i = 1; } else { ++i; }
                    #endregion
                }
            }
            re = $"{re}</div>";
            data[0] = stNO;
            data[1] = re;
            data[2] = stName;
            return data;
        }

    }
}
