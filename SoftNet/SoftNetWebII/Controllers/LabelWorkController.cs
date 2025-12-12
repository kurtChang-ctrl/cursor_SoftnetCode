using Base;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing;

using SoftNetWebII.Models;
using SoftNetWebII.Services;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Net.Http;
using System.Net.WebSockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

using static Microsoft.EntityFrameworkCore.DbLoggerCategory;

namespace SoftNetWebII.Controllers
{
    public class LabelWorkController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public LabelWorkController(SNWebSocketService websocket, SFC_Common sfc_Common)
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
        public IActionResult Index(string id)//網址參數 0=工站編號;1=Type(0=無1=工單2=計劃碼);2=值(工單或計劃碼);3=IndexSN;4=網頁來的
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            LabelWork tmp = new LabelWork();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string sql = "";
                tmp.OPNO = br.UserNO;
                tmp.OPNO_Name = br.UserName;
                if (id != null && id != "")
                {
                    string[] data = id.Split(';');
                    if (data.Length >= 4)
                    {
                        string displayType = "";
                        if (data.Length == 5) { tmp.BOOLonToggleMenu = data[4]; displayType = $";{data[4]}"; }
                        tmp.Station = data[0];
                        tmp.Type = data[1];
                        DataRow dr_M = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{tmp.Station}'");
                        if (dr_M == null)
                        { ViewBag.ERRMsg = $"查無 {tmp.Station}工站編號, 重新刷條碼 並 重新操作."; return View("ResuItTimeOUT", tmp); }
                        if (tmp.Type == "0" && dr_M["OrderNO"].ToString() != "")
                        {
                            if (dr_M["SimulationId"].ToString() != "") { return RedirectToAction("Index", "LabelWork", new { id = $"{tmp.Station};2;{dr_M["SimulationId"].ToString()};{dr_M["IndexSN"].ToString()}{displayType}" }); }
                            else { return RedirectToAction("Index", "LabelWork", new { id = $"{tmp.Station};1;{dr_M["OrderNO"].ToString()};{dr_M["IndexSN"].ToString()}{displayType}" }); }
                        }
                        tmp.TypeValue = data[2];
                        if (data[3].Trim() == "")
                        { tmp.IndexSN = "0"; }
                        else
                        { tmp.IndexSN = data[3]; }

                        #region 回傳可用工單, 有註冊的OP, 設定 Station_Config_Store_Type 網頁行為
                        DataRow dr2 = null;
                        DataRow dr3 = null;
                        dr2 = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{tmp.Station}'");
                        tmp.Station_Name = dr2["StationName"].ToString();
                        tmp.Has_Knives = bool.Parse(dr2["IsKnives"].ToString()) ? "1" : "0";
                        //###???子製程有問題
                        //sql = @$"select b.OrderNO,b.EstimatedStartTime,b.FactoryName,b.LineName,a.PP_Name,b.PartNO,b.PartName,a.DisplaySN,a.IndexSN,a.Station_Custom_IndexSN,a.IndexSN_Merge from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] as a
                        //                join SoftNetSYSDB.[dbo].[PP_WorkOrder] as b on a.PP_Name=b.PP_Name
                        //                where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO = '{tmp.Station}' and b.EndTime is NULL and b.EstimatedStartTime<='{DateTime.Now.AddMonths(2).ToString("MM/dd/yyyy HH:mm:ss.fff")}' order by b.EstimatedStartTime";

                        DateTime comp_CalendarDate = DateTime.Now.AddYears(10);//###??? 提早顯示的時間要參數化
                        sql = $@"select a.DOCNumberNO,a.NeedId from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] as a 
                                join SoftNetSYSDB.[dbo].[PP_WorkOrder] as b on a.DOCNumberNO=b.OrderNO and b.EndTime is NULL 
                                join SoftNetSYSDB.[dbo].[APS_NeedData] as c on c.Id=a.NeedId and c.State='6' and c.ServerId='{_Fun.Config.ServerId}'
                                where (a.StationNO = '{tmp.Station}' or a.StationNO_Merge like '%{tmp.Station},%') and a.CalendarDate<='{comp_CalendarDate.ToString("MM/dd/yyyy HH:mm:ss.fff")}' and a.DOCNumberNO!='' 
                                and (a.Time1_C!=0 or a.Time2_C!=0 or a.Time3_C!=0 or a.Time4_C!=0) group by a.DOCNumberNO,a.NeedId";
                        DataTable dt_PP = db.DB_GetData(sql);
                        if (dt_PP != null && dt_PP.Rows.Count > 0)
                        {
                            string inkey = "";
                            foreach (DataRow dr0 in dt_PP.Rows)
                            {
                                if (inkey == "") { inkey = $"'{dr0["DOCNumberNO"].ToString()}'"; } else { inkey = $"{inkey},'{dr0["DOCNumberNO"].ToString()}'"; }
                            }
                            Dictionary<string, List<string>> orderOrder = new Dictionary<string, List<string>>();
                            dt_PP = db.DB_GetData($"select DOCNumberNO,SimulationId from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where DOCNumberNO in ({inkey}) and (StationNO = '{tmp.Station}' or StationNO_Merge like '%{tmp.Station},%') and CalendarDate<='{comp_CalendarDate.ToString("MM/dd/yyyy HH:mm:ss.fff")}' group by DOCNumberNO,SimulationId order by DOCNumberNO");
                            if (dt_PP != null && dt_PP.Rows.Count > 0)
                            {
                                foreach (DataRow dr0 in dt_PP.Rows)
                                {
                                    if (orderOrder.ContainsKey(dr0["DOCNumberNO"].ToString()))
                                    {
                                        if (!orderOrder[dr0["DOCNumberNO"].ToString()].Contains(dr0["SimulationId"].ToString())) { orderOrder[dr0["DOCNumberNO"].ToString()].Add(dr0["SimulationId"].ToString()); }
                                    }
                                    else { orderOrder.Add(dr0["DOCNumberNO"].ToString(), new List<string> { dr0["SimulationId"].ToString() }); }
                                }
                            }
                            string sType2 = "";
                            string tmp_Index = "";
                            tmp.HasWO_List = new List<string[]>();
                            Dictionary<string, List<string[]>> groupsItemList = new Dictionary<string, List<string[]>>();
                            foreach (KeyValuePair<string, List<string>> kvp in orderOrder)
                            {
                                foreach (string sId in kvp.Value)
                                {
                                    if (sId == "")
                                    {
                                        //###???直接輸入工單的資料
                                    }
                                    else
                                    {
                                        dr2 = db.DB_GetFirstDataByDataRow($"select a.*,(a.NeedQTY+a.SafeQTY) as QTY,b.NeedType,b.CTName from SoftNetSYSDB.[dbo].[APS_Simulation] as a,SoftNetSYSDB.[dbo].[APS_NeedData] as b where a.NeedId=b.Id and a.SimulationId='{sId}'");
                                        if (dr2 != null)
                                        {
                                            if (dr2["NeedType"].ToString() == "1") { sType2 = $"訂單"; }
                                            else if (dr2["NeedType"].ToString() == "2") { sType2 = $"客戶"; }
                                            else if (dr2["NeedType"].ToString() == "5") { sType2 = $"底稿"; }
                                            else { sType2 = $"廠內"; }
                                            dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr2["PartNO"].ToString()}'");
                                            if (db.DB_GetQueryCount($"SELECT StationNO FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{tmp.Station}' and OrderNO='{dr2["DOCNumberNO"].ToString()}' and IndexSN={dr2["Source_StationNO_IndexSN"].ToString()}") > 0) { }
                                            else
                                            {
                                                if (dr2["Source_StationNO_Custom_DisplayName"].ToString() != "") { tmp_Index = dr2["Source_StationNO_Custom_DisplayName"].ToString(); }
                                                else if (dr2["Source_StationNO_Custom_IndexSN"].ToString() != "") { tmp_Index = dr2["Source_StationNO_Custom_IndexSN"].ToString(); }
                                                else { tmp_Index = dr2["Source_StationNO_IndexSN"].ToString(); }
                                                if (!groupsItemList.ContainsKey(dr2["PartNO"].ToString())) 
                                                {
                                                    List<string[]> tmp_l = new List<string[]>();
                                                    tmp_l.Add(new string[] { dr2["DOCNumberNO"].ToString(), dr2["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString(), sType2, dr2["CTName"].ToString(), dr2["Apply_PP_Name"].ToString(), tmp_Index, sId, dr2["Source_StationNO_Custom_IndexSN"].ToString(), dr2["Source_StationNO_Custom_DisplayName"].ToString(), dr2["QTY"].ToString(), dr2["Source_StationNO_IndexSN"].ToString() });
                                                    groupsItemList.Add(dr2["PartNO"].ToString(), tmp_l); 
                                                }
                                                else { groupsItemList[dr2["PartNO"].ToString()].Add(new string[] { dr2["DOCNumberNO"].ToString(), dr2["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString(), sType2, dr2["CTName"].ToString(), dr2["Apply_PP_Name"].ToString(), tmp_Index, sId, dr2["Source_StationNO_Custom_IndexSN"].ToString(), dr2["Source_StationNO_Custom_DisplayName"].ToString(), dr2["QTY"].ToString(), dr2["Source_StationNO_IndexSN"].ToString() }); }
                                                
                                            }
                                        }
                                    }
                                }
                            }
                            if (groupsItemList.Count > 0)
                            {
                                groupsItemList = groupsItemList.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value);
                                foreach (KeyValuePair<string, List<string[]>> kv in groupsItemList)
                                {
                                    foreach(string[] ss in kv.Value)
                                    {
                                        tmp.HasWO_List.Add(ss);
                                    }
                                }
                            }
                        }
                        if (dr_M != null)
                        {
                            switch (dr_M["State"].ToString())
                            {
                                case "1": tmp.Station_State = "生產中..."; break;
                                case "4": tmp.Station_State = "等待工單."; break;
                                default: tmp.Station_State = "停止..."; break;
                            }

                            tmp.StationNO_Custom_DisplayName = dr_M["StationNO_Custom_DisplayName"].ToString();
                            tmp.Station_Config_Store_Type = dr_M["Config_Store_Type"].ToString();

                            #region 檢查下一站是否為委外, 設定 Station_Config_Store_Type 網頁行為
                            if (!_Fun.Config.IsOutPackStationStore && dr_M["SimulationId"].ToString() != "")
                            {
                                dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_M["SimulationId"].ToString()}'");
                                dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr3["NeedId"].ToString()}' and PartNO='{dr3["Master_PartNO"].ToString()}' and Source_StationNO='{dr3["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN={dr3["IndexSN"].ToString()} and (Class='4' or Class='5') and Source_StationNO is not null");
                                if (dr3 != null)
                                {
                                    if (_Fun.Config.OutPackStationName == dr3["Apply_StationNO"].ToString()) { tmp.Station_Config_Store_Type = "0"; }
                                }
                            }
                            #endregion

                            #region 計算前站完成量
                            dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_M["SimulationId"].ToString()}'");
                            if (dr3 != null)
                            {
                                dr3 = db.DB_GetFirstDataByDataRow($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation] where Source_StationNO is not NULL and NeedId='{dr3["NeedId"].ToString()}' and Apply_PP_Name='{dr3["Apply_PP_Name"].ToString()}' and IndexSN={(int.Parse(dr3["IndexSN"].ToString()) - 1).ToString()} and Apply_StationNO='{dr3["Source_StationNO"].ToString()}'");
                                if (dr3 != null)
                                {
                                    dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].APS_PartNOTimeNote where SimulationId='{dr3["SimulationId"].ToString()}'");
                                    if (dr3 != null)
                                    {
                                        tmp.Transfer_QTY = dr3["Detail_QTY"].ToString();
                                    }
                                }
                            }
                            #endregion

                            #region 顯示資訊
                            DataRow tmp_dr2 = null;
                            int okQTY = 0;//本站
                            int failedQty = 0;//本站
                            int t_okQTY = 0;//其他站
                            int t_failedQty = 0;//其他站
                            int totCT = 0;//本站

                            DataTable tmp_dt = null;
                            if (dr_M["OrderNO"].ToString() != "")
                            {
                                tmp_dt = db.DB_GetData($@"SELECT sum(ProductFinishedQty) as OKQTY,sum(ProductFailedQty) as FailedQty,sum(CycleTime)*(sum(ProductFinishedQty)+sum(ProductFailedQty)) as TOTCT FROM SoftNetLogDB.[dbo].[SFC_StationDetail] where 
                                    ServerId='{_Fun.Config.ServerId}' and StationNO='{tmp.Station}' and PP_Name='{dr_M["PP_Name"].ToString()}' and OrderNO='{dr_M["OrderNO"].ToString()}' and IndexSN={dr_M["IndexSN"].ToString()} group by OP_NO");
                                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                {
                                    foreach (DataRow dr4 in tmp_dt.Rows)
                                    {
                                        okQTY += dr4.IsNull("OKQTY") ? 0 : int.Parse(dr4["OKQTY"].ToString());
                                        failedQty += dr4.IsNull("FailedQty") ? 0 : int.Parse(dr4["FailedQty"].ToString());
                                        totCT += dr4.IsNull("TOTCT") ? 0 : int.Parse(dr4["TOTCT"].ToString());
                                    }
                                    if (totCT != 0) { totCT = totCT / (okQTY + failedQty); }
                                }
                                tmp.WO_CT = totCT.ToString();

                                #region 累計合併站
                                tmp_dr2 = db.DB_GetFirstDataByDataRow($"select sum(ProductFinishedQty) as TOKQTY,sum(ProductFailedQty) as TFailedQty from SoftNetLogDB.[dbo].[SFC_StationDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{dr_M["StationNO"].ToString()}' and SimulationId='{dr_M["SimulationId"].ToString()}' and OrderNO='{dr_M["OrderNO"].ToString()}' and IndexSN_Merge='1' and PP_Name='{dr_M["PP_Name"].ToString()}' and IndexSN={dr_M["IndexSN"].ToString()}");
                                if (tmp_dr2 != null)
                                {
                                    t_okQTY = tmp_dr2.IsNull("TOKQTY") ? 0 : int.Parse(tmp_dr2["TOKQTY"].ToString());
                                    t_failedQty = tmp_dr2.IsNull("TFailedQty") ? 0 : int.Parse(tmp_dr2["TFailedQty"].ToString());
                                }
                                #endregion
                                tmp.TOT_OK_QTY = $"{okQTY.ToString()}+{t_okQTY.ToString()}";
                                tmp.TOT_Fail_QTY = $"{failedQty.ToString()}+{t_failedQty.ToString()}";

                                tmp_dr2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_M["PartNO"].ToString()}'");
                                if (tmp_dr2 != null)
                                {
                                    tmp.PartNO_INFO = $"{tmp_dr2["PartNO"].ToString()} [{tmp_dr2["PartName"].ToString()} {tmp_dr2["Specification"].ToString()}]";
                                }
                                tmp.RemarkTimeS = dr_M.IsNull("RemarkTimeS") ? "" : Convert.ToDateTime(dr_M["RemarkTimeS"]).ToString("yyyy-MM-dd HH:mm:ss");
                                tmp.RemarkTimeE = dr_M.IsNull("RemarkTimeE") ? "" : Convert.ToDateTime(dr_M["RemarkTimeE"]).ToString("yyyy-MM-dd HH:mm:ss");


                                tmp_dr2 = db.DB_GetFirstDataByDataRow($"select AVG(EfficientCycleTime) as ECT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_M["PartNO"].ToString()}' and StationNO='{dr_M["StationNO"].ToString()}' and PP_Name='{dr_M["PP_Name"].ToString()}' and IndexSN={dr_M["IndexSN"].ToString()} and DOCNO=''");
                                if (tmp_dr2 != null && !tmp_dr2.IsNull("ECT"))
                                {
                                    tmp.E_CT = tmp_dr2["ECT"].ToString();
                                }
                                if (dr_M["SimulationId"].ToString() != "")
                                {
                                    tmp_dr2 = db.DB_GetFirstDataByDataRow($"select NeedQTY as PNQTY FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr_M["SimulationId"].ToString()}'");
                                    if (tmp_dr2 != null && !tmp_dr2.IsNull("PNQTY"))
                                    { tmp.WO_QTY = tmp_dr2["PNQTY"].ToString(); }
                                }
                            }
                            #endregion

                            if (dr_M["OP_NO"].ToString().Trim() != "")
                            {
                                tmp.HasPO_List = new List<string[]>();
                                string[] s0 = dr_M["OP_NO"].ToString().Trim().Split(';');
                                foreach (string s in s0)
                                {
                                    dr2 = db.DB_GetFirstDataByDataRow($"select Name from SoftNetMainDB.[dbo].[User] where UserNO='{s}'");
                                    if (dr2 != null)
                                    {
                                        tmp.HasPO_List.Add(new string[] { s, dr2["Name"].ToString() });
                                    }
                                    else { tmp.HasPO_List.Add(new string[] { s, "" }); }
                                }
                            }

                        }
                        #endregion

                        switch (data[1])
                        {
                            case "0"://工站無資料
                                tmp.OrderNO = "";
                                tmp.SimulationId = "";
                                break;
                            case "1"://自訂工單
                                tmp.OrderNO = data[2];
                                tmp.SimulationId = "";
                                break;
                            case "2"://計畫工單
                                tmp.OrderNO = "";
                                tmp.SimulationId = data[2];
                                DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{data[2]}'");
                                if (dr_APS_Simulation != null && dr_APS_Simulation["DOCNumberNO"].ToString().Trim() != "'")
                                { tmp.OrderNO = dr_APS_Simulation["DOCNumberNO"].ToString().Trim(); }

                                break;
                        }
                    }
                }
            }
            return View(tmp);
        }

        [HttpPost]
        public ActionResult SetAction(LabelWork keys)//async Task<ActionResult>
        {
            var br = _Fun.GetBaseUser();
            if (keys == null || br == null || !br.IsLogin || br.UserNO.Trim() == "" || keys.Station == "") 
            { ViewBag.ERRMsg = $"作業失敗, 畫面已逾時, 請關閉網頁瀏覽器, 重新刷條碼 並 重新操作."; return View("ResuItTimeOUT",keys); }

            string key_OPNO = br.UserNO;
            string err = "";
            string sql = "";
            keys.ERRMsg = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                try
                {
                    DataRow dr_StationNO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                    if (dr_StationNO == null) { keys.ERRMsg = $"查無 {keys.Station} 工站資料, 請聯繫系統管理者."; goto break_FUN; }
                    switch (keys.State)
                    {
                        case "異動工單":
                            {
                                string stationNO_Custom_IndexSN = "";
                                string stationNO_Custom_DisplayName = "";
                                DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");

                                #region 檢查
                                if (dr["State"].ToString() == "1")
                                { keys.ERRMsg = $"本工站 {keys.Station} 還在運作中,請先停止才能設定工單."; goto break_FUN; }
                                string beforePartNO = dr["PartNO"].ToString();//紀錄上一個工站料號 (非工單料號)
                                string beforePP_Name = dr["PP_Name"].ToString();
                                string beforeIndexSN = dr["IndexSN"].ToString();
                                DataRow sfcdr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{keys.OrderNO}'");
                                if (sfcdr == null) { keys.ERRMsg = $"查無 {keys.OrderNO} 工單, 請聯繫系統管理者."; goto break_FUN; }

                                keys.IndexSN = keys.IndexSN.Trim();
                                if (keys.IndexSN == "0" || keys.IndexSN == "")
                                {
                                    DataTable tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and NeedId='{sfcdr["NeedId"].ToString()}' and (Source_StationNO='{keys.Station}' or StationNO_Merge like '%{keys.Station}%')");
                                    if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                    {
                                        if (tmp_dt.Rows.Count > 1)
                                        { keys.ERRMsg = $"{keys.OrderNO} 工單在此站有多重作業,故需指定何製程序號, 請聯繫系統管理者."; goto break_FUN; }
                                        stationNO_Custom_IndexSN = tmp_dt.Rows[0]["Source_StationNO_Custom_IndexSN"].ToString();
                                        stationNO_Custom_DisplayName = tmp_dt.Rows[0]["Source_StationNO_Custom_DisplayName"].ToString();
                                        keys.IndexSN= tmp_dt.Rows[0]["IndexSN"].ToString();
                                    }
                                    else
                                    { keys.ERRMsg = $"查無相關製程序號, 請聯繫系統管理者."; goto break_FUN; }
                                }
                                else
                                {
                                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and NeedId='{sfcdr["NeedId"].ToString()}' and (Source_StationNO='{keys.Station}' or StationNO_Merge like '%{keys.Station}%') and Source_StationNO_IndexSN={keys.IndexSN}");
                                    if (tmp_dr == null)
                                    { keys.ERRMsg = $"查無相關製程資料, 請聯繫系統管理者."; goto break_FUN; }
                                    else
                                    {
                                        stationNO_Custom_IndexSN = tmp_dr["Source_StationNO_Custom_IndexSN"].ToString();
                                        stationNO_Custom_DisplayName = tmp_dr["Source_StationNO_Custom_DisplayName"].ToString();
                                    }
                                }
                                #endregion

                                #region 檢查異動之前,前工單是否無報工
                                if (beforePartNO != "")
                                {
                                    int avgReportTime = 0;
                                    DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and PartNO='{beforePartNO}' and PP_Name='{beforePP_Name}' and IndexSN='{beforeIndexSN}' and ReportTime>10");
                                    if (dr_tmp != null && !dr_tmp.IsNull("AVGTime") && dr_tmp["AVGTime"].ToString().Trim() != "")
                                    {
                                        avgReportTime = int.Parse(dr_tmp["AVGTime"].ToString());
                                    }
                                    if (avgReportTime > 10)
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and PartNO='{beforePartNO}' and IndexSN={beforeIndexSN} and LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%開工%' order by LOGDateTime desc");
                                        if (dr_tmp != null)
                                        {
                                            DateTime tmp_edate = Convert.ToDateTime(dr_tmp["LOGDateTime"]);
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and PartNO='{beforePartNO}' and IndexSN={beforeIndexSN} and LOGDateTime>='{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%報工%'");
                                            if (dr_tmp == null)
                                            {
                                                int isARGs10_offset = 15;//###??? 10將來改參數
                                                if ((_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(),tmp_edate.AddMinutes(isARGs10_offset), DateTime.Now)) >= avgReportTime)
                                                {
                                                    keys.ERRMsg = $"疑似 前一次料號:{beforePartNO} 未完成報工. 若未報工,請至系統網站補報工.";
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion

                                string partNO = "";
                                string partName = "";
                                string typevalue = $"0;";
                                int ct = 0;
                                int num = 0;

                                #region 查有無需求碼
                                if (keys.SimulationId == "")
                                {
                                    typevalue = $"1;{keys.OrderNO}";
                                    //###???獎來要處裡自訂工單,暫時報錯誤
                                    keys.ERRMsg = $"程式功能異常, 請聯繫系統管理者."; goto break_FUN;
                                }
                                else
                                {
                                    typevalue = $"2;{keys.SimulationId}";
                                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{sfcdr["NeedId"].ToString()}' and SimulationId='{keys.SimulationId}'");
                                    if (tmp_dr != null)
                                    {
                                        num = int.Parse(tmp_dr["NeedQTY"].ToString()) + int.Parse(tmp_dr["SafeQTY"].ToString());
                                        ct = int.Parse(tmp_dr["Math_EfficientCT"].ToString());
                                        partNO = tmp_dr["PartNO"].ToString();
                                        partName = tmp_dr["Apply_PP_Name"].ToString();
                                        tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{tmp_dr["PartNO"].ToString()}'");
                                        if (tmp_dr != null)
                                        {
                                            partName = tmp_dr["PartName"].ToString().Replace("\"", "＂").Replace("'", "’"); 
                                        }
                                        tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where APS_StationNO='{keys.OrderNO}' AND SimulationId='{keys.SimulationId}'");
                                        if (tmp_dr != null)
                                        {
                                            num = int.Parse(tmp_dr["NeedQTY"].ToString());
                                        }
                                    }
                                }
                                #endregion

                                #region 更新Manufacture 製造現場狀態
                                if (!db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Station_Custom_IndexSN='{stationNO_Custom_IndexSN}',StationNO_Custom_DisplayName='{stationNO_Custom_DisplayName}',StartTime=NULL,RemarkTimeS=NULL,RemarkTimeE=NULL,EndTime=NULL,Label_ProjectType='0',OrderNO='{keys.OrderNO}',IndexSN={keys.IndexSN},OP_NO='{keys.OPNO}',Master_PP_Name='{sfcdr["PP_Name"].ToString()}',PP_Name='{sfcdr["PP_Name"].ToString()}',SimulationId='{keys.SimulationId}',PartNO='{partNO}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'"))
                                {
                                    keys.HasWO_List = new List<string[]>();
                                    keys.HasWO_List.Add(new string[] { keys.OPNO, keys.OPNO_Name });
                                    keys.ERRMsg = $"異動工單設定失敗."; goto break_FUN;
                                }
                                else { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{keys.OrderNO}' where SimulationId='{keys.SimulationId}'"); }
                                #endregion

                                db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{sfcdr["NeedId"].ToString()}','{keys.SimulationId}','{DateTime.Now.ToString("MM /dd/yyyy HH:mm:ss.fff")}','LabelWork','設定工站',NULL,'{keys.Station}','{partNO}','{keys.OrderNO}','{key_OPNO}',{keys.IndexSN})");
                                keys.Station_State = "停止...";

                                #region 更新Tag
                                DataRow dr_LabelStateINFO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr["Config_macID"].ToString()}'");
                                if (dr_LabelStateINFO != null && dr["Config_macID"].ToString().Trim() != "")
                                {
                                    string tmp_s = $"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{keys.Station};{typevalue};{keys.IndexSN}";
                                    DataRow totalData = _Fun.GetAvgCTWTandTotalOutput(db, false, sfcdr["OrderNO"].ToString(), keys.Station, keys.IndexSN);
                                    string dis_DetailQTY = "0";
                                    if (totalData != null)
                                    {
                                        dis_DetailQTY = totalData["TotalOutput"].ToString();
                                    }
                                    string isUpdate = "1";
                                    var json1 = "";
                                    var json_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{sfcdr["OrderNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{partNO}\",\"Text3\":\"品名:\",\"PartName\":\"{partName}\",\"Text4\":\"需求量:\",\"DetailQTY\":\"{dis_DetailQTY}\",\"Text5\":\"CT:\",\"EfficientCT\":\"{ct.ToString()}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                                    var writeShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{sfcdr["OrderNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{partNO}\",\"Text3\":\"品名:\",\"PartName\":\"{partName}\",\"Text4\":\"需求量:\",\"QTY\":\"{num.ToString()}\",\"Text5\":\"CT:\",\"EfficientCT\":\"{ct.ToString()}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                                    if (dr_LabelStateINFO["Version"].ToString().Trim() != "" && dr_LabelStateINFO["Version"].ToString().Trim().Substring(0, 2) == "42")
                                    {
                                        json_ShowValue = $"{json_ShowValue},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_StationNO["StationNO"].ToString()}\",\"text17\":\"{dr_StationNO["StationName"].ToString()}\"";
                                        writeShowValue = $"{writeShowValue},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_StationNO["StationNO"].ToString()}\",\"text17\":\"{dr_StationNO["StationName"].ToString()}\"";
                                        json1 = $"\"mac\":\"{dr["Config_macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{json_ShowValue}";
                                    }
                                    else
                                    {
                                        json_ShowValue = $"{json_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                        writeShowValue = $"{writeShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                        json1 = $"\"mac\":\"{dr["Config_macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{json_ShowValue}";
                                    }
                                    if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                    {
                                        _Fun.Tag_Write(db,dr["Config_macID"].ToString(),"設定工單", $"{json1},\"QTY\":\"{num.ToString()}\",\"ledrgb\":\"0\",\"ledstate\":0");
                                    }
                                    else { isUpdate = "0"; }
                                    if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{writeShowValue}',Ledrgb='0',Ledstate=0,StationNO='{keys.Station}',Type='1',OrderNO='{sfcdr["OrderNO"].ToString()}',IndexSN='{keys.IndexSN}',StoreNO='',StoreSpacesNO='',QTY={dis_DetailQTY},IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr["Config_macID"].ToString()}'"))
                                    {
                                    }
                                    if (isUpdate == "0")
                                    {
                                        keys.ERRMsg = $"{keys.Station} 工站 傳送電子訊號失敗, 請通知管理者"; goto break_FUN;
                                    }
                                }
                                #endregion
                            }
                            break;
                        case "設定人員":
                            {
                                string opNOs = keys.OPNO;
                                DataRow dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                if (dr != null && dr["OP_NO"].ToString().Trim() != "")
                                {
                                    List<string> tmp = dr["OP_NO"].ToString().Trim().Split(';').ToList();
                                    if (!tmp.Contains(keys.OPNO))
                                    {
                                        opNOs = $"{dr["OP_NO"].ToString().Trim()};{keys.OPNO}";
                                    }
                                    else { break; }
                                }
                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET OP_NO='{opNOs}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) 
                                VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelWork','設定人員','{dr["PP_Name"].ToString()}','{keys.Station}','{dr["PartNO"].ToString()}','{dr["OrderNO"].ToString()}','{key_OPNO}',{dr["IndexSN"].ToString()})");
                            }
                            break;
                        case "報工":
                            string message = "";
                            string stackTrace = "";
                            string ViewBagERRMsg = "";
                            bool is_reportOK = _SFC_Common.Reporting_LabelWork(db, dr_StationNO, br.UserNO, keys,false, ref message, ref stackTrace, ref ViewBagERRMsg);
                            if (ViewBagERRMsg != "")
                            { ViewBag.ERRMsg = $"作業失敗, 畫面已逾時, 請關閉網頁瀏覽器, 重新刷條碼 並 重新操作."; return View("ResuItTimeOUT", keys); }
                            else if (message != "")
                            {
                                ViewBag.ERRMsg = $"程式異常, {message},  請通知管理者.";
                                System.Threading.Tasks.Task task = _Log.ErrorAsync($"LabelWorkController.cs SetAction {keys.State} Exception: {message} {stackTrace}", true);
                                return View("ResuItTimeOUT", keys);
                            }
                            else
                            {
                                if (!is_reportOK) { goto break_FUN; }
                            }

                            /*
                            {
                                //###???若此處有改 TMM Service 55 ㄝ要改
                                DateTime startTime = DateTime.Now;
                                string meg = "";
                                try
                                {
                                    string[] data = new string[7] { keys.Station, keys.OrderNO, keys.IndexSN, keys.OKQTY.ToString(), keys.FailQTY.ToString(), keys.OPNO, keys.LocalIPPort };
                                    int outQTY = int.Parse(data[3]);//報工良品數量
                                    int failQTY = int.Parse(data[4]);//報工不良品數量
                                    DataRow dr_Manufacture = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[0]}'");
                                    DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{keys.SimulationId}'");

                                    #region 檢查網頁來源資料
                                    if (dr_Manufacture["OrderNO"].ToString() != keys.OrderNO || dr_Manufacture["IndexSN"].ToString() != keys.IndexSN) 
                                    {
                                        ViewBag.ERRMsg = $"作業失敗, 畫面已逾時, 請關閉網頁瀏覽器, 重新刷條碼 並 重新操作."; return View("ResuItTimeOUT", keys);
                                    }
                                    if (keys.OKQTY <= 0 && keys.FailQTY <= 0) { keys.ERRMsg = $"報工數量合計不能為0, 請重新報工."; goto break_FUN; }
                                    if (dr_Manufacture.IsNull("StartTime")) { keys.ERRMsg = $"{data[1]} 工站沒有開工紀錄, 請先執行開工, 才能報工."; goto break_FUN; }
                                    DataRow dr_WO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}'");
                                    if (dr_WO == null) { keys.ERRMsg = $"查無 {data[0]} 工單資料紀錄, 請聯繫系統管理者."; goto break_FUN; }
                                    sql = $"SELECT * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}' AND StationNO='{data[0]}' AND IndexSN={data[2]}";
                                    DataRow dr = db.DB_GetFirstDataByDataRow(sql);
                                    if (dr == null && dr_APS_Simulation!=null) 
                                    {
                                        string isend = dr_APS_Simulation["PartSN"].ToString().Trim() == "0" ? "1" : "0";
                                        string is_IndexSN_Merge= dr_APS_Simulation.IsNull("StationNO_Merge") ? "1" : "0";
                                        string sfc_sql = $@"INSERT INTO SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement (
                                                                                    [OrderNO],[StationNO],[StationName],[PP_Name],[IndexSN],[DisplaySN],
                                                                                    [IsLastStation],[Sub_PP_Name],[StatndCycleTime],[UpdateTime],[IndexSN_Merge],
                                                                                    [StartTime],[CumulativeTime],[AvarageCycleTime],[TotalCheckIn],[TotalCheckOut],
                                                                                    [TotalInput],[TotalOutput],[TotalFail],[TotalKeep],[FPY],
                                                                                    [YieldRate],[StationYieldRate],ServerId) VALUES 
                                                                                    ('{data[1]}','{data[0]}','{dr_StationNO["StationName"].ToString()}','{dr_Manufacture["PP_Name"].ToString()}',
                                                                                    {data[2]},0,'{isend}','{dr_Manufacture["PP_Name"].ToString()}',{dr_APS_Simulation["Math_StandardCT"].ToString()},'{DateTime.Now.ToString("MM/dd/yyyy H:mm:ss")}',
                                                                                    '{is_IndexSN_Merge}',null,0,0,0,0,0,0,0,0,0,0,0,'{_Fun.Config.ServerId}')";

                                        if (!db.DB_SetData(sfc_sql))
                                        {
                                            keys.ERRMsg = $"查無 PP_WorkOrder_Settlement 資料紀錄, 請聯繫系統管理者."; goto break_FUN;
                                        }
                                        else 
                                        { 
                                            dr = db.DB_GetFirstDataByDataRow(sql);
                                            if (dr == null) { keys.ERRMsg = $"查無 PP_WorkOrder_Settlement 資料紀錄, 請聯繫系統管理者."; goto break_FUN; }
                                        }
                                    }
                                    if (dr == null) { keys.ERRMsg = $"查無 PP_WorkOrder_Settlement 資料紀錄, 請聯繫系統管理者."; goto break_FUN; }

                                    if (!dr.IsNull("StartTime") && dr["StartTime"].ToString().Trim() != "")
                                    { startTime = Convert.ToDateTime(dr["StartTime"]); }
                                    #endregion

                                    #region 計算CT
                                    decimal ct = 0;
                                    DateTime rRemarkTimeS = startTime;
                                    if (dr_Manufacture.IsNull("RemarkTimeS"))
                                    { rRemarkTimeS = Convert.ToDateTime(dr_Manufacture["StartTime"]); }
                                    else
                                    { rRemarkTimeS = Convert.ToDateTime(dr_Manufacture["RemarkTimeS"]); }
                                    #region 先查相同人員與站與工單, 是否報工過
                                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and IndexSN={keys.IndexSN} and OrderNO='{keys.OrderNO}' and OperateType like '%報工%' and OP_NO like '%{keys.OPNO}%'");
                                    if (tmp_dr != null && Convert.ToDateTime(tmp_dr["LOGDateTime"]) > rRemarkTimeS) { rRemarkTimeS = Convert.ToDateTime(tmp_dr["LOGDateTime"]); }
                                    #endregion
                                    if (dr_Manufacture.IsNull("RemarkTimeE") || rRemarkTimeS >= Convert.ToDateTime(dr_Manufacture["RemarkTimeE"]))
                                    { 
                                        ct = _WebSocket.TimeCompute2Seconds(rRemarkTimeS, DateTime.Now) / (keys.OKQTY + keys.FailQTY);
                                        if (ct <= 0 && !dr_Manufacture.IsNull("RemarkTimeE") && Convert.ToDateTime(dr_Manufacture["RemarkTimeE"]) > Convert.ToDateTime(dr_Manufacture["RemarkTimeS"]))
                                        {
                                            ct = _WebSocket.TimeCompute2Seconds(Convert.ToDateTime(dr_Manufacture["RemarkTimeS"]), Convert.ToDateTime(dr_Manufacture["RemarkTimeE"])) / (keys.OKQTY + keys.FailQTY);
                                        }
                                    }
                                    else
                                    {
                                        ct = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), rRemarkTimeS, Convert.ToDateTime(dr_Manufacture["RemarkTimeE"])) / (keys.OKQTY + keys.FailQTY);
                                        if (ct <= 0 && Convert.ToDateTime(dr_Manufacture["RemarkTimeE"]) > rRemarkTimeS)
                                        {
                                            ct = _WebSocket.TimeCompute2Seconds(rRemarkTimeS, Convert.ToDateTime(dr_Manufacture["RemarkTimeE"])) / (keys.OKQTY + keys.FailQTY);
                                        }
                                    }
                                    decimal ct_log = ct < 1 ? 0 : ct;

                                    int ops = dr_Manufacture["OP_NO"].ToString().Split(';').Length;
                                    if (ops > 1) { ct = ct / ops; }
                                    #endregion

                                    #region 寫SFC_StationDetail
                                    string partNO = dr_Manufacture["PartNO"].ToString();
                                    string old_InTime = "";
                                    string old_OutTime = "";
                                    string old_ProductFinishedQty = "0";
                                    string old_ProductFailedQty = "0";
                                    int OP_Count = 1;
                                    string logTime = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff");

                                    DataRow dr_StationDetail = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail] WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}' AND StationNO='{data[0]}' AND IndexSN={data[2]}");
                                    if (dr_StationDetail == null)
                                    {
                                        //###???PP_Name暫時
                                        sql = string.Format(
                                            @"INSERT INTO SoftNetLogDB.[dbo].[SFC_StationDetail] (
                                                [LOGDateTime],[Master_PP_Name],[PP_Name],[OP_NO],[StationNO],
                                                [IndexSN],[IndexSN_Merge],[OrderNO],[PartNO],[InTime],
                                                [OutTime],[CycleTime],[InputFlag],[OutputFlag],[FailFlag],
                                                [Station_Type],[ProductFinishedQty],[ProductFailedQty],[SerialNO],[RMSName],SimulationId,ServerId) VALUES (
                                                '{0}','{1}','{2}','{3}','{4}',{5},'{6}','{7}','{8}','{9}',
                                                '{10}',{11},'{12}','{13}','{14}','{15}',{16},{17},'','{18}','{19}','{20}')",
                                            logTime,
                                            dr["PP_Name"].ToString(),
                                            dr["PP_Name"].ToString(),//###???暫時換dr["Sub_PP_Name"].ToString(),
                                            data[5],//OPNO
                                            data[0],
                                            data[2],
                                            dr["IndexSN_Merge"].ToString(),
                                            data[1],
                                            dr_WO["PartNO"].ToString(),
                                            startTime.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                            ct.ToString(),
                                            (outQTY + failQTY) > 0 ? 1 : 0,
                                            outQTY > 0 ? 1 : 0,
                                            failQTY > 0 ? 1 : 0,
                                            dr_StationNO["Station_Type"].ToString(),
                                            outQTY.ToString(),
                                            failQTY.ToString(), dr_StationNO["RMSName"].ToString(), keys.SimulationId, _Fun.Config.ServerId);
                                    }
                                    else
                                    {
                                        logTime = Convert.ToDateTime(dr_StationDetail["LOGDateTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff");
                                        old_InTime = Convert.ToDateTime(dr_StationDetail["InTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff");
                                        old_OutTime = Convert.ToDateTime(dr_StationDetail["OutTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff");
                                        old_ProductFinishedQty = dr_StationDetail["ProductFinishedQty"].ToString();
                                        old_ProductFailedQty = dr_StationDetail["ProductFailedQty"].ToString();
                                        if (int.Parse(dr_StationDetail["CycleTime"].ToString()) != 0) { ct = (ct + int.Parse(dr_StationDetail["CycleTime"].ToString())) > 0 ? Math.Round((ct + int.Parse(dr_StationDetail["CycleTime"].ToString())) / 2) : ct; }
                                        if (ct < 1) { ct = 0; }
                                        string sId = "";
                                        if (keys.SimulationId == "")
                                        { sId = ""; }
                                        else
                                        { sId = $"and SimulationId='{keys.SimulationId}'"; }
                                        sql = string.Format(
                                                @"UPDATE SoftNetLogDB.[dbo].[SFC_StationDetail] 
                                                SET [ProductFinishedQty]+={0}, [ProductFailedQty]+={1},
                                                [InTime]='{2}',[OutTime]='{3}',[CycleTime]={4} 
                                                WHERE ServerId='{9}' and OrderNO = '{5}' AND StationNO = '{6}' AND IndexSN={7} {8}",
                                                outQTY,
                                                failQTY,
                                                startTime.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                                ct.ToString(),
                                                data[1],
                                                data[0], data[2], sId, _Fun.Config.ServerId);
                                    }
                                    #endregion

                                    if (db.DB_SetData(sql))
                                    {
                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelWork','智慧報工','{dr_Manufacture["PP_Name"].ToString()}','{keys.Station}','{dr_Manufacture["PartNO"].ToString()}','{dr_Manufacture["OrderNO"].ToString()}','{br.UserNO}',{dr_Manufacture["IndexSN"].ToString()})");

                                        #region 更新標籤累計量
                                        if (outQTY != 0)
                                        {
                                            DataRow totalData = _Fun.GetAvgCTWTandTotalOutput(db, false, data[1], data[0], data[2]);
                                            string dis_DetailQTY = "0";
                                            if (totalData != null)
                                            {
                                                dis_DetailQTY = totalData["TotalOutput"].ToString().Trim();
                                                totalData = db.DB_GetFirstDataByDataRow($"select b.* from SoftNetMainDB.[dbo].[Manufacture] as a join SoftNetMainDB.[dbo].[LabelStateINFO] as b on b.macID=a.Config_macID where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{data[0]}'");
                                                if (totalData != null && dis_DetailQTY != totalData["QTY"].ToString().Trim())
                                                {
                                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set QTY='{dis_DetailQTY}',IsUpdate='0' where ServerId='{_Fun.Config.ServerId}' and macID='{totalData["macID"].ToString()}'");
                                                }
                                            }
                                        }
                                        #endregion

                                        #region log SFC_StationDetail_ChangeLOG紀錄
                                        int reportTime = 0;
                                        if (dr_StationDetail != null)
                                        {
                                            #region 計算上一次與現在時間差
                                            DataRow d2 = db.DB_GetFirstDataByDataRow($"SELECT LOGDateTimeID FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(dr_StationDetail["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by LOGDateTime,LOGDateTimeID desc");
                                            if (d2 != null)
                                            { reportTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(),Convert.ToDateTime(d2["LOGDateTimeID"]), DateTime.Now); }
                                            #endregion
                                        }
                                        string wsid = "NULL";
                                        if (keys.SimulationId != "") { wsid = $"'{keys.SimulationId}'"; }
                                        OP_Count = data[5].Split(";").Count();
                                        if (OP_Count <= 0) { OP_Count = 1; }
                                        string LOGDateTimeID = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                                        #region 查詢PP_EfficientDetail
                                        string eCT = "0";
                                        string upperCT = "0";
                                        string lowerCT = "0";
                                        DataRow dr_tmp_ct = db.DB_GetFirstDataByDataRow($"select AVG(EfficientCycleTime) as ECT,AVG(SD_UpperLimit) as UpperCT,AVG(SD_LowerLimit) as LowerCT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_Manufacture["PartNO"]}' and StationNO='{dr_Manufacture["StationNO"]}' and PP_Name='{dr_Manufacture["PP_Name"].ToString()}' and IndexSN={dr_Manufacture["IndexSN"].ToString()} and DOCNO=''");
                                        if (dr_tmp_ct != null)
                                        {
                                            if (!dr_tmp_ct.IsNull("ECT") && dr_tmp_ct["ECT"].ToString() != "") { eCT = dr_tmp_ct["ECT"].ToString(); }
                                            if (!dr_tmp_ct.IsNull("UpperCT") && dr_tmp_ct["UpperCT"].ToString() != "") { upperCT = dr_tmp_ct["UpperCT"].ToString(); }
                                            if (!dr_tmp_ct.IsNull("LowerCT") && dr_tmp_ct["LowerCT"].ToString() != "") { lowerCT = dr_tmp_ct["LowerCT"].ToString(); }
                                        }
                                        #endregion
                                        sql = $@"INSERT INTO SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] (LOGDateTime,LOGDateTimeID,OLD_InTime,OLD_OutTime,EditFinishedQty,EditFailedQty,OLD_ProductFinishedQty,OLD_ProductFailedQty,OP_Count,OP_NO,ServerId,StationNO,ReportTime,PartNO,SimulationId,PP_Name,IndexSN,CycleTime,ECT,LowerCT,UpperCT)
                                            VALUES ('{logTime}',
                                            '{LOGDateTimeID}',
                                            '{old_InTime}',
                                            '{old_OutTime}',
                                            {outQTY},
                                            {failQTY},
                                            {old_ProductFinishedQty},
                                            {old_ProductFailedQty},
                                            {OP_Count.ToString()},
                                            '{data[5]}','{_Fun.Config.ServerId}','{data[0]}',{reportTime.ToString()},'{partNO}',{wsid},'{dr["PP_Name"].ToString()}',{dr["IndexSN"].ToString()},{ct_log.ToString()},{eCT},{lowerCT},{upperCT})";
                                        if (db.DB_SetData(sql))
                                        {
                                            //###???
                                        }
                                        #endregion

                                        #region 計算效能 PP_EfficientDetail處理
                                        {
                                            List<double> allCT = new List<double>();//list for all avg value
                                            string top_flag = "";
                                            try
                                            {
                                                if (_Fun.Config.AdminKey03 != 0) { top_flag = $" TOP {_Fun.Config.AdminKey03} "; }
                                                DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} PP_Name,StationNO,PartNO as Sub_PartNO,CycleTime,WaitTime,(EditFinishedQty+EditFailedQty) as QTY from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG]
                                                    where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and PartNO='{dr_Manufacture["PartNO"].ToString()}' and PP_Name='{dr_WO["PP_Name"].ToString()}' and IndexSN={dr_Manufacture["IndexSN"].ToString()} and EditFinishedQty!=0 and CycleTime!=0");
                                                if (dt_Efficient != null && dt_Efficient.Rows.Count > 0)
                                                {
                                                    double efficient_CT = 0;
                                                    foreach (DataRow dr_tmp in dt_Efficient.Rows)
                                                    {
                                                        if (_Fun.Config.AdminKey14)
                                                        { efficient_CT = double.Parse(dr_tmp["CycleTime"].ToString()) + double.Parse(dr_tmp["WaitTime"].ToString()); }
                                                        else
                                                        { efficient_CT = double.Parse(dr_tmp["CycleTime"].ToString()); }
                                                        for (int tmp01 = 1; tmp01 <= (int)dr_tmp["QTY"]; tmp01++)//工單數目若為2 需算作兩筆
                                                        {
                                                            allCT.Add(efficient_CT);
                                                        }
                                                    }
                                                    if (allCT.Count > 0)
                                                    {
                                                        _WebSocket.SfcTimerloopthread_Tick_Efficient(db, allCT, keys.Station, dr_WO["PP_Name"].ToString(), dr_Manufacture["PP_Name"].ToString(), dr_Manufacture["IndexSN"].ToString(), dr_WO["PartNO"].ToString(), dr_Manufacture["PartNO"].ToString(), "");
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs 計算效能PP_EfficientDetail處理 Exception: {ex.Message} {ex.StackTrace}", true);
                                            }
                                        }
                                        #endregion

                                        #region 記錄刀工具使用時數
                                        DataTable tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStockII_Knives] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and IsDel='0'");
                                        if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                        {
                                            int useTime = 0;
                                            int useCount = outQTY + failQTY;
                                            string k_stime = "";
                                            foreach (DataRow d in tmp_dt.Rows)
                                            {
                                                if (!dr_Manufacture.IsNull("RemarkTimeS"))
                                                {
                                                    if (d.IsNull("StartTime")) { k_stime =$",StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}'"; } else { k_stime = ""; }
                                                    if (!dr_Manufacture.IsNull("RemarkTimeE"))
                                                    {
                                                        tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and PartNO='{dr_Manufacture["PartNO"].ToString()}' and LOGDateTime>'{Convert.ToDateTime(dr_Manufacture["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}' and OperateType like '%報工%'");
                                                        if ( tmp_dr != null )
                                                        {
                                                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                            ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.Station}','{dr_Manufacture["PartNO"].ToString()}',{(useCount).ToString()},0)");
                                                        }
                                                        else
                                                        {
                                                            useTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), Convert.ToDateTime(dr_Manufacture["RemarkTimeS"].ToString()), DateTime.Now);
                                                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                            ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.Station}','{dr_Manufacture["PartNO"].ToString()}',{(useCount).ToString()},{useTime.ToString()})");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        useTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), Convert.ToDateTime(dr_Manufacture["RemarkTimeS"].ToString()), DateTime.Now);
                                                        db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                        ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.Station}','{dr_Manufacture["PartNO"].ToString()}',{(useCount).ToString()},{useTime.ToString()})");
                                                    }
                                                    db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStockII_Knives] set TOTWorkTime+={useTime.ToString()},TOTCount+={useCount.ToString()}{k_stime} where ServerId='{_Fun.Config.ServerId}' and KId='{d["KId"].ToString()}'");
                                                }
                                            }
                                        }
                                        #endregion

                                        #region 修正工站開始日期
                                        if (dr_Manufacture["State"].ToString() == "1")
                                        { db.DB_SetData($"update SoftNetMainDB.[dbo].[Manufacture] set RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[0]}'"); }
                                        #endregion

                                        _WebSocket.Update_PP_WorkOrder_Settlement(db, data[1], keys.SimulationId);

                                        //###??? 不良數量尚未處裡
                                        if (keys.SimulationId != "")
                                        {
                                            bool isNeedQTY_OK = false;//判斷本站數量已足夠
                                            string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();

                                            string in_NO = "AC01";//###??? 暫時寫死領料單別
                                            string inOK_NO = "BC01";//###??? 暫時寫死入庫單別

                                            DataRow dr_APS_PartNOTimeNote = null;
                                            if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{data[1].Trim()}',Detail_QTY+={data[3]},Detail_Fail_QTY+={data[4]} where SimulationId='{keys.SimulationId}'"))
                                            {
                                                dr_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SimulationId}'");
                                                if ((int.Parse(dr_APS_PartNOTimeNote["Detail_QTY"].ToString()) + int.Parse(dr_APS_PartNOTimeNote["Detail_Fail_QTY"].ToString()) - int.Parse(dr_APS_PartNOTimeNote["NeedQTY"].ToString())) >= 0)
                                                { isNeedQTY_OK = true; }
                                            }

                                            //尋找相關BOM原物料
                                            #region 扣Keep量 與 處理領料單單據
                                            DataTable dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_APS_Simulation["NeedId"].ToString()}' and Apply_PP_Name='{dr_Manufacture["PP_Name"].ToString()}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and IndexSN={data[2]} order by PartSN desc");
                                            if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                                            {
                                                string docNumberNO = "";
                                                foreach (DataRow d in dt_APS_Simulation.Rows)
                                                {
                                                    #region 處裡移轉量 APS_PartNOTimeNote
                                                    if (!d.IsNull("Source_StationNO") && (d["Class"].ToString() == "4" || d["Class"].ToString() == "5"))
                                                    {
                                                        if (d["PartSN"].ToString() == "0")
                                                        {
                                                            #region 工單最後一站預開入庫單  
                                                            string tmp_no = "";
                                                            string in_StoreNO = "";
                                                            string in_StoreSpacesNO = "";
                                                            tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 Order by b.StoreOrder");
                                                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                                            {
                                                                int tmp_int = outQTY;
                                                                #region 有計畫Keep量  by StoreOrder順序扣
                                                                foreach (DataRow d2 in tmp_dt.Rows)
                                                                {
                                                                    if (in_StoreNO == "")
                                                                    {
                                                                        in_StoreNO = d2["StoreNO"].ToString();
                                                                        in_StoreSpacesNO = d2["StoreSpacesNO"].ToString();
                                                                    }
                                                                    if (tmp_int > 0)
                                                                    {
                                                                        if (int.Parse(d2["KeepQTY"].ToString()) >= tmp_int)
                                                                        {
                                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY-={tmp_int} where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                            _WebSocket.Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO,true);
                                                                            tmp_int = 0;
                                                                            break;
                                                                        }
                                                                        else
                                                                        {
                                                                            int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                            _WebSocket.Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_01, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO, true);
                                                                            tmp_int -= tmp_01;
                                                                        }
                                                                    }
                                                                }
                                                                if (tmp_int > 0)
                                                                {
                                                                    #region 計畫量不夠扣, 入實體倉
                                                                    if (in_StoreNO == "")
                                                                    {
                                                                        DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' and StoreNO!=''");
                                                                        if (tmp != null)
                                                                        {
                                                                            in_StoreNO = tmp["StoreNO"].ToString();
                                                                            in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                                                            _WebSocket.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO, true);
                                                                        }
                                                                        else
                                                                        {
                                                                            #region 查找適合入庫儲別
                                                                            _WebSocket.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                                                            #endregion
                                                                            #region 無倉紀錄, 加空倉
                                                                            _WebSocket.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO, true);
                                                                            #endregion
                                                                        }
                                                                    }
                                                                    #endregion

                                                                }
                                                                #endregion
                                                            }
                                                            else
                                                            {
                                                                #region 查找適合入庫儲別
                                                                _WebSocket.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                                                #endregion
                                                                #region 無倉紀錄, 加空倉
                                                                _WebSocket.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, int.Parse(data[3]), "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO, true);
                                                                #endregion
                                                            }

                                                            sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{tmp_no}',Next_StoreQTY+={data[3]} where SimulationId='{d["SimulationId"].ToString()}'";
                                                            
                                                            if (db.DB_SetData(sql))
                                                            {
                                                                #region 處理工站移轉時間

                                                                #endregion

                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            #region 非最後一站
                                                            int tmp_int = (int.Parse(d["BOMQTY"].ToString()) * outQTY);
                                                            DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d["SimulationId"].ToString()}'");
                                                            int wr_next_StationQTY = 0;
                                                            #region 檢查上一階之前是否有入退庫數量(BC01)
                                                            if (int.Parse(tmp["Next_StoreQTY"].ToString()) > 0)
                                                            {
                                                                int store_tmp = tmp_int;
                                                                DataTable dt_Store = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{tmp["Store_DOCNumberNO"].ToString()}' order by ArrivalDate");
                                                                if (dt_Store != null && dt_Store.Rows.Count > 0)
                                                                {
                                                                    int wrQTY = 0;
                                                                    foreach (DataRow dr_DOC3stockII in dt_Store.Rows)
                                                                    {
                                                                        if (store_tmp <= 0) { break; }
                                                                        if (store_tmp >= int.Parse(dr_DOC3stockII["QTY"].ToString()))
                                                                        {
                                                                            store_tmp -= int.Parse(dr_DOC3stockII["QTY"].ToString());
                                                                            wrQTY = int.Parse(dr_DOC3stockII["QTY"].ToString());
                                                                            wr_next_StationQTY += wrQTY;
                                                                        }
                                                                        else
                                                                        {
                                                                            //拆單處置
                                                                            wrQTY = store_tmp;
                                                                            wr_next_StationQTY += store_tmp;
                                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set QTY={store_tmp} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}'");
                                                                            db.DB_SetData($@"INSERT INTO [dbo].[DOC3stockII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],ArrivalDate) VALUES 
                                                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{dr_DOC3stockII["DOCNumberNO"].ToString()}','{dr_DOC3stockII["PartNO"].ToString()}',{dr_DOC3stockII["Price"].ToString()},'{dr_DOC3stockII["Unit"].ToString()}',{(int.Parse(dr_DOC3stockII["QTY"].ToString()) - store_tmp).ToString()}
                                                                                                ,'{dr_DOC3stockII["Remark"].ToString()}','{dr_DOC3stockII["SimulationId"].ToString()}','{dr_DOC3stockII["IsOK"].ToString()}','{dr_DOC3stockII["IN_StoreNO"].ToString()}','{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}','{dr_DOC3stockII["OUT_StoreNO"].ToString()}','{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}','{Convert.ToDateTime(dr_DOC3stockII["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                                            store_tmp = 0;
                                                                        }
                                                                        if (!bool.Parse(dr_DOC3stockII["IsOK"].ToString()))
                                                                        {
                                                                            #region 寫入庫存
                                                                            if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                                                                            { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={wrQTY} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["IN_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}'"); }
                                                                            else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={wrQTY} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'"); }
                                                                            #endregion
                                                                            #region 計算單據CT,平均,有效, 寫SFC_StationProjectDetail
                                                                            int typeTotalTime = 0;
                                                                            string writeSQL = "";
                                                                            if (!dr_DOC3stockII.IsNull("StartTime")) { typeTotalTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC3stockII["StartTime"].ToString()), DateTime.Now); }
                                                                            else { writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime}{writeSQL} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}' and IsOK='0'");
                                                                            string efficient_partNO = dr_DOC3stockII["PartNO"].ToString();
                                                                            string efficient_pp_Name = "";
                                                                            string E_stationNO = "";
                                                                            if (dr_DOC3stockII["SimulationId"].ToString() != "")
                                                                            {
                                                                                DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC3stockII["SimulationId"].ToString()}'");
                                                                                efficient_pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                                                                                if (!dr_tmp.IsNull("Source_StationNO") && dr_tmp["Source_StationNO"].ToString() != "")
                                                                                { E_stationNO = dr_tmp["Source_StationNO"].ToString(); }
                                                                                else { E_stationNO = dr_tmp["Apply_StationNO"].ToString(); }
                                                                            }
                                                                            DataTable dt_Efficient = db.DB_GetData($@"select TOP {_Fun.Config.AdminKey03} CT from SoftNetMainDB.[dbo].[DOC3stockII] where SUBSTRING(Id,2,2)='{_Fun.Config.ServerId}' and SUBSTRING(DOCNumberNO,1,4)='{dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4)}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and CT>0");
                                                                            List<double> allCT = new List<double>();
                                                                            if (dt_Efficient != null && dt_Efficient.Rows.Count > 0)
                                                                            {
                                                                                for (int i2 = 0; i2 < dt_Efficient.Rows.Count; i2++)
                                                                                {
                                                                                    foreach (DataRow dr2 in dt_Efficient.Rows)
                                                                                    {
                                                                                        allCT.Add(double.Parse(dr2["CT"].ToString()));
                                                                                    }
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                if (typeTotalTime != 0)
                                                                                { allCT.Add(typeTotalTime); }
                                                                            }
                                                                            if (allCT.Count > 0)
                                                                            {
                                                                                _WebSocket.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, efficient_pp_Name, efficient_pp_Name,"0", efficient_partNO, efficient_partNO, dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4));
                                                                            }
                                                                            #endregion
                                                                        }
                                                                        //開領料單
                                                                        string tmpDOCNO = dr_DOC3stockII["DOCNumberNO"].ToString();
                                                                        DataRow tmp_store = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and IN_StoreNO!=''");
                                                                        if (tmp_store != null)
                                                                        {
                                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                            _WebSocket.Create_DOC3stock(db, d, tmp_store["IN_StoreNO"].ToString(), tmp_store["IN_StoreSpacesNO"].ToString(), "", "", in_NO, wrQTY, "", "", $"{stationno}站 入庫後再領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmpDOCNO, br.UserNO, true);
                                                                        }
                                                                    }
                                                                    if (wr_next_StationQTY > 0)
                                                                    {
                                                                        //將入庫數量Next_StoreQTY扣除回寫Next_StationQTY
                                                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set Next_StoreQTY-={wr_next_StationQTY} where SimulationId='{d["SimulationId"].ToString()}'");
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                            int detail_QTY = int.Parse(tmp["Detail_QTY"].ToString());
                                                            int next_StationQTY = int.Parse(tmp["Next_StationQTY"].ToString());

                                                            int next_next_detail_QTY = 0;
                                                            #region 檢查下一階是否有偷先報工未移轉
                                                            DataRow dr_next_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SimulationId}'");
                                                            next_next_detail_QTY = int.Parse(dr_next_APS_PartNOTimeNote["Detail_QTY"].ToString());
                                                            if (next_next_detail_QTY > 0)
                                                            {
                                                                //detail_QTY - next_StationQTY=上一階數量
                                                                if ((next_StationQTY + tmp_int) < next_next_detail_QTY) { next_next_detail_QTY -= (next_StationQTY + tmp_int); }
                                                                else { next_next_detail_QTY = 0; }
                                                            }
                                                            #endregion

                                                            if ((detail_QTY - next_StationQTY) >= tmp_int)
                                                            {
                                                                //在製移轉 
                                                                sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Next_APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}',Next_StationQTY+={tmp_int + next_next_detail_QTY} where SimulationId='{d["SimulationId"].ToString()}'";
                                                                if (db.DB_SetData(sql))
                                                                {
                                                                    #region 處理工站移轉時間
                                                                    
                                                                    #endregion

                                                                }
                                                            }
                                                            else
                                                            {
                                                                bool is_run = true;
                                                                //將在製移轉 剩餘
                                                                tmp_int -= (detail_QTY - next_StationQTY);
                                                                sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Next_APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}',Next_StationQTY={tmp["Detail_QTY"].ToString()} where SimulationId='{d["SimulationId"].ToString()}'";
                                                                if (db.DB_SetData(sql) && int.Parse(tmp["Detail_QTY"].ToString()) > 0)
                                                                {
                                                                    #region 處理工站移轉時間
                                                                   
                                                                    #endregion

                                                                }

                                                                #region 先檢查是否已有單據, 且已移轉多少量
                                                                int stockQTY = 0;
                                                                tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) as qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}'");
                                                                if (tmp != null && !tmp.IsNull("qty"))
                                                                {
                                                                    stockQTY = int.Parse(tmp["qty"].ToString());
                                                                    if (stockQTY <= (detail_QTY + next_StationQTY)) { stockQTY = 0; }
                                                                }
                                                                if ((stockQTY - tmp_int) >= 0)
                                                                {
                                                                    is_run = false;
                                                                }
                                                                else
                                                                {
                                                                    tmp_int = tmp_int - stockQTY;
                                                                }
                                                                #endregion

                                                                if (tmp_int <= 0) { is_run = false; }
                                                                if (is_run)
                                                                {
                                                                    tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 order by b.StoreOrder");
                                                                    if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                                                    {
                                                                        #region 有計畫Keep量  by StoreOrder順序扣
                                                                        foreach (DataRow d2 in tmp_dt.Rows)
                                                                        {
                                                                            if (tmp_int > 0)
                                                                            {
                                                                                if (int.Parse(d2["KeepQTY"].ToString()) >= tmp_int)
                                                                                {
                                                                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY-={tmp_int} where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                    _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO.Trim());
                                                                                    tmp_int = 0;
                                                                                    break;
                                                                                }
                                                                                else
                                                                                {
                                                                                    int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                                                    //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_01} where Id='{d2["Id"].ToString()}'");
                                                                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                    _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO.Trim());
                                                                                    tmp_int -= tmp_01;
                                                                                }
                                                                            }
                                                                        }
                                                                        if (tmp_int > 0)
                                                                        {
                                                                            #region 計畫量不夠扣, 扣實體倉
                                                                            DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                                            {
                                                                                foreach (DataRow d2 in tmp_dt2.Rows)
                                                                                {
                                                                                    if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                                    {
                                                                                        //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_int} where Id='{d2["Id"].ToString()}'");
                                                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                        _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                                        tmp_int = 0;
                                                                                        break;
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                                        //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_01} where Id='{d2["Id"].ToString()}'");
                                                                                        if (tmp_01 != 0)
                                                                                        {
                                                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                            _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                                            tmp_int -= tmp_01;
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                            #endregion

                                                                            #region 實體倉不購扣, 扣空倉
                                                                            if (tmp_int > 0)
                                                                            {
                                                                                #region 查找適合出庫儲別
                                                                                string out_StoreNO = "";
                                                                                string out_StoreSpacesNO = "";
                                                                                _WebSocket.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                                                #endregion
                                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                _WebSocket.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO.Trim());
                                                                            }
                                                                            #endregion
                                                                        }
                                                                        #endregion
                                                                    }
                                                                    else
                                                                    {
                                                                        #region 沒計畫量, 扣實體倉
                                                                        DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                                        if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                                        {
                                                                            foreach (DataRow d2 in tmp_dt2.Rows)
                                                                            {
                                                                                if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                                {
                                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                    _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno} 沒計畫量,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                                    tmp_int = 0;
                                                                                    break;
                                                                                }
                                                                                else
                                                                                {
                                                                                    int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                                    if (tmp_01 != 0)
                                                                                    {
                                                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                        _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno} 沒計畫量,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                                        tmp_int -= tmp_01;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                        if (tmp_int > 0)
                                                                        {
                                                                            #region 查找適合庫儲別
                                                                            string out_StoreNO = "";
                                                                            string out_StoreSpacesNO = "";
                                                                            _WebSocket.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                                            #endregion

                                                                            #region 實體倉不購扣, 扣空倉
                                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                            _WebSocket.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno} 沒計畫量,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                            #endregion
                                                                        }
                                                                        #endregion
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                    }
                                                    else
                                                    {
                                                        bool is_run = true;
                                                        #region 原物料 扣庫存帳 TotalStock,TotalStockII
                                                        int tmp_int = (int.Parse(d["BOMQTY"].ToString()) * outQTY);
                                                        #region 先檢查是否已有單據, 且已移轉多少量
                                                        int detailQTY = tmp_int;
                                                        int stockQTY = 0;
                                                        DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO!=''");
                                                        if (tmp != null)
                                                        {
                                                            docNumberNO = tmp["DOCNumberNO"].ToString();
                                                            detailQTY += (int.Parse(tmp["Detail_QTY"].ToString()) + int.Parse(tmp["Next_StationQTY"].ToString()) + int.Parse(tmp["Next_StoreQTY"].ToString()));
                                                            tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) as qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{docNumberNO}'");
                                                            if (tmp != null && !tmp.IsNull("qty"))
                                                            {
                                                                stockQTY = int.Parse(tmp["qty"].ToString());
                                                            }
                                                            if ((stockQTY - detailQTY) >= 0)
                                                            {
                                                                is_run = false;
                                                            }
                                                            else
                                                            {
                                                                tmp_int = detailQTY - stockQTY;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) as qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}'");
                                                            if (tmp != null && !tmp.IsNull("qty"))
                                                            {
                                                                stockQTY = int.Parse(tmp["qty"].ToString());
                                                            }
                                                            if ((stockQTY - detailQTY) >= 0)
                                                            {
                                                                is_run = false;
                                                                tmp = db.DB_GetFirstDataByDataRow($"select top 1 DOCNumberNO from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}'");
                                                                if (tmp != null) { docNumberNO = tmp["DOCNumberNO"].ToString(); }
                                                            }
                                                            else
                                                            {
                                                                tmp_int = detailQTY - stockQTY;
                                                            }
                                                        }
                                                        #endregion

                                                        if (tmp_int <= 0) { is_run = false; }
                                                        int in_APS_PartNOTimeNote_QTY = tmp_int;

                                                        if (is_run)
                                                        {
                                                            tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 order by b.StoreOrder");
                                                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                                            {
                                                                #region 有計畫Keep量  by StoreOrder順序扣
                                                                foreach (DataRow d2 in tmp_dt.Rows)
                                                                {
                                                                    if (tmp_int > 0)
                                                                    {
                                                                        if (int.Parse(d2["KeepQTY"].ToString()) >= tmp_int)
                                                                        {
                                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY-={tmp_int} where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                            _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                            tmp_int = 0;
                                                                            break;
                                                                        }
                                                                        else
                                                                        {
                                                                            int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                            _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                            tmp_int -= tmp_01;
                                                                        }
                                                                    }
                                                                }
                                                                if (tmp_int > 0)
                                                                {
                                                                    #region 有計畫量不夠扣, 扣實體倉
                                                                    DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                                    if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                                    {
                                                                        foreach (DataRow d2 in tmp_dt2.Rows)
                                                                        {
                                                                            if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                            {
                                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno} 有計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                                tmp_int = 0;
                                                                                break;
                                                                            }
                                                                            else
                                                                            {
                                                                                int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                                if (tmp_01 != 0)
                                                                                {
                                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                    _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno} 有計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                                    tmp_int -= tmp_01;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                    #endregion

                                                                    #region 實體倉不購扣, 扣空倉
                                                                    if (tmp_int > 0)
                                                                    {
                                                                        #region 查找適合庫儲別
                                                                        string out_StoreNO = "";
                                                                        string out_StoreSpacesNO = "";
                                                                        _WebSocket.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                                        #endregion
                                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                        _WebSocket.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                    }
                                                                    #endregion
                                                                }
                                                                #endregion
                                                            }
                                                            else
                                                            {
                                                                #region 沒計畫量, 扣實體倉
                                                                if (tmp_int > 0)
                                                                {
                                                                    DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                                    if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                                    {
                                                                        foreach (DataRow d2 in tmp_dt2.Rows)
                                                                        {
                                                                            if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                            {
                                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno} 沒計畫量,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                                tmp_int = 0;
                                                                                break;
                                                                            }
                                                                            else
                                                                            {
                                                                                int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                                if (tmp_01 != 0)
                                                                                {
                                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                    _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno} 沒計畫量,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                                    tmp_int -= tmp_01;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                if (tmp_int > 0)
                                                                {
                                                                    #region 實體倉不購扣, 扣空倉
                                                                    #region 查找適合庫儲別
                                                                    string out_StoreNO = "";
                                                                    string out_StoreSpacesNO = "";
                                                                    _WebSocket.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                                    #endregion
                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                    _WebSocket.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno} 沒計畫量,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO, true);
                                                                    #endregion
                                                                }
                                                                #endregion
                                                            }
                                                        }
                                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Next_APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}',DOCNumberNO='{docNumberNO}',Next_StationQTY+={in_APS_PartNOTimeNote_QTY} where SimulationId='{d["SimulationId"].ToString()}'");
                                                        if (d["DOCNumberNO"].ToString().Trim() == "" && docNumberNO != "")
                                                        { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET DOCNumberNO='{docNumberNO}' where SimulationId='{d["SimulationId"].ToString()}'"); }
                                                        #endregion
                                                    }
                                                    #endregion

                                                    #region 修正上階子計畫完成
                                                    if (isNeedQTY_OK && !bool.Parse(d["IsOK"].ToString()))
                                                    {
                                                        if (!d.IsNull("Source_StationNO") && (d["Class"].ToString() == "4" || d["Class"].ToString() == "5"))
                                                        {
                                                            if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsOK='1' where SimulationId='{d["SimulationId"].ToString()}'"))
                                                            {
                                                                db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set ArrivalDate='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where IsOK='0' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=1,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{d["NeedId"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set ArrivalDate='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where IsOK='0' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        }
                                                    }
                                                    #endregion
                                                }
                                            }
                                            #endregion

                                            #region 判斷下一站是否為委外加工
                                            if (_Fun.Config.OutPackStationName == dr_APS_Simulation["Apply_StationNO"].ToString())
                                            {
                                                DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr_APS_Simulation["SimulationId"].ToString()}'");
                                                int docQTY = int.Parse(tmp["Detail_QTY"].ToString());
                                                int changQTY = docQTY;
                                                tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_APS_Simulation["NeedId"].ToString()}' and Source_StationNO='{dr_APS_Simulation["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN={(int.Parse(dr_APS_Simulation["Source_StationNO_IndexSN"].ToString()) + 1).ToString()}");
                                                if (tmp != null)//tmp為下一站的 APS_Simulation
                                                {
                                                    string docNumberNO = tmp["DOCNumberNO"].ToString();
                                                    #region 扣已有單據數量
                                                    DataRow doc_DOC4II = db.DB_GetFirstDataByDataRow($"select sum(QTY) as okQTY from SoftNetMainDB.[dbo].[DOC4ProductionII] where SimulationId='{tmp["SimulationId"].ToString()}'");
                                                    if (doc_DOC4II != null && !doc_DOC4II.IsNull("okQTY") && doc_DOC4II["okQTY"].ToString() != "") { docQTY -= int.Parse(doc_DOC4II["okQTY"].ToString()); }
                                                    #endregion
                                                    if (docQTY > 0)
                                                    {
                                                        string tmp_down_SID = "NULL";
                                                        string tmp_down_Source_StationNO = "NULL";
                                                        //用ArrivalDate網前計算StartTime
                                                        string tmp_down_StartTime = Convert.ToDateTime(dr_APS_Simulation["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss");
                                                        string tmp_down_ArrivalDate = Convert.ToDateTime(tmp["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss");
                                                        DataRow tmp_up = dr_APS_Simulation;
                                                        DataRow tmp_down = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{tmp["NeedId"].ToString()}' and Source_StationNO='{tmp["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN={(int.Parse(tmp["Source_StationNO_IndexSN"].ToString()) + 1).ToString()} and PartSN<{tmp["PartSN"].ToString()} order by PartSN desc");
                                                        if (tmp_down != null)
                                                        {
                                                            tmp_down_SID = $"'{tmp_down["SimulationId"].ToString()}'";
                                                            tmp_down_Source_StationNO = $"'{tmp_down["Source_StationNO"].ToString()}'";
                                                            tmp_down_ArrivalDate = Convert.ToDateTime(tmp_down["StartDate"]).ToString("yyyy-MM-dd HH:mm:ss");
                                                        }
                                                        #region 查找適合廠商
                                                        float price = 0;
                                                        string mFNO = _WebSocket.SelectDOC4ProductionMFNO(db, tmp["PartNO"].ToString(), tmp["SimulationId"].ToString(), in_NO, ref price);
                                                        #endregion
                                                        #region 查找適合入庫儲別
                                                        string in_StoreNO = "";
                                                        string in_StoreSpacesNO = "";
                                                        _WebSocket.SelectINStore(db, tmp["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "PA02",true);
                                                        #endregion

                                                        if (_WebSocket.Create_DOC4stock(db, tmp, mFNO, price, in_StoreNO, in_StoreSpacesNO, "PA02", docQTY, "", "", "工站報工,開下一站委外加工", tmp_down_StartTime, tmp_down_ArrivalDate, br.UserNO.Trim(), ref docNumberNO))
                                                        {
                                                            DataRow tmp_9 = db.DB_GetFirstDataByDataRow($"SELECT PartNO,IS_WorkingPaper,IS_Store_Test FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{tmp["PartNO"].ToString()}'");
                                                            if (bool.Parse(tmp_9["IS_WorkingPaper"].ToString()))
                                                            {
                                                                tmp_9 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where PartNO='{tmp["PartNO"].ToString()}' and WorkType='2' and SimulationId='{tmp["SimulationId"].ToString()}' and MFNO='{mFNO}' and DOCNumberNO='{docNumberNO}'");
                                                                if (tmp_9 == null)
                                                                {
                                                                    sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkingPaper] (ServerId,[Id],[WorkType],[PartNO],[Class],[IsOK],[NeedId],[SimulationId],[UP_SimulationId],[Down_SimulationId],[NeedQTY],[Price],[Unit],[MFNO],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],[APS_StationNO],[APS_StationNO_SID],[StartTime],[ArrivalDate],[EndTime],[UpdateTime],DOCNumberNO)
                                                                VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('P')}','2','{tmp["PartNO"].ToString()}','{tmp["Class"].ToString()}','0','{tmp["NeedId"].ToString()}','{tmp["SimulationId"].ToString()}','{tmp_up["SimulationId"].ToString()}',{tmp_down_SID},{docQTY},{price},'PCS','{mFNO}','{in_StoreNO}','{in_StoreSpacesNO}','','',
                                                                {tmp_down_Source_StationNO},{tmp_down_SID},'{tmp_down_StartTime}','{tmp_down_ArrivalDate}',NULL,'{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{docNumberNO}')";
                                                                    db.DB_SetData(sql);
                                                                    db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsWPaper='1',DOCNumberNO='{docNumberNO}' where SimulationId='{tmp["SimulationId"].ToString()}'");
                                                                }
                                                            }
                                                        }
                                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{docNumberNO}' where SimulationId='{tmp["SimulationId"].ToString()}'");
                                                    }
                                                }
                                                if (isNeedQTY_OK)
                                                {
                                                    DataTable dt_APS_WorkingPaper = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkingPaper] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{dr_APS_Simulation["SimulationId"].ToString()}' and WorkType='2' and IsOK='0'");
                                                    if (dt_APS_WorkingPaper != null && dt_APS_WorkingPaper.Rows.Count > 0)
                                                    {
                                                        foreach (DataRow d2 in dt_APS_WorkingPaper.Rows)
                                                        {
                                                            if (d2.IsNull("StartTime") ||(!d2.IsNull("StartTime") && Convert.ToDateTime(d2["StartTime"])<DateTime.Now))
                                                            {
                                                                db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkingPaper] set StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where Id='{d2["Id"].ToString()}'");
                                                            }
                                                        }
                                                    }
                                                }

                                            }
                                            #endregion

                                            #region 消除 APS_WorkTimeNote 工站負荷
                                            DataRow tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SimulationId}'");
                                            if (tmp_del != null && int.Parse(tmp_del["Time1_C"].ToString()) == 1 && int.Parse(tmp_del["Time2_C"].ToString()) == 0 && int.Parse(tmp_del["Time3_C"].ToString()) == 0 && int.Parse(tmp_del["Time4_C"].ToString()) == 0)
                                            { }
                                            else
                                            {
                                                if (tmp_del != null)
                                                {
                                                    string stationNO_Merge = "";
                                                    int delMath_UseTime = 0; int tmp_ct = 0; int tmp_wt = 0; int tmp_st = 0; int tmp_1 = 0; int tmp_2 = 0; int tmp_3 = 0; int tmp_4 = 0;
                                                    tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SimulationId}'");
                                                    if (tmp_del != null)
                                                    {
                                                        tmp_ct = int.Parse(tmp_del["Math_EfficientCT"].ToString());
                                                        tmp_wt = int.Parse(tmp_del["Math_EfficientWT"].ToString());
                                                        tmp_st = int.Parse(tmp_del["Math_StandardCT"].ToString());
                                                        if ((tmp_ct + tmp_wt) != 0)
                                                        { delMath_UseTime += (tmp_ct + tmp_wt) * (outQTY + failQTY); }
                                                        else if (tmp_st != 0)
                                                        { delMath_UseTime += tmp_st * (outQTY + failQTY); }
                                                        else
                                                        { delMath_UseTime += (int)ct * (outQTY + failQTY); }
                                                        if (!tmp_del.IsNull("StationNO_Merge") && tmp_del["StationNO_Merge"].ToString().Trim() != "")
                                                        {
                                                            stationNO_Merge = tmp_del["StationNO_Merge"].ToString().Trim().Substring(0, tmp_del["StationNO_Merge"].ToString().Trim().Length - 1);
                                                            stationNO_Merge = $" in ('{stationNO_Merge.Replace(",", "','")}')";
                                                        }
                                                    }
                                                    #region 先消自己
                                                    dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SimulationId}' and StationNO='{data[0]}' order by CalendarDate");
                                                    if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                                                    {
                                                        foreach (DataRow d in dt_APS_Simulation.Rows)
                                                        {
                                                            tmp_1 = int.Parse(d["Time1_C"].ToString()); tmp_2 = int.Parse(d["Time2_C"].ToString()); tmp_3 = int.Parse(d["Time3_C"].ToString()); tmp_4 = int.Parse(d["Time4_C"].ToString());
                                                            if (delMath_UseTime > 0)
                                                            {
                                                                if (tmp_1 > 0)
                                                                {
                                                                    if (delMath_UseTime > tmp_1) { delMath_UseTime -= tmp_1; tmp_1 = 0; } else { tmp_1 -= delMath_UseTime; delMath_UseTime = 0; }
                                                                }
                                                                if (delMath_UseTime > 0 && tmp_2 > 0)
                                                                {
                                                                    if (delMath_UseTime > tmp_2) { delMath_UseTime -= tmp_2; tmp_2 = 0; } else { tmp_2 -= delMath_UseTime; delMath_UseTime = 0; }
                                                                }
                                                                if (delMath_UseTime > 0 && tmp_3 > 0)
                                                                {
                                                                    if (delMath_UseTime > tmp_3) { delMath_UseTime -= tmp_3; tmp_3 = 0; } else { tmp_3 -= delMath_UseTime; delMath_UseTime = 0; }
                                                                }
                                                                if (delMath_UseTime > 0 && tmp_4 > 0)
                                                                {
                                                                    if (delMath_UseTime > tmp_4) { delMath_UseTime -= tmp_4; tmp_4 = 0; } else { tmp_4 -= delMath_UseTime; delMath_UseTime = 0; }
                                                                }
                                                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET Time1_C={tmp_1},Time2_C={tmp_2},Time3_C={tmp_3},Time4_C={tmp_4} where Id='{d["Id"].ToString()}'");
                                                            }
                                                            if (delMath_UseTime <= 0) { break; }
                                                        }
                                                    }
                                                    #endregion

                                                    #region 不夠消, 消其他合併站
                                                    if (delMath_UseTime > 0 && stationNO_Merge != "")
                                                    {
                                                        dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SimulationId}' and StationNO {stationNO_Merge} order by CalendarDate");
                                                        if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                                                        {
                                                            foreach (DataRow d in dt_APS_Simulation.Rows)
                                                            {
                                                                tmp_1 = int.Parse(d["Time1_C"].ToString());
                                                                tmp_2 = int.Parse(d["Time2_C"].ToString());
                                                                tmp_3 = int.Parse(d["Time3_C"].ToString());
                                                                tmp_4 = int.Parse(d["Time4_C"].ToString());
                                                                if (delMath_UseTime > 0)
                                                                {
                                                                    if (tmp_1 > 0)
                                                                    {
                                                                        if (delMath_UseTime > tmp_1) { delMath_UseTime -= tmp_1; tmp_1 = 1; } else { tmp_1 -= delMath_UseTime; delMath_UseTime = 0; }
                                                                    }
                                                                    if (delMath_UseTime > 0 && tmp_2 > 0)
                                                                    {
                                                                        if (delMath_UseTime > tmp_2) { delMath_UseTime -= tmp_2; tmp_2 = 1; } else { tmp_2 -= delMath_UseTime; delMath_UseTime = 0; }
                                                                    }
                                                                    if (delMath_UseTime > 0 && tmp_3 > 0)
                                                                    {
                                                                        if (delMath_UseTime > tmp_3) { delMath_UseTime -= tmp_3; tmp_3 = 1; } else { tmp_3 -= delMath_UseTime; delMath_UseTime = 0; }
                                                                    }
                                                                    if (delMath_UseTime > 0 && tmp_4 > 0)
                                                                    {
                                                                        if (delMath_UseTime > tmp_4) { delMath_UseTime -= tmp_4; tmp_4 = 1; } else { tmp_4 -= delMath_UseTime; delMath_UseTime = 0; }
                                                                    }
                                                                    db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET Time1_C={tmp_1},Time2_C={tmp_2},Time3_C={tmp_3},Time4_C={tmp_4} where Id='{d["Id"].ToString()}'");
                                                                }
                                                                if (delMath_UseTime <= 0) { break; }
                                                            }
                                                        }
                                                    }
                                                    #endregion

                                                    tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SimulationId}'");
                                                    if (tmp_del != null && int.Parse(tmp_del["Time1_C"].ToString()) == 0 && int.Parse(tmp_del["Time2_C"].ToString()) == 0 && int.Parse(tmp_del["Time3_C"].ToString()) == 0 && int.Parse(tmp_del["Time4_C"].ToString()) == 0 && int.Parse(tmp_del["NeedQTY"].ToString()) > (int.Parse(tmp_del["Detail_QTY"].ToString()) + int.Parse(tmp_del["Detail_Fail_QTY"].ToString())))
                                                    {
                                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET Time1_C=1 where SimulationId='{keys.SimulationId}'");
                                                    }
                                                }
                                            }
                                            #endregion
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    meg = $"後臺錯誤: {ex.Message}";
                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"後臺錯誤: {ex.Message} {ex.StackTrace}", true);
                                }
                            }
                            */

                            break;
                        case "領料":
                            {
                                List<string> data = StationSetOutStore(keys, ref err);//開立單據
                                if (err != "")
                                { keys.ERRMsg = err; }
                                else
                                {
                                    if (data.Count > 0)
                                    {
                                        string sID = "";
                                        foreach (string s in data)
                                        {
                                            if (sID == "") { sID = $" and Id in ('{s}'"; }
                                            else { sID += $",'{s}'"; }
                                        }
                                        if (sID != "") { sID += ")"; }

                                        DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and OUT_StoreNO!='' {sID} order by OUT_StoreNO,OUT_StoreSpacesNO,PartNO");
                                        if (dt != null && dt.Rows.Count > 0)
                                        {
                                            string nametmp = dt.Rows[0]["OUT_StoreNO"].ToString();
                                            keys.MES_String = $"<p>倉庫編號:{nametmp}<p/><table id='tableRead' class='table table-bordered xg-table' cellspacing='0'>";
                                            keys.MES_String = $"{keys.MES_String}<thead><tr><th>儲位</th><th>料號</th><th>數量</th><th>單位</th><th>單據編號</th></tr></thead><tbody>";
                                            foreach (DataRow dr in dt.Rows)
                                            {
                                                db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' where IsOK='0' and Id='{dr["Id"].ToString()}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and ArrivalDate='{Convert.ToDateTime(dr["ArrivalDate"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'");

                                                if (dr["OUT_StoreNO"].ToString() != nametmp)
                                                {
                                                    keys.MES_String = $"{keys.MES_String}</tbody></table>";
                                                    nametmp = dr["OUT_StoreNO"].ToString();
                                                    keys.MES_String = $"{keys.MES_String}<p>倉庫編號:{nametmp}<p/><table id='tableRead' class='table table-bordered xg-table' cellspacing='0'>";
                                                    keys.MES_String = $"{keys.MES_String}<thead><tr><th>儲位</th><th>料號</th><th>數量</th><th>單位</th><th>單據編號</th></tr></thead><tbody>";
                                                }
                                                keys.MES_String = $"{keys.MES_String}<tr><th>{dr["OUT_StoreSpacesNO"].ToString()}</th><th>{dr["PartNO"].ToString()}</th><th>{dr["QTY"].ToString()}</th><th>{dr["Unit"].ToString()}</th><th>{dr["DOCNumberNO"].ToString()}</th></tr>";
                                            }
                                            keys.MES_String = $"{keys.MES_String}</tbody></table>";
                                            DataRow d2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelWork','智慧領料','{d2["PP_Name"].ToString()}','{keys.Station}','{d2["PartNO"].ToString()}','{d2["OrderNO"].ToString()}','{keys.OPNO}',{keys.IndexSN})");
                                        }
                                    }
                                    else
                                    {
                                        if (keys.SimulationId != "")
                                        {
                                            DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and SimulationId='{keys.SimulationId}' order by OUT_StoreNO,OUT_StoreSpacesNO,PartNO");
                                            if (dt != null && dt.Rows.Count > 0)
                                            {
                                                string nametmp = dt.Rows[0]["OUT_StoreNO"].ToString();
                                                keys.MES_String = $"<p>倉庫編號:{nametmp}<p/><table id='tableRead' class='table table-bordered xg-table' cellspacing='0'>";
                                                keys.MES_String = $"{keys.MES_String}<thead><tr><th>儲位</th><th>料號</th><th>數量</th><th>單位</th><th>單據編號</th></tr></thead><tbody>";
                                                foreach (DataRow dr in dt.Rows)
                                                {
                                                    db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' where IsOK='0' and Id='{dr["Id"].ToString()}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and ArrivalDate='{Convert.ToDateTime(dr["ArrivalDate"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'");
                                                    if (dr["OUT_StoreNO"].ToString() != nametmp)
                                                    {
                                                        keys.MES_String = $"{keys.MES_String}</tbody></table>";
                                                        nametmp = dr["OUT_StoreNO"].ToString();
                                                        keys.MES_String = $"{keys.MES_String}<p>倉庫編號:{nametmp}<p/><table id='tableRead' class='table table-bordered xg-table' cellspacing='0'>";
                                                        keys.MES_String = $"{keys.MES_String}<thead><tr><th>儲位</th><th>料號</th><th>數量</th><th>單位</th><th>單據編號</th></tr></thead><tbody>";
                                                    }
                                                    keys.MES_String = $"{keys.MES_String}<tr><th>{dr["OUT_StoreSpacesNO"].ToString()}</th><th>{dr["PartNO"].ToString()}</th><th>{dr["QTY"].ToString()}</th><th>{dr["Unit"].ToString()}</th><th>{dr["DOCNumberNO"].ToString()}</th></tr>";
                                                }
                                                keys.MES_String = $"{keys.MES_String}</tbody></table>";
                                            }
                                            else
                                            { keys.MES_String = "無料可領....."; }
                                        }
                                        else
                                        { keys.MES_String = "無料可領....."; }
                                    }
                                }
                            }
                            break;
                        case "開工":
                            {
                                err = _SFC_Common.ChangeStatus(keys.Station, $"{keys.Station},1", "LabelWork");//###???ipport暫時放Station ,之後要測試會有何問題
                                if (err == "")
                                {
                                    DataRow dr = db.DB_GetFirstDataByDataRow($"select b.* from [dbo].[Manufacture] as a, SoftNetMainDB.[dbo].[LabelStateINFO] as b where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{keys.Station}' and a.Config_macID=b.macID");
                                    if (dr != null)
                                    {
                                        var ledrgb = "0";
                                        var ledstate = "0";
                                        var json = "";

                                        #region 更新標籤
                                        DataRow dr_Manufacture = db.DB_GetFirstDataByDataRow($"select a.*,b.OrderNO as BWO from SoftNetMainDB.[dbo].[Manufacture] as a,SoftNetMainDB.[dbo].[LabelStateINFO] as b where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{dr["StationNO"].ToString()}' and a.OrderNO!='' and a.PP_Name!='' and a.Config_macID=b.macID");
                                        if (dr_Manufacture != null && dr_Manufacture["OrderNO"].ToString().Trim() != "")
                                        {
                                            string macID = dr_Manufacture["Config_macID"].ToString();
                                            if (macID != "")
                                            {
                                                ledrgb = "ff00";
                                                DataRow totalData = _Fun.GetAvgCTWTandTotalOutput(db, false, dr["OrderNO"].ToString(), dr["StationNO"].ToString(), dr["IndexSN"].ToString());
                                                string dis_DetailQTY = "0";
                                                if (totalData != null)
                                                {
                                                    dis_DetailQTY = totalData["TotalOutput"].ToString().Trim();
                                                }
                                                if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set Ledrgb='{ledrgb}',Ledstate=0 where ServerId='{_Fun.Config.ServerId}' and macID='{macID}'"))
                                                {
                                                    //###???暫時拿掉
                                                    //if (dr_Manufacture["BWO"].ToString().Trim() == "" || dr_Manufacture["OrderNO"].ToString().Trim() != dr_Manufacture["BWO"].ToString().Trim())
                                                    //{
                                                    string simulationId = "";
                                                    DataRow sfcdr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_Manufacture["OrderNO"].ToString().Trim()}'");

                                                    //###???若不是SID數量要查BOM,CT時間,料號,品名
                                                    #region 查有無需求碼
                                                    string partNO = "";
                                                    string partName = "";
                                                    string typevalue = $"0;";
                                                    int ct = 0;
                                                    int num = int.Parse(sfcdr["Quantity"].ToString());
                                                    if (!sfcdr.IsNull("NeedId") && sfcdr["NeedId"].ToString().Trim() != "")
                                                    {
                                                        DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{sfcdr["NeedId"].ToString()}' and SimulationId='{dr_Manufacture["SimulationId"].ToString()}'");
                                                        if (tmp_dr != null)
                                                        {
                                                            simulationId = tmp_dr["SimulationId"].ToString();
                                                            typevalue = $"2;{simulationId}";
                                                            num = int.Parse(tmp_dr["NeedQTY"].ToString()) + int.Parse(tmp_dr["SafeQTY"].ToString());
                                                            ct = int.Parse(tmp_dr["Math_EfficientCT"].ToString());
                                                            partNO = tmp_dr["PartNO"].ToString();
                                                            tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{tmp_dr["PartNO"].ToString()}'");
                                                            if (tmp_dr != null)
                                                            {
                                                                partName = tmp_dr["PartName"].ToString().Replace("\"", "＂").Replace("'", "’");
                                                            }

                                                        }
                                                    }
                                                    else { typevalue = $"1;{dr_Manufacture["OrderNO"].ToString().Trim()}"; }
                                                    #endregion

                                                    
                                                    string isUpdate = "1";
                                                    DataRow dr_LabelStateINFO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{macID}'");
                                                    if (dr_LabelStateINFO != null && macID != "")
                                                    {
                                                        DataRow dr_Staion = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_LabelStateINFO["StationNO"].ToString()}'");
                                                        string tmp_s = $"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{dr_Manufacture["StationNO"].ToString()};{typevalue};{dr_Manufacture["IndexSN"].ToString()}";
                                                        var json1 = "";
                                                        var json_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{sfcdr["OrderNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{partNO}\",\"Text3\":\"品名:\",\"PartName\":\"{partName}\",\"Text4\":\"需求量:\",\"DetailQTY\":\"{dis_DetailQTY}\",\"Text5\":\"CT:\",\"EfficientCT\":\"{ct.ToString()}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                                                        var writeShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{sfcdr["OrderNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{partNO}\",\"Text3\":\"品名:\",\"PartName\":\"{partName}\",\"Text4\":\"需求量:\",\"QTY\":\"{num.ToString()}\",\"Text5\":\"CT:\",\"EfficientCT\":\"{ct.ToString()}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                                                        if (dr_LabelStateINFO["Version"].ToString().Trim() != "" && dr_LabelStateINFO["Version"].ToString().Trim().Substring(0, 2) == "42")
                                                        {
                                                            json_ShowValue = $"{json_ShowValue},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_Staion["StationNO"].ToString()}\",\"text17\":\"{dr_Staion["StationName"].ToString()}\"";
                                                            writeShowValue = $"{writeShowValue},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_Staion["StationNO"].ToString()}\",\"text17\":\"{dr_Staion["StationName"].ToString()}\"";
                                                            json1 = $"\"mac\":\"{macID}\",\"mappingtype\":744,\"styleid\":52,{json_ShowValue}";
                                                        }
                                                        else
                                                        {
                                                            json_ShowValue = $"{json_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                                            writeShowValue = $"{writeShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                                            json1 = $"\"mac\":\"{macID}\",\"mappingtype\":71,\"styleid\":48,{json_ShowValue}";
                                                        }

                                                        if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                        {
                                                            json = $"{json1},\"QTY\":\"{num.ToString()}\",\"ledrgb\":\"{ledrgb}\",\"ledstate\":{ledstate}";
                                                            _Fun.Tag_Write(db,macID, "智慧開工", json);
                                                        }
                                                        else { isUpdate = "0"; }
                                                        if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{writeShowValue}',StationNO='{dr_Manufacture["StationNO"].ToString()}',Type='1',OrderNO='{dr_Manufacture["OrderNO"].ToString()}',IndexSN='{dr_Manufacture["IndexSN"].ToString()}',StoreNO='',StoreSpacesNO='',QTY={dis_DetailQTY},IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{macID}'"))
                                                        {

                                                        }
                                                    }
                                                    
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                    keys.Station_State = "生產中...";
                                }
                                else
                                {
                                    keys.ERRMsg = err;
                                }
                            }
                            break;
                        case "停工":
                            {
                                var ledrgb = "0";
                                DataRow dr_Manufacture = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                err = _SFC_Common.ChangeStatus(keys.Station, $"{keys.Station},2", "LabelWork"); //###???ipport暫時放keys.Station ,之後要測試會有何問題
                                if (err == "")
                                {
                                    keys.Station_State = "停止...";
                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set Ledrgb='{ledrgb}',Ledstate=0 where ServerId='{_Fun.Config.ServerId}' and macID='{dr_Manufacture["Config_MutiWO"].ToString()}'");
                                }
                            }
                            break;
                        case "入料":
                            {
                                StationSetInStore(keys);
                                DataRow d2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelWork','智慧入料','{d2["PP_Name"].ToString()}','{keys.Station}','{d2["PartNO"].ToString()}','{d2["OrderNO"].ToString()}','{keys.OPNO}',{keys.IndexSN})");
                                if (keys.SimulationId != "")
                                {
                                    DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and SimulationId='{keys.SimulationId}' order by OUT_StoreNO,OUT_StoreSpacesNO,PartNO");
                                    if (dt != null && dt.Rows.Count > 0)
                                    {
                                        string nametmp = dt.Rows[0]["OUT_StoreNO"].ToString();
                                        keys.MES_String = $"<p>倉庫編號:{nametmp}<p/><table id='tableRead' class='table table-bordered xg-table' cellspacing='0'>";
                                        keys.MES_String = $"{keys.MES_String}<thead><tr><th>儲位</th><th>料號</th><th>數量</th><th>單位</th><th>單據編號</th></tr></thead><tbody>";
                                        foreach (DataRow dr in dt.Rows)
                                        {
                                            if (dr["OUT_StoreNO"].ToString() != nametmp)
                                            {
                                                keys.MES_String = $"{keys.MES_String}</tbody></table>";
                                                nametmp = dr["OUT_StoreNO"].ToString();
                                                keys.MES_String = $"{keys.MES_String}<p>倉庫編號:{nametmp}<p/><table id='tableRead' class='table table-bordered xg-table' cellspacing='0'>";
                                                keys.MES_String = $"{keys.MES_String}<thead><tr><th>儲位</th><th>料號</th><th>數量</th><th>單位</th><th>單據編號</th></tr></thead><tbody>";
                                            }
                                            keys.MES_String = $"{keys.MES_String}<tr><th>{dr["OUT_StoreSpacesNO"].ToString()}</th><th>{dr["PartNO"].ToString()}</th><th>{dr["QTY"].ToString()}</th><th>{dr["Unit"].ToString()}</th><th>{dr["DOCNumberNO"].ToString()}</th></tr>";
                                        }
                                        keys.MES_String = $"{keys.MES_String}</tbody></table>";
                                    }
                                    else
                                    { keys.MES_String = "無料可入....."; }
                                }
                                else
                                { keys.MES_String = "無料可入....."; }
                            }
                            break;
                        case "關單工站移轉":
                            {
                                if (keys.OrderNO == null || keys.OrderNO == "") { keys.ERRMsg = $"查無工單,無法設定關畢工站."; goto break_FUN; }
                                bool Is_Station_Config_Store_Type = false;
                                bool isLastStation = false;
                                string status = "4";
                                List<string> station_list = new List<string>();
                                List<string> station_list_NO_StationNO_Merge = new List<string>();
                                DataRow dr_M = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                if (dr_M["State"].ToString() == "1") { keys.ERRMsg = $"請先停止生產, 在設定關畢工站."; goto break_FUN; }
                                DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_M["SimulationId"].ToString()}'");
                                string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();
                                string label_WO = dr_M["OrderNO"].ToString().Trim();
                                string needID = "";
                                #region 檢查下一站是否為委外, 若是改停止動作加關閉
                                if (!_Fun.Config.IsOutPackStationStore && dr_M["SimulationId"].ToString() != "")
                                {
                                    DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_M["SimulationId"].ToString()}'");
                                    if (tmp != null)
                                    {
                                        if (_Fun.Config.OutPackStationName == tmp["Apply_StationNO"].ToString()) { Is_Station_Config_Store_Type = true; }
                                    }
                                }
                                #endregion
                                if (Is_Station_Config_Store_Type)
                                {
                                    var ledrgb = "0";
                                    DataRow dr_Manufacture = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                    err = _SFC_Common.ChangeStatus(keys.Station, $"{keys.Station},2", "LabelWork"); //###???ipport暫時放keys.Station ,之後要測試會有何問題
                                    if (err == "")
                                    {
                                        keys.Station_State = "停止...";
                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Label_ProjectType='0',OrderNO='',OP_NO='',Master_PP_Name='',PP_Name='',IndexSN=0,Station_Custom_IndexSN='',StationNO_Custom_DisplayName='',State='4',PartNO='',StartTime=NULL,RemarkTimeS=NULL,RemarkTimeE=NULL,EndTime=NULL where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                        #region 更新關閉的電子Tag
                                        DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                        if (dr_tmp != null)
                                        {
                                            string tmp_s = "";
                                            string tmp_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"\",\"Text2\":\"\",\"PartNO\":\"\",\"Text3\":\"\",\"PartName\":\"\",\"Text4\":\"\",\"QTY\":\"\",\"Text5\":\"\",\"EfficientCT\":\"\",\"Text6\":\"\",\"Rate\":\"\",\"Text7\":\"累計量:\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{keys.Station};0;;0\",\"outtime\":0";
                                            if (dr_tmp["Version"].ToString().Trim() != "" && dr_tmp["Version"].ToString().Trim().Substring(0, 2) == "42")
                                            {
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                                tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{keys.Station}\",\"text17\":\"{dr_tmp["StationName"].ToString()}\"";
                                                tmp_s = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{tmp_ShowValue}";
                                            }
                                            else
                                            {
                                                tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"{keys.Station}\",\"text17\":\"\"";
                                                tmp_s = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{tmp_ShowValue},\"ledrgb\":\"0\",\"ledstate\":0";
                                            }
                                            if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                            {
                                                _Fun.Tag_Write(db,dr_M["Config_macID"].ToString(), "智慧停工", tmp_s);
                                            }
                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{tmp_ShowValue}',Ledrgb='0',Ledstate=0,StationNO='{keys.Station}',Type='1',OrderNO='',IndexSN='',StoreNO='',StoreSpacesNO='',QTY=0,IsUpdate='1' where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                        }
                                        #endregion
                                    }
                                }
                                else
                                {
                                    DataRow dr = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_M["OrderNO"].ToString()}'");//###??? and WOStatus=0  要加是否已關閉判斷
                                    if (dr != null)
                                    {
                                        needID = dr.IsNull("NeedId") ? "" : dr["NeedId"].ToString();
                                        dr = db.DB_GetFirstDataByDataRow($"select a.*,b.IsLastStation FROM SoftNetMainDB.[dbo].[Manufacture] as a, SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] as b where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{keys.Station}' and a.OrderNO='{dr_M["OrderNO"].ToString()}' and a.IndexSN='{dr_M["IndexSN"].ToString()}' and a.StationNO=b.StationNO and a.IndexSN=b.IndexSN and a.OrderNO=b.OrderNO and a.PP_Name=b.PP_Name");
                                        if (dr != null)
                                        {
                                            isLastStation = bool.Parse(dr["IsLastStation"].ToString());
                                            #region 判斷是否最後一站
                                            if (isLastStation)
                                            {
                                                //SFC_Common SFC_FUN = new SFC_Common("1", _Fun.Config.Db);
                                                //bool isRun_PP_ProductProcess_Item = true;
                                                //if (db.DB_GetQueryCount($"SELECT * FROM SoftNetSYSDB.[dbo].PP_WO_Process_Item WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='" + dr_M["OrderNO"].ToString() + "'") > 0)
                                                //{ isRun_PP_ProductProcess_Item = false; }
                                                //DataTable dt_WO_Stations = SFC_FUN.Process_ALLSation_RE_Custom(_Fun.Config.ServerId, "1", _Fun.Config.Db, dr_M["PP_Name"].ToString(), "ORDER BY IndexSN, PP_Name ASC", isRun_PP_ProductProcess_Item, dr_M["OrderNO"].ToString());
                                                //if (dt_WO_Stations != null && dt_WO_Stations.Rows.Count > 0)
                                                //{
                                                //    foreach (DataRow d in dt_WO_Stations.Rows)
                                                //    {
                                                //        station_list.Add($"'{d["Station NO"].ToString()}'");
                                                //    }
                                                //}
                                                //SFC_FUN.Dispose();
                                                DataTable dt_WO_Stations = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and NeedId='{needID}'");
                                                if (dt_WO_Stations != null && dt_WO_Stations.Rows.Count > 0)
                                                {
                                                    foreach (DataRow d in dt_WO_Stations.Rows)
                                                    {
                                                        if (!station_list.Contains(d["Source_StationNO"].ToString()))
                                                        {
                                                            station_list.Add(d["Source_StationNO"].ToString());
                                                            if (!d.IsNull("StationNO_Merge"))
                                                            {
                                                                foreach (string s in d["StationNO_Merge"].ToString().Split(','))
                                                                {
                                                                    if (s.Trim() != "" && !station_list.Contains(s))
                                                                    {
                                                                        station_list.Add(s);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                dt_WO_Stations = db.DB_GetData($"select Apply_StationNO from SoftNetSYSDB.[dbo].APS_Simulation where NeedId='{needID}' and PartSN>=0 group by Apply_StationNO");
                                                if (dt_WO_Stations != null && dt_WO_Stations.Rows.Count > 0)
                                                {
                                                    foreach (DataRow d2 in dt_WO_Stations.Rows)
                                                    {
                                                        station_list_NO_StationNO_Merge.Add($"'{d2["Apply_StationNO"].ToString()}'");
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].APS_Simulation where NeedId='{needID}' and SimulationId='{dr_M["SimulationId"].ToString()}'");
                                                if (dr_tmp != null)
                                                {
                                                    if (!station_list.Contains(dr_tmp["Apply_StationNO"].ToString())) { station_list.Add(dr_tmp["Apply_StationNO"].ToString()); }
                                                    station_list_NO_StationNO_Merge.Add($"'{dr_tmp["Apply_StationNO"].ToString()}'");
                                                }
                                            }
                                            #endregion

                                            //###???以下有改 則RUNTimeServer 干涉關公單ㄝ要改

                                            //###???DOCNO暫時寫死BC01
                                            //###???入庫要考慮倉庫最大安置量
                                            #region 半成品 or 成品入庫 與 餘料入庫 Class='4' or Class='5'
                                            DataTable tmp_dt = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Apply_StationNO in ({string.Join(",", station_list_NO_StationNO_Merge)}) and Apply_PP_Name='{dr_M["PP_Name"].ToString()}' and (Class='4' or Class='5') and Source_StationNO is not null");
                                            if (tmp_dt != null)
                                            {
                                                foreach (DataRow d in tmp_dt.Rows)
                                                {
                                                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                    if (tmp_dr != null)
                                                    {
                                                        #region 多生產或多領入庫
                                                        int qty = int.Parse(tmp_dr["Detail_QTY"].ToString()) - (int.Parse(tmp_dr["Next_StationQTY"].ToString()) + int.Parse(tmp_dr["Next_StoreQTY"].ToString()));
                                                        if (qty > 0)
                                                        {
                                                            string tmp_no = "";
                                                            string in_StoreNO = "";
                                                            string in_StoreSpacesNO = "";
                                                            DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}'");
                                                            if (tmp == null)
                                                            {
                                                                #region 查找適合庫儲別
                                                                _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "BC01");
                                                                #endregion
                                                                _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;ManufactureController", ref tmp_no, br.UserNO, true);
                                                            }
                                                            else
                                                            {
                                                                in_StoreNO = tmp["StoreNO"].ToString();
                                                                in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                                                _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;ManufactureController", ref tmp_no, br.UserNO, true);
                                                            }
                                                            sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{tmp_no}',Next_StoreQTY+={qty} where SimulationId='{d["SimulationId"].ToString()}'";
                                                            db.DB_SetData(sql);
                                                        }
                                                        #endregion
                                                    }
                                                }
                                            }
                                            #endregion

                                            //###???暫時寫死 EB01
                                            #region 非半成品成品 原物料 餘料退回入庫 Class!='4' and Class!='5'
                                            /* 改由RUNTimeService 12 處理
                                            sql = "";
                                            List<string> sID_list = new List<string>();
                                            if (isLastStation)
                                            { sql = $"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and ((Class!='4' and Class!='5') or NoStation='1')"; }
                                            else
                                            {
                                                DataTable dt_tmp = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and Apply_PP_Name='{dr_M["PP_Name"].ToString()}' and ((Class!='4' and Class!='5') or Source_StationNO is null)");
                                                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                                {
                                                    foreach (DataRow d in dt_tmp.Rows)
                                                    { sID_list.Add($"'{d["SimulationId"].ToString()}'"); }
                                                    sql = $"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and SimulationId in ({string.Join(",", sID_list)})";
                                                }
                                            }
                                            DataTable dt_APS_PartNOTimeNote = db.DB_GetData(sql);
                                            if (dt_APS_PartNOTimeNote != null && dt_APS_PartNOTimeNote.Rows.Count > 0)
                                            {
                                                DataRow tmp_dr = null;
                                                int sQTY = 0;
                                                int useQYU = 0;
                                                foreach (DataRow d in dt_APS_PartNOTimeNote.Rows)
                                                {
                                                    tmp_dr = db.DB_GetFirstDataByDataRow($@"SELECT sum(a.QTY) as Total FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                                                            join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.DOCNumberNO,1,4) and b.DOCType='3' and b.ServerId='{_Fun.Config.ServerId}'
                                                                                            where a.SimulationId='{d["SimulationId"].ToString()}'");
                                                    if (tmp_dr != null && !tmp_dr.IsNull("Total"))
                                                    {
                                                        useQYU = int.Parse(d["Detail_QTY"].ToString()) + int.Parse(d["Next_StationQTY"].ToString()) + int.Parse(d["Next_StoreQTY"].ToString());
                                                        sQTY = int.Parse(tmp_dr["Total"].ToString());
                                                        if ((sQTY - useQYU) > 0)
                                                        {
                                                            useQYU = sQTY - useQYU;//退回量
                                                            int wQTY = useQYU;
                                                            tmp_dt = db.DB_GetData($@"SELECT a.*,c.NeedId,c.SimulationDate FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                                                            join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.DOCNumberNO,1,4) and b.DOCType='3' and b.ServerId='{_Fun.Config.ServerId}'
                                                                                            join SoftNetSYSDB.[dbo].[APS_Simulation] as c on c.SimulationId=a.SimulationId
                                                                                            where a.SimulationId='{d["SimulationId"].ToString()}' order by a.IsOK,a.Id,a.OUT_StoreNO,a.OUT_StoreSpacesNO");
                                                            string docNumberNO = "";
                                                            foreach (DataRow d2 in tmp_dt.Rows)
                                                            {
                                                                if ((int.Parse(d2["QTY"].ToString()) - useQYU) > 0)
                                                                {
                                                                    if (!bool.Parse(d2["IsOK"].ToString()))
                                                                    { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] SET QTY-={useQYU.ToString()} where Id='{d2["Id"].ToString()}' and DOCNumberNO='{d2["DOCNumberNO"].ToString()}'"); }
                                                                    else
                                                                    { _WebSocket.Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", useQYU, "", d2["Id"].ToString(), $"生產結束餘料退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;ManufactureController", ref docNumberNO, br.UserNO, true); }
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    if (!bool.Parse(d2["IsOK"].ToString()))
                                                                    { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] SET QTY=0,Remark='生產結束清除用量' where Id='{d2["Id"].ToString()}' and DOCNumberNO='{d2["DOCNumberNO"].ToString()}'"); }
                                                                    else
                                                                    { _WebSocket.Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", int.Parse(d2["QTY"].ToString()), "", d2["Id"].ToString(), $"生產結束餘料退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;ManufactureController", ref docNumberNO, br.UserNO, true); }
                                                                    useQYU -= int.Parse(d2["QTY"].ToString());
                                                                    if (useQYU <= 0) { break; }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            */
                                            #endregion

                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsOK='1' where SimulationId='{dr_M["SimulationId"].ToString()}'");
                                        }
                                    }

                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Label_ProjectType='0',OrderNO='',OP_NO='',Master_PP_Name='',PP_Name='',IndexSN=0,Station_Custom_IndexSN='',StationNO_Custom_DisplayName='',State='4',PartNO='',StartTime=NULL,RemarkTimeS=NULL,RemarkTimeE=NULL,EndTime=NULL where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                    db.DB_SetData(@$"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES 
                                                     ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{needID}','{dr_M["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelWork','智慧關站','{dr_M["PP_Name"].ToString()}','{keys.Station}','{dr_M["PartNO"].ToString()}','{dr_M["OrderNO"].ToString()}','{keys.OPNO}',{keys.IndexSN})");

                                    #region 送Service處理後續
                                    dr = db.DB_GetFirstDataByDataRow($"SELECT RMSName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                    if (isLastStation)
                                    {
                                        status = "5";//關站加關工單

                                        //###??? 要考慮 APS_NeedData與APS_Simulation , 同一個NeedId可能有多個 工單時, 發到Softnet Service的_CloseWO Code會有問題

                                    }
                                    //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                    if (dr != null && _WebSocket.RmsSend(dr["RMSName"].ToString(), 1, $"WebChangeStationStatus,{status},{keys.Station},WEBProg,{keys.Station},{dr_M["OP_NO"].ToString()},{dr_M["OrderNO"].ToString()},{dr_M["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
                                    {
                                        if (needID != "")
                                        {
                                            if (status == "5")
                                            {
                                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{needID}'");
                                                db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needID}'");
                                            }
                                            else if (status == "4")
                                            {
                                                db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needID}' and SimulationId='{dr_M["SimulationId"].ToString()}'");
                                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{needID}' and SimulationId='{dr_M["SimulationId"].ToString()}'");
                                            }
                                        }
                                        keys.SimulationId = "";
                                        keys.OrderNO = "";
                                        keys.StationNO_Custom_DisplayName = "";
                                        keys.IndexSN = "0";
                                        keys.Type = "0";
                                        keys.OKQTY = 0;
                                        keys.FailQTY = 0;
                                        keys.Station_State = "等待工單.";

                                        #region 更新電子Tag
                                        if (dr_M["Config_macID"].ToString().Trim() != "")
                                        {
                                            string isUpdate = "1";
                                            if (!_Fun.Is_Tag_Connect) { isUpdate = "0"; }
                                            if (isLastStation)
                                            {
                                                DataRow dr_tmp = null;
                                                string macID = "";
                                                foreach (string s in station_list)
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{s}'");
                                                    if (dr_tmp != null && dr_tmp["State"].ToString().Trim() != "1" && dr_tmp["Config_macID"].ToString().Trim() != "")
                                                    {
                                                        if (dr_tmp["OrderNO"].ToString().Trim() != "" && label_WO != dr_tmp["OrderNO"].ToString().Trim()) { continue; }
                                                        macID = dr_tmp["Config_macID"].ToString().Trim();
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and Type='1' and macID='{macID}'");
                                                        if (dr_tmp != null)
                                                        {
                                                            string tmp_s = "";
                                                            string tmp_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"\",\"Text2\":\"\",\"PartNO\":\"\",\"Text3\":\"\",\"PartName\":\"\",\"Text4\":\"\",\"QTY\":\"\",\"Text5\":\"\",\"EfficientCT\":\"\",\"Text6\":\"\",\"Rate\":\"\",\"Text7\":\"累計量:\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{s};0;;0\",\"outtime\":0";
                                                            if (dr_tmp["Version"].ToString().Trim() != "" && dr_tmp["Version"].ToString().Trim().Substring(0, 2) == "42")
                                                            {
                                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{s}'");
                                                                tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{s}\",\"text17\":\"{dr_tmp["StationName"].ToString()}\"";
                                                                tmp_s = $"\"mac\":\"{macID}\",\"mappingtype\":744,\"styleid\":52,{tmp_ShowValue}";
                                                            }
                                                            else
                                                            {
                                                                tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"{s}\",\"text17\":\"\"";
                                                                tmp_s = $"\"mac\":\"{macID}\",\"mappingtype\":71,\"styleid\":48,{tmp_ShowValue},\"ledrgb\":\"0\",\"ledstate\":0";
                                                            }
                                                            if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                            {
                                                                _Fun.Tag_Write(db,macID, "智慧關站", tmp_s);
                                                            }
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{tmp_ShowValue}',Ledrgb='0',Ledstate=0,StationNO='{s}',Type='1',OrderNO='',IndexSN='',StoreNO='',StoreSpacesNO='',QTY=0,IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{macID}'");
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                                if (dr_tmp != null)
                                                {
                                                    string tmp_s = "";
                                                    string tmp_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"\",\"Text2\":\"\",\"PartNO\":\"\",\"Text3\":\"\",\"PartName\":\"\",\"Text4\":\"\",\"QTY\":\"\",\"Text5\":\"\",\"EfficientCT\":\"\",\"Text6\":\"\",\"Rate\":\"\",\"Text7\":\"累計量:\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{keys.Station};0;;0\",\"outtime\":0";
                                                    if (dr_tmp["Version"].ToString().Trim() != "" && dr_tmp["Version"].ToString().Trim().Substring(0, 2) == "42")
                                                    {
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                                        tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{keys.Station}\",\"text17\":\"{dr_tmp["StationName"].ToString()}\"";
                                                        tmp_s = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{tmp_ShowValue}";
                                                    }
                                                    else
                                                    {
                                                        tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"{keys.Station}\",\"text17\":\"\"";
                                                        tmp_s = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{tmp_ShowValue},\"ledrgb\":\"0\",\"ledstate\":0";
                                                    }
                                                    if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                    {
                                                        _Fun.Tag_Write(db,dr_M["Config_macID"].ToString(),"智慧關站", tmp_s);
                                                    }
                                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{tmp_ShowValue}',Ledrgb='0',Ledstate=0,StationNO='{keys.Station}',Type='1',OrderNO='',IndexSN='',StoreNO='',StoreSpacesNO='',QTY=0,IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                                }
                                            }
                                        }
                                        #endregion

                                    }
                                    else
                                    {
                                        keys.ERRMsg = $"後台服務無作用,請檢查服務是否不正常."; goto break_FUN;
                                    }
                                    #endregion
                                }
                            }
                            break;
                        case "ESOP":
                            {
                                if (keys.SimulationId == null || keys.SimulationId == "") { keys.ERRMsg = $"此製程無ESOP資料."; }
                                else
                                {
                                    ViewBag.ESOP_Files = "";
                                    ViewBag.ESOP_PDF_Files = "";
                                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{keys.SimulationId}'");
                                    if (tmp_dr != null)
                                    {
                                        tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where Apply_PP_Name='{tmp_dr["Apply_PP_Name"].ToString()}' and Apply_StationNO='{tmp_dr["Source_StationNO"].ToString()}' and IndexSN={tmp_dr["Source_StationNO_IndexSN"].ToString()}");
                                        if (tmp_dr != null && !tmp_dr.IsNull("Apply_BOMId") && tmp_dr["Apply_BOMId"].ToString() != "")
                                        {
                                            tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where Id='{tmp_dr["Apply_BOMId"].ToString()}'");
                                            if (tmp_dr != null)
                                            {
                                                if (!tmp_dr.IsNull("Station_DIS_Remark") && tmp_dr["Station_DIS_Remark"].ToString() != "")
                                                { keys.MES_String = tmp_dr["Station_DIS_Remark"].ToString(); }
                                                if (!tmp_dr.IsNull("ESOP_Files") && tmp_dr["ESOP_Files"].ToString() != "")
                                                {
                                                    string key_file = "";
                                                    string key_pdffile = "";
                                                    string partESOPName = "";
                                                    foreach (string s in tmp_dr["ESOP_Files"].ToString().Split(";"))
                                                    {
                                                        if (tmp_dr["StationNO_Custom_DisplayName"].ToString().Trim() != "")
                                                        { partESOPName = tmp_dr["StationNO_Custom_DisplayName"].ToString().Trim(); }
                                                        else if (tmp_dr["Station_Custom_IndexSN"].ToString().Trim() != "")
                                                        { partESOPName = tmp_dr["Station_Custom_IndexSN"].ToString().Trim(); }
                                                        else { partESOPName = tmp_dr["IndexSN"].ToString().Trim(); }
                                                        partESOPName = partESOPName.Replace("/", "／").Replace("\\", "＼").Replace(":", "：").Replace("*", "＊").Replace("?", "？").Replace("\"", "＂").Replace("<", "＜").Replace(">", "＞").Replace("|", "｜");
                                                        if (s.Trim() != "")
                                                        {
                                                            if (s.Split('.')[(s.Split('.').Length - 1)] == "pdf")
                                                            {
                                                                if (key_pdffile == "") { key_pdffile = $"{tmp_dr["PartNO"].ToString().Trim()}/{tmp_dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}/{s}"; }
                                                                else { key_pdffile = $"{key_pdffile};{tmp_dr["PartNO"].ToString().Trim()}/{tmp_dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}/{s}"; }

                                                            }
                                                            else
                                                            {
                                                                if (key_file == "") { key_file = $"{tmp_dr["PartNO"].ToString().Trim()}/{tmp_dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}/{s}"; }
                                                                else { key_file = $"{key_file};{tmp_dr["PartNO"].ToString().Trim()}/{tmp_dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}/{s}"; }
                                                            }
                                                        }
                                                    }
                                                    ViewBag.ESOP_Files = key_file;
                                                    ViewBag.ESOP_PDF_Files = key_pdffile;
                                                }
                                            }
                                            else { keys.ERRMsg = "此製程無ESOP資料."; }
                                        }
                                        else { keys.ERRMsg = "此製程無ESOP資料."; }
                                    }
                                    else { keys.ERRMsg = "此製程無ESOP資料."; }
                                }
                                ViewBag.Model = keys;

                                return View("DisplayESOP",keys);
                            }
                        case "Knives":
                            {
                                string meg = "查無刀具或製工具使用歷程";
                                DataTable dt = db.DB_GetData($@"select b.* from SoftNetMainDB.[dbo].[TotalStockII_Knives] as a
                                                join SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] as b on a.KId=b.KId
                                                where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{keys.Station}' and a.IsDel='0' order by KId,LOGDateTime desc,PartNO");
                                if (dt != null && dt.Rows.Count > 0)
                                {
                                    DataRow tmp_dr = null;
                                    string tmp_09 = "";
                                    List<string> kId_List = new List<string>();
                                    meg = $"<div><table data-role='table' data-mode='columntoggle' class='ui-responsive ui-shadow' id='DisplayDataTable_E' border='1'><thead><tr><th>KId</th><th>歷程日期</th><th>歷程料件編號</th><th>品名</th><th>規格</th><th>單次生產量</th><th>單次時數</th></tr></thead><tbody>";
                                    foreach (DataRow dr in dt.Rows)
                                    {
                                        tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT *  FROM SoftNetMainDB.[dbo].TotalStockII_Knives where KId='{dr["KId"].ToString()}'");
                                        if (tmp_dr != null && !tmp_dr.IsNull("RecoverTime"))
                                        {
                                            if (Convert.ToDateTime(tmp_dr["RecoverTime"]) >= Convert.ToDateTime(dr["LOGDateTime"])) { continue; }
                                        }
                                        if (!kId_List.Contains(dr["KId"].ToString())) { kId_List.Add(dr["KId"].ToString()); }
                                        TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(dr["WorkTime"].ToString()));
                                        tmp_09 = $"{(int)standardTime_DIS.TotalHours}時{standardTime_DIS.Minutes}分{standardTime_DIS.Seconds}秒";
                                        tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                                        if (tmp_dr != null)
                                        { meg = $"{meg}<tr><td>{dr["KId"].ToString()}</td><td>{dr["LOGDateTime"].ToString()}</td><td>{dr["PartNO"].ToString()}</td><td>{tmp_dr["PartName"].ToString()}</td><td>{tmp_dr["Specification"].ToString()}</td><td>{dr["WorkQTY"].ToString()}</td><td>{tmp_09}</td></tr>"; }
                                        else { meg = $"{meg}<tr><td>{dr["KId"].ToString()}</td><td>{dr["LOGDateTime"].ToString()}</td><td>{dr["PartNO"].ToString()}</td><td></td><td></td><td>{dr["WorkQTY"].ToString()}</td><td>{tmp_09}</td></tr>"; }
                                    }
                                    meg = $"{meg}</tbody></table></div>";
                                    if (kId_List.Count > 0)
                                    {
                                        string tmp_01 = "";
                                        string tmp = $"<div><table data-role='table' data-mode='columntoggle' class='ui-responsive ui-shadow' id='DisplayDataTable_E_KId' border='1'><thead><tr><th>KId</th><th>目前位置</th><th>總使用次數</th><th>總使用時數</th></tr></thead><tbody>"; ;
                                        foreach (string s in kId_List)
                                        {
                                            tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[TotalStockII_Knives] where ServerId='{_Fun.Config.ServerId}' and KId='{s}' and IsDel='0'");
                                            if (tmp_dr != null)
                                            {
                                                TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(tmp_dr["TOTWorkTime"].ToString()));
                                                tmp_09 = $"{(int)standardTime_DIS.TotalHours}時{standardTime_DIS.Minutes}分{standardTime_DIS.Seconds}秒";
                                                if (tmp_dr["StationNO"].ToString().Trim() != "") { tmp_01 = $"在線:{tmp_dr["StationNO"].ToString()}站"; }
                                                else { tmp_01 = $"在庫:{tmp_dr["StoreNO"].ToString()}&nbsp;{tmp_dr["StoreSpacesNO"].ToString()}"; }
                                                tmp = $"{tmp}<tr><td>{s}</td><td>{tmp_01}</td><td>{tmp_dr["TOTCount"].ToString()}</td><td>{tmp_09}</td></tr>";
                                            }
                                        }
                                        tmp = $"{tmp}</tbody></table></div>";
                                        meg = $"<label>歷程明細</label>{tmp}<br />{meg}";
                                    }
                                }
                                keys.MES_String = meg;
                            }
                            break;
                    }
                    if (keys.MES_String == "") { keys.MES_String = $"{keys.State} 完成作業."; }
                }
                catch (Exception ex)
                {
                    ViewBag.ERRMsg = $"程式異常, {ex.Message},  請通知管理者.";
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"LabelWorkController.cs SetAction {keys.State} Exception: {ex.Message} {ex.StackTrace}", true);
                    return View("ResuItTimeOUT",keys);
                }
            }
        break_FUN:
            ViewBag.Model = keys;
            //ViewData.Model = keys;
            return View("ResuItDisplay",keys);
        }

        private List<string> StationSetOutStore(LabelWork keys,ref string  err)//領料
        {
            List<string> re = new List<string>();
            if (keys.Station != "" && (keys.OrderNO != "" || keys.SimulationId != ""))
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    List<string> storeNOs = new List<string>();
                    string in_NO02 = "AC01";//###???暫時寫死
                    string sql = "";

                    string for_Apply_StationNO_BY_Main_Source_StationNO = keys.Station;
                    DataRow dr_WO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{keys.OrderNO}'");
                    if (keys.SimulationId.Trim() != "")
                    {
                        #region 先查有無計畫,且有無單據
                        string docNumberNO = "";
                        DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{keys.SimulationId}'");
                        for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();
                        //本站用料明細
                        DataTable dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_WO["NeedId"].ToString()}' and Apply_PP_Name='{dr_WO["PP_Name"].ToString()}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and IndexSN={keys.IndexSN}");
                        string before_ST = "";
                        string sID = "";//本站用料明細的所有SimulationId ID
                        if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                        {
                            foreach (DataRow d in dt_APS_Simulation.Rows)
                            {
                                if (sID == "") { sID = $"a.SimulationId in ('{d["SimulationId"].ToString()}'"; }
                                else { sID += $",'{d["SimulationId"].ToString()}'"; }
                                if (!d.IsNull("Source_StationNO") && (d["Class"].ToString() == "4" || d["Class"].ToString() == "5"))
                                {
                                    if (before_ST == "") { before_ST = $"SimulationId in ('{d["SimulationId"].ToString()}'"; }
                                    else { before_ST += $",'{d["SimulationId"].ToString()}'"; }
                                }
                            }
                            if (sID != "") { sID += ")"; }
                            if (before_ST != "") { before_ST += ")"; }
                        }
                        if (sID != "")
                        {
                            sql = $@"SELECT a.*,c.NeedId,c.SimulationDate FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a
                                  join SoftNetSYSDB.[dbo].[APS_NeedData] as b on b.Id=a.NeedId and b.State='6' and b.ServerId='{_Fun.Config.ServerId}'
                                  join SoftNetSYSDB.[dbo].APS_Simulation as c on c.SimulationId=a.SimulationId
                                  where {sID} and (a.DOCNumberNO is null or a.DOCNumberNO='')  and ((a.Class!='4' and a.Class!='5') or a.NoStation='1') and a.Class!='7'";
                            DataTable dt_APS_PartNOTimeNote = db.DB_GetData(sql);
                            if (dt_APS_PartNOTimeNote != null && dt_APS_PartNOTimeNote.Rows.Count > 0)
                            {
                                bool is_run = true;
                                int tmp_int = 0;
                                DataRow tmp = null;

                                foreach (DataRow d in dt_APS_PartNOTimeNote.Rows)
                                {
                                    is_run = true;
                                    tmp_int = int.Parse(d["NeedQTY"].ToString());//需求數量
                                    //tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}'");
                                    //if (tmp != null)
                                    //{ tmp_int = int.Parse(tmp["NeedQTY"].ToString()) + int.Parse(tmp["SafeQTY"].ToString()); } //需求數量

                                    #region 先檢查DOC3stockII是否已有單據,若有寫入StartTime
                                    int stockQTY = 0;
                                    tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) as qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO02}'");
                                    if (tmp != null && !tmp.IsNull("qty"))
                                    {
                                        stockQTY = int.Parse(tmp["qty"].ToString());
                                    }
                                    if ((stockQTY - tmp_int) >= 0)
                                    {
                                        is_run = false;
                                        tmp = db.DB_GetFirstDataByDataRow($"select top 1 DOCNumberNO from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO02}'");
                                        if (tmp != null)
                                        {
                                            docNumberNO = tmp["DOCNumberNO"].ToString();
                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where IsOK='0' and SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO02}'");
                                        }
                                    }
                                    else
                                    {
                                        tmp_int -= stockQTY;
                                    }
                                    #endregion
                                    if (is_run && tmp_int>0)
                                    {
                                        //查有無Keep量
                                        DataTable tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 order by b.StoreOrder");
                                        if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                        {
                                            #region 有計畫Keep量  by StoreOrder順序扣
                                            foreach (DataRow d2 in tmp_dt.Rows)
                                            {
                                                if (tmp_int > 0)
                                                {
                                                    if (int.Parse(d2["KeepQTY"].ToString()) >= tmp_int)
                                                    {
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY-={tmp_int} where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                        tmp_int -= tmp_01;
                                                    }
                                                }
                                            }
                                            #endregion

                                            if (tmp_int > 0)
                                            {
                                                #region 計畫量不夠扣, 扣實體倉
                                                DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                {
                                                    foreach (DataRow d2 in tmp_dt2.Rows)
                                                    {
                                                        if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                        {
                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                            _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                            tmp_int = 0;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                            if (tmp_01 != 0)
                                                            {
                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                                tmp_int -= tmp_01;
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion

                                                #region 實體倉不購扣, 扣空倉
                                                if (tmp_int > 0)
                                                {
                                                    #region 查找適合庫儲別
                                                    string out_StoreNO = "";
                                                    string out_StoreSpacesNO = "";
                                                    _SFC_Common.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, "AC01");
                                                    #endregion
                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                    _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                }
                                                #endregion
                                            }
                                        }
                                        else
                                        {
                                            #region 沒keep量, 扣實體倉
                                            DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}'order by QTY desc");
                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                            {
                                                foreach (DataRow d2 in tmp_dt2.Rows)
                                                {
                                                    if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                    {
                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                        if (tmp_01 != 0)
                                                        {
                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                            _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                            tmp_int -= tmp_01;
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion

                                            #region 實體倉不購扣, 扣空倉
                                            if (tmp_int > 0)
                                            {
                                                #region 查找適合庫儲別
                                                string out_StoreNO = "";
                                                string out_StoreSpacesNO = "";
                                                _SFC_Common.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, "AC01");
                                                #endregion
                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                            }
                                            #endregion
                                        }
                                    }
                                    if (docNumberNO != "")
                                    {
                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set DOCNumberNO='{docNumberNO}' where NeedId='{d["NeedId"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{d["NeedId"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='' and ((Class!='4' and Class!='5') or Source_StationNO is null)");
                                        if (tmp != null)
                                        { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set DOCNumberNO='{docNumberNO}' where NeedId='{d["NeedId"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'"); }
                                    }
                                }
                            }
                        }
                        #endregion

                        #region 上一站可能有半成品先入庫, 此時要領出
                        //###??? before_ST 有錯誤
                        if (before_ST != "")
                        {
                            dt_APS_Simulation = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{dr_WO["NeedId"].ToString()}' and (Store_DOCNumberNO is not null or Store_DOCNumberNO !='') and {before_ST}");
                            if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                            {
                                foreach (DataRow d in dt_APS_Simulation.Rows)
                                {
                                    int tmp_int = int.Parse(d["Next_StoreQTY"].ToString());
                                    DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["Store_DOCNumberNO"].ToString()}'");
                                    if (tmp != null)
                                    {
                                        DataRow tmp2 = db.DB_GetFirstDataByDataRow($"select sum(QTY) as qty,StoreNO,StoreSpacesNO from SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and StoreNO='{tmp["IN_StoreNO"].ToString()}' and StoreSpacesNO='{tmp["IN_StoreSpacesNO"].ToString()}' and  PartNO='{tmp["PartNO"].ToString()}' group by StoreNO,StoreSpacesNO");
                                        if (tmp2 != null && int.Parse(tmp2["qty"].ToString()) >= tmp_int)
                                        {
                                            string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                            _SFC_Common.Create_DOC3stock(db, d, tmp2["StoreNO"].ToString(), tmp2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 原再製,領出生產", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                        }
                                        else
                                        {
                                            if (tmp != null) { tmp_int -= int.Parse(tmp["qty"].ToString()); }
                                            #region 扣實體倉
                                            DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                            {
                                                foreach (DataRow d2 in tmp_dt2.Rows)
                                                {
                                                    if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                    {
                                                        string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 原再製,領出生產", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                        if (tmp_01 != 0)
                                                        {
                                                            string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                            _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"{stationno} 原再製,領出生產", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                            tmp_int -= tmp_01;
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion

                                            #region 實體倉不購扣, 扣空倉
                                            if (tmp_int > 0)
                                            {
                                                #region 查找適合庫儲別
                                                string out_StoreNO = "";
                                                string out_StoreSpacesNO = "";
                                                _SFC_Common.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, "AC01");
                                                #endregion    
                                                string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", "AC01", tmp_int, "", "", $"{stationno} 原再製,領出生產", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                            }
                                            #endregion
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        if (sID != "")
                        {
                            dt_APS_Simulation = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] as a where a.IsOK='0' and {sID}");
                            if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                            {
                                foreach (DataRow d in dt_APS_Simulation.Rows)
                                {
                                    if (!storeNOs.Contains(d["OUT_StoreNO"].ToString())) { storeNOs.Add(d["OUT_StoreNO"].ToString()); }
                                    re.Add(d["Id"].ToString());
                                }
                            }
                        }
                    }
                    else
                    {
                        //開工單領料單
                        /*
                        if (keys.OrderNO.Trim() != "")
                        {
                            DataRow dr_PP_WorkOrder = db.DB_GetFirstDataByDataRow($"SELECT *  FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{keys.OrderNO.Trim()}'");
                            if (dr_PP_WorkOrder != null)
                            {
                                DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select a.Id,a.PartNO,a.Apply_PP_Name,a.Apply_StationNO,a.IsEnd,b.Class from SoftNetMainDB.[dbo].[BOM] as a,dbo.[Material] as b where a.Main_Item='0' and b.ServerId='{_Fun.Config.ServerId}' and a.Apply_PP_Name='{dr_PP_WorkOrder["PP_Name"].ToString()}' and a.Apply_PartNO='{dr_PP_WorkOrder["PartNO"].ToString()}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and IndexSN={keys.IndexSN} and a.PartNO=b.PartNO order by EffectiveDate desc");
                                DataTable dt = db.DB_GetData($"select *,'' as NeedId,'' as SimulationId from SoftNetMainDB.[dbo].[BOMII] where BOMId='{dr_tmp["Id"].ToString()}' order by sn");
                                if (dt != null && dt.Rows.Count > 0)
                                {
                                    string docNumberNO = "";
                                    int tmp_int = 0;
                                    int isARGs10_offset = 15;//###??? 10將來改參數
                                    string today = DateTime.Now.AddMinutes(isARGs10_offset).ToString("yyyy-MM-dd HH:mm:ss.fff");
                                    foreach (DataRow d in dt.Rows)
                                    {
                                        tmp_int = int.Parse(dr_PP_WorkOrder["Quantity"].ToString()) * int.Parse(d["BOMQTY"].ToString());
                                        sql = $@"select a.StoreNO,a.StoreSpacesNO,sum(a.QTY)-sum(b.KeepQTY) as TotQTY from SoftNetMainDB.[dbo].[TotalStock] as a
                                            join SoftNetMainDB.[dbo].[TotalStockII] as b on a.Id=b.Id
                                            where a.ServerId='{_Fun.Config.ServerId}' and A.PartNO='{d["PartNO"].ToString()}' group by a.StoreNO,a.StoreSpacesNO order by TotQTY desc";
                                        DataTable dt_TotalStock = db.DB_GetData(sql);
                                        if (dt_TotalStock != null && dt_TotalStock.Rows.Count > 0)
                                        {
                                            foreach (DataRow d2 in dt_TotalStock.Rows)
                                            {
                                                if (!d2.IsNull("TotQTY") && int.Parse(d2["TotQTY"].ToString()) > 0)
                                                {
                                                    if (int.Parse(d2["TotQTY"].ToString()) >= tmp_int)
                                                    {
                                                        _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"工單:{keys.OrderNO.Trim()}", today, "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                    }
                                                    else
                                                    {
                                                        tmp_int -= int.Parse(d2["TotQTY"].ToString());
                                                        _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", int.Parse(d2["TotQTY"].ToString()), "", "", $"工單:{keys.OrderNO.Trim()}", today, "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (docNumberNO != "")
                                    {
                                        DataTable dt_tmp = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and DOCNumberNO='{docNumberNO}'");
                                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                        {
                                            foreach (DataRow d in dt_tmp.Rows)
                                            {
                                                if (!storeNOs.Contains(d["OUT_StoreNO"].ToString())) { storeNOs.Add(d["OUT_StoreNO"].ToString()); }
                                                re.Add(d["Id"].ToString());
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        */
                    }
                    DataRow dr_storeNO = null;
                    foreach (string s in storeNOs)
                    {
                        dr_storeNO = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{s}' and Config_macID!=''");
                        if (dr_storeNO != null)
                        {
                            db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and Keys='領料_{keys.OPNO}' and StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and StoreNO='{s}'");
                            db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[BarCode_TMP] (ServerId,[Keys],[StationNO],[StoreNO],[Value],FailTime) VALUES ('{_Fun.Config.ServerId}','領料_{keys.OPNO}','{for_Apply_StationNO_BY_Main_Source_StationNO}','{s}','{keys.OrderNO},{keys.IndexSN},{keys.SimulationId}','{DateTime.Now.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss")}')");
                        }
                    }
                }
            }
            return re;
        }

        private void StationSetInStore(LabelWork keys)//入料
        {
            string sql = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                List<string> storeNOs = new List<string>();
                //###???要判斷是否重複刷8code
                if (keys.SimulationId.Trim() != "")
                {
                    DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{keys.SimulationId}'");
                    string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();
                    DataTable dt_APS_PartNOTimeNote = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and SimulationId='{keys.SimulationId}'");
                    if (dt_APS_PartNOTimeNote != null && dt_APS_PartNOTimeNote.Rows.Count > 0)
                    {
                        string in_NO = "BC01";//###??? 暫時寫死單別
                        foreach (DataRow d in dt_APS_PartNOTimeNote.Rows)
                        {
                            int qty = int.Parse(d["Detail_QTY"].ToString()) - int.Parse(d["Next_StationQTY"].ToString());
                            string tmp_no = "";
                            string in_StoreNO = "";
                            string in_StoreSpacesNO = "";
                            string returnID = "";
                            DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}'");
                            if (tmp == null)
                            {
                                #region 查找適合庫儲別
                                _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "BC01");
                                #endregion
                                _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", ref returnID, "工站生產入庫", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetInStore;LabelWorkController", ref tmp_no, keys.OPNO, true);
                            }
                            else
                            {
                                in_StoreNO = tmp["StoreNO"].ToString();
                                in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", ref returnID, "工站生產入庫", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetInStore;LabelWorkController", ref tmp_no, keys.OPNO, true);
                            }
                            if (!storeNOs.Contains(in_StoreNO)) { storeNOs.Add(in_StoreNO); }
                            if (returnID != "" && tmp_no != "")
                            { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' where Id='{returnID}' and DOCNumberNO='{tmp_no}'"); }
                            sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{tmp_no}',Next_StoreQTY+={qty} where SimulationId='{d["SimulationId"].ToString()}'";
                            db.DB_SetData(sql);
                        }
                    }
                }
                else
                {
                    //###??? 無需求碼 查SFC_StationDetail
                }
				DataRow dr_storeNO = null;
                foreach (string s in storeNOs)
                {
                    dr_storeNO = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{s}' and Config_macID!=''");
                    if (dr_storeNO != null)
                    {
                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and Keys='入料_{keys.OPNO}' and StationNO='{keys.Station}' and StoreNO='{s}'");
                        db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[BarCode_TMP] (ServerId,[Keys],[StationNO],[StoreNO],[Value],FailTime) VALUES ('{_Fun.Config.ServerId}','入料_{keys.OPNO}','{keys.Station}','{s}','{keys.OrderNO},{keys.IndexSN},{keys.SimulationId}','{DateTime.Now.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss")}')");
                    }
                }
            }
        }


        [HttpPost]
        public string SET_Sort_CMD(string key)
        {
            string re = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                try
                {
                    if (key != null && key != "")
                    {
                        string[] data = key.Split(',');
                        if (data.Length > 2)
                        {
                            string orderCMD = "";
                            switch (data[1])
                            {
                                case "PartNO":
                                    {
                                        if (data[2] == "0") { orderCMD = "order by PartNO,OrderNO,RemarkTimeS desc"; }
                                        else { orderCMD = "order by PartNO desc,OrderNO,RemarkTimeS desc"; }

                                        #region 回傳目前工站既有資訊
                                        DataTable dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[0]}' and EndTime is NULL {orderCMD}");
                                        if (dt != null && dt.Rows.Count > 0)
                                        {
                                            re = "";
                                            DataRow tmp_dr2 = null;
                                            int okQTY = 0;//本站
                                            int failedQty = 0;//本站
                                            int t_okQTY = 0;//其他站
                                            int t_failedQty = 0;//其他站
                                            int totCT = 0;//本站實際CT
                                            string state = "";
                                            string partNO = "";
                                            string partName = "";
                                            string specification = "";
                                            string remarkTimeS = "";
                                            string remarkTimeE = "";
                                            string source_StationNO_Custom_IndexSN = "";
                                            string source_StationNO_Custom_DisplayName = "";
                                            string eCT = "0";
                                            DataTable tmp_dt = null;
                                            string tmp_Index = "";
                                            string fontClass = "";
                                            foreach (DataRow dr in dt.Rows)
                                            {
                                                okQTY = 0; failedQty = 0; totCT = 0; t_okQTY = 0; t_failedQty = 0; partNO = ""; partName = ""; specification = ""; state = ""; source_StationNO_Custom_IndexSN = ""; source_StationNO_Custom_DisplayName = "";

                                                tmp_dt = db.DB_GetData($@"SELECT sum(ProductFinishedQty) as OKQTY,sum(ProductFailedQty) as FailedQty,sum(CycleTime)*(sum(ProductFinishedQty)+sum(ProductFailedQty)) as TOTCT FROM SoftNetLogDB.[dbo].[SFC_StationDetail] where 
                                                                        ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and PP_Name='{dr["PP_Name"].ToString()}' and OrderNO='{dr["OrderNO"].ToString()}' and IndexSN={dr["IndexSN"].ToString()} group by OP_NO");
                                                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                                {
                                                    foreach (DataRow dr2 in tmp_dt.Rows)
                                                    {
                                                        okQTY += dr2.IsNull("OKQTY") ? 0 : int.Parse(dr2["OKQTY"].ToString());
                                                        failedQty += dr2.IsNull("FailedQty") ? 0 : int.Parse(dr2["FailedQty"].ToString());
                                                        totCT += dr2.IsNull("TOTCT") ? 0 : int.Parse(dr2["TOTCT"].ToString());
                                                    }
                                                    if (totCT != 0) { totCT = totCT / (okQTY + failedQty); }
                                                }

                                                #region 累計合併站
                                                tmp_dr2 = db.DB_GetFirstDataByDataRow($"select sum(TotalOutput) as TOKQTY,sum(TotalFail) as TFailedQty from SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{dr["StationNO"].ToString()}' and OrderNO='{dr["OrderNO"].ToString()}' and IndexSN_Merge='1' and PP_Name='{dr["PP_Name"].ToString()}' and IndexSN={dr["IndexSN"].ToString()}");
                                                if (tmp_dr2 != null)
                                                {
                                                    t_okQTY = tmp_dr2.IsNull("TOKQTY") ? 0 : int.Parse(tmp_dr2["TOKQTY"].ToString());
                                                    t_failedQty = tmp_dr2.IsNull("TFailedQty") ? 0 : int.Parse(tmp_dr2["TFailedQty"].ToString());
                                                }
                                                #endregion

                                                tmp_dr2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                                                if (tmp_dr2 != null)
                                                {
                                                    partNO = tmp_dr2["PartNO"].ToString();
                                                    partName = tmp_dr2["PartName"].ToString();
                                                    specification = tmp_dr2["Specification"].ToString();
                                                }
                                                remarkTimeS = dr.IsNull("RemarkTimeS") ? "" : Convert.ToDateTime(dr["RemarkTimeS"]).ToString("MM-dd HH:mm");
                                                remarkTimeE = dr.IsNull("RemarkTimeE") ? "" : Convert.ToDateTime(dr["RemarkTimeE"]).ToString("MM-dd HH:mm");
                                                if (remarkTimeS != "" && remarkTimeE == "") { state = "生產"; }
                                                else if (remarkTimeE != "" && !dr.IsNull("StartTime")) { state = "停止"; }

                                                tmp_dr2 = db.DB_GetFirstDataByDataRow($"select AVG(EfficientCycleTime) as ECT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}' and StationNO='{dr["StationNO"].ToString()}' and PP_Name='{dr["PP_Name"].ToString()}' and IndexSN={dr["IndexSN"].ToString()} and DOCNO=''");
                                                if (tmp_dr2 != null && !tmp_dr2.IsNull("ECT"))
                                                {
                                                    eCT = tmp_dr2["ECT"].ToString();
                                                }
                                                tmp_dr2 = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                                                if (tmp_dr2 != null)
                                                {
                                                    source_StationNO_Custom_IndexSN = tmp_dr2["Source_StationNO_Custom_IndexSN"].ToString();
                                                    source_StationNO_Custom_DisplayName = tmp_dr2["Source_StationNO_Custom_DisplayName"].ToString();
                                                }
                                                #region 計算前站完成量
                                                string next_StationQTY = "0";
                                                DataRow dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                                                if (dr3 != null)
                                                {
                                                    dr3 = db.DB_GetFirstDataByDataRow($"select b.* from SoftNetSYSDB.[dbo].[APS_Simulation] as a,SoftNetSYSDB.[dbo].APS_PartNOTimeNote as b where a.Source_StationNO is not NULL and a.NeedId='{dr3["NeedId"].ToString()}' and a.Source_StationNO_IndexSN={(int.Parse(dr3["IndexSN"].ToString()) - 1).ToString()} and a.Apply_StationNO='{dr3["Source_StationNO"].ToString()}' and a.SimulationId=b.SimulationId");
                                                    if (dr3 != null)
                                                    {
                                                        next_StationQTY = dr3["Detail_QTY"].ToString();
                                                    }
                                                }
                                                #endregion
                                                if (source_StationNO_Custom_DisplayName != "") { tmp_Index = source_StationNO_Custom_DisplayName; }
                                                else if (source_StationNO_Custom_IndexSN != "") { tmp_Index = source_StationNO_Custom_IndexSN; }
                                                else { tmp_Index = dr["IndexSN"].ToString(); }


                                                re = $"{re}<tr><td><input type='checkbox' name='optradio' value='{dr["Id"].ToString()}'></td>";
                                                fontClass = "stateNOT";
                                                if (state == "生產") { fontClass = "stateGreen"; }
                                                else if (state == "停止") { fontClass = "stateRed"; }
                                                re = $"{re}<td class='{fontClass}'>{state}</td>";
                                                re = $"{re}<td>{partNO}</td><td>{partName}&nbsp;{specification}</td><td>{tmp_Index}</td><td>{dr["OrderNO"].ToString()}</td><td>{dr["PNQTY"].ToString()}</td><td>{okQTY.ToString()}+{t_okQTY.ToString()}</td><td>{failedQty.ToString()}+{t_failedQty.ToString()}</td>";
                                                re = $"{re}<td>{next_StationQTY}</td><td>{totCT.ToString()}</td><td>{eCT}</td><td>{remarkTimeS}</td><td>{remarkTimeE}</td><td>{dr["PP_Name"].ToString()}</td></tr>";
                                            }
                                        }
                                        #endregion

                                    }
                                    break;
                            }
                        }
                    }
                }
                catch
                {
                    re = "";
                }
            }
            return re;
        }


    }
}
