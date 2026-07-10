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
                        //                where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO = '{tmp.Station}' and b.EndTime is NULL and b.EstimatedStartTime<='{DateTime.Now.AddMonths(2).ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by b.EstimatedStartTime";

                        DateTime comp_CalendarDate = DateTime.Now.AddYears(10);//###??? 提早顯示的時間要參數化
                        // 單次查詢取得需求分類：可用＝State=6 且 PauseTime 為 NULL；暫停＝State=P（避免重複跑兩輪相同 SQL）
                        sql = $@"select a.DOCNumberNO,a.NeedId,c.State from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] as a 
                                join SoftNetSYSDB.[dbo].[PP_WorkOrder] as b on a.DOCNumberNO=b.OrderNO and b.EndTime is NULL 
                                join SoftNetSYSDB.[dbo].[APS_NeedData] as c on c.Id=a.NeedId and c.ServerId='{_Fun.Config.ServerId}'
                                and (c.State='6' or c.State='P')
                                where (a.StationNO = '{tmp.Station}' or a.StationNO_Merge like '%{tmp.Station},%') and a.CalendarDate<='{comp_CalendarDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}' and a.DOCNumberNO!='' 
                                and (a.Time1_C!=0 or a.Time2_C!=0 or a.Time3_C!=0 or a.Time4_C!=0) group by a.DOCNumberNO,a.NeedId,c.State";
                        DataTable dt_PP = db.DB_GetData(sql);
                        if (dt_PP != null && dt_PP.Rows.Count > 0)
                        {
                            HashSet<string> pairActive = new HashSet<string>(StringComparer.Ordinal);
                            HashSet<string> pairPaused = new HashSet<string>(StringComparer.Ordinal);
                            foreach (DataRow dr0 in dt_PP.Rows)
                            {
                                string docNo = dr0["DOCNumberNO"].ToString();
                                string needId = dr0["NeedId"].ToString();
                                string needState = dr0["State"].ToString();
                                string pairKey = docNo + "\u001f" + needId;
                                if (needState == "P") { pairPaused.Add(pairKey); }
                                else if (needState == "6") { pairActive.Add(pairKey); }
                            }
                            HashSet<string> docKeysForIn = new HashSet<string>(StringComparer.Ordinal);
                            foreach (DataRow dr0 in dt_PP.Rows)
                            {
                                docKeysForIn.Add(dr0["DOCNumberNO"].ToString());
                            }
                            string inkey = string.Join(",", docKeysForIn.Select(d => $"'{d}'"));
                            Dictionary<string, List<string>> orderOrder = new Dictionary<string, List<string>>();
                            dt_PP = db.DB_GetData($"select DOCNumberNO,SimulationId from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where DOCNumberNO in ({inkey}) and (StationNO = '{tmp.Station}' or StationNO_Merge like '%{tmp.Station},%') and CalendarDate<='{comp_CalendarDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}' group by DOCNumberNO,SimulationId order by DOCNumberNO");
                            if (dt_PP != null && dt_PP.Rows.Count > 0)
                            {
                                foreach (DataRow dr0 in dt_PP.Rows)
                                {
                                    string docKey = dr0["DOCNumberNO"].ToString();
                                    if (orderOrder.ContainsKey(docKey))
                                    {
                                        if (!orderOrder[docKey].Contains(dr0["SimulationId"].ToString())) { orderOrder[docKey].Add(dr0["SimulationId"].ToString()); }
                                    }
                                    else { orderOrder.Add(docKey, new List<string> { dr0["SimulationId"].ToString() }); }
                                }
                            }
                            string sType2 = "";
                            string tmp_Index = "";
                            tmp.HasWO_List = new List<string[]>();
                            tmp.HasWO_List_Paused = new List<string[]>();
                            Dictionary<string, List<string[]>> groupsItemList = new Dictionary<string, List<string[]>>();
                            Dictionary<string, List<string[]>> groupsItemListPaused = new Dictionary<string, List<string[]>>();
                            HashSet<string> activeRowDedup = new HashSet<string>(StringComparer.Ordinal);
                            HashSet<string> pausedRowDedup = new HashSet<string>(StringComparer.Ordinal);
                            Dictionary<string, DataRow> simCache = new Dictionary<string, DataRow>(StringComparer.Ordinal);
                            Dictionary<string, DataRow> materialCache = new Dictionary<string, DataRow>(StringComparer.Ordinal);
                            Dictionary<string, bool> manufactureBoundCache = new Dictionary<string, bool>(StringComparer.Ordinal);
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
                                        if (!simCache.TryGetValue(sId, out dr2))
                                        {
                                            dr2 = db.DB_GetFirstDataByDataRow($"select a.*,(a.NeedQTY+a.SafeQTY) as QTY,b.NeedType,b.CTName from SoftNetSYSDB.[dbo].[APS_Simulation] as a,SoftNetSYSDB.[dbo].[APS_NeedData] as b where a.NeedId=b.Id and a.SimulationId='{sId}'");
                                            if (dr2 != null) { simCache[sId] = dr2; }
                                        }
                                        if (dr2 != null)
                                        {
                                            string pairKey2 = dr2["DOCNumberNO"].ToString() + "\u001f" + dr2["NeedId"].ToString();
                                            bool toActive = pairActive.Contains(pairKey2);
                                            bool toPaused = pairPaused.Contains(pairKey2);
                                            // 同一 Need 同時命中 Active/Paused 時，Paused 優先，避免同筆重複出現在兩區。
                                            if (toActive && toPaused) { toActive = false; }
                                            if (!toActive && !toPaused) { continue; }

                                            if (dr2["NeedType"].ToString() == "1") { sType2 = $"訂單"; }
                                            else if (dr2["NeedType"].ToString() == "2") { sType2 = $"客戶"; }
                                            else if (dr2["NeedType"].ToString() == "5") { sType2 = $"底稿"; }
                                            else { sType2 = $"廠內"; }
                                            string partNoKey = dr2["PartNO"].ToString();
                                            if (!materialCache.TryGetValue(partNoKey, out dr3))
                                            {
                                                dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNoKey}'");
                                                if (dr3 != null) { materialCache[partNoKey] = dr3; }
                                            }
                                            string mfgKey = tmp.Station + "\u001f" + dr2["DOCNumberNO"].ToString() + "\u001f" + dr2["Source_StationNO_IndexSN"].ToString();
                                            if (!manufactureBoundCache.TryGetValue(mfgKey, out bool isBound))
                                            {
                                                isBound = db.DB_GetQueryCount($"SELECT StationNO FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{tmp.Station}' and OrderNO='{dr2["DOCNumberNO"].ToString()}' and IndexSN={dr2["Source_StationNO_IndexSN"].ToString()}") > 0;
                                                manufactureBoundCache[mfgKey] = isBound;
                                            }
                                            if (isBound) { }
                                            else
                                            {
                                                if (dr2["Source_StationNO_Custom_DisplayName"].ToString() != "") { tmp_Index = dr2["Source_StationNO_Custom_DisplayName"].ToString(); }
                                                else if (dr2["Source_StationNO_Custom_IndexSN"].ToString() != "") { tmp_Index = dr2["Source_StationNO_Custom_IndexSN"].ToString(); }
                                                else { tmp_Index = dr2["Source_StationNO_IndexSN"].ToString(); }
                                                int baseQTY = int.Parse(dr2["QTY"].ToString());
                                                int priorFail = _SFC_Common.SumPriorDetailFailQTY(db, dr2["NeedId"].ToString(), dr2["Apply_PP_Name"].ToString(), int.Parse(dr2["IndexSN"].ToString()));
                                                string disQTY = _SFC_Common.GetEffectiveNeedQTY(baseQTY, priorFail).ToString();
                                                var okFail = _SFC_Common.GetStationOkFailDisplay(db, _Fun.Config.ServerId, tmp.Station, dr2["Apply_PP_Name"].ToString(), dr2["DOCNumberNO"].ToString(), dr2["Source_StationNO_IndexSN"].ToString(), sId);
                                                string[] rowArr = new string[] { dr2["DOCNumberNO"].ToString(), dr2["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString(), sType2, dr2["CTName"].ToString(), dr2["Apply_PP_Name"].ToString(), tmp_Index, sId, dr2["Source_StationNO_Custom_IndexSN"].ToString(), dr2["Source_StationNO_Custom_DisplayName"].ToString(), disQTY, dr2["Source_StationNO_IndexSN"].ToString(), okFail.okDisplay, okFail.failDisplay };
                                                string rowKey = dr2["DOCNumberNO"].ToString() + "\u001f" + dr2["SimulationId"].ToString() + "\u001f" + dr2["Source_StationNO_IndexSN"].ToString();
                                                void AddToGroups(Dictionary<string, List<string[]>> gdict)
                                                {
                                                    string pk = dr2["PartNO"].ToString();
                                                    if (!gdict.ContainsKey(pk))
                                                    {
                                                        gdict.Add(pk, new List<string[]> { rowArr });
                                                    }
                                                    else { gdict[pk].Add(rowArr); }
                                                }
                                                if (toActive && activeRowDedup.Add(rowKey)) { AddToGroups(groupsItemList); }
                                                if (toPaused && pausedRowDedup.Add(rowKey)) { AddToGroups(groupsItemListPaused); }
                                            }
                                        }
                                    }
                                }
                            }
                            void FlattenSorted(Dictionary<string, List<string[]>> src, List<string[]> dest)
                            {
                                if (src.Count == 0) { return; }
                                var ordered = src.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value);
                                foreach (KeyValuePair<string, List<string[]>> kv in ordered)
                                {
                                    foreach (string[] ss in kv.Value)
                                    {
                                        dest.Add(ss);
                                    }
                                }
                            }
                            FlattenSorted(groupsItemList, tmp.HasWO_List);
                            FlattenSorted(groupsItemListPaused, tmp.HasWO_List_Paused);
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
                                    {
                                        int baseWoQTY = int.Parse(tmp_dr2["PNQTY"].ToString());
                                        DataRow drSim = db.DB_GetFirstDataByDataRow($"select NeedId,Apply_PP_Name,IndexSN from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_M["SimulationId"].ToString()}'");
                                        if (drSim != null)
                                        {
                                            int priorFail = _SFC_Common.SumPriorDetailFailQTY(db, drSim["NeedId"].ToString(), drSim["Apply_PP_Name"].ToString(), int.Parse(drSim["IndexSN"].ToString()));
                                            baseWoQTY = _SFC_Common.GetEffectiveNeedQTY(baseWoQTY, priorFail);
                                        }
                                        tmp.WO_QTY = baseWoQTY.ToString();
                                    }
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
            { ViewBag.ERRMsg = $"作業失敗, 畫面已逾時, 請關閉網頁瀏覽器, 重新刷條碼 並 重新操作."; return View("ResuItTimeOUT", keys); }

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
                                        keys.IndexSN = tmp_dt.Rows[0]["IndexSN"].ToString();
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
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and PartNO='{beforePartNO}' and IndexSN={beforeIndexSN} and LOGDateTime<='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' and OperateType like '%開工%' order by LOGDateTime desc");
                                        if (dr_tmp != null)
                                        {
                                            DateTime tmp_edate = Convert.ToDateTime(dr_tmp["LOGDateTime"]);
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and PartNO='{beforePartNO}' and IndexSN={beforeIndexSN} and LOGDateTime>='{tmp_edate.ToString("yyyy-MM-dd HH:mm:ss")}' and OperateType like '%報工%'");
                                            if (dr_tmp == null)
                                            {
                                                int isARGs10_offset = 15;//###??? 10將來改參數
                                                if ((_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), tmp_edate.AddMinutes(isARGs10_offset), DateTime.Now)) >= avgReportTime)
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
                                        num = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, sfcdr["NeedId"].ToString(), keys.SimulationId, int.Parse(tmp_dr["NeedQTY"].ToString())) + int.Parse(tmp_dr["SafeQTY"].ToString());
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
                                            num = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, sfcdr["NeedId"].ToString(), keys.SimulationId, int.Parse(tmp_dr["NeedQTY"].ToString()));
                                        }
                                    }
                                }
                                #endregion

                                #region 更新Manufacture 製造現場狀態
                                if (!db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Station_Custom_IndexSN='{stationNO_Custom_IndexSN}',StationNO_Custom_DisplayName='{stationNO_Custom_DisplayName}',StartTime=NULL,RemarkTimeS=NULL,RemarkTimeE=NULL,EndTime=NULL,Label_ProjectType='0',OrderNO='{keys.OrderNO}',IndexSN={keys.IndexSN},OP_NO='{keys.OPNO}',Master_PP_Name='{sfcdr["PP_Name"].ToString()}',PP_Name='{sfcdr["PP_Name"].ToString()}',SimulationId='{keys.SimulationId}',PartNO='{partNO}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'"))
                                {
                                    keys.HasWO_List = new List<string[]>();
                                    keys.HasWO_List_Paused = new List<string[]>();
                                    keys.HasWO_List.Add(new string[] { keys.OPNO, keys.OPNO_Name });
                                    keys.ERRMsg = $"異動工單設定失敗."; goto break_FUN;
                                }
                                else { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{keys.OrderNO}' where SimulationId='{keys.SimulationId}'"); }
                                #endregion

                                db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{sfcdr["NeedId"].ToString()}','{keys.SimulationId}','{DateTime.Now.ToString("MM /dd/yyyy HH:mm:ss.fff")}','LabelWork','設定工單',NULL,'{keys.Station}','{partNO}','{keys.OrderNO}','{key_OPNO}',{keys.IndexSN})");
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
                                        _Fun.Tag_Write(db, dr["Config_macID"].ToString(), "設定工單", $"{json1},\"QTY\":\"{num.ToString()}\",\"ledrgb\":\"0\",\"ledstate\":0");
                                    }
                                    else { isUpdate = "0"; }
                                    if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{writeShowValue}',Ledrgb='0',Ledstate=0,StationNO='{keys.Station}',Type='1',OrderNO='{sfcdr["OrderNO"].ToString()}',IndexSN='{keys.IndexSN}',StoreNO='',StoreSpacesNO='',QTY={dis_DetailQTY},IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr["Config_macID"].ToString()}'"))
                                    {
                                    }
                                    if (isUpdate == "0")
                                    {
                                        keys.ERRMsg = $"{keys.Station} 工站 傳送電子訊號失敗, 請通知管理者"; 
                                    }
                                }
                                #endregion
                                ApsNeedPauseResumeQueue.TryEnqueueFromOrderAndSimulation(db, keys.OrderNO, keys.SimulationId, keys.Station, "LabelWork_異動工單");
                                if (keys.ERRMsg != "") { goto break_FUN; }
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
                                VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','LabelWork','設定人員','{dr["PP_Name"].ToString()}','{keys.Station}','{dr["PartNO"].ToString()}','{dr["OrderNO"].ToString()}','{key_OPNO}',{dr["IndexSN"].ToString()})");
                            }
                            break;
                        case "報工":
                            string message = "";
                            string stackTrace = "";
                            string ViewBagERRMsg = "";
                            if (keys.OKQTY <= 0 && keys.FailQTY <= 0)
                            {
                                keys.ERRMsg = $"報工數量, 不能良品與不良都為 0."; goto break_FUN;
                            }
                            bool is_reportOK = _SFC_Common.Reporting_LabelWork(db, dr_StationNO, br.UserNO, keys, false, ref message, ref stackTrace, ref ViewBagERRMsg);
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
                                ApsNeedPauseResumeQueue.TryEnqueueFromManufactureStation(db, keys.Station, "LabelWork_報工");
                            }

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
                                            else { sID = $"{sID},'{s}'"; }
                                        }
                                        if (sID != "") { sID = $"{sID})"; }

                                        DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and OUT_StoreNO!='' {sID} order by OUT_StoreNO,OUT_StoreSpacesNO,PartNO");
                                        if (dt != null && dt.Rows.Count > 0)
                                        {
                                            string nametmp = dt.Rows[0]["OUT_StoreNO"].ToString();
                                            keys.MES_String = $"<p>倉庫編號:{nametmp}<p/><table id='tableRead' class='table table-bordered xg-table' cellspacing='0'>";
                                            keys.MES_String = $"{keys.MES_String}<thead><tr><th>儲位</th><th>料號</th><th>數量</th><th>單位</th><th>單據編號</th></tr></thead><tbody>";
                                            foreach (DataRow dr in dt.Rows)
                                            {
                                                db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' where IsOK='0' and Id='{dr["Id"].ToString()}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and ArrivalDate='{Convert.ToDateTime(dr["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}'");

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
                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','LabelWork','智慧領料','{d2["PP_Name"].ToString()}','{keys.Station}','{d2["PartNO"].ToString()}','{d2["OrderNO"].ToString()}','{keys.OPNO}',{keys.IndexSN})");
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
                                                    db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' where IsOK='0' and Id='{dr["Id"].ToString()}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and ArrivalDate='{Convert.ToDateTime(dr["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}'");
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
                                    if (keys.ERRMsg == "")
                                    { ApsNeedPauseResumeQueue.TryEnqueueFromManufactureStation(db, keys.Station, "LabelWork_領料"); }
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
                                                            num = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, sfcdr["NeedId"].ToString(), simulationId, int.Parse(tmp_dr["NeedQTY"].ToString())) + int.Parse(tmp_dr["SafeQTY"].ToString());
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
                                                            _Fun.Tag_Write(db, macID, "智慧開工", json);
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
                                    ApsNeedPauseResumeQueue.TryEnqueueFromManufactureStation(db, keys.Station, "LabelWork_開工");
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
                                    ApsNeedPauseResumeQueue.TryEnqueueFromManufactureStation(db, keys.Station, "LabelWork_停工");
                                }
                            }
                            break;
                        case "入料":
                            {
                                StationSetInStore(keys);
                                DataRow d2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                                db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','LabelWork','智慧入料','{d2["PP_Name"].ToString()}','{keys.Station}','{d2["PartNO"].ToString()}','{d2["OrderNO"].ToString()}','{keys.OPNO}',{keys.IndexSN})");
                                ApsNeedPauseResumeQueue.TryEnqueueFromManufactureStation(db, keys.Station, "LabelWork_入料");
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
                                                _Fun.Tag_Write(db, dr_M["Config_macID"].ToString(), "智慧停工", tmp_s);
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
                                                     ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{needID}','{dr_M["SimulationId"].ToString()}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','LabelWork','智慧關站','{dr_M["PP_Name"].ToString()}','{keys.Station}','{dr_M["PartNO"].ToString()}','{dr_M["OrderNO"].ToString()}','{keys.OPNO}',{keys.IndexSN})");

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
                                                                _Fun.Tag_Write(db, macID, "智慧關站", tmp_s);
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
                                                        _Fun.Tag_Write(db, dr_M["Config_macID"].ToString(), "智慧關站", tmp_s);
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
                                                    var sb_key_file = new StringBuilder();
                                                    var sb_key_pdffile = new StringBuilder();
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
                                                            string filePath = $"{tmp_dr["PartNO"].ToString().Trim()}/{tmp_dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}/{s}";
                                                            if (s.Split('.')[(s.Split('.').Length - 1)] == "pdf")
                                                            {
                                                                if (sb_key_pdffile.Length > 0) { sb_key_pdffile.Append(";"); }
                                                                sb_key_pdffile.Append(filePath);
                                                            }
                                                            else
                                                            {
                                                                if (sb_key_file.Length > 0) { sb_key_file.Append(";"); }
                                                                sb_key_file.Append(filePath);
                                                            }
                                                        }
                                                    }
                                                    string key_file = sb_key_file.ToString();
                                                    string key_pdffile = sb_key_pdffile.ToString();
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

                                return View("DisplayESOP", keys);
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
                                    var sb_meg = new StringBuilder();
                                    sb_meg.Append("<div><table data-role='table' data-mode='columntoggle' class='ui-responsive ui-shadow' id='DisplayDataTable_E' border='1'><thead><tr><th>KId</th><th>歷程日期</th><th>歷程料件編號</th><th>品名</th><th>規格</th><th>單次生產量</th><th>單次時數</th></tr></thead><tbody>");
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
                                        { sb_meg.Append($"<tr><td>{dr["KId"].ToString()}</td><td>{dr["LOGDateTime"].ToString()}</td><td>{dr["PartNO"].ToString()}</td><td>{tmp_dr["PartName"].ToString()}</td><td>{tmp_dr["Specification"].ToString()}</td><td>{dr["WorkQTY"].ToString()}</td><td>{tmp_09}</td></tr>"); }
                                        else { sb_meg.Append($"<tr><td>{dr["KId"].ToString()}</td><td>{dr["LOGDateTime"].ToString()}</td><td>{dr["PartNO"].ToString()}</td><td></td><td></td><td>{dr["WorkQTY"].ToString()}</td><td>{tmp_09}</td></tr>"); }
                                    }
                                    sb_meg.Append("</tbody></table></div>");
                                    if (kId_List.Count > 0)
                                    {
                                        string tmp_01 = "";
                                        var sb_tmp = new StringBuilder();
                                        sb_tmp.Append("<div><table data-role='table' data-mode='columntoggle' class='ui-responsive ui-shadow' id='DisplayDataTable_E_KId' border='1'><thead><tr><th>KId</th><th>目前位置</th><th>總使用次數</th><th>總使用時數</th></tr></thead><tbody>");
                                        foreach (string s in kId_List)
                                        {
                                            tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[TotalStockII_Knives] where ServerId='{_Fun.Config.ServerId}' and KId='{s}' and IsDel='0'");
                                            if (tmp_dr != null)
                                            {
                                                TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(tmp_dr["TOTWorkTime"].ToString()));
                                                tmp_09 = $"{(int)standardTime_DIS.TotalHours}時{standardTime_DIS.Minutes}分{standardTime_DIS.Seconds}秒";
                                                if (tmp_dr["StationNO"].ToString().Trim() != "") { tmp_01 = $"在線:{tmp_dr["StationNO"].ToString()}站"; }
                                                else { tmp_01 = $"在庫:{tmp_dr["StoreNO"].ToString()}&nbsp;{tmp_dr["StoreSpacesNO"].ToString()}"; }
                                                sb_tmp.Append($"<tr><td>{s}</td><td>{tmp_01}</td><td>{tmp_dr["TOTCount"].ToString()}</td><td>{tmp_09}</td></tr>");
                                            }
                                        }
                                        sb_tmp.Append("</tbody></table></div>");
                                        string tmp = sb_tmp.ToString();
                                        meg = $"<label>歷程明細</label>{tmp}<br />{sb_meg.ToString()}";
                                    }
                                    else
                                    {
                                        meg = sb_meg.ToString();
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
                    return View("ResuItTimeOUT", keys);
                }
            }
        break_FUN:
            ViewBag.Model = keys;
            //ViewData.Model = keys;
            return View("ResuItDisplay", keys);
        }

        private List<string> StationSetOutStore(LabelWork keys, ref string err)//領料
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
                                else { sID = $"{sID},'{d["SimulationId"].ToString()}'"; }
                                if (!d.IsNull("Source_StationNO") && (d["Class"].ToString() == "4" || d["Class"].ToString() == "5"))
                                {
                                    if (before_ST == "") { before_ST = $"SimulationId in ('{d["SimulationId"].ToString()}'"; }
                                    else { before_ST = $"{before_ST},'{d["SimulationId"].ToString()}'"; }
                                }
                            }
                            if (sID != "") { sID = $"{sID})"; }
                            if (before_ST != "") { before_ST = $"{before_ST})"; }
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
                                    tmp_int = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, d["NeedId"].ToString(), d["SimulationId"].ToString(), int.Parse(d["NeedQTY"].ToString()));//有效需求量
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
                                    if (is_run && tmp_int > 0)
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
                                                        string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
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
                                                            string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                            _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                            tmp_int = 0;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                            if (tmp_01 > 0)
                                                            {
                                                                string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
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
                                                    string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
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
                                                        string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, keys.OPNO, true);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                        if (tmp_01 > 0)
                                                        {
                                                            string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
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
                                                string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
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
                                                        if (tmp_01 > 0)
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
                            { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' where Id='{returnID}' and DOCNumberNO='{tmp_no}'"); }
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
                                            var sb = new StringBuilder();
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
                                                tmp_dr2 = db.DB_GetFirstDataByDataRow($"select sum(ProductFinishedQty) as TOKQTY,sum(ProductFailedQty) as TFailedQty from SoftNetLogDB.[dbo].[SFC_StationDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{dr["StationNO"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}' and OrderNO='{dr["OrderNO"].ToString()}' and IndexSN_Merge='1' and PP_Name='{dr["PP_Name"].ToString()}' and IndexSN={dr["IndexSN"].ToString()}");
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

                                                int effectivePNQTY = int.Parse(dr["PNQTY"].ToString());
                                                if (dr["SimulationId"].ToString() != "")
                                                {
                                                    DataRow drNoteForQty = db.DB_GetFirstDataByDataRow($"select NeedQTY as PNQTY FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr["SimulationId"].ToString()}'");
                                                    if (drNoteForQty != null && !drNoteForQty.IsNull("PNQTY"))
                                                    {
                                                        DataRow drSimForQty = db.DB_GetFirstDataByDataRow($"select NeedId from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                                                        if (drSimForQty != null)
                                                        {
                                                            effectivePNQTY = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, drSimForQty["NeedId"].ToString(), dr["SimulationId"].ToString(), int.Parse(drNoteForQty["PNQTY"].ToString()));
                                                        }
                                                    }
                                                }

                                                sb.Append($"<tr><td><input type='checkbox' name='optradio' value='{dr["Id"].ToString()}'></td>");
                                                fontClass = "stateNOT";
                                                if (state == "生產") { fontClass = "stateGreen"; }
                                                else if (state == "停止") { fontClass = "stateRed"; }
                                                sb.Append($"<td class='{fontClass}'>{state}</td>");
                                                sb.Append($"<td>{partNO}</td><td>{partName}&nbsp;{specification}</td><td>{tmp_Index}</td><td>{dr["OrderNO"].ToString()}</td><td>{effectivePNQTY.ToString()}</td><td>{okQTY.ToString()}+{t_okQTY.ToString()}</td><td>{failedQty.ToString()}+{t_failedQty.ToString()}</td>");
                                                sb.Append($"<td>{next_StationQTY}</td><td>{totCT.ToString()}</td><td>{eCT}</td><td>{remarkTimeS}</td><td>{remarkTimeE}</td><td>{dr["PP_Name"].ToString()}</td></tr>");
                                            }
                                            re = sb.ToString();
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

