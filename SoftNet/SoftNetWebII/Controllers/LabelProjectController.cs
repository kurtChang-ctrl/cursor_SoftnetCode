using Base;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore.SqlServer.Extensions.Internal;
using Microsoft.Extensions.Logging;

using SoftNetWebII.Models;
using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;

namespace SoftNetWebII.Controllers
{
    public class LabelProjectController : ApiCtrl
    {
        //private SocketClientService _SNsocket = null;
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        //public LabelProjectController(SocketClientService snsocket, SNWebSocketService websocket)
        public LabelProjectController(SNWebSocketService websocket, SFC_Common sfc_Common)
        {
            //if (_SNsocket == null)
            //{ _SNsocket = snsocket; }
            if (_WebSocket == null)
            { _WebSocket = websocket; }
            if (_SFC_Common == null)
            {
                _SFC_Common = sfc_Common;
            }
        }

        
        public IActionResult Index(string id)//網址參數 工站編號
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            LabelProject tmp = new LabelProject();
            if (id != null)
            {
                string sql = "";
                tmp.OPNO = br.UserNO;
                tmp.OPNO_Name = br.UserName;
                tmp.StationNO = id;
                tmp.Station_State = "無任何設定";

                #region 回傳可用製程名稱,料號, 有註冊的OP
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{tmp.StationNO}'");
                    if (dr_PP_Station == null)
                    { tmp.ERRMsg = $"系統中未建立 [{tmp.StationNO}] 工作站, 請通知管理者."; }
                    else
                    {
                        tmp.StationName = dr_PP_Station["StationName"].ToString();

                        //###???子製程有問題
                        #region 回傳 學習模式可用製程名稱 與 料號
                        sql = @$"select * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and StationNO = '{tmp.StationNO}' Order by PP_Name,IndexSN,Station_Custom_IndexSN";
                        DataTable dt_PP = db.DB_GetData(sql);
                        if (dt_PP != null && dt_PP.Rows.Count > 0)
                        {
                            tmp.HasPP_Name_List = new List<string[]>();
                            DataRow tmp_dr2 = null;
                            DataTable tmp_dt3 = null;
                            List<string> tmp_HasPartNO = new List<string>();
                            tmp.HasPartNO_List = new List<string[]>();

                            foreach (DataRow dr in dt_PP.Rows)
                            {
                                tmp.HasPP_Name_List.Add(new string[] { dr["PP_Name"].ToString(), dr["IndexSN"].ToString(), dr["Station_Custom_IndexSN"].ToString() });
                                #region 依 PP_ProductProcess_Item_PartNO 定義的該製成有哪些料可選用
                                tmp_dr2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item_PartNO] where Id='{dr["Id"].ToString()}'");
                                if (tmp_dr2 != null)
                                {
                                    if (tmp_dr2["By_Class"].ToString().Trim() != "")
                                    {
                                        tmp_dt3 = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and Class='{tmp_dr2["By_Class"].ToString()}' order by PartNO");
                                        if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                        {
                                            foreach (DataRow dr3 in tmp_dt3.Rows)
                                            {
                                                if (!tmp_HasPartNO.Contains(dr3["PartNO"].ToString()))
                                                {
                                                    tmp_HasPartNO.Add(dr3["PartNO"].ToString());
                                                    tmp.HasPartNO_List.Add(new string[] { dr3["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString() });
                                                }
                                            }
                                        }
                                    }
                                    if (tmp_dr2["By_PartType"].ToString().Trim() != "")
                                    {
                                        tmp_dt3 = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartType='{tmp_dr2["By_PartType"].ToString()}' order by PartNO");
                                        if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                        {
                                            foreach (DataRow dr3 in tmp_dt3.Rows)
                                            {
                                                if (!tmp_HasPartNO.Contains(dr3["PartNO"].ToString()))
                                                {
                                                    tmp_HasPartNO.Add(dr3["PartNO"].ToString());
                                                    tmp.HasPartNO_List.Add(new string[] { dr3["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString() });
                                                }
                                            }
                                        }
                                    }
                                    if (tmp_dr2["By_Model"].ToString().Trim() != "")
                                    {
                                        tmp_dt3 = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartType='{tmp_dr2["By_Model"].ToString()}' order by PartNO");
                                        if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                        {
                                            foreach (DataRow dr3 in tmp_dt3.Rows)
                                            {
                                                if (!tmp_HasPartNO.Contains(dr3["PartNO"].ToString()))
                                                {
                                                    tmp_HasPartNO.Add(dr3["PartNO"].ToString());
                                                    tmp.HasPartNO_List.Add(new string[] { dr3["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString() });
                                                }
                                            }
                                        }
                                    }
                                    if (!tmp_dr2.IsNull("Default_Use_PartNOs") && tmp_dr2["Default_Use_PartNOs"].ToString().Trim() != "")
                                    {
                                        List<string> tmp_s = dr["Default_Use_PartNOs"].ToString().Split(";").ToList();
                                        tmp_s.Sort();
                                        foreach (string s in tmp_s)
                                        {
                                            if (s != "")
                                            {
                                                DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{s}'");
                                                if (!tmp_HasPartNO.Contains(s))
                                                {
                                                    tmp_HasPartNO.Add(s);
                                                    tmp.HasPartNO_List.Add(new string[] { tmp_dr["PartNO"].ToString(), tmp_dr["PartName"].ToString(), tmp_dr["Specification"].ToString() });
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion

                        #region 回傳 目前工站既有資訊, 有註冊的OP
                        DataRow dr2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{tmp.StationNO}'");
                        if (dr2 != null && dr2["OP_NO"].ToString().Trim() != "")
                        {
                            tmp.PP_Name = dr2["PP_Name"].ToString();
                            tmp.IndexSN = dr2["IndexSN"].ToString();
                            tmp.OrderNO = dr2["OrderNO"].ToString();
                            tmp.PartNO = dr2["PartNO"].ToString();
                            if (tmp.PartNO != "")
                            {
                                DataRow dr_Material = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{tmp.PartNO}'");
                                if (dr_Material != null)
                                {
                                    tmp.PartName = dr_Material["PartName"].ToString();
                                    tmp.Specification = dr_Material["Specification"].ToString();
                                }
                            }
                            if (dr2["State"].ToString() == "1") { tmp.Station_State = "生產中......"; }
                            else if (dr2["State"].ToString() == "3") { tmp.Station_State = "工站暫停."; }
                            else if (dr2["State"].ToString() == "2") { tmp.Station_State = "工站停止中......"; }
                            else if (dr2["State"].ToString() == "4") { tmp.Station_State = "已完成工站設定, 等待開始生產."; }
                            tmp.HasPO_List = new List<string[]>();
                            string[] s0 = dr2["OP_NO"].ToString().Trim().Split(';');
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
                        DataRow dr_SFC_StationProjectDetail = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{tmp.StationNO}' and PP_Name='{tmp.PP_Name}' and OP_NO={tmp.OPNO} and OrderNO='{tmp.OrderNO}' and PartNO='{tmp.PartNO}' and StartTime is not null and EndTime is null order by StartTime desc");
                        if (dr_SFC_StationProjectDetail != null)
                        {
                            //同一站可能有多作業員, 所有人合計總工時
                            DataTable dt_SFC_StationProjectDetail = db.DB_GetData($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where Id='{dr_SFC_StationProjectDetail["Id"].ToString()}' and StationNO='{tmp.StationNO}'");
                            TimeSpan workTime = TimeSpan.Zero;
                            foreach (DataRow d in dt_SFC_StationProjectDetail.Rows)
                            {
                                tmp.TOTALOKQTY += int.Parse(d["ProductFinishedQty"].ToString());
                                tmp.TOTALFailQTY += int.Parse(d["ProductFailedQty"].ToString());
                                if (!d.IsNull("StartTime"))
                                {
                                    workTime = workTime.Add(_SFC_Common.GetCT(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(d["StartTime"]), DateTime.Now));
                                }
                            }
                            tmp.TOTALWorkTime = $"{((int)workTime.TotalHours).ToString()}小時 {workTime.Minutes.ToString()}分";
                        }
                        #endregion
                    }
                }
                #endregion
            }
            else { tmp.ERRMsg = "網頁無工站的來源參數, 請關閉瀏覽器中所有系統的網頁, 重新刷條碼 並 重新操作."; }
            //若將來要存條碼 tmp.QRCode=_Fun.Craete2BarCode(id, 300, 300);
            return View(tmp);
        }
        
        public ActionResult GetBarCode(string id)
        {
            string[] tmp = id.Split(';');
            byte[] imageData = null;
            if (tmp.Length == 4)
            {
                imageData = _Fun.Craete2BarCode(tmp[0], int.Parse(tmp[1]), int.Parse(tmp[2]), tmp[3]);
            }
            else
            {
                if (tmp[0] != "" && tmp[1] != "")
                {
                    string meg = $"{tmp[0]};{tmp[1]}";
                    imageData = _Fun.Craete2BarCode(meg, int.Parse(tmp[2]), int.Parse(tmp[3]), tmp[4]);
                }
            }
            return File(imageData, "image/jpeg");
        }
        public ActionResult CreateBarCode(LabelCreateBarCode key)
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            key.StationNOList = new List<string[]>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' and Station_Type='1'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow d in tmp_dt.Rows)
                    {
                        key.StationNOList.Add(new string[] {"", d["StationNO"].ToString(), d["StationName"].ToString() });
                    }
                }
                tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow d in tmp_dt.Rows)
                    {
                        key.StationNOList.Add(new string[] { d["StoreNO"].ToString(), d["StoreName"].ToString(), "" });
                    }
                }
            }
            return View(key);
        }
        public ActionResult StoreSpacesNOBarCode(LabelCreateBarCode key)
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            key.StationNOList = new List<string[]>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO!='' and StoreSpacesNO!=''");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow d in tmp_dt.Rows)
                    {
                        key.StationNOList.Add(new string[] { d["StoreNO"].ToString(), d["StoreSpacesNO"].ToString(), d["StoreSpacesName"].ToString() });
                    }
                }
            }
            return View(key);
        }

        
        [HttpPost]
        public ActionResult SetAction(LabelProject keys)
        {
            var br = _Fun.GetBaseUser();
            if (keys == null || br == null || !br.IsLogin || br.UserNO.Trim() == "") 
            {
                
                ViewBag.ErrType = "";
                ViewBag.ERRMsg = $"作業失敗, 畫面已逾時, 請關閉瀏覽器中所有系統的網頁, 重新刷條碼 並 重新操作.";
                ViewBag.Report = "";
                ViewData.Model = keys;
                ViewBag.StationNO = "";
                if (keys != null && keys.StationNO != "")
                { ViewBag.StationNO = keys.StationNO; }
                return View("ResuItTimeOUT"); 
            }

            if (keys.State!= "發Mail") { keys.ERRMsg = ""; }
            ViewBag.Report = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                try
                {

                    DataRow tmp_dr = null;
                    DataRow dr_M = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                    if (keys.State != "發Mail")
                    {
                        if (dr_M == null) { keys.ERRMsg = $"系統異常, 原因:資料庫Manufacture設定異常, 請通知管理者."; goto break_FUN; }
                        if (dr_M["Config_macID"].ToString().Trim() == "") { keys.ERRMsg = $"系統異常, 查無標籤設定資料, 請通知管理者."; goto break_FUN; }
                    }
                    DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");

                    switch (keys.State)
                    {
                        case "歷史查詢":
                            {
                                string report = "";
                                if (keys.OPNO == null || keys.OPNO.Trim() == "") { keys.ERRMsg = $"作業失敗, 畫面已逾時, 請關閉瀏覽器中所有系統的網頁, 重新刷條碼 並 重新操作."; goto break_FUN; }
                                DataTable dt_tmp = db.DB_GetData($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and LOGDateTime>='{DateTime.Now.AddDays(-14).ToString("MM/dd/yyyy HH:mm:ss.fff")}'  order by LOGDateTime desc");
                                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                {
                                    report = "<div>顯示 2週內的所有報工紀錄</div><div><table data-role='table' data-mode='columntoggle' class='ui-responsive' id='myTable'><thead><tr><th>報工日期</th><th data-priority='1'>製程名稱</th><th>報工人員</th><th>料號</th><th>OK量</th><th>Fail量</th><th data-priority='2'>實際工時</th><th data-priority='2'>標準工時</th></tr></thead><tbody>";
                                    foreach (DataRow d in dt_tmp.Rows)
                                    {
                                        report = $"{report}<tr><td>{Convert.ToDateTime(d["LOGDateTime"]).ToString("MM/dd HH:mm")}</td><td>{d["PP_Name"].ToString()}</td><td>{d["OP_NO"].ToString()}</td><td>{d["PartNO"].ToString()}</td><td>{d["OKQTY"].ToString()}</td><td>{d["FailQTY"].ToString()}</td><td>{d["CycleTime"].ToString()}</td><td>{float.Parse(d["EfficientCycleTime"].ToString()).ToString("0.0")}</td></tr>";
                                    }
                                    report = $"{report}</tbody></table></div>";
                                }
                                else { report = $"<div>{keys.StationNO} 目前查無14天內的報工紀錄.</div>"; }
                                ViewBag.Report = report;
                            }
                            break;
                        case "設定工站":
                            {
                                #region 網頁內容檢查
                                DataRow dr = null;
                                if (dr_M["State"].ToString() == "1") { keys.ERRMsg = $"本工站 {keys.StationNO} 已在生產中, 請先停止才能設定."; goto break_FUN; }
                                if (keys.PP_Name == null) { keys.PP_Name = ""; }
                                if (keys.OrderNO == null) { keys.OrderNO = ""; }
                                if (keys.PartNO == null) { keys.PartNO = ""; }
                                if (keys.PP_Name == "")
                                { keys.ERRMsg = $"製程名稱必須要有值, 請重新設定."; goto break_FUN; }
                                else if (keys.OrderNO == "" && keys.PartNO == "")
                                { keys.ERRMsg = $"工單編號 或 料件編號 其中一項不能空白, 請重新設定."; goto break_FUN; }
                                if (keys.OrderNO.Trim() != "")
                                {
                                    DataRow sfcdr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{keys.OrderNO}'");
                                    if (sfcdr == null) { keys.ERRMsg = $"查無 {keys.OrderNO} 工單資料, 請重新輸入."; goto break_FUN; }
                                }
                                if (keys.PP_Name.Trim() != "")
                                {
                                    if (keys.PP_Name.IndexOf(";") >= 0)
                                    {
                                        string[] data = keys.PP_Name.Split(';');
                                        keys.PP_Name = data[0];
                                        keys.IndexSN = data[1];
                                        if (keys.IndexSN.Trim() == "") { keys.IndexSN = "0"; }
                                    }
                                    dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{keys.PP_Name}' and StationNO='{keys.StationNO}' and IndexSN={keys.IndexSN}");
                                    if (dr == null)
                                    { keys.ERRMsg = $"查無相關製程順序.製程={keys.PP_Name} 順序={keys.IndexSN}, 請通知管理者."; goto break_FUN; }
                                }
                                if (keys.PartNO.Trim() != "")
                                {
                                    dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{keys.PartNO}'");
                                    if (dr == null)
                                    { keys.ERRMsg = $"查無相關料件編號, 請通知管理者."; goto break_FUN; }
                                }
                                #endregion

                                #region 檢查設定之前是否無報工, 且 ReportTime>5
                                if (dr_M["PartNO"].ToString().Trim() != "")
                                {
                                    int avgReportTime = 0;
                                    DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_M["PP_Name"].ToString()}' and IndexSN={dr_M["IndexSN"].ToString()} and PartNO='{dr_M["PartNO"].ToString()}' and ReportTime>5");
                                    if (dr_tmp != null && !dr_tmp.IsNull("AVGTime") && dr_tmp["AVGTime"].ToString().Trim() != "")
                                    {
                                        avgReportTime = int.Parse(dr_tmp["AVGTime"].ToString());
                                    }
                                    if (avgReportTime > 0)
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PartNO='{dr_M["PartNO"].ToString()}' and LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%開工%' order by LOGDateTime desc");
                                        if (dr_tmp != null)
                                        {
                                            DateTime tmp_edate = Convert.ToDateTime(dr_tmp["LOGDateTime"]);
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PartNO='{dr_M["PartNO"].ToString()}' and LOGDateTime>'{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%報工%'");
                                            if (dr_tmp == null)
                                            {
                                                if ((_SFC_Common.TimeCompute2Seconds(tmp_edate, DateTime.Now)) >= avgReportTime)
                                                {
                                                    keys.ERRMsg = $"疑似 前一次料號:{dr_M["PartNO"].ToString()} 未完成報工. 若未報工,請先報工, 否則請先執行 關站設定.";
                                                }
                                            }
                                        }
                                    }
                                }
                                if (keys.ERRMsg != "") { goto break_FUN; }
                                #endregion


                                #region 更新Manufacture 製造現場狀態
                                if (!db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET StartTime=NULL,RemarkTimeS=NULL,RemarkTimeE=NULL,EndTime=NULL,Label_ProjectType='1',State='2',OrderNO='{keys.OrderNO}',IndexSN={keys.IndexSN},OP_NO='{keys.OPNO}',Master_PP_Name='{keys.PP_Name}',PP_Name='{keys.PP_Name}',PartNO='{keys.PartNO}',SimulationId='' where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'"))
                                {
                                    keys.ERRMsg = $"工站設定失敗, 請通知管理者."; goto break_FUN;
                                }
                                #endregion


                                #region 處理SFC_StationProjectDetail
                                int dis_DetailQTY = 0;
                                TimeSpan workTime = TimeSpan.Zero;
                                //先查有沒有以前做過設定
                                DataTable dt_SFC_StationProjectDetail = db.DB_GetData($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{keys.PP_Name}' and OrderNO='{keys.OrderNO}' and OP_NO='{keys.OPNO}' and PartNO='{keys.PartNO}' and IndexSN={keys.IndexSN} and EndTime is NULL");
                                if (dt_SFC_StationProjectDetail != null && dt_SFC_StationProjectDetail.Rows.Count > 0)
                                {
                                    dt_SFC_StationProjectDetail = db.DB_GetData($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where Id='{dt_SFC_StationProjectDetail.Rows[0]["Id"].ToString()}'");
                                    foreach (DataRow d in dt_SFC_StationProjectDetail.Rows)
                                    {
                                        keys.TOTALOKQTY += int.Parse(d["ProductFinishedQty"].ToString());
                                        keys.TOTALFailQTY += int.Parse(d["ProductFailedQty"].ToString());
                                        dis_DetailQTY += keys.TOTALOKQTY + keys.TOTALFailQTY;
                                        if (!d.IsNull("StartTime"))
                                        { 
                                            workTime = workTime.Add(_SFC_Common.GetCT(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(d["StartTime"]), DateTime.Now));
                                        }
                                    }
                                }
                                else
                                {
                                    db.DB_SetData(@$"INSERT INTO SoftNetLogDB.[dbo].[SFC_StationProjectDetail] (ServerId,Id,StationNO,Master_PP_Name,PP_Name,OP_NO,IndexSN,OrderNO,PartNO,RMSName) VALUES 
                                                ('{_Fun.Config.ServerId}','{_Str.NewId('A')}','{keys.StationNO}','{keys.PP_Name}','{keys.PP_Name}','{keys.OPNO}',{keys.IndexSN},'{keys.OrderNO}','{keys.PartNO}','{dr_PP_Station["RMSName"].ToString()}')");
                                }
                                #endregion

                                #region 更新Tag
                                if (_Fun.Has_Tag_httpClient)
                                {
                                    string tmp_s = $"http://{_Fun.Config.LocalWebURL}/LabelProject/Index/{keys.StationNO}";
                                    string isUpdate = "1";
                                    var json4 = "";
                                    if (_Fun.Is_Tag_Connect)
                                    {
                                        json4 = $"\"Text1\":\"工站編號:\",\"StationNO\":\"{keys.StationNO}\",\"StationName\":\"{keys.StationName}\",\"Text2\":\"製程名稱:\",\"PP_Name\":\"{keys.PP_Name}\",\"Text3\":\"工單編號:\",\"OrderNO\":\"{keys.OrderNO}\",\"Text4\":\"料號:\",\"PartNO\":\"{keys.PartNO}\",\"Text5\":\"作業人員:\",\"OPNO\":\"{keys.OPNO}\",\"Text6\":\"累計工時:\",\"WorkTime\":\"{((int)workTime.TotalHours).ToString()}.{workTime.Minutes.ToString()}\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                                        var json = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":45,\"styleid\":54,{json4}";
                                        json = $"{json},\"QTY\":\"{dis_DetailQTY.ToString()}\",\"ledrgb\":\"0\",\"ledstate\":0";
                                        _Fun.Tag_Write(db,dr_M["Config_macID"].ToString(),"", json);
                                    }
                                    else { isUpdate = "0"; }
                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{json4}',Ledrgb='0',Ledstate=0,StationNO='{keys.StationNO}',Type='4',OrderNO='{keys.OrderNO}',IndexSN='{keys.IndexSN}',StoreNO='',StoreSpacesNO='',QTY={dis_DetailQTY.ToString()},IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                }
                                #endregion

                                keys.Station_State = "設定完成.";
                                keys.MES_String = "完成 工站設定作業....";
                                db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelProject','設定工站','{keys.PP_Name}','{keys.StationNO}','{keys.PartNO}','{keys.OrderNO}','{keys.OPNO}')");
                            }
                            break;
                        case "設定人員":
                            {
                                if (keys.OPNO == null || keys.OPNO.Trim() == "") { keys.ERRMsg = $"作業失敗, 畫面已逾時, 請關閉瀏覽器中所有系統的網頁, 重新刷條碼 並 重新操作."; goto break_FUN; }
                                if (keys.PP_Name == null) { keys.PP_Name = ""; }
                                if (keys.OrderNO == null) { keys.OrderNO = ""; }
                                if (keys.PartNO == null) { keys.PartNO = ""; }
                                if (keys.PP_Name == "" || keys.IndexSN == "" || dr_M["PP_Name"].ToString().Trim() == "") { keys.ERRMsg = $"製程名稱必須要有值, 請重新工站設定."; goto break_FUN; }
                                string name = keys.OPNO;
                                int dis_DetailQTY = 0;
                                TimeSpan workTime = TimeSpan.Zero;
                                if (dr_M["OP_NO"].ToString().Trim() != "")
                                {
                                    List<string> tmp = dr_M["OP_NO"].ToString().Trim().Split(';').ToList();
                                    if (tmp.Count > 0)
                                    {
                                        tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{keys.PP_Name}' and OP_NO='{tmp[0]}' and OrderNO='{keys.OrderNO}' and PartNO='{keys.PartNO}' and IndexSN={keys.IndexSN} and EndTime is NULL order by Id");
                                        if (tmp_dr != null)
                                        {
                                            DataTable dt_SFC_StationProjectDetail = db.DB_GetData($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where Id='{tmp_dr["Id"].ToString()}'");
                                            foreach (DataRow d in dt_SFC_StationProjectDetail.Rows)
                                            {
                                                keys.TOTALOKQTY += int.Parse(d["ProductFinishedQty"].ToString());
                                                keys.TOTALFailQTY += int.Parse(d["ProductFailedQty"].ToString());
                                                dis_DetailQTY += keys.TOTALOKQTY + keys.TOTALFailQTY;
                                                if (!d.IsNull("StartTime"))
                                                { workTime = workTime.Add(_SFC_Common.GetCT(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(d["StartTime"]), DateTime.Now)); }
                                            }
                                        }
                                    }
                                    if (!tmp.Contains(keys.OPNO))
                                    {
                                        name = $"{dr_M["OP_NO"].ToString().Trim()};{keys.OPNO}";
                                        string RMSName = "";
                                        string st = "NULL";
                                        string rst = "NULL";
                                        if (tmp_dr != null)
                                        {
                                            RMSName = tmp_dr["RMSName"].ToString();
                                            if (!tmp_dr.IsNull("StartTime")) { st = $"'{Convert.ToDateTime(tmp_dr["StartTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                        }
                                        if (dr_M["State"].ToString() == "1")
                                        {
                                            if (st == "NULL") { st = $"'{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                            rst = $"'{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'";
                                        }
                                        db.DB_SetData(@$"INSERT INTO SoftNetLogDB.[dbo].[SFC_StationProjectDetail] (ServerId,Id,StationNO,Master_PP_Name,PP_Name,OP_NO,IndexSN,OrderNO,PartNO,RMSName,StartTime,RemarkTimeS) VALUES 
                                                        ('{_Fun.Config.ServerId}','{tmp_dr["Id"].ToString()}','{dr_M["StationNO"].ToString()}','{dr_M["Master_PP_Name"].ToString()}','{dr_M["PP_Name"].ToString()}','{keys.OPNO}',{dr_M["IndexSN"].ToString()},'{dr_M["OrderNO"].ToString()}','{dr_M["PartNO"].ToString()}','{RMSName}',{st},{rst})");
                                    }
                                    else { keys.ERRMsg = $"作業人員已存在本工站人員名單."; goto break_FUN; }
                                }
                                if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET OP_NO='{name}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'"))
                                {
                                    #region 更新Tag
                                    if (_Fun.Has_Tag_httpClient)
                                    {
                                        string tmp_s = $"http://{_Fun.Config.LocalWebURL}/LabelProject/Index/{keys.StationNO}";
                                        string isUpdate = "1";
                                        var json4 = "";
                                        var jion_ShowValue = "";
                                        if (_Fun.Is_Tag_Connect)
                                        {
                                            jion_ShowValue = $"\"Text1\":\"工站編號:\",\"StationNO\":\"{keys.StationNO}\",\"StationName\":\"{keys.StationName}\",\"Text2\":\"製程名稱:\",\"PP_Name\":\"{keys.PP_Name}\",\"Text3\":\"工單編號:\",\"OrderNO\":\"{keys.OrderNO}\",\"Text4\":\"料號:\",\"PartNO\":\"{keys.PartNO}\",\"Text5\":\"作業人員:\",\"OPNO\":\"多人共同作業\",\"Text6\":\"累計工時:\",\"WorkTime\":\"{((int)workTime.TotalHours).ToString()}.{workTime.Minutes.ToString()}\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                                            json4 = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":45,\"styleid\":54,{jion_ShowValue}";
                                            tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                            var json = $"{json4},\"QTY\":\"{dis_DetailQTY.ToString()}\",\"ledrgb\":\"{tmp_dr["Ledrgb"].ToString()}\",\"ledstate\":{tmp_dr["Ledstate"].ToString()}";
                                            _Fun.Tag_Write(db,dr_M["Config_macID"].ToString(),"", json);
                                        }
                                        else { isUpdate = "0"; }
                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{jion_ShowValue}',QTY={dis_DetailQTY.ToString()},IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                    }
                                    #endregion
                                    keys.MES_String = "完成 人員新增作業....";
                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelProject','設定人員','{keys.PP_Name}','{keys.StationNO}','{keys.PartNO}','{keys.OrderNO}','{keys.OPNO}')");
                                }
                            }
                            break;
                        case "報工":
                            {
                                if (keys.OKQTY == 0 && keys.FailQTY==0) { keys.ERRMsg = $"未填寫數量, 請重新設定."; goto break_FUN; }
                                if (dr_M["PP_Name"].ToString().Trim() == "") { keys.ERRMsg = $"報工前, 需先做工站設定."; goto break_FUN; }
                                if (keys.OPNO == null || keys.OPNO.Trim() == "") { keys.ERRMsg = $"人員資料為空, 錯誤."; goto break_FUN; }
                                if (keys.PP_Name == null) { keys.PP_Name = ""; }
                                if (keys.OrderNO == null) { keys.OrderNO = ""; }
                                if (keys.PartNO == null) { keys.PartNO = ""; }
                                if (keys.PP_Name == "") { keys.ERRMsg = $"製程名稱必須要有值, 請重新設定."; goto break_FUN; }
                                DateTime startTime = DateTime.Now;
                                if (keys.OKQTY > 0 || keys.FailQTY > 0)
                                {
                                    keys.TOTALOKQTY = 0;
                                    keys.TOTALFailQTY = 0;
                                    DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{keys.PP_Name}' and OP_NO='{keys.OPNO}' and OrderNO='{keys.OrderNO}' and PartNO='{keys.PartNO}' and IndexSN={keys.IndexSN} and EndTime is NULL");
                                    if (dr != null)
                                    {
                                        if (dr.IsNull("StartTime"))
                                        { keys.ERRMsg = $"工站未設定開始過, 無法報工"; goto break_FUN; }
                                        else
                                        {
                                            #region 計算報工 CT,平均,有效, 寫SFC_StationProjectDetail
                                            DateTime rRemarkTimeS = new DateTime();
                                            if (dr.IsNull("RemarkTimeS"))
                                            { rRemarkTimeS = Convert.ToDateTime(dr["StartTime"]); }
                                            else
                                            { rRemarkTimeS = Convert.ToDateTime(dr["RemarkTimeS"]); }
                                            List<double> allCT = new List<double>();
                                            string docNO = "";
                                            string top_flag = $" TOP {_Fun.Config.AdminKey03} ";

                                            decimal ct = 0;

                                            if (dr.IsNull("RemarkTimeE") || rRemarkTimeS >= Convert.ToDateTime(dr["RemarkTimeE"]))
                                            { ct = Math.Round((decimal)_SFC_Common.GetCT(db, dr_PP_Station["CalendarName"].ToString(), rRemarkTimeS, DateTime.Now).TotalSeconds / (keys.OKQTY + keys.FailQTY)); }
                                            else
                                            { ct = Math.Round((decimal)_SFC_Common.GetCT(db, dr_PP_Station["CalendarName"].ToString(), rRemarkTimeS, Convert.ToDateTime(dr["RemarkTimeE"])).TotalSeconds / (keys.OKQTY + keys.FailQTY)); }
                                            decimal ct_log = ct < 1 ? 0 : ct;
                                            if (int.Parse(dr["CycleTime"].ToString()) != 0) { ct = (ct + int.Parse(dr["CycleTime"].ToString())) > 0 ? Math.Round((ct + int.Parse(dr["CycleTime"].ToString())) / 2) : ct; }
                                            if (ct < 1) { ct = 0; }
                                            string remark = "";
                                            DataRow dr_SFC_StationProjectDetail = null;
                                            if (dr_M["State"].ToString() == "1")
                                            {
                                                dr_SFC_StationProjectDetail = db.DB_GetFirstDataByDataRow($@"SELECT (sum(SD_LowerLimit*CountQTY)/sum(CountQTY)) as CCT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] 
                                                where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr["PP_Name"].ToString()}' and PartNO='{dr["PartNO"].ToString()}' and DOCNO='' 
                                                group by ServerId,StationNO,PP_Name,PartNO");
                                                if (dr_SFC_StationProjectDetail != null && !dr_SFC_StationProjectDetail.IsNull("CCT") && int.Parse(dr_SFC_StationProjectDetail["CCT"].ToString()) > 2)
                                                {
                                                    float tmp = float.Parse(dr_SFC_StationProjectDetail["CCT"].ToString()) * 0.85f;
                                                    if (ct_log >= (decimal)tmp) { remark = $",RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                }
                                            }
                                            if (keys.Remark != null && keys.Remark.Trim() != "") { remark = $"{remark},Remark='{keys.Remark.Trim()}'"; }
                                            if (db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SFC_StationProjectDetail] SET ProductFinishedQty+={keys.OKQTY.ToString()},ProductFailedQty+={keys.FailQTY.ToString()},CycleTime={ct.ToString()}{remark} where ServerId='{_Fun.Config.ServerId}' and Id='{dr["Id"].ToString()}' and StationNO='{keys.StationNO}' and OP_NO='{keys.OPNO}'"))
                                            {
                                                remark = $"RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'";
                                                tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                                                if (tmp_dr.IsNull("StartTime")) { remark = $"{remark},StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET {remark} where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                                            }
                                            dr_SFC_StationProjectDetail = db.DB_GetFirstDataByDataRow($"SELECT sum(ProductFinishedQty) as OKQTY,sum(ProductFailedQty) as FAILQTY FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{dr["Id"].ToString()}'");
                                            if (dr_SFC_StationProjectDetail != null)
                                            {
                                                keys.TOTALOKQTY = int.Parse(dr_SFC_StationProjectDetail["OKQTY"].ToString());
                                                keys.TOTALFailQTY = int.Parse(dr_SFC_StationProjectDetail["FAILQTY"].ToString());
                                            }
                                            //###???未來專案版改智慧版 要改DB內要改PartNO,Sub_PartNO,Master_PP_Name,PP_Name
                                            string tmp_s = "";
                                            if (dr_M["PartNO"].ToString().Trim() != "") { tmp_s = $" and PartNO='{dr_M["PartNO"].ToString().Trim()}'"; }
                                            DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} * from SoftNetLogDB.[dbo].SFC_StationProjectDetail where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_M["StationNO"].ToString()}' and PP_Name='{dr_M["PP_Name"].ToString()}' {tmp_s} and ProductFinishedQty!=0 and CycleTime!=0 order by StationNO,OP_NO,PP_Name,PartNO");
                                            if (dt_Efficient != null && dt_Efficient.Rows.Count > 0)
                                            {
                                                string partNO = dt_Efficient.Rows[0].IsNull("PartNO") ? "" : dt_Efficient.Rows[0]["PartNO"].ToString();
                                                string E_stationNO = dt_Efficient.Rows[0].IsNull("StationNO") ? "" : dt_Efficient.Rows[0]["StationNO"].ToString();
                                                string pp_Name = dt_Efficient.Rows[0].IsNull("PP_Name") ? "" : dt_Efficient.Rows[0]["PP_Name"].ToString();
                                                for (int i2 = 0; i2 < dt_Efficient.Rows.Count; i2++)
                                                {
                                                    DataRow dr2 = dt_Efficient.Rows[i2];
                                                    if (E_stationNO != (dr2.IsNull("StationNO") ? "" : dr2["StationNO"].ToString()) || partNO != (dr2.IsNull("PartNO") ? "" : dr2["PartNO"].ToString()) || pp_Name != (dr2.IsNull("PP_Name") ? "" : dr2["PP_Name"].ToString()))
                                                    {
                                                        _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, dr2["Master_PP_Name"].ToString(),"0", dr2["PP_Name"].ToString(), partNO, partNO, docNO);
                                                        E_stationNO = dr2.IsNull("StationNO") ? "" : dr2["StationNO"].ToString();
                                                        partNO = dr2.IsNull("PartNO") ? "" : dr2["PartNO"].ToString();
                                                        pp_Name = dr2.IsNull("PP_Name") ? "" : dr2["PP_Name"].ToString();
                                                        allCT.Clear();
                                                    }
                                                    for (int tmp01 = 1; tmp01 <= (int)dr2["ProductFinishedQty"]; tmp01++)//工單數目若為n 需算作n筆
                                                    {
                                                        if (_Fun.Config.AdminKey14)
                                                        { allCT.Add(double.Parse(dr2["CycleTime"].ToString()) + double.Parse(dr2["WaitTime"].ToString())); }
                                                        else
                                                        { allCT.Add(double.Parse(dr2["CycleTime"].ToString())); }
                                                    }
                                                }
                                                if (allCT.Count > 0)
                                                {
                                                    _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, pp_Name, pp_Name,"0", partNO, partNO, docNO);
                                                }
                                            }
                                            string efficientCycleTime = "0";
                                            string custom_SD_LowerLimit = "0";
                                            dr_SFC_StationProjectDetail = db.DB_GetFirstDataByDataRow($@"SELECT (sum(EfficientCycleTime*CountQTY)/sum(CountQTY)) as ECT,(sum(Custom_SD_LowerLimit*CountQTY)/sum(CountQTY)) as CCT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] 
                                            where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr["PP_Name"].ToString()}' and PartNO='{dr["PartNO"].ToString()}' and DOCNO='' 
                                            group by ServerId,StationNO,PP_Name,PartNO");
                                            if (dr_SFC_StationProjectDetail != null)
                                            {
                                                efficientCycleTime = dr_SFC_StationProjectDetail.IsNull("ECT") ? "0" : dr_SFC_StationProjectDetail["ECT"].ToString();
                                                custom_SD_LowerLimit = dr_SFC_StationProjectDetail.IsNull("CCT") ? "0" : dr_SFC_StationProjectDetail["CCT"].ToString();
                                            }
                                            #endregion

                                            #region 紀錄報工log
                                            int reportTime = 0;

                                            #region 計算上一次與現在時間差
                                            DataRow d2 = db.DB_GetFirstDataByDataRow($"SELECT LOGDateTimeID FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and Id='{dr["Id"].ToString()}' order by LOGDateTime,LOGDateTimeID desc");
                                            if (d2 != null)
                                            { reportTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(d2["LOGDateTimeID"]), DateTime.Now); }
                                            #endregion


                                            db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] (ServerId,[LOGDateTime],[Id],[StationNO],[OP_NO],[OKQTY],[FailQTY],[CycleTime],[WaitTime],PP_Name,IndexSN,PartNO,OrderNO,EfficientCycleTime,Custom_SD_LowerLimit,ReportTime) VALUES (
                                                                '{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{dr["Id"].ToString()}','{keys.StationNO}','{keys.OPNO}',{keys.OKQTY.ToString()},{keys.FailQTY.ToString()},{ct_log},0,'{dr["PP_Name"].ToString()}',{dr["IndexSN"].ToString()},'{dr["PartNO"].ToString()}','{dr["OrderNO"].ToString()}',{efficientCycleTime},{custom_SD_LowerLimit},{reportTime})");
                                            #endregion

                                            #region 更新標籤累計量, 30秒內更新
                                            if (_Fun.Has_Tag_httpClient)
                                            {
                                                int dis_DetailQTY = keys.OKQTY + keys.FailQTY;
                                                DataRow dr_QTY = db.DB_GetFirstDataByDataRow($"SELECT sum(ProductFinishedQty+ProductFailedQty) as QTY FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where Id='{dr["Id"].ToString()}' and StationNO='{keys.StationNO}'");
                                                if (dr_QTY != null && !dr_QTY.IsNull("QTY") && dr_QTY["QTY"].ToString() != "")
                                                { dis_DetailQTY = int.Parse(dr_QTY["QTY"].ToString()); }
                                                DataRow totalData = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                                if (totalData != null && dis_DetailQTY.ToString().Trim() != totalData["QTY"].ToString().Trim())
                                                {
                                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set QTY='{dis_DetailQTY.ToString()}',IsUpdate='0' where ServerId='{_Fun.Config.ServerId}' and macID='{totalData["macID"].ToString()}'");
                                                }
                                                keys.MES_String = "完成報工作業, 稍後會更新標籤數量.";
                                            }
                                            else { keys.MES_String = "完成報工作業."; }
                                            #endregion
                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelProject','學習模式報工','{keys.PP_Name}','{keys.StationNO}','{keys.PartNO}','{keys.OrderNO}','{keys.OPNO}')");
                                        }
                                    }
                                    else
                                    { keys.ERRMsg = $"查無歷史啟動過資料,無法報工, 請通知管理者"; goto break_FUN; }
                                }

                            }
                            break;
                        case "開始":
                            {
                                if (dr_M["State"].ToString() == "1") { goto break_FUN; }
                                if (dr_M["PP_Name"].ToString().Trim() == "") { keys.ERRMsg = $"開始前, 需先做工站設定."; goto break_FUN; }
                                //keys.ERRMsg = _SFC_Common.LabelProject_Start_Stop(db, "1", keys.StationNO, dr_M, keys.OPNO, ref keys);
                                if (keys.ERRMsg == "")
                                {
                                    keys.Station_State = "生產中...";
                                    keys.MES_String = "完成工站啟動生產.";
                                }
                            }
                            break;
                        case "停止":
                            {
                                if (dr_M["State"].ToString() == "2") { goto break_FUN; }
                                if (dr_M["PP_Name"].ToString().Trim() == "") { keys.ERRMsg = $"停止前, 需先工站設定."; goto break_FUN; }
                                //keys.ERRMsg = _SFC_Common.LabelProject_Start_Stop(db, "2", keys.StationNO, dr_M, keys.OPNO, ref keys);
                                if (keys.ERRMsg == "")
                                {
                                    keys.Station_State = "工站停止中...";
                                    keys.MES_String = "已停止工站生產作業.";
                                }
                            }
                            break;
                        case "關站":
                            {
                                #region 檢查設定之前是否無報工, 且 ReportTime>5
                                int avgReportTime = 0;
                                DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_M["PP_Name"].ToString()}' and IndexSN={dr_M["IndexSN"].ToString()} and PartNO='{dr_M["PartNO"].ToString()}' and ReportTime>5");
                                if (dr_tmp != null && !dr_tmp.IsNull("AVGTime") && dr_tmp["AVGTime"].ToString().Trim() != "")
                                {
                                    avgReportTime = int.Parse(dr_tmp["AVGTime"].ToString());
                                }
                                if (avgReportTime > 0)
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PartNO='{dr_M["PartNO"].ToString()}' and LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%開工%' order by LOGDateTime desc");
                                    if (dr_tmp != null)
                                    {
                                        DateTime tmp_edate = Convert.ToDateTime(dr_tmp["LOGDateTime"]);
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PartNO='{dr_M["PartNO"].ToString()}' and LOGDateTime>'{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%報工%'");
                                        if (dr_tmp == null)
                                        {
                                            if ((_SFC_Common.TimeCompute2Seconds(tmp_edate, DateTime.Now)) >= avgReportTime)
                                            {
                                                keys.ERRMsg = $"疑似 前一次料號:{dr_M["PartNO"].ToString()} 未完成報工. 若未報工,請先報工, 否則請先執行 關站設定.";
                                            }
                                        }
                                    }
                                }
                                if (keys.ERRMsg != "") { goto break_FUN; }
                                #endregion


                                if (dr_M["PP_Name"].ToString().Trim() == "") { keys.ERRMsg = $"關站前, 需有工站設定."; goto break_FUN; }
                                //keys.ERRMsg = _SFC_Common.LabelProject_Start_Stop(db, "4", keys.StationNO, dr_M, keys.OPNO, ref keys);
                                if (keys.ERRMsg == "")
                                {
                                    keys.Station_State = "已完成設定,待機中";
                                    keys.MES_String = "已關閉本次生產作業, 可進入工站設定, 設定下次生產";
                               }
                            }
                            break;
                        case "發Mail":
                            {
                                System.Threading.Tasks.Task task = _Log.ErrorAsync($"網頁回傳錯誤通知:內容 = {keys.ERRMsg}", true);
                                ViewBag.ERRMsg = "";
                                ViewBag.ErrType = "";
                                ViewBag.Report = "訊息已發出Mail到郵箱, 已通知廠商處理.";
                                ViewData.Model = keys;
                                ViewBag.StationNO = "";
                                if (keys != null && keys.StationNO != "")
                                { ViewBag.StationNO = keys.StationNO; }
                            }
                            return View("ResuItTimeOUT");
                    }
                }
                catch (Exception ex)
                {
                    ViewBag.ErrType = "SystemError";
                    ViewBag.ERRMsg = $"系統異常, 程式出現錯誤, 請通知管理者.";
                    ViewBag.Report = "";
                    string state = "";
                    if (keys != null) { state = keys.State; }
                    ViewBag.StationNO = "";
                    if (keys != null && keys.StationNO != "")
                    { ViewBag.StationNO = keys.StationNO; }
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"LabelProjectController.cs SetAction {state} Exception: {ex.Message} {ex.StackTrace}", true);
                    return View("ResuItTimeOUT");
                }
            }
        break_FUN:
            //ViewBag.Test1 = keys;
            ViewData.Model = keys;
            return View("ResuItDisplay");
        }
        

    }
}
