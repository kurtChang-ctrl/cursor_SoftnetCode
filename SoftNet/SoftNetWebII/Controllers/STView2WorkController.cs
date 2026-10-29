using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Models;
using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using System;
using System.Linq;
using System.Text;
using Microsoft.Extensions.Primitives;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion.Internal;
using System.Drawing;
using System.Runtime.CompilerServices;
using static System.Net.Mime.MediaTypeNames;
using Base;





namespace SoftNetWebII.Controllers
{
    public class STView2WorkController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public STView2WorkController(SNWebSocketService websocket, SFC_Common sfc_Common)
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
        public ActionResult Index(MUTIStationObj keys)
        {
            keys.MES_Report = "";
            List<string[]> StationNOList = new List<string[]>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and Station_Type='8'");
                DataRow tmp = null;
                if (dt != null && dt.Rows.Count > 0)
                {
                    #region 確認 Manufacture 資料
                    foreach (DataRow dr in dt.Rows)
                    {
                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                        if (tmp == null)
                        { db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Manufacture] ([StationNO],[ServerId],[Config_MutiWO],[Label_ProjectType]) VALUES ('{dr["StationNO"].ToString()}','{_Fun.Config.ServerId}','1','0')"); }
                        else
                        {
                            if (!bool.Parse(tmp["Config_MutiWO"].ToString()))
                            { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Config_MutiWO='1',Label_ProjectType='0' where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'"); }
                        }
                        StationNOList.Add(new string[] { dr["StationNO"].ToString(), dr["StationName"].ToString() });
                    }
                    #endregion
                }

                string report = "";
                if (keys.StationNO != null && keys.StationNO.Trim() != "")
                {
                    DataTable dt_tmp = db.DB_GetData($"SELECT a.*,b.PP_Name FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] as a,SoftNetLogDB.[dbo].[SFC_StationDetail] as b where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{keys.StationNO}' and a.LOGDateTimeID>='{DateTime.Now.AddDays(-14).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and a.LOGDateTime=b.LOGDateTime and a.ServerId=b.ServerId and a.StationNO=b.StationNO order by a.LOGDateTime,a.OP_NO desc");
                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                    {
                        string opName = "";
                        string partName = "";
                        string eCT = "";
                        string avgReportTime = "";
                        var sb_report = new StringBuilder();
                        sb_report.Append("<div>前14日歷史報工紀錄</div><div><table data-role='table' data-mode='columntoggle' class='ui-responsive' id='myTable'><thead><tr><th>報工日期</th><th data-priority='1'>製程名稱</th><th>工單編號</th><th>報工人員</th><th>料號</th><th>品名  規格</th><th>報工OK量</th><th>Fail量</th><th data-priority='2'>目標CT</th><th data-priority='3'>實際CT</th><th data-priority='4'>平均報工時間</th><th data-priority='4'>實際花費時間</th></tr></thead><tbody>");
                        foreach (DataRow d in dt_tmp.Rows)
                        {
                            #region 目標CT,報工時間
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{d["PP_Name"].ToString()}' and IndexSN={d["IndexSN"].ToString()} and PartNO='{d["PartNO"].ToString()}'");
                            if (tmp != null) { eCT = tmp["EfficientCycleTime"].ToString(); } else { eCT = "0"; }
                            DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PartNO='{d["PartNO"].ToString()}' and ReportTime>10");
                            if (dr_tmp != null && !dr_tmp.IsNull("AVGTime") && dr_tmp["AVGTime"].ToString().Trim() != "")
                            { avgReportTime = dr_tmp["AVGTime"].ToString(); }
                            else { avgReportTime = "0"; }
                            #endregion

                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[User] where ServerId='{_Fun.Config.ServerId}' and UserNO='{d["OP_NO"].ToString()}'");
                            if (tmp != null) { opName = $"{tmp["UserNO"].ToString()} {tmp["Name"].ToString()}"; } else { opName = ""; }
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}'");
                            if (tmp != null) { partName = $"{tmp["PartName"].ToString()} {tmp["Specification"].ToString()}"; } else { partName = ""; }
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail] where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(d["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and StationNO='{keys.StationNO}'");
                            sb_report.Append($"<tr><td>{Convert.ToDateTime(d["LOGDateTimeID"]).ToString("yyyy-MM-dd HH:mm:ss")}</td><td>{tmp["PP_Name"].ToString()}</td><td>{tmp["OrderNO"].ToString()}</td><td>{opName}</td><td>{d["PartNO"].ToString()}</td><td>{partName}</td><td>{d["EditFinishedQty"].ToString()}</td><td>{d["EditFailedQty"].ToString()}</td><td>{eCT}</td><td>{d["CycleTime"].ToString()}</td><td>{avgReportTime}</td><td>{d["CycleTime"].ToString()}</td></tr>");
                        }
                        sb_report.Append("</tbody></table></div>");
                        report = sb_report.ToString();
                        keys.MES_Report = report;
                    }
                }
            }
            ViewBag.StationNOList = StationNOList;
            keys.StationNOList = StationNOList;


            return View(keys);
        }

        public ActionResult Read()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            MUTIStationObj keys = new MUTIStationObj();
            if (keys.StationNO == null || keys.StationNO.Trim() == "")
            {
                List<string[]> StationNOList = new List<string[]>();
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and Station_Type='8'");
                    DataRow tmp = null;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                            if (tmp == null)
                            { db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Manufacture] ([StationNO],[ServerId],[Config_MutiWO],[Label_ProjectType]) VALUES ('{dr["StationNO"].ToString()}','{_Fun.Config.ServerId}','1','0')"); }
                            else
                            {
                                /*
                                keys.Station_Config_Store_Type = tmp["Config_Store_Type"].ToString();
                                #region 檢查下一站是否為委外, 設定 Station_Config_Store_Type 網頁行為
                                if (!_Fun.Config.IsOutPackStationStore && tmp["SimulationId"].ToString() != "")
                                {
                                    DataRow dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{tmp["SimulationId"].ToString()}'");
                                    if (dr3 != null)
                                    {
                                        if (_Fun.Config.OutPackStationName == dr3["Apply_StationNO"].ToString()) { keys.Station_Config_Store_Type = "0"; }
                                    }
                                }
                                #endregion
                                */
                                if (!bool.Parse(tmp["Config_MutiWO"].ToString()))
                                { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Config_MutiWO='1',Label_ProjectType='0' where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'"); }
                            }
                            StationNOList.Add(new string[] { dr["StationNO"].ToString(), dr["StationName"].ToString() });
                        }
                    }
                }
                ViewBag.StationNOList = StationNOList;
                keys.StationNOList = StationNOList;
            }
            else
            {
                keys.ERRMsg = "";
            }
            return View(keys);
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

        [HttpPost]
        public ActionResult MUTIStation(MUTIStationObj keys)//選完工站後動作
        {
            try
            {
                if (keys == null) { keys = new MUTIStationObj(); keys.ERRMsg = $"作業失敗, 畫面已逾時, 請關閉網頁瀏覽器, 重新登入帳號 並 重新操作."; }
                else
                {
                    if (keys.StationNO == null || keys.StationNO.Trim() == "")
                    {
                        keys.ERRMsg = "你沒有選擇工站, 請按 回到工站選單畫面 重新選擇工站.";
                    }
                    else
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            if (false)
                            //if (_Fun.Config.RUNMode != '1')
                            {
                                //###???將來要寫智慧模式
                            }
                            else
                            {
                                if (!GetINFO(db, ref keys))
                                {
                                    ViewBag.ErrType = "SystemError";
                                    keys.OutError = true;
                                    ViewBag.ERRMsg = keys.ERRMsg;
                                    ViewBag.Report = "";
                                    return View("ResuItTimeOUT");
                                }
                            }
                        }
                    }
                }
                ViewBag.DATA = keys;
            }
            catch (Exception ex)
            {
                string state = "";
                if (keys != null)
                {
                    state = keys.State;
                    keys.OutError = true;
                }
                ViewBag.Report = "";
                ViewBag.ErrType = "SystemError";
                ViewBag.ERRMsg = $"程式異常, 導致系統無法使用, 請通知管理者.";
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs MUTIStation {state} Exception: {ex.Message} {ex.StackTrace}", true);
                return View("ResuItTimeOUT");
            }
            if (keys != null) { keys.Select_ID = ""; }

            return View(keys);
        }

        /// <summary>
        /// Read 選完工站後先檢查 APS_Simulation_ErrorData 異常; 有資料則顯示確認頁, 無則直接進入 MUTIStation.
        /// </summary>
        [HttpPost]
        public ActionResult StationErrorCheck(MUTIStationObj keys)
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            try
            {
                if (keys == null) { keys = new MUTIStationObj(); }
                if (keys.StationNO == null || keys.StationNO.Trim() == "")
                {
                    keys.ERRMsg = "你沒有選擇工站, 請按 回到工站選單畫面 重新選擇工站.";
                    return RedirectToAction("Read");
                }
                keys.StationNO = keys.StationNO.Trim();
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataRow dr_Station = db.DB_GetFirstDataByDataRow($"SELECT StationNO,StationName FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                    if (dr_Station != null) { keys.StationName = dr_Station["StationName"].ToString(); }

                    keys.StationError_List = _SFC_Common.LoadStationPendingErrors(db, keys.StationNO);
                    if (keys.StationError_List == null || keys.StationError_List.Count == 0)
                    {
                        // 查無異常, 直接進入多工站作業畫面
                        if (!GetINFO(db, ref keys))
                        {
                            ViewBag.ErrType = "SystemError";
                            keys.OutError = true;
                            ViewBag.ERRMsg = keys.ERRMsg;
                            ViewBag.Report = "";
                            return View("ResuItTimeOUT");
                        }
                        ViewBag.DATA = keys;
                        keys.Select_ID = "";
                        return View("MUTIStation", keys);
                    }
                }
                ViewBag.DATA = keys;
                return View(keys);
            }
            catch (Exception ex)
            {
                ViewBag.Report = "";
                ViewBag.ErrType = "SystemError";
                ViewBag.ERRMsg = $"程式異常, 導致系統無法使用, 請通知管理者.";
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs StationErrorCheck Exception: {ex.Message} {ex.StackTrace}", true);
                return View("ResuItTimeOUT");
            }
        }

        [HttpPost]
        public ActionResult SetStation_Open(MUTIStationObj keys)
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home");
            }
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                try
                {
                    if (keys == null)
                    {
                        ViewBag.Report = "";
                        ViewBag.ErrType = "SystemError";
                        keys.OutError = true;
                        ViewBag.ERRMsg = "瀏覽器已逾時,請將瀏覽器的葉面關閉後, 再從新開啟瀏覽器再登入."; ;
                        return View("ResuItTimeOUT");
                    }
                    if (keys.StationNO == null || keys.StationNO.Trim() == "")
                    { keys.ERRMsg = "你沒有選擇工站, 請按 回到工站選單畫面 重新選擇工站."; }
                    else
                    {
                        if (keys.Select_ID == null || keys.Select_ID.Trim() == "")
                        {
                            keys.ERRMsg = "必須有選擇工作項目, 請重新選擇.";
                        }
                        else
                        {
                            keys.ERRMsg = "";


                            string[] data = keys.Select_ID.Split(';');
                            DataRow dr_MII = null;
                            string tmp = "";
                            var apsNeedResumeDedupe_Open = new HashSet<string>();
                            foreach (string id in data)
                            {
                                if (id.Trim() != "")
                                {
                                    dr_MII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}' and EndTime is NULL");
                                    if (dr_MII != null)
                                    {
                                        if (dr_MII.IsNull("RemarkTimeS") || (!dr_MII.IsNull("RemarkTimeS") && !dr_MII.IsNull("RemarkTimeE")))
                                        {
                                            #region 送Service處理
                                            DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT RMSName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                                            DataRow dr_M = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");

                                            string status = "1";
                                            string opNO = "";
                                            if (_Fun.GetBaseUser() != null) { opNO = _Fun.GetBaseUser().UserNO; }
                                            if (dr_M == null)
                                            {
                                                keys.ERRMsg = $"{keys.ERRMsg}<br>查無工站 {keys.StationNO} 現場設定(Manufacture),請通知管理者.";
                                                continue;
                                            }

                                            //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                            if (dr != null && _WebSocket.RmsSend(dr["RMSName"].ToString(), 1, $"WebChangeStationStatus,{status},{keys.StationNO},WEBProg,{keys.StationNO},{opNO},{dr_MII["OrderNO"].ToString()},{dr_MII["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
                                            {
                                                if (dr_MII.IsNull("RemarkTimeS")) { tmp = $",RemarkTimeS='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                                                if (dr_MII.IsNull("StartTime")) { tmp = $"{tmp},StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                                                if (db.DB_SetData($"update SoftNetMainDB.[dbo].[ManufactureII] set RemarkTimeE=NULL{tmp} where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}'"))
                                                {
                                                    DataRow dr_SimulationId = _SFC_Common.GetApsSimulationForManufactureII(db, dr_MII);
                                                    if (dr_SimulationId == null)
                                                    { keys.ERRMsg = $"{keys.ERRMsg}<br>{dr_MII["PartNO"]} 找不到對應排程資料,請通知管理者."; continue; }
                                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{dr_SimulationId["SimulationId"].ToString()}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','STView2Work','智慧開工','{dr_MII["PP_Name"].ToString()}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");
                                                    keys.MES_String = "完成工站啟動生產.";
                                                    ApsNeedPauseResumeQueue.TryEnqueueFromManufactureII(db, apsNeedResumeDedupe_Open, dr_MII, "STView2Work_智慧開工");
                                                }
                                            }
                                            else
                                            {
                                                keys.ERRMsg = $"{keys.ERRMsg}<br>後台 Service無作用,無法設定工站狀態,  請通知管理者.";
                                            }
                                            #endregion

                                        }
                                    }
                                }
                            }
                        }
                        if (!GetINFO(db, ref keys))
                        {
                            ViewBag.Report = "";
                            ViewBag.ErrType = "SystemError";
                            keys.OutError = true;
                            ViewBag.ERRMsg = keys.ERRMsg;
                            return View("ResuItTimeOUT");
                        }
                    }
                }
                catch (Exception ex)
                {
                    string state = "";
                    if (keys != null)
                    {
                        state = keys.State;
                        keys.OutError = true;
                    }
                    ViewBag.Report = "";
                    ViewBag.ErrType = "SystemError";
                    ViewBag.ERRMsg = $"程式異常, 導致系統無法使用, 請通知管理者.";
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs SetStation_Open {state} Exception: {ex.Message} {ex.StackTrace}", true);
                    return View("ResuItTimeOUT");
                }
            }
            keys.Select_ID = "";
            ViewBag.DATA = keys;
            return View("MUTIStation");
        }
        [HttpPost]
        public ActionResult SetStation_Stop(MUTIStationObj keys)
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home");
            }
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                try
                {
                    if (keys == null)
                    {
                        ViewBag.Report = "";
                        ViewBag.ErrType = "SystemError";
                        keys.OutError = true;
                        ViewBag.ERRMsg = "瀏覽器已逾時,請將瀏覽器的葉面關閉後, 再從新開啟瀏覽器再登入."; ;
                        return View("ResuItTimeOUT");
                    }
                    if (keys.StationNO == null || keys.StationNO.Trim() == "")
                    { keys.ERRMsg = "你沒有選擇工站, 請按 回到工站選單畫面 重新選擇工站."; }
                    else
                    {
                        if (keys.Select_ID == null || keys.Select_ID.Trim() == "")
                        {
                            keys.ERRMsg = "必須有選擇工作項目, 請重新選擇.";
                        }
                        else
                        {
                            keys.ERRMsg = "";
                            string[] data = keys.Select_ID.Split(';');
                            DataRow dr_MII = null;
                            var apsNeedResumeDedupe_Stop = new HashSet<string>();
                            foreach (string id in data)
                            {
                                if (id.Trim() != "")
                                {
                                    dr_MII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}' and EndTime is NULL ");
                                    if (dr_MII != null)
                                    {
                                        if (!dr_MII.IsNull("StartTime") || !dr_MII.IsNull("RemarkTimeS"))
                                        {
                                            if (dr_MII.IsNull("RemarkTimeE"))
                                            {

                                                #region 送Service處理
                                                DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT RMSName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                                                DataRow dr_M = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");

                                                string status = "2";
                                                string opNO = "";
                                                if (_Fun.GetBaseUser() != null) { opNO = _Fun.GetBaseUser().UserNO; }
                                                if (dr_M == null)
                                                {
                                                    keys.ERRMsg = $"{keys.ERRMsg}<br>查無工站 {keys.StationNO} 現場設定(Manufacture),請通知管理者.";
                                                    continue;
                                                }

                                                //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                                if (dr != null && _WebSocket.RmsSend(dr["RMSName"].ToString(), 1, $"WebChangeStationStatus,{status},{keys.StationNO},WEBProg,{keys.StationNO},{opNO},{dr_MII["OrderNO"].ToString()},{dr_MII["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
                                                {
                                                    if (db.DB_SetData($"update SoftNetMainDB.[dbo].[ManufactureII] set RemarkTimeE='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}'"))
                                                    {
                                                        keys.MES_String = "已停止工站生產作業.";
                                                        DataRow dr_SimulationId = _SFC_Common.GetApsSimulationForManufactureII(db, dr_MII);
                                                        if (dr_SimulationId == null)
                                                        { keys.ERRMsg = $"{keys.ERRMsg}<br>{dr_MII["PartNO"]} 找不到對應排程資料,請通知管理者."; continue; }
                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{dr_SimulationId["SimulationId"].ToString()}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','STView2Work','智慧停工','{dr_MII["PP_Name"].ToString()}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");
                                                        ApsNeedPauseResumeQueue.TryEnqueueFromManufactureII(db, apsNeedResumeDedupe_Stop, dr_MII, "STView2Work_智慧停工");
                                                    }
                                                }
                                                else
                                                {
                                                    keys.ERRMsg = $"{keys.ERRMsg}<br>後台 Service無作用,無法設定工站狀態,  請通知管理者.";
                                                }
                                                #endregion


                                            }
                                        }
                                        else
                                        { keys.ERRMsg = $"{dr_MII["PartNO"].ToString()} 該項目, 未執行開工."; }
                                    }
                                }
                            }
                        }
                        if (!GetINFO(db, ref keys))
                        {
                            ViewBag.Report = "";
                            ViewBag.ErrType = "SystemError";
                            keys.OutError = true;
                            ViewBag.ERRMsg = keys.ERRMsg;
                            return View("ResuItTimeOUT");
                        }
                    }
                }
                catch (Exception ex)
                {
                    string state = "";
                    if (keys != null)
                    {
                        state = keys.State;
                        keys.OutError = true;
                    }
                    ViewBag.Report = "";
                    ViewBag.ErrType = "SystemError";
                    ViewBag.ERRMsg = $"程式異常, 導致系統無法使用, 請通知管理者.";
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs SetStation_Stop {state} Exception: {ex.Message} {ex.StackTrace}", true);
                    return View("ResuItTimeOUT");
                }
            }
            keys.Select_ID = "";
            ViewBag.DATA = keys;
            return View("MUTIStation");
        }
        [HttpPost]
        public ActionResult SetStation_Close(MUTIStationObj keys)
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home");
            }
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                try
                {
                    if (keys == null)
                    {
                        ViewBag.Report = "";
                        ViewBag.ErrType = "SystemError";
                        keys.OutError = true;
                        ViewBag.ERRMsg = "瀏覽器已逾時,請將瀏覽器的葉面關閉後, 再從新開啟瀏覽器再登入."; ;
                        return View("ResuItTimeOUT");
                    }
                    if (keys.StationNO == null || keys.StationNO.Trim() == "")
                    { keys.ERRMsg = "你沒有選擇工站, 請按 回到工站選單畫面 重新選擇工站."; }
                    else
                    {
                        if (keys.Select_ID == null || keys.Select_ID.Trim() == "")
                        {
                            keys.ERRMsg = "必須有選擇工作項目, 請重新選擇.";
                        }
                        else
                        {
                            keys.ERRMsg = "";
                            string[] data = keys.Select_ID.Split(';');
                            DataRow dr_MII = null;
                            DataRow dr = null;
                            string beforePartNO = "";
                            string beforeIndexSN = "";
                            string sql = "";
                            bool Is_Station_Config_Store_Type = false;
                            foreach (string id in data)
                            {
                                if (id.Trim() != "")
                                {
                                    // 每個關站項目獨立判定 status, 避免上一筆最後站把 status 污染成 5 後續誤關/跳過
                                    string status = "4";
                                    Is_Station_Config_Store_Type = false;
                                    dr_MII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}'");
                                    if (dr_MII != null)
                                    {
                                        #region 檢查下一站是否為委外, 若是改停止動作加關閉
                                        if (!_Fun.Config.IsOutPackStationStore && dr_MII["SimulationId"].ToString() != "")
                                        {
                                            dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_MII["SimulationId"].ToString()}'");
                                            if (dr != null)
                                            {
                                                if (_Fun.Config.OutPackStationName == dr["Apply_StationNO"].ToString()) { Is_Station_Config_Store_Type = true; }
                                            }
                                        }
                                        #endregion
                                        if (Is_Station_Config_Store_Type)
                                        {
                                            //下站為加工改停止動作加關閉
                                            if (!dr_MII.IsNull("RemarkTimeS") && dr_MII.IsNull("RemarkTimeE"))
                                            { keys.ERRMsg = $"{keys.ERRMsg}<br>{dr_MII["PartNO"].ToString()} 狀態為 生產中, 無法關閉項目, 該項目請先停工."; }
                                            else
                                            {
                                                DataRow dr_tmp = null;
                                                beforePartNO = dr_MII["PartNO"].ToString();
                                                beforeIndexSN = dr_MII["IndexSN"].ToString();

                                                #region 檢查之前是否無報工
                                                int avgReportTime = 0;
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()}  and PartNO='{beforePartNO}' and ReportTime>5");
                                                if (dr_tmp != null && !dr_tmp.IsNull("AVGTime") && dr_tmp["AVGTime"].ToString().Trim() != "")
                                                {
                                                    avgReportTime = int.Parse(dr_tmp["AVGTime"].ToString());
                                                }
                                                if (avgReportTime > 0)
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={beforeIndexSN} and PartNO='{dr_MII["PartNO"].ToString()}' and LOGDateTime<='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' and OperateType like '%開工%' order by LOGDateTime desc");
                                                    if (dr_tmp != null)
                                                    {
                                                        DateTime tmp_edate = Convert.ToDateTime(dr_tmp["LOGDateTime"]);
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={beforeIndexSN} and PartNO='{dr_MII["PartNO"].ToString()}' and LOGDateTime>'{tmp_edate.ToString("yyyy-MM-dd HH:mm:ss")}' and OperateType like '%報工%'");
                                                        if (dr_tmp == null)
                                                        {
                                                            if ((_SFC_Common.TimeCompute2Seconds(tmp_edate, DateTime.Now)) >= avgReportTime)
                                                            {
                                                                // 僅警告: 前端關站 confirm 已顯示報工不足量, 使用者按「是」後仍可關站
                                                                keys.ERRMsg = $"{keys.ERRMsg}<br>疑似 料號:{dr_MII["PartNO"].ToString()} 未完成報工, 若未報工須請管理者輔助報工, 否則影響下一站生產";
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion

                                                #region 送Service處理
                                                dr = db.DB_GetFirstDataByDataRow($"SELECT RMSName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                                                DataRow dr_M = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");

                                                string opNO = "";
                                                if (_Fun.GetBaseUser() != null) { opNO = _Fun.GetBaseUser().UserNO; }
                                                if (dr_M == null)
                                                {
                                                    keys.ERRMsg = $"{keys.ERRMsg}<br>查無工站 {keys.StationNO} 現場設定(Manufacture),請通知管理者.";
                                                    continue;
                                                }

                                                //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                                if (dr != null && _WebSocket.RmsSend(dr["RMSName"].ToString(), 1, $"WebChangeStationStatus,2,{keys.StationNO},WEBProg,{keys.StationNO},{opNO},{dr_MII["OrderNO"].ToString()},{dr_MII["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
                                                {
                                                    string needID = "";
                                                    dr = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_MII["OrderNO"].ToString()}'");//###??? and WOStatus=0
                                                    if (dr != null)
                                                    {
                                                        needID = dr.IsNull("NeedId") ? "" : dr["NeedId"].ToString();
                                                    }
                                                    DataRow dr_SimulationId = _SFC_Common.GetApsSimulationForManufactureII(db, dr_MII);
                                                    if (dr_SimulationId == null)
                                                    {
                                                        keys.ERRMsg = $"{keys.ERRMsg}<br>{dr_MII["PartNO"]} 找不到對應排程資料(SimulationId={dr_MII["SimulationId"]}),請通知管理者.";
                                                        continue;
                                                    }
                                                    string mii_SimulationId = dr_SimulationId["SimulationId"].ToString();
                                                    if (needID != "")
                                                    {
                                                        db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needID}' and SimulationId='{mii_SimulationId}'");
                                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{needID}' and SimulationId='{mii_SimulationId}'");
                                                    }
                                                    db.DB_SetData(@$"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES 
                                                                        ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{mii_SimulationId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','STView2Work','智慧關站','{dr_MII["PP_Name"].ToString()}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");
                                                    string tmp_s = "";
                                                    if (dr_MII.IsNull("StartTime")) { tmp_s = $",StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                                                    if (dr_MII.IsNull("RemarkTimeS")) { tmp_s = $"{tmp_s},RemarkTimeS='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                                                    if (dr_MII.IsNull("RemarkTimeE")) { tmp_s = $"{tmp_s},RemarkTimeE='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                                                    db.DB_SetData($"update SoftNetMainDB.[dbo].[ManufactureII] set EndTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'{tmp_s} where Id='{id}'");
                                                }
                                                else
                                                {
                                                    keys.ERRMsg = $"{keys.ERRMsg}<br>後台 Service無作用,無法設定工站狀態,  請通知管理者.";
                                                }
                                                #endregion


                                            }
                                        }
                                        else
                                        {
                                            DataRow dr_APS_Simulation = _SFC_Common.GetApsSimulationForManufactureII(db, dr_MII);
                                            if (dr_APS_Simulation == null)
                                            {
                                                keys.ERRMsg = $"{keys.ERRMsg}<br>{dr_MII["PartNO"]} 找不到對應排程資料(SimulationId={dr_MII["SimulationId"]}),請通知管理者.";
                                                continue;
                                            }
                                            string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();
                                            string mii_SimulationId = dr_APS_Simulation["SimulationId"].ToString();

                                            beforePartNO = dr_MII["PartNO"].ToString();
                                            beforeIndexSN = dr_MII["IndexSN"].ToString();

                                            if (!dr_MII.IsNull("RemarkTimeS") && dr_MII.IsNull("RemarkTimeE"))
                                            { keys.ERRMsg = $"{keys.ERRMsg}<br>{dr_MII["PartNO"].ToString()} 狀態為 生產中, 無法關閉項目, 該項目請先停工."; }
                                            else
                                            {
                                                DataRow dr_tmp = null;
                                                DataTable tmp_dt = null;

                                                #region 檢查之前是否無報工
                                                int avgReportTime = 0;
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()}  and PartNO='{beforePartNO}' and ReportTime>5");
                                                if (dr_tmp != null && !dr_tmp.IsNull("AVGTime") && dr_tmp["AVGTime"].ToString().Trim() != "")
                                                {
                                                    avgReportTime = int.Parse(dr_tmp["AVGTime"].ToString());
                                                }
                                                if (avgReportTime > 0)
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={beforeIndexSN} and PartNO='{dr_MII["PartNO"].ToString()}' and LOGDateTime<='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' and OperateType like '%開工%' order by LOGDateTime desc");
                                                    if (dr_tmp != null)
                                                    {
                                                        DateTime tmp_edate = Convert.ToDateTime(dr_tmp["LOGDateTime"]);
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={beforeIndexSN} and PartNO='{dr_MII["PartNO"].ToString()}' and LOGDateTime>'{tmp_edate.ToString("yyyy-MM-dd HH:mm:ss")}' and OperateType like '%報工%'");
                                                        if (dr_tmp == null)
                                                        {
                                                            if ((_SFC_Common.TimeCompute2Seconds(tmp_edate, DateTime.Now)) >= avgReportTime)
                                                            {
                                                                // 僅警告: 前端關站 confirm 已顯示報工不足量, 使用者按「是」後仍可關站
                                                                keys.ERRMsg = $"{keys.ERRMsg}<br>疑似 料號:{dr_MII["PartNO"].ToString()} 未完成報工, 若未報工須請管理者輔助報工, 否則影響下一站生產";
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion

                                                #region 關站處理
                                                string needID = "";
                                                bool isLastStation = false;
                                                if (status == "4")
                                                {
                                                    #region 關站處理
                                                    dr = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_MII["OrderNO"].ToString()}'");//###??? and WOStatus=0
                                                    if (dr != null)
                                                    {
                                                        needID = dr.IsNull("NeedId") ? "" : dr["NeedId"].ToString();
                                                        dr = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and OrderNO='{dr_MII["OrderNO"].ToString()}' and IndexSN='{dr_MII["IndexSN"].ToString()}'");
                                                        if (dr != null)
                                                        {
                                                            isLastStation = bool.Parse(dr["IsLastStation"].ToString());
                                                            List<string> station_list_NO_StationNO_Merge = new List<string>();
                                                            #region 判斷是否最後一站
                                                            if (isLastStation)
                                                            {
                                                                #region 不含合併站
                                                                DataTable dt_station_list_NO_StationNO_Merge = db.DB_GetData($"select Apply_StationNO from SoftNetSYSDB.[dbo].APS_Simulation where NeedId='{needID}' and PartSN>=0 group by Apply_StationNO");
                                                                if (dt_station_list_NO_StationNO_Merge != null && dt_station_list_NO_StationNO_Merge.Rows.Count > 0)
                                                                {
                                                                    foreach (DataRow d in dt_station_list_NO_StationNO_Merge.Rows)
                                                                    {
                                                                        station_list_NO_StationNO_Merge.Add($"'{d["Apply_StationNO"].ToString()}'");
                                                                    }
                                                                }
                                                                #endregion
                                                            }
                                                            else
                                                            {
                                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].APS_Simulation where  NeedId='{needID}' and SimulationId='{mii_SimulationId}'");
                                                                if (dr_tmp != null)
                                                                {
                                                                    station_list_NO_StationNO_Merge.Add($"'{dr_tmp["Apply_StationNO"].ToString()}'");
                                                                }
                                                            }
                                                            #endregion

                                                            if (station_list_NO_StationNO_Merge.Count > 0)
                                                            {
                                                                #region 半成品 or 成品 入庫 與 餘料入庫
                                                                tmp_dt = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Apply_StationNO in ({string.Join(",", station_list_NO_StationNO_Merge)}) and Apply_PP_Name='{dr_MII["PP_Name"].ToString()}' and (Class='4' or Class='5') and Source_StationNO is not null");
                                                                if (tmp_dt != null)
                                                                {
                                                                    foreach (DataRow d in tmp_dt.Rows)
                                                                    {
                                                                        DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                        if (tmp_dr != null)
                                                                        {
                                                                            int qty = int.Parse(tmp_dr["Detail_QTY"].ToString()) - (int.Parse(tmp_dr["Next_StationQTY"].ToString()) + int.Parse(tmp_dr["Next_StoreQTY"].ToString()));
                                                                            if (qty > 0)
                                                                            {
                                                                                #region 成品入庫
                                                                                //最後一站開入庫單 //###???DOCNO暫時寫死BC01
                                                                                string tmp_no = "";
                                                                                string in_StoreNO = "";
                                                                                string in_StoreSpacesNO = "";
                                                                                DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}'");
                                                                                if (tmp == null)
                                                                                {
                                                                                    #region 查找適合庫儲別
                                                                                    _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "BC01");
                                                                                    #endregion
                                                                                    _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref tmp_no, br.UserNO);

                                                                                }
                                                                                else
                                                                                {
                                                                                    in_StoreNO = tmp["StoreNO"].ToString();
                                                                                    in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                                                                    _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref tmp_no, br.UserNO);
                                                                                }
                                                                                sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{tmp_no}',Next_StoreQTY+={qty} where SimulationId='{d["SimulationId"].ToString()}'";
                                                                                db.DB_SetData(sql);
                                                                                #endregion
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                #endregion
                                                            }
                                                            #region 非半成品 or 成品 餘料入庫  //###???暫時寫死 EB01
                                                            /* 改由RUNTimeService 12 處理
                                                            sql = "";
                                                            if (isLastStation)
                                                            { sql = $"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and ((Class!='4' and Class!='5') or NoStation='1')"; }
                                                            else
                                                            {
                                                                List<string> tmp_list = new List<string>();
                                                                DataTable dt_tmp = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and Apply_PP_Name='{dr_MII["PP_Name"].ToString()}' and ((Class!='4' and Class!='5') or Source_StationNO is null)");
                                                                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                                                {
                                                                    foreach (DataRow d in dt_tmp.Rows)
                                                                    { tmp_list.Add($"'{d["SimulationId"].ToString()}'"); }
                                                                    sql = $"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and SimulationId in ({string.Join(",", tmp_list)})";
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
                                                                            tmp_dt = db.DB_GetData($@"SELECT a.*,c.SimulationDate FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                                                            join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.DOCNumberNO,1,4) and b.DOCType='3' and b.ServerId='{_Fun.Config.ServerId}'
                                                                                            join SoftNetSYSDB.[dbo].[APS_Simulation] as c on c.SimulationId=a.SimulationId
                                                                                            where a.SimulationId='{d["SimulationId"].ToString()}' order by OUT_StoreNO,OUT_StoreSpacesNO,IsOK");
                                                                            string docNumberNO = "";
                                                                            foreach (DataRow d2 in tmp_dt.Rows)
                                                                            {
                                                                                if ((int.Parse(d2["QTY"].ToString()) - useQYU) > 0)
                                                                                {
                                                                                    _WebSocket.Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", useQYU, "", d2["Id"].ToString(), $"工單結束退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref docNumberNO, br.UserNO);
                                                                                    break;
                                                                                }
                                                                                else
                                                                                {
                                                                                    _WebSocket.Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", int.Parse(d2["QTY"].ToString()), "", d2["Id"].ToString(), $"工單結束退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref docNumberNO, br.UserNO);
                                                                                    useQYU -= int.Parse(d2["QTY"].ToString());
                                                                                    if (useQYU <= 0) { break; }
                                                                                }
                                                                            }
                                                                            sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{docNumberNO}',Next_StoreQTY+={wQTY} where SimulationId='{d["SimulationId"].ToString()}'";
                                                                            db.DB_SetData(sql);
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            */
                                                            #endregion

                                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsOK='1' where SimulationId='{mii_SimulationId}'");
                                                            if (needID != "") { _SFC_Common.CloseOrphanManufactureIIForNeed(db, needID); }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                #endregion

                                                #region 送Service處理
                                                DataRow dr_M = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                                                dr = db.DB_GetFirstDataByDataRow($"SELECT RMSName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                                                if (dr_M == null)
                                                {
                                                    keys.ERRMsg = $"{keys.ERRMsg}<br>查無工站 {keys.StationNO} 現場設定(Manufacture),請通知管理者.";
                                                    continue;
                                                }
                                                if (isLastStation)
                                                {
                                                    bool chack_state = true;
                                                    //###??? 還有單工, 程式應該還要改
                                                    #region 加判斷所有前站數量是否相符
                                                    tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_MII["OrderNO"].ToString()}'");
                                                    if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                                    {
                                                        foreach (DataRow dr_check in tmp_dt.Rows)
                                                        {
                                                            if (!dr_check.IsNull("RemarkTimeS") && dr_check.IsNull("RemarkTimeE")) { chack_state = false; break; }
                                                        }
                                                    }
                                                    #endregion
                                                    if (chack_state) { status = "5"; }//關站加關工單
                                                }

                                                string opNO = "";
                                                if (_Fun.GetBaseUser() != null) { opNO = _Fun.GetBaseUser().UserNO; }
                                                string type = "";
                                                switch (status)
                                                {
                                                    case "1":
                                                        type = "智慧開工"; break;
                                                    case "2":
                                                        type = "智慧停工"; break;
                                                    case "4":
                                                    case "5":
                                                        type = "智慧關站"; break;
                                                }
                                                DataRow dr_SimulationId = _SFC_Common.GetApsSimulationForManufactureII(db, dr_MII);
                                                if (dr_SimulationId == null)
                                                {
                                                    keys.ERRMsg = $"{keys.ERRMsg}<br>{dr_MII["PartNO"]} 找不到對應排程資料(SimulationId={dr_MII["SimulationId"]}),請通知管理者.";
                                                    continue;
                                                }
                                                //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                                // 先送 Service 成功後才寫 EndTime / OperateLog, 避免後台失敗時現場已結案
                                                if (dr != null && _WebSocket.RmsSend(dr["RMSName"].ToString(), 1, $"WebChangeStationStatus,{status},{keys.StationNO},WEBProg,{keys.StationNO},{opNO},{dr_MII["OrderNO"].ToString()},{dr_MII["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
                                                {
                                                    string tmp_s = "";
                                                    if (dr_MII.IsNull("StartTime")) { tmp_s = $",StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                                                    if (dr_MII.IsNull("RemarkTimeS")) { tmp_s = $"{tmp_s},RemarkTimeS='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                                                    if (dr_MII.IsNull("RemarkTimeE")) { tmp_s = $"{tmp_s},RemarkTimeE='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                                                    db.DB_SetData($"update SoftNetMainDB.[dbo].[ManufactureII] set EndTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'{tmp_s} where Id='{id}'");
                                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{dr_SimulationId["SimulationId"].ToString()}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','STView2Work','{type}','{dr_MII["PP_Name"].ToString()} 系統指派','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");

                                                    if (needID != "")
                                                    {
                                                        if (status == "5")
                                                        {
                                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{needID}'");
                                                            db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needID}'");
                                                        }
                                                        else if (status == "4")
                                                        {
                                                            db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needID}' and SimulationId='{mii_SimulationId}'");
                                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{needID}' and SimulationId='{mii_SimulationId}'");
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    keys.MES_String = $"{keys.MES_String}<br>{keys.StationNO} 後台無作用,無法設定工站狀態, 請通知管理者.";
                                                }
                                                #endregion

                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (!GetINFO(db, ref keys))
                        {
                            ViewBag.Report = "";
                            ViewBag.ErrType = "SystemError";
                            keys.OutError = true;
                            ViewBag.ERRMsg = keys.ERRMsg;
                            return View("ResuItTimeOUT");
                        }
                    }
                }
                catch (Exception ex)
                {
                    string state = "";
                    if (keys != null)
                    {
                        state = keys.State;
                        keys.OutError = true;
                    }
                    ViewBag.Report = "";
                    ViewBag.ErrType = "SystemError";
                    ViewBag.ERRMsg = $"程式異常, 導致系統無法使用, 請通知管理者.";
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs SetStation_Close {state} Exception: {ex.Message} {ex.StackTrace}", true);
                    return View("ResuItTimeOUT");
                }
            }
            keys.Select_ID = "";
            ViewBag.DATA = keys;
            return View("MUTIStation");
        }

        [HttpPost]
        public ActionResult SetStationConfig(string keys) //工站設定   ipport,站1,站2,,,;工單;IndexSN;作業員
        {
            return View();
        }

        [HttpPost]
        public ActionResult SetAction(MUTIStationObj keys)
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home");
            }
            string sql = "";
            keys.ERRMsg = "";
            keys.MES_String = "";
            ViewBag.ErrType = "";
            if (keys == null)
            {
                ViewBag.Report = "";
                ViewBag.ErrType = "SystemError";
                keys.OutError = true;
                ViewBag.ERRMsg = "瀏覽器已逾時,請將瀏覽器的葉面關閉後, 再從新開啟瀏覽器再登入."; ;
                return View("ResuItTimeOUT");
            }
            if (keys.StationNO == null || keys.StationNO.Trim() == "")
            {
                keys.ERRMsg = "你沒有選擇工站 或 畫面逾時, 請按 回到工站選單畫面 重新選擇工站.";
                keys.Select_ID = "";
                ViewBag.DATA = keys;
            }
            else
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    try
                    {
                        DataRow tmp_dr = null;
                        DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                        if (dr_PP_Station == null) { keys.ERRMsg = $"查無 {keys.StationNO} 工站資料紀錄, 請通知管理者."; goto break_FUN; }

                        switch (keys.State)
                        {
                            case "新增項目":
                                {
                                    DataRow dr = null;
                                    //###???1108要改主動維護 State
                                    //if (dr_M["State"].ToString() == "1") { keys.ERRMsg = $"本工站 {keys.StationNO} 已運作中, 請先停止才能設定."; goto break_FUN; }

                                    #region 檢查 SI_OrderNO,SI_IndexSN
                                    string indexSN = "";
                                    string stationNO_Custom_IndexSN = "";
                                    string stationNO_Custom_DisplayName = "";

                                    DataRow sfcdr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{keys.SI_OrderNO}'");
                                    if (sfcdr == null) { keys.ERRMsg = $"查無 {keys.SI_OrderNO} 工單."; goto break_FUN; }
                                    else
                                    {
                                        keys.SI_PP_Name = sfcdr["PP_Name"].ToString(); keys.SI_PartNO = sfcdr["PartNO"].ToString();
                                        if (!sfcdr.IsNull("NeedId") && sfcdr["NeedId"].ToString() != "")
                                        {
                                            if (keys.SI_SimulationId == "") { keys.ERRMsg = $" {keys.SI_OrderNO} 工單, 因程式有誤, 無法選取, 請通知管理者."; goto break_FUN; }
                                        }
                                    }
                                    keys.SI_IndexSN = keys.SI_IndexSN.Trim();
                                    if (keys.SI_IndexSN == "0" || keys.SI_IndexSN == "")
                                    {
                                        DataTable tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and NeedId='{sfcdr["NeedId"].ToString()}' and Source_StationNO='{keys.StationNO}'");
                                        if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                        {
                                            if (tmp_dt.Rows.Count == 1)
                                            {
                                                indexSN = tmp_dt.Rows[0]["Source_StationNO_IndexSN"].ToString(); keys.SI_IndexSN = indexSN; keys.SI_PP_Name = tmp_dt.Rows[0]["PP_Name"].ToString();
                                                stationNO_Custom_IndexSN = tmp_dt.Rows[0]["Source_StationNO_Custom_IndexSN"].ToString();
                                                stationNO_Custom_DisplayName = tmp_dt.Rows[0]["Source_StationNO_Custom_DisplayName"].ToString();
                                            }
                                            else
                                            { keys.ERRMsg = $"{keys.SI_OrderNO} 工單在此站有多重作業,故需指定何製程序號"; goto break_FUN; }
                                        }
                                        else
                                        { keys.ERRMsg = $"查無相關製程序號. IndexSN={keys.SI_IndexSN}, 請通知管理者."; goto break_FUN; }
                                    }
                                    else
                                    {
                                        tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and Apply_PP_Name='{sfcdr["PP_Name"].ToString()}' and (Source_StationNO='{keys.StationNO}' or StationNO_Merge like '%{keys.StationNO},%') and Source_StationNO_IndexSN={keys.SI_IndexSN}");
                                        if (tmp_dr != null)
                                        {
                                            indexSN = tmp_dr["Source_StationNO_IndexSN"].ToString(); keys.SI_PP_Name = tmp_dr["Apply_PP_Name"].ToString();
                                            stationNO_Custom_IndexSN = tmp_dr["Source_StationNO_Custom_IndexSN"].ToString();
                                            stationNO_Custom_DisplayName = tmp_dr["Source_StationNO_Custom_DisplayName"].ToString();
                                        }
                                        else
                                        { keys.ERRMsg = $"查無相關製程順序. IndexSN={keys.SI_IndexSN}, 請通知管理者."; goto break_FUN; }
                                    }

                                    #endregion

                                    //###???若不是SID數量要查BOM,CT時間,料號,品名
                                    #region 查有無需求碼
                                    string partNO = "";
                                    string partName = "";
                                    string typevalue = $"0;";

                                    int ct = 0;
                                    int num = int.Parse(sfcdr["Quantity"].ToString());
                                    if (keys.SI_SimulationId != "")
                                    {
                                        tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{sfcdr["NeedId"].ToString()}' and SimulationId='{keys.SI_SimulationId}'");
                                        if (tmp_dr != null)
                                        {
                                            keys.SI_PartNO = tmp_dr["PartNO"].ToString();
                                            typevalue = $"2;{keys.SI_SimulationId}";
                                            num = _SFC_Common.GetEffectiveNeedQTYBySimulation(db, sfcdr["NeedId"].ToString(), keys.SI_SimulationId, int.Parse(tmp_dr["NeedQTY"].ToString())) + int.Parse(tmp_dr["SafeQTY"].ToString());
                                            ct = int.Parse(tmp_dr["Math_EfficientCT"].ToString());
                                            partNO = tmp_dr["PartNO"].ToString();
                                            tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{tmp_dr["PartNO"].ToString()}'");
                                            if (tmp_dr != null)
                                            {
                                                partName = tmp_dr["PartName"].ToString();
                                            }

                                        }
                                    }
                                    else { typevalue = $"1;{keys.SI_OrderNO}"; }
                                    #endregion


                                    if (keys.SI_PP_Name == null) { keys.SI_PP_Name = ""; }
                                    if (keys.SI_OrderNO == null) { keys.SI_OrderNO = ""; }
                                    if (keys.SI_PartNO == null) { keys.SI_PartNO = ""; }
                                    if (keys.SI_PP_Name == "")
                                    { keys.ERRMsg = $"製程名稱必須要有值, 請重新設定, 請通知管理者."; goto break_FUN; }
                                    else if (keys.SI_OrderNO == "" || keys.SI_PartNO == "")
                                    { keys.ERRMsg = $"工單編號 與 料件編號 不能空白, 請通知管理者."; goto break_FUN; }
                                    if (keys.SI_SimulationId == null || keys.SI_SimulationId == "")
                                    { keys.ERRMsg = $"系統資料有誤, 請通知管理者."; goto break_FUN; }

                                    #region 更新ManufactureII 製造現場狀態
                                    tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and OrderNO='{keys.SI_OrderNO}' and IndexSN={keys.SI_IndexSN} and PP_Name='{keys.SI_PP_Name}' and PartNO='{keys.SI_PartNO}' and EndTime is NULL");
                                    if (tmp_dr == null)
                                    {
                                        string id = _Str.NewId('C');
                                        if (db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[ManufactureII] ([Id],[StationNO],[ServerId],[OrderNO],[Master_PP_Name],[PP_Name],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],[PartNO],[SimulationId],PNQTY)
                                              VALUES ('{id}','{keys.StationNO}','{_Fun.Config.ServerId}','{keys.SI_OrderNO}','{keys.SI_PP_Name}','{keys.SI_PP_Name}',{keys.SI_IndexSN},'{stationNO_Custom_IndexSN}','{stationNO_Custom_DisplayName}','{keys.SI_PartNO}','{keys.SI_SimulationId}',{num})"))
                                        {
                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{keys.SI_OrderNO}' where SimulationId='{keys.SI_SimulationId}'");
                                            /*
                                            if (keys.SI_SimulationId != "")
                                            {
                                                DataRow dr_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] WHERE SimulationId='{keys.SI_SimulationId}'");
                                                if (dr_APS_PartNOTimeNote != null)
                                                {
                                                    if (dr_APS_PartNOTimeNote.IsNull("APS_StationNO") || dr_APS_PartNOTimeNote["APS_StationNO"].ToString().Trim() == "" || dr_APS_PartNOTimeNote["DOCNumberNO"].ToString().Trim() == "")
                                                    {
                                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET APS_StationNO='{keys.StationNO}',DOCNumberNO='{keys.SI_OrderNO}' where SimulationId='{keys.SI_SimulationId}'");
                                                    }
                                                }
                                            }
                                            */
                                            keys.MES_String = "完成 工站設定作業....";
                                        }
                                        else
                                        { keys.ERRMsg = $"寫入資料, 程式異常, 請通知管理者."; goto break_FUN; }
                                    }
                                    else
                                    {
                                        keys.ERRMsg = $"目前已有相同生產項目還未關閉,無法重複相同項目, 請重新設定."; goto break_FUN;
                                    }
                                    #endregion
                                    DataRow dr_SimulationId = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{keys.SI_SimulationId}'");
                                    if (dr_SimulationId == null)
                                    { keys.ERRMsg = $"系統資料有誤, 找不到排程資料(SimulationId={keys.SI_SimulationId}), 請通知管理者."; goto break_FUN; }
                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{keys.SI_SimulationId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','STView2Work','設定工單','{keys.SI_PP_Name}','{keys.StationNO}','{keys.SI_PartNO}','{keys.SI_OrderNO}','{br.UserNO}',{keys.SI_IndexSN})");
                                    ApsNeedPauseResumeQueue.TryEnqueueFromOrderAndSimulation(db, keys.SI_OrderNO, keys.SI_SimulationId, keys.StationNO, "STView2Work_新增項目");

                                    keys.SI_FailQTY = 0;
                                    keys.SI_OKQTY = 0;
                                    keys.SI_OPNO = "";
                                    keys.SI_PP_Name = "";
                                    keys.SI_PartName = "";
                                    keys.SI_IndexSN = "";
                                    keys.SI_OrderNO = "";
                                    keys.SI_PartNO = "";

                                }
                                break;
                            case "報工":
                                string message = "";
                                string stackTrace = "";
                                if (keys.SI_OKQTY<=0 && keys.SI_FailQTY<=0)
                                {
                                    keys.ERRMsg = $"報工數量, 不能良品與不良都為 0."; goto break_FUN;
                                }
                                bool is_reportOK = _SFC_Common.Reporting_STView2Work(db, dr_PP_Station, br.UserNO, keys, false, ref message, ref stackTrace);
                                if (message != "")
                                {
                                    ViewBag.Report = "";
                                    ViewBag.ErrType = "SystemError";
                                    keys.OutError = true;
                                    keys.ERRMsg = $"程式異常, 導致系統無法使用, 請通知管理者.";
                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs 多工報工 Exception: {message} {stackTrace}", true);
                                    return View("ResuItTimeOUT");
                                }
                                else
                                {
                                    if (!is_reportOK) { goto break_FUN; }
                                    ApsNeedPauseResumeQueue.TryEnqueueFromManufactureStation(db, keys.StationNO, "STView2Work_報工");
                                }


                                break;
                            case "發Mail":
                                {
                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"網頁回傳錯誤通知:內容 = {keys.ERRMsg}", true);
                                    ViewBag.ERRMsg = "";
                                    keys.OutError = true;
                                    ViewBag.ErrType = "";
                                    ViewBag.Report = "訊息已發出 Mail 到郵箱.";
                                }
                                return View("ResuItTimeOUT");
                            case "ESOP":
                                {
                                    if (keys.Select_ID == null || keys.Select_ID.Split(';').Length > 2) { keys.ERRMsg = $"查看 ESOP 需先選擇一個項目"; goto break_FUN; }
                                    var sb_id = new StringBuilder();
                                    foreach (string s in keys.Select_ID.Split(';'))
                                    {
                                        if (s.Trim() != "")
                                        {
                                            if (sb_id.Length == 0) { sb_id.Append($"'{s}'"); }
                                            else { sb_id.Append($",'{s}'"); }
                                        }
                                    }
                                    string id = sb_id.ToString();
                                    DataTable tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id in ({id})");
                                    if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                    {
                                        foreach (DataRow dr in tmp_dt.Rows)
                                        {
                                            tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                                            if (tmp_dr != null)
                                            {
                                                tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{tmp_dr["NeedId"].ToString()}' and Apply_PP_Name='{tmp_dr["Apply_PP_Name"].ToString()}' and Apply_StationNO='{tmp_dr["Source_StationNO"].ToString()}' and IndexSN={tmp_dr["Source_StationNO_IndexSN"].ToString()}");
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
                                                                    if (s.Split('.')[(s.Split('.').Length - 1)] == "pdf")
                                                                    {
                                                                        if (sb_key_pdffile.Length == 0) { sb_key_pdffile.Append($"{tmp_dr["PartNO"].ToString().Trim()}/{tmp_dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}/{s}"); }
                                                                        else { sb_key_pdffile.Append($";{tmp_dr["PartNO"].ToString().Trim()}/{tmp_dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}/{s}"); }

                                                                    }
                                                                    else
                                                                    {
                                                                        if (sb_key_file.Length == 0) { sb_key_file.Append($"{tmp_dr["PartNO"].ToString().Trim()}/{tmp_dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}/{s}"); }
                                                                        else { sb_key_file.Append($";{tmp_dr["PartNO"].ToString().Trim()}/{tmp_dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}/{s}"); }
                                                                    }
                                                                }
                                                            }
                                                            ViewBag.ESOP_Files = sb_key_file.ToString();
                                                            ViewBag.ESOP_PDF_Files = sb_key_pdffile.ToString();
                                                        }
                                                    }
                                                    else { ViewBag.ERRMsg = "此製程無ESOP資料."; }
                                                }
                                                else { ViewBag.ERRMsg = "此製程無ESOP資料."; }
                                            }
                                            else { ViewBag.ERRMsg = "此製程無ESOP資料."; }
                                        }
                                    }
                                    else { ViewBag.ERRMsg = "此製程無ESOP資料."; }



                                    return View("DisplayESOP");
                                }
                            case "Knives":
                                {
                                    string meg = "查無刀具或製工具使用歷程";
                                    DataTable dt = db.DB_GetData($@"select b.* from SoftNetMainDB.[dbo].[TotalStockII_Knives] as a
                                                join SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] as b on a.KId=b.KId
                                                where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{keys.StationNO}' and a.IsDel='0' order by KId,LOGDateTime desc,PartNO");
                                    if (dt != null && dt.Rows.Count > 0)
                                    {
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
                                        meg = sb_meg.ToString();
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
                                            meg = $"<label>歷程明細</label>{tmp}<br />{meg}";
                                        }
                                    }
                                    ViewBag.Report = meg;
                                }
                                return View("DisplayKnives");
                            case "領料":
                                {
                                    if (keys.Select_ID == null || keys.Select_ID == "") { keys.ERRMsg = $"領料前, 需先選擇項目."; goto break_FUN; }
                                    string err = "";
                                    DataRow dr_MII = null;
                                    ViewBag.EStore = "";
                                    var apsNeedResumeDedupe_Pick = new HashSet<string>();
                                    foreach (string id in keys.Select_ID.Split(';'))
                                    {
                                        if (id == "") { continue; }
                                        err = "";
                                        dr_MII = db.DB_GetFirstDataByDataRow($"SELECT *  FROM SoftNetMainDB.[dbo].[ManufactureII] where Id='{id}'");
                                        List<string> data = MutiStationSetOutStore(dr_MII, br.UserNO, ref err);//開立單據
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
                                                    ViewBag.EStore = $"<p>倉庫編號:{nametmp}<p/><table id='tableRead' class='table table-bordered xg-table' cellspacing='0'>";
                                                    ViewBag.EStore = $"{ViewBag.EStore}<thead><tr><th>儲位</th><th>料號</th><th>數量</th><th>單位</th><th>單據編號</th></tr></thead><tbody>";
                                                    foreach (DataRow dr in dt.Rows)
                                                    {
                                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' where IsOK='0' and Id='{dr["Id"].ToString()}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and ArrivalDate='{Convert.ToDateTime(dr["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}'");

                                                        if (dr["OUT_StoreNO"].ToString() != nametmp)
                                                        {
                                                            ViewBag.EStore = $"{ViewBag.EStore}</tbody></table>";
                                                            nametmp = dr["OUT_StoreNO"].ToString();
                                                            ViewBag.EStore = $"{ViewBag.EStore}<p>倉庫編號:{nametmp}<p/><table id='tableRead' class='table table-bordered xg-table' cellspacing='0'>";
                                                            ViewBag.EStore = $"{ViewBag.EStore}<thead><tr><th>儲位</th><th>料號</th><th>數量</th><th>單位</th><th>單據編號</th></tr></thead><tbody>";
                                                        }
                                                        ViewBag.EStore = $"{ViewBag.EStore}<tr><th>{dr["OUT_StoreSpacesNO"].ToString()}</th><th>{dr["PartNO"].ToString()}</th><th>{dr["QTY"].ToString()}</th><th>{dr["Unit"].ToString()}</th><th>{dr["DOCNumberNO"].ToString()}</th></tr>";
                                                    }
                                                    ViewBag.EStore = $"{ViewBag.EStore}</tbody></table>";
                                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','STView2Work','智慧領料','{dr_MII["PP_Name"].ToString()}','{dr_MII["StationNO"].ToString()}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");
                                                }
                                            }
                                            else
                                            {
                                                if (dr_MII["SimulationId"].ToString() != "")
                                                {
                                                    DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and SimulationId='{dr_MII["SimulationId"].ToString()}' order by OUT_StoreNO,OUT_StoreSpacesNO,PartNO");
                                                    if (dt != null && dt.Rows.Count > 0)
                                                    {
                                                        string nametmp = dt.Rows[0]["OUT_StoreNO"].ToString();
                                                        ViewBag.EStore = $"<p>倉庫編號:{nametmp}<p/><table id='tableRead' class='table table-bordered xg-table' cellspacing='0'>";
                                                        ViewBag.EStore = $"{ViewBag.EStore}<thead><tr><th>儲位</th><th>料號</th><th>數量</th><th>單位</th><th>單據編號</th></tr></thead><tbody>";
                                                        foreach (DataRow dr in dt.Rows)
                                                        {
                                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' where IsOK='0' and Id='{dr["Id"].ToString()}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and ArrivalDate='{Convert.ToDateTime(dr["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}'");

                                                            if (dr["OUT_StoreNO"].ToString() != nametmp)
                                                            {
                                                                ViewBag.EStore = $"{ViewBag.EStore}</tbody></table>";
                                                                nametmp = dr["OUT_StoreNO"].ToString();
                                                                ViewBag.EStore = $"{ViewBag.EStore}<p>倉庫編號:{nametmp}<p/><table id='tableRead' class='table table-bordered xg-table' cellspacing='0'>";
                                                                ViewBag.EStore = $"{ViewBag.EStore}<thead><tr><th>儲位</th><th>料號</th><th>數量</th><th>單位</th><th>單據編號</th></tr></thead><tbody>";
                                                            }
                                                            ViewBag.EStore = $"{ViewBag.EStore}<tr><th>{dr["OUT_StoreSpacesNO"].ToString()}</th><th>{dr["PartNO"].ToString()}</th><th>{dr["QTY"].ToString()}</th><th>{dr["Unit"].ToString()}</th><th>{dr["DOCNumberNO"].ToString()}</th></tr>";
                                                        }
                                                        ViewBag.EStore = $"{ViewBag.EStore}</tbody></table>";
                                                    }
                                                    else
                                                    { ViewBag.EStore = $"{ViewBag.EStore}<p>工單:{dr_MII["OrderNO"].ToString()} &nbsp;無料可領....."; }
                                                }
                                                else
                                                { ViewBag.EStore = $"{ViewBag.EStore}<p>工單:{dr_MII["OrderNO"].ToString()} &nbsp;無料可領....."; }
                                            }
                                        }
                                        if (err == "" && dr_MII != null)
                                        { ApsNeedPauseResumeQueue.TryEnqueueFromManufactureII(db, apsNeedResumeDedupe_Pick, dr_MII, "STView2Work_領料"); }
                                    }
                                }
                                return View("DisplayEStore");
                            case "入料":
                                {
                                    if (keys.Select_ID == null || keys.Select_ID == "") { keys.ERRMsg = $"入料前, 需先選擇項目."; goto break_FUN; }
                                    DataRow dr_MII = null;
                                    ViewBag.EStore = "";
                                    var apsNeedResumeDedupe_In = new HashSet<string>();
                                    foreach (string id in keys.Select_ID.Split(';'))
                                    {
                                        if (id == "") { continue; }
                                        dr_MII = db.DB_GetFirstDataByDataRow($"SELECT *  FROM SoftNetMainDB.[dbo].[ManufactureII] where Id='{id}'");
                                        if (dr_MII == null) { keys.ERRMsg = $"{keys.ERRMsg}<br>查無選定項目,無法入料."; continue; }
                                        MutiStationSetInStore(dr_MII, br.UserNO);
                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','STView2Work','智慧入料','{dr_MII["PP_Name"].ToString()}','{dr_MII["StationNO"].ToString()}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");
                                        if (dr_MII["SimulationId"].ToString() != "")
                                        {
                                            DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and SimulationId='{dr_MII["SimulationId"].ToString()}' order by OUT_StoreNO,OUT_StoreSpacesNO,PartNO");
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
                                            { ViewBag.EStore = $"{ViewBag.EStore}<p>工單:{dr_MII["OrderNO"].ToString()}&nbsp;無料可入....."; }
                                        }
                                        else
                                        { ViewBag.EStore = $"{ViewBag.EStore}<p>工單:{dr_MII["OrderNO"].ToString()}&nbsp;無料可入....."; }
                                        if (dr_MII != null)
                                        { ApsNeedPauseResumeQueue.TryEnqueueFromManufactureII(db, apsNeedResumeDedupe_In, dr_MII, "STView2Work_入料"); }
                                    }
                                }
                                return View("DisplayEStore");
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Report = "";
                        ViewBag.ErrType = "SystemError";
                        keys.OutError = true;
                        keys.ERRMsg = $"程式異常, 導致系統無法使用, 請通知管理者.";
                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs SetAction Exception: {ex.Message} {ex.StackTrace}", true);
                        return View("ResuItTimeOUT");
                    }
                break_FUN:
                    if (!GetINFO(db, ref keys))
                    {
                        ViewBag.Report = "";
                        ViewBag.ErrType = "SystemError";
                        keys.OutError = true;
                        ViewBag.ERRMsg = keys.ERRMsg;
                        return View("ResuItTimeOUT");
                    }
                    keys.Select_ID = "";
                    ViewBag.DATA = keys;
                }
            }
            return View("MUTIStation");
        }

        private bool GetINFO(DBADO db, ref MUTIStationObj keys)
        {
            try
            {
                string sql = "";
                keys.HasWorkData_List = new List<string[]>();//Id,狀態,PartNO,PartName,Specification,PP_Name,IndexSN,OrderNO,okQTY,failedQty,totCT,目標cT,開始時間,停止時間,生產量

                DataRow keys_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                if (keys_dr == null)
                {
                    keys.ERRMsg = $"查無您選擇的 {keys.StationNO} 工站, 請重新操作.";
                    return false;
                }
                keys.StationName = keys_dr["StationName"].ToString();
                keys.Has_Knives = bool.Parse(keys_dr["IsKnives"].ToString()) ? "1" : "0";

                //###???子製程有問題
                #region 回傳可用工單, 有註冊的OP
                //sql = @$"select b.OrderNO,b.EstimatedStartTime,b.FactoryName,b.LineName,a.PP_Name,b.PartNO,b.PartName,a.DisplaySN,a.IndexSN,a.Station_Custom_IndexSN,a.IndexSN_Merge from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] as a
                //                        join SoftNetSYSDB.[dbo].[PP_WorkOrder] as b on a.PP_Name=b.PP_Name
                //                        where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO = '{keys.StationNO}' and b.EndTime is NULL and b.EstimatedStartTime<='{DateTime.Now.AddMonths(2).ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by b.EstimatedStartTime";

                DateTime comp_CalendarDate = DateTime.Now.AddYears(10);//###??? 提早顯示的時間要參數化
                // 單次查詢：可用＝State=6 且 PauseTime 為 NULL；暫停＝State=P（與 LabelWork 一致）
                sql = $@"select a.DOCNumberNO,a.NeedId,c.State from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] as a 
                                join SoftNetSYSDB.[dbo].[PP_WorkOrder] as b on a.DOCNumberNO=b.OrderNO and b.EndTime is NULL 
                                join SoftNetSYSDB.[dbo].[APS_NeedData] as c on c.Id=a.NeedId and c.ServerId='{_Fun.Config.ServerId}'
                                and (c.State='6' or c.State='P')
                                where (a.StationNO = '{keys.StationNO}' or a.StationNO_Merge like '%{keys.StationNO},%') and a.CalendarDate<='{comp_CalendarDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}' and a.DOCNumberNO!='' 
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
                    //###??? 之後還要補不是NeedID,直接輸入工單的資料 加到inkey
                    Dictionary<string, List<string>> orderOrder = new Dictionary<string, List<string>>();
                    dt_PP = db.DB_GetData($"select DOCNumberNO,SimulationId from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where DOCNumberNO in ({inkey}) and (StationNO = '{keys.StationNO}' or StationNO_Merge like '%{keys.StationNO},%') and CalendarDate<='{comp_CalendarDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by CalendarDate,DOCNumberNO");
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
                    keys.HasWO_List = new List<string[]>();
                    keys.HasWO_List_Paused = new List<string[]>();
                    DataRow dr2 = null;
                    DataRow dr3 = null;
                    Dictionary<string, DataRow> simCache = new Dictionary<string, DataRow>(StringComparer.Ordinal);
                    Dictionary<string, DataRow> materialCache = new Dictionary<string, DataRow>(StringComparer.Ordinal);
                    Dictionary<string, bool> manufactureIICache = new Dictionary<string, bool>(StringComparer.Ordinal);
                    HashSet<string> activeRowDedup = new HashSet<string>(StringComparer.Ordinal);
                    HashSet<string> pausedRowDedup = new HashSet<string>(StringComparer.Ordinal);
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

                                    if (dr2["NeedType"].ToString() == "1") { sType2 = $"訂單需求"; }
                                    else if (dr2["NeedType"].ToString() == "2") { sType2 = $"客戶需求"; }
                                    else if (dr2["NeedType"].ToString() == "5") { sType2 = $"底稿需求"; }
                                    else { sType2 = $"廠內需求"; }
                                    string partNoKey = dr2["PartNO"].ToString();
                                    if (!materialCache.TryGetValue(partNoKey, out dr3))
                                    {
                                        dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNoKey}'");
                                        if (dr3 != null) { materialCache[partNoKey] = dr3; }
                                    }
                                    string mfgIIKey = keys.StationNO + "\u001f" + dr2["DOCNumberNO"].ToString() + "\u001f" + dr2["Source_StationNO_IndexSN"].ToString();
                                    if (!manufactureIICache.TryGetValue(mfgIIKey, out bool isBoundII))
                                    {
                                        isBoundII = db.DB_GetQueryCount($"SELECT Id FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO = '{keys.StationNO}' and OrderNO='{dr2["DOCNumberNO"].ToString()}' and IndexSN='{dr2["Source_StationNO_IndexSN"].ToString()}'") > 0;
                                        manufactureIICache[mfgIIKey] = isBoundII;
                                    }
                                    if (!isBoundII)
                                    {
                                        string rowKey = dr2["DOCNumberNO"].ToString() + "\u001f" + dr2["SimulationId"].ToString() + "\u001f" + dr2["Source_StationNO_IndexSN"].ToString();
                                        string[] rowWo = new string[] { dr2["DOCNumberNO"].ToString(), dr2["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString(), sType2, dr2["CTName"].ToString(), dr2["Apply_PP_Name"].ToString(), dr2["Source_StationNO_IndexSN"].ToString(), dr2["SimulationId"].ToString(), dr2["Source_StationNO_Custom_IndexSN"].ToString(), dr2["Source_StationNO_Custom_DisplayName"].ToString(), dr2["QTY"].ToString() };
                                        if (toActive && activeRowDedup.Add(rowKey)) { keys.HasWO_List.Add(rowWo); }
                                        if (toPaused && pausedRowDedup.Add(rowKey)) { keys.HasWO_List_Paused.Add(new string[] { dr2["DOCNumberNO"].ToString(), dr2["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString(), sType2, dr2["CTName"].ToString(), dr2["Apply_PP_Name"].ToString(), dr2["Source_StationNO_IndexSN"].ToString(), dr2["SimulationId"].ToString(), dr2["Source_StationNO_Custom_IndexSN"].ToString(), dr2["Source_StationNO_Custom_DisplayName"].ToString(), dr2["QTY"].ToString() }); }
                                    }
                                }
                            }
                        }
                    }
                }

                #endregion

                #region 回傳人員名單
                keys.HasPO_List = new List<string[]>();
                sql = @$"select a.*,b.Name from SoftNetMainDB.[dbo].[User] as a join SoftNetMainDB.[dbo].[Dept] as b on a.DeptId=b.Id  where a.ServerId='{_Fun.Config.ServerId}' order by b.Name";
                DataTable dt_User = db.DB_GetData(sql);
                if (dt_User != null && dt_User.Rows.Count > 0)
                {
                    foreach (DataRow d in dt_User.Rows)
                    {
                        keys.HasPO_List.Add(new string[] { d["UserNO"].ToString(), d["Name"].ToString(), d["DeptId"].ToString() });
                    }
                }
                #endregion

                #region 回傳目前工站既有資訊
                string orderdata = "order by PartNO,IndexSN,OrderNO,RemarkTimeS desc";
                if (keys.OrderBY_CMD != null && keys.OrderBY_CMD != "") { orderdata = keys.OrderBY_CMD; }
                DataTable dt = db.DB_GetData($"select *,(select PartName from SoftNetMainDB.[dbo].[Material] as b where b.PartNO=a.PartNO) as PartName from SoftNetMainDB.[dbo].[ManufactureII] as a where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and EndTime is NULL {orderdata}");
                if (dt != null && dt.Rows.Count > 0)
                {
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
                    string woClose_ck = "";
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
                        //tmp_dr2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                        if (tmp_dr2 != null)
                        {
                            tmp_dr2 = db.DB_GetFirstDataByDataRow($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation] where Source_StationNO is not NULL and NeedId='{tmp_dr2["NeedId"].ToString()}' and Apply_PP_Name='{tmp_dr2["Apply_PP_Name"].ToString()}' and IndexSN={(int.Parse(tmp_dr2["IndexSN"].ToString()) - 1).ToString()} and Apply_StationNO='{tmp_dr2["Source_StationNO"].ToString()}'");
                            if (tmp_dr2 != null)
                            {
                                tmp_dr2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].APS_PartNOTimeNote where SimulationId='{tmp_dr2["SimulationId"].ToString()}'");
                                if (tmp_dr2 != null)
                                {
                                    next_StationQTY = tmp_dr2["Detail_QTY"].ToString();
                                }
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
                        if ((okQTY + t_okQTY + failedQty + t_failedQty) >= effectivePNQTY) { woClose_ck = ""; }
                        else { woClose_ck = $"工單:{dr["OrderNO"].ToString()} 需求量:{effectivePNQTY} 不足量:{(effectivePNQTY - okQTY - t_okQTY - failedQty - t_failedQty).ToString()} 數量."; }
                        keys.HasWorkData_List.Add(new string[] { dr["Id"].ToString(), state, partNO, partName, specification, dr["PP_Name"].ToString(), tmp_Index, dr["OrderNO"].ToString(), $"{okQTY.ToString()}+{t_okQTY.ToString()}", $" {failedQty.ToString()}+{t_failedQty.ToString()}", totCT.ToString(), eCT, remarkTimeS, remarkTimeE, effectivePNQTY.ToString(), source_StationNO_Custom_IndexSN, source_StationNO_Custom_DisplayName, next_StationQTY, woClose_ck });
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                keys.ERRMsg = $"程式異常, 導致系統無法使用, 請通知管理者.";
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs GetINFO Exception: {ex.Message} {ex.StackTrace}", true);
                return false;
            }
            return true;
        }

        private List<string> MutiStationSetOutStore(DataRow dr_MII, string userNO, ref string err)//領料
        {
            List<string> re = new List<string>();
            if (dr_MII["StationNO"].ToString() != "" && (dr_MII["OrderNO"].ToString() != "" || dr_MII["SimulationId"].ToString() != ""))
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    List<string> storeNOs = new List<string>();
                    string in_NO02 = "AC01";//###???暫時寫死
                    string sql = "";

                    string for_Apply_StationNO_BY_Main_Source_StationNO = dr_MII["StationNO"].ToString();
                    DataRow dr_WO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_MII["OrderNO"].ToString()}'");
                    if (dr_MII["SimulationId"].ToString().Trim() != "")
                    {
                        #region 先查有無計畫,且有無單據
                        string docNumberNO = "";
                        DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_MII["SimulationId"].ToString()}'");
                        for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();
                        //本站用料明細
                        DataTable dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_WO["NeedId"].ToString()}' and Apply_PP_Name='{dr_WO["PP_Name"].ToString()}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and IndexSN={dr_MII["IndexSN"].ToString()}");
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
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
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
                                                            _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
                                                            tmp_int = 0;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                            if (tmp_01 > 0)
                                                            {
                                                                string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                                _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
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
                                                    _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
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
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                        if (tmp_01 > 0)
                                                        {
                                                            string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                            _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
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
                                                _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
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
                            dt_APS_Simulation = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{dr_WO["NeedId"].ToString()}' and Store_DOCNumberNO is not null and Store_DOCNumberNO !='' and {before_ST}");
                            if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                            {
                                foreach (DataRow d in dt_APS_Simulation.Rows)
                                {
                                    int tmp_int = int.Parse(d["Next_StoreQTY"].ToString());
                                    DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["Store_DOCNumberNO"].ToString()}'");
                                    if (tmp != null)
                                    {
                                        DataRow tmp2 = db.DB_GetFirstDataByDataRow($"select sum(QTY) as qty,StoreNO,StoreSpacesNO from SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and  ServerId='{_Fun.Config.ServerId}' and StoreNO='{tmp["IN_StoreNO"].ToString()}' and StoreSpacesNO='{tmp["IN_StoreSpacesNO"].ToString()}' and  PartNO='{tmp["PartNO"].ToString()}' group by StoreNO,StoreSpacesNO");
                                        if (tmp2 != null && int.Parse(tmp2["qty"].ToString()) >= tmp_int)
                                        {
                                            string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                            _SFC_Common.Create_DOC3stock(db, d, tmp2["StoreNO"].ToString(), tmp2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 原再製,領出生產", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
                                        }
                                        else
                                        {
                                            if (tmp != null) { tmp_int -= int.Parse(tmp["qty"].ToString()); }
                                            #region 扣實體倉
                                            DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and  ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                            {
                                                foreach (DataRow d2 in tmp_dt2.Rows)
                                                {
                                                    if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                    {
                                                        string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 原再製,領出生產", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                        if (tmp_01 > 0)
                                                        {
                                                            string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                            _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"{stationno} 原再製,領出生產", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
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
                                                _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", "AC01", tmp_int, "", "", $"{stationno} 原再製,領出生產", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
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
                    DataRow dr_storeNO = null;
                    foreach (string s in storeNOs)
                    {
                        dr_storeNO = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{s}' and Config_macID!=''");
                        if (dr_storeNO != null)
                        {
                            db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and Keys='領料_{userNO}' and StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and StoreNO='{s}'");
                            db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[BarCode_TMP] (ServerId,[Keys],[StationNO],[StoreNO],[Value],FailTime) VALUES ('{_Fun.Config.ServerId}','領料_{userNO}','{for_Apply_StationNO_BY_Main_Source_StationNO}','{s}','{dr_MII["OrderNO"].ToString()},{dr_MII["IndexSN"].ToString()},{dr_MII["SimulationId"].ToString()}','{DateTime.Now.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss")}')");
                        }
                    }
                }
            }
            return re;
        }

        private void MutiStationSetInStore(DataRow dr_MII, string userNO)//入料
        {
            string sql = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                List<string> storeNOs = new List<string>();
                //###???要判斷是否重複刷8code
                if (dr_MII["SimulationId"].ToString() != "")
                {
                    DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_MII["SimulationId"].ToString()}'");
                    if (dr_APS_Simulation == null) { return; }
                    string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();
                    DataTable dt_APS_PartNOTimeNote = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and SimulationId='{dr_MII["SimulationId"].ToString()}'");
                    if (dt_APS_PartNOTimeNote != null && dt_APS_PartNOTimeNote.Rows.Count > 0)
                    {
                        string in_NO = "BC01";//###??? 暫時寫死單別
                        foreach (DataRow d in dt_APS_PartNOTimeNote.Rows)
                        {
                            int qty = int.Parse(d["Detail_QTY"].ToString()) - (int.Parse(d["Next_StationQTY"].ToString()) + int.Parse(d["Next_StoreQTY"].ToString()));
                            if (qty <= 0) { continue; }
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
                                _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", ref returnID, "工站生產入庫", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetInStore;LabelWorkController", ref tmp_no, userNO, true);
                            }
                            else
                            {
                                in_StoreNO = tmp["StoreNO"].ToString();
                                in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", ref returnID, "工站生產入庫", Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetInStore;LabelWorkController", ref tmp_no, userNO, true);
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
                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and Keys='入料_{userNO}' and StationNO='{dr_MII["StationNO"].ToString()}' and StoreNO='{s}'");
                        db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[BarCode_TMP] (ServerId,[Keys],[StationNO],[StoreNO],[Value],FailTime) VALUES ('{_Fun.Config.ServerId}','入料_{userNO}','{dr_MII["StationNO"].ToString()}','{s}','{dr_MII["OrderNO"].ToString()},{dr_MII["IndexSN"].ToString()},{dr_MII["SimulationId"].ToString()}','{DateTime.Now.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss")}')");
                    }
                }
            }
        }


        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(Ctrl, dt));
        }


        private STView2WorkService EditService()
        {
            return new STView2WorkService(Ctrl);
        }

        //讀取要修改的資料(Get Updated Json)
        [HttpPost]
        public async Task<ContentResult> GetUpdJson(string key)
        {
            return JsonToCnt(await EditService().GetUpdJsonAsync(key));
        }

        //新增(DB)
        public async Task<JsonResult> Create(string json)
        {
            return Json(await EditService().CreateAsync(_Str.ToJson(json)));
        }
        //修改(DB)
        public async Task<JsonResult> Update(string key, string json)
        {
            return Json(await EditService().UpdateAsync(key, _Str.ToJson(json)));
        }

        //刪除(DB)
        public async Task<JsonResult> Delete(string key)
        {
            return Json(await EditService().DeleteAsync(key));
        }





        public ActionResult FGD01()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            MUTIStationObj keys = new MUTIStationObj();
            if (keys.StationNO == null || keys.StationNO.Trim() == "")
            {
                List<string[]> StationNOList = new List<string[]>();
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and Station_Type='8'");
                    DataRow tmp = null;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                            if (tmp == null)
                            { db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Manufacture] ([StationNO],[ServerId],[Config_MutiWO],[Label_ProjectType]) VALUES ('{dr["StationNO"].ToString()}','{_Fun.Config.ServerId}','1','0')"); }
                            else
                            {
                                keys.Station_Config_Store_Type = tmp["Config_Store_Type"].ToString();
                                #region 檢查下一站是否為委外, 設定 Station_Config_Store_Type 網頁行為
                                if (!_Fun.Config.IsOutPackStationStore && tmp["SimulationId"].ToString() != "")
                                {
                                    DataRow dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{tmp["SimulationId"].ToString()}'");
                                    dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr3["NeedId"].ToString()}' and PartNO='{dr3["Master_PartNO"].ToString()}' and Source_StationNO='{dr3["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN={dr3["IndexSN"].ToString()} and (Class='4' or Class='5') and Source_StationNO is not null");
                                    if (dr3 != null)
                                    {
                                        if (_Fun.Config.OutPackStationName == dr3["Apply_StationNO"].ToString()) { keys.Station_Config_Store_Type = "0"; }
                                    }
                                }
                                #endregion

                                if (!bool.Parse(tmp["Config_MutiWO"].ToString()))
                                { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Config_MutiWO='1',Label_ProjectType='0' where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'"); }
                            }
                            StationNOList.Add(new string[] { dr["StationNO"].ToString(), dr["StationName"].ToString() });
                        }
                    }
                }
                ViewBag.StationNOList = StationNOList;
                keys.StationNOList = StationNOList;
            }
            else
            {
                keys.ERRMsg = "";
            }
            return View(keys);
        }

    }
}
