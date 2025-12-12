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
                    DataTable dt_tmp = db.DB_GetData($"SELECT a.*,b.PP_Name FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] as a,SoftNetLogDB.[dbo].[SFC_StationDetail] as b where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{keys.StationNO}' and a.LOGDateTimeID>='{DateTime.Now.AddDays(-14).ToString("MM/dd/yyyy HH:mm:ss.fff")}' and a.LOGDateTime=b.LOGDateTime and a.ServerId=b.ServerId and a.StationNO=b.StationNO order by a.LOGDateTime,a.OP_NO desc");
                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                    {
                        string opName = "";
                        string partName = "";
                        string eCT = "";
                        string avgReportTime = "";
                        report = "<div>前14日歷史報工紀錄</div><div><table data-role='table' data-mode='columntoggle' class='ui-responsive' id='myTable'><thead><tr><th>報工日期</th><th data-priority='1'>製程名稱</th><th>工單編號</th><th>報工人員</th><th>料號</th><th>品名  規格</th><th>報工OK量</th><th>Fail量</th><th data-priority='2'>目標CT</th><th data-priority='3'>實際CT</th><th data-priority='4'>平均報工時間</th><th data-priority='4'>實際花費時間</th></tr></thead><tbody>";
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
                            report = $"{report}<tr><td>{Convert.ToDateTime(d["LOGDateTimeID"]).ToString("yyyy/MM/dd HH:mm:ss")}</td><td>{tmp["PP_Name"].ToString()}</td><td>{tmp["OrderNO"].ToString()}</td><td>{opName}</td><td>{d["PartNO"].ToString()}</td><td>{partName}</td><td>{d["EditFinishedQty"].ToString()}</td><td>{d["EditFailedQty"].ToString()}</td><td>{eCT}</td><td>{d["CycleTime"].ToString()}</td><td>{avgReportTime}</td><td>{d["CycleTime"].ToString()}</td></tr>";
                        }
                    }
                    report = $"{report}</tbody></table></div>";
                    keys.MES_Report = report;
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
                                            if (_Fun.GetBaseUser() != null) { opNO = _Fun.GetBaseUser().UserName; }

                                            //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                            if (dr != null && _WebSocket.RmsSend(dr["RMSName"].ToString(), 1, $"WebChangeStationStatus,{status},{keys.StationNO},WEBProg,{keys.StationNO},{opNO},{dr_MII["OrderNO"].ToString()},{dr_MII["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
                                            {
                                                if (dr_MII.IsNull("RemarkTimeS")) { tmp = $",RemarkTimeS='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                                                if (dr_MII.IsNull("StartTime")) { tmp = $"{tmp},StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                                                if (db.DB_SetData($"update SoftNetMainDB.[dbo].[ManufactureII] set RemarkTimeE=NULL{tmp} where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}'"))
                                                {
                                                    DataRow dr_SimulationId = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_MII["SimulationId"].ToString()}'");
                                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{dr_MII["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2Work','智慧開工','{dr_MII["PP_Name"].ToString()}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");
                                                    keys.MES_String = "完成工站啟動生產.";
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
                                                if (_Fun.GetBaseUser() != null) { opNO = _Fun.GetBaseUser().UserName; }

                                                //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                                if (dr != null && _WebSocket.RmsSend(dr["RMSName"].ToString(), 1, $"WebChangeStationStatus,{status},{keys.StationNO},WEBProg,{keys.StationNO},{opNO},{dr_MII["OrderNO"].ToString()},{dr_MII["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
                                                {
                                                    if (db.DB_SetData($"update SoftNetMainDB.[dbo].[ManufactureII] set RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}'"))
                                                    {
                                                        keys.MES_String = "已停止工站生產作業.";
                                                        DataRow dr_SimulationId = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_MII["SimulationId"].ToString()}'");
                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{dr_MII["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2Work','智慧停工','{dr_MII["PP_Name"].ToString()}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");
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
                            string status = "4";
                            string sql = "";
                            bool Is_Station_Config_Store_Type = false;
                            foreach (string id in data)
                            {
                                if (id.Trim() != "")
                                {
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

                                                #region 檢查之前是否無報工
                                                int avgReportTime = 0;
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()}  and PartNO='{beforePartNO}' and ReportTime>5");
                                                if (dr_tmp != null && !dr_tmp.IsNull("AVGTime") && dr_tmp["AVGTime"].ToString().Trim() != "")
                                                {
                                                    avgReportTime = int.Parse(dr_tmp["AVGTime"].ToString());
                                                }
                                                if (avgReportTime > 0)
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={beforeIndexSN} and PartNO='{dr_MII["PartNO"].ToString()}' and LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%開工%' order by LOGDateTime desc");
                                                    if (dr_tmp != null)
                                                    {
                                                        DateTime tmp_edate = Convert.ToDateTime(dr_tmp["LOGDateTime"]);
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={beforeIndexSN} and PartNO='{dr_MII["PartNO"].ToString()}' and LOGDateTime>'{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%報工%'");
                                                        if (dr_tmp == null)
                                                        {
                                                            if ((_SFC_Common.TimeCompute2Seconds(tmp_edate, DateTime.Now)) >= avgReportTime)
                                                            {
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
                                                if (_Fun.GetBaseUser() != null) { opNO = _Fun.GetBaseUser().UserName; }

                                                //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                                if (dr != null && _WebSocket.RmsSend(dr["RMSName"].ToString(), 1, $"WebChangeStationStatus,2,{keys.StationNO},WEBProg,{keys.StationNO},{opNO},{dr_MII["OrderNO"].ToString()},{dr_MII["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
                                                {
                                                    string needID = "";
                                                    dr = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_MII["OrderNO"].ToString()}'");//###??? and WOStatus=0
                                                    if (dr != null)
                                                    {
                                                        needID = dr.IsNull("NeedId") ? "" : dr["NeedId"].ToString();
                                                    }
                                                    if (needID != "")
                                                    {
                                                        db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needID}' and SimulationId='{dr_MII["SimulationId"].ToString()}'");
                                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{needID}' and SimulationId='{dr_MII["SimulationId"].ToString()}'");
                                                    }
                                                    DataRow dr_SimulationId = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_MII["SimulationId"].ToString()}'");
                                                    db.DB_SetData(@$"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES 
                                                                        ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{dr_MII["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2Work','智慧關站','{dr_MII["PP_Name"].ToString()}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");
                                                    string tmp_s = "";
                                                    if (dr_MII.IsNull("StartTime")) { tmp_s = $",StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                    if (dr_MII.IsNull("RemarkTimeS")) { tmp_s = $"{tmp_s},RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                    if (dr_MII.IsNull("RemarkTimeE")) { tmp_s = $"{tmp_s},RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                    db.DB_SetData($"update SoftNetMainDB.[dbo].[ManufactureII] set EndTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'{tmp_s} where Id='{id}'");
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
                                            DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_MII["SimulationId"].ToString()}'");
                                            string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();

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
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={beforeIndexSN} and PartNO='{dr_MII["PartNO"].ToString()}' and LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%開工%' order by LOGDateTime desc");
                                                    if (dr_tmp != null)
                                                    {
                                                        DateTime tmp_edate = Convert.ToDateTime(dr_tmp["LOGDateTime"]);
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={beforeIndexSN} and PartNO='{dr_MII["PartNO"].ToString()}' and LOGDateTime>'{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%報工%'");
                                                        if (dr_tmp == null)
                                                        {
                                                            if ((_SFC_Common.TimeCompute2Seconds(tmp_edate, DateTime.Now)) >= avgReportTime)
                                                            {
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
                                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].APS_Simulation where  NeedId='{needID}' and SimulationId='{dr_MII["SimulationId"].ToString()}'");
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

                                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsOK='1' where SimulationId='{dr_MII["SimulationId"].ToString()}'");
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                #endregion

                                                #region 送Service處理
                                                DataRow dr_M = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                                                dr = db.DB_GetFirstDataByDataRow($"SELECT RMSName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
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
                                                string tmp_s = "";
                                                if (dr_MII.IsNull("StartTime")) { tmp_s = $",StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                if (dr_MII.IsNull("RemarkTimeS")) { tmp_s = $"{tmp_s},RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                if (dr_MII.IsNull("RemarkTimeE")) { tmp_s = $"{tmp_s},RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                db.DB_SetData($"update SoftNetMainDB.[dbo].[ManufactureII] set EndTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'{tmp_s} where Id='{id}'");

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
                                                DataRow dr_SimulationId = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_MII["SimulationId"].ToString()}'");
                                                db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{dr_MII["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2Work','{type}','{dr_MII["PP_Name"].ToString()} 系統指派','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");

                                                //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                                if (dr != null && _WebSocket.RmsSend(dr["RMSName"].ToString(), 1, $"WebChangeStationStatus,{status},{keys.StationNO},WEBProg,{keys.StationNO},{opNO},{dr_MII["OrderNO"].ToString()},{dr_MII["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
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
                                                            db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needID}' and SimulationId='{dr_MII["SimulationId"].ToString()}'");
                                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{needID}' and SimulationId='{dr_MII["SimulationId"].ToString()}'");
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
                                        keys.SI_PP_Name = sfcdr["PP_Name"].ToString(); keys.SI_PartNO= sfcdr["PartNO"].ToString();
                                        if (!sfcdr.IsNull("NeedId") && sfcdr["NeedId"].ToString()!="")
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
                                            num = int.Parse(tmp_dr["NeedQTY"].ToString()) + int.Parse(tmp_dr["SafeQTY"].ToString());
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
                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{keys.SI_SimulationId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2Work','新增項目','{keys.SI_PP_Name}','{keys.StationNO}','{keys.SI_PartNO}','{keys.SI_OrderNO}','{br.UserNO}',{keys.SI_IndexSN})");

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
                                bool is_reportOK= _SFC_Common.Reporting_STView2Work(db, dr_PP_Station, br.UserNO, keys,false, ref message, ref stackTrace);
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
                                }

                                /*
                                {
                                    #region 檢查
                                    if (keys.Select_ID == null || keys.SI_Slect_OPNOs == null || keys.Select_ID == "" || keys.SI_Slect_OPNOs == "") { keys.ERRMsg = $"報工前, 需先選擇項目 或 操作員."; goto break_FUN; }
                                    if (keys.SI_OKQTY == 0 && keys.SI_FailQTY == 0) { keys.ERRMsg = $"數量均無值, 請重新輸入."; goto break_FUN; }
                                    string id = "";
                                    foreach (string s in keys.Select_ID.Split(';'))
                                    {
                                        if (s.Trim() != "")
                                        {
                                            if (id == "") { id = s; }
                                            else { keys.ERRMsg = $"報工 一次只能選擇一個項目, 請重新選擇."; goto break_FUN; }
                                        }
                                    }
                                    DataRow dr_WO = null;
                                    DataRow dr_MII = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}' and EndTime is NULL");
                                    if (dr_MII == null) { keys.ERRMsg = $"該項目已關閉生產, 無法報工."; goto break_FUN; }
                                    if (dr_MII.IsNull("StartTime")) { keys.ERRMsg = $"項目未設定開工過, 無法報工, 請先按開工."; goto break_FUN; }
                                    if (dr_MII.IsNull("RemarkTimeS")) { keys.ERRMsg = $"項目未設定開工過, 無法報工, 請先按開工."; goto break_FUN; }
                                    keys.SI_SimulationId = dr_MII["SimulationId"].ToString();
                                    keys.SI_IndexSN = dr_MII["IndexSN"].ToString();
                                    keys.SI_OrderNO = dr_MII["OrderNO"].ToString();
                                    keys.SI_PP_Name = dr_MII["PP_Name"].ToString();
                                    keys.SI_PartNO = dr_MII["PartNO"].ToString();
                                    #endregion

                                    DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{keys.SI_SimulationId}'");
                                    DateTime startTime = DateTime.Now;
                                    if (keys.SI_OKQTY > 0 || keys.SI_FailQTY > 0)
                                    {
                                        string[] data = new string[7] { keys.StationNO, keys.SI_OrderNO, keys.SI_IndexSN, keys.SI_OKQTY.ToString(), keys.SI_FailQTY.ToString(), keys.SI_Slect_OPNOs, keys.LocalIPPort };

                                        #region 檢查網頁來源資料
                                        dr_WO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}'");
                                        if (dr_WO == null) { keys.ERRMsg = $"查無 {data[0]} 工單資料紀錄."; goto break_FUN; }
                                        sql = $"SELECT * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}' AND StationNO='{data[0]}' AND IndexSN={data[2]}";
                                        DataRow dr = db.DB_GetFirstDataByDataRow(sql);
                                        if (dr == null && dr_APS_Simulation!=null) 
                                        {
                                            string isend = dr_APS_Simulation["PartSN"].ToString().Trim() == "0" ? "1" : "0";
                                            string is_IndexSN_Merge = dr_APS_Simulation.IsNull("StationNO_Merge") ? "1" : "0";
                                            string sfc_sql = $@"INSERT INTO SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement (
                                                                                    [OrderNO],[StationNO],[StationName],[PP_Name],[IndexSN],[DisplaySN],
                                                                                    [IsLastStation],[Sub_PP_Name],[StatndCycleTime],[UpdateTime],[IndexSN_Merge],
                                                                                    [StartTime],[CumulativeTime],[AvarageCycleTime],[TotalCheckIn],[TotalCheckOut],
                                                                                    [TotalInput],[TotalOutput],[TotalFail],[TotalKeep],[FPY],
                                                                                    [YieldRate],[StationYieldRate],ServerId) VALUES 
                                                                                    ('{data[1]}','{data[0]}','{dr_PP_Station["StationName"].ToString()}','{dr_MII["PP_Name"].ToString()}',
                                                                                    {data[2]},0,'{isend}','{dr_MII["PP_Name"].ToString()}',{dr_APS_Simulation["Math_StandardCT"].ToString()},'{DateTime.Now.ToString("MM/dd/yyyy H:mm:ss")}',
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
                                        if (dr != null && !dr.IsNull("StartTime") && dr["StartTime"].ToString().Trim() != "")
                                        { startTime = Convert.ToDateTime(dr["StartTime"]); }
                                        #endregion

                                        int outQTY = keys.SI_OKQTY;
                                        int failQTY = keys.SI_FailQTY;
                                        decimal ct = 0;
                                        DateTime rRemarkTimeS = Convert.ToDateTime(dr_MII["RemarkTimeS"]);

                                        #region 計算CT

                                        //檢查rRemarkTimeS是否需要更改適合的時間, 查上一次報工的時間
                                        string select_OP = "";
                                        if (keys.SI_Slect_OPNOs.Split(';').Length > 1) 
                                        { 
                                            foreach(string s in keys.SI_Slect_OPNOs.Split(';'))
                                            {
                                                if (select_OP == "") { select_OP = $"OP_NO like '%{s}%'"; }
                                                else { select_OP = $"{select_OP} or OP_NO like '%{s}%'"; }
                                            }
                                            select_OP = $"({select_OP})";
                                        }
                                        else { select_OP = $"OP_NO like '%{keys.SI_Slect_OPNOs}%'"; }
                                        tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={keys.SI_IndexSN} and OrderNO='{keys.SI_OrderNO}' and OperateType like '%報工%' and {select_OP}  order by LOGDateTime desc");
                                        if (tmp_dr != null && Convert.ToDateTime(tmp_dr["LOGDateTime"]) > rRemarkTimeS) 
                                        { 
                                            rRemarkTimeS = Convert.ToDateTime(tmp_dr["LOGDateTime"]); 
                                        }
                                        if (dr_MII.IsNull("RemarkTimeE") || rRemarkTimeS >= Convert.ToDateTime(dr_MII["RemarkTimeE"]))
                                        {
                                            ct = _WebSocket.TimeCompute2Seconds(rRemarkTimeS, DateTime.Now) / (keys.SI_OKQTY + keys.SI_FailQTY);
                                            if (ct <= 0 && !dr_MII.IsNull("RemarkTimeE") && Convert.ToDateTime(dr_MII["RemarkTimeE"]) > Convert.ToDateTime(dr_MII["RemarkTimeS"]))
                                            {
                                                ct = _WebSocket.TimeCompute2Seconds(Convert.ToDateTime(dr_MII["RemarkTimeS"]), Convert.ToDateTime(dr_MII["RemarkTimeE"])) / (keys.SI_OKQTY + keys.SI_FailQTY);
                                            }
                                        }
                                        else
                                        { 
                                            ct = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), rRemarkTimeS, Convert.ToDateTime(dr_MII["RemarkTimeE"])) / (keys.SI_OKQTY + keys.SI_FailQTY);
                                            if (ct <= 0 && Convert.ToDateTime(dr_MII["RemarkTimeE"]) > rRemarkTimeS)
                                            {
                                                ct = _WebSocket.TimeCompute2Seconds(rRemarkTimeS, Convert.ToDateTime(dr_MII["RemarkTimeE"])) / (keys.SI_OKQTY + keys.SI_FailQTY);
                                            }
                                        }
                                        decimal ct_log = ct < 1 ? 0 : ct;
                                        #endregion

                                        int ops = keys.SI_Slect_OPNOs.Split(';').Length;
                                        if (ops > 1) { ct = ct / ops; }

                                        #region 寫SFC_StationDetail
                                        string logTime = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff");
                                        string lOGDateTime = logTime;
                                        string old_InTime = "";
                                        string old_OutTime = "";
                                        string old_ProductFinishedQty = "0";
                                        string old_ProductFailedQty = "0";
                                        int OP_Count = data[5].Split(';').Count() + 1;
                                        if (OP_Count <= 0) { OP_Count = 1; }
                                        DataRow dr_StationDetail = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail] WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO = '{data[1]}' AND StationNO = '{data[0]}' AND IndexSN={data[2]}");
                                        if (dr_StationDetail == null)
                                        {
                                            //###???PP_Name暫時
                                            old_InTime = startTime.ToString("yyyy-MM-dd HH:mm:ss.fff");
                                            old_OutTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
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
                                                data[5],
                                                data[0],
                                                data[2],
                                                dr["IndexSN_Merge"].ToString(),
                                                data[1],
                                                dr_MII["PartNO"].ToString(),
                                                old_InTime,
                                                old_OutTime,
                                                ct.ToString(),
                                                (outQTY + failQTY) > 0 ? 1 : 0,
                                                outQTY > 0 ? 1 : 0,
                                                failQTY > 0 ? 1 : 0,
                                                dr_PP_Station["Station_Type"].ToString(),
                                                outQTY.ToString(),
                                                failQTY.ToString(), dr_PP_Station["RMSName"].ToString(), keys.SI_SimulationId, _Fun.Config.ServerId);
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

                                            sql = string.Format(
                                                    @"UPDATE SoftNetLogDB.[dbo].[SFC_StationDetail] 
                                                SET [ProductFinishedQty]+={0}, [ProductFailedQty]+={1},
                                                [InTime]='{2}',[OutTime]='{3}',[CycleTime]={4} 
                                                WHERE ServerId='{9}' and OrderNO = '{5}' AND StationNO = '{6}' AND IndexSN={7} AND LOGDateTime = '{8}'",
                                                    outQTY,
                                                    failQTY,
                                                    startTime.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                                    ct.ToString(),
                                                    data[1],
                                                    data[0], data[2], logTime, _Fun.Config.ServerId);
                                        }
                                        #endregion

                                        if (db.DB_SetData(sql))
                                        {

                                            #region log SFC_StationDetail_ChangeLOG紀錄
                                            int reportTime = 0;
                                            if (dr_StationDetail != null)
                                            {

                                                #region 計算上一次報工與現在時間差
                                                DataRow d2 = db.DB_GetFirstDataByDataRow($"SELECT LOGDateTimeID FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(dr_StationDetail["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by LOGDateTime,LOGDateTimeID desc");
                                                if (d2 != null)
                                                { reportTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(d2["LOGDateTimeID"]), DateTime.Now); }
                                                #endregion
                                            }
                                            string wsid = "NULL";
                                            if (keys.SI_SimulationId != "") { wsid = $"'{keys.SI_SimulationId}'"; }
                                            string partNO = keys.SI_PartNO;
                                            string LOGDateTimeID = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                                            #region 取得PP_EfficientDetail紀錄
                                            string eCT = "0";
                                            string upperCT = "0";
                                            string lowerCT = "0";
                                            DataRow dr_tmp_ct = db.DB_GetFirstDataByDataRow($"select AVG(EfficientCycleTime) as ECT,AVG(SD_UpperLimit) as UpperCT,AVG(SD_LowerLimit) as LowerCT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StationNO='{data[0]}' and PP_Name='{dr["PP_Name"].ToString()}' and IndexSN={dr["IndexSN"].ToString()} and DOCNO=''");
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
                                                string _s = "";
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
                                                    where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PartNO='{keys.SI_PartNO}' and PP_Name='{dr_WO["PP_Name"].ToString()}' and IndexSN={data[2]} and EditFinishedQty!=0 and CycleTime!=0");
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
                                                            _WebSocket.SfcTimerloopthread_Tick_Efficient(db, allCT, keys.StationNO, dr_WO["PP_Name"].ToString(), keys.SI_PP_Name, data[2], dr_WO["PartNO"].ToString(), keys.SI_PartNO, "");
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
                                            DataTable tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStockII_Knives] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IsDel='0'");
                                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                            {
                                                int useTime = 0;
                                                int useCount = outQTY + failQTY;
                                                string k_stime = "";
                                                foreach (DataRow d in tmp_dt.Rows)
                                                {
                                                    if (!dr_MII.IsNull("RemarkTimeS"))
                                                    {
                                                        if (d.IsNull("StartTime")) { k_stime = $",StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}'"; } else { k_stime = ""; }
                                                        if (!dr_MII.IsNull("RemarkTimeE"))
                                                        {
                                                            tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PartNO='{dr_MII["PartNO"].ToString()}' and LOGDateTime>'{Convert.ToDateTime(dr_MII["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}' and OperateType like '%報工%'");
                                                            if (tmp_dr != null)
                                                            {
                                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                            ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}',{(useCount).ToString()},0)");
                                                            }
                                                            else
                                                            {
                                                                useTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(dr_MII["RemarkTimeS"].ToString()), DateTime.Now);
                                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                            ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}',{(useCount).ToString()},{useTime.ToString()})");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            useTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(dr_MII["RemarkTimeS"].ToString()), DateTime.Now);
                                                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                        ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}',{(useCount).ToString()},{useTime.ToString()})");
                                                        }
                                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStockII_Knives] set TOTWorkTime+={useTime.ToString()},TOTCount+={useCount.ToString()}{k_stime} where ServerId='{_Fun.Config.ServerId}' and KId='{d["KId"].ToString()}'");
                                                    }
                                                }
                                            }
                                            #endregion

                                            _WebSocket.Update_PP_WorkOrder_Settlement(db, data[1], keys.SI_SimulationId);

                                            //###??? 不良數量尚未處裡
                                            if (keys.SI_SimulationId != "")
                                            {
                                                string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();
                                                bool isNeedQTY_OK = false;//判斷本站數量已足夠
                                                string in_NO = "AC01";//###??? 暫時寫死領料單別
                                                string inOK_NO = "BC01";//###??? 暫時寫死入庫單別

                                                //本階報工當筆資料寫入
                                                DataRow dr_APS_PartNOTimeNote = null;
                                                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{data[1].Trim()}',Detail_QTY+={data[3]},Detail_Fail_QTY+={data[4]} where SimulationId='{keys.SI_SimulationId}'"))
                                                {
                                                    dr_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SI_SimulationId}'");
                                                    if ((int.Parse(dr_APS_PartNOTimeNote["Detail_QTY"].ToString()) + int.Parse(dr_APS_PartNOTimeNote["Detail_Fail_QTY"].ToString()) - int.Parse(dr_APS_PartNOTimeNote["NeedQTY"].ToString())) >= 0)
                                                    { isNeedQTY_OK = true; }
                                                }

                                                //尋找相關BOM原物料
                                                #region 由本階工站,查上一階 扣Keep量 與 處理領料單單據 與若上階有先入庫,要先領出
                                                DataTable dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_APS_Simulation["NeedId"].ToString()}' and Apply_PP_Name='{keys.SI_PP_Name}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and IndexSN={data[2]} order by PartSN desc");
                                                if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                                                {
                                                    string docNumberNO = "";
                                                    foreach (DataRow d in dt_APS_Simulation.Rows)
                                                    {
                                                        #region 上一階是工站, 處裡移轉量 APS_PartNOTimeNote
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
                                                                    #region 有計畫先扣Keep量  先by StoreOrder順序扣, 在預開入庫單 
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
                                                                                _WebSocket.Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmp_no, br.UserNO);
                                                                                tmp_int = 0;
                                                                                break;
                                                                            }
                                                                            else
                                                                            {
                                                                                int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                                _WebSocket.Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_01, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmp_no, br.UserNO);
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
                                                                                _WebSocket.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmp_no, br.UserNO);
                                                                            }
                                                                            else
                                                                            {
                                                                                #region 查找適合庫儲別
                                                                                _WebSocket.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                                                                #endregion
                                                                                _WebSocket.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmp_no, br.UserNO);
                                                                            }
                                                                        }
                                                                        #endregion
                                                                    }
                                                                    #endregion
                                                                }
                                                                else
                                                                {
                                                                    #region 查找適合庫儲別
                                                                    _WebSocket.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                                                    #endregion
                                                                    #region 無倉紀錄, 加空倉
                                                                    _WebSocket.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, int.Parse(data[3]), "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmp_no, br.UserNO);
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
                                                                                _WebSocket.Create_DOC3stock(db, d, tmp_store["IN_StoreNO"].ToString(), tmp_store["IN_StoreSpacesNO"].ToString(), "", "", in_NO, wrQTY, "", "", $"{stationno} 入庫後再領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmpDOCNO, br.UserNO);
                                                                            }
                                                                        }
                                                                        if (wr_next_StationQTY>0)
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
                                                                DataRow dr_next_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SI_SimulationId}'");
                                                                next_next_detail_QTY = int.Parse(dr_next_APS_PartNOTimeNote["Detail_QTY"].ToString());
                                                                if (next_next_detail_QTY>0)
                                                                {
                                                                    //detail_QTY - next_StationQTY=上一階數量
                                                                    if ((next_StationQTY + tmp_int) < next_next_detail_QTY) { next_next_detail_QTY -= (next_StationQTY + tmp_int); }
                                                                    else { next_next_detail_QTY = 0; }
                                                                }
                                                                #endregion

                                                                if ((detail_QTY - next_StationQTY) >= tmp_int)
                                                                {
                                                                    //先在製移轉 
                                                                    sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Next_APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}',Next_StationQTY+={(tmp_int+ next_next_detail_QTY).ToString()} where SimulationId='{d["SimulationId"].ToString()}'";
                                                                    if (db.DB_SetData(sql))
                                                                    {
                                                                        #region 處理工站移轉時間

                                                                        #endregion
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    bool is_run = true;
                                                                    //將在製 剩餘移轉完
                                                                    tmp_int -= (detail_QTY - next_StationQTY);
                                                                    sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Next_APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}',Next_StationQTY={tmp["Detail_QTY"].ToString()} where SimulationId='{d["SimulationId"].ToString()}'";
                                                                    if (db.DB_SetData(sql) && int.Parse(tmp["Detail_QTY"].ToString()) > 0)
                                                                    {
                                                                        #region 處理工站移轉時間

                                                                        #endregion
                                                                    }

                                                                    #region 先檢查是否已有相關單據(已領過AC01), 且已移轉多少量
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
                                                                                        _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
                                                                                        tmp_int = 0;
                                                                                        break;
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                                                        //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_01} where Id='{d2["Id"].ToString()}'");
                                                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                        _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
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
                                                                                            _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
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
                                                                                                _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
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
                                                                                    _WebSocket.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
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
                                                                                        _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
                                                                                        tmp_int = 0;
                                                                                        break;
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                                        if (tmp_01 != 0)
                                                                                        {
                                                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                            _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
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
                                                                                _WebSocket.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
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
                                                            #region 上一階是原物料 扣庫存帳 TotalStock,TotalStockII
                                                            int tmp_int = (int.Parse(d["BOMQTY"].ToString()) * outQTY);

                                                            #region 先檢查是否已有單據, 且已移轉多少量
                                                            int detailQTY = tmp_int;
                                                            int stockQTY = 0;
                                                            DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO!=''");
                                                            if (tmp != null)
                                                            {
                                                                docNumberNO = tmp["DOCNumberNO"].ToString();
                                                                detailQTY += (int.Parse(tmp["Detail_QTY"].ToString()) + int.Parse(tmp["Next_StationQTY"].ToString()) - int.Parse(tmp["Next_StoreQTY"].ToString()));
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
                                                                                _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
                                                                                tmp_int = 0;
                                                                                break;
                                                                            }
                                                                            else
                                                                            {
                                                                                int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
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
                                                                                    _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
                                                                                    tmp_int = 0;
                                                                                    break;
                                                                                }
                                                                                else
                                                                                {
                                                                                    int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                                    if (tmp_01 != 0)
                                                                                    {
                                                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                        _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
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
                                                                            _WebSocket.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
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
                                                                                    _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
                                                                                    tmp_int = 0;
                                                                                    break;
                                                                                }
                                                                                else
                                                                                {
                                                                                    int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                                    if (tmp_01 != 0)
                                                                                    {
                                                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                        _WebSocket.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
                                                                                        tmp_int -= tmp_01;
                                                                                    }
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
                                                                        _WebSocket.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, br.UserNO);
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
                                                            _WebSocket.SelectINStore(db, tmp["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "PA02", true);
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
                                                                if (d2.IsNull("StartTime") || (!d2.IsNull("StartTime") && Convert.ToDateTime(d2["StartTime"]) < DateTime.Now))
                                                                {
                                                                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkingPaper] set StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where Id='{d2["Id"].ToString()}'");
                                                                }
                                                            }
                                                        }
                                                    }

                                                }
                                                #endregion

                                                #region 消除 APS_WorkTimeNote 工站負荷
                                                DataRow tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SI_SimulationId}'");
                                                if (tmp_del != null && int.Parse(tmp_del["Time1_C"].ToString()) == 1 && int.Parse(tmp_del["Time2_C"].ToString()) == 0 && int.Parse(tmp_del["Time3_C"].ToString()) == 0 && int.Parse(tmp_del["Time4_C"].ToString()) == 0)
                                                { }
                                                else
                                                {
                                                    if (tmp_del != null)
                                                    {
                                                        string stationNO_Merge = "";
                                                        int delMath_UseTime = 0; int tmp_ct = 0; int tmp_wt = 0; int tmp_st = 0; int tmp_1 = 0; int tmp_2 = 0; int tmp_3 = 0; int tmp_4 = 0;
                                                        tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SI_SimulationId}'");
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
                                                        dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SI_SimulationId}' and StationNO='{data[0]}' order by CalendarDate");
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
                                                            dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SI_SimulationId}' and StationNO {stationNO_Merge} order by CalendarDate");
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

                                                        tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SI_SimulationId}'");
                                                        if (tmp_del != null && int.Parse(tmp_del["Time1_C"].ToString()) == 0 && int.Parse(tmp_del["Time2_C"].ToString()) == 0 && int.Parse(tmp_del["Time3_C"].ToString()) == 0 && int.Parse(tmp_del["Time4_C"].ToString()) == 0 && int.Parse(tmp_del["NeedQTY"].ToString()) > (int.Parse(tmp_del["Detail_QTY"].ToString()) + int.Parse(tmp_del["Detail_Fail_QTY"].ToString())))
                                                        {
                                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET Time1_C=1 where SimulationId='{keys.SI_SimulationId}'");
                                                        }
                                                    }
                                                }
                                                #endregion
                                            }

                                            keys.SI_FailQTY = 0;
                                            keys.SI_OKQTY = 0;

                                            keys.SI_PP_Name = "";
                                            keys.SI_PartName = "";
                                            keys.SI_IndexSN = "";
                                            keys.SI_OrderNO = "";
                                            keys.SI_PartNO = "";
                                            keys.MES_String = "完成報工作業.";
                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2Work','網頁報工','{dr_MII["PP_Name"].ToString()}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{keys.SI_Slect_OPNOs}',{dr_MII["IndexSN"].ToString()})");
                                            keys.SI_OPNO = "";
                                            keys.SI_Slect_OPNOs = "";
                                        }

                                    }
                                }
                                */

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
                                    if (keys.Select_ID == null || keys.Select_ID.Split(';').Length>2) { keys.ERRMsg = $"查看 ESOP 需先選擇一個項目"; goto break_FUN; }
                                    string id = "";
                                    foreach (string s in keys.Select_ID.Split(';'))
                                    {
                                        if (s.Trim() != "")
                                        {
                                            if (id == "") { id = $"'{s}'"; }
                                            else { id = $"{id},'{s}'"; }
                                        }
                                    }
                                    DataTable tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id in ({id})");
                                    if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                    {
                                        foreach(DataRow dr in tmp_dt.Rows)
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
                                    ViewBag.Report = meg;
                                }
                                return View("DisplayKnives");
                            case "領料":
                                {
                                    if (keys.Select_ID == null || keys.Select_ID == "") { keys.ERRMsg = $"領料前, 需先選擇項目."; goto break_FUN; }
                                    string err = "";
                                    DataRow dr_MII = null;
                                    ViewBag.EStore = "";
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
                                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' where IsOK='0' and Id='{dr["Id"].ToString()}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and ArrivalDate='{Convert.ToDateTime(dr["ArrivalDate"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'");

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
                                                    DataRow d2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_MII["StationNO"].ToString()}'");
                                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelWork','智慧領料','{d2["PP_Name"].ToString()}','{dr_MII["StationNO"].ToString()}','{d2["PartNO"].ToString()}','{d2["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");
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
                                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' where IsOK='0' and Id='{dr["Id"].ToString()}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and ArrivalDate='{Convert.ToDateTime(dr["ArrivalDate"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'");

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
                                    }
                                }
                                return View("DisplayEStore");
                            case "入料":
                                {
                                    if (keys.Select_ID == null || keys.Select_ID == "") { keys.ERRMsg = $"入料前, 需先選擇項目."; goto break_FUN; }
                                    DataRow dr_MII = null;
                                    ViewBag.EStore = "";
                                    foreach (string id in keys.Select_ID.Split(';'))
                                    {
                                        if (id == "") { continue; }
                                        dr_MII = db.DB_GetFirstDataByDataRow($"SELECT *  FROM SoftNetMainDB.[dbo].[ManufactureII] where Id='{id}'");
                                        MutiStationSetInStore(dr_MII, br.UserNO);
                                        DataRow d2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_MII["StationNO"].ToString()}'");
                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelWork','智慧入料','{d2["PP_Name"].ToString()}','{dr_MII["StationNO"].ToString()}','{d2["PartNO"].ToString()}','{d2["OrderNO"].ToString()}','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");
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
                //                        where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO = '{keys.StationNO}' and b.EndTime is NULL and b.EstimatedStartTime<='{DateTime.Now.AddMonths(2).ToString("MM/dd/yyyy HH:mm:ss.fff")}' order by b.EstimatedStartTime";

                DateTime comp_CalendarDate = DateTime.Now.AddYears(10);//###??? 提早顯示的時間要參數化
                sql = $@"select a.DOCNumberNO,a.NeedId from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] as a 
                                join SoftNetSYSDB.[dbo].[PP_WorkOrder] as b on a.DOCNumberNO=b.OrderNO and b.EndTime is NULL 
                                join SoftNetSYSDB.[dbo].[APS_NeedData] as c on c.Id=a.NeedId and c.State='6' and c.ServerId='{_Fun.Config.ServerId}'
                                where (a.StationNO = '{keys.StationNO}' or a.StationNO_Merge like '%{keys.StationNO},%') and a.CalendarDate<='{comp_CalendarDate.ToString("MM/dd/yyyy HH:mm:ss.fff")}' and a.DOCNumberNO!='' 
                                and (a.Time1_C!=0 or a.Time2_C!=0 or a.Time3_C!=0 or a.Time4_C!=0) group by a.DOCNumberNO,a.NeedId";
                DataTable dt_PP = db.DB_GetData(sql);
                if (dt_PP != null && dt_PP.Rows.Count > 0)
                {
                    string inkey = "";
                    foreach (DataRow dr0 in dt_PP.Rows)
                    {
                        if (inkey == "") { inkey = $"'{dr0["DOCNumberNO"].ToString()}'"; } else { inkey = $"{inkey},'{dr0["DOCNumberNO"].ToString()}'"; }
                    }
                    //###??? 之後還要補不是NeedID,直接輸入工單的資料 加到inkey
                    Dictionary<string, List<string>> orderOrder = new Dictionary<string, List<string>>();
                    dt_PP = db.DB_GetData($"select DOCNumberNO,SimulationId from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where DOCNumberNO in ({inkey}) and (StationNO = '{keys.StationNO}' or StationNO_Merge like '%{keys.StationNO},%') and CalendarDate<='{comp_CalendarDate.ToString("MM/dd/yyyy HH:mm:ss.fff")}' order by CalendarDate,DOCNumberNO");
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
                    DataRow dr2 = null;
                    DataRow dr3 = null;
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
                                    if (dr2["NeedType"].ToString() == "1") { sType2 = $"訂單需求"; }
                                    else if (dr2["NeedType"].ToString() == "2") { sType2 = $"客戶需求"; }
                                    else if (dr2["NeedType"].ToString() == "5") { sType2 = $"底稿需求"; }
                                    else { sType2 = $"廠內需求"; }
                                    dr3 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr2["PartNO"].ToString()}'");
                                    if (db.DB_GetQueryCount($"SELECT Id FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO = '{keys.StationNO}' and OrderNO='{dr2["DOCNumberNO"].ToString()}' and IndexSN='{dr2["Source_StationNO_IndexSN"].ToString()}'") > 0) { }
                                    else
                                    { keys.HasWO_List.Add(new string[] { dr2["DOCNumberNO"].ToString(), dr2["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString(), sType2, dr2["CTName"].ToString(), dr2["Apply_PP_Name"].ToString(), dr2["Source_StationNO_IndexSN"].ToString(), dr2["SimulationId"].ToString(), dr2["Source_StationNO_Custom_IndexSN"].ToString(), dr2["Source_StationNO_Custom_DisplayName"].ToString(), dr2["QTY"].ToString() }); }
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
                            eCT= tmp_dr2["ECT"].ToString();
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
                        if ((okQTY + t_okQTY + failedQty + t_failedQty)>= int.Parse(dr["PNQTY"].ToString())) { woClose_ck = ""; }
                        else { woClose_ck = $"工單:{dr["OrderNO"].ToString()} 工單量:{dr["PNQTY"].ToString()} 不足量:{(int.Parse(dr["PNQTY"].ToString()) - okQTY + t_okQTY + failedQty + t_failedQty).ToString()} 數量."; }
                        keys.HasWorkData_List.Add(new string[] { dr["Id"].ToString(), state, partNO, partName, specification, dr["PP_Name"].ToString(), tmp_Index, dr["OrderNO"].ToString(), $"{okQTY.ToString()}+{t_okQTY.ToString()}", $" {failedQty.ToString()}+{t_failedQty.ToString()}", totCT.ToString(), eCT, remarkTimeS, remarkTimeE, dr["PNQTY"].ToString(), source_StationNO_Custom_IndexSN, source_StationNO_Custom_DisplayName , next_StationQTY,woClose_ck });
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

        private List<string> MutiStationSetOutStore(DataRow dr_MII,string userNO, ref string err)//領料
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
                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
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
                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                            _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
                                                            tmp_int = 0;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                            if (tmp_01 != 0)
                                                            {
                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
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
                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
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
                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "StationSetOutStore;LabelWorkController", ref docNumberNO, userNO, true);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                        if (tmp_01 != 0)
                                                        {
                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
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
                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
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
                            dt_APS_Simulation = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{dr_WO["NeedId"].ToString()}' and (Store_DOCNumberNO is not null or Store_DOCNumberNO !='') and {before_ST}");
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
                                                        if (tmp_01 != 0)
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
                    string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();
                    DataTable dt_APS_PartNOTimeNote = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and SimulationId='{dr_MII["SimulationId"].ToString()}'");
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
                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and Keys='入料_{userNO}' and StationNO='{dr_MII["StationNO"].ToString()}' and StoreNO='{s}'");
                        db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[BarCode_TMP] (ServerId,[Keys],[StationNO],[StoreNO],[Value],FailTime) VALUES ('{_Fun.Config.ServerId}','入料_{userNO}','{dr_MII["StationNO"].ToString()}','{s}','{dr_MII["OrderNO"].ToString()},{dr_MII["IndexSN"].ToString()},{dr_MII["StationNO"].ToString()}','{DateTime.Now.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss")}')");
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
