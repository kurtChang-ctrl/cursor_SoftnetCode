using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using System;
using BaseApi.Services;
using SoftNetWebII.Models;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using Base;




namespace SoftNetWebII.Controllers
{
    public class STView2Controller : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public STView2Controller(SNWebSocketService websocket, SFC_Common sfc_Common)
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
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            keys.MES_Report = "";
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
                            if (!bool.Parse(tmp["Config_MutiWO"].ToString()))
                            { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Config_MutiWO='1',Label_ProjectType='0' where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'"); }
                        }
                        StationNOList.Add(new string[] { dr["StationNO"].ToString(), dr["StationName"].ToString() });
                    }
                }

                string report = "";
                if (keys.StationNO != null && keys.StationNO.Trim() != "")
                {
                    DataTable dt_tmp = db.DB_GetData($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and  LOGDateTime>='{DateTime.Now.AddDays(-14).ToString("MM/dd/yyyy HH:mm:ss.fff")}'  order by LOGDateTime,OP_NO desc");
                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                    {
                        string opName = "";
                        string partName = "";
                        report = "<div>前14日歷史報工紀錄</div><div><table data-role='table' data-mode='columntoggle' class='ui-responsive' id='myTable'><thead><tr><th>報工日期</th><th data-priority='1'>製程名稱</th><th>報工人員</th><th>料號</th><th>品名  規格</th><th>OK量</th><th>Fail量</th><th data-priority='2'>平均CT</th></tr></thead><tbody>";
                        foreach (DataRow d in dt_tmp.Rows)
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[User] where ServerId='{_Fun.Config.ServerId}' and UserNO='{d["OP_NO"].ToString()}'");
                            if (tmp != null) { opName = $"{tmp["UserNO"].ToString()} {tmp["Name"].ToString()}"; } else { opName = ""; }
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}'");
                            if (tmp != null) { partName = $"{tmp["PartName"].ToString()} {tmp["Specification"].ToString()}"; } else { partName = ""; }
                            report = $"{report}<tr><td>{Convert.ToDateTime(d["LOGDateTime"]).ToString("yyyy/MM/dd HH:mm:ss")}</td><td>{d["PP_Name"].ToString()}</td><td>{opName}</td><td>{d["PartNO"].ToString()}</td><td>{partName}</td><td>{d["OKQTY"].ToString()}</td><td>{d["FailQTY"].ToString()}</td><td>{d["CycleTime"].ToString()}</td></tr>";
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
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2Controller.cs MUTIStation {state} Exception: {ex.Message} {ex.StackTrace}", true);
                return View("ResuItTimeOUT");
            }
            return View(keys);
        }

        [HttpPost]
        public ActionResult SetStation_Open(MUTIStationObj keys)
        {
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                try
                {
                    var br = _Fun.GetBaseUser();
                    if (br == null || !br.IsLogin || br.UserNO.Trim() == "")
                    {
                        return RedirectToAction("Login", "Home");
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
                            foreach (string id in data)
                            {
                                if (id.Trim() != "")
                                {
                                    dr_MII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}' and EndTime is NULL");
                                    if (dr_MII != null)
                                    {
                                        if (dr_MII.IsNull("RemarkTimeS") || (!dr_MII.IsNull("RemarkTimeS") && !dr_MII.IsNull("RemarkTimeE")))
                                        {
                                            if (dr_MII.IsNull("StartTime")) { tmp = $",StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                                            db.DB_SetData($"update [dbo].[ManufactureII] set RemarkTimeE=NULL,RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'{tmp} where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}'");
                                            keys.MES_String = "完成工站啟動生產.";
                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2','學習模式開工','{dr_MII["PP_Name"].ToString()} ','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','','{br.UserNO}',{dr_MII["PartNO"].ToString()})");//###???dr_M["OrderNO"].ToString()暫時放空
                                            DataTable tmp_dt = db.DB_GetData($"select * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and OrderNO='{dr_MII["OrderNO"].ToString()}' and PartNO='{dr_MII["PartNO"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()} and EndTime is NULL");
                                            if (tmp_dt != null && tmp_dt.Rows.Count>0)
                                            {
                                                dr_MII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}' and EndTime is NULL");
                                                string tmp_s = "";
                                                foreach (DataRow d2 in tmp_dt.Rows)
                                                {
                                                    tmp_s = "";
                                                    if (d2.IsNull("StartTime")) { tmp_s = $"{tmp_s},StartTime='{Convert.ToDateTime(dr_MII["StartTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                    if (d2.IsNull("RemarkTimeS")) { tmp_s = $"{tmp_s},RemarkTimeS='{Convert.ToDateTime(dr_MII["RemarkTimeS"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                    db.DB_SetData($"update SoftNetLogDB.[dbo].[SFC_StationProjectDetail] set RemarkTimeE=NULL{tmp_s}  where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and OrderNO='{dr_MII["OrderNO"].ToString()}' and PartNO='{dr_MII["PartNO"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()} and EndTime is NULL");
                                                }
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
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2Controller.cs SetStation_Open {state} Exception: {ex.Message} {ex.StackTrace}", true);
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
            if (br == null || !br.IsLogin || br.UserNO.Trim() == "")
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
                                                db.DB_SetData($"update [dbo].[ManufactureII] set RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}'");
                                                keys.MES_String = "已停止工站生產作業.";
                                                db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2','學習模式停工','{dr_MII["PP_Name"].ToString()} ','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','','{br.UserNO}',{dr_MII["IndexSN"].ToString()})");//###???{dr_MII["OrderNO"].ToString()} 暫時放空
                                                DataTable tmp_dt = db.DB_GetData($"select * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and OrderNO='{dr_MII["OrderNO"].ToString()}' and PartNO='{dr_MII["PartNO"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()} and EndTime is NULL");
                                                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                                {
                                                    dr_MII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}' and EndTime is NULL");
                                                    string tmp_s = "";
                                                    foreach (DataRow d2 in tmp_dt.Rows)
                                                    {
                                                        tmp_s = "";
                                                        if (d2.IsNull("StartTime")) { tmp_s = $"{tmp_s},StartTime='{Convert.ToDateTime(dr_MII["StartTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                        if (d2.IsNull("RemarkTimeS")) { tmp_s = $"{tmp_s},RemarkTimeS='{Convert.ToDateTime(dr_MII["RemarkTimeS"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                        db.DB_SetData($"update SoftNetLogDB.[dbo].[SFC_StationProjectDetail] set RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'{tmp_s}  where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and OrderNO='{dr_MII["OrderNO"].ToString()}' and PartNO='{dr_MII["PartNO"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()} and EndTime is NULL");
                                                    }
                                                }
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
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2Controller.cs SetStation_Stop {state} Exception: {ex.Message} {ex.StackTrace}", true);
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
                            string tmp = "";
                            string opNO = "";
                            string beforePartNO = "";
                            string beforeIndexSN = "";
                            foreach (string id in data)
                            {
                                if (id.Trim() != "")
                                {

                                    dr_MII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}'");
                                    if (dr_MII != null)
                                    {
                                        beforePartNO = dr_MII["PartNO"].ToString();
                                        beforeIndexSN = dr_MII["IndexSN"].ToString();
                                        if (!dr_MII.IsNull("RemarkTimeS") && dr_MII.IsNull("RemarkTimeE"))
                                        { keys.ERRMsg = $"{dr_MII["PartNO"].ToString()} 狀態為 生產中, 無法關閉項目, 該項目請先停工."; }
                                        else
                                        {
                                            #region 檢查之前是否無報工
                                            int avgReportTime = 0;
                                            DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()}  and PartNO='{beforePartNO}' and ReportTime>5");
                                            if (dr_tmp != null && !dr_tmp.IsNull("AVGTime") && dr_tmp["AVGTime"].ToString().Trim() != "")
                                            {
                                                avgReportTime = int.Parse(dr_tmp["AVGTime"].ToString());
                                            }
                                            if (avgReportTime > 0)
                                            {
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={beforeIndexSN} and PartNO='{beforePartNO}' and LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%開工%' order by LOGDateTime desc");
                                                if (dr_tmp != null)
                                                {
                                                    DateTime tmp_edate = Convert.ToDateTime(dr_tmp["LOGDateTime"]);
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={beforeIndexSN} and PartNO='{beforePartNO}' and LOGDateTime>'{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%報工%'");
                                                    if (dr_tmp == null)
                                                    {
                                                        if ((_SFC_Common.TimeCompute2Seconds(tmp_edate, DateTime.Now)) >= avgReportTime)
                                                        {
                                                            keys.ERRMsg = $"疑似 前一次料號:{dr_MII["PartNO"].ToString()} 未完成報工, 若未報工須請管理者, 由後臺輔助報工";
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion

                                            tmp = "";
                                            if (dr_MII.IsNull("StartTime")) { tmp = $",StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                            if (dr_MII.IsNull("RemarkTimeS")) { tmp = $"{tmp},RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                            if (dr_MII.IsNull("RemarkTimeE")) { tmp = $"{tmp},RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                            db.DB_SetData($"update [dbo].[ManufactureII] set EndTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' {tmp} where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}'");

                                            dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and OrderNO='{dr_MII["OrderNO"].ToString()}' and PartNO='{dr_MII["PartNO"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()} and EndTime is NULL");
                                            if (dr != null)
                                            { db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SFC_StationProjectDetail] SET EndTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}',RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and Id='{dr["Id"].ToString()}' and StationNO='{keys.StationNO}'"); }
                                            keys.MES_String = "已關閉本次生產作業, 可進入工站設定, 設定下次生產";
                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2','學習模式關站','{dr_MII["PP_Name"].ToString()} ','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','','{br.UserNO}',{dr_MII["PartNO"].ToString()})");//###???{dr_MII["OrderNO"].ToString()} 暫時放空
                                            
                                            DataTable tmp_dt = db.DB_GetData($"select * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and OrderNO='{dr_MII["OrderNO"].ToString()}' and PartNO='{dr_MII["PartNO"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()} and EndTime is NULL");
                                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                            {
                                                dr_MII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}' and EndTime is NULL");
                                                if (dr_MII != null)
                                                {
                                                    string tmp_s = "";
                                                    foreach (DataRow d2 in tmp_dt.Rows)
                                                    {
                                                        if (d2.IsNull("StartTime")) { tmp_s = $"{tmp_s},StartTime='{Convert.ToDateTime(dr_MII["StartTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                        if (d2.IsNull("RemarkTimeS")) { tmp_s = $"{tmp_s},RemarkTimeS='{Convert.ToDateTime(dr_MII["RemarkTimeS"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                        if (d2.IsNull("RemarkTimeE")) { tmp_s = $"{tmp_s},RemarkTimeE='{Convert.ToDateTime(dr_MII["RemarkTimeE"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                                        db.DB_SetData($"update SoftNetLogDB.[dbo].[SFC_StationProjectDetail] set EndTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'{tmp_s}  where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and OrderNO='{dr_MII["OrderNO"].ToString()}' and PartNO='{dr_MII["PartNO"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()} and EndTime is NULL");
                                                    }
                                                }
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
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2Controller.cs SetStation_Close {state} Exception: {ex.Message} {ex.StackTrace}", true);
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
                keys.ERRMsg = "你沒有選擇工站, 請按 回到工站選單畫面 重新選擇工站.";
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
                        //DataRow dr_M = db.DB_GetFirstDataByDataRow($"SELECT * from [dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                        //if (dr_M == null) { keys.ERRMsg = $"資料庫 Manufacture設定異常, 請通知管理者."; goto break_FUN; }
                        DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");

                        switch (keys.State)
                        {
                            case "新增項目":
                                {
                                    #region 網頁內容檢查
                                    DataRow dr = null;
                                    string stationNO_Custom_IndexSN = "";
                                    string stationNO_Custom_DisplayName = "";
                                    //###???1108要改主動維護 State
                                    //if (dr_M["State"].ToString() == "1") { keys.ERRMsg = $"本工站 {keys.StationNO} 已運作中, 請先停止才能設定."; goto break_FUN; }
                                    if (keys.SI_PP_Name == null) { keys.SI_PP_Name = ""; }
                                    if (keys.SI_OrderNO == null) { keys.SI_OrderNO = ""; }
                                    if (keys.SI_PartNO == null) { keys.SI_PartNO = ""; }
                                    if (keys.SI_PP_Name == "")
                                    { keys.ERRMsg = $"製程名稱必須要有值, 請重新設定."; goto break_FUN; }
                                    else if (keys.SI_OrderNO == "" && keys.SI_PartNO == "")
                                    { keys.ERRMsg = $"工單編號 或 料件編號 其中一項不能空白, 請重新設定."; goto break_FUN; }
                                    if (keys.SI_OrderNO.Trim() != "")
                                    {
                                        DataRow sfcdr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{keys.SI_OrderNO}'");
                                        if (sfcdr == null) { keys.ERRMsg = $"查無 {keys.SI_OrderNO} 工單資料, 請重新輸入."; goto break_FUN; }
                                    }
                                    if (keys.SI_PP_Name.Trim() != "")
                                    {
                                        if (keys.SI_PP_Name.IndexOf(";") >= 0)
                                        {
                                            string[] data = keys.SI_PP_Name.Split(';');
                                            keys.SI_PP_Name = data[0];
                                            keys.SI_IndexSN = data[1];
                                        }
                                        dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{keys.SI_PP_Name}' and StationNO='{keys.StationNO}' and IndexSN={keys.SI_IndexSN}");
                                        if (dr == null)
                                        { keys.ERRMsg = $"查無相關製程順序.製程={keys.SI_PP_Name} 順序={keys.SI_IndexSN}, 請通知管理者."; goto break_FUN; }
                                        else
                                        {
                                            stationNO_Custom_IndexSN = dr["Station_Custom_IndexSN"].ToString();
                                            stationNO_Custom_DisplayName = dr["DisplayName"].ToString();
                                        }
                                    }
                                    if (keys.SI_PartNO.Trim() != "")
                                    {
                                        dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{keys.SI_PartNO}'");
                                        if (dr == null)
                                        { keys.ERRMsg = $"查無相關料件編號:{keys.SI_PartNO}, 請重新設定."; goto break_FUN; }
                                    }
                                    #endregion

                                    #region 更新ManufactureII 製造現場狀態
                                    tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and OrderNO='{keys.SI_OrderNO}' and IndexSN={keys.SI_IndexSN} and PP_Name='{keys.SI_PP_Name}' and PartNO='{keys.SI_PartNO}' and EndTime is NULL");
                                    if (tmp_dr == null)
                                    {
                                        string id = _Str.NewId('C');
                                        db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[ManufactureII] ([Id],[StationNO],[ServerId],[OrderNO],[Master_PP_Name],[PP_Name],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],[PartNO],[SimulationId])
                                              VALUES ('{id}','{keys.StationNO}','{_Fun.Config.ServerId}','{keys.SI_OrderNO}','{keys.SI_PP_Name}','{keys.SI_PP_Name}',{keys.SI_IndexSN},'{stationNO_Custom_IndexSN}','{stationNO_Custom_DisplayName}','{keys.SI_PartNO}','')");
                                        //tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{keys.SI_PartNO}'");
                                        keys.MES_String = "完成 工站設定作業....";
                                    }
                                    else
                                    {
                                        keys.ERRMsg = $"目前已有相同生產項目還未關閉,無法重複相同項目, 請重新設定."; goto break_FUN;
                                    }
                                    #endregion

                                    keys.SI_FailQTY = 0;
                                    keys.SI_OKQTY = 0;
                                    keys.SI_OPNO = "";
                                    keys.SI_PP_Name = "";
                                    keys.SI_PartName = "";
                                    keys.SI_IndexSN = "";
                                    keys.SI_OrderNO = "";
                                    keys.SI_PartNO = "";
                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2','新增項目','{keys.SI_PP_Name} ','{keys.StationNO}','{keys.SI_PartNO}','{keys.SI_OrderNO}','{br.UserNO}')");

                                }
                                break;
                            case "報工":
                                {
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
                                    DataRow dr_MII = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}' and EndTime is NULL");
                                    if (dr_MII == null) { keys.ERRMsg = $"該項目已關閉生產, 無法報工."; goto break_FUN; }
                                    if (dr_MII.IsNull("StartTime")) { keys.ERRMsg = $"項目未設定開始過, 無法報工"; goto break_FUN; }
                                    DateTime startTime = DateTime.Now;
                                    if (keys.SI_OKQTY > 0 || keys.SI_FailQTY > 0)
                                    {
                                        DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and OP_NO='{keys.SI_Slect_OPNOs}' and OrderNO='{dr_MII["OrderNO"].ToString()}' and PartNO='{dr_MII["PartNO"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()} and EndTime is NULL");
                                        if (dr == null)
                                        {
                                            db.DB_SetData(@$"INSERT INTO SoftNetLogDB.[dbo].[SFC_StationProjectDetail] (ServerId,Id,StationNO,Master_PP_Name,PP_Name,OP_NO,IndexSN,OrderNO,PartNO,RMSName,StartTime) VALUES 
                                                ('{_Fun.Config.ServerId}','{_Str.NewId('A')}','{keys.StationNO}','{dr_MII["Master_PP_Name"].ToString()}','{dr_MII["PP_Name"].ToString()}','{keys.SI_Slect_OPNOs}',{dr_MII["IndexSN"].ToString()},'{dr_MII["OrderNO"].ToString()}','{dr_MII["PartNO"].ToString()}','{dr_PP_Station["RMSName"].ToString()}','{Convert.ToDateTime(dr_MII["StartTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}')");
                                            dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and OP_NO='{keys.SI_Slect_OPNOs}' and OrderNO='{dr_MII["OrderNO"].ToString()}' and PartNO='{dr_MII["PartNO"].ToString()}' and IndexSN={dr_MII["IndexSN"].ToString()} and EndTime is NULL");
                                        }
                                        if (dr == null)
                                        {
                                            keys.ERRMsg = $"系統異常, 查無標籤設定資料, 請通知管理者."; goto break_FUN;
                                        }
                                        else
                                        {
                                            #region 計算報工 CT,平均,有效, 寫SFC_StationProjectDetail
                                            DateTime rRemarkTimeS = new DateTime();
                                            if (dr_MII.IsNull("RemarkTimeS"))
                                            { rRemarkTimeS = Convert.ToDateTime(dr_MII["StartTime"]); }
                                            else
                                            { rRemarkTimeS = Convert.ToDateTime(dr_MII["RemarkTimeS"]); }
                                            List<double> allCT = new List<double>();
                                            string docNO = "";
                                            string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
                                            decimal ct = 0;
                                            if (dr_MII.IsNull("RemarkTimeE") || rRemarkTimeS >= Convert.ToDateTime(dr_MII["RemarkTimeE"]))
                                            { ct = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), rRemarkTimeS, DateTime.Now) / (keys.SI_OKQTY + keys.SI_FailQTY); }
                                            else
                                            { ct = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), rRemarkTimeS, Convert.ToDateTime(dr_MII["RemarkTimeE"])) / (keys.SI_OKQTY + keys.SI_FailQTY); }
                                            decimal ct_log = ct < 1 ? 0 : ct;
                                            if (int.Parse(dr["CycleTime"].ToString()) != 0) { ct = (ct + int.Parse(dr["CycleTime"].ToString())) > 0 ? Math.Round((ct + int.Parse(dr["CycleTime"].ToString())) / 2) : ct; }
                                            if (ct < 1) { ct = 0; }
                                            string remark = "";
                                            DataRow dr_SFC_StationProjectDetail = null;
                                                dr_SFC_StationProjectDetail = db.DB_GetFirstDataByDataRow($@"SELECT (sum(SD_LowerLimit*CountQTY)/sum(CountQTY)) as CCT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] 
                                                where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PP_Name='{dr["PP_Name"].ToString()}' and PartNO='{dr["PartNO"].ToString()}' and DOCNO='' 
                                                group by ServerId,StationNO,PP_Name,PartNO");
                                            if (dr_SFC_StationProjectDetail != null && !dr_SFC_StationProjectDetail.IsNull("CCT") && dr_SFC_StationProjectDetail["CCT"].ToString().Trim() == "" && int.Parse(dr_SFC_StationProjectDetail["CCT"].ToString()) > 2)
                                            {
                                                float tmp = float.Parse(dr_SFC_StationProjectDetail["CCT"].ToString()) * 0.85f;
                                                if (ct_log >= (decimal)tmp) { remark = $",RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                                            }
                                            
                                            if (keys.Remark != null && keys.Remark.Trim() != "") { remark = $",Remark='{keys.Remark.Trim()}'"; }
                                            if (db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SFC_StationProjectDetail] SET ProductFinishedQty+={keys.SI_OKQTY.ToString()},ProductFailedQty+={keys.SI_FailQTY.ToString()},CycleTime={ct}{remark} where Id='{dr["Id"].ToString()}' and StationNO='{keys.StationNO}' and OP_NO='{keys.SI_Slect_OPNOs}'"))
                                            {
                                            }
                                            //###???未來專案版改智慧版 要改DB內要改PartNO,Sub_PartNO,Master_PP_Name,PP_Name
                                            string tmp_s = "";
                                            if (dr_MII["PartNO"].ToString().Trim() != "") { tmp_s = $" and PartNO='{dr_MII["PartNO"].ToString().Trim()}'"; }
                                            DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} * from SoftNetLogDB.[dbo].SFC_StationProjectDetail where ServerId='{_Fun.Config.ServerId}' and PP_Name='{dr_MII["PP_Name"].ToString()}' and StationNO='{dr_MII["StationNO"].ToString()}'{tmp_s} and ProductFinishedQty!=0 and CycleTime!=0 order by StationNO,OP_NO,PP_Name,PartNO");
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
                                                        _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, dr2["Master_PP_Name"].ToString(), dr2["PP_Name"].ToString(),"0", partNO, partNO, docNO);
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
                                            //先確認目前行事曆時段日期
                                            DateTime tmp_date = DateTime.Now;
                                            DateTime tmp_startdate = DateTime.Now;
                                            int reportTime = 0;
                                            DataRow d2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_PP_Station["CalendarName"].ToString()}' and Holiday='{tmp_date.AddDays(-1).ToString("MM/dd/yyyy")}'");
                                            if (d2 == null)
                                            { d2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_PP_Station["CalendarName"].ToString()}' and Holiday='{tmp_date.ToString("MM/dd/yyyy")}'"); }
                                            if (d2 != null)
                                            {
                                                bool isRUN = true;
                                                if (int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                                {
                                                    if (tmp_date > new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0) && tmp_date < new DateTime(tmp_date.AddDays(1).Year, tmp_date.AddDays(1).Month, tmp_date.AddDays(1).Day, int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 59))
                                                    { isRUN = false; tmp_date = tmp_date.AddDays(-1); }
                                                }
                                                if (isRUN && int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                                {
                                                    if (tmp_date > new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), 0) && tmp_date < new DateTime(tmp_date.AddDays(1).Year, tmp_date.AddDays(1).Month, tmp_date.AddDays(1).Day, int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[1]), 59))
                                                    { isRUN = false; tmp_date = tmp_date.AddDays(-1); }
                                                }
                                                if (isRUN)
                                                { d2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_PP_Station["CalendarName"].ToString()}' and Holiday='{tmp_date.ToString("MM/dd/yyyy")}'"); }
                                                if (d2 != null)
                                                {
                                                    tmp_startdate = Convert.ToDateTime(d2["Holiday"]);
                                                    tmp_startdate = new DateTime(tmp_startdate.Year, tmp_startdate.Month, tmp_startdate.Day, int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                                }
                                                //計算上一次與現在時間差
                                                d2 = db.DB_GetFirstDataByDataRow($"SELECT LOGDateTime FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and CONVERT(varchar(100),LOGDateTime, 23)='{tmp_startdate.ToString("yyyy-MM-dd")}' order by LOGDateTime,LOGDateTimeID desc");
                                                if (d2 != null)
                                                { reportTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(d2["LOGDateTimeID"]), DateTime.Now); }
                                                if (reportTime < 0) { reportTime = 0; }
                                            }
                                            db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] (ServerId,[LOGDateTime],[Id],[StationNO],[OP_NO],[OKQTY],[FailQTY],[CycleTime],[WaitTime],PP_Name,IndexSN,PartNO,OrderNO,EfficientCycleTime,Custom_SD_LowerLimit,ReportTime) VALUES (
                                                                '{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{dr["Id"].ToString()}','{keys.StationNO}','{keys.SI_Slect_OPNOs}',{keys.SI_OKQTY.ToString()},{keys.SI_FailQTY.ToString()},{ct_log},0,'{dr["PP_Name"].ToString()}',{dr["IndexSN"].ToString()},'{dr["PartNO"].ToString()}','{dr["OrderNO"].ToString()}',{efficientCycleTime},{custom_SD_LowerLimit},{reportTime})");
                                            #endregion

                                            keys.SI_FailQTY = 0;
                                            keys.SI_OKQTY = 0;

                                            keys.SI_PP_Name = "";
                                            keys.SI_PartName = "";
                                            keys.SI_IndexSN = "";
                                            keys.SI_OrderNO = "";
                                            keys.SI_PartNO = "";
                                            keys.MES_String = "完成報工作業.";
                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2','多工站報工','{dr_MII["PP_Name"].ToString()} {keys.SI_Slect_OPNOs}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{br.UserNO}')");
                                            keys.SI_OPNO = "";
                                            keys.SI_Slect_OPNOs = "";
                                        }
                                    }
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
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Report = "";
                        ViewBag.ErrType = "SystemError";
                        keys.OutError = true;
                        keys.ERRMsg = $"程式異常, 導致系統無法使用, 請通知管理者.";
                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2Controller.cs SetAction Exception: {ex.Message} {ex.StackTrace}", true);
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
                keys.HasWorkData_List = new List<string[]>();//Id,狀態,PartNO,PartName,Specification,PP_Name,IndexSN,OrderNO,okQTY,failedQty,totCT,目標cT,開始時間,停止時間

                DataRow keys_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                if (keys_dr == null)
                {
                    keys.ERRMsg = $"查無您選擇的 {keys.StationNO} 工站, 請重新操作.";
                    return false;
                }
                keys.StationName = keys_dr["StationName"].ToString();
                //###???子製程有問題
                #region 回傳學習模式可用製程名稱 與 料號
                sql = @$"select * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and StationNO = '{keys.StationNO}' Order by PP_Name,IndexSN,Station_Custom_IndexSN";
                DataTable dt_PP = db.DB_GetData(sql);
                if (dt_PP != null && dt_PP.Rows.Count > 0)
                {
                    keys.HasPP_Name_List = new List<string[]>();
                    DataRow tmp_dr2 = null;
                    DataTable tmp_dt3 = null;
                    List<string> tmp_HasPartNO = new List<string>();
                    keys.HasPartNO_List = new List<string[]>();
                    foreach (DataRow dr in dt_PP.Rows)
                    {
                        keys.HasPP_Name_List.Add(new string[] { dr["PP_Name"].ToString(), dr["IndexSN"].ToString(), dr["Station_Custom_IndexSN"].ToString() });
                        tmp_dr2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item_PartNO] where Id='{dr["Id"].ToString()}'");
                        if (tmp_dr2 != null)
                        {

                            if (tmp_dr2["By_Class"].ToString().Trim() != "")
                            {
                                tmp_dt3 = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and Class='{tmp_dr2["By_Class"].ToString()}'");
                                if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                {
                                    foreach (DataRow dr3 in tmp_dt3.Rows)
                                    {
                                        if (!tmp_HasPartNO.Contains(dr3["PartNO"].ToString()))
                                        {
                                            tmp_HasPartNO.Add(dr3["PartNO"].ToString());
                                            keys.HasPartNO_List.Add(new string[] { dr3["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString() });
                                        }
                                    }
                                }
                            }
                            if (tmp_dr2["By_PartType"].ToString().Trim() != "")
                            {
                                tmp_dt3 = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartType='{tmp_dr2["By_PartType"].ToString()}'");
                                if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                {
                                    foreach (DataRow dr3 in tmp_dt3.Rows)
                                    {
                                        if (!tmp_HasPartNO.Contains(dr3["PartNO"].ToString()))
                                        {
                                            tmp_HasPartNO.Add(dr3["PartNO"].ToString());
                                            keys.HasPartNO_List.Add(new string[] { dr3["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString() });
                                        }
                                    }
                                }
                            }
                            if (tmp_dr2["By_Model"].ToString().Trim() != "")
                            {
                                tmp_dt3 = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartType='{tmp_dr2["By_Model"].ToString()}'");
                                if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                {
                                    foreach (DataRow dr3 in tmp_dt3.Rows)
                                    {
                                        if (!tmp_HasPartNO.Contains(dr3["PartNO"].ToString()))
                                        {
                                            tmp_HasPartNO.Add(dr3["PartNO"].ToString());
                                            keys.HasPartNO_List.Add(new string[] { dr3["PartNO"].ToString(), dr3["PartName"].ToString(), dr3["Specification"].ToString() });
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
                                        tmp_dr2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{s}'");
                                        if (!tmp_HasPartNO.Contains(s))
                                        {
                                            tmp_HasPartNO.Add(s);
                                            keys.HasPartNO_List.Add(new string[] { tmp_dr2["PartNO"].ToString(), tmp_dr2["PartName"].ToString(), tmp_dr2["Specification"].ToString() });
                                        }
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
                DataTable dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and EndTime is NULL");
                if (dt != null && dt.Rows.Count > 0)
                {
                    DataRow tmp_dr2 = null;
                    int okQTY = 0;
                    int failedQty = 0;
                    int totCT = 0;
                    string state = "";
                    string partNO = "";
                    string partName = "";
                    string specification = "";
                    string remarkTimeS = "";
                    string remarkTimeE = "";
                    DataTable tmp_dt = null;
                    foreach (DataRow dr in dt.Rows)
                    {
                        okQTY = 0; failedQty = 0; totCT = 0; partNO = ""; partName = ""; specification = ""; state = "";

                        tmp_dt = db.DB_GetData($@"SELECT sum(ProductFinishedQty) as OKQTY,sum(ProductFailedQty) as FailedQty,sum(CycleTime)*(sum(ProductFinishedQty)+sum(ProductFailedQty)) as TOTCT FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where 
                                    ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and PP_Name='{dr["PP_Name"].ToString()}' and OrderNO='{dr["OrderNO"].ToString()}' and PartNO='{dr["PartNO"].ToString()}' and IndexSN={dr["IndexSN"].ToString()}
                                    and EndTime is null and StartTime is not null group by OP_NO");
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

                        tmp_dr2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                        if (tmp_dr2 != null)
                        {
                            partNO = tmp_dr2["PartNO"].ToString();
                            partName = tmp_dr2["PartName"].ToString();
                            specification = tmp_dr2["Specification"].ToString();
                        }
                        remarkTimeS = dr.IsNull("RemarkTimeS") ? "" : Convert.ToDateTime(dr["RemarkTimeS"]).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        remarkTimeE = dr.IsNull("RemarkTimeE") ? "" : Convert.ToDateTime(dr["RemarkTimeE"]).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        if (remarkTimeS != "" && remarkTimeE == "") { state = "生產"; }
                        else if (remarkTimeE != "" && !dr.IsNull("StartTime")) { state = "停止"; }
                        keys.HasWorkData_List.Add(new string[] { dr["Id"].ToString(), state, partNO, partName, specification, dr["PP_Name"].ToString(), dr["IndexSN"].ToString(), dr["OrderNO"].ToString(), okQTY.ToString(), failedQty.ToString(), totCT.ToString(), "0", remarkTimeS, remarkTimeE });
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                keys.ERRMsg = $"程式異常, 導致系統無法使用, 請通知管理者.";
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2Controller.cs GetINFO Exception: {ex.Message} {ex.StackTrace}", true);
                return false;
            }
            return true;
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(Ctrl, dt));
        }


        private STView2Service EditService()
        {
            return new STView2Service(Ctrl);
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



    }//class
}
