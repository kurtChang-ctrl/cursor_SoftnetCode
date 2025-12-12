using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseWeb.Models;
using BaseWeb.Services;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class SimulationController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;

        public SimulationController(SNWebSocketService socket, SFC_Common sfc_Common)
        {
            if (_WebSocket == null)
            {
                _WebSocket = socket;
            }
            if (_SFC_Common == null)
            {
                _SFC_Common = sfc_Common;
            }
        }
        public ActionResult Read()
        {
            List<IdStrDto> factoryName = new List<IdStrDto>();
            List<IdStrDto> calendarName = new List<IdStrDto>();
            List<IdStrDto> productProcess = new List<IdStrDto>();
            List<IdStrDto> partNO = new List<IdStrDto>();
            List<string[]> partNOANDpartName = new List<string[]>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT FactoryName FROM [dbo].[Factory] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        factoryName.Add(new IdStrDto(dr["FactoryName"].ToString(), dr["FactoryName"].ToString()));
                    }
                }
                dt = db.DB_GetData($"SELECT ServerId,CalendarName FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' group by ServerId,CalendarName");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        calendarName.Add(new IdStrDto(dr["CalendarName"].ToString(), dr["CalendarName"].ToString()));
                    }
                }
                dt = db.DB_GetData($"SELECT LineName,PP_Name FROM SoftNetSYSDB.[dbo].[PP_ProductProcess] where ServerId='{_Fun.Config.ServerId}' group by LineName,PP_Name");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        productProcess.Add(new IdStrDto(dr["PP_Name"].ToString(), dr["PP_Name"].ToString()));
                    }
                }
                dt = db.DB_GetData($"SELECT PartNO,PartName,Specification FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and (Class='4' or Class='5')");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        partNO.Add(new IdStrDto(dr["PartNO"].ToString(), dr["PartNO"].ToString()));
                        partNOANDpartName.Add(new string[] { dr["PartNO"].ToString(), dr["PartName"].ToString(), dr["Specification"].ToString() });

                    }
                }

            }
            ViewBag.FactoryName = factoryName;
            ViewBag.CalendarName = calendarName;
            ViewBag.ProductProcess = productProcess;
            ViewBag.PartNO = partNO;
            ViewBag.PartNO_PartName = partNOANDpartName;

            return View();
        }
        [HttpPost]
        public string HTML_ChangeClass(string keys) 
        {
            string meg = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT PartNO,PartName,Specification FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and Class='{keys}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (meg == "") { meg = $"{dr["PartNO"].ToString()};{dr["PartName"].ToString()};{dr["Specification"].ToString()}"; }
                        else { meg = $"{meg},{dr["PartNO"].ToString()};{dr["PartName"].ToString()};{dr["Specification"].ToString()}"; }

                    }
                }
            }
            return meg;
        }

        [HttpPost]
        public async Task<JsonResult> ImportExcel(IFormFile file)
        {
            string uname = "";
            if (_Fun.GetBaseUser() != null)
            { uname = _Fun.GetBaseUser().UserName; }
            var importDto = new ExcelImportDto<SimulationDto>()
            {
                ImportType = "Simulation",
                TplPath = _Xp.DirTpl + "UserImport.xlsx",
                FnSaveImportRows = SaveImportRows,
                CreatorName = uname,
            };
            string dirUpload= @"D:\書_sampleCode\HrAdm-master\_upload\UserImport\";
            var model =  await _WebExcel.ImportByFileAsync(file, dirUpload, importDto);
            return Json(model);
        }
        private List<string> SaveImportRows(List<SimulationDto> okRows)
        {
            var results = new List<string>();


            /*
            var db = _Xp.GetDb();
            var deptIds = db.Dept.Select(a => a.DeptId).ToList();
            foreach (var row in okRows)
            {
                //check rules: deptId
                if (!deptIds.Contains(row.DeptId))
                {
                    results.Add("DeptId Wrong.");
                    continue;
                }

                //check rules: Account not repeat
                if (db.User.Any(a => a.Account == row.Account))
                {
                    results.Add("Account Existed.");
                    continue;
                }

                #region set entity model & save db
                db.User.Add(new User()
                {
                    Id = _Str.NewId('Z'),
                    Name = row.Name,
                    Account = row.Account,
                    Pwd = row.Pwd,
                    DeptId = row.DeptId,
                    Status = true,
                });

                //save db
                try
                {
                    db.SaveChanges();
                    results.Add("");
                }
                catch (Exception ex)
                {
                    results.Add(ex.InnerException.Message);
                }
                #endregion
            }

            db.Dispose();
            */
            return results;
        }

        [HttpPost]
        public string GOTORUNPP(string keys) //正式排程 0=ipport,1.需求碼1,需求碼2....
        {
            string meg = "";
            string[] data = keys.Split(',');
            bool run = false;
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                List<string> neddIds = new List<string>();
                DataRow dr_M = null;
                DataTable dt_APS_Simulation = null;
                DataRow tmp;
                DataRow tmp2;
                string ipport = data[0];
                string sql = "";
                for (int i = 1; i < data.Length; i++)
                {
                    dr_M = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{data[i]}' order by NeedSimulationDate");
                    if (dr_M != null)
                    {
                        switch (dr_M["State"].ToString())
                        {
                            case "1":
                                meg = $"{meg}<br>模擬中,無法轉計畫."; break;
                            case "2":
                                string ERR = "";
                                run = _SFC_Common.Create_WorkOrder(db, data[i], dr_M["CalendarName"].ToString(), ref ERR);
                                if (ERR != "") { meg = $"<br>{ERR}"; }
                                if (run)
                                {
                                    if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='6',StateINFO=NULL,UpdateTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and Id='{data[i]}'")) { run = true; }
                                }
                                break;
                            case "3":
                            case "5":
                                meg = $"{meg}<br>需重新模擬."; break;
                            case "4":
                                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='7',StateINFO=NULL,UpdateTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and Id='{data[i]}'")) { run = true; }
                                break;
                            case "8":
                                meg = $"{meg}<br>重新模擬或刪除此需求."; break;
                            default:
                                meg = $"{meg}<br>系統定義有誤, 請通知系統管理者."; break;
                        }
                    }
                    else
                    {
                        meg = $"{meg}<br>無法轉正式計畫";
                    }
                }
                if (run) { _SFC_Common.RunSetSimulation(null, ipport, null, '3'); }
            }
            return meg;
        }

        [HttpPost]
        public string SetSimulation(string ipport, string arg,string ids,string set1, string set2, string set3) //0=ipport,1.排成設定參數,需求碼2....
        {
            string meg = "";
            List<string> data = ids.Split(',').ToList();//需求碼s
            RunSimulation_Arg args = new RunSimulation_Arg();

            #region 紀錄排程設定參數 args
            List<char> cs = arg.PadRight(20, '0').ToArray().ToList();

            if (arg != null)
            {

                foreach (char c in cs)
                {
                    if (c == '0') { args.ARGs.Add(false); }
                    else { args.ARGs.Add(true); }
                }
            }
            else { args.ARGs.AddRange(new bool[20]); }
            #endregion

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow dr_tmp = null;
                #region 前置檢查
                if (data.Count <= 0) { meg = $"網頁為勾選需求"; return meg; }
                for (int i = 0; i < data.Count; i++)
                {
                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT a.* from SoftNetSYSDB.[dbo].[APS_NeedData] as a,SoftNetMainDB.[dbo].[Material] as b where a.ServerId='{_Fun.Config.ServerId}' and b.ServerId='{_Fun.Config.ServerId}' and a.Id='{data[i]}' and a.PartNO=b.PartNO");
                    if (dr_tmp == null) { meg = $"需求中有料號不存在"; return meg; }
                }
                #endregion

                int data_Count = data.Count;
                #region 當多筆時依客戶權重先排順序 與 初始化
                string sortID = "";
                for (int i = 0; i < data.Count; i++)
                {
                    if (i == 0)
                    { sortID = $"('{data[i]}'"; }
                    else { sortID = $"{sortID},'{data[i]}'"; }
                }
                if (sortID != "") { sortID += ")"; }
                //依客戶檔CTDataWeights欄位由大到小 模擬
                DataTable dt = db.DB_GetData($"select a.Id,a.State,(select CTDataWeights from SoftNetMainDB.[dbo].[CTData] as b where a.NeedSource=b.CTNO) as CTDataWeights FROM SoftNetSYSDB.[dbo].[APS_NeedData] as a where a.ServerId='{_Fun.Config.ServerId}' and a.Id in {sortID}  order by CTDataWeights desc");
                data.Clear();
                sortID = "";
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow d in dt.Rows)
                    {
                        #region 初始化,刪除前次模擬需求碼
                        if (d["State"].ToString() == "1") { meg = $"{meg}<br>{d["Id"].ToString()} 已在模擬,須等完成才能再次模擬"; continue; }
                        else
                        {
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{d["Id"].ToString()}'");
                            db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{d["Id"].ToString()}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{d["Id"].ToString()}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{d["Id"].ToString()}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{d["Id"].ToString()}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{d["Id"].ToString()}'");
                            db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] where Trigger_NeedId='{d["Id"].ToString()}'");
                        }
                        #endregion

                        data.Add(d["Id"].ToString());
                        if (sortID == "")
                        { sortID = $"('{d["Id"].ToString()}'"; }
                        else { sortID = $"{sortID},'{d["Id"].ToString()}'"; }
                    }
                }
                #endregion

                //###??? 將來要加料件 權重先排順序 與 訂單,計畫,補存貨誰先 , 在用基因演算法, 取得多選的優先列表, 先回給使用者調整, 再跑 RunSetSimulation_thread_0(演算程式)

                if (sortID != "") { sortID += ")"; }
                else
                {
                    if (meg.Trim() == "")
                    { meg = $"{meg}<br> 無可安排的排程"; }
                }
                if (meg.Trim() == "")
                {
                    if (!db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='1',KeyA='{String.Join(",", cs.ToArray())}',UpdateTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}',NeedSimulationDate=NULL where ServerId='{_Fun.Config.ServerId}' and Id in {sortID}"))
                    {
                        if (data_Count > 1)
                        {
                            meg = $"{meg}<br> 有部份需求,排程模擬失敗";
                        }
                        else
                        { meg = $"{meg}<br> 排程模擬失敗"; }
                    }
                    else
                    {
                        _SFC_Common.RunSetSimulation(null, ipport, null, '3');
                        if (set1 == "9")
                        { _SFC_Common.RunSetSimulation(args, ipport, data, '9'); }
                        else
                        { _SFC_Common.RunSetSimulation(args, ipport, data, '1'); }
                    }
                }
            }
            return meg;
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private SimulationService EditService()
        {
            return new SimulationService(Ctrl);
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
            //return Json(await EditService().CreateAsync(_Str.ToJson(json)));
            ResultDto row = await EditService().CreateAsync(_Str.ToJson(json));
            if (row.ErrorMsg == "")
            {
                /*
                JObject json2 = _Str.ToJson(json);
                var rows = json2["_rows"] as JArray;
                string[] data = json.Split(';');//['LOGDateTime', 'Id', 'ServerId', 'StationNO', 'OP_NO'];
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    //db.DB_SetData($"Delete FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{data[2]}' and StationNO='{data[3]}' and Id='{data[1]}' and OP_NO='{data[4]}'");
                }
                */
            }

            return Json(row);
        }

        public async Task<JsonResult> Update(string key, string json)
        {
            //return Json(await EditService().UpdateAsync(key, _Str.ToJson(json)));
            ResultDto row = await EditService().UpdateAsync(key, _Str.ToJson(json));
            if (row.ErrorMsg == "")
            {
                JObject json2 = _Str.ToJson(json);
                var rows = json2["_rows"] as JArray;
                if (rows != null && rows.Count>0 && rows[0]["Id"].ToString() != "")
                {
                    string id = rows[0]["Id"].ToString().Trim();

                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                    {


                        db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{id}'");
                        db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{id}'");
                        db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{id}'");
                        db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{id}'");
                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='3',StateINFO='資料修改過',UpdateTime=null where Id='{id}'");
                        db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{id}'"); ;
                        db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{id}'");
                        db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] where Trigger_NeedId='{id}'");

                    }
                }
            }

            return Json(row);
        }

        public async Task<JsonResult> Delete(string key)
        {
            //return Json(await EditService().DeleteAsync(key));
            return Json(await EditService().OtherDeleteAsync(key));
        }
    }
}
