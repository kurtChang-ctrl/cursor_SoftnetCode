using Base;
using Base.Models;
using Base.Services;
using BaseApi.Services;
using BaseWeb.Models;
using BaseWeb.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting.Internal;

using SoftNetWebII.Models;
using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class ConfigController : Controller
    {
        public IActionResult Index()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }

            return View();
        }
        public IActionResult ImportCustomization()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            return View();
        }

        public IActionResult ThreadState(SystemConfigObj keys)
        {
            if (keys != null && keys.Send_RUNTimeServer_Thread_State != null && keys.Send_RUNTimeServer_Thread_State != "")
            {
                if (keys.Send_RUNTimeServer_Thread_State.ToLower() == "close")
                {
                    _Fun.Is_Thread_ForceClose = true;
                }
            }
            if (keys.FunType != null)
            {
                if (keys.FunType.ToLower() == "ok")
                {
                    _Fun.Is_Thread_For_Test = true;
                }
                else if (keys.FunType.ToLower() == "ng")
                {
                    _Fun.Is_Thread_For_Test = false;
                }
            }
            return View(keys);
        }


        [HttpPost]
        public async Task<string> SetCommand(string ipport, string type, string arg) //0=ipport,1=指令,2=參數
        {
            ViewBag.ERRMsg = "";
            string re = "";
            switch (type)
            {
                case "1"://mail test
                    string mailHtml = $"<form action='http://{_Fun.Config.LocalWebURL}/WorkingPaper/MailAutoAction/;PZZ1BFMJ02V2'><p>測試訊號1</p>><p>測試訊號1</p>><p>測試訊號1</p><hr /><input type=submit value='按此紐 將立即幫助您完成Test發送' style='width:100%';height='60px' /><hr /></form>";
                    await _Log.ErrorAsync(mailHtml, true);
                    break;
                case "2"://Create bom esop folder
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            DataTable tmp_dt = db.DB_GetData($@"select a.Id,a.PartNO,a.Apply_PP_Name,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName from SoftNetMainDB.[dbo].[BOM] as a 
                                                            join SoftNetMainDB.[dbo].[Material] as b on a.ServerId=b.ServerId and a.PartNO=b.PartNO and (b.Class='4' or b.Class='5')
                                                            where a.ServerId='{_Fun.Config.ServerId}' group by a.Id,a.PartNO,a.Apply_PP_Name,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName order by a.PartNO,a.Apply_PP_Name,a.IndexSN");

                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                            {
                                string partESOPName = "";
                                var root = "";
                                foreach (DataRow dr in tmp_dt.Rows)
                                {
                                    if (dr["PartNO"].ToString()== "TG-A08-2")
                                    {
                                        string _s = "";
                                    }
                                    if (dr["StationNO_Custom_DisplayName"].ToString().Trim() != "")
                                    { partESOPName = dr["StationNO_Custom_DisplayName"].ToString().Trim(); }
                                    else if (dr["Station_Custom_IndexSN"].ToString().Trim() != "")
                                    { partESOPName = dr["Station_Custom_IndexSN"].ToString().Trim(); }
                                    else { partESOPName = dr["IndexSN"].ToString().Trim(); }
                                    try
                                    {
                                        partESOPName = partESOPName.Replace("/", "／").Replace("\\", "＼").Replace(":", "：").Replace("*", "＊").Replace("?", "？").Replace("\"", "＂").Replace("<", "＜").Replace(">", "＞").Replace("|", "｜");
                                        if (partESOPName.IndexOf("/") > 0 || partESOPName.IndexOf("\\") > 0 || partESOPName.IndexOf(";") > 0)
                                        { re = $"{re}<br />工序名稱不能有/或\\或;符號. {dr["PartNO"].ToString().Trim()}/{dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}"; }
                                        else
                                        {
                                            root = $"{Directory.GetCurrentDirectory()}/wwwroot/ESOP/{_Fun.Config.ServerId}/{dr["PartNO"].ToString().Trim()}/{dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}";
                                            Directory.CreateDirectory(root);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        re = $"{re}<br />{ex.Message}";
                                    }
                                }
                            }
                        }
                    }
                    break;
                case "3"://複寫BOM ESOP資料夾目錄的文檔至資料庫
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            DataTable tmp_dt = db.DB_GetData($@"select a.Id,a.PartNO,a.Apply_PP_Name,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName from SoftNetMainDB.[dbo].[BOM] as a 
                                                            join SoftNetMainDB.[dbo].[Material] as b on a.ServerId=b.ServerId and a.PartNO=b.PartNO and (b.Class='4' or b.Class='5')
                                                            where a.ServerId='{_Fun.Config.ServerId}' group by a.Id,a.PartNO,a.Apply_PP_Name,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName order by a.PartNO,a.Apply_PP_Name,a.IndexSN");

                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                            {
                                string partESOPName = "";
                                var root = "";
                                string[] files = null;
                                string tmp = "";
                                foreach (DataRow dr in tmp_dt.Rows)
                                {
                                    tmp = "";
                                    if (dr["StationNO_Custom_DisplayName"].ToString().Trim() != "")
                                    { partESOPName = dr["StationNO_Custom_DisplayName"].ToString().Trim(); }
                                    else if (dr["Station_Custom_IndexSN"].ToString().Trim() != "")
                                    { partESOPName = dr["Station_Custom_IndexSN"].ToString().Trim(); }
                                    else { partESOPName = dr["IndexSN"].ToString().Trim(); }
                                    try
                                    {
                                        partESOPName= partESOPName.Replace("/", "／").Replace("\\", "＼").Replace(":", "：").Replace("*", "＊").Replace("?", "？").Replace("\"", "＂").Replace("<", "＜").Replace(">", "＞").Replace("|", "｜");
                                        root = $"{Directory.GetCurrentDirectory()}/wwwroot/ESOP/{_Fun.Config.ServerId}/{dr["PartNO"].ToString().Trim()}/{dr["Apply_PP_Name"].ToString().Trim()}/{partESOPName}";
                                        if (Directory.Exists(root))
                                        {
                                            files = Directory.GetFiles(root);
                                            if (files.Length > 0)
                                            {
                                                foreach (string s in files)
                                                {
                                                    string _s = Path.GetFileName(s);
                                                    if (s.IndexOf("/") > 0)
                                                    {
                                                        if (tmp == "") { tmp = Path.GetFileName(s); }
                                                        else
                                                        { tmp = $"{tmp};{Path.GetFileName(s)}"; }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        re = $"{re}<br />{ex.Message}";
                                    }
                                    if (tmp != "")
                                    { db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set ESOP_Files='{tmp}' where Id='{dr["Id"].ToString()}'"); }
                                    else { db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set ESOP_Files=NULL where Id='{dr["Id"].ToString()}'"); }
                                }
                            }
                        }
                    }
                    break;
                case "4":
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            DataRow tmp_dt2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='02'");
                            if (tmp_dt2 == null) { return re; }
                            DateTime st = new DateTime(2023, 1, 1);
                            DateTime et = new DateTime(2026, 10, 27);
                            while (st <= et)
                            {
                                st = st.AddDays(1);
                                if (st.DayOfWeek == DayOfWeek.Saturday)
                                {
                                    tmp_dt2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday='{st.ToString("yyyy/MM/dd")}' and ServerId='02'");
                                    if (tmp_dt2 == null)
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_HolidayCalendar] ([ServerId],[CalendarName],[Holiday],[Flag_Morning],[Flag_Afternoon],[Flag_Night],[Flag_Graveyard],[Shift_Morning],[Shift_Afternoon],[Shift_Night],[Shift_Graveyard]) VALUES 
('02','2021行事曆','{st.ToString("yyyy/MM/dd")}','1','1','0','0','08:00,10:00,10:10,12:00','13:00,15:00,15:10,17:00','13:00,15:00,15:10,17:00','13:00,15:00,15:10,17:00')");
                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_HolidayCalendar] ([ServerId],[CalendarName],[Holiday],[Flag_Morning],[Flag_Afternoon],[Flag_Night],[Flag_Graveyard],[Shift_Morning],[Shift_Afternoon],[Shift_Night],[Shift_Graveyard]) VALUES 
('02','CNC加工24時行事曆','{st.ToString("yyyy/MM/dd")}','1','1','0','0','08:00,10:00,10:10,12:00','13:00,15:00,15:10,17:00','13:00,15:00,15:10,17:00','13:00,15:00,15:10,17:00')");
                                    }
                                }
                            }
                        }
                    }
                    break;
                case "5":
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            DataTable tmp_dt = db.DB_GetData($"select * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and (Class='6' or Class='7')");

                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                            {
                                string partESOPName = "";
                                var root = "";
                                foreach (DataRow dr in tmp_dt.Rows)
                                {
                                    if (!db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','xxx','','{dr["PartNO"].ToString()}',0)"))
                                    { re = $"{re}<br />{dr["PartNO"].ToString()}沒成功"; }
                                }
                            }
                        }
                    }
                    break;
                case "A1"://刪無用的ManufactureII, 料件有APS_Default_StoreNO,APS_Default_StoreSpacesNO檢查TotalStock是否有資料
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            DataRow tmp = null;
                            #region  刪無用ManufactureII
                            DataTable dt_ManufactureII2 = null;
                            DataTable dt = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and State='9'");
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                foreach (DataRow dr_II in dt.Rows)
                                {
                                    dt_ManufactureII2 = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_II["Id"].ToString()}'");
                                    if (dt_ManufactureII2 != null && dt_ManufactureII2.Rows.Count > 0)
                                    {
                                        string ii_ID = "";
                                        foreach (DataRow dr_III in dt_ManufactureII2.Rows)
                                        {
                                            if (ii_ID == "") { ii_ID = $"'{dr_III["SimulationId"].ToString()}'"; }
                                            else { ii_ID = $"{ii_ID},'{dr_III["SimulationId"].ToString()}'"; }
                                        }
                                        if (ii_ID != "")
                                        {
                                            db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[ManufactureII] where SimulationId in ({ii_ID})");
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region 檢查TotalStock是否有資料
                            dt = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and APS_Default_StoreNO!=''");
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                string tmp_sq= "";
                                foreach (DataRow dr_II in dt.Rows)
                                {
                                    if (dr_II["PartNO"].ToString()== "SBG-03-10")
                                    {
                                        string _s = "";
                                    }
                                    if (dr_II["APS_Default_StoreSpacesNO"].ToString() != "") { tmp_sq = $" and StoreNO='{dr_II["APS_Default_StoreNO"].ToString()}' and StoreSpacesNO='{dr_II["APS_Default_StoreSpacesNO"].ToString()}'"; }
                                    else { tmp_sq = $" and StoreNO='{dr_II["APS_Default_StoreNO"].ToString()}'"; }
                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_II["PartNO"].ToString()}' {tmp_sq}");
                                    if (tmp==null)
                                    {
                                       if(db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{dr_II["APS_Default_StoreNO"].ToString()}','{dr_II["APS_Default_StoreSpacesNO"].ToString()}','{dr_II["PartNO"].ToString()}',0)"))
                                        {
                                            string _s = "";
                                        }
                                    }
                                }
                            }
                            #endregion

                            //#region 檢查DOC3stockII 的IN_StoreNO,OUT_StoreNO
                            //dt = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}'");
                            //if (dt != null && dt.Rows.Count > 0)
                            //{
                            //    string tmp_sq = "";
                            //    foreach (DataRow dr_II in dt.Rows)
                            //    {
                            //        tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_II["PartNO"].ToString()}'");
                            //        if (tmp!=null)
                            //        {
                            //            if (tmp["APS_Default_StoreNO"].ToString()!="")
                            //            {
                            //                if (dr_II["IN_StoreNO"].ToString()!="" && dr_II["IN_StoreNO"].ToString()!= tmp["APS_Default_StoreNO"].ToString() && dr_II["IN_StoreSpacesNO"].ToString() != tmp["APS_Default_StoreSpacesNO"].ToString())
                            //                {
                            //                    db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set IN_StoreNO='{tmp["APS_Default_StoreNO"].ToString()}',IN_StoreSpacesNO='{tmp["APS_Default_StoreSpacesNO"].ToString()}' where Id='{dr_II["Id"].ToString()}' and  DOCNumberNO='{dr_II["DOCNumberNO"].ToString()}' and  ArrivalDate='{Convert.ToDateTime(dr_II["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and  PartNO='{dr_II["PartNO"].ToString()}'");
                            //                }
                            //                if (dr_II["OUT_StoreNO"].ToString() != "" && dr_II["OUT_StoreNO"].ToString() != tmp["APS_Default_StoreNO"].ToString() && dr_II["OUT_StoreSpacesNO"].ToString() != tmp["APS_Default_StoreSpacesNO"].ToString())
                            //                {
                            //                    db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set OUT_StoreNO='{tmp["APS_Default_StoreNO"].ToString()}',OUT_StoreSpacesNO='{tmp["APS_Default_StoreSpacesNO"].ToString()}' where Id='{dr_II["Id"].ToString()}' and  DOCNumberNO='{dr_II["DOCNumberNO"].ToString()}' and  ArrivalDate='{Convert.ToDateTime(dr_II["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and  PartNO='{dr_II["PartNO"].ToString()}'");
                            //                }

                            //            }
                            //        }

                            //    }
                            //}
                            //#endregion
                        }
                    }
                    break;
                case "99"://製程回寫BOM StationNO_Merge,MFNO
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            DataTable dt_BOM = null;
                            /*
                            DataTable dt_BOM = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and (IndexSN_Merge='1' or StationNO='{_Fun.Config.OutPackStationName}') order by PP_Name,IndexSN");
                            if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                            {
                                DataRow tmp = null;
                                bool is_f = true;
                                string indexSN_Merge = "";
                                string b_id = dt_BOM.Rows[0]["BOMId"].ToString();
                                for (int i=0;i<dt_BOM.Rows.Count;i++)
                                {
                                    DataRow dr = dt_BOM.Rows[i];

                                    if (dr["StationNO"].ToString() == _Fun.Config.OutPackStationName)
                                    {
                                        if (dr["MFNO"].ToString() != "")
                                        { db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set MFNO='{dr["MFNO"].ToString()}' where ServerId='{_Fun.Config.ServerId}' and Id='{dr["BOMId"].ToString()}'"); }
                                        continue;
                                    }

                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{dr["BOMId"].ToString()}'");
                                    if (tmp != null)
                                    {
                                        if (dr["StationNO"].ToString() == tmp["Apply_StationNO"].ToString()) { continue; }
                                    }
                                    else { continue; }
                                    if (is_f || b_id == dr["BOMId"].ToString())
                                    {
                                        if (is_f) { b_id = dr["BOMId"].ToString(); }
                                        is_f = false;
                                        if (indexSN_Merge.IndexOf($"{dr["StationNO"].ToString()},") < 0)
                                        { indexSN_Merge = $"{indexSN_Merge}{dr["StationNO"].ToString()},"; }
                                    }
                                    else
                                    {
                                        is_f = true;
                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set StationNO_Merge='{indexSN_Merge}' where ServerId='{_Fun.Config.ServerId}' and Id='{b_id}'");
                                        indexSN_Merge = "";
                                        i -= 1;
                                    }
                                }
                            }
                            */

                            {
                                List<string> bomII = new List<string>();
                                dt_BOM = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Main_Item='1' order by Apply_PP_Name");
                                if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                                {
                                    DataRow tmp = null;
                                    bool iserr = false;
                                    string isEnd_ID = "";
                                    List<string> log_BOM_All_id = new List<string>();
                                    int indexSN = 0;
                                    foreach (DataRow dr_B in dt_BOM.Rows)
                                    {
                                        log_BOM_All_id.Clear();
                                        #region 檢查IndexSN
                                        DataRow dr_BOM = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{dr_B["Id"].ToString()}'");
                                        if (dr_BOM != null)
                                        {
                                            indexSN = int.Parse(dr_BOM["IndexSN"].ToString()) - 1;
                                            log_BOM_All_id.Add(dr_B["Id"].ToString());
                                            for (int i = indexSN; i >= 1; i--)
                                            {
                                                if (db.DB_GetQueryCount($"select Id from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_BOM["PartNO"].ToString()}' and Apply_PP_Name='{dr_BOM["Apply_PP_Name"].ToString()}' and IndexSN={i}") != 1)
                                                { iserr = true; break; }
                                                else
                                                {
                                                    tmp = db.DB_GetFirstDataByDataRow($"select Id from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_BOM["PartNO"].ToString()}' and Apply_PP_Name='{dr_BOM["Apply_PP_Name"].ToString()}' and IndexSN={i}");
                                                    log_BOM_All_id.Add(tmp["Id"].ToString());
                                                    if (i == 1)
                                                    {
                                                        isEnd_ID = tmp["Id"].ToString();
                                                    }
                                                }
                                            }
                                        }
                                        else { iserr = true; }
                                        #endregion
                                        if (iserr)
                                        {
                                            string _s1 = "此BOM表資料有問題無法發行, 請聯繫系統管理員.";
                                            if (!bomII.Contains(dr_BOM["Apply_PP_Name"].ToString()))
                                            {
                                                bomII.Add(dr_BOM["Apply_PP_Name"].ToString());
                                                re = $"{re}<br />{dr_BOM["Apply_PP_Name"].ToString()} 失敗.";
                                            }

                                        }
                                        else
                                        {
                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IsConfirm='1' where ServerId='{_Fun.Config.ServerId}' and Id='{dr_B["Id"].ToString()}'");
                                            if (isEnd_ID != "") { db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IsEnd='1' where ServerId='{_Fun.Config.ServerId}' and Id='{isEnd_ID}'"); }

                                            #region 重新整理製程檔
                                            db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[PP_ProductProcess] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{dr_BOM["Apply_PP_Name"].ToString()}' and Apply_PartNO='{dr_BOM["Apply_PartNO"].ToString()}'");
                                            db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{dr_BOM["Apply_PP_Name"].ToString()}' and Apply_PartNO='{dr_BOM["Apply_PartNO"].ToString()}'");
                                            db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess] (ServerId,PP_Name,Apply_PartNO,CalendarName,UpdateTime,FactoryName,LineName) VALUES
                                            ('{_Fun.Config.ServerId}','{dr_BOM["Apply_PP_Name"].ToString()}','{dr_BOM["Apply_PartNO"].ToString()}','{_Fun.Config.DefaultCalendarName}','{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}')");
                                            string indexSN_Merge = "0";
                                            string outPackType = "0";
                                            foreach (string id in log_BOM_All_id)
                                            {
                                                tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{id}'");
                                                if (tmp.IsNull("StationNO_Merge")) { indexSN_Merge = "0"; } else { indexSN_Merge = "1"; }
                                                if (tmp["Apply_StationNO"].ToString() != _Fun.Config.OutPackStationName) { outPackType = "0"; } else { outPackType = "1"; }
                                                db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,Apply_PartNO,PP_Name,PartNO,StationNO,IndexSN,DisplaySN,Station_Custom_IndexSN,DisplayName,IndexSN_Merge,OutPackType,MFNO,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{tmp["Apply_PartNO"].ToString()}','{tmp["Apply_PP_Name"].ToString()}','{tmp["PartNO"].ToString()}','{tmp["Apply_StationNO"].ToString()}',{tmp["IndexSN"].ToString()},0,'{tmp["Station_Custom_IndexSN"].ToString()}','{tmp["StationNO_Custom_DisplayName"].ToString()}','{indexSN_Merge}','{outPackType}','{tmp["MFNO"].ToString()}','{tmp["Id"].ToString()}')");
                                                if (!tmp.IsNull("StationNO_Merge") && tmp["StationNO_Merge"].ToString() != "")
                                                {
                                                    foreach (string s in tmp["StationNO_Merge"].ToString().Split(','))
                                                    {
                                                        if (s.Trim() == "") { continue; }
                                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,Apply_PartNO,PP_Name,PartNO,StationNO,IndexSN,DisplaySN,Station_Custom_IndexSN,DisplayName,IndexSN_Merge,OutPackType,MFNO,BOMId) VALUES
                                                                        ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{tmp["Apply_PartNO"].ToString()}','{tmp["Apply_PP_Name"].ToString()}','{tmp["PartNO"].ToString()}','{s}',{tmp["IndexSN"].ToString()},0,'{tmp["Station_Custom_IndexSN"].ToString()}','{tmp["StationNO_Custom_DisplayName"].ToString()}','{indexSN_Merge}','{outPackType}','{tmp["MFNO"].ToString()}','{tmp["Id"].ToString()}')");
                                                    }
                                                }
                                            }
                                            DataTable dt_BOM3 = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{tmp["Apply_PP_Name"].ToString()}' order by IndexSN");
                                            if (dt_BOM3 != null && dt_BOM3.Rows.Count > 0)
                                            {
                                                int displayNO = 0;
                                                foreach (DataRow row in dt_BOM3.Rows)
                                                {
                                                    displayNO += 1;
                                                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] set DisplaySN={displayNO.ToString()} where ServerId='{_Fun.Config.ServerId}' and Id='{row["Id"].ToString()}'");
                                                }
                                            }
                                            #endregion
                                        }
                                    }
                                }
                            }

                        }
                    }
                    break;
                case "98"://大正寫A01~A05共用站
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            DataTable dt_BOM = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_StationNO like 'A%' order by Apply_PP_Name");
                            if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                            {
                                string StationNO_Merge = "";
                                foreach (DataRow dr in dt_BOM.Rows)
                                {
                                    StationNO_Merge = "";
                                    switch (dr["Apply_StationNO"].ToString())
                                    {
                                        case "A01": StationNO_Merge = "A02,A03,A04,A05,"; break;
                                        case "A02": StationNO_Merge = "A01,A03,A04,A05,"; break;
                                        case "A03": StationNO_Merge = "A02,A01,A04,A05,"; break;
                                        case "A04": StationNO_Merge = "A02,A03,A01,A05,"; break;
                                        case "A05": StationNO_Merge = "A02,A03,A04,A01,"; break;
                                        default:
                                            string _s = "";
                                            break;
                                    }
                                    if (StationNO_Merge!="")
                                    {
                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set StationNO_Merge='{StationNO_Merge}' where Id='{dr["Id"].ToString()}'");
                                    }
                                    else
                                    {
                                        string _s = "";
                                    }

                                }
                            }
                        }
                    }
                    break;
                case "97":
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            DataTable dt_BOM = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and Source_StationNO like 'A%' order by Apply_PP_Name");
                            if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                            {
                                string StationNO_Merge = "";
                                foreach (DataRow dr in dt_BOM.Rows)
                                {
                                    StationNO_Merge = "";
                                    switch (dr["Source_StationNO"].ToString())
                                    {
                                        case "A01": StationNO_Merge = "A02,A03,A04,A05,"; break;
                                        case "A02": StationNO_Merge = "A01,A03,A04,A05,"; break;
                                        case "A03": StationNO_Merge = "A02,A01,A04,A05,"; break;
                                        case "A04": StationNO_Merge = "A02,A03,A01,A05,"; break;
                                        case "A05": StationNO_Merge = "A02,A03,A04,A01,"; break;
                                        default:
                                            string _s = "";
                                            break;
                                    }
                                    if (StationNO_Merge != "")
                                    {
                                        db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set StationNO_Merge='{StationNO_Merge}' where SimulationId='{dr["SimulationId"].ToString()}'");
                                        db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set StationNO_Merge='{StationNO_Merge}' where SimulationId='{dr["SimulationId"].ToString()}'");
                                        db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set StationNO_Merge='{StationNO_Merge}' where SimulationId='{dr["SimulationId"].ToString()}'");
                                    }
                                    else
                                    {
                                        string _s = "";
                                    }

                                }
                            }
                        }
                    }
                    break;
                case "96"://確認BOMII的class
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            DataTable dt_BOM = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}'");
                            if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                            {
                                DataRow dr_Material = null;
                                foreach (DataRow dr in dt_BOM.Rows)
                                {
                                    dr_Material = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                                    if (dr_Material != null)
                                    {
                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[BOMII] set Class='{dr_Material["Class"].ToString()}' where Id='{dr["Id"].ToString()}'");
                                    }
                                }
                            }
                        }
                    }
                    break;
                case "95"://修正干涉停工時間
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            DataTable dt_BOM = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[OperateLog] where OperateType like '%停工%' and StationNO!='B01' and StationNO!='B02' order by LOGDateTime ");
                            if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                            {
                                DateTime time = DateTime.Now;
                                foreach (DataRow dr in dt_BOM.Rows)
                                {
                                    time = Convert.ToDateTime(dr["LOGDateTime"]);
                                    if (time.Hour>17 && time.Minute > 30)
                                    {
                                        string _s = "";
                                    }
                                }
                            }
                        }
                    }
                    break;
                case "94"://查TotalStock 
                    {
                        using (DBADO db = new DBADO("1", _Fun.Config.Db))
                        {
                            DataTable dt_BOM = db.DB_GetData($"select ServerId,StoreNO,StoreSpacesNO,PartNO,count(*) as TOT from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' group by ServerId,StoreNO,StoreSpacesNO,PartNO");
                            if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                            {
                                DataRow tmp = null;
                                foreach (DataRow dr in dt_BOM.Rows)
                                {
                                    if (int.Parse(dr["TOT"].ToString()) > 1)
                                    {
                                        //tmp = db.DB_GetFirstDataByDataRow($@"select top 1 a.* from SoftNetMainDB.[dbo].[TotalStock] as a
                                        //        join SoftNetMainDB.[dbo].[TotalStockII] as b on a.Id!=b.Id 
                                        //        where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{dr["StoreNO"].ToString()}' and a.StoreSpacesNO='{dr["StoreSpacesNO"].ToString()}' and a.PartNO='{dr["PartNO"].ToString()}'");
                                        tmp = db.DB_GetFirstDataByDataRow($@"select top 1 a.* from SoftNetMainDB.[dbo].[TotalStock] as a
                                            where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{dr["StoreNO"].ToString()}' and a.StoreSpacesNO='{dr["StoreSpacesNO"].ToString()}' and a.PartNO='{dr["PartNO"].ToString()}'");
                                        if (tmp != null)
                                        { db.DB_SetData($"delete from SoftNetMainDB.[dbo].[TotalStock] where Id='{tmp["Id"].ToString()}'"); }
                                    }
                                }
                            }
                        }
                    }
                    break;
                case "a":
                    re = "廠商資料表 Excel格式不正確";
                    break;
                case "b":
                    re = "員工資料表 Excel格式不正確";
                    break;
                case "c":
                    re = "工站基本資料 Excel格式不正確";
                    break;
                case "d":
                    re = "產品結構 Excel格式不正確";
                    break;
                case "e":
                    re = "生產工序表 Excel格式不正確";
                    break;

            }
            return re;
        }

        [HttpPost]
        public async Task<JsonResult> ImportExcel_A(IFormFile file)
        {
            string type = "大正_庫存數量導入";
            return Json(await Go_ImportExcel(file, type));
        }
        public async Task<JsonResult> ImportExcel_B(IFormFile file)
        {
            string type = "大正_成品_3";
            return Json(await Go_ImportExcel(file, type));
        }
        public async Task<JsonResult> ImportExcel_C(IFormFile file)
        {
            string type = "大正_素材_3";
            return Json(await Go_ImportExcel(file, type));
        }
        public async Task<JsonResult> ImportExcel_D(IFormFile file)
        {
            string type = "大正_採購半成品_3";
            return Json(await Go_ImportExcel(file, type));
        }
        public async Task<JsonResult> ImportExcel_E(IFormFile file)
        {
            string type = "大正_機械倉_關聯料號";
            return Json(await Go_ImportExcel(file, type));
        }
        public async Task<JsonResult> ImportExcel_F(IFormFile file)
        {
            string type = "大正_電子燈_關聯料號";
            return Json(await Go_ImportExcel(file, type));
        }



        private async Task<ResultImportDto> Go_ImportExcel(IFormFile file,string type)
        {
            string uname = "";
            if (_Fun.GetBaseUser() != null)
            { uname = _Fun.GetBaseUser().UserName; }
            if (type=="")
            {
                ResultImportDto tmp = new ResultImportDto()
                {
                    ErrorMsg = "格式無法判讀, 請重新選擇正確檔案.",
                };
                return tmp;
            }
            var importDto = new ExcelImportDto<SimulationDto>()
            {
                //###??? 依需求命名   ImportType = "PPName_AND_BOM",
                ImportType = type,//"萬_PPName_AND_BOM_二次",
                TplPath = _Xp.DirTpl + "UserImport.xlsx",
                FnSaveImportRows = SaveImportRows,
                CreatorName = uname,
            };
            string dirUpload = @"D:\書_sampleCode\HrAdm-master\_upload\UserImport\";
            var model = await _WebExcel.ImportByFileAsync(file, dirUpload, importDto);
            return model;
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

    }
}
