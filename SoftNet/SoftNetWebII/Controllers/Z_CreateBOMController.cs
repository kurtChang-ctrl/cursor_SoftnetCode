using Base.Services;
using BaseApi.Controllers;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Models;
using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Data;
using System;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Net.WebSockets;
using Base;

namespace SoftNetWebII.Controllers
{
    public class Z_CreateBOMController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public Z_CreateBOMController(SNWebSocketService websocket, SFC_Common sfc_Common)
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
        public ActionResult Index1(Z_CreateBOM_Data keys)
        {
            keys.MES_Report = "";
            keys.ERRMsg = "";
            List<string[]> StationNOList = new List<string[]>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and Station_Type='8'");
                DataRow tmp = null;
                if (dt != null && dt.Rows.Count > 0)
                {
                    #region 確認 Manufacture資料 新增StationNOList資料

                    //###??? 已派工的車要拿掉
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
            }
            keys.StationNOList = StationNOList;
            return View(keys);
        }


        [HttpPost]
        public ActionResult SetAction(Z_CreateBOM_Data keys)
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home");
            }
            string sql = "";
            keys.ERRMsg = "";
            keys.MES_Report = "";
            if (keys.StationNO == null || keys.StationNO.Trim() == "")
            {
                keys.ERRMsg = "你沒有選擇台車 或 畫面逾時, 請按 回到工站選單畫面 重新選擇工站.";
            }
            else
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    try
                    {
                        DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}'");
                        if (dr_PP_Station == null) { keys.ERRMsg = $"查無 {keys.StationNO} 台車工站資料紀錄, 請通知管理者."; goto break_FUN; }

                        keys.StationNO = "01車";
                        int isARGs10_offset = 15;//###??? 15將來改參數

                        string ids = _Str.NewId('X');
                        sql = @$"INSERT INTO SoftNetSYSDB.[dbo].[APS_NeedData] (UpdateTime,[ServerId],[Id],[NeedType],[NeedSource],[NeedDate],CalendarName,[PartNO],[NeedQTY],[BOMId],[Apply_PP_Name],[BufferTime],[ChangeQTY],KeyA,FactoryName) 
                                 VALUES ('{DateTime.Now.AddMinutes(isARGs10_offset).ToString("yyyy-MM-dd HH:mm:ss")}','{_Fun.Config.ServerId}','{ids}','5','電視拆解','{DateTime.Now.AddDays(3).ToString("yyyy/MM/dd HH:mm:ss")}','2021行事曆','{keys.StationNO}',{keys.QTY.ToString()},'Z001A4PB1C2Q','電視拆解製程',8,0,'0,1,0,0,0,0,1,1,0,1,1,0,1,0,0,0,0,0,0,0','{_Fun.Config.DefaultFactoryName}')";
                        db.DB_SetData(sql);


                        string meg = "";
                        string arg = "01000011011010000000";
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



                        string sortID = "";
                        for (int i = 0; i < data.Count; i++)
                        {
                            if (i == 0)
                            { sortID = $"('{data[i]}'"; }
                            else { sortID = $"{sortID},'{data[i]}'"; }
                        }
                        if (sortID != "") { sortID += ")"; }

                        if (meg.Trim() == "")
                        {
                            if (!db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='1',KeyA='{String.Join(",", cs.ToArray())}',UpdateTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}',NeedSimulationDate=NULL where ServerId='{_Fun.Config.ServerId}' and Id in {sortID}"))
                            { keys.ERRMsg = $"排程模擬失敗, 請通知管理者."; }
                            else
                            {
                                List<string> needID_data = new List<string>() { ids };
                                _SFC_Common.RunSetSimulation(args, "", needID_data, '5');
                            }
                        }




                    }
                    catch (Exception ex)
                    {
                        keys.ERRMsg = $"程式異常, 導致系統無法使用, 請通知管理者.";
                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs SetAction Exception: {ex.Message} {ex.StackTrace}", true);
                    }
                break_FUN:
                    string _s = "";
                }
            }
            return View("ResuItTimeOUT");
        }
    }
}
