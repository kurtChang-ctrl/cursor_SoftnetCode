using Base;
using Base.Services;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Net.WebSockets;

namespace SoftNetWebII.Controllers
{
    public class TestController : Controller
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public TestController(SNWebSocketService websocket, SFC_Common sfc_Common)
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
        public IActionResult Read()
        {
            return View();
        }
        public IActionResult Read2()
        {
            return View();
        }
        public IActionResult Index1()
        {
            return View();
        }
        public IActionResult Index4()
        {
            string re = "";
            DateTime sTime = new DateTime(2024, 12, 13);
            DateTime eTime = new DateTime(2034, 12, 31);
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow dr_tmp = null;
                while (sTime <= eTime)
                {
                    if (sTime.DayOfWeek != DayOfWeek.Saturday && sTime.DayOfWeek != DayOfWeek.Sunday)
                    {
                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT *  FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and Holiday='{sTime.ToString("yyyy/MM/dd")}' and CalendarName='2021行事曆'");
                        if (dr_tmp==null)
                        {
                            re = $"{re}<br />{sTime.ToString("yyyy/MM/dd")}";
                            db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_HolidayCalendar] ([ServerId],[CalendarName],[Holiday],[Flag_Morning],[Flag_Afternoon],[Flag_Night],[Flag_Graveyard],[Shift_Morning],[Shift_Afternoon],[Shift_Night],[Shift_Graveyard]) 
                                            VALUES ('{_Fun.Config.ServerId}','2021行事曆','{sTime.ToString("yyyy/MM/dd")}','1','1','0','0','08:00,10:00,10:10,12:00','13:00,15:00,15:10,17:00','17:30,19:30,19:40,21:00','23:30,02:30,02:40,06:00')");
                        }
                    }
                    sTime = sTime.AddDays(1);
                }
            }
            ViewBag.HtmlOutput = re;
            return View();
        }
        public IActionResult Index5()
        {
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT *,(Time1_C+Time2_C+Time3_C+Time4_C) as TOT FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where Time1_C>0 or Time2_C>0 or Time3_C>0 or Time4_C>0");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time_TOT={dr["TOT"].ToString()} where Id='{dr["Id"].ToString()}'");
                    }
                }
            }
            return View();
        }
        public IActionResult Index6()
        {
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT *,(EditFinishedQty+EditFailedQty) as QTY FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and EditFinishedQty!=0 and CycleTime!=0");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        #region 計算效能 PP_EfficientDetail處理
                        {
                            List<double> allCT = new List<double>();//list for all avg value
                            string top_flag = "";
                            try
                            {
                                if (_Fun.Config.AdminKey03 != 0) { top_flag = $" TOP {_Fun.Config.AdminKey03} "; }
                                DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} PP_Name,StationNO,PartNO as Sub_PartNO,CycleTime,WaitTime,(EditFinishedQty+EditFailedQty) as QTY from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG]
                                                    where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and PartNO='{dr["PartNO"].ToString()}' and PP_Name='{dr["PP_Name"].ToString()}' and IndexSN={dr["IndexSN"].ToString()} and EditFinishedQty!=0 and CycleTime!=0");
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
                                        _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, dr["StationNO"].ToString(), dr["PP_Name"].ToString(), dr["PP_Name"].ToString(), dr["IndexSN"].ToString(), dr["PartNO"].ToString(), dr["PartNO"].ToString(), "");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs 計算效能PP_EfficientDetail處理 Exception: {ex.Message} {ex.StackTrace}", true);
                            }
                        }
                        #endregion
                    }
                }
            }
            return View();
        }
    }
}
