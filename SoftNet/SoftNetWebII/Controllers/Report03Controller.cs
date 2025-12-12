using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;

using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Net.WebSockets;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class Report03Controller : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public Report03Controller( SNWebSocketService websocket, SFC_Common sfc_Common)
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

        public ActionResult Index01()
        {
            List<IdStrDto> ppNameData = new List<IdStrDto>();
            List<IdStrDto> opNODate = new List<IdStrDto>();
            List<IdStrDto> stationNO = new List<IdStrDto>();

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[User] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count>0)
                {
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        opNODate.Add(new IdStrDto(dr["UserNO"].ToString(), dr["Name"].ToString()));
                    }
                }
                tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_ProductProcess] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        ppNameData.Add(new IdStrDto(dr["PP_Name"].ToString(), dr["PP_Name"].ToString()));
                    }
                }
                tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        //stationNO.Add(new IdStrDto(dr["StationNO"].ToString(),$"{dr["StationNO"].ToString()} {dr["StationName"].ToString()}" ));
                        stationNO.Add(new IdStrDto(dr["StationNO"].ToString(), dr["StationName"].ToString()));
                    }
                }
            }
            ViewBag.OPNOName = opNODate;
            ViewBag.PPNameData = ppNameData;
            ViewBag.StationNOData = stationNO;

            return View();
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(Ctrl, dt));
        }

        private Report03Service EditService()
        {
            return new Report03Service(Ctrl);
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
            //return Json(await EditService().UpdateAsync(key, _Str.ToJson(json)));
            ResultDto row = await EditService().UpdateAsync(key, _Str.ToJson(json));
            if (row.ErrorMsg == "")
            {
                //JObject json2 = _Str.ToJson(json);
                //var rows = json2["_rows"] as JArray;
                string[] data = key.Split(';');//['LOGDateTime', 'Id', 'ServerId', 'StationNO', 'OP_NO'];
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataRow dr_log = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{data[2]}' and StationNO='{data[3]}' and Id='{data[1]}' and OP_NO='{data[4]}' and LOGDateTime='{data[0]}'");
                    DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{data[2]}' and StationNO='{data[3]}' and Id='{data[1]}' and OP_NO='{data[4]}'");
                    if (dr != null && dr_log != null)
                    {
                        int okqty =int.Parse(dr_log["OKQTY"].ToString()) - int.Parse(dr["ProductFinishedQty"].ToString());
                        int failqty = int.Parse(dr_log["FailQTY"].ToString()) - int.Parse(dr["ProductFailedQty"].ToString());
                        string tmp = "";
                        if (!dr.IsNull("RemarkTimeS") && !dr.IsNull("RemarkTimeE"))
                        {
                            int ctQTY = int.Parse(dr["ProductFinishedQty"].ToString()) + int.Parse(dr["ProductFailedQty"].ToString()) + okqty + failqty;
                            dr_log = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{data[2]}' and StationNO='{data[3]}'");
                            decimal ct = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_log["CalendarName"].ToString(), Convert.ToDateTime(dr["RemarkTimeS"]), Convert.ToDateTime(dr["RemarkTimeE"])) / ctQTY;
                            decimal ct_log = ct < 1 ? 0 : ct;
                            if (int.Parse(dr["CycleTime"].ToString()) != 0) { ct = (ct + int.Parse(dr["CycleTime"].ToString())) > 0 ? Math.Round((ct + int.Parse(dr["CycleTime"].ToString())) / 2) : ct; }
                            if (ct < 1) { ct = 0; }
                            tmp = $",CycleTime={ct.ToString()}";
                        }
                        if (db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SFC_StationProjectDetail] SET ProductFinishedQty+={okqty.ToString()},ProductFailedQty+={failqty.ToString()}{tmp} where ServerId='{dr["ServerId"].ToString()}' and Id='{dr["Id"].ToString()}' and StationNO='{dr["StationNO"].ToString()}' and OP_NO='{dr["OP_NO"].ToString()}'"))
                        {
                        }

                    }
                }
            }

            return Json(row);
        }

        //刪除(DB)
        public async Task<JsonResult> Delete(string key)
        {
            //return Json(await EditService().DeleteAsync(key));
            ResultDto row = await EditService().DeleteAsync(key);
            if (row.ErrorMsg == "")
            {
                //JObject json2 = _Str.ToJson(json);
                //var rows = json2["_rows"] as JArray;
                string[] data = key.Split(';');//['LOGDateTime', 'Id', 'ServerId', 'StationNO', 'OP_NO'];
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    db.DB_SetData($"Delete FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{data[2]}' and StationNO='{data[3]}' and Id='{data[1]}' and OP_NO='{data[4]}'");
                }
            }

            return Json(row);

        }
    }

}
