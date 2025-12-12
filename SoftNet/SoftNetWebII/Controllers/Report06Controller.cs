using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using System;
using System.Text.Encodings.Web;
using Base;

namespace SoftNetWebII.Controllers
{
    public class Report06Controller : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public Report06Controller(SNWebSocketService websocket, SFC_Common sfc_Common)
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

        public ActionResult EditWorkTime()
        {
            List<IdStrDto> ppNameData = new List<IdStrDto>();
            List<IdStrDto> opNODate = new List<IdStrDto>();
            List<IdStrDto> stationNO = new List<IdStrDto>();

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[User] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
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

        private Report04Service EditService()
        {
            return new Report04Service(Ctrl);
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
                int test = -2;
                test += test;
                //JObject json2 = _Str.ToJson(json);
                //var rows = json2["_rows"] as JArray;
                string[] data = key.Split(';');//"ServerId", "LOGDateTime", "LOGDateTimeID", "StationNO", "PartNO"
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataRow dr_log = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{data[0]}' and StationNO='{data[3]}' and LOGDateTimeID='{data[2]}' and PartNO='{data[4]}' and LOGDateTime='{data[1]}'");
                    DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail] where ServerId='{data[0]}' and StationNO='{data[3]}' and LOGDateTime='{data[1]}' and PartNO='{data[4]}'");
                    if (dr != null && dr_log != null)
                    {
                        int okqty = int.Parse(dr_log["EditFinishedQty"].ToString()) + int.Parse(dr_log["OLD_ProductFinishedQty"].ToString()) - int.Parse(dr["ProductFinishedQty"].ToString());
                        //if (okqty > 0) { okqty = int.Parse(dr_log["EditFinishedQty"].ToString()) - okqty; } else { okqty = int.Parse(dr_log["EditFinishedQty"].ToString()) + okqty; }
                        int failqty = int.Parse(dr_log["EditFailedQty"].ToString()) + int.Parse(dr_log["OLD_ProductFailedQty"].ToString()) - int.Parse(dr["ProductFailedQty"].ToString());
                        //if (failqty > 0) { failqty = int.Parse(dr_log["EditFailedQty"].ToString()) - failqty; } else { failqty = int.Parse(dr_log["EditFailedQty"].ToString()) + failqty; }
                        string tmp = "";

                        #region 修正SFC_StationDetail的 CycleTime 與 SFC_StationDetail_ChangeLOG的OLD_ProductFinishedQty,OLD_ProductFailedQty
                        DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{data[0]}' and StationNO='{data[3]}' and PartNO='{data[4]}' and LOGDateTime='{data[1]}' order by LOGDateTimeID");
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            int cts = 0;
                            int qtys = 0;
                            bool is_edit = false;

                            int old_p_fail = 0;
                            int old_p_ok = 0;
                            foreach (DataRow d in dt.Rows)
                            {
                                if (int.Parse(d["EditFinishedQty"].ToString()) > 0 && int.Parse(dr["CycleTime"].ToString()) > 0)
                                {
                                    qtys += int.Parse(d["EditFinishedQty"].ToString());
                                    cts += (int.Parse(d["EditFinishedQty"].ToString()) * int.Parse(dr["CycleTime"].ToString()));
                                }
                                if (is_edit) { db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] SET OLD_ProductFinishedQty={old_p_ok.ToString()},OLD_ProductFailedQty={old_p_fail.ToString()}{tmp} where ServerId='{dr["ServerId"].ToString()}' and LOGDateTime='{Convert.ToDateTime(dr["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and StationNO='{dr["StationNO"].ToString()}' and SerialNO='{dr["SerialNO"].ToString()}'"); }
                                if (d["LOGDateTimeID"].ToString() == data[2])
                                {
                                    is_edit = true;
                                }
                                old_p_ok = int.Parse(d["OLD_ProductFinishedQty"].ToString()) + int.Parse(d["EditFinishedQty"].ToString());
                                old_p_fail = int.Parse(d["OLD_ProductFailedQty"].ToString()) + int.Parse(d["EditFailedQty"].ToString());
                            }
                            if (cts > 0) { tmp = $",CycleTime={(cts / qtys).ToString()}"; }
                        }
                        #endregion

                        if (db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SFC_StationDetail] SET ProductFinishedQty+={okqty.ToString()},ProductFailedQty+={failqty.ToString()}{tmp} where ServerId='{dr["ServerId"].ToString()}' and LOGDateTime='{Convert.ToDateTime(dr["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and StationNO='{dr["StationNO"].ToString()}' and SerialNO='{dr["SerialNO"].ToString()}'"))
                        {
                            //okqty值有正負
                            #region 修正 APS_PartNOTimeNote的Detail_QTY
                            dr_log = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr["SimulationId"].ToString()}'");
                            if (dr_log != null)
                            {
                                int ok_Detail = int.Parse(dr_log["Detail_QTY"].ToString());
                                int fail_Detail = int.Parse(dr_log["Detail_Fail_QTY"].ToString());
                                string wr_sql = "";
                                if (int.Parse(dr_log["Next_StoreQTY"].ToString()) != 0)
                                {
                                    wr_sql = $"Next_StationQTY+={okqty.ToString()},Detail_QTY+={okqty.ToString()}";

                                    dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{data[0]}' and DOCNumberNO='{dr_log["Store_DOCNumberNO"].ToString()}' and SimulationId='{dr_log["SimulationId"].ToString()}' order by IsOK");
                                    if (dt != null && dt.Rows.Count > 0)
                                    {
                                        DataRow dr4 = null;
                                        foreach (DataRow d2 in dt.Rows)
                                        {
                                            if (okqty != 0 && (okqty > 0 || int.Parse(d2["QTY"].ToString()) >= Math.Abs(okqty)))
                                            {

                                                dr4 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{data[0]}' and DOCNumberNO='{dr_log["Store_DOCNumberNO"].ToString()}' and Id='{d2["Id"].ToString()}' and IN_StoreNO='{d2["IN_StoreNO"].ToString()}' and OUT_StoreNO='{d2["OUT_StoreNO"].ToString()}'");
                                                if (bool.Parse(dr4["IsOK"].ToString()))
                                                {
                                                    if (!d2.IsNull("IN_StoreNO") && d2["IN_StoreNO"].ToString() != "")
                                                    { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] SET QTY+={okqty} where ServerId='{data[0]}' and PartNO='{d2["PartNO"].ToString()}' and StoreNO='{d2["IN_StoreNO"].ToString()}' and StoreSpacesNO='{d2["IN_StoreSpacesNO"].ToString()}'"); }
                                                    else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] SET QTY+={okqty} where ServerId='{data[0]}' and PartNO='{d2["PartNO"].ToString()}' and StoreNO='{d2["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{d2["OUT_StoreSpacesNO"].ToString()}'"); }
                                                }
                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] SET QTY+={okqty} where ServerId='{data[0]}' and DOCNumberNO='{dr_log["Store_DOCNumberNO"].ToString()}' and Id='{d2["Id"].ToString()}' and IN_StoreNO='{d2["IN_StoreNO"].ToString()}' and OUT_StoreNO='{d2["OUT_StoreNO"].ToString()}'");
                                                break;
                                            }
                                            else
                                            {
                                                dr4 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{data[0]}' and DOCNumberNO='{dr_log["Store_DOCNumberNO"].ToString()}' and Id='{d2["Id"].ToString()}' and IN_StoreNO='{d2["IN_StoreNO"].ToString()}' and OUT_StoreNO='{d2["OUT_StoreNO"].ToString()}'");
                                                if (bool.Parse(dr4["IsOK"].ToString()))
                                                {
                                                    if (!d2.IsNull("IN_StoreNO") && d2["IN_StoreNO"].ToString() != "")
                                                    { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] SET QTY+={okqty} where ServerId='{data[0]}' and PartNO='{d2["PartNO"].ToString()}' and StoreNO='{d2["IN_StoreNO"].ToString()}' and StoreSpacesNO='{d2["IN_StoreSpacesNO"].ToString()}'"); }
                                                    else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] SET QTY+={okqty} where ServerId='{data[0]}' and PartNO='{d2["PartNO"].ToString()}' and StoreNO='{d2["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{d2["OUT_StoreSpacesNO"].ToString()}'"); }
                                                }
                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] SET QTY=0 where ServerId='{data[0]}' and DOCNumberNO='{dr_log["Store_DOCNumberNO"].ToString()}' and Id='{d2["Id"].ToString()}' and IN_StoreNO='{d2["IN_StoreNO"].ToString()}' and OUT_StoreNO='{d2["OUT_StoreNO"].ToString()}'");
                                                okqty += int.Parse(d2["QTY"].ToString());
                                            }
                                        }
                                    }
                                }
                                else { wr_sql = $"Detail_QTY+={okqty.ToString()}"; }
                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET {wr_sql} where SimulationId='{dr["SimulationId"].ToString()}'");
                            }
                            #endregion
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
                    //db.DB_SetData($"Delete FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{data[2]}' and StationNO='{data[3]}' and Id='{data[1]}' and OP_NO='{data[4]}'");
                }
            }

            return Json(row);

        }
    }
}

