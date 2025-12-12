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
    public class Report05Controller : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public Report05Controller(SNWebSocketService websocket, SFC_Common sfc_Common)
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

        public ActionResult Index()
        {
            List<IdStrDto> oP_NO = new List<IdStrDto>();
            List<IdStrDto> ppNameData = new List<IdStrDto>();
            List<IdStrDto> partNO = new List<IdStrDto>();
            List<IdStrDto> stationNO = new List<IdStrDto>();
            List<IdStrDto> operateType = new List<IdStrDto>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        partNO.Add(new IdStrDto(dr["PartNO"].ToString(), dr["PartNO"].ToString()));
                    }
                }
                tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[User] where ServerId='{_Fun.Config.ServerId}'");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        oP_NO.Add(new IdStrDto(dr["UserNO"].ToString(), $"{dr["UserNO"].ToString()} {dr["Name"].ToString()}" ));
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
            operateType.Add(new IdStrDto("開工", "開工"));
            operateType.Add(new IdStrDto("停工", "停工"));
            operateType.Add(new IdStrDto("關站", "關站"));
            operateType.Add(new IdStrDto("報工", "報工"));
            operateType.Add(new IdStrDto("領料", "領料"));
            operateType.Add(new IdStrDto("入料", "入料"));
            operateType.Add(new IdStrDto("干涉", "干涉"));
            operateType.Add(new IdStrDto("新增項目", "新增項目"));
            operateType.Add(new IdStrDto("設定工站", "設定工站"));


            ViewBag.OP_NO = oP_NO;
            ViewBag.OperateType = operateType;
            ViewBag.PartNO = partNO;
            ViewBag.PPNameData = ppNameData;
            ViewBag.StationNOData = stationNO;

            return View();
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(Ctrl, dt));
        }

        private Report05Service EditService()
        {
            return new Report05Service(Ctrl);
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
    }
}