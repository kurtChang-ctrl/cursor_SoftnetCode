using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class Select3Controller : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;

        public Select3Controller(SNWebSocketService websocket, SFC_Common sfc_Common)
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

        public ActionResult DataUpdate()
        {
            List<IdStrDto> storeNO = new List<IdStrDto>();
            List<IdStrDto> partNO = new List<IdStrDto>();
            List<IdStrDto> storeSpacesNO = new List<IdStrDto>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        storeNO.Add(new IdStrDto(dr["StoreNO"].ToString(), dr["StoreName"].ToString()));
                    }
                }
                dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        partNO.Add(new IdStrDto(dr["PartNO"].ToString(), dr["PartName"].ToString()));
                    }
                }
                dt = db.DB_GetData($"SELECT StoreSpacesName,StoreSpacesNO FROM SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' group by StoreSpacesName,StoreSpacesNO");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        storeSpacesNO.Add(new IdStrDto(dr["StoreSpacesNO"].ToString(), $"{dr["StoreSpacesNO"].ToString()} {dr["StoreSpacesName"].ToString()}" ));
                    }
                }
            }
            ViewBag.StoreNO = storeNO;
            ViewBag.PartNO = partNO;
            ViewBag.StoreSpacesNO = storeSpacesNO;
            return View();
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private Select3Service EditService()
        {
            return new Select3Service(Ctrl);
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
            return Json(await EditService().CreateAsync(_Str.ToJson(json)));
        }

        public async Task<JsonResult> Update(string key, string json)
        {
            return Json(await EditService().UpdateAsync(key, _Str.ToJson(json)));
        }

        public async Task<JsonResult> Delete(string key)
        {
            return Json(await EditService().DeleteAsync(key));
        }
    }
}

