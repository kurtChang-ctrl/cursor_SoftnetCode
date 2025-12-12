using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;

using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class StationController : ApiCtrl
    {
        public ActionResult Read()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            List<IdStrDto> station_Type = new List<IdStrDto>();
            station_Type.Add(new IdStrDto("1", "作業站"));
            station_Type.Add(new IdStrDto("2", "維修站"));
            station_Type.Add(new IdStrDto("3", "控制站"));
            station_Type.Add(new IdStrDto("7", "虛擬站"));
            station_Type.Add(new IdStrDto("8", "多工站"));
            ViewBag.Station_Type = station_Type;

            List<IdStrDto> stationUI_type = new List<IdStrDto>();
            stationUI_type.Add(new IdStrDto("1", "追朔"));
            stationUI_type.Add(new IdStrDto("2", "不追朔"));
            ViewBag.StationUI_type = stationUI_type;

            List<IdStrDto> factoryName = new List<IdStrDto>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT FactoryName FROM [dbo].[Factory] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    int i = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        factoryName.Add(new IdStrDto(dr["FactoryName"].ToString(), dr["FactoryName"].ToString()));
                    }
                }
            }
            ViewBag.FactoryName = factoryName;

            List<IdStrDto> calendarName = new List<IdStrDto>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT ServerId,CalendarName FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' group by ServerId,CalendarName");
                if (dt != null && dt.Rows.Count > 0)
                {
                    int i = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        calendarName.Add(new IdStrDto(dr["CalendarName"].ToString(), dr["CalendarName"].ToString()));
                    }
                }
            }
            ViewBag.CalendarName = calendarName;

            return View();
        }
        [HttpPost]
        public string Update_Manufacture(string keys)
        {
            string meg = "";

            string[] data = keys.Split(',');
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string sql = $"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Config_WaitTime_Formula = '{data[1]}', Config_CycleTime_Formula = '{data[2]}', Config_CalculatePauseTime = '{data[3]}', Config_OutQtyTagName = '{data[4]}', Config_FailQtyTagName = '{data[5]}', Config_WaitTagValueDIO = '{data[6]}', Config_IsTagValueCumulative = {data[7]}, Config_IsWaitTargetFinish = '{data[8]}'";
                sql= $"{sql} where ServerId='{_Fun.Config.ServerId}' and StationNO = '{data[0]}'";
                if (!db.DB_SetData(sql))
                { return "寫入失敗!"; }
            }
            return meg;
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private StationService EditService()
        {
            return new StationService(Ctrl);
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
            return Json(await EditService().OtherCrateAsync(_Str.ToJson(json)));
        }

        public async Task<JsonResult> Update(string key, string json)
        {
            ResultDto row = await EditService().UpdateAsync(key, _Str.ToJson(json));
            if (row.ErrorMsg=="")
            {
                JObject json2 = _Str.ToJson(json);
                var rows = json2["_rows"] as JArray;
                string Config_MutiWO = "0";
                if (rows[0]["Station_Type"].ToString().Trim() == "8") { Config_MutiWO = "1"; }
                string sql = $"UPDATE SoftNetMainDB.[dbo].[Manufacture] Set Station_Type='{Config_MutiWO}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{rows[0]["StationNO"].ToString()}'";
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    db.DB_SetData(sql);
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
