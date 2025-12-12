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
    public class DashboardController : ApiCtrl
    {
        public IActionResult Index()
        {
            return View();
        }
        public ActionResult CapacityIncreaseRate()
        {
            List<IdStrDto> factoryName = new List<IdStrDto>();
            List<IdStrDto> lineName = new List<IdStrDto>();
            List<IdStrDto> orderNo = new List<IdStrDto>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData("SELECT FactoryName FROM SoftNetMainDB.[dbo].[Factory]");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        factoryName.Add(new IdStrDto(dr["FactoryName"].ToString(), dr["FactoryName"].ToString()));
                    }
                }
                dt = db.DB_GetData("SELECT LineName FROM SoftNetSYSDB.[dbo].[PP_Station] group by LineName");
                if (dt != null && dt.Rows.Count > 0)
                {
                    //###???將來要加where FactoryName
                    foreach (DataRow dr in dt.Rows)
                    {
                        lineName.Add(new IdStrDto(dr["LineName"].ToString(), dr["LineName"].ToString()));
                    }
                }
                dt = db.DB_GetData($"SELECT OrderNO FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and EndTime is NULL");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        orderNo.Add(new IdStrDto(dr["OrderNO"].ToString(), dr["OrderNO"].ToString()));
                    }
                }
            }
            ViewBag.OrderNO = orderNo;
            ViewBag.LineName = lineName;
            ViewBag.FactoryName = factoryName;
            return View();
        }

        [HttpPost]
        public string SetStationDashboard(string keys) //工站設定   ipport,站1,站2,,,;工單;製程;作業員
        {
            string re = "";
            string[] data1 = keys.Split(',');

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO like '%%'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    re = $"{dt.Rows.Count.ToString()},";
                    for (int i = 1; i <= dt.Rows.Count; i++)
                    {
                        DataRow dr = dt.Rows[i - 1];
                        re += $"<div id='container{i.ToString()}'></div>";
                        if ( i % 3 == 0 )
                        { re += "<p />"; }
                    }
                }
                else
                { re = "0,"; }
            }

            return re;
        }


        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private DashboardService EditService()
        {
            return new DashboardService(Ctrl);
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
