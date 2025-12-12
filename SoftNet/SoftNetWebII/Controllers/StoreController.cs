using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using SoftNetWebII.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Threading.Tasks;

using System.Data;
using Newtonsoft.Json.Linq;
using System.Linq;
using Base;

namespace SoftNetWebII.Controllers
{
    //[XgProgAuth]
    public class StoreController : ApiCtrl
    {
        public async Task<ActionResult> Read()
        {
            //test
            //_Fun.Except("exception test");
            List<IdStrDto> className = new List<IdStrDto>();
            className.Add(new IdStrDto("實體倉", "實體倉"));
            className.Add(new IdStrDto("虛擬倉", "虛擬倉"));
            className.Add(new IdStrDto("大正倉", "大正倉"));
            ViewBag.Class = className;

            List<IdStrDto> default_IN_OUT = new List<IdStrDto>();
            default_IN_OUT.Add(new IdStrDto("0", "不限定類型"));
            default_IN_OUT.Add(new IdStrDto("1", "適用原物料"));
            default_IN_OUT.Add(new IdStrDto("2", "適用生產/加工件"));
            default_IN_OUT.Add(new IdStrDto("3", "大正智能倉"));
            default_IN_OUT.Add(new IdStrDto("4", "限定刀工具"));
            default_IN_OUT.Add(new IdStrDto("5", "適用成品倉"));

            ViewBag.Default_IN_OUT = default_IN_OUT;

            List<IdStrDto> factoryName = new List<IdStrDto>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData("SELECT FactoryName FROM [dbo].[Factory]");
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

            return View();
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(Ctrl, dt));
        }

        private StoreService EditService()
        {
            return new StoreService(Ctrl);
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
            //JObject data = _Str.ToJson(json);
            //var rows_row = data["_rows"] as JArray;
            //var rows_child =data["_childs"] as JArray;
            //if ((JToken)data["_childs"] != null)
            //{
            //    JToken item = (JToken)data["_childs"];
            //    if (item != null)
            //    {
            //        foreach (JToken item2 in item)
            //        {
            //            if (item2["StoreSpacesNO"] != null)
            //            {
            //                string _s = item2["StoreSpacesNO"].ToString();
            //            }
            //        }
            //    }
            //}

            ResultDto row = await EditService().UpdateAsync(key, _Str.ToJson(json));
            if (row.ErrorMsg == "")
            {
                string[] data=key.Split(';');
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataTable dt = db.DB_GetData($"select StoreSpacesNO,count(*) as tot FROM SoftNetMainDB.[dbo].[StoreII] where ServerId='{data[0]}' and StoreNO='{data[1]}' group by StoreSpacesNO");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        foreach (DataRow d in dt.Rows)
                        {
                            if (!d.IsNull("tot") && d["tot"].ToString()!="" && int.Parse(d["tot"].ToString())!=1)
                            {
                                row.ErrorMsg = "異動後, 發現儲位有重複編號, 請重新進入修改功能查閱與修正, 否則系統進出料會有問題, 或請與系統管理者聯繫.";
                            }
                        }
                    }
                }
            }
            return Json(row);
            //return Json(await EditService().UpdateAsync(key, _Str.ToJson(json)));
        }

        //刪除(DB)
        public async Task<JsonResult> Delete(string key)
        {
            return Json(await EditService().DeleteAsync(key));
        }



    }//class
}