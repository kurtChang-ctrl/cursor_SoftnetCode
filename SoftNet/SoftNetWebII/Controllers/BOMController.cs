using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using SoftNetWebII.Services;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Data;
using BaseApi.Services;
using Base;

namespace SoftNetWebII.Controllers
{
    //[XgProgAuth]
    public class BOMController : ApiCtrl
    {
        public async Task<ActionResult> Read()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            //test
            //_Fun.Except("exception test");
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                List<IdStrDto> apply_PP_Name = new List<IdStrDto>();
                DataTable dt = db.DB_GetData($"SELECT PP_Name FROM SoftNetSYSDB.[dbo].[PP_ProductProcess] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        apply_PP_Name.Add(new IdStrDto(dr["PP_Name"].ToString(), dr["PP_Name"].ToString()));
                    }
                }
                ViewBag.Apply_PP_Name = apply_PP_Name;

                List<IdStrDto> className = new List<IdStrDto>();
                className.Add(new IdStrDto("1", "原物料"));
                className.Add(new IdStrDto("2", "採購件"));
                className.Add(new IdStrDto("3", "委外件"));
                className.Add(new IdStrDto("4", "製造半成品"));
                className.Add(new IdStrDto("5", "製造成品"));
                className.Add(new IdStrDto("6", "加工刀具"));
                className.Add(new IdStrDto("7", "工具製具"));
                ViewBag.Class = className;

                List<IdStrDto> isEnd = new List<IdStrDto>();
                isEnd.Add(new IdStrDto("0", "否"));
                isEnd.Add(new IdStrDto("1", "是"));
                ViewBag.IsEnd = isEnd;
            }
            return View();
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await new BOMRead().GetPageAsync(Ctrl, dt));
        }

        private BOMEdit EditService()
        {
            return new BOMEdit(Ctrl);
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



    }//class
}