using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using Microsoft.AspNetCore.Mvc;
using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class DOCController : ApiCtrl
    {
        public ActionResult Read()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            List<IdStrDto> dOCTypeName = new List<IdStrDto>();
            dOCTypeName.Add(new IdStrDto("1", "進貨"));
            dOCTypeName.Add(new IdStrDto("2", "銷貨"));
            dOCTypeName.Add(new IdStrDto("3", "存貨_領料"));
            dOCTypeName.Add(new IdStrDto("4", "存貨_入料"));
            //dOCTypeName.Add(new IdStrDto("5", "存貨_調撥"));//###???目前系統設計不能有調撥,只能用領料,入料
            dOCTypeName.Add(new IdStrDto("6", "生產"));
            dOCTypeName.Add(new IdStrDto("7", "委外"));
            dOCTypeName.Add(new IdStrDto("8", "檢驗"));
            dOCTypeName.Add(new IdStrDto("9", "財務"));
            ViewBag.DOCType = dOCTypeName;

            return View();
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private DOCService EditService()
        {
            return new DOCService(Ctrl);
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
