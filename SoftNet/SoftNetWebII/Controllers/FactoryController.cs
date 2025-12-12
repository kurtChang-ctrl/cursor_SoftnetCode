using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using Microsoft.AspNetCore.Mvc;
using SoftNetWebII.Services;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class FactoryController : ApiCtrl
    {
        public ActionResult Read()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            return View();
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await new FactoryRead().GetPageAsync(dt, Ctrl));
        }

        private FactoryEdit EditService()
        {
            return new FactoryEdit(Ctrl);
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
