using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using Microsoft.AspNetCore.Mvc;
using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class DOCSSignController : ApiCtrl
    {
        public ActionResult Read()
        {

            return View();
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private DOCSSignService EditService()
        {
            return new DOCSSignService(Ctrl);
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

