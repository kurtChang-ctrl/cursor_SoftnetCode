using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseWeb.Services;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Services;
using System.Data;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class FlowSignController : ApiCtrl
    {
        public ActionResult Read()
        {
            //var locale0 = _Xp.GetLocale0();
            //ViewBag.SignStatuses2 = await _XpCode.GetSignStatuses2Async(locale0);
            return View();
        }
        [HttpPost]
        public RedirectToActionResult SignRedirect(string keys)
        {
            return RedirectToAction("Read", "/DOCBuy", new { key = keys });
        }
        [HttpPost]
        public async Task<string> SignOK(string ip, string id, string status, string note)
        {

            string meg = "";
            string[] data = id.Split(',');
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow dr;
                for (int i = 0; i < data.Length; i++)
                {
                    meg += await _XgFlow.SignRowAsync(data[i], (status == "Y"), "", "", false);
                }
                //switch (data[i].Substring(0,4))
                //    {
                //        case "":
                //            break;
                //    }
                //    dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[APS_NeedData] where Id='{data[i]}' and State='2' order by NeedSimulationDate");
                //    if (dr != null)
                //    {
                //        //###???暫時  if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='6' where Id='{data[i]}'")) { run = true; }
                //    }
                //}
            }
            return "";
        }


        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(Ctrl, dt));
        }

        private FlowSignService EditService()
        {
            return new FlowSignService(Ctrl);
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

        //TODO: add your code
        //get file/image
        public async Task<FileResult> ViewFile(string table, string fid, string key, string ext)
        {
            return await _Xp.ViewLeaveAsync(fid, key, ext);
        }
    }
}
