using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class MFDataController : ApiCtrl
    {
        public ActionResult Read()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            List<IdStrDto> storeNO = new List<IdStrDto>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT StoreNO,StoreName FROM SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    storeNO.Add(new IdStrDto("", ""));
                    foreach (DataRow dr in dt.Rows)
                    {
                        storeNO.Add(new IdStrDto(dr["StoreNO"].ToString(), dr["StoreName"].ToString()));
                    }
                }
            }
            ViewBag.StoreNO = storeNO;

            return View();
        }


        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private MFDataService EditService()
        {
            return new MFDataService(Ctrl);
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
            string[] data = key.Split(';');
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT MFNO FROM SoftNetMainDB.[dbo].[MFData] where ServerId='{data[0]}' and MFNO='{data[1]}'");
                if (tmp == null)
                {
                    ResultDto row = new ResultDto();
                    row.ErrorMsg = "主欄位 不能被異動.";
                    return Json(row);
                }
            }
            return Json(await EditService().UpdateAsync(key, _Str.ToJson(json)));
        }

        public async Task<JsonResult> Delete(string key)
        {
			string[] data = key.Split(';');
			ResultDto row = new ResultDto();
			using (DBADO db = new DBADO("1", _Fun.Config.Db))
			{
				DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO FROM SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and MFNO='{data[1]}'");
				if (tmp != null) { row.ErrorMsg = "BOM表中有此廠商設定, 故無法刪除."; goto break_FUN; }
				tmp = db.DB_GetFirstDataByDataRow($"SELECT a.MFNO FROM SoftNetMainDB.[dbo].[DOC1Buy] as a join SoftNetMainDB.[dbo].[DOC1BuyII] as b on a.DOCNumberNO=b.DOCNumberNO and b.IsOK='0' where ServerId='{_Fun.Config.ServerId}' and MFNO='{data[1]}' and IsOK='0'");
				if (tmp != null) { row.ErrorMsg = "進貨單據有此料號, 故無法刪除."; goto break_FUN; }
				tmp = db.DB_GetFirstDataByDataRow($"SELECT a.MFNO FROM SoftNetMainDB.[dbo].[DOC4Production] as a join SoftNetMainDB.[dbo].[DOC4ProductionII] as b on a.DOCNumberNO=b.DOCNumberNO and b.IsOK='0' where ServerId='{_Fun.Config.ServerId}' MFNO PartNO='{data[1]}' and IsOK='0'");
				if (tmp != null) { row.ErrorMsg = "委外加工單據有此料號, 故無法刪除."; goto break_FUN; }

				//return Json(await EditService().DeleteAsync(key));
				row = await EditService().DeleteAsync(key);

			}
		break_FUN:

			return Json(row);
		}
    }
}

