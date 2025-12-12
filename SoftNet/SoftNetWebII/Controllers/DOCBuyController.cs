using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using SoftNetWebII.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Threading.Tasks;

using System.Data;
using SoftNetWebII.Models;
using Base.Enums;
using Base;

namespace SoftNetWebII.Controllers
{
    //[XgProgAuth]
    public class DOCBuyController : ApiCtrl
    {
        public ActionResult Read(string key)
        {
            List<IdStrDto> className = new List<IdStrDto>();
            className.Add(new IdStrDto("1", "原物料"));
            className.Add(new IdStrDto("2", "採購件"));
            className.Add(new IdStrDto("3", "委外件"));
            className.Add(new IdStrDto("4", "製造半成品"));
            className.Add(new IdStrDto("5", "製造成品"));
            className.Add(new IdStrDto("6", "加工刀具"));
            className.Add(new IdStrDto("7", "工具製具"));

            List<IdStrDto> mFNO = new List<IdStrDto>();
            List<IdStrDto> dOCNO = new List<IdStrDto>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT MFNO,MFName FROM [dbo].[MFData] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    int i = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        mFNO.Add(new IdStrDto(dr["MFNO"].ToString(), $"{dr["MFNO"].ToString()} {dr["MFName"].ToString()}"));
                    }
                }
                dt = db.DB_GetData($"SELECT DOCNO,DOCName FROM [dbo].[DOCRole] where DOCType='1' and ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    int i = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        dOCNO.Add(new IdStrDto(dr["DOCNO"].ToString(), $"{dr["DOCNO"].ToString()} {dr["DOCName"].ToString()}"));
                    }
                }
            }
            ViewBag.DOCNO = dOCNO;
            ViewBag.Class = className;
            ViewBag.MFNO = mFNO;
            if (key != null && key != "")
            { ViewBag.ViexID = key; }
            else
            { ViewBag.ViexID = ""; }
            return View();
        }

        public async Task<ActionResult> SingRead(string key)
        {
            return JsonToCnt(await EditService().GetUpdJsonAsync(key));
        }

        [HttpPost]
        public async Task<ContentResult> GetPartNOData(string deptId, string account)
        {
            if (!string.IsNullOrEmpty(account))
                account += '%';
            var sql = $@"select Model,PartNO,PartName,Specification,PartType,Unit from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and Class=iif(@Class is null, Class, @Class) and (@Model is null or Model like @Model) order by Class,Model,PartNO";
            //var rows = await _Db.GetJsonsAsync(sql, new List<object>() { "Model", account, "Class", deptId });
            DBADO db = new DBADO("1", _Fun.Config.Db);
            var rows = db.DB_Test(sql, new List<object>() { "Model", account, "Class" });
            db.Dispose();
            //string _s = rows.ToString();
            return Content(rows == null ? "" : rows.ToString(), ContentTypeEstr.Json);
        }



        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(Ctrl, dt));
        }

        private DOCBuyService EditService()
        {
            return new DOCBuyService(Ctrl);
        }

        //讀取要修改的資料(Get Updated Json)
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

        //新增(DB)
        public async Task<JsonResult> Create(string json)
        {
            //return Json(await EditService().CreateAsync(_Str.ToJson(json)));
            return Json(await EditService().OtherCrateAsync(_Str.ToJson(json)));
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
        public ActionResult GetSignRows(string id)
        {
            return PartialView(_Xp.SignRowsView, EditService().GetSignRows(id));
        }





    }//class
}
