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
    public class MaterialController : ApiCtrl
    {
        public ActionResult Read()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            List<IdStrDto> typeName = new List<IdStrDto>();
            typeName.Add(new IdStrDto("0", "實料"));
            typeName.Add(new IdStrDto("1", "虛料"));

            List<IdStrDto> className = new List<IdStrDto>();
            className.Add(new IdStrDto("1", "原物料"));
            className.Add(new IdStrDto("2", "採購件"));
            className.Add(new IdStrDto("3", "委外件"));
            className.Add(new IdStrDto("4", "製造半成品"));
            className.Add(new IdStrDto("5", "製造成品"));
            className.Add(new IdStrDto("6", "刀具"));
            className.Add(new IdStrDto("7", "工具製具"));

            List<IdStrDto> unitName = new List<IdStrDto>();
            unitName.Add(new IdStrDto("PCS", "PCS"));
            unitName.Add(new IdStrDto("個", "個"));
            unitName.Add(new IdStrDto("包", "包"));
            unitName.Add(new IdStrDto("箱", "箱"));
            unitName.Add(new IdStrDto("公克", "公克"));
            unitName.Add(new IdStrDto("公斤", "公斤"));
            unitName.Add(new IdStrDto("公噸", "公噸"));
            unitName.Add(new IdStrDto("公升", "公升"));

			List<IdStrDto> aPS_Default_MFNO = new List<IdStrDto>();
			List<IdStrDto> aPS_Default_StoreNO = new List<IdStrDto>();
			using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
				DataTable dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' order by MFNO");
				if (dt != null && dt.Rows.Count > 0)
				{
                    string name = "";
					aPS_Default_MFNO.Add(new IdStrDto("", ""));
					foreach (DataRow dr in dt.Rows)
                    {
                        if (dr["SName"].ToString() != "") { name = dr["SName"].ToString(); }
                        else { name = dr["MFNO"].ToString(); }
						aPS_Default_MFNO.Add(new IdStrDto(dr["MFNO"].ToString(), name));
					}
				}
				dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' order by StoreNO");
				if (dt != null && dt.Rows.Count > 0)
				{
					string name = "";
					aPS_Default_StoreNO.Add(new IdStrDto("", ""));
					foreach (DataRow dr in dt.Rows)
					{
						if (dr["StoreName"].ToString() != "") { name = dr["StoreName"].ToString(); }
						else { name = dr["StoreNO"].ToString(); }
						aPS_Default_StoreNO.Add(new IdStrDto(dr["StoreNO"].ToString(), name));
					}
				}
			}
            //@少儲位
            ViewBag.APS_Default_MFNO = aPS_Default_MFNO;
			ViewBag.APS_Default_StoreNO = aPS_Default_StoreNO;

			ViewBag.PartType = typeName;
            ViewBag.Unit = unitName;
            ViewBag.Class = className;
            return View();
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await new MaterialRead().GetPageAsync(dt, Ctrl));
        }

        private MaterialEdit EditService()
        {
            return new MaterialEdit(Ctrl);
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
            //ResultDto row = await EditService().CreateAsync(_Str.ToJson(json));
            //if (row.ErrorMsg == "")
            //{
            //    JObject json2 = _Str.ToJson(json);
            //    var rows = json2["_rows"] as JArray;
            //    if (rows != null && rows.Count > 0)
            //    {
            //        bool isrun = false;
            //        string partNO = rows[0]["PartNO"].ToString();
            //        string aPS_Default_StoreNO = "";
            //        string aPS_Default_StoreSpacesNO = "";
            //        if (rows[0]["APS_Default_StoreNO"] != null) { aPS_Default_StoreNO = rows[0]["APS_Default_StoreNO"].ToString(); isrun = true; }
            //        if (rows[0]["APS_Default_StoreSpacesNO"] != null) { aPS_Default_StoreSpacesNO = rows[0]["APS_Default_StoreSpacesNO"].ToString(); isrun = true; }
            //        if (isrun)
            //        {
            //            //EditDto editDto = EditService().GetDto();
            //            //partNO = editDto.PkeyFids[1];
            //            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            //            {
            //                DataRow tmp = null;
            //                if (aPS_Default_StoreSpacesNO != "")
            //                { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{aPS_Default_StoreNO}' and StoreSpacesNO='{aPS_Default_StoreSpacesNO}'"); }
            //                else
            //                { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{aPS_Default_StoreNO}'"); }
            //                if (tmp == null)
            //                {
            //                    db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{aPS_Default_StoreNO}','{aPS_Default_StoreNO}','{partNO}',0)");
            //                }
            //            }
            //        }
            //    }
            //}
            //return Json(row);
        }

        public async Task<JsonResult> Update(string key, string json)
        {
            string[] data = key.Split(';');
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO FROM SoftNetMainDB.[dbo].[Material] where ServerId='{data[0]}' and PartNO='{data[1]}'");
                if (tmp == null)
                {
                    ResultDto row2 = new ResultDto();
                    row2.ErrorMsg = "主欄位 不能被異動.";
                    return Json(row2);
                }
            }
            return Json(await EditService().UpdateAsync(key, _Str.ToJson(json)));
            //ResultDto row = await EditService().UpdateAsync(key, _Str.ToJson(json));
            //if (row.ErrorMsg == "")
            //{
            //    JObject json2 = _Str.ToJson(json);
            //    var rows = json2["_rows"] as JArray;
            //    if (rows != null && rows.Count > 0)
            //    {
            //        bool isrun = false;
            //        string partNO = key.Split(';')[1];
            //        string aPS_Default_StoreNO = "";
            //        string aPS_Default_StoreSpacesNO = "";
            //        if (rows[0]["APS_Default_StoreNO"] != null) { aPS_Default_StoreNO = rows[0]["APS_Default_StoreNO"].ToString(); isrun = true; }
            //        if (rows[0]["APS_Default_StoreSpacesNO"] != null) { aPS_Default_StoreSpacesNO = rows[0]["APS_Default_StoreSpacesNO"].ToString(); isrun = true; }
            //        if (isrun)
            //        {
            //            //EditDto editDto = EditService().GetDto();
            //            //partNO = editDto.PkeyFids[1];
            //            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            //            {
            //                DataRow tmp = null;
            //                if (aPS_Default_StoreSpacesNO != "")
            //                { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{aPS_Default_StoreNO}' and StoreSpacesNO='{aPS_Default_StoreSpacesNO}'"); }
            //                else
            //                { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{aPS_Default_StoreNO}'"); }
            //                if (tmp == null)
            //                {
            //                    db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{aPS_Default_StoreNO}','{aPS_Default_StoreNO}','{partNO}',0)");
            //                }
            //            }
            //        }
            //    }
            //}
            //return Json(row);
        }

        public async Task<JsonResult> Delete(string key)
        {
            string[] data = key.Split(';');
            ResultDto row = new ResultDto();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO FROM SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and PartNO='{data[1]}'");
                if (tmp != null) { row.ErrorMsg = "BOM表中有此料號, 故無法刪除."; goto break_FUN; }
                tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO FROM SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and PartNO='{data[1]}'");
                if (tmp != null) { row.ErrorMsg = "BOM表中有此料號, 故無法刪除."; goto break_FUN; }
                tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO FROM SoftNetMainDB.[dbo].[DOC1BuyII] where ServerId='{_Fun.Config.ServerId}' and PartNO='{data[1]}' and IsOK='0'");
                if (tmp != null) { row.ErrorMsg = "進貨單據有此料號, 故無法刪除."; goto break_FUN; }
                tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO FROM SoftNetMainDB.[dbo].[DOC2SalesII] where ServerId='{_Fun.Config.ServerId}' and PartNO='{data[1]}' and IsOK='0'");
                if (tmp != null) { row.ErrorMsg = "出貨單據有此料號, 故無法刪除."; goto break_FUN; }
                tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO FROM SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and PartNO='{data[1]}' and IsOK='0'");
                if (tmp != null) { row.ErrorMsg = "存貨單據有此料號, 故無法刪除."; goto break_FUN; }
                tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO FROM SoftNetMainDB.[dbo].[DOC4ProductionII] where ServerId='{_Fun.Config.ServerId}' and PartNO='{data[1]}' and IsOK='0'");
                if (tmp != null) { row.ErrorMsg = "委外加工單據有此料號, 故無法刪除."; goto break_FUN; }
                tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO FROM SoftNetMainDB.[dbo].[DOC5OUTII] where ServerId='{_Fun.Config.ServerId}' and PartNO='{data[1]}' and IsOK='0'");
                if (tmp != null) { row.ErrorMsg = "出貨單據有此料號, 故無法刪除."; goto break_FUN; }
                tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(QTY) as qty FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{data[1]}'");
                if (tmp != null && !tmp.IsNull("qty") && tmp["qty"].ToString().Trim()!="") { row.ErrorMsg = "倉儲尚有此料號存貨數量, 故無法刪除."; goto break_FUN; }
                tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO FROM SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and PartNO='{data[1]}' and IsOK='0'");
                if (tmp != null) { row.ErrorMsg = "生產排程尚有此料號數量, 故無法刪除."; goto break_FUN; }

                //return Json(await EditService().DeleteAsync(key));
                row = await EditService().DeleteAsync(key);
                if (row.ErrorMsg == "")
                {
                    string partNO = key.Split(';')[1];
                    //EditDto editDto = EditService().GetDto();
                    //partNO = editDto.PkeyFids[1];
                    DataTable dt = db.DB_GetData($"SELECT * SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}'");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            db.DB_SetData($"delete from SoftNetMainDB.[dbo].[TotalStockII_Knives] where ServerId='{_Fun.Config.ServerId}' and MId='{dr["Id"].ToString()}'");
                            db.DB_SetData($"delete from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and Id='{dr["Id"].ToString()}'");
                        }
                    }
                }
            }
        break_FUN:

            return Json(row);
        }
    }
}
