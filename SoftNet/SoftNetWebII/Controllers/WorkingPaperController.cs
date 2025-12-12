using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;

using SoftNetWebII.Services;
using StackExchange.Redis;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.WebSockets;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{

    public class WorkingPaperController : ApiCtrl
    {
        public ActionResult Index()
        {
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                List<IdStrDto> mFNO = new List<IdStrDto>();
                mFNO.Add(new IdStrDto("", ""));
                DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        mFNO.Add(new IdStrDto(dr["MFNO"].ToString(), dr["MFName"].ToString()));
                    }
                }
                ViewBag.MFNO = mFNO;

                List<IdStrDto> workType = new List<IdStrDto>();
                workType.Add(new IdStrDto("1", "排程採購"));
                workType.Add(new IdStrDto("2", "排程委外"));
                workType.Add(new IdStrDto("3", "補存貨量"));
                workType.Add(new IdStrDto("4", "廠內生產"));
                ViewBag.WorkType = workType;


            }
            return View();
        }
        public ActionResult MailAutoAction(string id)
        {
            string re = "";
            if (id!=null && id!="")
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataRow tmp = null;
                    string sendTime = DateTime.Now.AddMinutes(15).ToString("MM/dd/yyyy HH:mm:ss.fff"); //DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff"); ;
                    foreach (string s in id.Split(';'))
                    {
                        if (s != "")
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where IsOK='1' and ServerId='{_Fun.Config.ServerId}' and Id='{s}'");
                            if (tmp != null)
                            {
                                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkingPaper] SET IsOK='2',SendTime='{sendTime}' where ServerId='{_Fun.Config.ServerId}' and Id='{s}'"))
                                { re = $"{re}<p>料號:{tmp["PartNO"].ToString()} 完成:{tmp["DOCNumberNO"].ToString()}單據發送</p>"; }
                                else { re = $"{re}<p>{s} 失敗</p>"; }
                            }
                            else { re = $"{re}<p>{s} 可能已過期了.</p>"; }
                        }
                    }
                }
            }
            else
            {
                re = "此封干涉作業已不存在.";
            }
            ViewBag.HtmiDisplay = re;
            return View();
        }
        [HttpPost]
        public string GOTORUNDOC(string keys) //正式發出單據 0=ipport,1.Id1,Id2....
        {
            string meg = "";
            string[] data = keys.Split(',');
            bool run = false;
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                List<string> neddIds = new List<string>();
                DataRow dr_M = null;
                DataTable dt_APS_Simulation = null;
                DataRow tmp;
                DataRow tmp2;

                string sql = "";
                for (int i = 1; i < data.Length; i++)
                {

                    //###???未完成

                }
                
            }
            meg = "此功能程式尚未寫好.";
            return meg;
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private PageTableWervice EditService()
        {
            return new PageTableWervice(Ctrl);
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
            string c_MFNO = "";
            string c_Price = "";
            JObject data = _Str.ToJson(json);
            var rows = data["_rows"] as JArray;
            if (rows[0]["MFNO"] != null)
            {
                c_MFNO = rows[0]["MFNO"].ToString().Trim();
            }
            if (rows[0]["Price"] != null)
            {
                c_Price = rows[0]["Price"].ToString().Trim();
            }

            ResultDto row = await EditService().UpdateAsync(key, _Str.ToJson(json));
            if (row.ErrorMsg == "")
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[APS_WorkingPaper] where Id='{key}'");
                    if (dr != null)
                    {
                        if ((c_MFNO!="" || c_Price!="") && dr["DOCNumberNO"].ToString() != "")
                        {
                            switch (dr["WorkType"].ToString())
                            {
                                case "1":
                                    //###???
                                    break;
                                case "2":
                                    {
                                        string wr = "";
                                        if (c_Price != "") { c_Price = $"Price='{c_Price}'"; }
                                        if (c_MFNO != "") { c_MFNO = $"MFNO='{c_MFNO}'"; }
                                        if (c_MFNO != "" || c_Price != "") 
                                        {
                                            wr = $"{(c_Price==""?"": c_Price)}";
                                            if (wr == "") { wr = $"{(c_MFNO == "" ? "" : c_MFNO)}"; }
                                            else { wr = $"{wr}{(c_MFNO == "" ? "" : $",{c_MFNO}")}"; }
                                        }
                                        if (wr != "")
                                        {
                                            wr = $"set {wr}";
                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC4Production] {wr} where DOCNumberNO='{dr["DOCNumberNO"].ToString()}'");
                                            db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkingPaper] {wr} where DOCNumberNO='{dr["DOCNumberNO"].ToString()}'");
                                        }
                                    }
                                    break;
                                case "3":
                                    //###???
                                    break;
                                case "4":
                                    //###???
                                    break;
                            }

                        }
                    }
                }
            }
            return Json(row);
        }

        public async Task<JsonResult> Delete(string key)
        {
            return Json(await EditService().DeleteAsync(key));
        }
    }
}
