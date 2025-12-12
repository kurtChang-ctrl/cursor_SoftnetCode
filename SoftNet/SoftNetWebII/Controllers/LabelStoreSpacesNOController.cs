using Base;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Office2010.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

using SoftNetWebII.Models;
using SoftNetWebII.Services;
using StackExchange.Redis;
using System;
using System.Collections.Generic;
using System.Data;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using ZXing.QrCode.Internal;

namespace SoftNetWebII.Controllers
{
    public class LabelStoreSpacesNOController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public LabelStoreSpacesNOController(SNWebSocketService websocket, SFC_Common sfc_Common)
        {
            if (_WebSocket == null)
            {
                _WebSocket = websocket;
            }
            if (_SFC_Common == null)
            {
                _SFC_Common = sfc_Common;
            }
        }
        public IActionResult UpdateTagResult(API_UpdateTagResult data)
        {
            return View();
        }
        public IActionResult EnterKey(API_EnterKeyResult data)
        {
            return View();
        }
        public IActionResult StroeFUN(string id)//id=倉庫編號;儲位
        {
            var br = _Fun.GetBaseUser();
            if (id == null || br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            if (br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                ViewBag.ERRMsg = "網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
                ViewBag.StoreNO = "";
                ViewBag.StoreSpacesNO = "";
                return View("ResuItTimeOUT");
            }
            string[] data=id.Split(';');
            string store = data[0];
            string storeSpacesNO = data[1];
            string storeName = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow tmp = db.DB_GetFirstDataByDataRow($@"select a.*,b.StoreName from SoftNetMainDB.[dbo].[StoreII] as a
                                join SoftNetMainDB.[dbo].[Store] as b on a.StoreNO=b.StoreNO
                                where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{store}' and a.StoreSpacesNO='{storeSpacesNO}'");
                if (tmp != null)
                {
                    storeName = tmp["StoreName"].ToString();
                }
            }
            ViewBag.StoreNO = store;
            ViewBag.StoreName = storeName;
            ViewBag.StoreSpacesNO = storeSpacesNO;
            ViewBag.Station = "";
            return View();
        }

        private bool OpenWindow_EStore(string url,string token, List<API_EStore_Open> req)
        {
            bool re = false;
            if (url == "" || req.Count <= 0) { return re; }
            HttpClient httpClient = new HttpClient();
            string json = JsonConvert.SerializeObject(req);
            var jdata = new StringContent(json, Encoding.UTF8, "application/json");
            httpClient.DefaultRequestHeaders.Add("machine-token", token);
            var responseTask = httpClient.PostAsync(url, jdata);
            responseTask.Wait();
            var result = responseTask.Result;
            var readTask = result.Content.ReadAsStringAsync();
            readTask.Wait();

            if (result.IsSuccessStatusCode)
            {
                //json = readTask.Result;
                re = true;
            }
            else
            {
                //json = readTask.Result;
                re = false;
            }
            readTask.Dispose();
            responseTask.Dispose();
            httpClient.Dispose();

            return re;
        }
        private void GetEStoreURLandToken(string storeNO, ref string url,ref string token)
        {
            #region 大正貴訊號處理 開門
            if (_Fun.Config.Default_EStore_ControlURL != "")
            {
                string[] type_Data = null;
                for (int i = 0; i < _Fun.Config.Default_EStore_ControlURL.Split(';').Length; i++)
                {
                    string s = _Fun.Config.Default_EStore_ControlURL.Split(';')[i].Trim();
                    if (s == "") { continue; }
                    type_Data = s.Split(',');
                    if (type_Data.Length == 2 && type_Data[0] == storeNO)
                    {
                        url = type_Data[1];
                        token = _Fun.Config.Default_EStore_MachineToken.Split(';')[i].Trim();
                        break;
                    }
                }
            }
            else { return; }
            #endregion
        }
        private string OpenWindow_EStoreII_for_newList(DBADO db, string api_URL,string api_Token, string[] data, bool is_DOC)
        {
            string re = "";
            if (api_URL != "")
            {
                string docID = "";
                string docDOCNumberNO = "";
                bool isRUN = false;
                int out_qty = 0;
                DataRow dr_DOC3stockII = null;
                List<API_EStore_Open> req = new List<API_EStore_Open>();
                string[] zxy = null;
                for (int i = 0; i < data.Length; i++)
                {
                    if (data[i].Trim() == "") { continue; }
                    docID = data[i].Split(",")[0];
                    if (is_DOC)
                    {
                        docDOCNumberNO = data[i].Split(",")[1];
                        dr_DOC3stockII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where Id='{docID}' and DOCNumberNO='{docDOCNumberNO}' and IsOK='0'");
                        if (dr_DOC3stockII != null)
                        {
                            try
                            {
                                if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                                { zxy = dr_DOC3stockII["OUT_StoreSpacesNO"].ToString().Split('.'); }
                                else { zxy = dr_DOC3stockII["IN_StoreSpacesNO"].ToString().Split('.'); }

                                if (zxy.Length == 3)
                                {
                                    req.Add(new API_EStore_Open(int.Parse(zxy[0]), int.Parse(zxy[1]), int.Parse(zxy[2])));
                                }
                                else { re = $"{re}<br />料號:{dr_DOC3stockII["PartNO"].ToString()} 儲位:{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()} 定義有錯, 電子櫃無法開門."; }
                            }
                            catch { re = $"{re}<br />料號:{dr_DOC3stockII["PartNO"].ToString()} 儲位:{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()} 定義有錯, 電子櫃無法開門."; }
                        }
                    }
                    else
                    {
                        dr_DOC3stockII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and Id='{data[i].Trim()}'");
                        if (dr_DOC3stockII != null)
                        {
                            try
                            {
                                zxy = dr_DOC3stockII["StoreSpacesNO"].ToString().Split('.');
                                if (zxy.Length == 3)
                                {
                                    req.Add(new API_EStore_Open(int.Parse(zxy[0]), int.Parse(zxy[1]), int.Parse(zxy[2])));
                                }
                                else { re = $"{re}<br />料號:{dr_DOC3stockII["PartNO"].ToString()} 儲位:{dr_DOC3stockII["StoreSpacesNO"].ToString()} 定義有錯, 電子櫃無法開門."; }
                            }
                            catch { re = $"{re}<br />料號:{dr_DOC3stockII["PartNO"].ToString()} 儲位:{dr_DOC3stockII["StoreSpacesNO"].ToString()} 定義有錯, 電子櫃無法開門."; }
                        }
                    }
                }
                if (!OpenWindow_EStore(api_URL, api_Token, req))
                { re = $"{re}<br />電子櫃通訊有問題, 電子櫃無法開門."; }
            }
            else { re = $"電子櫃設定有問題,無定義電子櫃的URL參數, 電子櫃無法開門."; }
            return re;
        }
        public ActionResult SetAction_Check(StoreList keys)
        {
            var br = _Fun.GetBaseUser();
            if (keys == null || keys.StoreNO == null || keys.StoreNO == "" || br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                ViewBag.ERRMsg = "網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
                ViewBag.StoreNO = "";
                ViewBag.StoreSpacesNO = "";
                return View("ResuItTimeOUT");
            }
            keys.ERRMsg = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string api_URL = "";
                DataRow dr_Store = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}'");
                if (dr_Store != null)
                {
                    switch (keys.ActionType)
                    {
                        case "A"://針對倉庫領取
                            {
                                if (keys.Station == "" || keys.Select_ID == null || keys.Select_ID.Trim() == "")
                                {
                                    keys.ERRMsg = $"{keys.ERRMsg}<br />沒有選擇單據項目, 請重新選擇.";
                                }
                                else
                                {
                                    string[] data = keys.Select_ID.Split(';');
                                    string[] data_QTY = keys.Select_ID_QTY.Split(';');
                                    if (data.Length != data_QTY.Length)
                                    {
                                        keys.ERRMsg = $"{keys.ERRMsg}<br />勾選項目與數量不符合, 請重新選擇 或連繫系統管理員.";
                                    }
                                    else
                                    {
                                        DataRow dr_TotalStock = null;
                                        DataRow tmp_dr = null;
                                        int out_qty = 0;
                                        bool isRUN = false;
                                        string docNumberNO = "";
                                        string report = "";
                                        int sQTY = 0;

                                        for (int i = 0; i < data.Length; i++)
                                        {
                                            if (data[i].Trim() == "") { continue; }
                                            dr_TotalStock = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and Id='{data[i].Trim()}'");
                                            isRUN = int.TryParse(data_QTY[i], out out_qty);
                                            if (dr_TotalStock == null || !isRUN)
                                            {
                                                ViewBag.ERRMsg = $"{ViewBag.ERRMsg}<br />{data[i].Trim()}項目, 數量={data_QTY[i]} 沒成功.";
                                            }
                                            else
                                            {
                                                sQTY = int.Parse(dr_TotalStock["QTY"].ToString());
                                                if (sQTY == out_qty) { continue; }
                                                #region 寫入庫存
                                                tmp_dr = db.DB_GetFirstDataByDataRow($"select PartNO,'' as NeedId,'' as SimulationId from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_TotalStock["PartNO"].ToString()}'");
                                                if (tmp_dr != null)
                                                {
                                                    if (sQTY > out_qty)
                                                    {
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={(sQTY - out_qty).ToString()} where ServerId='{_Fun.Config.ServerId}' and Id='{data[i].Trim()}'");
                                                        //###???暫時寫死AC02
                                                        _SFC_Common.Create_DOC3stock(db, tmp_dr, dr_TotalStock["StoreNO"].ToString(), dr_TotalStock["StoreSpacesNO"].ToString(), "", "", "AC02", Math.Abs(out_qty), "", "", $"存貨盤損", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "SetAction_Check;LabelStoreSpacesNOController", ref docNumberNO, br.UserNO.Trim(), true, true);
                                                        report = $"{report}<br />單號:{docNumberNO}&nbsp;料號:{dr_TotalStock["PartNO"].ToString()}&nbsp;&nbsp;修正為:{data_QTY[i]}數量&nbsp;盤損:{(sQTY - out_qty).ToString()}";
                                                    }
                                                    else
                                                    {
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={(out_qty - sQTY).ToString()} where ServerId='{_Fun.Config.ServerId}' and Id='{data[i].Trim()}'");
                                                        //###???暫時寫死BC02
                                                        _SFC_Common.Create_DOC3stock(db, tmp_dr, "", "", dr_TotalStock["StoreNO"].ToString(), dr_TotalStock["StoreSpacesNO"].ToString(), "BC02", out_qty, "", "", $"存貨盤盈", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "SetAction_Check;LabelStoreSpacesNOController", ref docNumberNO, br.UserNO.Trim(), true, true);
                                                        report = $"{report}<br />單號:{docNumberNO}&nbsp;料號:{dr_TotalStock["PartNO"].ToString()}&nbsp;&nbsp;修正為:{data_QTY[i]}數量&nbsp;盤盈:{(out_qty - sQTY).ToString()}";
                                                    }
                                                }
                                                #endregion
                                            }
                                        }

                                        ViewBag.Report = report;
                                        ViewBag.StoreNO = keys.StoreNO;
                                        ViewBag.StoreSpacesNO = keys.StoreSpacesNO;
                                        return View("ResuItTimeOUT");
                                    }

                                }
                            }
                            break;
                    }


                    DataTable dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and ((OUT_StoreNO='{keys.StoreNO}' and OUT_StoreSpacesNO='{keys.StoreSpacesNO}' and SUBSTRING(DOCNumberNO,1,4)='AC02') or (IN_StoreNO='{keys.StoreNO}' and IN_StoreSpacesNO='{keys.StoreSpacesNO}' and SUBSTRING(DOCNumberNO,1,4)='BC02')) order by DOCNumberNO,OUT_StoreSpacesNO desc,PartNO");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.SIDList = dt;
                    }
                    dt = db.DB_GetData($@"SELECT a.*  FROM SoftNetMainDB.[dbo].[TotalStock] as a
                    join SoftNetMainDB.[dbo].[Material] as b on b.Class!='6' and b.Class!='7' and a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                    where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{keys.StoreNO}' and a.StoreSpacesNO='{keys.StoreSpacesNO}' order by PartNO,StoreSpacesNO desc,QTY desc");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.TotalStock_List = dt;
                    }
                }
                else
                {
                    ViewBag.ERRMsg = "網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
                    ViewBag.StoreNO = "";
                    ViewBag.StoreSpacesNO = "";
                    return View("ResuItTimeOUT");
                }
            }
            keys.Select_ID = "";
            keys.Select_ID_QTY = "";
            ViewBag.StoreList = keys;
            return View();
        }

        public ActionResult SetAction_MIN(StoreList keys)
        {
            var br = _Fun.GetBaseUser();
            if (keys == null || keys.StoreNO == null || keys.StoreNO == "" || br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                ViewBag.ERRMsg = "網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
                ViewBag.StoreNO = "";
                ViewBag.StoreSpacesNO = "";
                return View("ResuItTimeOUT");
            }
            keys.ERRMsg = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string api_URL = "";
                string api_Token = "";
                DataRow dr_Store = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}'");
                if (dr_Store != null)
                {
                    #region 大正貴訊號處理 開門
                    if (dr_Store["Store_Type"].ToString() == "A")
                    {
                        GetEStoreURLandToken(keys.StoreNO, ref api_URL, ref api_Token);
                        if (api_URL == "" || api_Token == "") { keys.ERRMsg = "此倉為電子控制倉, 但無API參數設定."; }
                    }
                    #endregion

                    switch (keys.ActionType)
                    {
                        case "A"://針對單據領取
                            {
                                if (keys.Station == "" || keys.Select_ID == null || keys.Select_ID.Trim() == "")
                                {
                                    keys.ERRMsg = $"{keys.ERRMsg}<br />沒有選擇單據項目, 請重新選擇.";
                                }
                                else
                                {
                                    string[] data = keys.Select_ID.Split(';');
                                    string[] data_QTY = keys.Select_ID_QTY.Split(';');
                                    if (data.Length != data_QTY.Length)
                                    {
                                        keys.ERRMsg = $"{keys.ERRMsg}<br />勾選項目與數量不符合, 請重新選擇 或連繫系統管理員.";
                                    }
                                    else
                                    {
                                        string docID = "";
                                        string docDOCNumberNO = "";
                                        bool isRUN = false;
                                        int out_qty = 0;
                                        DataRow dr_DOC3stockII = null;

                                        DataRow dr_tag = null;
                                        List<DataRow> doc_Id_List = new List<DataRow>();
                                        #region  開門
                                        if (api_URL != "")
                                        { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, true); }
                                        #endregion 

                                        for (int i = 0; i < data.Length; i++)
                                        {
                                            if (data[i].Trim() == "" || data_QTY[i] == "") { continue; }
                                            if (data[i].Trim() != "" && data[i].Split(",").Length > 1)
                                            {
                                                docID = data[i].Split(",")[0];
                                                docDOCNumberNO = data[i].Split(",")[1];
                                                int docQTY = 0;
                                                isRUN = int.TryParse(data_QTY[i], out out_qty);
                                                if (isRUN && out_qty != 0 && docID != "" && docDOCNumberNO != "")
                                                {
                                                    #region 計算單據CT,平均,有效,寫入庫存, 寫IsOK='1'
                                                    string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
                                                    dr_DOC3stockII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where Id='{docID}' and DOCNumberNO='{docDOCNumberNO}' and IsOK='0'");
                                                    if (dr_DOC3stockII != null)
                                                    {
                                                        docQTY = int.Parse(dr_DOC3stockII["QTY"].ToString());
                                                        if (out_qty != docQTY)
                                                        {
                                                            if (out_qty > docQTY)
                                                            {
                                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set QTY={out_qty} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}'");
                                                            }
                                                            else
                                                            {
                                                                //拆單處置
                                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set QTY={out_qty} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}'");
                                                                db.DB_SetData($@"INSERT INTO [dbo].[DOC3stockII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],ArrivalDate) VALUES 
                                                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{dr_DOC3stockII["DOCNumberNO"].ToString()}','{dr_DOC3stockII["PartNO"].ToString()}',{dr_DOC3stockII["Price"].ToString()},'{dr_DOC3stockII["Unit"].ToString()}',{(docQTY - out_qty).ToString()}
                                                                                                ,'{dr_DOC3stockII["Remark"].ToString()}','{dr_DOC3stockII["SimulationId"].ToString()}','{dr_DOC3stockII["IsOK"].ToString()}','{dr_DOC3stockII["IN_StoreNO"].ToString()}','{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}','{dr_DOC3stockII["OUT_StoreNO"].ToString()}','{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}','{Convert.ToDateTime(dr_DOC3stockII["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                            }
                                                        }
                                                        dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'");
                                                        if (dr_tag != null && dr_tag["Config_macID"].ToString() != "")
                                                        {
                                                            #region 儲位亮燈
                                                            doc_Id_List.Add(dr_DOC3stockII);
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            #region 寫入庫存
                                                            //###???調撥會有問題

                                                            if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                                                            { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={out_qty} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["IN_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}'"); }
                                                            else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={out_qty} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'"); }
                                                            #endregion

                                                            int typeTotalTime = 0;
                                                            string writeSQL = "";
                                                            if (!dr_DOC3stockII.IsNull("StartTime")) { typeTotalTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC3stockII["StartTime"].ToString()), DateTime.Now); }
                                                            else { writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime}{writeSQL} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}' and IsOK='0'");
                                                            string partNO = dr_DOC3stockII["PartNO"].ToString();
                                                            string pp_Name = "";
                                                            string E_stationNO = "";
                                                            if (dr_DOC3stockII["SimulationId"].ToString() != "")
                                                            {
                                                                DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC3stockII["SimulationId"].ToString()}'");
                                                                pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                                                                if (!dr_tmp.IsNull("Source_StationNO") && dr_tmp["Source_StationNO"].ToString() != "")
                                                                { E_stationNO = dr_tmp["Source_StationNO"].ToString(); }
                                                                else { E_stationNO = dr_tmp["Apply_StationNO"].ToString(); }
                                                            }
                                                            DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} CT from SoftNetMainDB.[dbo].[DOC3stockII] where SUBSTRING(Id,2,2)='{_Fun.Config.ServerId}' and SUBSTRING(DOCNumberNO,1,4)='{dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4)}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and CT>0");
                                                            List<double> allCT = new List<double>();
                                                            if (dt_Efficient != null && dt_Efficient.Rows.Count > 0)
                                                            {
                                                                for (int i2 = 0; i2 < dt_Efficient.Rows.Count; i2++)
                                                                {
                                                                    foreach (DataRow dr2 in dt_Efficient.Rows)
                                                                    {
                                                                        allCT.Add(double.Parse(dr2["CT"].ToString()));
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (typeTotalTime != 0)
                                                                { allCT.Add(typeTotalTime); }
                                                            }
                                                            if (allCT.Count > 0)
                                                            {
                                                                _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, pp_Name, pp_Name, "0", partNO, partNO, dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4));
                                                            }
                                                        }
                                                        ViewBag.Report = "已完成單據領取作業.";
                                                    }
                                                    #endregion
                                                }

                                            }
                                        }
                                        if (doc_Id_List.Count > 0)
                                        {
                                            #region 儲位亮燈
                                            var remacIDs = "";
                                            var json3 = "";
                                            string storeNO = "";
                                            string storeSpacesN = "";
                                            using (HttpClient httpClient = new HttpClient())
                                            {
                                                DateTime now = DateTime.Now;
                                                try
                                                {
                                                    foreach (DataRow d in doc_Id_List)
                                                    {

                                                        if (d["IN_StoreNO"].ToString() != "" && d["IN_StoreSpacesNO"].ToString() != "") { storeNO = d["IN_StoreNO"].ToString(); storeSpacesN = d["IN_StoreSpacesNO"].ToString(); }
                                                        else if (d["OUT_StoreNO"].ToString() != "" && d["OUT_StoreSpacesNO"].ToString() != "") { storeNO = d["OUT_StoreNO"].ToString(); storeSpacesN = d["OUT_StoreSpacesNO"].ToString(); }

                                                        DataRow tmp_SNO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[StoreII] where StoreNO='{storeNO}' and StoreSpacesNO='{storeSpacesN}'");
                                                        if (tmp_SNO != null)
                                                        {
                                                            if (remacIDs == "") { remacIDs = $"{tmp_SNO["Config_macID"].ToString().Trim()}"; }
                                                            else { remacIDs = $"{remacIDs},{tmp_SNO["Config_macID"].ToString().Trim()}"; }
                                                            if (_Fun.Is_Tag_Connect)
                                                            {
                                                                if (json3 == "")
                                                                { json3 = "[{" + $"\"mac\":\"{tmp_SNO["Config_macID"].ToString().Trim()}\",\"outtime\":0,\"ledrgb\":\"ff0000\",\"ledmode\":0,\"buzzer\":0,\"lednum\":255,\"reserve\":\"reserve\"" + "}"; }
                                                                else
                                                                { json3 = json3 + ",{" + $"\"mac\":\"{tmp_SNO["Config_macID"].ToString().Trim()}\",\"outtime\":0,\"ledrgb\":\"ff0000\",\"ledmode\":0,\"buzzer\":0,\"lednum\":255,\"reserve\":\"reserve\"" + "}"; }
                                                            }
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set SimulationId='{d["SimulationId"].ToString()}',DOCNumberNO='{d["DOCNumberNO"].ToString()}',Id='{d["Id"].ToString()}',Ledrgb='ff0000',Ledstate=0,IsUpdate='1' where ServerId='{_Fun.Config.ServerId}' and macID='{tmp_SNO["Config_macID"].ToString().Trim()}'");
                                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType]) VALUES ('{_Str.NewId('L')}','{tmp_SNO["Config_macID"].ToString()}','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈')");
                                                        }
                                                    }
                                                    if (json3 != "") { json3 += "]"; }
                                                    string url = $"http://{_Fun.Config.ElectronicTagsURL}/wms/associate/lightTagsLed";
                                                    var content = new StringContent(json3, Encoding.UTF8, "application/json");
                                                    HttpResponseMessage response = httpClient.PostAsync(url, content).Result;
                                                    if (!response.IsSuccessStatusCode)
                                                    {
                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType],[JSON]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈,發送Fail','{remacIDs}')");
                                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"儲位亮燈 傳送電子訊號失敗,請通知管理者", false);    //false here, not mailRoot, or endless roop !!
                                                    }
                                                    else
                                                    {
                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType],[JSON]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈,發送OK','{remacIDs}')");
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"LabelStroeController.cs 儲位亮燈 Exception: {ex.Message} {ex.StackTrace}", true);
                                                }
                                            }
                                            #endregion
                                        }
                                        ViewBag.StoreNO = keys.StoreNO;
                                        ViewBag.StoreSpacesNO = keys.StoreSpacesNO;
                                        return View("ResuItTimeOUT");
                                    }
                                }
                            }
                            break;
                        case "B"://針對倉庫領取
                            {

                                if (keys.Station == "" || keys.Select_ID == null || keys.Select_ID.Trim() == "")
                                {
                                    keys.ERRMsg = $"{keys.ERRMsg}<br />沒有選擇單據項目, 請重新選擇.";
                                }
                                else
                                {
                                    string[] data = keys.Select_ID.Split(';');
                                    string[] data_QTY = keys.Select_ID_QTY.Split(';');
                                    if (data.Length != data_QTY.Length)
                                    {
                                        keys.ERRMsg = $"{keys.ERRMsg}<br />勾選項目與數量不符合, 請重新選擇 或連繫系統管理員.";
                                    }
                                    else
                                    {
                                        DataRow dr_TotalStock = null;
                                        DataRow tmp_dr = null;
                                        int out_qty = 0;
                                        bool isRUN = false;
                                        string docNumberNO = "";
                                        string report = "";

                                        DataRow dr_DOC3stockII = null;
                                        DataRow dr_tag = null;
                                        List<DataRow> doc_Id_List = new List<DataRow>();
                                        #region  開門
                                        if (api_URL != "")
                                        { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, false); }
                                        #endregion 


                                        for (int i = 0; i < data.Length; i++)
                                        {
                                            if (data[i].Trim() == "") { continue; }
                                            dr_TotalStock = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and Id='{data[i].Trim()}'");
                                            isRUN = int.TryParse(data_QTY[i], out out_qty);
                                            if (dr_TotalStock == null || !isRUN)
                                            {
                                                ViewBag.ERRMsg = $"{ViewBag.ERRMsg}<br />{data[i].Trim()}項目, 數量={data_QTY[i]} 沒成功.";
                                            }
                                            else
                                            {
                                                //#region 寫入庫存
                                                //tmp_dr = db.DB_GetFirstDataByDataRow($"select PartNO,'' as NeedId,'' as SimulationId from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_TotalStock["PartNO"].ToString()}'");
                                                //if (tmp_dr != null)
                                                //{
                                                //    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={out_qty.ToString()} where ServerId='{_Fun.Config.ServerId}' and Id='{data[i].Trim()}'");
                                                //    //###???暫時寫死AC09
                                                //    _WebSocket.Create_DOC3stock(db, tmp_dr, dr_TotalStock["StoreNO"].ToString(), dr_TotalStock["StoreSpacesNO"].ToString(), "", "", "AC09", out_qty, "", "", $"人為領料", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "SetAction_MIN;LabelStroeController", ref docNumberNO, br.UserNO.Trim(), true, true);
                                                //    if (report == "")
                                                //    { report = $"系統自動產生 單號:{docNumberNO}&nbsp;&nbsp;明細如下:"; }
                                                //    report = $"{report}<br />料號:{dr_TotalStock["PartNO"].ToString()}&nbsp;&nbsp;領取:{data_QTY[i]}數量";
                                                //}
                                                //#endregion
                                                #region 寫入庫存
                                                tmp_dr = db.DB_GetFirstDataByDataRow($"select PartNO,'' as NeedId,'' as SimulationId from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_TotalStock["PartNO"].ToString()}'");
                                                if (tmp_dr != null)
                                                {
                                                    dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_TotalStock["StoreNO"].ToString()}' and StoreSpacesNO='{dr_TotalStock["StoreSpacesNO"].ToString()}'");
                                                    if (dr_tag != null && dr_tag["Config_macID"].ToString() != "")
                                                    {
                                                        string docID = "";
                                                        _SFC_Common.Create_DOC3stock(db, tmp_dr, dr_TotalStock["StoreNO"].ToString(), dr_TotalStock["StoreSpacesNO"].ToString(), "", "", "AC09", out_qty, "", ref docID, $"人為領料", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "SetAction_MIN;LabelStroeController", ref docNumberNO, br.UserNO.Trim(), true, false);
                                                        dr_DOC3stockII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where Id='{docID}' and DOCNumberNO='{docNumberNO}' and IsOK='0'");
                                                        if (dr_DOC3stockII != null)
                                                        {
                                                            #region 儲位亮燈
                                                            doc_Id_List.Add(dr_DOC3stockII);
                                                            #endregion
                                                        }
                                                    }
                                                    else
                                                    {
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={out_qty.ToString()} where ServerId='{_Fun.Config.ServerId}' and Id='{data[i].Trim()}'");
                                                        //###???暫時寫死AC09
                                                        _SFC_Common.Create_DOC3stock(db, tmp_dr, dr_TotalStock["StoreNO"].ToString(), dr_TotalStock["StoreSpacesNO"].ToString(), "", "", "AC09", out_qty, "", "", $"人為領料", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "SetAction_MIN;LabelStroeController", ref docNumberNO, br.UserNO.Trim(), true, true);
                                                    }

                                                    if (report == "")
                                                    { report = $"系統自動產生 單號:{docNumberNO}&nbsp;&nbsp;明細如下:"; }
                                                    report = $"{report}<br />料號:{dr_TotalStock["PartNO"].ToString()}&nbsp;&nbsp;領取:{data_QTY[i]}數量";
                                                }
                                                #endregion
                                            }
                                        }
                                        if (doc_Id_List.Count > 0)
                                        {
                                            #region 儲位亮燈
                                            var remacIDs = "";
                                            var json3 = "";
                                            string storeNO = "";
                                            string storeSpacesN = "";
                                            using (HttpClient httpClient = new HttpClient())
                                            {
                                                DateTime now = DateTime.Now;
                                                try
                                                {
                                                    foreach (DataRow d in doc_Id_List)
                                                    {

                                                        if (d["IN_StoreNO"].ToString() != "" && d["IN_StoreSpacesNO"].ToString() != "") { storeNO = d["IN_StoreNO"].ToString(); storeSpacesN = d["IN_StoreSpacesNO"].ToString(); }
                                                        else if (d["OUT_StoreNO"].ToString() != "" && d["OUT_StoreSpacesNO"].ToString() != "") { storeNO = d["OUT_StoreNO"].ToString(); storeSpacesN = d["OUT_StoreSpacesNO"].ToString(); }

                                                        DataRow tmp_SNO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[StoreII] where StoreNO='{storeNO}' and StoreSpacesNO='{storeSpacesN}'");
                                                        if (tmp_SNO != null)
                                                        {
                                                            if (remacIDs == "") { remacIDs = $"{tmp_SNO["Config_macID"].ToString().Trim()}"; }
                                                            else { remacIDs = $"{remacIDs},{tmp_SNO["Config_macID"].ToString().Trim()}"; }
                                                            if (_Fun.Is_Tag_Connect)
                                                            {
                                                                if (json3 == "")
                                                                { json3 = "[{" + $"\"mac\":\"{tmp_SNO["Config_macID"].ToString().Trim()}\",\"outtime\":0,\"ledrgb\":\"ff0000\",\"ledmode\":0,\"buzzer\":0,\"lednum\":255,\"reserve\":\"reserve\"" + "}"; }
                                                                else
                                                                { json3 = json3 + ",{" + $"\"mac\":\"{tmp_SNO["Config_macID"].ToString().Trim()}\",\"outtime\":0,\"ledrgb\":\"ff0000\",\"ledmode\":0,\"buzzer\":0,\"lednum\":255,\"reserve\":\"reserve\"" + "}"; }
                                                            }
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set SimulationId='{d["SimulationId"].ToString()}',DOCNumberNO='{d["DOCNumberNO"].ToString()}',Id='{d["Id"].ToString()}',Ledrgb='ff0000',Ledstate=0,IsUpdate='1' where ServerId='{_Fun.Config.ServerId}' and macID='{tmp_SNO["Config_macID"].ToString().Trim()}'");
                                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType]) VALUES ('{_Str.NewId('L')}','{tmp_SNO["Config_macID"].ToString()}','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈')");
                                                        }
                                                    }
                                                    if (json3 != "") { json3 += "]"; }
                                                    string url = $"http://{_Fun.Config.ElectronicTagsURL}/wms/associate/lightTagsLed";
                                                    var content = new StringContent(json3, Encoding.UTF8, "application/json");
                                                    HttpResponseMessage response = httpClient.PostAsync(url, content).Result;
                                                    if (!response.IsSuccessStatusCode)
                                                    {
                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType],[JSON]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈,發送Fail','{remacIDs}')");
                                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"儲位亮燈 傳送電子訊號失敗,請通知管理者", false);    //false here, not mailRoot, or endless roop !!
                                                    }
                                                    else
                                                    {
                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType],[JSON]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈,發送OK','{remacIDs}')");
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"LabelStroeController.cs 儲位亮燈 Exception: {ex.Message} {ex.StackTrace}", true);
                                                }
                                            }
                                            #endregion
                                        }
                                        ViewBag.Report = report;
                                        ViewBag.StoreNO = keys.StoreNO;
                                        ViewBag.StoreSpacesNO = keys.StoreSpacesNO;
                                        return View("ResuItTimeOUT");
                                    }

                                }
                            }
                            break;
                    }
                    if (keys.ActionType == "A" || keys.ActionType == "B")
                    {
                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and Keys='領料_{br.UserNO.Trim()}' and StoreNO='{keys.StoreNO}'");
                    }

                    DataTable dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and OUT_StoreNO='{keys.StoreNO}' and OUT_StoreSpacesNO='{keys.StoreSpacesNO}' and IsOK='0' order by DOCNumberNO,OUT_StoreSpacesNO desc,PartNO");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.SIDList = dt;
                    }
                    dt = db.DB_GetData($@"SELECT a.*  FROM SoftNetMainDB.[dbo].[TotalStock] as a
                    join SoftNetMainDB.[dbo].[Material] as b on b.Class!='6' and b.Class!='7' and a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                    where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{keys.StoreNO}' and a.StoreSpacesNO='{keys.StoreSpacesNO}' order by PartNO,StoreSpacesNO desc,QTY desc");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.TotalStock_List = dt;
                    }
                }
                else
                {
                    ViewBag.ERRMsg = "網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
                    ViewBag.StoreNO = "";
                    ViewBag.StoreSpacesNO = "";
                    return View("ResuItTimeOUT");
                }
            }
            keys.Select_ID = "";
            keys.Select_ID_QTY = "";
            ViewBag.StoreList = keys;
            return View();
        }
        public ActionResult SetAction_MOUT(StoreList keys)
        {
            var br = _Fun.GetBaseUser();
            if (keys == null || keys.StoreNO == null || keys.StoreNO == "" || br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                ViewBag.ERRMsg = "網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
                ViewBag.StoreNO = "";
                ViewBag.StoreSpacesNO = "";
                return View("ResuItTimeOUT");
            }
            keys.ERRMsg = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string api_URL = "";
                string api_Token = "";
                DataRow dr_Store = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}'");
                if (dr_Store != null)
                {
                    #region 大正貴訊號處理 開門
                    if (dr_Store["Store_Type"].ToString() == "A")
                    {
                        GetEStoreURLandToken(keys.StoreNO, ref api_URL, ref api_Token);
                        if (api_URL == "" || api_Token == "") { keys.ERRMsg = "此倉為電子控制倉, 但無API參數設定."; }
                    }
                    #endregion

                    string sql = "";
                    DataRow dr_tmp = null;
                    switch (keys.ActionType)
                    {
                        case "A"://針對單據入庫
                            {
                                if (keys.Station == "" || keys.Select_ID == null || keys.Select_ID.Trim() == "")
                                {
                                    keys.ERRMsg = $"{keys.ERRMsg}<br />沒有選擇單據項目, 請重新選擇.";
                                }
                                else
                                {
                                    string[] data = keys.Select_ID.Split(';');
                                    string[] data_QTY = keys.Select_ID_QTY.Split(';');
                                    if (data.Length != data_QTY.Length)
                                    {
                                        keys.ERRMsg = $"{keys.ERRMsg}<br />勾選項目與數量不符合, 請重新選擇 或連繫系統管理員.";
                                    }
                                    else
                                    {
                                        string docID = "";
                                        string docDOCNumberNO = "";
                                        bool isRUN = false;
                                        int out_qty = 0;
                                        DataRow dr_DOCRole = null;
                                        Dictionary<string, List<DataRow>> doc_Id_List = new Dictionary<string, List<DataRow>>();
                                        #region  開門
                                        if (api_URL != "")
                                        { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, true); }
                                        #endregion 

                                        for (int i = 0; i < data.Length; i++)
                                        {
                                            if (data[i].Trim() == "" || data_QTY[i] == "") { continue; }
                                            if (data[i].Trim() != "" && data[i].Split(",").Length > 1)
                                            {
                                                docID = data[i].Split(",")[0];
                                                docDOCNumberNO = data[i].Split(",")[1];
                                                int docQTY = 0;
                                                isRUN = int.TryParse(data_QTY[i], out out_qty);
                                                if (isRUN && out_qty != 0 && docID != "" && docDOCNumberNO != "")
                                                {
                                                    dr_DOCRole = db.DB_GetFirstDataByDataRow($"select DOCType from dbo.DOCRole where DOCNO='{docDOCNumberNO.Substring(0, 4)}'");
                                                    if (dr_DOCRole == null) { continue; }
                                                    else
                                                    {
                                                        switch (dr_DOCRole["DOCType"].ToString())
                                                        {
                                                            case "1": DOC1Buy(db, docID, docDOCNumberNO, docQTY, out_qty, ref doc_Id_List); break;
                                                            case "2":
                                                            case "3": 
                                                            case "4": 
                                                            case "5": DOC3stock(db, docID, docDOCNumberNO, docQTY, out_qty, ref doc_Id_List); break;
                                                            case "6": 
                                                            case "7": DOC4Production(db, docID, docDOCNumberNO, docQTY, out_qty, ref doc_Id_List); break;
                                                            case "8": DOC3stock(db, docID, docDOCNumberNO, docQTY, out_qty, ref doc_Id_List); break;
                                                            case "9": break;
                                                        }
                                                    }
                                                    ViewBag.Report = "已完成單據入庫作業.";
                                                }

                                            }
                                        }
                                        if (doc_Id_List.Count > 0)
                                        {
                                            #region 儲位亮燈
                                            var remacIDs = "";
                                            var json3 = "";
                                            string storeNO = "";
                                            string storeSpacesN = "";
                                            using (HttpClient httpClient = new HttpClient())
                                            {
                                                DateTime now = DateTime.Now;
                                                try
                                                {
                                                    foreach (KeyValuePair<string, List<DataRow>> obj in doc_Id_List)
                                                    {
                                                        foreach (DataRow d in obj.Value)
                                                        {
                                                            if (obj.Key == "DOC3")
                                                            {
                                                                if (d["IN_StoreNO"].ToString() != "" && d["IN_StoreSpacesNO"].ToString() != "") { storeNO = d["IN_StoreNO"].ToString(); storeSpacesN = d["IN_StoreSpacesNO"].ToString(); }
                                                                else if (d["OUT_StoreNO"].ToString() != "" && d["OUT_StoreSpacesNO"].ToString() != "") { storeNO = d["OUT_StoreNO"].ToString(); storeSpacesN = d["OUT_StoreSpacesNO"].ToString(); }
                                                            }
                                                            else
                                                            { storeNO = d["StoreNO"].ToString(); storeSpacesN = d["StoreSpacesNO"].ToString(); }
                                                            DataRow tmp_SNO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[StoreII] where StoreNO='{storeNO}' and StoreSpacesNO='{storeSpacesN}'");
                                                            if (tmp_SNO != null)
                                                            {
                                                                if (remacIDs == "") { remacIDs = $"{tmp_SNO["Config_macID"].ToString().Trim()}"; }
                                                                else { remacIDs = $"{remacIDs},{tmp_SNO["Config_macID"].ToString().Trim()}"; }
                                                                if (_Fun.Is_Tag_Connect)
                                                                {
                                                                    if (json3 == "")
                                                                    { json3 = "[{" + $"\"mac\":\"{tmp_SNO["Config_macID"].ToString().Trim()}\",\"outtime\":0,\"ledrgb\":\"ff0000\",\"ledmode\":0,\"buzzer\":0,\"lednum\":255,\"reserve\":\"reserve\"" + "}"; }
                                                                    else
                                                                    { json3 = json3 + ",{" + $"\"mac\":\"{tmp_SNO["Config_macID"].ToString().Trim()}\",\"outtime\":0,\"ledrgb\":\"ff0000\",\"ledmode\":0,\"buzzer\":0,\"lednum\":255,\"reserve\":\"reserve\"" + "}"; }
                                                                }
                                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set SimulationId='{d["SimulationId"].ToString()}',DOCNumberNO='{d["DOCNumberNO"].ToString()}',Id='{d["Id"].ToString()}',Ledrgb='ff0000',Ledstate=0,IsUpdate='1' where ServerId='{_Fun.Config.ServerId}' and macID='{tmp_SNO["Config_macID"].ToString().Trim()}'");
                                                                db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType]) VALUES ('{_Str.NewId('L')}','{tmp_SNO["Config_macID"].ToString()}','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈')");
                                                            }
                                                        }
                                                    }
                                                    if (json3 != "") { json3 += "]"; }
                                                    string url = $"http://{_Fun.Config.ElectronicTagsURL}/wms/associate/lightTagsLed";
                                                    var content = new StringContent(json3, Encoding.UTF8, "application/json");
                                                    HttpResponseMessage response = httpClient.PostAsync(url, content).Result;
                                                    if (!response.IsSuccessStatusCode)
                                                    {
                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType],[JSON]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈,發送Fail','{remacIDs}')");
                                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"儲位亮燈 傳送電子訊號失敗,請通知管理者", false);    //false here, not mailRoot, or endless roop !!
                                                    }
                                                    else
                                                    {
                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType],[JSON]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈,發送OK','{remacIDs}')");
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"LabelStroeController.cs 儲位亮燈 Exception: {ex.Message} {ex.StackTrace}", true);
                                                }
                                            }
                                            #endregion
                                        }

                                        ViewBag.StoreNO = keys.StoreNO;
                                        ViewBag.StoreSpacesNO = keys.StoreSpacesNO;
                                        return View("ResuItTimeOUT");
                                    }
                                }
                            }
                            break;
                        case "B"://針對倉庫入庫
                            {
                                if (keys.Station == "" || keys.Select_ID == null || keys.Select_ID.Trim() == "")
                                {
                                    keys.ERRMsg = $"{keys.ERRMsg}<br />沒有選擇單據項目, 請重新選擇.";
                                }
                                else
                                {
                                    string[] data = keys.Select_ID.Split(';');
                                    string[] data_QTY = keys.Select_ID_QTY.Split(';');
                                    if (data.Length != data_QTY.Length)
                                    {
                                        keys.ERRMsg = $"{keys.ERRMsg}<br />勾選項目與數量不符合, 請重新選擇 或連繫系統管理員.";
                                    }
                                    else
                                    {
                                        DataRow dr_TotalStock = null;
                                        DataRow tmp_dr = null;
                                        int out_qty = 0;
                                        bool isRUN = false;
                                        string docNumberNO = "";
                                        string report = "";

                                        DataRow dr_DOC = null;
                                        DataRow dr_tag = null;
                                        Dictionary<string, List<DataRow>> doc_Id_List = new Dictionary<string, List<DataRow>>();
                                        #region  開門
                                        if (api_URL != "")
                                        { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, false); }
                                        #endregion 

                                        for (int i = 0; i < data.Length; i++)
                                        {
                                            if (data[i].Trim() == "") { continue; }
                                            dr_TotalStock = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and Id='{data[i].Trim()}'");
                                            isRUN = int.TryParse(data_QTY[i], out out_qty);
                                            if (dr_TotalStock == null || !isRUN)
                                            {
                                                ViewBag.ERRMsg = $"{ViewBag.ERRMsg}<br />{data[i].Trim()}項目, 數量={data_QTY[i]} 沒成功.";
                                            }
                                            else
                                            {
                                                #region 寫入庫存
                                                tmp_dr = db.DB_GetFirstDataByDataRow($"select PartNO,'' as NeedId,'' as SimulationId from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_TotalStock["PartNO"].ToString()}'");
                                                if (tmp_dr != null)
                                                {
                                                    dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_TotalStock["StoreNO"].ToString()}' and StoreSpacesNO='{dr_TotalStock["StoreSpacesNO"].ToString()}'");
                                                    if (dr_tag != null && dr_tag["Config_macID"].ToString() != "")
                                                    {
                                                        string docID = "";
                                                        _SFC_Common.Create_DOC3stock(db, tmp_dr,"","", dr_TotalStock["StoreNO"].ToString(), dr_TotalStock["StoreSpacesNO"].ToString(), "BC09", out_qty, "", ref docID, $"人為入料", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "SetAction_MIN;LabelStroeController", ref docNumberNO, br.UserNO.Trim(), true, false);
                                                        dr_DOC = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where Id='{docID}' and DOCNumberNO='{docNumberNO}' and IsOK='0'");
                                                        if (dr_DOC != null)
                                                        {
                                                            #region 儲位亮燈
                                                            if (doc_Id_List.ContainsKey("DOC3"))
                                                            { doc_Id_List["DOC3"].Add(dr_DOC); }
                                                            else { doc_Id_List.Add("DOC3", new List<DataRow> { dr_DOC }); }
                                                            #endregion
                                                        }
                                                    }
                                                    else
                                                    {
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={out_qty.ToString()} where ServerId='{_Fun.Config.ServerId}' and Id='{data[i].Trim()}'");
                                                        //###???暫時寫死AC09
                                                        _SFC_Common.Create_DOC3stock(db, tmp_dr, "", "", dr_TotalStock["StoreNO"].ToString(), dr_TotalStock["StoreSpacesNO"].ToString(), "BC09", out_qty, "", "", $"人為入料", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "SetAction_MOUT;LabelStroeController", ref docNumberNO, br.UserNO.Trim(), true, true);
                                                    }
                                                    if (report == "")
                                                    { report = $"系統自動產生 單號:{docNumberNO}&nbsp;&nbsp;明細如下:"; }
                                                    report = $"{report}<br />料號:{dr_TotalStock["PartNO"].ToString()}&nbsp;&nbsp;入庫:{data_QTY[i]}數量";
                                                }
                                                #endregion
                                            }
                                        }
                                        if (doc_Id_List.Count > 0)
                                        {
                                            #region 儲位亮燈
                                            var remacIDs = "";
                                            var json3 = "";
                                            string storeNO = "";
                                            string storeSpacesN = "";
                                            using (HttpClient httpClient = new HttpClient())
                                            {
                                                DateTime now = DateTime.Now;
                                                try
                                                {
                                                    foreach (KeyValuePair<string, List<DataRow>> obj in doc_Id_List)
                                                    {
                                                        foreach (DataRow d in obj.Value)
                                                        {
                                                            if (obj.Key == "DOC3")
                                                            {
                                                                if (d["IN_StoreNO"].ToString() != "" && d["IN_StoreSpacesNO"].ToString() != "") { storeNO = d["IN_StoreNO"].ToString(); storeSpacesN = d["IN_StoreSpacesNO"].ToString(); }
                                                                else if (d["OUT_StoreNO"].ToString() != "" && d["OUT_StoreSpacesNO"].ToString() != "") { storeNO = d["OUT_StoreNO"].ToString(); storeSpacesN = d["OUT_StoreSpacesNO"].ToString(); }
                                                            }
                                                            else
                                                            { storeNO = d["StoreNO"].ToString(); storeSpacesN = d["StoreSpacesNO"].ToString(); }
                                                            DataRow tmp_SNO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[StoreII] where StoreNO='{storeNO}' and StoreSpacesNO='{storeSpacesN}'");
                                                            if (tmp_SNO != null)
                                                            {
                                                                if (remacIDs == "") { remacIDs = $"{tmp_SNO["Config_macID"].ToString().Trim()}"; }
                                                                else { remacIDs = $"{remacIDs},{tmp_SNO["Config_macID"].ToString().Trim()}"; }
                                                                if (_Fun.Is_Tag_Connect)
                                                                {
                                                                    if (json3 == "")
                                                                    { json3 = "[{" + $"\"mac\":\"{tmp_SNO["Config_macID"].ToString().Trim()}\",\"outtime\":0,\"ledrgb\":\"ff0000\",\"ledmode\":0,\"buzzer\":0,\"lednum\":255,\"reserve\":\"reserve\"" + "}"; }
                                                                    else
                                                                    { json3 = json3 + ",{" + $"\"mac\":\"{tmp_SNO["Config_macID"].ToString().Trim()}\",\"outtime\":0,\"ledrgb\":\"ff0000\",\"ledmode\":0,\"buzzer\":0,\"lednum\":255,\"reserve\":\"reserve\"" + "}"; }
                                                                }
                                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set SimulationId='{d["SimulationId"].ToString()}',DOCNumberNO='{d["DOCNumberNO"].ToString()}',Id='{d["Id"].ToString()}',Ledrgb='ff0000',Ledstate=0,IsUpdate='1' where ServerId='{_Fun.Config.ServerId}' and macID='{tmp_SNO["Config_macID"].ToString().Trim()}'");
                                                                db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType]) VALUES ('{_Str.NewId('L')}','{tmp_SNO["Config_macID"].ToString()}','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈')");
                                                            }
                                                        }
                                                    }
                                                    if (json3 != "") { json3 += "]"; }
                                                    string url = $"http://{_Fun.Config.ElectronicTagsURL}/wms/associate/lightTagsLed";
                                                    var content = new StringContent(json3, Encoding.UTF8, "application/json");
                                                    HttpResponseMessage response = httpClient.PostAsync(url, content).Result;
                                                    if (!response.IsSuccessStatusCode)
                                                    {
                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType],[JSON]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈,發送Fail','{remacIDs}')");
                                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"儲位亮燈 傳送電子訊號失敗,請通知管理者", false);    //false here, not mailRoot, or endless roop !!
                                                    }
                                                    else
                                                    {
                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType],[JSON]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈,發送OK','{remacIDs}')");
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"LabelStroeController.cs 儲位亮燈 Exception: {ex.Message} {ex.StackTrace}", true);
                                                }
                                            }
                                            #endregion
                                        }

                                        ViewBag.Report = report;
                                        ViewBag.StoreNO = keys.StoreNO;
                                        ViewBag.StoreSpacesNO = keys.StoreSpacesNO;
                                        return View("ResuItTimeOUT");
                                    }

                                }
                            }
                            break;
                    }
                    if (keys.ActionType == "A" || keys.ActionType == "B")
                    {
                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and Keys='入料_{br.UserNO.Trim()}' and StoreNO='{keys.StoreNO}'");
                    }


                    DataTable dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and IN_StoreNO='{keys.StoreNO}' and IN_StoreSpacesNO='{keys.StoreSpacesNO}' and IsOK='0' order by DOCNumberNO,IN_StoreSpacesNO desc,PartNO");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.SIDList = dt;
                    }
                    dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC1BuyII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' and StoreSpacesNO='{keys.StoreSpacesNO}'  and IsOK='0' order by DOCNumberNO,StoreSpacesNO desc,PartNO");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        if (keys.SIDList == null)
                        {
                            keys.SIDList = new DataTable();
                            keys.SIDList.Columns.Add("Id");
                            keys.SIDList.Columns.Add("DOCNumberNO");
                            keys.SIDList.Columns.Add("ArrivalDate", typeof(DateTime));
                            keys.SIDList.Columns.Add("ServerId");
                            keys.SIDList.Columns.Add("PartNO");
                            keys.SIDList.Columns.Add("Price");
                            keys.SIDList.Columns.Add("Unit");
                            keys.SIDList.Columns.Add("QTY");
                            keys.SIDList.Columns.Add("WeightQty");
                            keys.SIDList.Columns.Add("Remark");
                            keys.SIDList.Columns.Add("SimulationId");
                            keys.SIDList.Columns.Add("IsOK");
                            keys.SIDList.Columns.Add("IN_StoreNO");
                            keys.SIDList.Columns.Add("IN_StoreSpacesNO");
                            keys.SIDList.Columns.Add("OUT_StoreNO");
                            keys.SIDList.Columns.Add("OUT_StoreSpacesNO");
                            keys.SIDList.Columns.Add("EndTime", typeof(DateTime));
                            keys.SIDList.Columns.Add("StartTime", typeof(DateTime));
                            keys.SIDList.Columns.Add("CT");
                        }
                        DataRow dr_dtAdd = null;
                        foreach (DataRow dr in dt.Rows)
                        {

                            dr_dtAdd = keys.SIDList.NewRow();
                            dr_dtAdd[0] = dr["Id"].ToString(); dr_dtAdd[1] = dr["DOCNumberNO"].ToString(); dr_dtAdd[2] = dr.IsNull("ArrivalDate") ? DBNull.Value : Convert.ToDateTime(dr["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss");
                            dr_dtAdd[3] = dr["ServerId"].ToString(); dr_dtAdd[4] = dr["PartNO"].ToString(); dr_dtAdd[5] = dr["Price"].ToString(); dr_dtAdd[6] = dr["Unit"].ToString(); dr_dtAdd[7] = dr["QTY"].ToString();
                            dr_dtAdd[8] = dr["WeightQty"].ToString(); dr_dtAdd[9] = dr["Remark"].ToString(); dr_dtAdd[10] = dr["SimulationId"].ToString(); dr_dtAdd[11] = dr["IsOK"].ToString(); dr_dtAdd[12] = dr["StoreNO"].ToString(); dr_dtAdd[13] = dr["StoreSpacesNO"].ToString();
                            dr_dtAdd[14] = ""; dr_dtAdd[15] = ""; dr_dtAdd[16] = dr.IsNull("EndTime") ? DBNull.Value : Convert.ToDateTime(dr["EndTime"]).ToString("yyyy-MM-dd HH:mm:ss"); dr_dtAdd[17] = dr.IsNull("StartTime") ? DBNull.Value : Convert.ToDateTime(dr["StartTime"]).ToString("yyyy-MM-dd HH:mm:ss");
                            dr_dtAdd[18] = dr["CT"].ToString();
                            keys.SIDList.Rows.Add(dr_dtAdd);
                        }
                    }
                    dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC4ProductionII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' and StoreSpacesNO='{keys.StoreSpacesNO}' and IsOK='0' order by DOCNumberNO,StoreSpacesNO desc,PartNO");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        if (keys.SIDList == null)
                        {
                            keys.SIDList = new DataTable();
                            keys.SIDList.Columns.Add("Id");
                            keys.SIDList.Columns.Add("DOCNumberNO");
                            keys.SIDList.Columns.Add("ArrivalDate", typeof(DateTime));
                            keys.SIDList.Columns.Add("ServerId");
                            keys.SIDList.Columns.Add("PartNO");
                            keys.SIDList.Columns.Add("Price");
                            keys.SIDList.Columns.Add("Unit");
                            keys.SIDList.Columns.Add("QTY");
                            keys.SIDList.Columns.Add("WeightQty");
                            keys.SIDList.Columns.Add("Remark");
                            keys.SIDList.Columns.Add("SimulationId");
                            keys.SIDList.Columns.Add("IsOK");
                            keys.SIDList.Columns.Add("IN_StoreNO");
                            keys.SIDList.Columns.Add("IN_StoreSpacesNO");
                            keys.SIDList.Columns.Add("OUT_StoreNO");
                            keys.SIDList.Columns.Add("OUT_StoreSpacesNO");
                            keys.SIDList.Columns.Add("EndTime", typeof(DateTime));
                            keys.SIDList.Columns.Add("StartTime", typeof(DateTime));
                            keys.SIDList.Columns.Add("CT");
                        }
                        DataRow dr_dtAdd = null;
                        foreach (DataRow dr in dt.Rows)
                        {
                            dr_dtAdd = keys.SIDList.NewRow();
                            dr_dtAdd[0] = dr["Id"].ToString(); dr_dtAdd[1] = dr["DOCNumberNO"].ToString(); dr_dtAdd[2] = dr.IsNull("ArrivalDate") ? DBNull.Value : Convert.ToDateTime(dr["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss");
                            dr_dtAdd[3] = dr["ServerId"].ToString(); dr_dtAdd[4] = dr["PartNO"].ToString(); dr_dtAdd[5] = dr["Price"].ToString(); dr_dtAdd[6] = dr["Unit"].ToString(); dr_dtAdd[7] = dr["QTY"].ToString();
                            dr_dtAdd[8] = dr["WeightQty"].ToString(); dr_dtAdd[9] = dr["Remark"].ToString(); dr_dtAdd[10] = dr["SimulationId"].ToString(); dr_dtAdd[11] = dr["IsOK"].ToString(); dr_dtAdd[12] = dr["StoreNO"].ToString(); dr_dtAdd[13] = dr["StoreSpacesNO"].ToString();
                            dr_dtAdd[14] = ""; dr_dtAdd[15] = ""; dr_dtAdd[16] = dr.IsNull("EndTime") ? DBNull.Value : Convert.ToDateTime(dr["EndTime"]).ToString("yyyy-MM-dd HH:mm:ss"); dr_dtAdd[17] = dr.IsNull("StartTime") ? DBNull.Value : Convert.ToDateTime(dr["StartTime"]).ToString("yyyy-MM-dd HH:mm:ss");
                            dr_dtAdd[18] = dr["CT"].ToString();
                            keys.SIDList.Rows.Add(dr_dtAdd);
                        }
                    }
                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Store]  where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' and Class='大正倉'");
                    if (dr_tmp == null)
                    {
                        sql = $@"SELECT a.*  FROM SoftNetMainDB.[dbo].[TotalStock] as a
                        join SoftNetMainDB.[dbo].[Material] as b on b.Class!='6' and b.Class!='7' and a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                        where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{keys.StoreNO}' and a.StoreSpacesNO='{keys.StoreSpacesNO}' order by PartNO,StoreSpacesNO desc,QTY desc";
                    }
                    else
                    {
                        sql = $@"SELECT a.*  FROM SoftNetMainDB.[dbo].[TotalStock] as a
                        join SoftNetMainDB.[dbo].[Material] as b on (b.Class='6' or b.Class='7') and a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                        where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{keys.StoreNO}' and a.StoreSpacesNO='{keys.StoreSpacesNO}' order by PartNO,StoreSpacesNO desc,QTY desc";
                    }
                    dt = db.DB_GetData(sql);
                    //dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' order by PartNO,StoreSpacesNO desc,QTY desc");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.TotalStock_List = dt;
                    }
                    dt = db.DB_GetData($"select StoreSpacesNO from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' and StoreSpacesNO='{keys.StoreSpacesNO}' group by StoreSpacesNO order by StoreSpacesNO");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.StoreSpacesNO_List = new List<string>();
                        foreach (DataRow d in dt.Rows)
                        {
                            if (!keys.StoreSpacesNO_List.Contains(d["StoreSpacesNO"].ToString())) { keys.StoreSpacesNO_List.Add(d["StoreSpacesNO"].ToString()); }
                        }
                    }

                }
                else
                {
                    ViewBag.ERRMsg = $"{ViewBag.ERRMsg}<br />網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
                    ViewBag.StoreNO = "";
                    ViewBag.StoreSpacesNO = "";
                    return View("ResuItTimeOUT");
                }
            }
            keys.Select_ID = "";
            keys.Select_ID_QTY = "";
            ViewBag.StoreList = keys;
            return View();
        }
        public ActionResult SetAction_KINOUT(StoreList keys)
        {
            var br = _Fun.GetBaseUser();
            if (keys == null || keys.StoreNO == null || keys.StoreNO == "" || br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                ViewBag.ERRMsg = "網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
                ViewBag.StoreNO = "";
                ViewBag.StoreSpacesNO = "";
                return View("ResuItTimeOUT");
            }

            keys.ERRMsg = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string api_URL = "";
                string api_Token = "";
                DataRow dr_Store = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}'");
                if (dr_Store != null)
                {
                    #region 大正貴訊號處理 開門
                    if (dr_Store["Store_Type"].ToString() == "A")
                    {
                        GetEStoreURLandToken(keys.StoreNO, ref api_URL, ref api_Token);
                        if (api_URL == "" || api_Token == "") { keys.ERRMsg = "此倉為電子控制倉, 但無API參數設定."; }
                    }
                    #endregion

                    DataTable dt = db.DB_GetData($@"select a.*,b.Unit from SoftNetMainDB.[dbo].[TotalStock] as a
                    join SoftNetMainDB.[dbo].[Material] as b on b.PartNO=a.PartNO and b.Class='6' and b.ServerId='{_Fun.Config.ServerId}'
                    where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{keys.StoreNO}' and a.StoreSpacesNO='{keys.StoreSpacesNO}'");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.New_K_List = dt;
                    }
                    dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStockII_Knives] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' and StoreSpacesNO='{keys.StoreSpacesNO}' and StationNO='' and IsDel='0'");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.OLD_K_List = dt;
                    }
                    dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStockII_Knives] where ServerId='{_Fun.Config.ServerId}' and StationNO!='' and IsDel='0'");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.OLD_K_IN_Station_List = dt;
                    }
                    dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and IsKnives='1'");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.Station_List = new List<string[]>();
                        keys.Station_List.Add(new string[] { "", "" });
                        foreach (DataRow dr in dt.Rows)
                        {
                            keys.Station_List.Add(new string[] { dr["StationNO"].ToString(), dr["StationName"].ToString() });
                        }
                    }
                    dt = db.DB_GetData($"select StoreSpacesNO from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' and StoreSpacesNO='{keys.StoreSpacesNO}' group by StoreSpacesNO order by StoreSpacesNO");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.StoreSpacesNO_List = new List<string>();
                        foreach (DataRow d in dt.Rows)
                        {
                            if (!keys.StoreSpacesNO_List.Contains(d["StoreSpacesNO"].ToString())) { keys.StoreSpacesNO_List.Add(d["StoreSpacesNO"].ToString()); }
                        }
                    }
                    switch (keys.ActionType)
                    {
                        case "A"://領新刀
                            {
                                if (keys.Station == null || keys.Station == "" || keys.Select_ID == null || keys.Select_ID.Trim() == "")
                                {
                                    keys.ERRMsg = $"{keys.ERRMsg}<br />沒有選擇工作項目 或 工站, 請重新選擇.";
                                }
                                else
                                {
                                    string[] data = keys.Select_ID.Split(';');

                                    #region  開門
                                    if (api_URL != "")
                                    { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, false); }
                                    #endregion

                                    foreach (string id in data)
                                    {
                                        if (id.Trim() != "")
                                        {
                                            if (db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStock] set QTY-=1 where ServerId='{_Fun.Config.ServerId}' and Id='{id}'"))
                                            {
                                                db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives] (ServerId,[KId],[MId],[StationNO]) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('K')}','{id}','{keys.Station}')");
                                            }
                                        }
                                    }
                                    ViewBag.StoreNO = keys.StoreNO;
                                    ViewBag.StoreSpacesNO = keys.StoreSpacesNO;
                                    ViewBag.Report = "已完成新刀歷程建檔, 後續本刀請以舊刀方式處理..";
                                    return View("ResuItTimeOUT");
                                }
                            }
                            break;
                        case "B"://舊刀領取
                            {
                                if (keys.Station == null || keys.Station == "" || keys.Select_ID == null || keys.Select_ID.Trim() == "")
                                {
                                    keys.ERRMsg = $"{keys.ERRMsg}<br />沒有選擇舊刀項目 或 工站, 請重新選擇.";
                                }
                                else
                                {
                                    string[] data = keys.Select_ID.Split(';');

                                    #region  開門
                                    if (api_URL != "")
                                    { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, false); }
                                    #endregion

                                    foreach (string id in data)
                                    {
                                        if (id.Trim() != "")
                                        {
                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStockII_Knives] set StationNO='{keys.Station}',StoreNO='',StoreSpacesNO='' where KId='{id}'");
                                        }
                                    }
                                    ViewBag.StoreNO = keys.StoreNO;
                                    ViewBag.StoreSpacesNO = keys;
                                    ViewBag.Report = "已完成舊刀領取出庫.";
                                    return View("ResuItTimeOUT");
                                }
                            }
                            break;
                        case "C"://舊刀歸回
                            {
                                if (keys.Select_ID == null || keys.Select_ID.Trim() == "")
                                {
                                    keys.ERRMsg = $"{keys.ERRMsg}<br />沒有選擇舊刀項目, 請重新選擇.";
                                }
                                else
                                {
                                    string[] data = keys.Select_ID.Split(';');

                                    #region  開門
                                    if (api_URL != "")
                                    { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, false); }
                                    #endregion

                                    foreach (string id in data)
                                    {
                                        if (id.Trim() != "")
                                        {
                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStockII_Knives] set StationNO='',StoreNO='{keys.StoreNO}',StoreSpacesNO='{keys.StoreSpacesNO}' where KId='{id}'");
                                        }
                                    }
                                    ViewBag.StoreNO = keys.StoreNO;
                                    ViewBag.StoreSpacesNO = keys.StoreSpacesNO;
                                    ViewBag.Report = "已完成舊刀歸回入庫.";
                                    return View("ResuItTimeOUT");
                                }
                            }
                            break;
                        case "D"://作廢再生
                            {
                                if (keys.Select_ID == null || keys.Select_ID.Trim() == "")
                                {
                                    keys.ERRMsg = $"{keys.ERRMsg}<br />沒有選擇舊刀項目, 請重新選擇.";
                                }
                                else
                                {
                                    string[] data = keys.Select_ID.Split(';');
                                    foreach (string id in data)
                                    {
                                        if (id.Trim() != "")
                                        {
                                            if (keys.Select_DroneType == "1")
                                            {
                                                #region 作廢
                                                db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStockII_Knives] set IsDel='1' where KId='{id}'");
                                                #endregion
                                            }
                                            else
                                            {
                                                #region 再生
                                                db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStockII_Knives] set TOTWorkTime=0,TOTCount=0,RecoverTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' where KId='{id}'");
                                                #endregion
                                            }
                                        }
                                    }
                                    ViewBag.StoreNO = keys.StoreNO;
                                    ViewBag.StoreSpacesNO = keys.StoreSpacesNO;
                                    if (keys.Select_DroneType == "1")
                                    { ViewBag.Report = "已完成舊刀作廢作業."; }
                                    else { ViewBag.Report = "已完成舊刀再生作業, 此刀壽命依設定重新計算."; }
                                    return View("ResuItTimeOUT");
                                }


                            }
                            break;
                        case "E"://舊刀歷程
                            {

                            }
                            break;
                    }
                }
                else
                {
                    ViewBag.ERRMsg = "網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
                    ViewBag.StoreNO = "";
                    ViewBag.StoreSpacesNO = "";
                    return View("ResuItTimeOUT");
                }
            }
            keys.Select_ID = "";
            keys.Select_ID_QTY = "";
            ViewBag.StoreList = keys;
            return View();
        }
        [HttpPost]
        public ActionResult On_OLD_C_Select_ID(string data)
        {
            string meg = "";

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                #region 舊刀在現清單
                DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStockII_Knives] where StationNO='{data}' and IsDel='0'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (meg == "") { meg = $"{dr["KId"].ToString()}"; }
                        else
                        { meg = $"{meg},{dr["KId"].ToString()}"; }
                    }
                }
                #endregion
            }
            return Content(meg);
        }
        [HttpPost]
        public ActionResult On_E_GetDATA(string data)//刀使用歷程
        {
            string meg = "";
            if (data == "") { return Content(meg); }
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($@"select b.* from SoftNetMainDB.[dbo].[TotalStockII_Knives] as a
                                                join SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] as b on a.KId=b.KId
                                                where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{data}' and a.IsDel='0' order by KId,LOGDateTime desc,PartNO");
                if (dt != null && dt.Rows.Count > 0)
                {
                    DataRow tmp_dr = null;
                    string tmp_09 = "";
                    List<string> kId_List = new List<string>();
                    meg = $"<div><table data-role='table' data-mode='columntoggle' class='ui-responsive ui-shadow' id='DisplayDataTable_E' border='1'><thead><tr><th>KId</th><th>歷程日期</th><th>歷程料件編號</th><th>品名</th><th>規格</th><th>單次生產量</th><th>單次時數</th></tr></thead><tbody>";
                    foreach (DataRow dr in dt.Rows)
                    {
                        tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT *  FROM SoftNetMainDB.[dbo].TotalStockII_Knives where KId='{dr["KId"].ToString()}'");
                        if (tmp_dr != null && !tmp_dr.IsNull("RecoverTime"))
                        {
                            if (Convert.ToDateTime(tmp_dr["RecoverTime"]) >= Convert.ToDateTime(dr["LOGDateTime"])) { continue; }
                        }
                        if (!kId_List.Contains(dr["KId"].ToString())) { kId_List.Add(dr["KId"].ToString()); }
                        TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(dr["WorkTime"].ToString()));
                        tmp_09 = $"{(int)standardTime_DIS.TotalHours}時{standardTime_DIS.Minutes}分{standardTime_DIS.Seconds}秒";
                        tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                        if (tmp_dr != null)
                        { meg = $"{meg}<tr><td>{dr["KId"].ToString()}</td><td>{dr["LOGDateTime"].ToString()}</td><td>{dr["PartNO"].ToString()}</td><td>{tmp_dr["PartName"].ToString()}</td><td>{tmp_dr["Specification"].ToString()}</td><td>{dr["WorkQTY"].ToString()}</td><td>{tmp_09}</td></tr>"; }
                        else { meg = $"{meg}<tr><td>{dr["KId"].ToString()}</td><td>{dr["LOGDateTime"].ToString()}</td><td>{dr["PartNO"].ToString()}</td><td></td><td></td><td>{dr["WorkQTY"].ToString()}</td><td>{tmp_09}</td></tr>"; }
                    }
                    meg = $"{meg}</tbody></table></div>";
                    if (kId_List.Count > 0)
                    {
                        string tmp_01 = "";
                        string tmp = $"<div><table data-role='table' data-mode='columntoggle' class='ui-responsive ui-shadow' id='DisplayDataTable_E_KId' border='1'><thead><tr><th>KId</th><th>目前位置</th><th>總使用次數</th><th>總使用時數</th></tr></thead><tbody>"; ;
                        foreach (string s in kId_List)
                        {
                            tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[TotalStockII_Knives] where ServerId='{_Fun.Config.ServerId}' and KId='{s}' and IsDel='0'");
                            if (tmp_dr != null)
                            {
                                TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(tmp_dr["TOTWorkTime"].ToString()));
                                tmp_09 = $"{(int)standardTime_DIS.TotalHours}時{standardTime_DIS.Minutes}分{standardTime_DIS.Seconds}秒";
                                if (tmp_dr["StationNO"].ToString().Trim() != "") { tmp_01 = $"在線:{tmp_dr["StationNO"].ToString()}站"; }
                                else { tmp_01 = $"在庫:{tmp_dr["StoreNO"].ToString()}&nbsp;{tmp_dr["StoreSpacesNO"].ToString()}"; }
                                tmp = $"{tmp}<tr><td>{s}</td><td>{tmp_01}</td><td>{tmp_dr["TOTCount"].ToString()}</td><td>{tmp_09}</td></tr>";
                            }
                        }
                        tmp = $"{tmp}</tbody></table></div>";
                        meg = $"<label>歷程明細</label>{tmp}<br />{meg}";
                    }
                }
            }
            return Content(meg);
        }

        public ActionResult SetStoreChange(List<LabelDOC3stockII> keys)
        {
            return View(keys);
        }

        private void DOC3stock(DBADO db, string docID, string docDOCNumberNO, int docQTY, int out_qty, ref Dictionary<string, List<DataRow>> doc_Id_List)
        {
            string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
            #region 計算單據CT,平均,有效,寫入庫存, 寫IsOK='1'
            DataRow dr_DOC3stockII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where Id='{docID}' and DOCNumberNO='{docDOCNumberNO}' and IsOK='0'");
            if (dr_DOC3stockII != null)
            {
                docQTY = int.Parse(dr_DOC3stockII["QTY"].ToString());
                if (out_qty != docQTY)
                {
                    if (out_qty > docQTY)
                    {
                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set QTY={out_qty} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}'");
                    }
                    else
                    {
                        //拆單處置
                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set QTY={out_qty} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}'");
                        db.DB_SetData($@"INSERT INTO [dbo].[DOC3stockII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],ArrivalDate) VALUES 
                                                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{dr_DOC3stockII["DOCNumberNO"].ToString()}','{dr_DOC3stockII["PartNO"].ToString()}',{dr_DOC3stockII["Price"].ToString()},'{dr_DOC3stockII["Unit"].ToString()}',{(docQTY - out_qty).ToString()}
                                                                                                ,'{dr_DOC3stockII["Remark"].ToString()}','{dr_DOC3stockII["SimulationId"].ToString()}','{dr_DOC3stockII["IsOK"].ToString()}','{dr_DOC3stockII["IN_StoreNO"].ToString()}','{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}','{dr_DOC3stockII["OUT_StoreNO"].ToString()}','{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}','{Convert.ToDateTime(dr_DOC3stockII["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                    }
                }
                DataRow dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'");
                if (dr_tag != null && dr_tag["Config_macID"].ToString() != "")
                {
                    #region 儲位亮燈
                    if (doc_Id_List.ContainsKey("DOC3"))
                    { doc_Id_List["DOC3"].Add(dr_DOC3stockII); }
                    else { doc_Id_List.Add("DOC3", new List<DataRow> { dr_DOC3stockII }); }
                    #endregion
                }
                else
                {
                    #region 寫入庫存
                    if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                    { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={out_qty} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["IN_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}'"); }
                    else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={out_qty} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'"); }
                    #endregion

                    int typeTotalTime = 0;
                    string writeSQL = "";
                    if (!dr_DOC3stockII.IsNull("StartTime")) { typeTotalTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC3stockII["StartTime"].ToString()), DateTime.Now); }
                    else { writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime}{writeSQL} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}' and IsOK='0'");
                    string partNO = dr_DOC3stockII["PartNO"].ToString();
                    string pp_Name = "";
                    string E_stationNO = "";
                    if (dr_DOC3stockII["SimulationId"].ToString().Trim() != "")
                    {
                        DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC3stockII["SimulationId"].ToString().Trim()}'");
                        pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                        if (!dr_tmp.IsNull("Source_StationNO") && dr_tmp["Source_StationNO"].ToString() != "")
                        { E_stationNO = dr_tmp["Source_StationNO"].ToString(); }
                        else { E_stationNO = dr_tmp["Apply_StationNO"].ToString(); }
                    }
                    DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} CT from SoftNetMainDB.[dbo].[DOC3stockII] where SUBSTRING(Id,2,2)='{_Fun.Config.ServerId}' and SUBSTRING(DOCNumberNO,1,4)='{dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4)}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and CT>0");
                    List<double> allCT = new List<double>();
                    if (dt_Efficient != null && dt_Efficient.Rows.Count > 0)
                    {
                        for (int i2 = 0; i2 < dt_Efficient.Rows.Count; i2++)
                        {
                            foreach (DataRow dr2 in dt_Efficient.Rows)
                            {
                                allCT.Add(double.Parse(dr2["CT"].ToString()));
                            }
                        }
                    }
                    else
                    {
                        if (typeTotalTime != 0)
                        { allCT.Add(typeTotalTime); }
                    }
                    if (allCT.Count > 0)
                    {
                        _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, pp_Name, pp_Name, "0", partNO, partNO, dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4));
                    }
                }
            }
            #endregion

        }
        private void DOC1Buy(DBADO db, string docID, string docDOCNumberNO, int docQTY, int out_qty, ref Dictionary<string, List<DataRow>> doc_Id_List)
        {
            string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
            #region 計算單據CT,平均,有效,寫入庫存, 寫IsOK='1'
            DataRow dr_DOC1BuyII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC1BuyII] where Id='{docID}' and DOCNumberNO='{docDOCNumberNO}' and IsOK='0'");
            if (dr_DOC1BuyII != null)
            {
                docQTY = int.Parse(dr_DOC1BuyII["QTY"].ToString());
                if (out_qty != docQTY)
                {
                    if (out_qty > docQTY)
                    {
                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC1BuyII] set QTY={out_qty} where Id='{dr_DOC1BuyII["Id"].ToString()}' and DOCNumberNO='{dr_DOC1BuyII["DOCNumberNO"].ToString()}'");
                    }
                    else
                    {
                        //拆單處置
                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC1BuyII] set QTY={out_qty} where Id='{dr_DOC1BuyII["Id"].ToString()}' and DOCNumberNO='{dr_DOC1BuyII["DOCNumberNO"].ToString()}'");
                        db.DB_SetData($@"INSERT INTO [dbo].[DOC1BuyII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[StoreNO],[StoreSpacesNO],ArrivalDate) VALUES 
                                                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{dr_DOC1BuyII["DOCNumberNO"].ToString()}','{dr_DOC1BuyII["PartNO"].ToString()}',{dr_DOC1BuyII["Price"].ToString()},'{dr_DOC1BuyII["Unit"].ToString()}',{(docQTY - out_qty).ToString()}
                                                                                                ,'{dr_DOC1BuyII["Remark"].ToString()}','{dr_DOC1BuyII["SimulationId"].ToString()}','{dr_DOC1BuyII["IsOK"].ToString()}','{dr_DOC1BuyII["StoreNO"].ToString()}','{dr_DOC1BuyII["StoreSpacesNO"].ToString()}','{Convert.ToDateTime(dr_DOC1BuyII["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                    }
                }
                DataRow dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_DOC1BuyII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC1BuyII["OUT_StoreSpacesNO"].ToString()}'");
                if (dr_tag != null && dr_tag["Config_macID"].ToString() != "")
                {
                    #region 儲位亮燈
                    if (doc_Id_List.ContainsKey("DOC1"))
                    { doc_Id_List["DOC1"].Add(dr_DOC1BuyII); }
                    else { doc_Id_List.Add("DOC1", new List<DataRow> { dr_DOC1BuyII }); }
                    #endregion
                }
                else
                {
                    #region 寫入庫存
                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={out_qty} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC1BuyII["PartNO"].ToString()}' and StoreNO='{dr_DOC1BuyII["StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC1BuyII["StoreSpacesNO"].ToString()}'");
                    #endregion

                    int typeTotalTime = 0;
                    if (!dr_DOC1BuyII.IsNull("StartTime")) { typeTotalTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC1BuyII["StartTime"].ToString()), DateTime.Now); }
                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC1BuyII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime} where Id='{dr_DOC1BuyII["Id"].ToString()}' and DOCNumberNO='{dr_DOC1BuyII["DOCNumberNO"].ToString()}' and IsOK='0'");
                    string partNO = dr_DOC1BuyII["PartNO"].ToString();
                    string pp_Name = "";
                    string E_stationNO = "";
                    if (dr_DOC1BuyII["SimulationId"].ToString().Trim() != "")
                    {
                        DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC1BuyII["SimulationId"].ToString().Trim()}'");
                        pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                        if (!dr_tmp.IsNull("Source_StationNO") && dr_tmp["Source_StationNO"].ToString() != "")
                        { E_stationNO = dr_tmp["Source_StationNO"].ToString(); }
                        else { E_stationNO = dr_tmp["Apply_StationNO"].ToString(); }
                    }
                    DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} CT from SoftNetMainDB.[dbo].[DOC1BuyII] where SUBSTRING(Id,2,2)='{_Fun.Config.ServerId}' and SUBSTRING(DOCNumberNO,1,4)='{dr_DOC1BuyII["DOCNumberNO"].ToString().Substring(0, 4)}' and PartNO='{dr_DOC1BuyII["PartNO"].ToString()}' and CT>0");
                    List<double> allCT = new List<double>();
                    if (dt_Efficient != null && dt_Efficient.Rows.Count > 0)
                    {
                        for (int i2 = 0; i2 < dt_Efficient.Rows.Count; i2++)
                        {
                            foreach (DataRow dr2 in dt_Efficient.Rows)
                            {
                                allCT.Add(double.Parse(dr2["CT"].ToString()));
                            }
                        }
                    }
                    else
                    {
                        if (typeTotalTime != 0)
                        { allCT.Add(typeTotalTime); }
                    }
                    if (allCT.Count > 0)
                    {
                        _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, pp_Name, pp_Name, "0", partNO, partNO, dr_DOC1BuyII["DOCNumberNO"].ToString().Substring(0, 4));
                    }
                }
            }
            #endregion

        }
        private void DOC4Production(DBADO db, string docID, string docDOCNumberNO, int docQTY, int out_qty, ref Dictionary<string, List<DataRow>> doc_Id_List)
        {
            string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
            #region 計算單據CT,平均,有效,寫入庫存, 寫IsOK='1'
            DataRow dr_DOC4ProductionII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC4ProductionII] where Id='{docID}' and DOCNumberNO='{docDOCNumberNO}' and IsOK='0'");
            if (dr_DOC4ProductionII != null)
            {
                docQTY = int.Parse(dr_DOC4ProductionII["QTY"].ToString());
                if (out_qty != docQTY)
                {
                    if (out_qty > docQTY)
                    {
                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC4ProductionII] set QTY={out_qty} where Id='{dr_DOC4ProductionII["Id"].ToString()}' and DOCNumberNO='{dr_DOC4ProductionII["DOCNumberNO"].ToString()}'");
                    }
                    else
                    {
                        //拆單處置
                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC4ProductionII] set QTY={out_qty} where Id='{dr_DOC4ProductionII["Id"].ToString()}' and DOCNumberNO='{dr_DOC4ProductionII["DOCNumberNO"].ToString()}'");
                        db.DB_SetData($@"INSERT INTO [dbo].[DOC4ProductionII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[StoreNO],[StoreSpacesNO],ArrivalDate) VALUES 
                                                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{dr_DOC4ProductionII["DOCNumberNO"].ToString()}','{dr_DOC4ProductionII["PartNO"].ToString()}',{dr_DOC4ProductionII["Price"].ToString()},'{dr_DOC4ProductionII["Unit"].ToString()}',{(docQTY - out_qty).ToString()}
                                                                                                ,'{dr_DOC4ProductionII["Remark"].ToString()}','{dr_DOC4ProductionII["SimulationId"].ToString()}','{dr_DOC4ProductionII["IsOK"].ToString()}','{dr_DOC4ProductionII["StoreNO"].ToString()}','{dr_DOC4ProductionII["StoreSpacesNO"].ToString()}','{Convert.ToDateTime(dr_DOC4ProductionII["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                    }
                }
                DataRow dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_DOC4ProductionII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC4ProductionII["OUT_StoreSpacesNO"].ToString()}'");
                if (dr_tag != null && dr_tag["Config_macID"].ToString() != "")
                {
                    #region 儲位亮燈
                    if (doc_Id_List.ContainsKey("DOC4"))
                    { doc_Id_List["DOC4"].Add(dr_DOC4ProductionII); }
                    else { doc_Id_List.Add("DOC4", new List<DataRow> { dr_DOC4ProductionII }); }
                    #endregion
                }
                else
                {
                    #region 寫入庫存
                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={out_qty} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC4ProductionII["PartNO"].ToString()}' and StoreNO='{dr_DOC4ProductionII["StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC4ProductionII["StoreSpacesNO"].ToString()}'");
                    #endregion

                    int typeTotalTime = 0;
                    if (!dr_DOC4ProductionII.IsNull("StartTime")) { typeTotalTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC4ProductionII["StartTime"].ToString()), DateTime.Now); }
                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC4ProductionII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime} where Id='{dr_DOC4ProductionII["Id"].ToString()}' and DOCNumberNO='{dr_DOC4ProductionII["DOCNumberNO"].ToString()}' and IsOK='0'");
                    string partNO = dr_DOC4ProductionII["PartNO"].ToString();
                    string pp_Name = "";
                    string E_stationNO = "";
                    if (dr_DOC4ProductionII["SimulationId"].ToString().Trim() != "")
                    {
                        DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC4ProductionII["SimulationId"].ToString().Trim()}'");
                        pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                        if (!dr_tmp.IsNull("Source_StationNO") && dr_tmp["Source_StationNO"].ToString() != "")
                        { E_stationNO = dr_tmp["Source_StationNO"].ToString(); }
                        else { E_stationNO = dr_tmp["Apply_StationNO"].ToString(); }
                    }
                    DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} CT from SoftNetMainDB.[dbo].[DOC4ProductionII] where SUBSTRING(Id,2,2)='{_Fun.Config.ServerId}' and SUBSTRING(DOCNumberNO,1,4)='{dr_DOC4ProductionII["DOCNumberNO"].ToString().Substring(0, 4)}' and PartNO='{dr_DOC4ProductionII["PartNO"].ToString()}' and CT>0");
                    List<double> allCT = new List<double>();
                    if (dt_Efficient != null && dt_Efficient.Rows.Count > 0)
                    {
                        for (int i2 = 0; i2 < dt_Efficient.Rows.Count; i2++)
                        {
                            foreach (DataRow dr2 in dt_Efficient.Rows)
                            {
                                allCT.Add(double.Parse(dr2["CT"].ToString()));
                            }
                        }
                    }
                    else
                    {
                        if (typeTotalTime != 0)
                        { allCT.Add(typeTotalTime); }
                    }
                    if (allCT.Count > 0)
                    {
                        _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, pp_Name, pp_Name, "0", partNO, partNO, dr_DOC4ProductionII["DOCNumberNO"].ToString().Substring(0, 4));
                    }
                }
            }
            #endregion

        }


    }
}
