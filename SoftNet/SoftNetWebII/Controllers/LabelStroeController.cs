using Base;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
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
using System.Xml;
using ZXing.QrCode.Internal;

namespace SoftNetWebII.Controllers
{
    public class LabelStroeController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public LabelStroeController( SNWebSocketService websocket, SFC_Common sfc_Common)
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
        public IActionResult Index(string id)//id=倉庫編號;BOOLonToggleMenu
        {
			List<LabelDOC3stockII> re = new List<LabelDOC3stockII>();
            var br = _Fun.GetBaseUser();
            if (id==null || br == null || !br.IsLogin || br.UserNO.Trim() == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            string[] data = id.Split(';');

            ViewBag.StoreID = data[0];
            ViewBag.Station = "";
            ViewBag.ERRMsg = "";
            ViewBag.BOOLonToggleMenu = "";
            if (data.Length >1 && data[1]== "HTML")
            {
                data[1] = "";
                ViewBag.BOOLonToggleMenu = "HTML";
            }

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                bool isRUN = true;
                string sID = "";

				#region 先檢查是否排隊
				DataRow dr_isRUN = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and Type='2' and StoreNO='{data[0]}' and Ledrgb='0'");
				if (db.DB_GetQueryCount($"SELECT * from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{data[0]}' and (Keys='領料_{br.UserNO.Trim()}' or Keys='入料_{br.UserNO.Trim()}')") <= 0)
                {
                    return RedirectToAction("StroeFUN", "/LabelStroe", new { id = data[0] });
                }
                else
                {
                    if (dr_isRUN == null)
                    {
                        isRUN = false;
                    }
                }
                #endregion

                if (!isRUN)
                {
                    ViewBag.ERRMsg= "倉庫排隊中.......";
                    ViewBag.RetuenId = id;
                    return View();
                }
                else
                {
                    DataTable dt = null;
                    string in_NO = "";
                    List<string[]> reMutiData = new List<string[]>();
                    if (data.Length == 1 || (data.Length > 1 && data[1] == ""))
                    {
                        dt = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{data[0]}' and (Keys='領料_{br.UserNO.Trim()}' or Keys='入料_{br.UserNO.Trim()}') order by Keys,StationNO");
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            DataRow dr_wo = null;
                            foreach (DataRow dr in dt.Rows)
                            {
                                string[] key = dr["Value"].ToString().Split(',');//'0={keys.OrderNO},1={keys.IndexSN},2={keys.SimulationId}'
                                dr_wo = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{key[0].Trim()}'");
                                if (dr["Keys"].ToString().Trim().IndexOf("領料_") == 0)
                                {
                                    in_NO = "AC01";//###??? 暫時寫死領料單別 
                                    //###???沒寫上一站可能有半成品先入庫, 此時要領出
                                }
                                else
                                {
                                    in_NO = "BC01";//###??? 暫時寫死入料單別
                                }
                                if (key[2].Trim() != "")
                                {
                                    #region 查對應sID
                                    if (sID == "") { sID = $" and SimulationId in ('{key[2].Trim()}'"; }
                                    else { sID = $"{sID},'{key[2].Trim()}'"; }
                                    //DataTable dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_wo["NeedId"].ToString()}' and Apply_StationNO='{dr["StationNO"].ToString()}' and IndexSN={key[1]}");
                                    //if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                                    //{
                                    //    sID = $"and SimulationId in ('{key[2].Trim()}'";
                                    //    foreach (DataRow d in dt_APS_Simulation.Rows)
                                    //    {
                                    //        sID += $",'{d["SimulationId"].ToString()}'";
                                    //    }
                                    //    if (sID != "") { sID += ")"; }
                                    //}
                                    #endregion
                                    if (db.DB_GetQueryCount($"select Id from SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and (OUT_StoreNO='{data[0]}' or IN_StoreNO='{data[0]}') and SimulationId='{key[2]}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}'") > 0)
                                    { reMutiData.Add(new string[] { dr["Keys"].ToString(), dr["StationNO"].ToString(), dr["Value"].ToString() }); }
                                }
                                else
                                {
                                    if (db.DB_GetQueryCount($"select Id from SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and (OUT_StoreNO='{data[0]}' or IN_StoreNO='{data[0]}') and DOCNumberNO='{key[0].Trim()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}'") > 0)
                                    { reMutiData.Add(new string[] { dr["Keys"].ToString(), dr["StationNO"].ToString(), dr["Value"].ToString() }); }

                                }
                            }
                            if (sID != "") { sID = $"{sID})"; }
                            if (reMutiData.Count > 1)
                            {
                                ViewBag.MutiData = reMutiData;
                                return View();
                            }
                        }
                    }
                    else
                    {
                        //###???這裡應該都不會執行
                        for (int i = 1; i < data.Length; i++)
                        {
                            string[] s2 = data[i].Split(',');
                            if (s2.Length > 1)
                            {
                                DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and Keys='{s2[0]}' and StationNO='{s2[1]}' and StoreNO='{data[0]}'");
                                if (dr_tmp != null)
                                {
                                    reMutiData.Add(new string[] { dr_tmp["Keys"].ToString(), dr_tmp["StationNO"].ToString(), dr_tmp["Value"].ToString() });
                                }
                            }
                        }
                    }
                    //DataTable dt_DOC3stockII = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where Id=''");
                    if (reMutiData.Count <= 0) 
                    {
                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{data[0]}' and (Keys='領料_{br.UserNO.Trim()}' or Keys='入料_{br.UserNO.Trim()}')");
                        ViewBag.ERRMsg = "目前無可工作內容.......";
                        return View(); 
                    }
                    string tmp_StoreNO = "";
                    char id_Type = '1';//定義領/入行為
                    foreach (string[] mutiData in reMutiData)
                    {
                        if (mutiData[0].IndexOf("領料_") == 0)
                        {
                            id_Type = '1';
                            in_NO = "AC01";//###??? 暫時寫死領料單別
                            tmp_StoreNO = $"OUT_StoreNO='{data[0]}'";
                        }
                        else
                        {
                            id_Type = '2';
                            in_NO = "BC01";//###??? 暫時寫死入料單別
                            tmp_StoreNO = $"IN_StoreNO='{data[0]}'";
                        }
                        string[] s = mutiData[2].Trim().Split(',');//'0={keys.OrderNO},1={keys.Index},2={keys.SimulationId}'
                        DataTable re_dt = null;
                        if (s[2].Trim() != "")
                        {
                            if (sID != "")
                            {
                                re_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and {tmp_StoreNO} and SUBSTRING(DOCNumberNO,1,4)='{in_NO}' {sID}");
                                if (re_dt != null && re_dt.Rows.Count > 0)
                                {
                                    foreach (DataRow d2 in re_dt.Rows)
                                    {
                                        re.Add(new LabelDOC3stockII(id_Type,d2["Id"].ToString(), d2["DOCNumberNO"].ToString(), d2["PartNO"].ToString(), d2["Price"].ToString(), d2["Unit"].ToString(), d2["QTY"].ToString(), d2["Remark"].ToString(), d2["SimulationId"].ToString(), d2["IsOK"].ToString(), d2["IN_StoreNO"].ToString(), d2["IN_StoreSpacesNO"].ToString(), d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), d2["ArrivalDate"].ToString()));
                                    }
                                }
                            }
                        }
                        else
                        {
                            re_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and {tmp_StoreNO} and DOCNumberNO='{s[0].Trim()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}'");
                            if (re_dt != null && re_dt.Rows.Count > 0)
                            {
                                foreach (DataRow d2 in re_dt.Rows)
                                {
                                    re.Add(new LabelDOC3stockII(id_Type,d2["Id"].ToString(), d2["DOCNumberNO"].ToString(), d2["PartNO"].ToString(), d2["Price"].ToString(), d2["Unit"].ToString(), d2["QTY"].ToString(), d2["Remark"].ToString(), d2["SimulationId"].ToString(), d2["IsOK"].ToString(), d2["IN_StoreNO"].ToString(), d2["IN_StoreSpacesNO"].ToString(), d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), d2["ArrivalDate"].ToString()));
                                }
                            }
                        }
                    }
                }


                if (re.Count > 0)
                {
                    ViewBag.StoreID = null;
                    if (_Fun.Has_Tag_httpClient)
                    {
                        try
                        {
                            string storeNO = "";
                            string storeSpacesN = "";
                            string isUpdate = "1";
                            var json_ShowValue = "";
                            var json2 = "";
                            List<string> updateID_list = new List<string>();
                            DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_isRUN["StoreNO"].ToString()}'");
                            updateID_list.Add(dr_isRUN["macID"].ToString());
                            #region 倉庫亮燈
                            if (_Fun.Is_Tag_Connect)
                            {
                                json_ShowValue = $"\"Text1\":\"倉庫編號:\",\"StoreNO\":\"{dr_isRUN["StoreNO"].ToString()}\",\"Text2\":\"名稱:\",\"StoreName\":\"{tmp["StoreName"].ToString()}\",\"BarCode\":\"\",\"Text3\":\"狀態:\",\"State\":\"工作中...\",\"text7\":\"\",\"text8\":\"\",\"outtime\":0";
                                json2 = $"\"mac\":\"{dr_isRUN["macID"].ToString()}\",\"mappingtype\":23,\"styleid\":49,{json_ShowValue}";
                                var json_tmp = $"{json2},\"ledrgb\":\"ff00\",\"ledstate\":0";
                                _Fun.Tag_Write(db,dr_isRUN["macID"].ToString(), $"倉庫亮燈,工站領入", json_tmp);
                            }
                            else { isUpdate = "0"; }
                            #endregion

                            #region 儲位亮燈
                            var remacIDs = "";
                            var json3 = "";
                            string Store_DOC_ID = "NULL";
                            using (HttpClient httpClient = new HttpClient())
                            {
                                DateTime now = DateTime.Now;
                                try
                                {
                                    foreach (LabelDOC3stockII d in re)
                                    {
                                        //###???調撥會有問題
                                        if (d.IN_StoreNO != "" && d.IN_StoreSpacesNO != "") { storeNO = d.IN_StoreNO; storeSpacesN = d.IN_StoreSpacesNO; }
                                        else if (d.OUT_StoreNO != "" && d.OUT_StoreNO != "") { storeNO = d.OUT_StoreNO; storeSpacesN = d.OUT_StoreSpacesNO; }

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
                                                //_Fun.Tag_Write(dr_isRUN["macID"].ToString(), json3);
                                            }
                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set SimulationId='{d.SimulationId}',DOCNumberNO='{d.DOCNumberNO}',Id='{d.Id}',Ledrgb='ff0000',Ledstate=0,IsUpdate='1' where ServerId='{_Fun.Config.ServerId}' and macID='{tmp_SNO["Config_macID"].ToString().Trim()}'");
                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType]) VALUES ('{_Str.NewId('L')}','{tmp_SNO["Config_macID"].ToString()}','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位亮燈')");
                                        }
                                        else
                                        {
                                            #region 若無儲位,則記錄在倉庫燈
                                            if (Store_DOC_ID == "NULL") { Store_DOC_ID = $"'{d.DOCNumberNO},{d.Id}"; }
                                            else
                                            { Store_DOC_ID = $"{Store_DOC_ID};{d.DOCNumberNO},{d.Id}"; }
                                            #endregion
                                        }
                                    }
                                    if (Store_DOC_ID != "NULL") { Store_DOC_ID = $"{Store_DOC_ID}'"; }
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

                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{json_ShowValue}',Ledrgb='ff00',Ledstate=0,Type2macIDs='{remacIDs}',Store_DOC_ID={Store_DOC_ID},IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr_isRUN["macID"].ToString()}'");
                            db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{data[0]}' and (Keys='領料_{br.UserNO.Trim()}' or Keys='入料_{br.UserNO.Trim()}')");

                        }
                        catch (Exception ex)
                        {
                            ViewBag.ERRMsg = $"{data[0]} 倉庫 傳送電子訊號失敗,請通知管理者";
                        }
                    }
                }
            }
            return View();
        }
        public IActionResult StroeFUN(string id)//id=倉庫編號
        {
			var br = _Fun.GetBaseUser();
			if (br == null || !br.IsLogin || br.UserNO.Trim() == "")
			{
				ViewBag.ERRMsg = "網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
				ViewBag.StoreNO = "";
				return View("ResuItTimeOUT");
			}
			string store = id;
            string storeName = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{id}'");
                if (tmp!=null)
                {
                    storeName = tmp["StoreName"].ToString();
                }
            }
            ViewBag.StoreNO = store;
            ViewBag.StoreName = storeName;
            ViewBag.Station = "";

            return View();
        }
        
        private bool OpenWindow_EStore(string url,string token, List<API_EStore_Open> req)
        {
            bool re = false;
            if (url == "" || req.Count <= 0) { return re; }
            HttpClient httpClient = new HttpClient();
            try
            {
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
            }
            catch (Exception ex)
            {
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"大正倉燈 控制失敗 Exception: {ex.Message} {ex.StackTrace}", true);
                re = false;
            }
            httpClient.Dispose();

            return re;
        }
        private void GetEStoreURLandToken(string storeNO, ref string url, ref string token)
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
        private string OpenWindow_EStoreII_for_newList(DBADO db,string api_URL, string api_Token, string[] data, char select_Window_Type)//select_Window_Type 1=單據 2=儲位 3=指定
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

                    if (select_Window_Type=='1')
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
                                    req.Add(new API_EStore_Open(int.Parse(zxy[0]), int.Parse(zxy[2]), int.Parse(zxy[1])));
                                }
                                else { re = $"{re}<br />料號:{dr_DOC3stockII["PartNO"].ToString()} 儲位:{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()} 定義有錯, 電子櫃無法開門."; }
                            }
                            catch { re = $"{re}<br />料號:{dr_DOC3stockII["PartNO"].ToString()} 儲位:{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()} 定義有錯, 電子櫃無法開門."; }
                        }
                    }
                    else if (select_Window_Type == '2')
                    {
                        dr_DOC3stockII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and Id='{data[i].Trim()}'");
                        if (dr_DOC3stockII != null)
                        {
                            try
                            {
                                zxy = dr_DOC3stockII["StoreSpacesNO"].ToString().Split('.');
                                if (zxy.Length == 3)
                                {
                                    req.Add(new API_EStore_Open(int.Parse(zxy[0]), int.Parse(zxy[2]), int.Parse(zxy[1])));
                                }
                                else { re = $"{re}<br />料號:{dr_DOC3stockII["PartNO"].ToString()} 儲位:{dr_DOC3stockII["StoreSpacesNO"].ToString()} 定義有錯, 電子櫃無法開門."; }
                            }
                            catch { re = $"{re}<br />料號:{dr_DOC3stockII["PartNO"].ToString()} 儲位:{dr_DOC3stockII["StoreSpacesNO"].ToString()} 定義有錯, 電子櫃無法開門."; }
                        }
                    }
                    else
                    {
                        zxy = data[i].Split('.');
                        if (zxy.Length == 3)
                        {
                            req.Add(new API_EStore_Open(int.Parse(zxy[0]), int.Parse(zxy[2]), int.Parse(zxy[1])));
                        }
                        else { re = $"{re}<br /> 儲位:{data[i]} 定義有錯, 電子櫃無法開門."; }
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
                                        #region  開門 或 亮燈
                                        if (api_URL != "")
                                        { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, '1'); }
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
                                                            if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                                                            { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={out_qty} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["IN_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}'"); }
                                                            else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={out_qty} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'"); }

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
                                                            #endregion
                                                        }
                                                        ViewBag.Report = "已完成單據領取作業.";
                                                    }
                                                    #endregion
                                                }

                                            }
                                        }
                                        if (doc_Id_List.Count>0)
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
                                        { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, '2'); }
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
                    DataTable dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and OUT_StoreNO='{keys.StoreNO}' and IsOK='0' order by DOCNumberNO,OUT_StoreSpacesNO desc,PartNO");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.SIDList = dt;
                    }
                    dt = db.DB_GetData($@"SELECT a.*  FROM SoftNetMainDB.[dbo].[TotalStock] as a
                    join SoftNetMainDB.[dbo].[Material] as b on b.Class!='6' and b.Class!='7' and a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                    where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{keys.StoreNO}' order by PartNO,StoreSpacesNO desc,QTY desc");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.TotalStock_List = dt;
                    }
                }
                else
                {
                    ViewBag.ERRMsg = "網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
                    ViewBag.StoreNO = "";
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
			if (keys== null || keys.StoreNO==null || keys.StoreNO == "" || br == null || !br.IsLogin || br.UserNO.Trim() == "")
			{
				ViewBag.ERRMsg = "網頁 Timeout(網頁使用時間逾時, 請重新登入網頁).";
				ViewBag.StoreNO = "";
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
                                        { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, '1'); }
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
                                                            case "1": DOC1Buy(db, docID, docDOCNumberNO, docQTY, out_qty,ref doc_Id_List); break;
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
                                        { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, '2'); }
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
                                        return View("ResuItTimeOUT");
                                    }

                                }
                            }
                            break;
                        case "C"://設定儲位
                            {
                                if (keys.Station == "" || keys.Select_ID == null || keys.Select_ID.Trim() == "" || keys.TotalStock_List != null)
                                {
                                    keys.ERRMsg = "沒有改變任何儲位項目, 請重新選擇.";
                                }
                                else
                                {
                                    string[] data = keys.Select_ID.Split(';');
                                    if (data.Length > 1)
                                    {
                                        string[] tmp = null;
                                        for (int i = 0; i < data.Length; i++)
                                        {
                                            tmp = data[i].Split(",");
                                            if (tmp.Length > 1)
                                            { db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStock] set StoreSpacesNO='{tmp[1].Trim()}' where Id='{tmp[0]}'"); }
                                        }
                                        ViewBag.Report = "儲位設定完成.";
                                        string del_id = "";
                                        string id = "";
                                        string tmp_PartNO = "";
                                        string tmp_StoreSpacesNO = "";
                                        DataTable tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' order by PartNO,StoreSpacesNO desc");
                                        if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                        {
                                            DataRow dr_TotalStock = null;
                                            tmp_PartNO = tmp_dt.Rows[0]["PartNO"].ToString();
                                            id = tmp_dt.Rows[0]["Id"].ToString();
                                            for (int i = 1; i < tmp_dt.Rows.Count; i++)
                                            {
                                                dr_TotalStock = tmp_dt.Rows[i];
                                                if (tmp_PartNO == dr_TotalStock["PartNO"].ToString())
                                                {
                                                    if (tmp_StoreSpacesNO == dr_TotalStock["StoreSpacesNO"].ToString())
                                                    {
                                                        if (db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStock] set QTY+={dr_TotalStock["QTY"].ToString()} where Id='{id}'"))
                                                        {
                                                            if (del_id == "") { del_id = $"'{dr_TotalStock["Id"].ToString()}'"; }
                                                            else { del_id = $"{del_id},'{dr_TotalStock["Id"].ToString()}'"; }
                                                        }
                                                    }
                                                    else if (tmp_StoreSpacesNO != "" && dr_TotalStock["StoreSpacesNO"].ToString() == "")
                                                    {
                                                        if (db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStock] set QTY+={dr_TotalStock["QTY"].ToString()} where Id='{id}'"))
                                                        {
                                                            if (del_id == "") { del_id = $"'{dr_TotalStock["Id"].ToString()}'"; }
                                                            else { del_id = $"{del_id},'{dr_TotalStock["Id"].ToString()}'"; }
                                                        }
                                                    }
                                                    tmp_StoreSpacesNO = dr_TotalStock["StoreSpacesNO"].ToString();
                                                }
                                                else { tmp_PartNO = dr_TotalStock["PartNO"].ToString(); tmp_StoreSpacesNO = dr_TotalStock["StoreSpacesNO"].ToString(); }
                                                id = dr_TotalStock["Id"].ToString();
                                            }
                                            if (del_id != "")
                                            {
                                                if (db.DB_SetData($"delete from SoftNetMainDB.[dbo].[TotalStock] where Id in ({del_id})"))
                                                {
                                                    del_id = del_id.Replace("'", "");
                                                    tmp = del_id.Split(',');
                                                    ViewBag.Report = $"{ViewBag.Report}<br />並將如下料號明細數量合併並刪除相同料號.";
                                                    foreach (string s in tmp)
                                                    {
                                                        ViewBag.Report = $"{ViewBag.Report}<br />{s}";
                                                    }
                                                }
                                                else
                                                { ViewBag.ERRMsg = $"{ViewBag.ERRMsg}<br />系統發生資料異動嚴重錯誤, 請聯繫管理員."; }

                                            }
                                        }
                                        ViewBag.StoreNO = keys.StoreNO;
                                        return View("ResuItTimeOUT");
                                    }

                                }
                            }
                            break;
                        case "S"://搜尋
                            break;
                    }
                    if (keys.ActionType == "A" || keys.ActionType == "B")
                    {
                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and Keys='入料_{br.UserNO.Trim()}' and StoreNO='{keys.StoreNO}'");
                    }

                    bool hasRun = false;
                    DataTable dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and IN_StoreNO='{keys.StoreNO}' and IsOK='0' order by DOCNumberNO,IN_StoreSpacesNO desc,PartNO");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        if (keys.ActionType == "S")
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
                                hasRun = false;
                                if (keys.ActionType == "S")
                                {
                                    if (keys.S_PartNO!=null&& keys.S_PartNO != "")
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select PartNO from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and (PartNO like '%{keys.S_PartNO}%' or PartName like '%{keys.S_PartNO}%' or Specification like '%{keys.S_PartNO}%')");
                                        if (dr_tmp != null && !hasRun) { hasRun = true; }
                                    }
                                    if (keys.S_Station != null && keys.S_Station != "" && !hasRun)
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT DOCNumberNO FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and APS_StationNO='{keys.S_Station}'");
                                        if (dr_tmp != null && !hasRun) { hasRun = true; }
                                    }
                                }
                                else { hasRun = true; }
                                if (hasRun)
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
                        }
                        else
                        { keys.SIDList = dt; }
                    }
                    dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC1BuyII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' and IsOK='0' order by DOCNumberNO,StoreSpacesNO desc,PartNO");
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
                            keys.SIDList.Columns.Add("EndTime",typeof(DateTime));
                            keys.SIDList.Columns.Add("StartTime",typeof(DateTime));
                            keys.SIDList.Columns.Add("CT");
                        }
                        DataRow dr_dtAdd = null;
                        foreach (DataRow dr in dt.Rows)
                        {
                            hasRun = false;
                            if (keys.ActionType == "S")
                            {
                                if (keys.S_MFNO != null && keys.S_MFNO != "")
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($@"SELECT DOCNumberNO FROM SoftNetMainDB.[dbo].[DOC1Buy] as a
                                            join SoftNetMainDB.[dbo].[MFData] as b on a.MFNO=b.MFNO
                                            where a.ServerId='{_Fun.Config.ServerId}' and a.DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and (a.MFNO like '%{keys.S_MFNO}%' or b.MFName like '%{keys.S_MFNO}%')");
                                    if (dr_tmp != null && !hasRun) { hasRun = true; }
                                }
                                if (keys.S_PartNO != null && keys.S_PartNO != "" && !hasRun)
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select PartNO from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and (PartNO like '%{keys.S_PartNO}%' or PartName like '%{keys.S_PartNO}%' or Specification like '%{keys.S_PartNO}%')");
                                    if (dr_tmp != null && !hasRun) { hasRun = true; }
                                }
                                if (keys.S_Station != null && keys.S_Station != "" && !hasRun)
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT DOCNumberNO FROM SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and Source_StationNO='{keys.S_Station}'");
                                    if (dr_tmp != null && !hasRun) { hasRun = true; }
                                }
                            }
                            else { hasRun = true; }
                            if (hasRun)
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
                    }
                    dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC4ProductionII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' and IsOK='0' order by DOCNumberNO,StoreSpacesNO desc,PartNO");
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
                            hasRun = false;
                            if (keys.ActionType == "S")
                            {
                                if (keys.S_MFNO !=null && keys.S_MFNO != "")
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($@"SELECT DOCNumberNO FROM SoftNetMainDB.[dbo].[DOC4Production] as a
                                            join SoftNetMainDB.[dbo].[MFData] as b on a.MFNO=b.MFNO
                                            where a.ServerId='{_Fun.Config.ServerId}' and a.DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and (a.MFNO like '%{keys.S_MFNO}%' or b.MFName like '%{keys.S_MFNO}%')");
                                    if (dr_tmp != null && !hasRun) { hasRun = true; }
                                }
                                if (keys.S_PartNO != null && keys.S_PartNO != "" && !hasRun)
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select PartNO from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and (PartNO like '%{keys.S_PartNO}%' or PartName like '%{keys.S_PartNO}%' or Specification like '%{keys.S_PartNO}%')");
                                    if (dr_tmp != null && !hasRun) { hasRun = true; }
                                }
                                if (keys.S_Station !=null && keys.S_Station != "" && !hasRun)
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT DOCNumberNO FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}' and APS_StationNO='{keys.S_Station}'");
                                    if (dr_tmp != null && !hasRun) { hasRun = true; }
                                }
                            }
                            else { hasRun = true; }
                            if (hasRun)
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
                    }

                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Store]  where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' and Class='大正倉'");
                    if (dr_tmp == null)
                    {
                        sql = $@"SELECT a.*  FROM SoftNetMainDB.[dbo].[TotalStock] as a
                        join SoftNetMainDB.[dbo].[Material] as b on b.Class!='6' and b.Class!='7' and a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                        where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{keys.StoreNO}' order by PartNO,StoreSpacesNO desc,QTY desc";
                    }
                    else
                    {
                        sql = $@"SELECT a.*  FROM SoftNetMainDB.[dbo].[TotalStock] as a
                        join SoftNetMainDB.[dbo].[Material] as b on (b.Class='6' or b.Class='7') and a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                        where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{keys.StoreNO}' order by PartNO,StoreSpacesNO desc,QTY desc";
                    }
                    dt = db.DB_GetData(sql);
                    //dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' order by PartNO,StoreSpacesNO desc,QTY desc");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.TotalStock_List = dt;
                    }
                    dt = db.DB_GetData($"select StoreSpacesNO from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' group by StoreSpacesNO order by StoreSpacesNO");
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
                    where a.ServerId='{_Fun.Config.ServerId}' and a.StoreNO='{keys.StoreNO}'");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        keys.New_K_List = dt;
                    }
                    dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStockII_Knives] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' and StationNO='' and IsDel='0'");
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
                    dt = db.DB_GetData($"select StoreSpacesNO from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{keys.StoreNO}' group by StoreSpacesNO order by StoreSpacesNO");
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
                                    { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, '2'); }
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
                                    DataRow dr_tmp = null;
                                    List<string> list = new List<string>();
                                    foreach (string id in data)
                                    {
                                        if (id.Trim() == "") { continue; }
                                        else
                                        {
                                            
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetMainDB.[dbo].[TotalStockII_Knives] where KId='{id}'");
                                            if (dr_tmp != null && dr_tmp["StoreSpacesNO"].ToString().Trim()!="")
                                            {
                                                if (!list.Contains(dr_tmp["StoreSpacesNO"].ToString().Trim()))
                                                { list.Add(dr_tmp["StoreSpacesNO"].ToString().Trim()); }
                                            }
                                        }
                                    }
                                    if (list.Count > 0)
                                    {
                                        #region  開門
                                        if (api_URL != "")
                                        { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, list.ToArray(), '3'); }
                                        #endregion
                                    }

                                    foreach (string id in data)
                                    {
                                        if (id.Trim() != "")
                                        {
                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStockII_Knives] set StationNO='{keys.Station}',StoreNO='',StoreSpacesNO='' where KId='{id}'");
                                        }
                                    }
                                    ViewBag.StoreNO = keys.StoreNO;
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
                                    string[] data = keys.StoreSpacesNO.Split(';');

                                    #region  開門
                                    if (api_URL != "")
                                    { ViewBag.ERRMsg = OpenWindow_EStoreII_for_newList(db, api_URL, api_Token, data, '3'); }
                                    #endregion

                                    foreach (string id in data)
                                    {
                                        if (id.Trim() != "")
                                        {
                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStockII_Knives] set StationNO='',StoreNO='{keys.StoreNO}',StoreSpacesNO='{keys.StoreSpacesNO}' where KId='{id}'");
                                        }
                                    }
                                    ViewBag.StoreNO = keys.StoreNO;
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

        private void DOC3stock(DBADO db, string docID,string docDOCNumberNO,int docQTY,int out_qty, ref Dictionary<string, List<DataRow>> doc_Id_List)
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
                string storeNO = "";
                string storeSpacesNO = "";
                if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                { storeNO = dr_DOC3stockII["IN_StoreNO"].ToString(); storeSpacesNO = dr_DOC3stockII["IN_StoreSpacesNO"].ToString(); }
                else { storeNO = dr_DOC3stockII["OUT_StoreNO"].ToString(); storeSpacesNO = dr_DOC3stockII["OUT_StoreSpacesNO"].ToString(); }

                DataRow dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{storeNO}' and StoreSpacesNO='{storeSpacesNO}'");
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
        private void DOC1Buy(DBADO db, string docID, string docDOCNumberNO, int docQTY, int out_qty,ref Dictionary<string, List<DataRow>> doc_Id_List)
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
                DataRow  dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_DOC1BuyII["StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC1BuyII["StoreSpacesNO"].ToString()}'");
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
                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC4ProductionII] set QTY={out_qty.ToString()} where Id='{dr_DOC4ProductionII["Id"].ToString()}' and DOCNumberNO='{dr_DOC4ProductionII["DOCNumberNO"].ToString()}'");
                    }
                    else
                    {
                        //拆單處置
                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC4ProductionII] set QTY={out_qty.ToString()} where Id='{dr_DOC4ProductionII["Id"].ToString()}' and DOCNumberNO='{dr_DOC4ProductionII["DOCNumberNO"].ToString()}'");
                        db.DB_SetData($@"INSERT INTO [dbo].[DOC4ProductionII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[StoreNO],[StoreSpacesNO],ArrivalDate) VALUES 
                                                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{dr_DOC4ProductionII["DOCNumberNO"].ToString()}','{dr_DOC4ProductionII["PartNO"].ToString()}',{dr_DOC4ProductionII["Price"].ToString()},'{dr_DOC4ProductionII["Unit"].ToString()}',{(docQTY - out_qty).ToString()}
                                                                                                ,'{dr_DOC4ProductionII["Remark"].ToString()}','{dr_DOC4ProductionII["SimulationId"].ToString()}','{dr_DOC4ProductionII["IsOK"].ToString()}','{dr_DOC4ProductionII["StoreNO"].ToString()}','{dr_DOC4ProductionII["StoreSpacesNO"].ToString()}','{Convert.ToDateTime(dr_DOC4ProductionII["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                    }
                }
                DataRow dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_DOC4ProductionII["StoreNO"].ToString()}'");
                if (dr_tag["Class"].ToString() == "虛擬倉")
                {
                    var br = _Fun.GetBaseUser();
                    {
                        //###??? 來自 DOC4confirmController code, 以後要 function 化
                        string sql = "";
                        string storeNO = "";
                        string storeSpacesNO = "";
                        DataRow dr_DOC = null;
                        DataRow dr_tmp = null;
                        string s = docID;
                        sql = $"SELECT a.*,b.DOCType  FROM SoftNetMainDB.[dbo].[DOC4ProductionII] as a,SoftNetMainDB.[dbo].[DOCRole] as b where a.Id='{s}' and SUBSTRING(a.DOCNumberNO,1,4)=b.DOCNO and a.IsOK='0' and b.ServerId='{_Fun.Config.ServerId}'";//將來where條件要加DOCType=?
                        dr_DOC = db.DB_GetFirstDataByDataRow(sql);
                        if (dr_DOC != null)
                        {
                            int dr_DOC_QTY = int.Parse(dr_DOC["QTY"].ToString());

                            DataRow d = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC["SimulationId"].ToString()}'");
                            dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_WorkOrder] where NeedId='{d["NeedId"].ToString()}' and EndTime is NULL");
                            //###???搜尋工單的方式要改
                            if (dr_tmp != null)
                            {
                                _SFC_Common.Update_PP_WorkOrder_Settlement(db, dr_tmp["OrderNO"].ToString(), dr_DOC["SimulationId"].ToString());
                            }

                            if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set Detail_QTY+={dr_DOC_QTY.ToString()} where SimulationId='{dr_DOC["SimulationId"].ToString()}'"))
                            {
                                int ct = 0;
                                if (!dr_DOC.IsNull("StartTime")) { ct = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC["StartTime"]), DateTime.Now); }
                                if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC4ProductionII] set IsOK='1',CT={ct},EndTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where Id='{s}' and DOCNumberNO='{dr_DOC["DOCNumberNO"].ToString()}' and IsOK='0'"))
                                {
                                    #region 計算單據PP_EfficientDetail
                                    string partNO = dr_DOC["PartNO"].ToString();
                                    string pp_Name = "";
                                    string E_stationNO = "";
                                    if (dr_DOC["SimulationId"].ToString() != "")
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC["SimulationId"].ToString()}'");
                                        pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                                        if (!dr_tmp.IsNull("Source_StationNO") && dr_tmp["Source_StationNO"].ToString() != "")
                                        { E_stationNO = dr_tmp["Source_StationNO"].ToString(); }
                                        else { E_stationNO = dr_tmp["Apply_StationNO"].ToString(); }
                                    }
                                    DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} CT from SoftNetMainDB.[dbo].[DOC4ProductionII] where SUBSTRING(Id,2,2)='{_Fun.Config.ServerId}' and SUBSTRING(DOCNumberNO,1,4)='{dr_DOC["DOCNumberNO"].ToString().Substring(0, 4)}' and PartNO='{dr_DOC["PartNO"].ToString()}' and CT>0");
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
                                    if (allCT.Count > 0)
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC4Production] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO='{dr_DOC["DOCNumberNO"].ToString()}'");
                                        _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, pp_Name, pp_Name, "0", partNO, partNO, dr_DOC["DOCNumberNO"].ToString().Substring(0, 4), "", dr_tmp["MFNO"].ToString());
                                    }
                                    #endregion
                                }

                                #region 子計畫完成 與 其他報工code不同
                                if (d != null && !bool.Parse(d["IsOK"].ToString()))
                                {
                                    DataRow dr_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d["SimulationId"].ToString()}'");
                                    if (dr_APS_PartNOTimeNote != null)
                                    {
                                        if (((int.Parse(dr_APS_PartNOTimeNote["Detail_QTY"].ToString()) + int.Parse(dr_APS_PartNOTimeNote["Detail_Fail_QTY"].ToString()) + int.Parse(dr_APS_PartNOTimeNote["Next_StoreQTY"].ToString())) - int.Parse(dr_APS_PartNOTimeNote["NeedQTY"].ToString())) >= 0)
                                        {
                                            if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsOK='1' where SimulationId='{d["SimulationId"].ToString()}'"))
                                            {
                                                if (int.Parse(d["PartSN"].ToString()) >= 0)
                                                {
                                                    #region 若本站數量已足夠,修正上一站單據完成日, 並由RUNTimeServer是否要干涉 與 修正上一站APS_WorkTimeNote的工作負荷
                                                    DataTable dt_Befor_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{d["NeedId"].ToString()}' and Apply_StationNO='{d["Source_StationNO"].ToString()}' and IndexSN={d["Source_StationNO_IndexSN"].ToString()} and Master_PartNO='{d["PartNO"].ToString()}' order by PartSN desc");
                                                    if (dt_Befor_APS_Simulation != null && dt_Befor_APS_Simulation.Rows.Count > 0)
                                                    {
                                                        string idS = "";
                                                        string workStationNO_ids = "";
                                                        foreach (DataRow d2 in dt_Befor_APS_Simulation.Rows)
                                                        {
                                                            if (idS == "") { idS = $"'{d2["SimulationId"].ToString()}'"; }
                                                            else { idS = $"{idS},'{d2["SimulationId"].ToString()}'"; }
                                                            if (!d2.IsNull("Source_StationNO") && (d2["Class"].ToString() == "4" || d2["Class"].ToString() == "5"))
                                                            {
                                                                if (workStationNO_ids == "") { workStationNO_ids = $"'{d2["SimulationId"].ToString()}'"; }
                                                                else { workStationNO_ids = $"{workStationNO_ids},'{d2["SimulationId"].ToString()}'"; }
                                                            }
                                                        }
                                                        if (idS != "")
                                                        {
                                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC4ProductionII] set ArrivalDate='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where IsOK='0' and SimulationId in ({idS})");
                                                        }
                                                        if (workStationNO_ids != "")
                                                        {
                                                            db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=1,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{d["NeedId"].ToString()}' and SimulationId in ({workStationNO_ids})");
                                                        }
                                                    }
                                                    #endregion
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion

                                #region 若下1站有先領半成品AC03先報工, 則開BC03(退料)
                                string in_tmp_NO = "AC03";//###??? 暫時寫死 生產件,前站移轉不足,先領倉補
                                string out_tmp_NO = "BC03";//###??? 暫時寫死 生產件,前站移轉不足,先領倉補, 之後補報工退料

                                DataRow dr_AC03 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC["SimulationId"].ToString()}'");
                                if (dr_AC03 != null)
                                {
                                    DataRow dr_BC03 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_AC03["NeedId"].ToString()}' and Apply_PP_Name='{dr_AC03["Apply_PP_Name"].ToString()}' and PartSN={dr_AC03["PartSN"].ToString()}-1 and IndexSN={dr_AC03["IndexSN"].ToString()}+1");
                                    if (dr_BC03 != null)
                                    {
                                        dr_BC03 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr_BC03["SimulationId"].ToString()}'");
                                        if (dr_BC03 != null)
                                        {
                                            int store_tmp = dr_DOC_QTY;
                                            DataRow tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_tmp_NO}'");
                                            if (tmp != null && tmp["qty"].ToString() != "" && int.Parse(tmp["qty"].ToString()) > 0)
                                            {
                                                int tmp_AC03 = int.Parse(tmp["qty"].ToString());
                                                tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{out_tmp_NO}'");
                                                if (tmp != null && tmp["qty"].ToString() != "" && int.Parse(tmp["qty"].ToString()) > 0)
                                                {
                                                    tmp_AC03 -= int.Parse(tmp["qty"].ToString());
                                                    if (tmp_AC03 > 0)
                                                    {
                                                        if (store_tmp > tmp_AC03) { store_tmp = tmp_AC03; }
                                                        if (store_tmp > 0)
                                                        {
                                                            string doc = "";
                                                            tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_tmp_NO}'");
                                                            _SFC_Common.Create_DOC3stock(db, dr_AC03, "", "", tmp["OUT_StoreNO"].ToString(), tmp["OUT_StoreSpacesNO"].ToString(), out_tmp_NO, store_tmp, "", "", $"報工後,生產先領倉量退回 {dr_BC03["APS_StationNO"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref doc, "系統指派", true, true);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (tmp_AC03 > 0)
                                                    {
                                                        if (store_tmp > tmp_AC03) { store_tmp = tmp_AC03; }
                                                        if (store_tmp > 0)
                                                        {
                                                            string doc = "";
                                                            tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_tmp_NO}'");
                                                            _SFC_Common.Create_DOC3stock(db, dr_AC03, "", "", tmp["OUT_StoreNO"].ToString(), tmp["OUT_StoreSpacesNO"].ToString(), out_tmp_NO, store_tmp, "", "", $"報工後,生產先領倉量退回 {dr_BC03["APS_StationNO"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref doc, "系統指派", true, true);
                                                        }
                                                    }
                                                }
                                                sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Next_APS_StationNO='{dr_BC03["APS_StationNO"].ToString()}',Next_StationQTY+={store_tmp.ToString()} where SimulationId='{dr_DOC["SimulationId"].ToString()}'";
                                                if (db.DB_SetData(sql))
                                                { }
                                            }
                                        }
                                    }
                                }
                                #endregion


                                if (d != null && d["PartSN"].ToString() == "0")
                                {
                                    bool isOver = false;
                                    if (dr_DOC["SimulationId"].ToString() != "")
                                    {
                                        #region 判斷數量是否已到齊
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC["SimulationId"].ToString()}'");
                                        string stationNO = dr_tmp["Source_StationNO"].ToString();
                                        string for_Apply_StationNO_BY_Main_Source_StationNO = dr_tmp["Source_StationNO"].ToString();
                                        string pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                                        string indexSN = dr_tmp["Source_StationNO_IndexSN"].ToString();
                                        string orderNO = "";
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_tmp["NeedId"].ToString()}' and Apply_StationNO='{dr_tmp["Source_StationNO"].ToString()}' and IndexSN={dr_tmp["Source_StationNO_IndexSN"].ToString()} and Master_PartNO='{dr_tmp["PartNO"].ToString()}' order by PartSN desc");
                                        if (dr_tmp != null)
                                        {
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr_tmp["SimulationId"].ToString()}'");
                                            if (dr_tmp != null)
                                            {
                                                if (((int.Parse(dr_tmp["Detail_QTY"].ToString()) + int.Parse(dr_tmp["Detail_Fail_QTY"].ToString()) - int.Parse(dr_tmp["NeedQTY"].ToString()))) >= 0)
                                                {
                                                    isOver = true;
                                                    string needID = dr_tmp["NeedId"].ToString();
                                                    DataTable dt_all = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_tmp["NeedId"].ToString()}' order by PartSN desc");
                                                    if (dt_all != null && dt_all.Rows.Count > 0)
                                                    {
                                                        string idS = "";
                                                        foreach (DataRow d2 in dt_all.Rows)
                                                        {
                                                            //###??? XX01暫時寫死
                                                            if (orderNO == "" && d2["DOCNumberNO"].ToString() != "" && d2["DOCNumberNO"].ToString().Length > 4 && d2["DOCNumberNO"].ToString().Substring(0, 4) == "XX01") { orderNO = d2["DOCNumberNO"].ToString(); }
                                                            if (idS == "") { idS = $"'{d2["SimulationId"].ToString()}'"; }
                                                            else { idS = $"{idS},'{d2["SimulationId"].ToString()}'"; }
                                                        }
                                                        if (idS != "")
                                                        {
                                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC4ProductionII] set IsOK='1' where IsOK='0' and SimulationId in ({idS})");
                                                        }
                                                    }
                                                    db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[APS_WorkingPaper] where WorkType='2' and DOCNumberNO='{dr_DOC["DOCNumberNO"].ToString()}'");

                                                    #region 關站處理
                                                    DataRow dr = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}' and PP_Name='{pp_Name}' and OrderNO='{orderNO}' and IndexSN='{indexSN}'");
                                                    if (dr != null)
                                                    {
                                                        bool isLastStation = bool.Parse(dr["IsLastStation"].ToString());
                                                        List<string> station_list_NO_StationNO_Merge = new List<string>();
                                                        #region 判斷是否最後一站
                                                        if (isLastStation)
                                                        {
                                                            #region 不含合併站
                                                            DataTable dt_station_list_NO_StationNO_Merge = db.DB_GetData($"select Apply_StationNO from SoftNetSYSDB.[dbo].APS_Simulation where NeedId='{needID}' and PartSN>=0 group by Apply_StationNO");
                                                            if (dt_station_list_NO_StationNO_Merge != null && dt_station_list_NO_StationNO_Merge.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow d9 in dt_station_list_NO_StationNO_Merge.Rows)
                                                                {
                                                                    station_list_NO_StationNO_Merge.Add($"'{d9["Apply_StationNO"].ToString()}'");
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].APS_Simulation where  NeedId='{needID}' and SimulationId='{dr_DOC["SimulationId"].ToString()}'");
                                                            if (dr_tmp != null)
                                                            {
                                                                station_list_NO_StationNO_Merge.Add($"'{dr_tmp["Apply_StationNO"].ToString()}'");
                                                            }
                                                        }
                                                        #endregion

                                                        if (station_list_NO_StationNO_Merge.Count > 0)
                                                        {
                                                            #region 半成品 or 成品 入庫 與 餘料入庫
                                                            DataTable tmp_dt = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Apply_StationNO in ({string.Join(",", station_list_NO_StationNO_Merge)}) and Apply_PP_Name='{pp_Name}' and (Class='4' or Class='5') and Source_StationNO is not null");
                                                            if (tmp_dt != null)
                                                            {
                                                                foreach (DataRow d9 in tmp_dt.Rows)
                                                                {
                                                                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and SimulationId='{d9["SimulationId"].ToString()}'");
                                                                    if (tmp_dr != null)
                                                                    {
                                                                        int qty = int.Parse(tmp_dr["Detail_QTY"].ToString()) - (int.Parse(tmp_dr["Next_StationQTY"].ToString()) + int.Parse(tmp_dr["Next_StoreQTY"].ToString()));
                                                                        if (qty > 0)
                                                                        {
                                                                            #region 成品入庫
                                                                            //最後一站開入庫單 //###???DOCNO暫時寫死BC01
                                                                            string tmp_no = "";
                                                                            string in_StoreNO = "";
                                                                            string in_StoreSpacesNO = "";
                                                                            DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d9["SimulationId"].ToString()}'");
                                                                            if (tmp == null)
                                                                            {
                                                                                tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d9["PartNO"].ToString()}'");
                                                                                if (tmp == null)
                                                                                {
                                                                                    #region 查找適合庫儲別
                                                                                    _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "BC01");
                                                                                    #endregion
                                                                                    _SFC_Common.Create_DOC3stock(db, d9, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "委外回廠工單結束,生產件入庫", Convert.ToDateTime(d9["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref tmp_no, br.UserNO, true);
                                                                                }
                                                                                else
                                                                                {
                                                                                    in_StoreNO = tmp["StoreNO"].ToString();
                                                                                    in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                                                                    _SFC_Common.Create_DOC3stock(db, d9, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "委外回廠工單結束,生產件入庫", Convert.ToDateTime(d9["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref tmp_no, br.UserNO, true);
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                in_StoreNO = tmp["StoreNO"].ToString();
                                                                                in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                                                                _SFC_Common.Create_DOC3stock(db, d9, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "委外回廠工單結束,生產件入庫", Convert.ToDateTime(d9["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref tmp_no, br.UserNO, true);
                                                                            }
                                                                            sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{tmp_no}',Next_StoreQTY+={qty} where SimulationId='{d9["SimulationId"].ToString()}'";
                                                                            db.DB_SetData(sql);
                                                                            #endregion
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                        #region 非半成品 or 成品 餘料入庫  //###???暫時寫死 EB01
                                                        /* 改由RUNTimeService 12 處理
                                                        sql = "";
                                                        if (isLastStation)
                                                        { sql = $"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and ((Class!='4' and Class!='5') or NoStation='1')"; }
                                                        else
                                                        {
                                                            List<string> tmp_list = new List<string>();
                                                            DataTable dt_tmp = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and Apply_PP_Name='{pp_Name}' and ((Class!='4' and Class!='5') or Source_StationNO is null)");
                                                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow d9 in dt_tmp.Rows)
                                                                { tmp_list.Add($"'{d9["SimulationId"].ToString()}'"); }
                                                                sql = $"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and SimulationId in ({string.Join(",", tmp_list)})";
                                                            }
                                                        }
                                                        DataTable dt_APS_PartNOTimeNote = db.DB_GetData(sql);
                                                        if (dt_APS_PartNOTimeNote != null && dt_APS_PartNOTimeNote.Rows.Count > 0)
                                                        {
                                                            DataRow tmp_dr = null;
                                                            int sQTY = 0;
                                                            int useQYU = 0;
                                                            foreach (DataRow d9 in dt_APS_PartNOTimeNote.Rows)
                                                            {
                                                                tmp_dr = db.DB_GetFirstDataByDataRow($@"SELECT sum(a.QTY) as Total FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                                                            join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.DOCNumberNO,1,4) and b.DOCType='3' and b.ServerId='{_Fun.Config.ServerId}'
                                                                                            where a.SimulationId='{d9["SimulationId"].ToString()}'");
                                                                if (tmp_dr != null && !tmp_dr.IsNull("Total"))
                                                                {
                                                                    useQYU = int.Parse(d9["Detail_QTY"].ToString()) + int.Parse(d9["Next_StationQTY"].ToString()) + int.Parse(d9["Next_StoreQTY"].ToString());
                                                                    sQTY = int.Parse(tmp_dr["Total"].ToString());
                                                                    if ((sQTY - useQYU) > 0)
                                                                    {
                                                                        useQYU = sQTY - useQYU;//退回量
                                                                        int wQTY = useQYU;
                                                                        DataTable tmp_dt = db.DB_GetData($@"SELECT a.*,c.SimulationDate FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                                                            join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.DOCNumberNO,1,4) and b.DOCType='3' and b.ServerId='{_Fun.Config.ServerId}'
                                                                                            join SoftNetSYSDB.[dbo].[APS_Simulation] as c on c.SimulationId=a.SimulationId
                                                                                            where a.SimulationId='{d9["SimulationId"].ToString()}' order by OUT_StoreNO,OUT_StoreSpacesNO,IsOK");
                                                                        string docNumberNO = "";
                                                                        foreach (DataRow d2 in tmp_dt.Rows)
                                                                        {
                                                                            if ((int.Parse(d2["QTY"].ToString()) - useQYU) > 0)
                                                                            {
                                                                                _WebSocket.Create_DOC3stock(db, d9, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", useQYU, "", d2["Id"].ToString(), $"{orderNO}工單結束退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref docNumberNO, br.UserNO, true);
                                                                                break;
                                                                            }
                                                                            else
                                                                            {
                                                                                _WebSocket.Create_DOC3stock(db, d9, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", int.Parse(d2["QTY"].ToString()), "", d2["Id"].ToString(), $"{orderNO}工單結束退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref docNumberNO, br.UserNO, true);
                                                                                useQYU -= int.Parse(d2["QTY"].ToString());
                                                                                if (useQYU <= 0) { break; }
                                                                            }
                                                                        }
                                                                        sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{docNumberNO}',Next_StoreQTY+={wQTY} where SimulationId='{d["SimulationId"].ToString()}'";
                                                                        db.DB_SetData(sql);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        */
                                                        #endregion

                                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsOK='1' where SimulationId='{dr_DOC["SimulationId"].ToString()}'");
                                                    }
                                                    #endregion

                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                    if (!isOver)
                                    {
                                        #region 工單最後一站預開入庫單  
                                        string inOK_NO = "BC01";//###??? 暫時寫死入庫單別
                                        string tmp_no = "";
                                        string in_StoreNO = "";
                                        string in_StoreSpacesNO = "";
                                        DataTable tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 Order by b.StoreOrder");
                                        if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                        {
                                            int tmp_int = int.Parse(dr_DOC["QTY"].ToString());
                                            #region 有計畫Keep量  by StoreOrder順序扣
                                            foreach (DataRow d2 in tmp_dt.Rows)
                                            {
                                                if (in_StoreNO == "")
                                                {
                                                    in_StoreNO = d2["StoreNO"].ToString();
                                                    in_StoreSpacesNO = d2["StoreSpacesNO"].ToString();
                                                }
                                                if (tmp_int > 0)
                                                {
                                                    if (int.Parse(d2["KeepQTY"].ToString()) >= tmp_int)
                                                    {
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY-={tmp_int} where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        _SFC_Common.Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_int, "", "", $"{d["DOCNumberNO"].ToString()} 生產件入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO, true);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        _SFC_Common.Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_01, "", "", $"{d["DOCNumberNO"].ToString()} 生產件入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO, true);
                                                        tmp_int -= tmp_01;
                                                    }
                                                }
                                            }
                                            if (tmp_int > 0)
                                            {
                                                #region 計畫量不夠扣, 入實體倉
                                                if (in_StoreNO == "")
                                                {
                                                    DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' and StoreNO!=''");
                                                    if (tmp != null)
                                                    {
                                                        in_StoreNO = tmp["StoreNO"].ToString();
                                                        in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                                        _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, tmp_int, "", "", $"{d["DOCNumberNO"].ToString()} 生產件入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO, true);
                                                    }
                                                    else
                                                    {
                                                        #region 查找適合入庫儲別
                                                        _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                                        #endregion
                                                        #region 無倉紀錄, 加空倉
                                                        _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, tmp_int, "", "", $"工單:{d["DOCNumberNO"].ToString()} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO, true);
                                                        #endregion
                                                    }
                                                }
                                                #endregion

                                            }
                                            #endregion
                                        }
                                        else
                                        {
                                            #region 查找適合入庫儲別
                                            _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                            #endregion
                                            #region 無倉紀錄, 加空倉
                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                            _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, dr_DOC_QTY, "", "", $"{stationno} 工單:{d["DOCNumberNO"].ToString()} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO, true);
                                            #endregion
                                        }

                                        sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{tmp_no}',Next_StoreQTY+={dr_DOC_QTY} where SimulationId='{d["SimulationId"].ToString()}'";

                                        if (db.DB_SetData(sql))
                                        {
                                            #region 處理工站移轉時間
                                            /*
                                            DataRow tmp = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetLogDB.[dbo].[SFC_StationDetail] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["Source_StationNO"].ToString()}' and IndexSN='{d["Source_StationNO_IndexSN"].ToString()}' order by LOGDateTime desc");
                                            if (tmp != null)
                                            {
                                                int NextStationTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), Convert.ToDateTime(tmp["LOGDateTime"]), DateTime.Now);
                                                if (NextStationTime > 0)
                                                {
                                                    #region 回寫上一站報工移轉時間
                                                    bool isRUNNextStationTime = false;
                                                    DataTable dt_SFC_StationDetail_ChangeLOG = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(tmp["LOGDateTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}' and StationNO='{tmp["StationNO"].ToString()}' and PartNO='{d["PartNO"].ToString()}' order by LOGDateTime");

                                                    if (dt_SFC_StationDetail_ChangeLOG != null && dt_SFC_StationDetail_ChangeLOG.Rows.Count > 0)
                                                    {
                                                        for (int i = 1; i <= dt_SFC_StationDetail_ChangeLOG.Rows.Count; i++)
                                                        {
                                                            DataRow dLOG = dt_SFC_StationDetail_ChangeLOG.Rows[(i - 1)];
                                                            if (i == dt_SFC_StationDetail_ChangeLOG.Rows.Count && int.Parse(dLOG["NextStationTime"].ToString()) != 0)
                                                            {
                                                                db.DB_SetData($@"UPDATE SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] SET NextStationTime={((NextStationTime + int.Parse(dLOG["NextStationTime"].ToString())) / 2)} where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(dLOG["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' 
                                                                                        and LOGDateTimeID='{Convert.ToDateTime(dLOG["LOGDateTimeID"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and StationNO='{dLOG["StationNO"].ToString()}' and PartNO='{dLOG["PartNO"].ToString()}'");

                                                                isRUNNextStationTime = true;
                                                            }
                                                            else
                                                            {
                                                                if (int.Parse(dLOG["NextStationTime"].ToString()) == 0)
                                                                {
                                                                    db.DB_SetData($@"UPDATE SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] SET NextStationTime={NextStationTime} where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(dLOG["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' 
                                                                                        and LOGDateTimeID='{Convert.ToDateTime(dLOG["LOGDateTimeID"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and StationNO='{dLOG["StationNO"].ToString()}' and PartNO='{dLOG["PartNO"].ToString()}'");
                                                                    isRUNNextStationTime = true;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    #endregion

                                                    #region 統計 PP_EfficientDetail 移轉時間
                                                    if (isRUNNextStationTime)
                                                    {
                                                        DataTable dt_Efficient = db.DB_GetData($@"select TOP {_Fun.Config.AdminKey03} NextStationTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["Source_StationNO"].ToString()}' and PartNO='{d["PartNO"].ToString().Trim()}' and NextStationTime!=0 order by StationNO,PartNO");
                                                        if (dt_Efficient != null && dt_Efficient.Rows.Count > 0)
                                                        {
                                                            List<double> allCT = new List<double>();
                                                            foreach (DataRow dr2 in dt_Efficient.Rows)
                                                            {
                                                                allCT.Add(double.Parse(dr2["NextStationTime"].ToString()));
                                                            }
                                                            if (allCT.Count > 0)
                                                            {
                                                                _WebSocket.SfcTimerloopthread_Tick_Efficient(db, allCT, d["Source_StationNO"].ToString(), d["Apply_PP_Name"].ToString(), d["Apply_PP_Name"].ToString(), d["PartNO"].ToString(), d["PartNO"].ToString(), "", data[0]);
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                            }
                                            */
                                            #endregion

                                        }
                                        #endregion
                                    }
                                }

                            }
                        }
                    }
                }
                else
                {
                    dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_DOC4ProductionII["StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC4ProductionII["StoreSpacesNO"].ToString()}'");
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
                        DataRow dr_tmp = null;
                        #region 寫入庫存, 委外加工放入TotalStockII
                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={out_qty.ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC4ProductionII["PartNO"].ToString()}' and StoreNO='{dr_DOC4ProductionII["StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC4ProductionII["StoreSpacesNO"].ToString()}'");

                        dr_tmp = db.DB_GetFirstDataByDataRow($"select Id from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC4ProductionII["PartNO"].ToString()}' and StoreNO='{dr_DOC4ProductionII["StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC4ProductionII["StoreSpacesNO"].ToString()}'");
                        if (dr_tmp != null)
                        {
                            string id = dr_tmp["Id"].ToString();
                            dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[TotalStockII] where ServerId='{_Fun.Config.ServerId}' and Id='{id}' and SimulationId='{dr_DOC4ProductionII["SimulationId"].ToString()}'");
                            if (dr_tmp != null)
                            { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY+={out_qty.ToString()} where ServerId='{_Fun.Config.ServerId}' and Id='{id}' and SimulationId='{dr_tmp["SimulationId"].ToString()}'"); }
                            else
                            {
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{dr_DOC4ProductionII["SimulationId"].ToString()}'");
                                db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII] ([Id],NeedId,[SimulationId],[KeepQTY],ArrivalDate) VALUES ('{id}','{dr_DOC4ProductionII["NeedId"].ToString()}','{dr_DOC4ProductionII["SimulationId"].ToString()}',{out_qty.ToString()},'{Convert.ToDateTime(dr_tmp["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                            }

                        }
                        #endregion

                        int typeTotalTime = 0;
                        if (!dr_DOC4ProductionII.IsNull("StartTime")) { typeTotalTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC4ProductionII["StartTime"].ToString()), DateTime.Now); }
                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC4ProductionII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime} where Id='{dr_DOC4ProductionII["Id"].ToString()}' and DOCNumberNO='{dr_DOC4ProductionII["DOCNumberNO"].ToString()}' and IsOK='0'");
                        string partNO = dr_DOC4ProductionII["PartNO"].ToString();
                        string pp_Name = "";
                        string E_stationNO = "";
                        if (dr_DOC4ProductionII["SimulationId"].ToString().Trim() != "")
                        {
                            dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC4ProductionII["SimulationId"].ToString().Trim()}'");
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
            }
            #endregion

        }

    }
}
