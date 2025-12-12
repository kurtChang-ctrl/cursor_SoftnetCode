using Base;
using Base.Services;
using BaseApi.Services;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

using SoftNetWebII.Models;
using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.WebSockets;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;

namespace SoftNetWebII.Controllers
{
    public class TagResult_EnterKeyController : Controller
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public TagResult_EnterKeyController(SNWebSocketService websocket, SFC_Common sfc_Common)
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
        public ActionResult Index([FromBody] API_EnterKeyResult input)
        {
            if (input == null)
            {
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"TagResult_EnterKeyController.cs 接收標籤按鈕 keys=NULL", true);
                return StatusCode(200);
            }
            //_Fun.test.Add($"Time={DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")} mac={input.mac} result={input.result}");
            bool isOK = false;
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                if (input.mac != null)
                {
                    DateTime now = DateTime.Now;
                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ReceiveType],[INFO]) VALUES ('{_Str.NewId('L')}','{input.mac}','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','接收按鈕訊號','{input.result.ToString()}')");

                    DataRow dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{input.mac.Trim()}'");
                    if (dr != null)
                    {
                        var verLab = dr["Version"].ToString().Trim();
                        var showValue = dr["ShowValue"].ToString().Trim();
                        var json = "";
                        var ledrgb = "0";
                        var ledstate = "0";
                        string err = "";
                        switch (dr["Type"].ToString())
                        {
                            case "1"://工站按鈕
                            case "4":
                                {
                                    switch (input.result)
                                    {
                                        case 0://停工 or 開工
                                            DataRow dr_M = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                                            List<string> tmp = dr_M["OP_NO"].ToString().Trim().Split(';').ToList();
                                            if (dr_M["State"].ToString() == "1")
                                            {
                                                #region 停工
                                                ledrgb = "0";
                                                if (dr["Type"].ToString() == "4")
                                                {
                                                    //學習模式
                                                    LabelProject keyNull = new LabelProject();
                                                    //err = _SFC_Common.LabelProject_Start_Stop(db, "2", dr["StationNO"].ToString(), dr_M, tmp[0], ref keyNull);
                                                }
                                                else
                                                {
                                                    err = _SFC_Common.ChangeStatus(input.mac.Trim(), $"{dr["StationNO"].ToString().Trim()},2", "TagResult_EnterKey"); //###???ipport暫時放mac ,之後要測試會有何問題
                                                    if (err == "")
                                                    { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set Ledrgb='{ledrgb}',Ledstate=0 where ServerId='{_Fun.Config.ServerId}' and macID='{input.mac.Trim()}'"); }
                                                }

                                                if (err != "")
                                                {
                                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"TagResult_EnterKeyController.cs {dr["StationNO"].ToString()}'站 接收標籤停工按鈕 失敗, err={err}", true);
                                                }
                                                #endregion
                                            }
                                            else
                                            {
                                                #region 開工
                                                if (dr["Type"].ToString() == "4")
                                                {
                                                    //學習模式
                                                    LabelProject keyNull = new LabelProject();
                                                    //err = _WebSocket.LabelProject_Start_Stop(db, "1", dr["StationNO"].ToString(), dr_M, tmp[0], ref keyNull);
                                                }
                                                else
                                                {
                                                    err = _SFC_Common.ChangeStatus(input.mac.Trim(), $"{dr["StationNO"].ToString().Trim()},1", "TagResult_EnterKey");//###???ipport暫時放mac ,之後要測試會有何問題
                                                    if (err == "")
                                                    {
                                                        DataRow dr_Manufacture = db.DB_GetFirstDataByDataRow($"select a.*,b.OrderNO as BWO from SoftNetMainDB.[dbo].[Manufacture] as a,SoftNetMainDB.[dbo].[LabelStateINFO] as b where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{dr["StationNO"].ToString()}' and a.OrderNO!='' and a.PP_Name!='' and a.Config_macID=b.macID");
                                                        if (dr_Manufacture != null && dr_Manufacture["orderNO"].ToString().Trim() != "")
                                                        {
                                                            ledrgb = "ff00";
                                                            DataRow totalData = _Fun.GetAvgCTWTandTotalOutput(db, false, dr["OrderNO"].ToString(), dr["StationNO"].ToString(), dr["IndexSN"].ToString());
                                                            string dis_DetailQTY = "0";
                                                            if (totalData != null)
                                                            {
                                                                dis_DetailQTY = totalData["TotalOutput"].ToString().Trim();
                                                            }
                                                            if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set Ledrgb='{ledrgb}',Ledstate=0 where ServerId='{_Fun.Config.ServerId}' and macID='{input.mac.Trim()}'"))
                                                            {
                                                                //###???暫時拿掉
                                                                //if (dr_Manufacture["BWO"].ToString().Trim() == "" || dr_Manufacture["orderNO"].ToString().Trim() != dr_Manufacture["BWO"].ToString().Trim())
                                                                //{
                                                                string simulationId = "";
                                                                DataRow sfcdr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_Manufacture["OrderNO"].ToString().Trim()}'");

                                                                //###???若不是SID數量要查BOM,CT時間,料號,品名
                                                                #region 查有無需求碼
                                                                string partNO = "";
                                                                string partName = "";
                                                                string typevalue = $"0;";
                                                                int ct = 0;
                                                                int num = int.Parse(sfcdr["Quantity"].ToString());
                                                                if (!sfcdr.IsNull("NeedId") && sfcdr["NeedId"].ToString().Trim() != "")
                                                                {
                                                                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{sfcdr["NeedId"].ToString()}' and SimulationId='{dr_Manufacture["SimulationId"].ToString()}'");
                                                                    if (tmp_dr != null)
                                                                    {
                                                                        simulationId = tmp_dr["SimulationId"].ToString();
                                                                        typevalue = $"2;{simulationId}";
                                                                        num = int.Parse(tmp_dr["NeedQTY"].ToString()) + int.Parse(tmp_dr["SafeQTY"].ToString());
                                                                        ct = int.Parse(tmp_dr["Math_EfficientCT"].ToString());
                                                                        partNO = tmp_dr["PartNO"].ToString();
                                                                        tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{tmp_dr["PartNO"].ToString()}'");
                                                                        if (tmp_dr != null)
                                                                        {
                                                                            partName = tmp_dr["PartName"].ToString().Replace("\"", "＂").Replace("'", "’");
                                                                        }

                                                                    }
                                                                }
                                                                else { typevalue = $"1;{dr_Manufacture["OrderNO"].ToString().Trim()}"; }
                                                                #endregion

                                                                #region 更新標籤
                                                                string isUpdate = "1";
                                                                DataRow dr_LabelStateINFO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{input.mac.Trim()}'");
                                                                if (dr_LabelStateINFO != null && input.mac.Trim() != "")
                                                                {
                                                                    DataRow dr_Staion = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_LabelStateINFO["StationNO"].ToString()}'");
                                                                    string tmp_s = $"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{dr_Manufacture["StationNO"].ToString()};{typevalue};{dr_Manufacture["IndexSN"].ToString()}";
                                                                    var json1 = "";
                                                                    var json_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{sfcdr["OrderNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{partNO}\",\"Text3\":\"品名:\",\"PartName\":\"{partName}\",\"Text4\":\"需求量:\",\"DetailQTY\":\"{dis_DetailQTY}\",\"Text5\":\"CT:\",\"EfficientCT\":\"{ct.ToString()}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                                                                    var writeShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{sfcdr["OrderNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{partNO}\",\"Text3\":\"品名:\",\"PartName\":\"{partName}\",\"Text4\":\"需求量:\",\"QTY\":\"{num.ToString()}\",\"Text5\":\"CT:\",\"EfficientCT\":\"{ct.ToString()}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                                                                    if (dr_LabelStateINFO["Version"].ToString().Trim() != "" && dr_LabelStateINFO["Version"].ToString().Trim().Substring(0, 2) == "42")
                                                                    {
                                                                        json_ShowValue = $"{json_ShowValue},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_Staion["StationNO"].ToString()}\",\"text17\":\"{dr_Staion["StationName"].ToString()}\"";
                                                                        writeShowValue = $"{writeShowValue},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_Staion["StationNO"].ToString()}\",\"text17\":\"{dr_Staion["StationName"].ToString()}\"";
                                                                        json1 = $"\"mac\":\"{input.mac.Trim()}\",\"mappingtype\":744,\"styleid\":52,{json_ShowValue}";
                                                                    }
                                                                    else
                                                                    {
                                                                        json_ShowValue = $"{json_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                                                        writeShowValue = $"{writeShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                                                        json1 = $"\"mac\":\"{input.mac.Trim()}\",\"mappingtype\":71,\"styleid\":48,{json_ShowValue}";
                                                                    }

                                                                    if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                                    {
                                                                        json = $"{json1},\"QTY\":\"{num.ToString()}\",\"ledrgb\":\"{ledrgb}\",\"ledstate\":{ledstate}";
                                                                        _Fun.Tag_Write(db,input.mac.Trim(),"按鍵開工", json);
                                                                    }
                                                                    else { isUpdate = "0"; }
                                                                    if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{writeShowValue}',StationNO='{dr_Manufacture["StationNO"].ToString()}',Type='1',OrderNO='{dr_Manufacture["OrderNO"].ToString()}',IndexSN='{dr_Manufacture["IndexSN"].ToString()}',StoreNO='',StoreSpacesNO='',QTY={dis_DetailQTY},IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{input.mac.Trim()}'"))
                                                                    {

                                                                    }
                                                                }
                                                                #endregion
                                                                //}
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //###??? log
                                                    }
                                                }
                                                #endregion
                                            }
                                            break;
                                        case 1:
                                        case 2:
                                        case 3:
                                        case 5:
                                            //仓储3代，DI接口反馈result=2;
                                            //仓储24代，DI接口反馈result = 3;
                                            //仓储23代，DI接口反馈result=5;
                                            //isOK = Set_SFC_StationDetail(db, dr["StationNO"].ToString());
                                            ButtonSetDef(db, input.mac.Trim(), input.result);
                                            break;
                                    }
                                }
                                break;
                            case "2"://倉庫按鈕
                                {
                                    switch (input.result)
                                    {
                                        case 0://右鍵
                                            {
                                                //###注意3代標籤的版本
                                                string isUpdate = "1";
                                                DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr["StoreNO"].ToString()}'");
                                                var json_ShowValue = $"\"Text1\":\"倉庫編號:\",\"StoreNO\":\"{dr["StoreNO"].ToString()}\",\"Text2\":\"名稱:\",\"StoreName\":\"{tmp["StoreName"].ToString()}\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelStroe/Index/{dr["StoreNO"].ToString()}\",\"Text3\":\"狀態:\",\"State\":\"等待工作中...\",\"text7\":\"\",\"text8\":\"\",\"outtime\":0";
                                                json = $"\"mac\":\"{input.mac.Trim()}\",\"mappingtype\":23,\"styleid\":49,{json_ShowValue}";
                                                if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                {
                                                    json = $"{json},\"ledrgb\":\"0\",\"ledstate\":0";
                                                    _Fun.Tag_Write(db,input.mac.Trim(), "倉庫按鈕", json);
                                                }
                                                else { isUpdate = "0"; }
                                                if (!dr.IsNull("Store_DOC_ID") && dr["Store_DOC_ID"].ToString().Trim() != "")
                                                {
                                                    #region 無儲位,寫入庫存
                                                    DataRow dr_DOC3stockII = null;
                                                    foreach (string s in dr["Store_DOC_ID"].ToString().Trim().Split(';'))
                                                    {
                                                        string[] s2 = s.Split(',');
                                                        dr_DOC3stockII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where Id='{s2[1]}' and DOCNumberNO='{s2[0]}'");
                                                        if (dr_DOC3stockII != null)
                                                        {
                                                            if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                                                            { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={dr_DOC3stockII["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["IN_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}'"); }
                                                            else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={dr_DOC3stockII["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'"); }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (!dr.IsNull("Type2macIDs") && dr["Type2macIDs"].ToString().Trim() != "")
                                                {
                                                    string type2macIDs = dr["Type2macIDs"].ToString().Trim().Replace(",", "','");
                                                    DataTable dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID in ('{type2macIDs}')");
                                                    if (dt != null && dt.Rows.Count > 0)
                                                    {
                                                        if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                        {
                                                            using (HttpClient httpClient = new HttpClient())
                                                            {
                                                                try
                                                                {
                                                                    string json3 = "";
                                                                    foreach (DataRow d in dt.Rows)
                                                                    {
                                                                        if (d["Ledrgb"].ToString().Trim() != "0")
                                                                        {
                                                                            if (json3 == "")
                                                                            { json3 = "[{" + $"\"mac\":\"{d["macID"].ToString()}\",\"outtime\":0,\"ledrgb\":\"0\",\"ledmode\":3,\"buzzer\":0,\"lednum\":255,\"reserve\":\"reserve\"" + "}"; }
                                                                            else
                                                                            { json3 = json3 + ",{" + $"\"mac\":\"{d["macID"].ToString()}\",\"outtime\":0,\"ledrgb\":\"0\",\"ledmode\":3,\"buzzer\":0,\"lednum\":255,\"reserve\":\"reserve\"" + "}"; }

                                                                            #region 計算單據CT,平均,有效,寫入庫存, 寫SFC_StationProjectDetail
                                                                            string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
                                                                            DataRow dr_DOC3stockII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where Id='{d["Id"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}'");
                                                                            if (dr_DOC3stockII != null)
                                                                            {
                                                                                #region 寫入庫存
                                                                                if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                                                                                { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={dr_DOC3stockII["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["IN_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}'"); }
                                                                                else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={dr_DOC3stockII["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'"); }
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
                                                                            #endregion
                                                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType]) VALUES ('{_Str.NewId('L')}','{d["macID"].ToString()}','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位滅燈')");
                                                                        }
                                                                    }
                                                                    if (json3 != "") { json3 += "]"; }
                                                                    string url = $"http://{_Fun.Config.ElectronicTagsURL}/wms/associate/lightTagsLed";
                                                                    var content = new StringContent(json3, Encoding.UTF8, "application/json");
                                                                    HttpResponseMessage response = httpClient.PostAsync(url, content).Result;
                                                                    if (!response.IsSuccessStatusCode)
                                                                    {
                                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位滅燈,發送Fail')");
                                                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"儲位亮燈 傳送電子訊號失敗,請通知管理者", false);    //false here, not mailRoot, or endless roop !!
                                                                    }
                                                                    else
                                                                    {
                                                                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','儲位滅燈,發送OK')");
                                                                    }
                                                                }
                                                                catch (Exception ex)
                                                                {
                                                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"TagResult_EnterKeyController.cs 儲位滅燈 Exception: {ex.Message} {ex.StackTrace}", true);
                                                                }
                                                            }
                                                        }
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set SimulationId='',DOCNumberNO='',Id='',Ledrgb='0',Ledstate=0,IsUpdate='1' where ServerId='{_Fun.Config.ServerId}' and macID in ('{type2macIDs}')");
                                                    }
                                                }
                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{json_ShowValue}',Ledrgb='0',Ledstate=0,Type2macIDs=NULL,IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr["macID"].ToString()}'");
                                            }
                                            break;
                                        case 1:
                                            break;
                                        case 2:
                                            break;
                                        case 3:
                                            break;
                                    }
                                }
                                break;
                            case "3"://儲位按鈕
                                {
                                    if (dr["Id"].ToString().Trim() != "" && dr["DOCNumberNO"].ToString().Trim() != "")
                                    {
                                        string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
                                        DataRow dr_DOC3stockII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where Id='{dr["Id"].ToString()}' and DOCNumberNO='{dr["DOCNumberNO"].ToString()}'");

                                        if (dr_DOC3stockII != null)
                                        {

                                            #region 寫入庫存
                                            if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                                            { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={dr_DOC3stockII["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["IN_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}'"); }
                                            else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={dr_DOC3stockII["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'"); }
                                            #endregion

                                            #region 計算單據CT,平均,有效, 寫SFC_StationProjectDetail
                                            int typeTotalTime = 0;
                                            string writeSQL = "";
                                            if (!dr_DOC3stockII.IsNull("StartTime")) { typeTotalTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC3stockII["StartTime"].ToString()), DateTime.Now); }
                                            else { writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime}{writeSQL} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}' and IsOK='0'");
                                            //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set SimulationId='',DOCNumberNO='',Id='',Ledrgb='0',Ledstate=0 where ServerId='{_Fun.Config.ServerId}' and macID='{input.mac.Trim()}'");

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


                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set SimulationId='',DOCNumberNO='',Id='',Ledrgb='0',Ledstate=0,IsUpdate='1' where ServerId='{_Fun.Config.ServerId}' and macID='{dr["macID"].ToString()}'");

                                        return StatusCode(200);
                                    }
                                }
                                break;
                        }
                    }
                }
            }
            return StatusCode(200);
        }


        private void ButtonSetDef(DBADO db, string mac, int input)
        {
            string goRunFun = "0";
            bool isOK = false;
            DataRow dr_LabelMACs = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelMACs] where ServerId='{_Fun.Config.ServerId}' and macID='{mac}'");
            if (dr_LabelMACs != null)
            {
                switch (input)
                {
                    case 1: goRunFun = dr_LabelMACs["ButtonKey1"].ToString(); break;
                    case 2: goRunFun = dr_LabelMACs["ButtonKey2"].ToString(); break;
                    case 3: goRunFun = dr_LabelMACs["ButtonKey3"].ToString(); break;
                    case 4: goRunFun = dr_LabelMACs["ButtonKey4"].ToString(); break;
                    case 5: goRunFun = dr_LabelMACs["ButtonKey5"].ToString(); break;
                }
            }
            switch (goRunFun)
            {
                case "1"://報工
                    isOK = ReportingWork(db, mac);
                    break;
            }
        }

        private bool ReportingWork(DBADO db, string mac)//報工
        {

            LabelWork keys = new LabelWork();
            DataRow dr_Manufacture = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and Config_macID='{mac}'");
            keys.Station = dr_Manufacture["StationNO"].ToString();
            keys.OrderNO = dr_Manufacture["OrderNO"].ToString();
            keys.IndexSN = dr_Manufacture["IndexSN"].ToString();
            keys.OKQTY = 1;
            keys.FailQTY = 0;
            keys.OPNO = dr_Manufacture["OP_NO"].ToString();
            keys.SimulationId = dr_Manufacture["SimulationId"].ToString();
            DataRow dr_StationNO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
            string message = "";
            string stackTrace = "";
            string ViewBagERRMsg = "";
            bool is_reportOK = _SFC_Common.Reporting_LabelWork(db, dr_StationNO, "系統指派", keys, false, ref message, ref stackTrace, ref ViewBagERRMsg);
            if (!is_reportOK)
            { 
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"自動報工錯誤: {message} {stackTrace}", true);
                return false;
            }
            return true;
        }

    }
}
