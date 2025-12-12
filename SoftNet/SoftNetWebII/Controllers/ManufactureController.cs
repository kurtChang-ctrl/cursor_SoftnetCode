using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
//using DocumentFormat.OpenXml.Drawing.Charts;
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
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;
using static StackExchange.Redis.Role;

namespace SoftNetWebII.Controllers
{
    public class ManufactureController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;

        public ManufactureController(SNWebSocketService websocket, SFC_Common sfc_Common)
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
        public ActionResult Read()
        {
            List<IdStrDto> factoryName = new List<IdStrDto>();
            List<IdStrDto> lineName = new List<IdStrDto>();
            List<IdStrDto> orderNo = new List<IdStrDto>();
            List<string[]> hasPO_List = new List<string[]>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData("SELECT FactoryName FROM SoftNetMainDB.[dbo].[Factory]");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        factoryName.Add(new IdStrDto(dr["FactoryName"].ToString(), dr["FactoryName"].ToString()));
                    }
                }
                dt = db.DB_GetData($"SELECT LineName FROM SoftNetSYSDB.[dbo].[PP_Station] where  ServerId='{_Fun.Config.ServerId}' group by LineName");
                if (dt != null && dt.Rows.Count > 0)
                {
                    //###???將來要加where FactoryName
                    foreach (DataRow dr in dt.Rows)
                    {
                        lineName.Add(new IdStrDto(dr["LineName"].ToString(), dr["LineName"].ToString()));
                    }
                }
                dt = db.DB_GetData($"SELECT OrderNO FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and EndTime is NULL");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        orderNo.Add(new IdStrDto(dr["OrderNO"].ToString(), dr["OrderNO"].ToString()));
                    }
                }

                #region 回傳人員名單
               
                string sql = @$"select a.*,b.Name from SoftNetMainDB.[dbo].[User] as a join SoftNetMainDB.[dbo].[Dept] as b on a.DeptId=b.Id  where a.ServerId='{_Fun.Config.ServerId}' order by b.Name";
                DataTable dt_User = db.DB_GetData(sql);
                if (dt_User != null && dt_User.Rows.Count > 0)
                {
                    foreach (DataRow d in dt_User.Rows)
                    {
                        hasPO_List.Add(new string[] { d["UserNO"].ToString(), d["Name"].ToString(), d["DeptId"].ToString() });
                    }
                }
                #endregion
            }
            ViewBag.OrderNO = orderNo;
            ViewBag.LineName = lineName;
            ViewBag.FactoryName = factoryName;
            ViewBag.OPNO = hasPO_List;
            return View();
        }

        [HttpPost]
        public async Task<string> WEBSETChangeStatus(string ipport, string keys) //改變工站狀態   1=開始,2=停止,3=暫停,4=關站
        {
            string meg = "";
            string[] data = keys.Split(',');
            string status = data[data.Length - 1];
            string sql = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow dr = null;
                DataRow dr_M = null;
                bool isrun = false;
                for (int i = 0; i < (data.Length - 1); i++)
                {
                    isrun = false;
                    dr_M = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");
                    DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_M["SimulationId"].ToString()}'");
                    if (dr_M == null || dr_APS_Simulation == null) { meg = $"{meg}<br>{data[i]} 無法改變工站狀態, 可能無工單設定資料"; continue; }
                    string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();

                    #region 檢查 State合理性
                    if (dr_M["State"].ToString() == status) { continue; }
                    else if (dr_M["State"].ToString() == "2" && status == "3") { meg = $"{meg}<br>{data[i]} 無法設定暫停"; continue; }
                    else if (dr_M["State"].ToString() != "1" && status == "3") { meg = $"{meg}<br>{data[i]} 無法設定暫停"; continue; }
                    else if (status == "4" && data.Length > 2) { meg = $"{meg}<br>關站一次只能選取一站"; continue; }
                    DataRow dr_WO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_M["OrderNO"].ToString()}'");
                    if (dr_WO == null && !bool.Parse(dr_M["Label_ProjectType"].ToString()))
                    {
                        meg = $"{meg}<br>{data[i]} 查無工單,無法設定工站狀態"; continue;
                    }
                    string label_WO = dr_M["OrderNO"].ToString();
                    if (status == "1" && bool.Parse(dr_M["Label_ProjectType"].ToString()) && (dr_M["PP_Name"].ToString() == "" && dr_M["IndexSN"].ToString() == "" && dr_M["OP_NO"].ToString() == ""))
                    { meg = $"{meg}<br>{data[i]} 工站無設定,無法設定啟動"; continue; }
                    else if (status == "1" && !bool.Parse(dr_M["Label_ProjectType"].ToString()) && (dr_M["PP_Name"].ToString() == "" || dr_M["IndexSN"].ToString() == "" || dr_M["OP_NO"].ToString() == "")) 
                    { 
                        meg = $"{meg}<br>{data[i]} 工站無設定,無法設定啟動"; 
                        continue;
                    }
                    #endregion

                    #region for 學習模式版本
                    if (bool.Parse(dr_M["Label_ProjectType"].ToString()))
                    {
                        if (status == "3") { continue; }
                        string[] opNO = dr_M["OP_NO"].ToString().Split(';');
                        //LabelProject tmp = new LabelProject();
                        //string tmp_meg = _SFC_Common.LabelProject_Start_Stop(db, status, data[i], dr_M, opNO[0], ref tmp);
                        //if (tmp_meg!="") { meg = $"{tmp_meg}"; }
                        continue;
                    }
                    #endregion

                    switch(status)
                    {
                        case "1":
                            {
                                string err = _SFC_Common.ChangeStatus(data[i], $"{data[i]},1", "Manufacture");//###???ipport暫時放Station ,之後要測試會有何問題
                                if (err == "")
                                {
                                    dr = db.DB_GetFirstDataByDataRow($"select * from [dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");
                                    if (dr != null)
                                    {
                                        var ledrgb = "0";
                                        var ledstate = "0";
                                        var json = "";

                                        DataRow dr_Manufacture = db.DB_GetFirstDataByDataRow($"select a.*,b.OrderNO as BWO from SoftNetMainDB.[dbo].[Manufacture] as a,SoftNetMainDB.[dbo].[LabelStateINFO] as b where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{data[i]}' and a.OrderNO!='' and a.PP_Name!='' and a.Config_macID=b.macID");
                                        if (dr_Manufacture != null && dr_Manufacture["OrderNO"].ToString().Trim() != "")
                                        {
                                            string macID = dr_Manufacture["Config_macID"].ToString();
                                            ledrgb = "ff00";
                                            DataRow totalData = _Fun.GetAvgCTWTandTotalOutput(db, false, dr["OrderNO"].ToString(), dr["StationNO"].ToString(), dr["IndexSN"].ToString());
                                            string dis_DetailQTY = "0";
                                            if (totalData != null)
                                            {
                                                dis_DetailQTY = totalData["TotalOutput"].ToString().Trim();
                                            }
                                            if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set Ledrgb='{ledrgb}',Ledstate=0 where ServerId='{_Fun.Config.ServerId}' and macID='{macID}'"))
                                            {
                                                //###???暫時拿掉
                                                //if (dr_Manufacture["BWO"].ToString().Trim() == "" || dr_Manufacture["OrderNO"].ToString().Trim() != dr_Manufacture["BWO"].ToString().Trim())
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
                                                DataRow dr_LabelStateINFO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{macID}'");
                                                if (dr_LabelStateINFO != null && macID != "")
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
                                                        json1 = $"\"mac\":\"{macID}\",\"mappingtype\":744,\"styleid\":52,{json_ShowValue}";
                                                    }
                                                    else
                                                    {
                                                        json_ShowValue = $"{json_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                                        writeShowValue = $"{writeShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                                        json1 = $"\"mac\":\"{macID}\",\"mappingtype\":71,\"styleid\":48,{json_ShowValue}";
                                                    }

                                                    if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                    {
                                                        json = $"{json1},\"QTY\":\"{num.ToString()}\",\"ledrgb\":\"{ledrgb}\",\"ledstate\":{ledstate}";
                                                        _Fun.Tag_Write(db,macID,"網頁開工", json);
                                                    }
                                                    else { isUpdate = "0"; }
                                                    if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{writeShowValue}',StationNO='{dr_Manufacture["StationNO"].ToString()}',Type='1',OrderNO='{dr_Manufacture["OrderNO"].ToString()}',IndexSN='{dr_Manufacture["IndexSN"].ToString()}',StoreNO='',StoreSpacesNO='',QTY={dis_DetailQTY},IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{macID}'"))
                                                    {

                                                    }
                                                }
                                                #endregion
                                            }
                                        }

                                    }
                                }
                                else
                                {
                                    meg = $"{meg}<br>{data[i]} 開工失敗,原因:{err}.";
                                }
                            }
                            break;
                        case "2":
                            {
                                var ledrgb = "0";
                                DataRow dr_Manufacture = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");
                                string err = _SFC_Common.ChangeStatus(data[i], $"{data[i]},2", "Manufacture"); //###???ipport暫時放keys.Station ,之後要測試會有何問題
                                if (err == "")
                                {
                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set Ledrgb='{ledrgb}',Ledstate=0 where ServerId='{_Fun.Config.ServerId}' and macID='{dr_Manufacture["Config_MutiWO"].ToString()}'");
                                }
                                else
                                { meg = $"{meg}<br>{data[i]} 開工失敗,原因:{err}."; }
                                //###??? 缺少更新標籤
                            }
                            break;
                        case "4":
                            {
                                bool Is_Station_Config_Store_Type = false;
                                bool isLastStation = false;
                                List<string> station_list = new List<string>();
                                List<string> station_list_NO_StationNO_Merge = new List<string>();
                                if (dr_M["State"].ToString() == "1") { meg = $"{meg}<br>{data[i]} 工站無設定,無法設定啟動"; continue; }
                                if (dr_M["OrderNO"].ToString() == "") { meg = $"{meg}<br>{data[i]} 工站查無工單,無法設定關畢工站."; continue; }
                                for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();
                                string needID = "";
                                #region 檢查下一站是否為委外, 若是改停止動作加關閉
                                if (!_Fun.Config.IsOutPackStationStore && dr_M["SimulationId"].ToString() != "")
                                {
                                    DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_M["SimulationId"].ToString()}'");
                                    if (tmp != null)
                                    {
                                        if (_Fun.Config.OutPackStationName == tmp["Apply_StationNO"].ToString()) { Is_Station_Config_Store_Type = true; }
                                    }
                                }
                                #endregion
                                if (Is_Station_Config_Store_Type)
                                {
                                    var ledrgb = "0";
                                    string err = _SFC_Common.ChangeStatus(data[i], $"{data[i]},2", "Manufacture"); //###???ipport暫時放keys.Station ,之後要測試會有何問題
                                    if (err == "")
                                    {
                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Label_ProjectType='0',OrderNO='',OP_NO='',Master_PP_Name='',PP_Name='',IndexSN=0,Station_Custom_IndexSN='',StationNO_Custom_DisplayName='',State='4',PartNO='',StartTime=NULL,RemarkTimeS=NULL,RemarkTimeE=NULL,EndTime=NULL where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");
                                        #region 更新關閉的電子Tag
                                        DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                        if (dr_tmp != null)
                                        {
                                            string tmp_s = "";
                                            string tmp_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"\",\"Text2\":\"\",\"PartNO\":\"\",\"Text3\":\"\",\"PartName\":\"\",\"Text4\":\"\",\"QTY\":\"\",\"Text5\":\"\",\"EfficientCT\":\"\",\"Text6\":\"\",\"Rate\":\"\",\"Text7\":\"累計量:\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{data[i]};0;;0\",\"outtime\":0";
                                            if (dr_tmp["Version"].ToString().Trim() != "" && dr_tmp["Version"].ToString().Trim().Substring(0, 2) == "42")
                                            {
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");
                                                tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{data[i]}\",\"text17\":\"{dr_tmp["StationName"].ToString()}\"";
                                                tmp_s = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{tmp_ShowValue}";
                                            }
                                            else
                                            {
                                                tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"{data[i]}\",\"text17\":\"\"";
                                                tmp_s = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{tmp_ShowValue},\"ledrgb\":\"0\",\"ledstate\":0";
                                            }
                                            if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                            {
                                                _Fun.Tag_Write(db,dr_M["Config_macID"].ToString(), "網頁關站", tmp_s);
                                            }
                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{tmp_ShowValue}',Ledrgb='0',Ledstate=0,StationNO='{data[i]}',Type='1',OrderNO='',IndexSN='',StoreNO='',StoreSpacesNO='',QTY=0,IsUpdate='1' where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                        }
                                        #endregion
                                    }
                                }
                                else
                                {
                                    dr = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_M["OrderNO"].ToString()}'");//###??? and WOStatus=0  要加是否已關閉判斷
                                    if (dr != null)
                                    {
                                        needID = dr.IsNull("NeedId") ? "" : dr["NeedId"].ToString();
                                        dr = db.DB_GetFirstDataByDataRow($"select a.*,b.IsLastStation FROM SoftNetMainDB.[dbo].[Manufacture] as a, SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] as b where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{data[i]}' and a.OrderNO='{dr_M["OrderNO"].ToString()}' and a.IndexSN='{dr_M["IndexSN"].ToString()}' and a.StationNO=b.StationNO and a.IndexSN=b.IndexSN and a.OrderNO=b.OrderNO and a.PP_Name=b.PP_Name");
                                        if (dr != null)
                                        {
                                            isLastStation = bool.Parse(dr["IsLastStation"].ToString());
                                            #region 判斷是否最後一站
                                            if (isLastStation)
                                            {
                                                //SFC_Common SFC_FUN = new SFC_Common("1", _Fun.Config.Db);
                                                //bool isRun_PP_ProductProcess_Item = true;
                                                //if (db.DB_GetQueryCount($"SELECT * FROM SoftNetSYSDB.[dbo].PP_WO_Process_Item WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='" + dr_M["OrderNO"].ToString() + "'") > 0)
                                                //{ isRun_PP_ProductProcess_Item = false; }
                                                //DataTable dt_WO_Stations = SFC_FUN.Process_ALLSation_RE_Custom(_Fun.Config.ServerId, "1", _Fun.Config.Db, dr_M["PP_Name"].ToString(), "ORDER BY IndexSN, PP_Name ASC", isRun_PP_ProductProcess_Item, dr_M["OrderNO"].ToString());
                                                //if (dt_WO_Stations != null && dt_WO_Stations.Rows.Count > 0)
                                                //{
                                                //    foreach (DataRow d in dt_WO_Stations.Rows)
                                                //    {
                                                //        station_list.Add($"'{d["Station NO"].ToString()}'");
                                                //    }
                                                //}
                                                //SFC_FUN.Dispose();
                                                DataTable dt_WO_Stations = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and NeedId='{needID}'");
                                                if (dt_WO_Stations != null && dt_WO_Stations.Rows.Count > 0)
                                                {
                                                    foreach (DataRow d in dt_WO_Stations.Rows)
                                                    {
                                                        if (!station_list.Contains(d["Source_StationNO"].ToString()))
                                                        {
                                                            station_list.Add(d["Source_StationNO"].ToString());
                                                            if (!d.IsNull("StationNO_Merge"))
                                                            {
                                                                foreach (string s in d["StationNO_Merge"].ToString().Split(','))
                                                                {
                                                                    if (s.Trim() != "" && !station_list.Contains(s))
                                                                    {
                                                                        station_list.Add(s);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                dt_WO_Stations = db.DB_GetData($"select Apply_StationNO from SoftNetSYSDB.[dbo].APS_Simulation where NeedId='{needID}' and PartSN>=0 group by Apply_StationNO");
                                                if (dt_WO_Stations != null && dt_WO_Stations.Rows.Count > 0)
                                                {
                                                    foreach (DataRow d in dt_WO_Stations.Rows)
                                                    {
                                                        station_list_NO_StationNO_Merge.Add($"'{d["Apply_StationNO"].ToString()}'");
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].APS_Simulation where NeedId='{needID}' and SimulationId='{dr_M["SimulationId"].ToString()}'");
                                                if (dr_tmp != null)
                                                {
                                                    station_list.Add($"'{dr_tmp["Apply_StationNO"].ToString()}'");
                                                    station_list_NO_StationNO_Merge.Add($"'{dr_tmp["Apply_StationNO"].ToString()}'");
                                                }
                                            }
                                            #endregion

                                            //###???以下有改 則RUNTimeServer 干涉關公單ㄝ要改

                                            //###???DOCNO暫時寫死BC01
                                            //###???入庫要考慮倉庫最大安置量
                                            #region 半成品 or 成品入庫 與 餘料入庫 Class='4' or Class='5'
                                            DataTable tmp_dt = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Apply_StationNO in ({string.Join(",", station_list_NO_StationNO_Merge)}) and Apply_PP_Name='{dr_M["PP_Name"].ToString()}' and (Class='4' or Class='5') and Source_StationNO is not null");
                                            if (tmp_dt != null)
                                            {
                                                foreach (DataRow d in tmp_dt.Rows)
                                                {
                                                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                    if (tmp_dr != null)
                                                    {
                                                        #region 多生產或多領入庫
                                                        int qty = int.Parse(tmp_dr["Detail_QTY"].ToString()) - (int.Parse(tmp_dr["Next_StationQTY"].ToString()) + int.Parse(tmp_dr["Next_StoreQTY"].ToString()));
                                                        if (qty > 0)
                                                        {
                                                            string tmp_no = "";
                                                            string in_StoreNO = "";
                                                            string in_StoreSpacesNO = "";
                                                            DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}'");
                                                            if (tmp == null)
                                                            {
                                                                tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}'");
                                                                if (tmp == null)
                                                                {
                                                                    #region 查找適合庫儲別
                                                                    _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "BC01");
                                                                    #endregion
                                                                    _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;ManufactureController", ref tmp_no,"系統指派");
                                                                }
                                                                else
                                                                {
                                                                    in_StoreNO = tmp["StoreNO"].ToString();
                                                                    in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                                                    _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;ManufactureController", ref tmp_no, "系統指派");
                                                                }
                                                            }
                                                            else
                                                            {
                                                                in_StoreNO = tmp["StoreNO"].ToString();
                                                                in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                                                _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;ManufactureController", ref tmp_no, "系統指派");
                                                            }
                                                            sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{tmp_no}',Next_StoreQTY+={qty} where SimulationId='{d["SimulationId"].ToString()}'";
                                                            db.DB_SetData(sql);
                                                        }
                                                        #endregion
                                                    }
                                                }
                                            }
                                            #endregion

                                            //###???暫時寫死 EB01
                                            #region 非半成品成品 原物料 餘料退回入庫 Class!='4' and Class!='5'
                                            sql = "";
                                            List<string> sID_list = new List<string>();
                                            if (isLastStation)
                                            { sql = $"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and ((Class!='4' and Class!='5') or NoStation='1')"; }
                                            else
                                            {
                                                DataTable dt_tmp = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and Apply_PP_Name='{dr_M["PP_Name"].ToString()}' and ((Class!='4' and Class!='5') or Source_StationNO is null)");
                                                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                                {
                                                    foreach (DataRow d in dt_tmp.Rows)
                                                    { sID_list.Add($"'{d["SimulationId"].ToString()}'"); }
                                                    sql = $"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and SimulationId in ({string.Join(",", sID_list)})";
                                                }
                                            }
                                            DataTable dt_APS_PartNOTimeNote = db.DB_GetData(sql);
                                            if (dt_APS_PartNOTimeNote != null && dt_APS_PartNOTimeNote.Rows.Count > 0)
                                            {
                                                DataRow tmp_dr = null;
                                                int sQTY = 0;
                                                int useQYU = 0;
                                                foreach (DataRow d in dt_APS_PartNOTimeNote.Rows)
                                                {
                                                    tmp_dr = db.DB_GetFirstDataByDataRow($@"SELECT sum(a.QTY) as Total FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                                                            join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.DOCNumberNO,1,4) and b.DOCType='3' and b.ServerId='{_Fun.Config.ServerId}'
                                                                                            where a.SimulationId='{d["SimulationId"].ToString()}'");
                                                    if (tmp_dr != null && !tmp_dr.IsNull("Total"))
                                                    {
                                                        useQYU = int.Parse(d["Detail_QTY"].ToString()) + int.Parse(d["Next_StationQTY"].ToString()) + int.Parse(d["Next_StoreQTY"].ToString());
                                                        sQTY = int.Parse(tmp_dr["Total"].ToString());
                                                        if ((sQTY - useQYU) > 0)
                                                        {
                                                            useQYU = sQTY - useQYU;//退回量
                                                            int wQTY = useQYU;
                                                            tmp_dt = db.DB_GetData($@"SELECT a.*,c.NeedId,c.SimulationDate FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                                                            join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.DOCNumberNO,1,4) and b.DOCType='3' and b.ServerId='{_Fun.Config.ServerId}'
                                                                                            join SoftNetSYSDB.[dbo].[APS_Simulation] as c on c.SimulationId=a.SimulationId
                                                                                            where a.SimulationId='{d["SimulationId"].ToString()}' order by a.IsOK,a.Id,a.OUT_StoreNO,a.OUT_StoreSpacesNO");
                                                            string docNumberNO = "";
                                                            foreach (DataRow d2 in tmp_dt.Rows)
                                                            {
                                                                if ((int.Parse(d2["QTY"].ToString()) - useQYU) > 0)
                                                                {
                                                                    if (!bool.Parse(d2["IsOK"].ToString()))
                                                                    { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] SET QTY-={useQYU.ToString()} where Id='{d2["Id"].ToString()}' and DOCNumberNO='{d2["DOCNumberNO"].ToString()}'"); }
                                                                    else
                                                                    { _SFC_Common.Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", useQYU, "", d2["Id"].ToString(), $"生產結束餘料退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;ManufactureController", ref docNumberNO, "系統指派"); }
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    if (!bool.Parse(d2["IsOK"].ToString()))
                                                                    { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] SET QTY=0,Remark='生產結束清除用量' where Id='{d2["Id"].ToString()}' and DOCNumberNO='{d2["DOCNumberNO"].ToString()}'"); }
                                                                    else
                                                                    { _SFC_Common.Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", int.Parse(d2["QTY"].ToString()), "", d2["Id"].ToString(), $"生產結束餘料退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;ManufactureController", ref docNumberNO, "系統指派"); }
                                                                    useQYU -= int.Parse(d2["QTY"].ToString());
                                                                    if (useQYU <= 0) { break; }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion

                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsOK='1' where SimulationId='{dr_M["SimulationId"].ToString()}'");
                                        }
                                    }
                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Label_ProjectType='0',OrderNO='',OP_NO='',Master_PP_Name='',PP_Name='',IndexSN=0,Station_Custom_IndexSN='',StationNO_Custom_DisplayName='',State='4',PartNO='',StartTime=NULL,RemarkTimeS=NULL,RemarkTimeE=NULL,EndTime=NULL where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");
                                    db.DB_SetData(@$"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES 
                                                    ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','Manufacture','智慧關站','{dr_M["PP_Name"].ToString()}','{data[i]}','{dr_M["PartNO"].ToString()}','{dr_M["OrderNO"].ToString()}','{dr_M["OP_NO"].ToString()}',{dr_M["IndexSN"].ToString()})");

                                    #region 送Service處理後續
                                    dr = db.DB_GetFirstDataByDataRow($"SELECT RMSName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");
                                    if (isLastStation)
                                    {
                                        status = "5";//關站加關工單

                                        //###??? 要考慮 APS_NeedData與APS_Simulation , 同一個NeedId可能有多個 工單時, 發到Softnet Service的_CloseWO Code會有問題

                                    }
                                    //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                    if (dr != null && _WebSocket.RmsSend(dr["RMSName"].ToString(), 1, $"WebChangeStationStatus,{status},{data[i]},WEBProg,{data[i]},{dr_M["OP_NO"].ToString()},{dr_M["OrderNO"].ToString()},{dr_M["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
                                    {
                                        if (needID != "")
                                        {
                                            if (status == "5")
                                            {
                                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{needID}'");
                                                db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needID}'");
                                            }
                                            else if (status == "4")
                                            {
                                                db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needID}' and SimulationId='{dr_M["SimulationId"].ToString()}'");
                                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{needID}' and SimulationId='{dr_M["SimulationId"].ToString()}'");
                                            }
                                        }

                                        #region 更新電子Tag
                                        if (dr_M["Config_macID"].ToString().Trim() != "")
                                        {
                                            string isUpdate = "1";
                                            if (!_Fun.Is_Tag_Connect) { isUpdate = "0"; }
                                            if (isLastStation)
                                            {
                                                DataRow dr_tmp = null;
                                                string macID = "";
                                                foreach (string s in station_list)
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{s}'");
                                                    if (dr_tmp != null && dr_tmp["State"].ToString().Trim() != "1" && dr_tmp["Config_macID"].ToString().Trim() != "")
                                                    {
                                                        if (dr_tmp["OrderNO"].ToString().Trim() != "" && label_WO != dr_tmp["OrderNO"].ToString().Trim()) { continue; }
                                                        macID = dr_tmp["Config_macID"].ToString().Trim();
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and Type='1' and macID='{macID}'");
                                                        if (dr_tmp != null)
                                                        {
                                                            string tmp_s = "";
                                                            string tmp_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"\",\"Text2\":\"\",\"PartNO\":\"\",\"Text3\":\"\",\"PartName\":\"\",\"Text4\":\"\",\"QTY\":\"\",\"Text5\":\"\",\"EfficientCT\":\"\",\"Text6\":\"\",\"Rate\":\"\",\"Text7\":\"累計量:\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{s};0;;0\",\"outtime\":0";
                                                            if (dr_tmp["Version"].ToString().Trim() != "" && dr_tmp["Version"].ToString().Trim().Substring(0, 2) == "42")
                                                            {
                                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{s}'");
                                                                tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{s}\",\"text17\":\"{dr_tmp["StationName"].ToString()}\"";
                                                                tmp_s = $"\"mac\":\"{macID}\",\"mappingtype\":744,\"styleid\":52,{tmp_ShowValue}";
                                                            }
                                                            else
                                                            {
                                                                tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"{s}\",\"text17\":\"\"";
                                                                tmp_s = $"\"mac\":\"{macID}\",\"mappingtype\":71,\"styleid\":48,{tmp_ShowValue},\"ledrgb\":\"0\",\"ledstate\":0";
                                                            }
                                                            if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                            {
                                                                _Fun.Tag_Write(db,macID, "網頁關站", tmp_s);
                                                            }
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{tmp_ShowValue}',Ledrgb='0',Ledstate=0,StationNO='{s}',Type='1',OrderNO='',IndexSN='',StoreNO='',StoreSpacesNO='',QTY=0,IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{macID}'");
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                                if (dr_tmp != null)
                                                {
                                                    string tmp_s = "";
                                                    string tmp_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"\",\"Text2\":\"\",\"PartNO\":\"\",\"Text3\":\"\",\"PartName\":\"\",\"Text4\":\"\",\"QTY\":\"\",\"Text5\":\"\",\"EfficientCT\":\"\",\"Text6\":\"\",\"Rate\":\"\",\"Text7\":\"累計量:\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{data[i]};0;;0\",\"outtime\":0";
                                                    if (dr_tmp["Version"].ToString().Trim() != "" && dr_tmp["Version"].ToString().Trim().Substring(0, 2) == "42")
                                                    {
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");
                                                        tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{data[i]}\",\"text17\":\"{dr_tmp["StationName"].ToString()}\"";
                                                        tmp_s = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{tmp_ShowValue}";
                                                    }
                                                    else
                                                    {
                                                        tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"{data[i]}\",\"text17\":\"\"";
                                                        tmp_s = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{tmp_ShowValue},\"ledrgb\":\"0\",\"ledstate\":0";
                                                    }
                                                    if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                    {
                                                        _Fun.Tag_Write(db,dr_M["Config_macID"].ToString(), "網頁關站", tmp_s);
                                                    }
                                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{tmp_ShowValue}',Ledrgb='0',Ledstate=0,StationNO='{data[i]}',Type='1',OrderNO='',IndexSN='',StoreNO='',StoreSpacesNO='',QTY=0,IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                                }
                                            }
                                        }
                                        #endregion

                                    }
                                    else
                                    {
                                        meg = $"{meg}<br>{data[i]} 後台服務無作用,請檢查服務是否不正常.";
                                    }
                                    #endregion
                                }
                            }
                            break;
                    }

                }
            }
            return meg;
        }


        [HttpPost]
        public string Select_WO_List(string keys) 
        {
            string re = "";
            string[] data1 = keys.Split(',');
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow dr = null;
                DataRow tmp_dr = null;
                string indexSN = "";
                DataRow dr_tmp = null;
                DataTable dt_tmp = null;
                List<string> list_wo = new List<string>();
                List<string> list_indexSN = new List<string>();
                for (int i = 0; i < data1.Length; i++)
                {
                    dt_tmp = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where IsDEL='0' and (Class='4' or Class='5') and DOCNumberNO!='' and (Source_StationNO='{data1[i]}' or StationNO_Merge like '%{data1[i]},%')");
                    if (dt_tmp != null && dt_tmp.Rows.Count>0)
                    {
                        foreach (DataRow d in dt_tmp.Rows)
                        {
                            if (!list_wo.Contains(d["DOCNumberNO"].ToString())) { list_wo.Add(d["DOCNumberNO"].ToString()); }
                            if (d["Source_StationNO_Custom_DisplayName"].ToString() != "")
                            {
                                if (!list_indexSN.Contains(d["Source_StationNO_Custom_DisplayName"].ToString())) { list_indexSN.Add(d["Source_StationNO_Custom_DisplayName"].ToString()); }
                            }
                            else
                            {
                                if (!list_indexSN.Contains(d["Source_StationNO_IndexSN"].ToString())) { list_indexSN.Add(d["Source_StationNO_IndexSN"].ToString()); }
                            }
                        }
                    }
                }
                if (list_wo.Count > 0)
                {
                    re=string.Join(",", list_wo);
                }
                re = $"{re};";
                if (list_indexSN.Count > 0)
                {
                    re = $"{re}{string.Join(",", list_indexSN)}";
                }
            }
            return re;
        }
            [HttpPost]
        public string SetStationConfig(string keys) //工站設定   ipport,站1,站2,,,;工單;IndexSN;作業員
        {
            string[] data1 = keys.Split(',');
            string[] data2 = data1[data1.Length - 1].Split(';');
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow dr = null;
                DataRow tmp_dr = null;
                string indexSN = "";
                DataRow dr_StationNO = null;
                for (int i = 1; i < (data1.Length - 1); i++)
                {
                    indexSN = data2[1];
                    dr_StationNO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data1[i]}'");
                    dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data1[i]}'");
                    if (dr == null || dr["State"].ToString() == "1")
                    { return $"本工站 {data1[i]} 已運作中 或 查無工站設定 ,請先停止或確認工站基本資料, 才能設定工單."; }
                    string beforePartNO= dr["PartNO"].ToString();
                    string beforeIndexSN = dr["IndexSN"].ToString();
                    string stationNO_Custom_IndexSN = "";
                    string stationNO_Custom_DisplayName = "";
                    string beforePP_Name = dr["PP_Name"].ToString();
                    DataRow sfcdr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{data2[0]}'");
                    if (sfcdr == null) { return $"查無 {data2[0]} 工單."; }
                    if (data2[1].Trim() == "")
                    {
                        DataTable tmp_dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and NeedId='{sfcdr["NeedId"].ToString()}' and Source_StationNO='{data1[i]}'");
                        if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                        {
                            if (tmp_dt.Rows.Count == 1)
                            {
                                indexSN = tmp_dt.Rows[0]["IndexSN"].ToString();
                                stationNO_Custom_IndexSN = tmp_dt.Rows[0]["Station_Custom_IndexSN"].ToString();
                                stationNO_Custom_DisplayName = tmp_dt.Rows[0]["Source_StationNO_Custom_DisplayName"].ToString();
                            }
                            else
                            { return $"{data1[i]} 工單在此站有多重工序作業,故需指定何製程序號."; }
                        }
                        else
                        { return $"{data1[i]} 查無相關製程順序."; }
                    }
                    else
                    {
                        int tmp_i = 0;
                        tmp_dr = null;
                        if (int.TryParse(data2[1].Trim(), out tmp_i))
                        {
                            tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and NeedId='{sfcdr["NeedId"].ToString()}' and Source_StationNO='{data1[i]}' and Station_Custom_IndexSN={data2[1]}");
                        }
                        if (tmp_dr != null)
                        {
                            indexSN = tmp_dr["IndexSN"].ToString();
                            stationNO_Custom_IndexSN = tmp_dr["Station_Custom_IndexSN"].ToString();
                            stationNO_Custom_DisplayName = tmp_dr["Source_StationNO_Custom_DisplayName"].ToString();
                        }
                        else
                        { return $"{data1[i]} 查無相關製程順序."; }
                    }
                    string simulationId = "";
                    string partNO = "";
                    DataRow sId_dr = null;
                    string partName = "";
                    string typevalue = $"0;";
                    int num = 0;
                    int ct = 0;
                    #region 檢查之前是否無報工
                    if (beforePartNO != "")
                    {
                        int avgReportTime = 0;
                        DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data1[i]}' and PartNO='{beforePartNO}' and PP_Name='{beforePP_Name}' and IndexSN='{beforeIndexSN}' and  and ReportTime>10");
                        if (dr_tmp != null && !dr_tmp.IsNull("AVGTime") && dr_tmp["AVGTime"].ToString().Trim() != "")
                        {
                            avgReportTime = int.Parse(dr_tmp["AVGTime"].ToString());
                        }
                        if (avgReportTime > 0)
                        {
                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data1[i]}' and IndexSN={beforeIndexSN} and PartNO='{beforePartNO}' and LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%開工%' order by LOGDateTime desc");
                            if (dr_tmp != null)
                            {
                                DateTime tmp_edate = Convert.ToDateTime(dr_tmp["LOGDateTime"]);
                                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data1[i]}' and IndexSN={beforeIndexSN} and PartNO='{beforePartNO}' and LOGDateTime>'{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%報工%'");
                                if (dr_tmp == null)
                                {
                                    int isARGs10_offset = 15;//###??? 10將來改參數
                                    if ((_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), tmp_edate.AddMinutes(isARGs10_offset), DateTime.Now)) >= avgReportTime)
                                    {
                                        return $"{data1[i]}工站 疑似 前一次料號:{beforePartNO} 未完成報工. 若未報工,請先報工, 否則請先執行 關站設定.";
                                    }
                                }
                            }
                        }
                    }
                    #endregion


                    #region 查有無SID
                    if (!sfcdr.IsNull("NeedId") && sfcdr["NeedId"].ToString().Trim()!="")
                    {
                        sId_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] WHERE NeedId='{sfcdr["NeedId"].ToString()}' and DOCNumberNO = '{data2[0]}' AND Source_StationNO = '{data1[i]}' AND (Source_StationNO_IndexSN={indexSN} or Source_StationNO_Custom_IndexSN='{indexSN}')");
                        if (sId_dr!=null)
                        { 
                            simulationId = sId_dr["SimulationId"].ToString();
                            partNO = sId_dr["PartNO"].ToString();
                            typevalue = $"2;{simulationId}";
                            ct = int.Parse(sId_dr["Math_EfficientCT"].ToString());
                        }
                        else
                        {
                            sId_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] WHERE NeedId='{sfcdr["NeedId"].ToString()}' and DOCNumberNO = '{data2[0]}' AND StationNO_Merge like '%{data1[i]},%' AND (Source_StationNO_IndexSN={indexSN} or Source_StationNO_Custom_IndexSN='{indexSN}')");
                            if (sId_dr != null)
                            {
                                simulationId = sId_dr["SimulationId"].ToString();
                                partNO = sId_dr["PartNO"].ToString();
                                typevalue = $"2;{simulationId}";
                                ct = int.Parse(sId_dr["Math_EfficientCT"].ToString());
                            }
                            else { return $"{data1[i]} 查無相關排程需求碼,設定失敗."; }
                        }
                        sId_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}'");
                        if (sId_dr != null)
                        {
                            partName = sId_dr["PartName"].ToString().Replace("\"", "＂").Replace("'", "’");
                        }
                        sId_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where APS_StationNO='{data1[i]}' AND SimulationId='{simulationId}'");
                        if (sId_dr != null)
                        {
                            num = int.Parse(sId_dr["NeedQTY"].ToString());
                        }
                    }
                    #endregion


                    if (!db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Station_Custom_IndexSN='{stationNO_Custom_IndexSN}',StationNO_Custom_DisplayName='{stationNO_Custom_DisplayName}',StartTime=NULL,RemarkTimeS=NULL,RemarkTimeE=NULL,EndTime=NULL,Label_ProjectType='0',OrderNO='{data2[0]}',IndexSN={indexSN},OP_NO='{data2[2]}',Master_PP_Name='{sfcdr["PP_Name"].ToString()}',PP_Name='{sfcdr["PP_Name"].ToString()}',SimulationId='{simulationId}',PartNO='{partNO}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{data1[i]}'"))
                    {
                        return $"{data1[i]} 設定失敗.";
                    }
                    else
                    {
                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{data2[0]}' where SimulationId='{simulationId}'");
                        /*
                        if (simulationId != "")
                        {
                            DataRow dr_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] WHERE SimulationId='{simulationId}'");
                            if (dr_APS_PartNOTimeNote != null)
                            {
                                if (dr_APS_PartNOTimeNote.IsNull("APS_StationNO") || sfcdr["APS_StationNO"].ToString().Trim() == "" || sfcdr["DOCNumberNO"].ToString().Trim() == "")
                                {
                                    db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET APS_StationNO='{data1[i]}',DOCNumberNO='{data2[0]}' where SimulationId='{simulationId}'");
                                }
                            }
                        }
                        */
                    }
                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','Manufacture','設定工站','','{data1[i]}','{partNO}','{data2[0]}','{data2[2]}',{indexSN})");

                    #region 更新Tag
                    DataRow dr_LabelStateINFO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr["Config_macID"].ToString()}'");
                    if (dr_LabelStateINFO != null && dr["Config_macID"].ToString().Trim() != "")
                    {
                        string tmp_s = $"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{data1[i]};{typevalue};{indexSN}";
                        DataRow totalData = _Fun.GetAvgCTWTandTotalOutput(db, false, sfcdr["OrderNO"].ToString(), data1[i], indexSN);
                        string dis_DetailQTY = "0";
                        if (totalData != null)
                        {
                            dis_DetailQTY = totalData["TotalOutput"].ToString();
                        }
                        string isUpdate = "1";
                        var json1 = "";
                        var json_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{sfcdr["OrderNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{partNO}\",\"Text3\":\"品名:\",\"PartName\":\"{partName}\",\"Text4\":\"需求量:\",\"DetailQTY\":\"{dis_DetailQTY}\",\"Text5\":\"CT:\",\"EfficientCT\":\"{ct.ToString()}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                        var writeShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{sfcdr["OrderNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{partNO}\",\"Text3\":\"品名:\",\"PartName\":\"{partName}\",\"Text4\":\"需求量:\",\"QTY\":\"{num.ToString()}\",\"Text5\":\"CT:\",\"EfficientCT\":\"{ct.ToString()}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                        if (dr_LabelStateINFO["Version"].ToString().Trim() != "" && dr_LabelStateINFO["Version"].ToString().Trim().Substring(0, 2) == "42")
                        {
                            json_ShowValue = $"{json_ShowValue},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_StationNO["StationNO"].ToString()}\",\"text17\":\"{dr_StationNO["StationName"].ToString()}\"";
                            writeShowValue = $"{writeShowValue},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_StationNO["StationNO"].ToString()}\",\"text17\":\"{dr_StationNO["StationName"].ToString()}\"";
                            json1 = $"\"mac\":\"{dr["Config_macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{json_ShowValue}";
                        }
                        else
                        {
                            json_ShowValue = $"{json_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                            writeShowValue = $"{writeShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                            json1 = $"\"mac\":\"{dr["Config_macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{json_ShowValue}";
                        }
                        if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                        {
                            _Fun.Tag_Write(db,dr["Config_macID"].ToString(),"網頁設定工單", $"{json1},\"QTY\":\"{num.ToString()}\",\"ledrgb\":\"0\",\"ledstate\":0");
                        }
                        else { isUpdate = "0"; }
                        if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{writeShowValue}',Ledrgb='0',Ledstate=0,StationNO='{data1[i]}',Type='1',OrderNO='{sfcdr["OrderNO"].ToString()}',IndexSN='{indexSN}',StoreNO='',StoreSpacesNO='',QTY={dis_DetailQTY},IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr["Config_macID"].ToString()}'"))
                        {
                        }

                    }
                    #endregion

                }
            }
            lock (_WebSocket.lock__WebSocketList)
            {
                if (_WebSocket._WebSocketList.ContainsKey(data1[0]))
                {
                    _WebSocket.Send(_WebSocket._WebSocketList[data1[0]].socket, "StationStatusChange");
                }
            }
            return "";
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private ManufactureService EditService()
        {
            return new ManufactureService(Ctrl);
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
