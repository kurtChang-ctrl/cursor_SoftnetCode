using Base.Services;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Models;
using System.Collections.Generic;
using System.Data;
using System;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Wordprocessing;
using BaseApi.Services;
using Newtonsoft.Json;
using System.Linq;
using System.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Security.Cryptography;
using Base;


namespace SoftNetWebII.Controllers
{
    public class CreateBOMController : Controller
    {
        public IActionResult Index(CreateBOMOBJ key)
        {
            string re = "";
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow dr_Material = null;
                DataRow tmp = null;
                DataRow dr_selectBOM = null;
                DataTable dt_tmp = null;
                string s_tmp = "";

                if (key == null || key.COMType == "" || key.MBOMId == "")
                {
                    key = new CreateBOMOBJ();
                    key.COMType = "0";
                    key.SClass = "4";
                    key.MBOMId = "";
                    key.SApply_PartNO = "";
                    key.ERRMsg = "";
                }
                else
                {
                    key.ERRMsg = "";
                    if (key.SelectBOMId != null && key.SelectBOMId != "")
                    {
                        key.MBOMId = key.SelectBOMId;
                        key.SelectBOMId = "";
                        key.COMType = "9";
                        DataRow tmp_Mbom = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{key.MBOMId}'");
                        if (tmp_Mbom != null) 
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{tmp_Mbom["Apply_PartNO"].ToString()}'");
                            key.SClass = tmp["Class"].ToString();
                            key.SApply_PartNO = tmp_Mbom["Apply_PartNO"].ToString(); 
                        }
                        else
                        {
                            key.ERRMsg = $"選擇已存在的BOM母件料號, 非系統轉換碼, 請選擇正確的轉換碼.";
                            key.MBOMId = "";
                            key.SelectBOMId = "";
                            key.COMType = "0";
                        }
                    }
                }

                #region BOM母件清單 NeedIdList
                if (key.NeedIdList == null || key.NeedIdList.Count == 0)
                {
                    if (key.NeedIdList == null) { key.NeedIdList = new List<string>(); }
                    dt_tmp = db.DB_GetData($"select a.Id,a.Apply_PartNO,a.Version,b.PartName,b.Specification,b.Class from SoftNetMainDB.[dbo].[BOM] as a,SoftNetMainDB.[dbo].[Material] as b where a.ServerId='{_Fun.Config.ServerId}' and b.ServerId='{_Fun.Config.ServerId}' and a.Main_Item='1' and a.Apply_PartNO=b.PartNO group by a.Id,a.Apply_PartNO,a.Version,b.PartName,b.Specification,b.Class order by b.Class,a.Apply_PartNO");
                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                    {
                        string partNOClass = "";
                        key.NeedIdList.Add("\x01\x01\x01\x01\x01");
                        foreach (DataRow d in dt_tmp.Rows)
                        {
                            if (d["Class"].ToString() == "4") { partNOClass = "半成品"; }
                            else if (d["Class"].ToString() == "5") { partNOClass = "成品"; }
                            else { partNOClass = "錯誤BOM"; }
                            key.NeedIdList.Add($"{d["Id"].ToString()}\x01{d["Apply_PartNO"].ToString()}\x01{d["PartName"].ToString()}\x01{d["Specification"].ToString()}\x01{partNOClass}\x01{d["Version"].ToString()}");
                        }
                    }
                }
                #endregion

                #region 料件Class清單
                if (key.HasClass_List == null || key.HasClass_List.Count == 0)
                {
                    if (key.HasClass_List == null) { key.HasClass_List = new List<string[]>(); }
                    key.HasClass_List = new List<string[]>(); key.HasClass_List.Add(new string[] { "", "" });
                    key.HasClass_List.Add(new string[] { "1", "原物料" });
                    key.HasClass_List.Add(new string[] { "2", "採購件" });
                    key.HasClass_List.Add(new string[] { "3", "委外件" });
                    key.HasClass_List.Add(new string[] { "4", "半成品" });
                    key.HasClass_List.Add(new string[] { "5", "成品" });
                    key.HasClass_List.Add(new string[] { "6", "刀具" });
                    key.HasClass_List.Add(new string[] { "7", "治工具" });
                }
                #endregion

                #region 料件清單
                if (key.HasPartNO_List == null || key.HasPartNO_List.Count == 0)
                {
                    if (key.HasPartNO_List == null) { key.HasPartNO_List = new List<string>(); }
                    key.HasPartNO_List = new List<string>(); key.HasPartNO_List.Add("\x01\x01\x01");
                    if (key.SClass != "") { s_tmp = $" and Class='{key.SClass}'"; }
                    dt_tmp = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' {s_tmp} order by Class,PartNO");
                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                    {
                        foreach (DataRow d in dt_tmp.Rows)
                        {
                            key.HasPartNO_List.Add($"{d["PartNO"].ToString()}\x01{d["PartName"].ToString()}\x01{d["Specification"].ToString()}\x01{d["Class"].ToString()}");
                        }
                    }
                }
                #endregion

                #region 廠商清單
                if (key.HasMFNO_List == null || key.HasMFNO_List.Count == 0)
                {
                    if (key.HasMFNO_List == null) { key.HasMFNO_List = new List<string[]>(); }
                    key.HasMFNO_List = new List<string[]>(); key.HasMFNO_List.Add(new string[] { "", "" });
                    dt_tmp = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' order by MFNO");
                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                    {
                        foreach (DataRow d in dt_tmp.Rows)
                        {
                            key.HasMFNO_List.Add(new string[] { d["MFNO"].ToString(), d["MFName"].ToString() });
                        }
                    }
                }
                #endregion

                #region 可用工站
                s_tmp = "";
                if (key.HasStationNO_List == null || key.HasStationNO_List.Count == 0)
                {
                    if (key.HasStationNO_List == null) { key.HasStationNO_List = new List<string[]>(); }
                    key.HasStationNO_List = new List<string[]>(); key.HasStationNO_List.Add(new string[] { "", "" });
                    dt_tmp = db.DB_GetData($"SELECT StationNO,StationName FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' order by StationNO");
                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                    {
                        foreach (DataRow d in dt_tmp.Rows)
                        {
                            key.HasStationNO_List.Add(new string[] { d["StationNO"].ToString(), d["StationName"].ToString() });
                        }
                    }
                }
                #endregion

                if (key.ERRMsg == "")
                {
                    switch (key.COMType)
                    {
                        case "0":
                            key.SPartNO = "";
                            key.SPartName = "";
                            break;
                        case "1"://建立母階料號
                            {
                                #region 檢查
                                if (key.SPP_Name == null || key.SPP_Name.Trim() == "") { key.ERRMsg = "製程名稱不能為空白"; key.COMType = "0"; goto break_FUN; }
                                tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_PP_Name='{key.SPP_Name}' and Apply_PartNO='{key.SPartNO}'");
                                if (tmp != null) { key.ERRMsg = $"已有重複料件與製程的名稱, 請另外重新命名製程名稱, 或修改已存在的料件BOM表."; key.COMType = "0"; goto break_FUN; }

                                if (key.SApply_StationNO == null || key.SApply_StationNO == "") { key.ERRMsg = $"{key.ERRMsg}主生產工站不能為空白."; key.COMType = "0"; goto break_FUN; }
                                dr_Material = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{key.SApply_StationNO}'");
                                if (dr_Material == null) { key.ERRMsg = "查無主生產工站編號, 請重新輸入."; key.COMType = "0"; goto break_FUN; }
                                if (key.SPartNO == null || key.SPartNO.Trim() == "") { key.ERRMsg = "無母件料號不能為空白"; key.COMType = "0"; goto break_FUN; }
                                dr_Material = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{key.SPartNO}'");
                                if (dr_Material == null) { key.ERRMsg = "查無母件料號, 請重新輸入."; key.COMType = "0"; goto break_FUN; }
                                else { if (dr_Material["Class"].ToString() != "4" && dr_Material["Class"].ToString() != "5") { key.ERRMsg = $"母階料號不能為非生產類型的料號."; key.COMType = "0"; goto break_FUN; } }
                                string mFNO = "";
                                if (key.SApply_StationNO != _Fun.Config.OutPackStationName) { key.MFNO = ""; }
                                else
                                {
                                    if (key.MFNO != null && key.MFNO != "")
                                    {
                                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' and MFNO='{key.MFNO}'");
                                        if (tmp == null) { key.ERRMsg = "查無委外加工編號"; key.COMType = "0"; goto break_FUN; }
                                        else { mFNO = key.MFNO; }
                                    }
                                    else { key.ERRMsg = $"工作站為委外加工, 須指定一個廠商編號."; key.COMType = "0"; goto break_FUN; }
                                }

                                string stationNO_Merge = "NULL";
                                if (key.SStationNO_IndexSN_Merge != null && key.SStationNO_IndexSN_Merge != "")
                                {
                                    foreach (string s in key.SStationNO_IndexSN_Merge.Split(','))
                                    {
                                        if (s.Trim() == "") { continue; }
                                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{s}'");
                                        if (tmp == null) { key.ERRMsg = $"{key.ERRMsg}<br />查無 {s} 此共用生產站編號"; }
                                    }
                                    if (key.SStationNO_IndexSN_Merge.Contains($"{key.SApply_StationNO},")) { key.ERRMsg = $"{key.ERRMsg}<br />主生產工站 與 共用生產站 有重複."; }
                                    stationNO_Merge = $"'{key.SStationNO_IndexSN_Merge}'";
                                }
                                string outPackType = "0";
                                if (key.SApply_StationNO == _Fun.Config.OutPackStationName)
                                {
                                    outPackType = "1";
                                    if (key.SStationNO_IndexSN_Merge != null && key.SStationNO_IndexSN_Merge != "") { key.ERRMsg = $"{key.ERRMsg}<br />若為委外加工, 則不能有共用生產站"; }
                                }
                                #endregion

                                if (key.ERRMsg != "") { key.COMType = "0"; goto break_FUN; }
                                else
                                {
                                    #region new BOM,BOMII
                                    key.MBOMId = _Str.NewId('Z');
                                    key.SApply_PartNO = key.SPartNO;
                                    db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[ServerId],[Apply_PartNO],PartNO,[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],UpdateTime,IsChackQTY,IsChackIsOK,OutPackType,Station_DIS_Remark,StationNO_Merge,MFNO,IsShare_PP_Name) VALUES
                                            ('{key.MBOMId}','{_Fun.Config.ServerId}','{key.SApply_PartNO}','{key.SPartNO}','1','{DateTime.Now.ToString("yyyy/MM/dd")}','{DateTime.Now.AddYears(10).ToString("yyyy/MM/dd")}','1.0000','{key.SPP_Name}','{key.SApply_StationNO}','0',1,'','{key.SStationNO_Custom_DisplayName}','{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','{key.SIsChackQTY.ToString()}','{key.SIsChackIsOK.ToString()}','{outPackType}','{key.SStation_DIS_Remark}',{stationNO_Merge},'{mFNO}','0')");
                                    db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,[Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class],[AttritionRate]) VALUES
                                            ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{key.MBOMId}',1,'{key.SPartNO}',{key.SBOMQTY},'{dr_Material["Class"].ToString()}',0)");
                                    #endregion
                                    key.COMType = "9";
                                    re = RecursiveBOM(db, key);
                                }
                            }
                            break;
                        case "2"://顯示生產件修改畫面
                            {
                                dr_selectBOM = db.DB_GetFirstDataByDataRow($"select a.*,b.PartName,b.Class from SoftNetMainDB.[dbo].[BOM] as a,SoftNetMainDB.[dbo].[Material] as b where a.ServerId='{_Fun.Config.ServerId}' and b.ServerId='{_Fun.Config.ServerId}' and a.Id='{key.BOMId}' and a.PartNO=b.PartNO");
                                if (dr_selectBOM != null)
                                {
                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and BOMId='{key.BOMId}' and PartNO='{dr_selectBOM["PartNO"].ToString()}'");
                                    if (tmp != null) { key.SBOMQTY = int.Parse(tmp["BOMQTY"].ToString()); }
                                    else { key.SBOMQTY = 1; }
                                    key.BOMId = dr_selectBOM["Id"].ToString();
                                    key.IndexSN = int.Parse(dr_selectBOM["IndexSN"].ToString());
                                    key.SPartNO = dr_selectBOM["PartNO"].ToString();
                                    key.SPartName = dr_selectBOM["PartName"].ToString();
                                    if (key.SClass == null || key.SClass == "") { key.SClass = dr_selectBOM["Class"].ToString(); }
                                    else
                                    {
                                        if (key.SClass != dr_selectBOM["Class"].ToString())
                                        {
                                            key.SPartNO = ""; key.SPartName = "";
                                        }
                                    }
                                    key.SApply_StationNO = dr_selectBOM["Apply_StationNO"].ToString();
                                    key.SPP_Name = dr_selectBOM["Apply_PP_Name"].ToString();
                                    key.MFNO = dr_selectBOM["MFNO"].ToString();
                                    //key.Fun_S_List = "E 2";
                                    key.SStationNO_Custom_DisplayName = dr_selectBOM["StationNO_Custom_DisplayName"].ToString();
                                    key.SStation_DIS_Remark = dr_selectBOM["Station_DIS_Remark"].ToString();
                                    key.SIsChackIsOK = bool.Parse(dr_selectBOM["IsChackIsOK"].ToString()) ? 1 : 0;
                                    key.SIsChackQTY = bool.Parse(dr_selectBOM["IsChackQTY"].ToString()) ? 1 : 0;
                                    if (!dr_selectBOM.IsNull("StationNO_Merge"))
                                    {
                                        key.SStationNO_IndexSN_Merge = dr_selectBOM["StationNO_Merge"].ToString();
                                    }
                                    else { key.SStationNO_IndexSN_Merge = ""; }
                                }
                            }
                            break;
                        case "3"://顯示非生產件修改畫面
                            {
                                //key.HasClass_List.Clear();
                                //key.HasClass_List.Add(new string[] { "1", "原物料" });
                                //key.HasClass_List.Add(new string[] { "2", "採購件" });
                                //key.HasClass_List.Add(new string[] { "3", "委外件" });
                                dr_selectBOM = db.DB_GetFirstDataByDataRow($"select a.BOMQTY,b.PartNO,b.PartName,b.Class from SoftNetMainDB.[dbo].[BOMII] as a,SoftNetMainDB.[dbo].[Material] as b where b.ServerId='{_Fun.Config.ServerId}' and a.Id='{key.BOMId}' and a.PartNO=b.PartNO");
                                if (dr_selectBOM != null)
                                {
                                    if (key.SClass != null && key.SClass != dr_selectBOM["Class"].ToString())
                                    { key.SPartNO = ""; key.SPartName = ""; }
                                    else
                                    {
                                        key.SPartNO = dr_selectBOM["PartNO"].ToString();
                                        key.SPartName = dr_selectBOM["PartName"].ToString();
                                        key.SBOMQTY = int.Parse(dr_selectBOM["BOMQTY"].ToString());
                                        key.SClass = dr_selectBOM["Class"].ToString();
                                    }
                                }
                            }
                            break;
                        case "4"://子階 增,刪
                            {
                                string type1 = "";//2=生產件,3=非生產件
                                string type2 = "";//功能
                                DataRow tmp_MBOM = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{key.MBOMId}'");
                                if (tmp_MBOM == null)
                                { key.ERRMsg = $"查無系統BOM編號, 目前無法異動, 請重新查詢母件BOM表."; }
                                else
                                {
                                    type1 = key.Fun_S_List.Substring(2); type2 = key.Fun_S_List.Substring(0, 1);
                                    if (key.Fun_S_List != "")
                                    {
                                        int indexNO = -1;
                                        #region 查找對應BOMM或BOMII資料 dr_selectBOM
                                        if (type1 == "2" && type2 != "B")
                                        {
                                            dr_selectBOM = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{key.BOMId}'");
                                            indexNO = int.Parse(dr_selectBOM["IndexSN"].ToString());
                                        }
                                        else
                                        {
                                            if (type2 != "B")
                                            { dr_selectBOM = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and Id='{key.BOMId}'"); }
                                        }
                                        #endregion

                                        dr_Material = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{key.SPartNO}'");
                                        if (dr_Material == null || (dr_selectBOM == null && type2 != "B")) { key.ERRMsg = "料件編號 或 BOM表, 有錯誤."; }
                                        else
                                        {
                                            #region 確認可以異動
                                            tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Apply_PP_Name='{tmp_MBOM["Apply_PP_Name"].ToString()}' and PartNO='{tmp_MBOM["PartNO"].ToString()}' and (State='1' or State='2' or State='6')");
                                            if (tmp != null)
                                            {
                                                if (type1 == "2" && type2 == "D") 
                                                { key.ERRMsg = $"料號:{tmp_MBOM["PartNO"].ToString()} 製程:{tmp_MBOM["Apply_PP_Name"].ToString()} 已有在模擬中或生產中, 目前無法異動BOM的結構, 請至干涉修改."; }
                                            }
                                            #endregion
                                            if (key.ERRMsg=="")
                                            {
                                                if (type2 == "B")
                                                {
                                                    #region 從製程新增 非生產件
                                                    int sn = db.DB_GetQueryCount($"SELECT sn FROM SoftNetMainDB.[dbo].[BOMII] where BOMId='{key.BOMId}' order by sn desc");
                                                    db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,[Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class],[AttritionRate]) VALUES
                                                        ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{key.BOMId}',({sn + 1}),'{key.SPartNO}',{key.SBOMQTY},'{dr_Material["Class"].ToString()}',0)");
                                                    #region 刪除虛擬BOMII 的主BOM表料件 
                                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{key.BOMId}'");
                                                    if (tmp != null)
                                                    {
                                                        string befor_PartNO = tmp["PartNO"].ToString();
                                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_PP_Name='{tmp["Apply_PP_Name"].ToString()}' and Apply_PartNO='{tmp["Apply_PartNO"].ToString()}' and IndexSN<{tmp["IndexSN"].ToString()} order by IndexSN desc");
                                                        if (tmp != null)
                                                        {
                                                            //tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and BOMId='{tmp["Id"].ToString()}' and PartNO!='{tmp["PartNO"].ToString()}' and (Class=4 or Class=5)");
                                                            //if (tmp != null)
                                                            //{ db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and BOMId='{key.BOMId}' and PartNO='{befor_PartNO}'"); }
                                                            db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and BOMId='{key.BOMId}' and PartNO='{befor_PartNO}'");
                                                        }
                                                    }
                                                    #endregion
                                                    #endregion
                                                }
                                                else if (type2 == "C")
                                                {
                                                    #region 從製程新增 生產件
                                                    if (type1 == "2")
                                                    {
                                                        if (dr_Material["Class"].ToString() != "4" && dr_Material["Class"].ToString() != "5")
                                                        {
                                                            key.ERRMsg = $"新增製程 與 製程主料, 不能設定料件為非生產件.";
                                                        }
                                                        else
                                                        {
                                                            if (key.SPP_Name != null && key.SPP_Name != "" && key.SApply_StationNO != "" && indexNO >= 1)
                                                            {
                                                                #region 檢查
                                                                if (key.SApply_StationNO != "")
                                                                {
                                                                    DataRow dr_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{key.SApply_StationNO}'");
                                                                    if (dr_Station == null) { key.ERRMsg = "查無主生產工站編號, 請重新輸入."; }
                                                                }
                                                                else { key.ERRMsg = $"{key.ERRMsg}<br />主生產工站編號, 須指定一個廠內工作站 或 廠外委外作業站."; }
                                                                if (key.SApply_StationNO != _Fun.Config.OutPackStationName)
                                                                { key.MFNO = ""; }
                                                                if (key.MFNO != null && key.MFNO != "")
                                                                {
                                                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' and MFNO='{key.MFNO}'");
                                                                    if (tmp == null) { key.ERRMsg = $"{key.ERRMsg}<br />查無委外加工編號"; }
                                                                }
                                                                if (key.SApply_StationNO == _Fun.Config.OutPackStationName && (key.MFNO == null || key.MFNO == ""))
                                                                {
                                                                    key.ERRMsg = $"{key.ERRMsg}<br />工作站為委外加工, 須給指定一個廠商編號.";
                                                                }
                                                                string indexSN_Merge = "NULL";
                                                                if (key.SStationNO_IndexSN_Merge != null && key.SStationNO_IndexSN_Merge != "")
                                                                {
                                                                    foreach (string s in key.SStationNO_IndexSN_Merge.Split(','))
                                                                    {
                                                                        if (s == "") { continue; }
                                                                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{s}'");
                                                                        if (tmp == null) { key.ERRMsg = $"{key.ERRMsg}<br />查無 {s} 此共用生產站編號"; }
                                                                    }
                                                                    if (key.SStationNO_IndexSN_Merge.Contains(key.SApply_StationNO)) { key.ERRMsg = $"{key.ERRMsg}<br />主生產工站 與 共用生產站 不能有重複, 請取消共用站與主站同名的共用站."; }
                                                                    indexSN_Merge = $"'{key.SStationNO_IndexSN_Merge}'";
                                                                }
                                                                string outPackType = "0";
                                                                if (key.SApply_StationNO == _Fun.Config.OutPackStationName)
                                                                {
                                                                    outPackType = "1";
                                                                    if (key.SStationNO_IndexSN_Merge != null && key.SStationNO_IndexSN_Merge != "") { key.ERRMsg = $"{key.ERRMsg}<br />若為委外加工, 則不能有共用生產站"; }
                                                                }
                                                                if (key.SPP_Name == null || key.SPP_Name == "") { key.ERRMsg = $"{key.ERRMsg}<br />製程名稱不能為空白."; }
                                                                else if (key.SPartNO == null || key.SPartNO == "") { key.ERRMsg = $"{key.ERRMsg}<br />料件編號不能為空白."; }
                                                                else if (dr_Material["Class"].ToString() != "4" && dr_Material["Class"].ToString() != "5")
                                                                {
                                                                    key.ERRMsg = $"{key.ERRMsg}<br />設定不能為非生產件, 若此階層要改為非生產件, 請將此階刪除,從上階製程新增此非生產件, 或[修改]選項改為[新增]選項.";
                                                                }
                                                                #endregion

                                                                if (key.ERRMsg == "")
                                                                {
                                                                    string id = _Str.NewId('Z');
                                                                    string updateTimeKeyID = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                                                                    //###???若有分支,以下code會有問題
                                                                    #region 先異動所有階的 IndexSN, 與新增
                                                                    dt_tmp = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_PP_Name='{key.SPP_Name}' and Apply_PartNO='{key.SApply_PartNO}' order by IndexSN desc");
                                                                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                                                    {
                                                                        int contrlIndexNO = dt_tmp.Rows.Count + 1;
                                                                        for (int i = 0; i < dt_tmp.Rows.Count; i++)
                                                                        {
                                                                            DataRow dr = dt_tmp.Rows[i];
                                                                            if ((contrlIndexNO - indexNO) == 1)
                                                                            {
                                                                                db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IndexSN={contrlIndexNO.ToString()} where Id='{dr["Id"].ToString()}'"); contrlIndexNO -= 1;
                                                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[ServerId],[Apply_PartNO],PartNO,[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],UpdateTime,IsChackQTY,IsChackIsOK,OutPackType,Station_DIS_Remark,StationNO_Merge,MFNO,IsShare_PP_Name) VALUES
                                                                            ('{id}','{_Fun.Config.ServerId}','{key.SApply_PartNO}','{key.SPartNO}','0','{DateTime.Now.ToString("yyyy/MM/dd")}','{DateTime.Now.AddYears(10).ToString("yyyy/MM/dd")}','','{key.SPP_Name}','{key.SApply_StationNO}','0',{contrlIndexNO.ToString()},'','{key.SStationNO_Custom_DisplayName}','{updateTimeKeyID}','{key.SIsChackQTY.ToString()}','{key.SIsChackIsOK.ToString()}','{outPackType}','{key.SStation_DIS_Remark}',{indexSN_Merge},'{key.MFNO}','0')");
                                                                                contrlIndexNO -= 1;
                                                                            }
                                                                            else { db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IndexSN={contrlIndexNO.ToString()} where Id='{dr["Id"].ToString()}'"); contrlIndexNO -= 1; }
                                                                        }
                                                                        db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,[Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class],[AttritionRate]) VALUES
                                                                        ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{id}',1,'{key.SPartNO}',{key.SBOMQTY},'{dr_Material["Class"].ToString()}',0)");
                                                                    }
                                                                    #endregion

                                                                }
                                                            }
                                                            else { key.ERRMsg = $"系統程式異常, 請聯繫系統管理者."; }
                                                        }
                                                    }
                                                    else if (type1 == "3")
                                                    {
                                                        #region 從原物料新增 生產件
                                                        int sn = db.DB_GetQueryCount($"SELECT sn FROM SoftNetMainDB.[dbo].[BOMII] where BOMId='{key.BOMId}' order by sn desc");
                                                        db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,[Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class],[AttritionRate]) VALUES
                                                        ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{key.BOMId}',({sn + 1}),'{key.SPartNO}',{key.SBOMQTY},'{dr_Material["Class"].ToString()}',0)");

                                                        #region 刪除虛擬BOMII 的主BOM表料件 
                                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and BOMId='{key.BOMId}'");
                                                        if (tmp != null)
                                                        {
                                                            string befor_PartNO = tmp["PartNO"].ToString();
                                                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_PP_Name='{tmp["Apply_PP_Name"].ToString()}' and Apply_PartNO='{tmp["Apply_PartNO"].ToString()}' and IndexSN<{tmp["IndexSN"].ToString()} order by IndexSN desc");
                                                            if (tmp != null)
                                                            {
                                                                //tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and BOMId='{tmp["Id"].ToString()}' and PartNO!='{tmp["PartNO"].ToString()}' and (Class=4 or Class=5)");
                                                                //if (tmp != null)
                                                                //{ db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and BOMId='{key.BOMId}' and PartNO='{befor_PartNO}'"); }
                                                                db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and BOMId='{key.BOMId}' and PartNO='{befor_PartNO}'");
                                                            }
                                                        }
                                                        #endregion
                                                        #endregion
                                                    }
                                                    #endregion
                                                }
                                                else if (type2 == "D")
                                                {
                                                    #region 刪除
                                                    if (type1 == "2")
                                                    {
                                                        #region 記錄要刪BOM IndexNO所有Id
                                                        List<string> log_BOM_All_list = new List<string>();
                                                        log_BOM_All_list.Add(key.BOMId);
                                                        DataTable dt_log_BOM = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOMII] where BOMId='{key.BOMId}' order by sn");
                                                        if (dt_log_BOM != null && dt_log_BOM.Rows.Count > 0)
                                                        {
                                                            DEL_RecursiveBOM2(db, dt_log_BOM, indexNO, dr_selectBOM["Apply_PP_Name"].ToString(), dr_selectBOM["Apply_PartNO"].ToString(), ref log_BOM_All_list);
                                                        }
                                                        #endregion
                                                        if (log_BOM_All_list.Count > 0)
                                                        {
                                                            string id = "";
                                                            foreach (string s in log_BOM_All_list)
                                                            {
                                                                if (id == "") { id = $"'{s}'"; }
                                                                else { id = $"{id},'{s}'"; }
                                                            }
                                                            if (id != "")
                                                            {
                                                                db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BOM] where id in ({id})");
                                                                db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BOMII] where BOMId in ({id})");
                                                                db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IsConfirm='0' where ServerId='{_Fun.Config.ServerId}' and Id='{key.MBOMId}'");
                                                                #region 重新整理BOM檔 indexNO
                                                                int newIndexNO = int.Parse(tmp_MBOM["IndexSN"].ToString()) - indexNO;
                                                                if (newIndexNO > 0)
                                                                {
                                                                    db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IndexSN-={newIndexNO.ToString()} where Apply_PartNO='{dr_selectBOM["Apply_PartNO"].ToString()}' and Apply_PP_Name='{dr_selectBOM["Apply_PP_Name"].ToString()}'");
                                                                }
                                                                #endregion
                                                            }
                                                            if (key.BOMId == key.MBOMId)
                                                            {
                                                                key.COMType = "0";
                                                                key.SClass = "4";
                                                                key.MBOMId = "";
                                                                key.SApply_PartNO = "";
                                                                key.ERRMsg = "";
                                                                goto break_FUN;
                                                            }
                                                        }
                                                    }
                                                    else if (type1 == "3")
                                                    {
                                                        //刪子BOMII
                                                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BOMII] where Id='{key.BOMId}'");
                                                    }
                                                    #endregion
                                                }
                                                else if (type2 == "E")
                                                {
                                                    #region 修改
                                                    DataRow dr_tmp2 = null;
                                                    if (type1 == "3")
                                                    {
                                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[BOMII] set PartNO='{key.SPartNO}',Class='{dr_Material["Class"].ToString()}',BOMQTY={key.SBOMQTY} where Id='{key.BOMId}'");
                                                    }
                                                    else
                                                    {
                                                        #region 檢查
                                                        if (key.BOMId == key.MBOMId && dr_selectBOM["PartNO"].ToString() != key.SPartNO) { key.ERRMsg = "主製程的主料號, 不能異動, 若需修改, 建議將此製程刪除重建."; }
                                                        if (key.SApply_StationNO != "")
                                                        {
                                                            DataRow dr_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{key.SApply_StationNO}'");
                                                            if (dr_Station == null) { key.ERRMsg = $"{key.ERRMsg}<br />查無主生產工站編號, 請重新輸入."; }
                                                        }
                                                        else { key.ERRMsg = $"{key.ERRMsg}<br />主生產工站編號, 須指定一個廠內工作站 或 廠外委外作業站."; }
                                                        if (key.SApply_StationNO != _Fun.Config.OutPackStationName)
                                                        { key.MFNO = ""; }
                                                        if (key.MFNO != null && key.MFNO != "")
                                                        {
                                                            dr_tmp2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' and MFNO='{key.MFNO}'");
                                                            if (dr_tmp2 == null) { key.ERRMsg = $"{key.ERRMsg}<br />查無委外加工編號"; }
                                                        }
                                                        if (key.SApply_StationNO == _Fun.Config.OutPackStationName && (key.MFNO == null || key.MFNO == ""))
                                                        {
                                                            key.ERRMsg = $"{key.ERRMsg}<br />工作站為委外加工, 須給指定一個廠商編號.";
                                                        }
                                                        string indexSN_Merge = "NULL";
                                                        if (key.SStationNO_IndexSN_Merge != null && key.SStationNO_IndexSN_Merge != "")
                                                        {
                                                            foreach (string s in key.SStationNO_IndexSN_Merge.Split(','))
                                                            {
                                                                if (s == "") { continue; }
                                                                dr_tmp2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{s}'");
                                                                if (dr_tmp2 == null) { key.ERRMsg = $"{key.ERRMsg}<br />查無 {s} 此共用生產站編號"; }
                                                            }
                                                            if (key.SStationNO_IndexSN_Merge.Contains(key.SApply_StationNO)) { key.ERRMsg = $"{key.ERRMsg}<br />主生產工站 與 共用生產站 不能有重複, 請取消共用站與主站同名的共用站."; }
                                                            indexSN_Merge = $"'{key.SStationNO_IndexSN_Merge}'";
                                                        }
                                                        string outPackType = "0";
                                                        if (key.SApply_StationNO == _Fun.Config.OutPackStationName)
                                                        {
                                                            outPackType = "1";
                                                            if (key.SStationNO_IndexSN_Merge != null && key.SStationNO_IndexSN_Merge != "") { key.ERRMsg = $"{key.ERRMsg}<br />若為委外加工, 則不能有共用生產站"; }
                                                        }
                                                        if (key.SPP_Name == null || key.SPP_Name == "") { key.ERRMsg = $"{key.ERRMsg}<br />製程名稱不能為空白."; }
                                                        else if (key.SPartNO == null || key.SPartNO == "") { key.ERRMsg = $"{key.ERRMsg}<br />料件編號不能為空白."; }
                                                        else if (dr_Material["Class"].ToString() != "4" && dr_Material["Class"].ToString() != "5")
                                                        {
                                                            key.ERRMsg = $"{key.ERRMsg}<br />設定不能為非生產件, 若此階層要改為非生產件, 請將此階刪除,從上階製程新增此非生產件, 或[修改]選項改為[新增]選項.";
                                                        }
                                                        #endregion
                                                        if (key.ERRMsg == "")
                                                        {
                                                            if (db.DB_SetData($@"update SoftNetMainDB.[dbo].[BOM] set MFNO='{key.MFNO}',StationNO_Merge={indexSN_Merge},OutPackType='{outPackType}',PartNO='{key.SPartNO}',StationNO_Custom_DisplayName='{key.SStationNO_Custom_DisplayName}',
                                                        Station_DIS_Remark='{key.SStation_DIS_Remark}',Apply_StationNO='{key.SApply_StationNO}',IsChackQTY='{key.SIsChackQTY}',IsChackIsOK='{key.SIsChackIsOK}',UpdateTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}' where Id='{key.BOMId}'"))
                                                            {
                                                                if (key.SApply_StationNO != dr_selectBOM["Apply_StationNO"].ToString() || key.SPartNO != dr_selectBOM["PartNO"].ToString()) { db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IsConfirm='0' where ServerId='{_Fun.Config.ServerId}' and Id='{key.MBOMId}'"); }
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                            }
                                        }
                                    }
                                    else { key.ERRMsg = "沒選擇 刪除 或 新增 或 修改 的選項"; }
                                }
                                if (key.ERRMsg == "")
                                {
                                    key.COMType = "9";
                                    re = RecursiveBOM(db, key);
                                }
                                else
                                {
                                    if (type1 == "2") { key.COMType = "2"; }
                                    if (type1 == "3") { key.COMType = "3"; }
                                }
                            }
                            break;
                        case "7"://複製BOM
                            {
                                if (key.SPartNO == null || key.SPartNO == "" || key.SCOPY_PP_Name == null || key.SCOPY_PP_Name == "") { key.ERRMsg = $"{key.ERRMsg}<br />要複製的料件編號或製程名稱不能為空白."; }
                                else
                                {
                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_PP_Name='{key.SCOPY_PP_Name}' and Apply_PartNO='{key.SPartNO}'");
                                    if (tmp != null) { key.ERRMsg = $"{key.ERRMsg}<br />已有重複的製程名稱, 請另外重新命名."; }
                                    else
                                    {
                                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{key.MBOMId}' and Main_Item='1'");
                                        if (tmp == null) { key.ERRMsg = $"{key.ERRMsg}<br />被複製的料件編號, 找不到BOM資訊."; }
                                        {
                                            string newPartNO = key.SPartNO;
                                            string oldPartNO = tmp["PartNO"].ToString();
                                            string newBOMId = "";
                                            Copy_BOM(db, tmp, key.SCOPY_PP_Name, key.SPartNO, newPartNO, oldPartNO, ref newBOMId);

                                            key.MBOMId = newBOMId;
                                            key.SApply_PartNO = key.SPartNO;
                                            key.COMType = "9";
                                            re = RecursiveBOM(db, key);
                                        }
                                    }
                                }
                            }
                            break;
                        case "8"://發行
                            {
                                bool iserr = false;
                                string isEnd_ID = "";
                                List<string> log_BOM_All_id = new List<string>();
                                int indexSN = 0;
                                DataTable dt_BOM = null;
                                DataRow dr_BOM = null;
                                #region 檢查IndexSN
                                dt_BOM = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_PP_Name='{key.SPP_Name}' and Apply_PartNO='{key.SApply_PartNO}'");
                                if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                                {
                                    tmp = db.DB_GetFirstDataByDataRow($"select Id,IndexSN from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{key.MBOMId}' and Main_Item='1'");
                                    if (int.Parse(tmp["IndexSN"].ToString()) != dt_BOM.Rows.Count)
                                    { iserr = true; }
                                }
                                if (!iserr)
                                {
                                    dr_BOM = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{key.MBOMId}'");
                                    if (dr_BOM != null)
                                    {
                                        indexSN = int.Parse(dr_BOM["IndexSN"].ToString()) - 1;
                                        log_BOM_All_id.Add(key.MBOMId);
                                        for (int i = indexSN; i >= 1; i--)
                                        {
                                            if (db.DB_GetQueryCount($"select Id from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_PartNO='{dr_BOM["Apply_PartNO"].ToString()}' and Apply_PP_Name='{dr_BOM["Apply_PP_Name"].ToString()}' and IndexSN={i}") != 1)
                                            { iserr = true; break; }
                                            else
                                            {
                                                tmp = db.DB_GetFirstDataByDataRow($"select Id from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_PartNO='{dr_BOM["Apply_PartNO"].ToString()}' and Apply_PP_Name='{dr_BOM["Apply_PP_Name"].ToString()}' and IndexSN={i}");
                                                log_BOM_All_id.Add(tmp["Id"].ToString());
                                                if (i == 1)
                                                {
                                                    isEnd_ID = tmp["Id"].ToString();
                                                }
                                            }
                                        }
                                    }
                                    else { iserr = true; }
                                }
                                #endregion
                                if (iserr)
                                {
                                    //###???以下程式,若BOM有分支, 下列程式是錯的
                                    dt_BOM = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_PP_Name='{key.SPP_Name}' and Apply_PartNO='{key.SApply_PartNO}' order by Main_Item,IndexSN");
                                    if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                                    {
                                        indexSN = dt_BOM.Rows.Count;
                                        for (int i = 1; i <= indexSN; i++)
                                        {
                                            DataRow dr = dt_BOM.Rows[(i - 1)];
                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IndexSN={i.ToString()} where ServerId='{_Fun.Config.ServerId}' and Id='{dr["Id"].ToString()}'");
                                        }
                                        key.ERRMsg = "此BOM表資料有問題無法發行, 系統已嘗試自我修正, 請檢查修正後結構是否正確.";
                                    }
                                    else
                                    { key.ERRMsg = "此BOM表資料有問題無法發行, 請聯繫系統管理員."; }
                                }
                                else
                                {
                                    #region 檢查與修正 BOMII
                                    dt_BOM = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_PP_Name='{key.SPP_Name}' and Apply_PartNO='{key.SApply_PartNO}' order by IndexSN");
                                    if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                                    {
                                        indexSN = dt_BOM.Rows.Count;
                                        for (int i = 1; i <= indexSN; i++)
                                        {
                                            DataRow dr = dt_BOM.Rows[(i - 1)];
                                            DataTable dt_BOM2 = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and BOMId='{dr["Id"].ToString()}'");
                                            if (dt_BOM2 != null && dt_BOM2.Rows.Count > 0)
                                            {
                                                if (dt_BOM2.Rows.Count > 1)
                                                {
                                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and BOMId='{dr["Id"].ToString()}' and PartNO='{dr["PartNO"].ToString()}'");
                                                    if (tmp != null)
                                                    {
                                                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and BOMId='{dr["Id"].ToString()}' and PartNO='{dr["PartNO"].ToString()}'");
                                                        dt_BOM2 = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOMII] where ServerId='{_Fun.Config.ServerId}' and BOMId='{dr["Id"].ToString()}'");
                                                    }
                                                    if (dt_BOM2 != null && dt_BOM2.Rows.Count > 0)
                                                    {
                                                        for (int j = 1; j <= dt_BOM2.Rows.Count; j++)
                                                        {
                                                            DataRow dr2 = dt_BOM2.Rows[(j - 1)];
                                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[BOMII] set sn={j.ToString()} where ServerId='{_Fun.Config.ServerId}' and Id='{dr2["Id"].ToString()}'");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        dr_Material = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                                                        db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,[Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class],[AttritionRate]) VALUES
                                                        ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{dr["Id"].ToString()}',1,'{dr["PartNO"].ToString()}',1,'{dr_Material["Class"].ToString()}',0)");
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                dr_Material = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,[Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class],[AttritionRate]) VALUES
                                                        ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{dr["Id"].ToString()}',1,'{dr["PartNO"].ToString()}',1,'{dr_Material["Class"].ToString()}',0)");
                                            }

                                        }
                                    }
                                    #endregion

                                    db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IsConfirm='1' where ServerId='{_Fun.Config.ServerId}' and Id='{key.MBOMId}'");
                                    if (isEnd_ID != "") { db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IsEnd='1' where ServerId='{_Fun.Config.ServerId}' and Id='{isEnd_ID}'"); }

                                    #region 重新整理製程檔
                                    db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[PP_ProductProcess] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{dr_BOM["Apply_PP_Name"].ToString()}' and Apply_PartNO='{dr_BOM["Apply_PartNO"].ToString()}'");
                                    db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{dr_BOM["Apply_PP_Name"].ToString()}' and Apply_PartNO='{dr_BOM["Apply_PartNO"].ToString()}'");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess] (ServerId,PP_Name,Apply_PartNO,CalendarName,UpdateTime,FactoryName,LineName) VALUES
                                            ('{_Fun.Config.ServerId}','{dr_BOM["Apply_PP_Name"].ToString()}','{dr_BOM["Apply_PartNO"].ToString()}','{_Fun.Config.DefaultCalendarName}','{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}')");
                                    string indexSN_Merge = "0";
                                    string outPackType = "0";
                                    foreach (string id in log_BOM_All_id)
                                    {
                                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{id}'");
                                        if (tmp.IsNull("StationNO_Merge")) { indexSN_Merge = "0"; } else { indexSN_Merge = "1"; }
                                        if (tmp["Apply_StationNO"].ToString() != _Fun.Config.OutPackStationName) { outPackType = "0"; } else { outPackType = "1"; }
                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,Apply_PartNO,PP_Name,PartNO,StationNO,IndexSN,DisplaySN,Station_Custom_IndexSN,DisplayName,IndexSN_Merge,OutPackType,MFNO,BOMId) VALUES
                                            ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{tmp["Apply_PartNO"].ToString()}','{tmp["Apply_PP_Name"].ToString()}','{tmp["PartNO"].ToString()}','{tmp["Apply_StationNO"].ToString()}',{tmp["IndexSN"].ToString()},0,'{tmp["Station_Custom_IndexSN"].ToString()}','{tmp["StationNO_Custom_DisplayName"].ToString()}','{indexSN_Merge}','{outPackType}','{tmp["MFNO"].ToString()}','{tmp["Id"].ToString()}')");
                                        if (!tmp.IsNull("StationNO_Merge") && tmp["StationNO_Merge"].ToString() != "")
                                        {
                                            foreach (string s in tmp["StationNO_Merge"].ToString().Split(','))
                                            {
                                                if (s.Trim() == "") { continue; }
                                                db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,Apply_PartNO,PP_Name,PartNO,StationNO,IndexSN,DisplaySN,Station_Custom_IndexSN,DisplayName,IndexSN_Merge,OutPackType,MFNO,BOMId) VALUES
                                            ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{tmp["Apply_PartNO"].ToString()}','{tmp["Apply_PP_Name"].ToString()}','{tmp["PartNO"].ToString()}','{s}',{tmp["IndexSN"].ToString()},0,'{tmp["Station_Custom_IndexSN"].ToString()}','{tmp["StationNO_Custom_DisplayName"].ToString()}','{indexSN_Merge}','{outPackType}','{tmp["MFNO"].ToString()}','{tmp["Id"].ToString()}')");
                                            }
                                        }
                                    }
                                    dt_BOM = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{tmp["Apply_PP_Name"].ToString()}' and Apply_PartNO='{tmp["Apply_PartNO"].ToString()}' order by IndexSN");
                                    if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                                    {
                                        int displayNO = 0;
                                        foreach (DataRow row in dt_BOM.Rows)
                                        {
                                            displayNO += 1;
                                            db.DB_SetData($"update SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] set DisplaySN={displayNO.ToString()} where Id='{row["Id"].ToString()}'");
                                        }
                                    }
                                    #endregion
                                }
                                key.COMType = "9";
                                re = RecursiveBOM(db, key);
                            }
                            break;
                        case "A"://取消發行
                            {
                                DataRow tmp_MBOM = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{key.MBOMId}'");
                                if (tmp_MBOM == null)
                                { key.ERRMsg = $"查無系統BOM編號, 目前無法異動, 請重新查詢母件BOM表."; }
                                else
                                {
                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Apply_PP_Name='{tmp_MBOM["Apply_PP_Name"].ToString()}' and PartNO='{tmp_MBOM["PartNO"].ToString()}' and (State='1' or State='2' or State='6')");
                                    if (tmp != null)
                                    {
                                        key.ERRMsg = $"料號:{tmp_MBOM["PartNO"].ToString()} 製程:{tmp_MBOM["Apply_PP_Name"].ToString()} 已有在模擬中或生產中, 目前無法取消發行.";
                                    }
                                    else
                                    {
                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IsConfirm='0' where ServerId='{_Fun.Config.ServerId}' and Id='{key.MBOMId}'");
                                        db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[PP_ProductProcess] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{tmp_MBOM["Apply_PP_Name"].ToString()}' and Apply_PartNO='{tmp_MBOM["Apply_PartNO"].ToString()}'");
                                        db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{tmp_MBOM["Apply_PP_Name"].ToString()}' and Apply_PartNO='{tmp_MBOM["Apply_PartNO"].ToString()}'");
                                    }
                                }
                                key.COMType = "9";
                                re = RecursiveBOM(db, key);
                            }
                            break;
                        case "9"://顯示BOM Tree結構
                            re = RecursiveBOM(db, key);
                            break;
                    }
                }
            break_FUN:
                ViewBag.HtmlOutputBOM = re;
            }

            return View(key);
        }

        private string RecursiveBOM(DBADO db, CreateBOMOBJ key)
        {
            string re = "";
            #region 顯示BOM Tree結構
            if (key.MBOMId != null && key.MBOMId != "")
            {
                DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{key.MBOMId}' and Main_Item='1'");
                if (tmp == null) { return re; }
                key.SPP_Name = tmp["Apply_PP_Name"].ToString();
                key.SApply_PartNO = tmp["Apply_PartNO"].ToString();
                DataRow dr_Material = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{tmp["PartNO"].ToString()}'");
                key.SClass = dr_Material["Class"].ToString();
                string pp_Name_Display = $"[{tmp["StationNO_Custom_DisplayName"].ToString()}]";
                string isConfirm = "";
                if (!bool.Parse(tmp["IsConfirm"].ToString())) { isConfirm = $"<button type='submit' style='border: 2px blue none;background-color: #CDCD9A;' onclick='_me.isConfirmMBOMId(\"{key.MBOMId}\")'>正式發行BOM表</button>"; }
                else { isConfirm = $"<button type='submit' style='border: 2px blue none;background-color: #CDCD9A;' onclick='_me.isCancelMBOMId(\"{key.MBOMId}\")'>取消發行</button>"; }
                string indexNO = tmp["IndexSN"].ToString();
                if (indexNO == "0") { indexNO = "未發行"; }
                string copyBOM = "";
                DataTable dt_BOM = null;

                #region 產生BOM複製按鈕
                if (bool.Parse(tmp["IsConfirm"].ToString()) && key.HasPartNO_List != null && key.HasPartNO_List.Count > 0)
                {
                    copyBOM = "<input type='text' name='SPartNO' list='options7' value=''><datalist id='options7'>";
                    copyBOM = $"{copyBOM}<option value=''></option>";
                    var rudlist = key.HasPartNO_List.Where(x => (x.Split('\x01')[3] == key.SClass));
                    string[] item = null;
                    foreach (string s in rudlist)
                    {
                        item = s.Split('\x01');
                        var t01 = key.NeedIdList.Where(x => x.Split('\x01')[1] == item[0]).ToList();
                        if (t01!=null && t01.Count>0)
                        {
                            item[3] = "已有BOM";
                        } else 
                        { 
                            item[3] = "無";
                        }
                        copyBOM = $"{copyBOM}<option value='{item[0]}'>[{item[3]}]品名:{item[1]}&emsp;{item[2]} </option>";
                    }
                    copyBOM = $"&emsp;&emsp;<label>複製新BOM料號:</label>{copyBOM}</datalist><label>製程名稱:</label><input type='text' name='SCOPY_PP_Name' value=''><button type='submit' style='border: 2px blue none;background-color: #CDCD9A;' onclick='_me.isCopyBOM(\"{key.MBOMId}\")'>複製BOM表</button>";
                }
                #endregion

                re = $"<figure>   <figcaption>  母料號:{dr_Material["PartNO"].ToString()}  品名:{dr_Material["PartName"].ToString()}  規格:{dr_Material["Specification"].ToString()}  {isConfirm}  {copyBOM}</figcaption><hr><figcaption>用料結構 與 生產流程表</figcaption><ul class='tree'>   <li><button type='submit' style='border: 2px blue none;' onclick='_me.onCreateBOM(\"2,{key.MBOMId},{tmp["Apply_PP_Name"].ToString()},{key.MBOMId}\")'>{tmp["PartNO"].ToString()}[{dr_Material["PartName"].ToString()}][{dr_Material["Specification"].ToString()}][{indexNO}][{tmp["Apply_StationNO"].ToString()}]{pp_Name_Display}</button>";
                dt_BOM = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOMII] where BOMId='{tmp["Id"].ToString()}' order by sn");
                if (dt_BOM != null && dt_BOM.Rows.Count > 0)
                {
                    RecursiveBOM2(db, key.MBOMId, 0, 1, dt_BOM, int.Parse(tmp["IndexSN"].ToString()), tmp["Apply_PP_Name"].ToString(), tmp["Apply_PartNO"].ToString(), key.MBOMId, ref re);
                }
                else
                { re = $"{re}</li>"; }
                re = $"{re}</ul></figure>";
            }
            else { re = "無母件料號, BOM表編號."; }
            #endregion
            return re;
        }
        private void RecursiveBOM2(DBADO db, string needId, int sn, int sn2, DataTable dr_M, int indexSN, string pp_Name, string apply_PartNO, string mBOMId, ref string re)
        {
            if (dr_M == null || dr_M.Rows.Count <= 0) { return; }
            DataRow dr_tmp = null;
            DataTable dr_MII = null;
            DataRow dr_Material = null;
            if (sn != sn2) { re = $"{re}<ul>"; sn = sn2; }
            foreach (DataRow dr1 in dr_M.Rows)
            {
                dr_Material = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr1["PartNO"].ToString()}'");
                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where Apply_PartNO='{apply_PartNO}' and Apply_PP_Name='{pp_Name}' and IndexSN={(indexSN - 1)}");
                if (dr_tmp != null)
                {
                    string pp_Name_Display = $"[{dr_tmp["StationNO_Custom_DisplayName"].ToString()}]";
                    string indexNO = dr_tmp["IndexSN"].ToString();
                    if (indexNO == "0") { indexNO = "未發行"; }
                    re = $"{re}   <li><button type='submit' style='border: 2px blue none;' onclick='_me.onCreateBOM(\"2,{mBOMId},{pp_Name},{dr_tmp["Id"].ToString()}\")'>{dr1["PartNO"].ToString()}[{dr_Material["PartName"].ToString()}][{dr_Material["Specification"].ToString()}][{indexNO}][{dr_tmp["Apply_StationNO"].ToString()}]{pp_Name_Display}</button>";
                    dr_MII = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOMII] where BOMId='{dr_tmp["Id"].ToString()}' order by sn");
                    if (dr_MII != null && dr_MII.Rows.Count > 0)
                    {
                        RecursiveBOM2(db, needId, sn2, (sn2 + 1), dr_MII, (indexSN - 1), pp_Name, apply_PartNO, mBOMId, ref re);
                    }
                }
                else
                {
                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where Id='{mBOMId}'");
                    if (dr_tmp != null)
                    {
                        if (dr1["PartNO"].ToString() == dr_tmp["PartNO"].ToString()) { continue; }
                    }
                    re = $"{re}   <li><button type='submit' onclick='_me.onCreateBOM(\"3,{mBOMId},{pp_Name},{dr1["Id"].ToString()}\")'>用料:{dr1["PartNO"].ToString()}[{dr_Material["PartName"].ToString()}]</button>";
                }
                re = $"{re}</li>";
            }
            re = $"{re}</ul>";
        }


        /*
        private bool getBOMII_8(DBADO db, int indexSN, string BOMId, string pp_Name, ref string isEnd_ID)
        {
            DataTable dt_BOMII = db.DB_GetData($@"select a.*,b.PartName,b.Specification,b.Class FROM SoftNetMainDB.[dbo].[BOMII] as a
                                                join SoftNetMainDB.[dbo].[Material] as b on b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO 
                                                where BOMId='{BOMId}' order by b.Class,sn");
            if (dt_BOMII != null && dt_BOMII.Rows.Count > 0)
            {
                indexSN -= 1;
                if (indexSN >= 1)
                {
                    DataRow tmp = null;
                    bool iserr = false;
                    foreach (DataRow drII in dt_BOMII.Rows)
                    {
                        if (drII["Class"].ToString() == "4" || drII["Class"].ToString() == "5")
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"select a.Id,a.PartNO,a.Apply_PP_Name,a.Apply_StationNO,a.IsEnd,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName,a.OutPackType,b.Class,IsChackQTY,IsChackIsOK from SoftNetMainDB.[dbo].[BOM] as a,dbo.[Material] as b where a.Main_Item='0' and a.IndexSN={indexSN} and b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{drII["PartNO"].ToString()}' and a.Apply_PP_Name='{pp_Name}' and a.PartNO=b.PartNO order by IndexSN desc");
                            if (tmp == null) { return true; }
                            else
                            {
                                if (indexSN > 1)
                                {
                                    iserr = getBOMII_8(db, indexSN, tmp["Id"].ToString(), pp_Name, ref isEnd_ID);
                                    if (iserr) { return true; }
                                }
                                else
                                {
                                    if (indexSN == 1) { isEnd_ID = tmp["Id"].ToString(); }
                                }
                            }
                        }
                        if (iserr) { return true; }
                    }
                }
            }
            return false;
        }
        */
        private void DEL_RecursiveBOM2(DBADO db, DataTable dr_M, int indexSN, string pp_Name, string apply_PartNO, ref List<string> del_list_M)
        {
            if (dr_M == null || dr_M.Rows.Count <= 0) { return; }
            DataRow dr_tmp = null;
            DataTable dr_MII = null;
            foreach (DataRow dr1 in dr_M.Rows)
            {
                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where Apply_PartNO='{apply_PartNO}' and Apply_PP_Name='{pp_Name}' and IndexSN={(indexSN - 1)}");
                if (dr_tmp != null)
                {
                    del_list_M.Add(dr_tmp["Id"].ToString());
                    dr_MII = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOMII] where BOMId='{dr_tmp["Id"].ToString()}' order by sn");
                    if (dr_MII != null && dr_MII.Rows.Count > 0) { DEL_RecursiveBOM2(db, dr_MII, (indexSN - 1), dr_tmp["Apply_PP_Name"].ToString(), apply_PartNO, ref del_list_M); }
                }
            }
        }
        private void Copy_BOM(DBADO db, DataRow bom, string new_PP_Name, string new_apply_PartNO, string new_partNO, string old_partNO, ref string newBOMId)
        {
            string sql = "";
            string newId = "";
            string partNNO = "";
            string old_PP_Name = bom["Apply_PP_Name"].ToString();
            string old_Apply_PartNO = bom["Apply_PartNO"].ToString();
            int old_IndexSN = int.Parse(bom["IndexSN"].ToString());
            if (newBOMId == "")
            {
                newBOMId = _Str.NewId('Z');
                newId = newBOMId;
                string stationNO_Merge = bom.IsNull("StationNO_Merge") ? "NULL" : $"'{bom["StationNO_Merge"].ToString()}'";
                sql = $@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[ServerId],[Apply_PartNO],PartNO,[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],UpdateTime,IsChackQTY,IsChackIsOK,OutPackType,Station_DIS_Remark,StationNO_Merge,MFNO,IsShare_PP_Name) VALUES
                                                    ('{newBOMId}','{_Fun.Config.ServerId}','{new_apply_PartNO}','{new_apply_PartNO}','1','{DateTime.Now.ToString("yyyy/MM/dd")}','{DateTime.Now.AddYears(10).ToString("yyyy/MM/dd")}','1.0000','{new_PP_Name}','{bom["Apply_StationNO"].ToString()}','0',
                                                    {bom["IndexSN"].ToString()},'{bom["Station_Custom_IndexSN"].ToString()}','{bom["StationNO_Custom_DisplayName"].ToString()}','{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','{bom["IsChackQTY"].ToString()}','{bom["IsChackIsOK"].ToString()}','{bom["OutPackType"].ToString()}','{bom["Station_DIS_Remark"].ToString()}',{stationNO_Merge},'{bom["MFNO"].ToString()}','{bom["IsShare_PP_Name"].ToString()}')";
            }
            else
            {
                newId = _Str.NewId('Z');
                if (old_partNO == bom["PartNO"].ToString()) { partNNO = new_partNO; } else { partNNO = bom["PartNO"].ToString(); }
                string stationNO_Merge = bom.IsNull("StationNO_Merge") ? "NULL" : $"'{bom["StationNO_Merge"].ToString()}'";
                sql = $@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[ServerId],[Apply_PartNO],PartNO,[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],UpdateTime,IsChackQTY,IsChackIsOK,OutPackType,Station_DIS_Remark,StationNO_Merge,MFNO,IsShare_PP_Name) VALUES
                                                    ('{newId}','{_Fun.Config.ServerId}','{new_apply_PartNO}','{partNNO}','0','{DateTime.Now.ToString("yyyy/MM/dd")}','{DateTime.Now.AddYears(10).ToString("yyyy/MM/dd")}','','{new_PP_Name}','{bom["Apply_StationNO"].ToString()}','0',
                                                    {bom["IndexSN"].ToString()},'{bom["Station_Custom_IndexSN"].ToString()}','{bom["StationNO_Custom_DisplayName"].ToString()}','{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','{bom["IsChackQTY"].ToString()}','{bom["IsChackIsOK"].ToString()}','{bom["OutPackType"].ToString()}','{bom["Station_DIS_Remark"].ToString()}',{stationNO_Merge},'{bom["MFNO"].ToString()}','{bom["IsShare_PP_Name"].ToString()}')";
            }
            db.DB_SetData(sql);

            DataTable bomII = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOMII] where BOMId='{bom["Id"].ToString()}' order by sn");
            foreach (DataRow dr1 in bomII.Rows)
            {
                if (old_partNO == dr1["PartNO"].ToString()) { partNNO = new_partNO; } else { partNNO = dr1["PartNO"].ToString(); }
                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,[Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class],[AttritionRate]) VALUES
                                            ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{newId}',{dr1["sn"].ToString()},'{partNNO}',{dr1["BOMQTY"].ToString()},'{dr1["Class"].ToString()}',{dr1["AttritionRate"].ToString()})");
            }
            if (old_IndexSN <= 1) { return; }
            else
            {
                bom = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where  Apply_PP_Name='{old_PP_Name}' and Apply_PartNO='{old_Apply_PartNO}' and IndexSN={(old_IndexSN - 1).ToString()}");
                if (bom != null)
                { Copy_BOM(db, bom, new_PP_Name, new_apply_PartNO, new_partNO, old_partNO,ref newBOMId); }
            }
        }



    }
}
