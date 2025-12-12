using Base;
using Base.Services;
using Microsoft.AspNetCore.Mvc;

using System;
using System.Data;

namespace SoftNetWebII.Controllers
{
    public class ZTestController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Test02()
        {
            string re = "NG";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                DataTable dt = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[BOM] where ServerId='02' order by PartNO,IndexSN");
                if (dt != null)
                {
                    string id = "";
                    /*
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (dr["IndexSN"].ToString() == "1")
                        {
                            db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IsEnd='1' where ServerId='02' and Id='{dr["Id"].ToString()}'");
                        }
                        if (dr["Version"].ToString() == "1.000")
                        {
                            db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set IsConfirm='1' where ServerId='02' and Id='{dr["Id"].ToString()}'");
                        }
                        if (dr["Apply_StationNO"].ToString() == "複合車床")
                        {
                            db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set Apply_StationNO='LMC_2' where ServerId='02' and Id='{dr["Id"].ToString()}'");
                        }
                        if (dr["Apply_StationNO"].ToString() == "車床")
                        {
                            db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set Apply_StationNO='LN_1' where ServerId='02' and Id='{dr["Id"].ToString()}'");
                        }
                        if (dr["Apply_StationNO"].ToString() == "銑床")
                        {
                            db.DB_SetData($"update SoftNetMainDB.[dbo].[BOM] set Apply_StationNO='MC_1' where ServerId='02' and Id='{dr["Id"].ToString()}'");
                        }
                    }
                    */
                    if (db.DB_GetQueryCount($"SELECT *  FROM SoftNetSYSDB.[dbo].[PP_ProductProcess] where ServerId='02'") <= 0)
                    {
                        int j = 0;
                        dt = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[BOM] where ServerId='02' order by PartNO,IndexSN");
                        string logname = "";
                        try
                        {
                            foreach (DataRow dr in dt.Rows)
                            {
                                if (db.DB_GetQueryCount($"SELECT *  FROM SoftNetSYSDB.[dbo].[PP_ProductProcess] where PP_Name='{dr["PartNO"].ToString()}_加工製程'") <= 0)
                                {
                                    logname = dr["PartNO"].ToString();
                                    if (id != dr["PartNO"].ToString())
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess] (ServerId,PP_Name,UpdateTime,CalendarName,FactoryName,LineName) VALUES
                                                                ('02','{dr["PartNO"].ToString()}_加工製程','{logDate}','{_Fun.Config.DefaultCalendarName}','萬銓','G008')");
                                        id = dr["PartNO"].ToString();
                                        j = 0;
                                    }
                                }
                                db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','{dr["Apply_StationNO"].ToString()}','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                if (dr["Apply_StationNO"].ToString() == "MC_1")
                                {
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','MC_2','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','MC_3','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','MC_4','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','MC_5','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','MC_6','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','MC_7','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','MC_8','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");

                                }
                                else if (dr["Apply_StationNO"].ToString() == "LN_1")
                                {
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LN_2','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LN_3','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LN_4','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LN_5','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LN_6','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LN_7','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LN_8','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                }
                                else if (dr["Apply_StationNO"].ToString() == "LMC_2")
                                {
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LMC_3','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LMC_4','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LMC_5','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LMC_6','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LMC_7','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LMC_8','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LMC_9','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");
                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey,BOMId) VALUES
                                                                ('{_Str.NewId('Y')}','02','萬銓','G008','{dr["PartNO"].ToString()}_加工製程',{(++j).ToString()},{dr["IndexSN"].ToString()},'1','LMC_10','{dr["StationNO_Custom_DisplayName"].ToString()}','','0','0','{dr["Id"].ToString()}')");

                                }
                            }
                        }
                        catch
                        {
                            string _s = logname;
                            string S = "";
                        }
                    }


                    re = "OK";
                }
            }
            ViewBag.Class = re;
            return View();
        }

        public IActionResult Test01()
        {
            string re = "NG";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                if (db.DB_GetQueryCount("SELECT * from SoftNetMainDB.[dbo].[BOM] where PartNO='YLVPIZ012'") <= 0)
                {
                    DataTable dt = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='01' and Class='5'");
                    if (dt != null)
                    {
                        string id = "";
                        foreach (DataRow dr in dt.Rows)
                        {
                            //組裝製程
                            id = _Str.NewId('Z');
                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[PartNO],[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN]) VALUES
                                        ('{id}','{dr["PartNO"].ToString()}','1','2024-01-01','2034-12-31','1.0000','組裝製程','H01','0',3,'')");
                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,[Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class],[AttritionRate],[PP_Name]) VALUES
                                        ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{id}',1,'{dr["PartNO"].ToString()}',1,'{dr["Class"].ToString()}',0,'')");

                            id = _Str.NewId('Z');
                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[PartNO],[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN]) VALUES
                                        ('{id}','{dr["PartNO"].ToString()}','0','2024-01-01','2034-12-31','','組裝製程','G01','0',2,'')");
                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] ([Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class],[AttritionRate],[PP_Name]) VALUES
                                        ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{id}',1,'{dr["PartNO"].ToString()}',1,'{dr["Class"].ToString()}',0,'')");

                            id = _Str.NewId('Z');
                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[PartNO],[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN]) VALUES
                                        ('{id}','{dr["PartNO"].ToString()}','0','2024-01-01','2034-12-31','','組裝製程','F01','1',1,'')");
                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,[Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class],[AttritionRate],[PP_Name]) VALUES
                                        ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{id}',1,'{dr["PartNO"].ToString()}',1,'{dr["Class"].ToString()}',0,'')");
                        }
                        re = "OK";
                    }
                }
            }
            ViewBag.Class = re;
            return View();
        }

        public IActionResult Test03()
        {
            ViewBag.DisplayHTML = "";
            ViewBag.StationNO = "";
            ViewBag.StationName = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' order by StationNO");
                if (dt != null && dt.Rows.Count>0)
                {
                    int i = 1;
                    string re = "";
                    string stNO = "";
                    string stName = "";
                    float value = 0;
                    DataRow dr_tmp = null;
                    int ct = 0;
                    int ect = 0;
                    int lct = 0;
                    int uct = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        value = 0.01f;
                        #region 計算值
                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 7 AVG(CycleTime+WaitTime) as CT,AVG(ECT) as ECT,AVG(LowerCT) as LCT,AVG(UpperCT) as UCT FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CycleTime!=0");
                        if (dr_tmp != null && !dr_tmp.IsNull("CT") && dr_tmp["CT"].ToString() != "" && int.Parse(dr_tmp["CT"].ToString()) > 0 && int.Parse(dr_tmp["ECT"].ToString()) > 0)
                        {
                            ct = int.Parse(dr_tmp["CT"].ToString());
                            if (dr_tmp["LCT"].ToString() != "" && int.Parse(dr_tmp["LCT"].ToString()) >= ct)
                            {
                                value = 100;
                            }
                            else if (dr_tmp["UCT"].ToString() != "" && ct >= int.Parse(dr_tmp["UCT"].ToString()))
                            {
                                value = 0.01f;
                            }
                            else if (dr_tmp["ECT"].ToString() != "" && dr_tmp["LCT"].ToString() != "" && int.Parse(dr_tmp["ECT"].ToString()) >= ct && int.Parse(dr_tmp["LCT"].ToString()) >= ct)
                            {
                                uct = int.Parse(dr_tmp["UCT"].ToString());
                                value = ((uct - ct) / uct) * 100;
                                value = (100 - value) * (60 / 100) + 40;
                            }
                            else if (dr_tmp["ECT"].ToString() != "" && int.Parse(dr_tmp["ECT"].ToString()) >= ct)
                            {
                                ect = int.Parse(dr_tmp["ECT"].ToString());
                                value = ((ect - ct) / ect) * 100;
                                value = (100 - value) * (40 / 100);
                            }
                        }
                        if (value <= 0) { value = 0.01f; }
                        if (value >100) { value = 100; }
                        #endregion

                        if (stNO == "") { stNO = $"{dr["StationNO"].ToString()},{value.ToString("0.00")}"; }
                        else { stNO = $"{stNO};{dr["StationNO"].ToString()},{value.ToString("0.00")}"; }
                        if (stName == "") { stName = $"{dr["StationName"].ToString()}"; }
                        else { stName = $"{stName};{dr["StationName"].ToString()}"; }

                        #region display StationNO
                        if (i == 1) { re = $"{re}<div class='row'>"; }
                        re = $"{re}<div id='{dr["StationNO"].ToString()}' class='gaugeSVGVeiw'></div>";
                        if (i >= 4) { re = $"{re}</div>"; i = 1; } else { ++i; }
                        #endregion
                    }
                    ViewBag.StationNO = stNO;
                    ViewBag.DisplayHTML = re;
                    ViewBag.StationName = stName;

                }
            }
            return View();
        }
    }
}
