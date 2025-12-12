using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
//using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Net.WebSockets;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    
    public class DOCconfirmController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public DOCconfirmController(SNWebSocketService websocket, SFC_Common sfc_Common)
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
        public ActionResult Index()
        {
            List<IdStrDto> station_Type = new List<IdStrDto>();
            station_Type.Add(new IdStrDto("1", "作業站"));
            station_Type.Add(new IdStrDto("2", "維修站"));
            station_Type.Add(new IdStrDto("3", "控制站"));
            station_Type.Add(new IdStrDto("7", "虛擬站"));
            station_Type.Add(new IdStrDto("8", "多工站"));
            ViewBag.Station_Type = station_Type;

            List<IdStrDto> stationUI_type = new List<IdStrDto>();
            stationUI_type.Add(new IdStrDto("1", "追朔"));
            stationUI_type.Add(new IdStrDto("2", "不追朔"));
            ViewBag.StationUI_type = stationUI_type;

            List<IdStrDto> factoryName = new List<IdStrDto>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData("SELECT FactoryName FROM [dbo].[Factory]");
                if (dt != null && dt.Rows.Count > 0)
                {
                    int i = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        factoryName.Add(new IdStrDto(dr["FactoryName"].ToString(), dr["FactoryName"].ToString()));
                    }
                }
            }

            ViewBag.FactoryName = factoryName;

            return View();
        }


        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(dt, Ctrl));
        }

        private DOCconfirmService EditService()
        {
            return new DOCconfirmService(Ctrl);
        }

        [HttpPost]
        public string ConfirmDOC1Buy(string keys) //採購確認  0=ipport,1=單號子項的Id,2=單號子項的Id,,,,,
        {
            string meg = "";
            string[] data = keys.Split(',');
            string ipport = data[0];
            List<string> lists = new List<string>();
            for (int i = 1; i < data.Length; i++)
            {
                lists.Add(data[i]);
            }
            if (lists.Count > 0)
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    string sql = "";
                    DataRow tmp_dr = null;
                    string storeNO = "";
                    string storeSpacesNO = "";
                    Dictionary<string, List<string[]>> _poolTEST_NO = new Dictionary<string, List<string[]>>();//暫存檢驗單
                    Dictionary<string, List<string[]>> _poolStore_NO = new Dictionary<string, List<string[]>>();//暫存入庫單

                    foreach (string s in lists)
                    {
                        sql = $"SELECT a.*,b.DOCType  FROM SoftNetMainDB.[dbo].[DOC1BuyII] as a,SoftNetMainDB.[dbo].[DOCRole] as b where a.Id='{s}' and SUBSTRING(a.DOCNumberNO,1,4)=b.DOCNO and b.ServerId='{_Fun.Config.ServerId}'";//將來where條件要加DOCType=?
                        DataRow dr_DOC1Buy = db.DB_GetFirstDataByDataRow(sql);
                        if (dr_DOC1Buy != null)
                        {
                            switch (dr_DOC1Buy["DOCType"].ToString())
                            {
                                case "3"://領類
                                case "4"://入類
                                case "5"://調撥
                                case "1"://採購類
                                case "2"://銷貨類
                                case "6"://生產類
                                case "7"://委外類
                                case "8"://檢驗類
                                         //case "9"://財務類
                                         //###???未完成
                                    storeNO = dr_DOC1Buy["StoreNO"].ToString();
                                    storeSpacesNO = dr_DOC1Buy["StoreSpacesNO"].ToString();
                                    break;
                            }
                            int ct = 0;
                            if (!dr_DOC1Buy.IsNull("StartTime")) { ct = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC1Buy["StartTime"]), DateTime.Now); }
                            if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC1BuyII] set IsOK='1',CT={ct},EndTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where Id='{s}' and DOCNumberNO='{dr_DOC1Buy["DOCNumberNO"].ToString()}' and IsOK='0'"))
                            {
                                #region 計算單據PP_EfficientDetail
                                string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
                                string partNO = dr_DOC1Buy["PartNO"].ToString();
                                string pp_Name = "";
                                string E_stationNO = "";
                                if (dr_DOC1Buy["SimulationId"].ToString() != "")
                                {
                                    DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC1Buy["SimulationId"].ToString()}'");
                                    pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                                    if (!dr_tmp.IsNull("Source_StationNO") && dr_tmp["Source_StationNO"].ToString() != "")
                                    { E_stationNO = dr_tmp["Source_StationNO"].ToString(); }
                                    else { E_stationNO = dr_tmp["Apply_StationNO"].ToString(); }
                                }
                                DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} CT from SoftNetMainDB.[dbo].[DOC1BuyII] where SUBSTRING(Id,2,2)='{_Fun.Config.ServerId}' and SUBSTRING(DOCNumberNO,1,4)='{dr_DOC1Buy["DOCNumberNO"].ToString().Substring(0, 4)}' and PartNO='{dr_DOC1Buy["PartNO"].ToString()}' and CT>0");
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
                                    DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC1Buy] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO='{dr_DOC1Buy["DOCNumberNO"].ToString()}'");
                                    _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, pp_Name, pp_Name,"0", partNO, partNO, dr_DOC1Buy["DOCNumberNO"].ToString().Substring(0, 4), "", dr_tmp["MFNO"].ToString());
                                }
                                #endregion


                                #region 檢驗,入庫單
                                tmp_dr = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC1Buy["PartNO"].ToString()}'");
                                if (tmp_dr != null)
                                {
                                    if (bool.Parse(tmp_dr["IS_Store_Test"].ToString()))
                                    {
                                        //檢驗單
                                        if (_poolTEST_NO.ContainsKey(dr_DOC1Buy["DOCNumberNO"].ToString()))
                                        { _poolTEST_NO[dr_DOC1Buy["DOCNumberNO"].ToString()].Add(new string[] { dr_DOC1Buy["PartNO"].ToString(), Convert.ToDateTime(dr_DOC1Buy["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff"), dr_DOC1Buy["QTY"].ToString(), dr_DOC1Buy["Unit"].ToString(), dr_DOC1Buy["SimulationId"].ToString(), dr_DOC1Buy["Id"].ToString() }); }
                                        else
                                        {
                                            List<string[]> tmp = new List<string[]>();
                                            tmp.Add(new string[] { dr_DOC1Buy["PartNO"].ToString(), Convert.ToDateTime(dr_DOC1Buy["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff"), dr_DOC1Buy["qty"].ToString(), dr_DOC1Buy["Unit"].ToString(), dr_DOC1Buy["SimulationId"].ToString(), dr_DOC1Buy["Id"].ToString() });
                                            _poolTEST_NO.Add(dr_DOC1Buy["DOCNumberNO"].ToString(), tmp);
                                        }
                                    }
                                    else
                                    {
                                        //入庫單
                                        if (_poolStore_NO.ContainsKey(dr_DOC1Buy["DOCNumberNO"].ToString()))
                                        { _poolStore_NO[dr_DOC1Buy["DOCNumberNO"].ToString()].Add(new string[] { dr_DOC1Buy["PartNO"].ToString(), Convert.ToDateTime(dr_DOC1Buy["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff"), dr_DOC1Buy["QTY"].ToString(), dr_DOC1Buy["Unit"].ToString(), dr_DOC1Buy["SimulationId"].ToString(), dr_DOC1Buy["Id"].ToString() }); }
                                        else
                                        {
                                            List<string[]> tmp = new List<string[]>();
                                            tmp.Add(new string[] { dr_DOC1Buy["PartNO"].ToString(), Convert.ToDateTime(dr_DOC1Buy["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff"), dr_DOC1Buy["QTY"].ToString(), dr_DOC1Buy["Unit"].ToString(), dr_DOC1Buy["SimulationId"].ToString(), dr_DOC1Buy["Id"].ToString() });
                                            _poolStore_NO.Add(dr_DOC1Buy["DOCNumberNO"].ToString(), tmp);
                                        }
                                    }
                                }
                                #endregion

                            }


                        }
                    }
                    int no = 0;
                    string buyNO = "";
                    string docNO = "";
                    string date = DateTime.Now.ToString("yyyyMMdd");

                    if (_poolTEST_NO.Count > 0)
                    {
                        docNO = "PA01";//###???寫死
                        #region 處理檢驗單號
                        no = 0;
                        tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT DOCNumberNO FROM SoftNetMainDB.[dbo].[DOC3stock] where ServerId='{_Fun.Config.ServerId}' and DOCNO='{docNO}' order by DOCNumberNO desc");
                        if (tmp_dr != null)
                        {
                            no = int.Parse(tmp_dr["DOCNumberNO"].ToString().Trim().Substring(12));
                        }
                        foreach (KeyValuePair<string, List<string[]>> r in _poolTEST_NO)
                        {
                            buyNO = $"PA01{date}{(++no).ToString().PadLeft(4, '0')}";
                            sql = $"INSERT INTO SoftNetMainDB.[dbo].[DOC3stock] (ServerId,DOCNO,DOCNumberNO,DOCDate,UserId,DOCType,SourceNO,TotalMoney,TaxMoney,FlowLevel,FlowStatus) VALUES ('{_Fun.Config.ServerId}','{docNO}','{buyNO}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','Alex','8','{r.Key}',100,5,0,'Y')";
                            if (db.DB_SetData(sql))
                            {
                                //s 0="PartNO 1=CalendarDate 2=qty 3=Unit 4=SimulationId 5=上筆關聯Id
                                foreach (string[] s in r.Value)
                                {
                                    db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[DOC3stockII] (ServerId,Id,DOCNumberNO,PartNO,Price,Unit,QTY,SimulationId,ArrivalDate,IN_StoreNO,IN_StoreSpacesNO) VALUES ('{_Fun.Config.ServerId}','{s[5]}','{buyNO}','{s[0]}',0,'{s[3]}','{s[2]}','{s[4]}','{s[1]}','a1','')");
                                }
                            }
                        }
                        #endregion

                    }
                    if (_poolStore_NO.Count > 0)
                    {
                        docNO = "AA02";//###???寫死

                        #region 處理入庫單號
                        no = 0;
                        tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT DOCNumberNO FROM SoftNetMainDB.[dbo].[DOC3stock] where ServerId='{_Fun.Config.ServerId}' and DOCNO='{docNO}' order by DOCNumberNO desc");
                        if (tmp_dr != null)
                        {
                            no = int.Parse(tmp_dr["DOCNumberNO"].ToString().Trim().Substring(12));
                        }
                        foreach (KeyValuePair<string, List<string[]>> r in _poolStore_NO)
                        {
                            buyNO = $"{docNO}{date}{(++no).ToString().PadLeft(4, '0')}";
                            sql = $"INSERT INTO SoftNetMainDB.[dbo].[DOC3stock] (ServerId,DOCNO,DOCNumberNO,DOCDate,UserId,DOCType,SourceNO,TotalMoney,TaxMoney,FlowLevel,FlowStatus) VALUES ('{_Fun.Config.ServerId}','{docNO}','{buyNO}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','Alex','4','{r.Key}',100,5,0,'Y')";
                            if (db.DB_SetData(sql))
                            {
                                foreach (string[] s in r.Value)
                                {
                                    db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[DOC3stockII] (ServerId,Id,DOCNumberNO,PartNO,Price,Unit,QTY,SimulationId,ArrivalDate,IN_StoreNO,IN_StoreSpacesNO) VALUES ('{_Fun.Config.ServerId}','{s[5]}','{buyNO}','{s[0]}',0,'{s[3]}','{s[2]}','{s[4]}','{s[1]}','a1','')");
                                }
                            }
                        }
                        #endregion
                    }
                }
            }



     
            /*
            List<string> Ids = new List<string>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                //nos = "'000000000001','Z0117W3PH045','Z0117W3PH06F','Z0117W3PH1R7','Z0117W3PH1R8'";
                //string ReadSql = $@"select b.SourceNO,a.* from SoftNetMainDB.[dbo].[DOC3stockII] as a 
                //                        join SoftNetMainDB.[dbo].[DOC3stock] as b on a.DOCNumberNO=b.DOCNumberNO
                //                        where a.IsOK='0' and a.Id in ({nos}) order by b.SourceNO,a.Id";

                string ReadSql = $@"select b.SourceNO,a.* from SoftNetMainDB.[dbo].[DOC3stockII] as a 
                                        join SoftNetMainDB.[dbo].[DOC3stock] as b on a.DOCNumberNO=b.DOCNumberNO
                                        where a.IsOK='0'  order by b.SourceNO,a.Id";


                DataTable tmp_dt = db.DB_GetData(ReadSql);
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    foreach (DataRow d in tmp_dt.Rows)
                    {
                        if (d["SourceNO"].ToString().Trim() == "")
                        {
                            Ids.Add($"'{d["Id"].ToString()}'"); continue;
                        }
                        string _s = $"select * from SoftNetMainDB.[dbo].[DOC1BuyII] where IsOK='1' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["SourceNO"].ToString().Trim()}'";
                        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[DOC1BuyII] where IsOK='1' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["SourceNO"].ToString().Trim()}'") > 0)
                        {
                            Ids.Add($"'{d["Id"].ToString()}'");
                        }
                        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[DOC2SalesII] where IsOK='1' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["SourceNO"].ToString().Trim()}'") > 0)
                        {
                            Ids.Add($"'{d["Id"].ToString()}'");
                        }
                        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='1' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["SourceNO"].ToString().Trim()}'") > 0)
                        {
                            Ids.Add($"'{d["Id"].ToString()}'");
                        }
                        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[DOC4ProductionII] where IsOK='1' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["SourceNO"].ToString().Trim()}'") > 0)
                        {
                            Ids.Add($"'{d["Id"].ToString()}'");
                        }
                        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[DOC5OUTII] where IsOK='1' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["SourceNO"].ToString().Trim()}'") > 0)
                        {
                            Ids.Add($"'{d["Id"].ToString()}'");
                        }
                    }
                    if (Ids.Count > 0)
                    {
                        //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1' where Id in ({string.Join(",", Ids)})");


                        //tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where Id in ({string.Join(",", Ids)})");
                        //if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC1BuyII] set IsOK='1' where Id in ({nos})"))
                        //{ meg = "無法完成!"; }
                    }
                }
            }
           */

            return meg;
        }


    }
}
