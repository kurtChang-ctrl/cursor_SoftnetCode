using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Net.WebSockets;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class DOC3confirmController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public DOC3confirmController(SNWebSocketService websocket, SFC_Common sfc_Common)
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

        private DOC3confirmService EditService()
        {
            return new DOC3confirmService(Ctrl);
        }

        [HttpPost]
        public string ConfirmDOC3Stock(string keys) //存貨確認  0=ipport,1=單號子項的Id,2=單號子項的Id,,,,,
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
                    string storeNO = "";
                    string storeSpacesNO = "";
                    Dictionary<string, List<string[]>> _poolTEST_NO = new Dictionary<string, List<string[]>>();//暫存檢驗單
                    Dictionary<string, List<string[]>> _poolStore_NO = new Dictionary<string, List<string[]>>();//暫存入庫單
                    string qtyType = "+";
                    foreach (string s in lists)
                    {
                        sql = $"SELECT a.*,b.DOCType,b.DOCNO  FROM SoftNetMainDB.[dbo].[DOC3stockII] as a,SoftNetMainDB.[dbo].[DOCRole] as b where a.Id='{s}' and SUBSTRING(a.DOCNumberNO,1,4)=b.DOCNO and b.ServerId='{_Fun.Config.ServerId}'";//將來where條件要加DOCType=?
                        DataRow dr_DOC3 = db.DB_GetFirstDataByDataRow(sql);
                        if (dr_DOC3 != null)
                        {
                            switch (dr_DOC3["DOCType"].ToString())
                            {
                                case "3"://領類
                                    qtyType = "-";
                                    storeNO = dr_DOC3["OUT_StoreNO"].ToString();
                                    storeSpacesNO = dr_DOC3["OUT_StoreSpacesNO"].ToString();
                                    break;
                                case "4"://入類
                                    qtyType = "+";
                                    storeNO = dr_DOC3["IN_StoreNO"].ToString();
                                    storeSpacesNO = dr_DOC3["IN_StoreSpacesNO"].ToString();
                                    break;
                                case "5"://調撥
                                    #region 先做減倉
                                    qtyType = "-";
                                    storeNO = dr_DOC3["OUT_StoreNO"].ToString();
                                    storeSpacesNO = dr_DOC3["OUT_StoreSpacesNO"].ToString();
                                    sql = $"select Id from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3["PartNO"].ToString()}' and StoreNO='{storeNO}' and StoreSpacesNO='{storeSpacesNO}'";
                                    DataRow tmp_dr2 = db.DB_GetFirstDataByDataRow(sql);
                                    if (tmp_dr2 != null)
                                    {
                                        if (!db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY{qtyType}='{dr_DOC3["QTY"].ToString()}' where ServerId='{_Fun.Config.ServerId}' and Id='{tmp_dr2["Id"].ToString()}'"))
                                        { meg = $"{meg}<br> 單號:{dr_DOC3["DOCNumberNO"].ToString()} 無法完成調撥作業!";continue; }
                                    }
                                    else
                                    {
                                        /*
                                        //無儲位
                                        sql = $"select * from SoftNetMainDB.[dbo].[TotalStock] where PartNO='{dr_DOC3["PartNO"].ToString()}' and StoreNO='{dr_DOC3["StoreNO"].ToString()}'";
                                        tmp_dr2 = db.DB_GetFirstDataByDataRow(sql);
                                        if (tmp_dr2 != null)
                                        {
                                            if (!db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY{qtyType}='{dr_DOC3["QTY"].ToString()}' where Id='{tmp_dr2["Id"].ToString()}'"))
                                            { meg = $"{meg}<br> 單號:{dr_DOC3["DOCNumberNO"].ToString()} 無法完成入庫!"; }
                                        }
                                        else
                                        {
                                            //無料號
                                            string blankId = _Str.NewId('Z');
                                            if (!db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] ([Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{blankId}','{storeNO}','{storeSpacesNO}','{dr_DOC3["PartNO"].ToString()}',{qtyType}{dr_DOC3["QTY"].ToString()})"))
                                            { meg = $"{meg}<br> 單號:{dr_DOC3["DOCNumberNO"].ToString()} 無法完成入庫!"; }
                                            else
                                            {
                                                db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock_Blank_LOG] ([Id],StoreId,[logDate],[SimulationId],QTY,[Remark]) VALUES ('{_Str.NewId('Z')}','{blankId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{dr_DOC3["SimulationId"].ToString()}',{qtyType}{dr_DOC3["QTY"].ToString()},'入庫但無對應倉')");
                                            }
                                        }
                                        */
                                        meg = $"{meg}<br> 單號:{dr_DOC3["DOCNumberNO"].ToString()} 無法完成調撥作業!"; continue;
                                    }
                                    #endregion
                                    qtyType = "+";
                                    storeNO = dr_DOC3["IN_StoreNO"].ToString();
                                    storeSpacesNO = dr_DOC3["IN_StoreSpacesNO"].ToString();
                                    break;
                                default: continue;
                            }

                            int typeTotalTime = 0;
                            string writeSQL = "";
                            if (!dr_DOC3.IsNull("StartTime")) { typeTotalTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName,Convert.ToDateTime(dr_DOC3["StartTime"]), DateTime.Now); }
                            else{ writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                            if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime}{writeSQL} where Id='{s}' and DOCNumberNO='{dr_DOC3["DOCNumberNO"].ToString()}' and IsOK='0'"))
                            {

                                #region 計算單據PP_EfficientDetail
                                string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
                                string partNO = dr_DOC3["PartNO"].ToString();
                                string pp_Name = "";
                                string E_stationNO = "";
                                if (dr_DOC3["SimulationId"].ToString() != "")
                                {
                                    DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC3["SimulationId"].ToString()}'");
                                    pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                                    if (!dr_tmp.IsNull("Source_StationNO") && dr_tmp["Source_StationNO"].ToString() != "")
                                    { E_stationNO = dr_tmp["Source_StationNO"].ToString(); }
                                    else { E_stationNO = dr_tmp["Apply_StationNO"].ToString(); }
                                }
                                DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} CT from SoftNetMainDB.[dbo].[DOC3stockII] where SUBSTRING(Id,2,2)='{_Fun.Config.ServerId}' and SUBSTRING(DOCNumberNO,1,4)='{dr_DOC3["DOCNumberNO"].ToString().Substring(0, 4)}' and PartNO='{dr_DOC3["PartNO"].ToString()}' and CT>0");
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
                                    _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, pp_Name, pp_Name, partNO,"0", partNO, dr_DOC3["DOCNumberNO"].ToString().Substring(0, 4));
                                }
                                #endregion

                                sql = $"select Id from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3["PartNO"].ToString()}' and StoreNO='{storeNO}' and StoreSpacesNO='{storeSpacesNO}'";
                                DataRow tmp_dr2 = db.DB_GetFirstDataByDataRow(sql);
                                if (tmp_dr2 != null)
                                {
                                    if (!db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY{qtyType}='{dr_DOC3["QTY"].ToString()}' where ServerId='{_Fun.Config.ServerId}' and Id='{tmp_dr2["Id"].ToString()}'"))
                                    { meg = $"{meg}<br> 單號:{dr_DOC3["DOCNumberNO"].ToString()} 無法完成入庫!"; }
                                }
                                else
                                {
                                    //無儲位
                                    sql = $"select * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3["PartNO"].ToString()}' and StoreNO='{storeNO}'";
                                    tmp_dr2 = db.DB_GetFirstDataByDataRow(sql);
                                    if (tmp_dr2 != null)
                                    {
                                        if (!db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY{qtyType}='{dr_DOC3["QTY"].ToString()}' where ServerId='{_Fun.Config.ServerId}' and Id='{tmp_dr2["Id"].ToString()}'"))
                                        { meg = $"{meg}<br> 單號:{dr_DOC3["DOCNumberNO"].ToString()} 無法完成入庫!"; }
                                    }
                                    else
                                    {
                                        #region 查找適合庫儲別
                                        _SFC_Common.SelectINStore(db, dr_DOC3["PartNO"].ToString(), ref storeNO, ref storeSpacesNO, dr_DOC3["DOCNO"].ToString());
                                        #endregion
                                        if (storeNO != "")
                                        {
                                            tmp_dr2 = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{storeNO}'");
                                            if (tmp_dr2 != null)
                                            {
                                                string blankId = _Str.NewId('Z');
                                                if (!db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY],Class) VALUES ('{_Fun.Config.ServerId}','{blankId}','{storeNO}','{storeSpacesNO}','{dr_DOC3["PartNO"].ToString()}',{qtyType}{dr_DOC3["QTY"].ToString()},'{tmp_dr2["Class"].ToString()}')"))
                                                { meg = $"{meg}<br> 單號:{dr_DOC3["DOCNumberNO"].ToString()} 無法完成入庫!"; }
                                                else
                                                {
                                                    db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock_Blank_LOG] (ServerId,[Id],StoreId,[logDate],[SimulationId],QTY,[Remark]) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{blankId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{dr_DOC3["SimulationId"].ToString()}',{qtyType}{dr_DOC3["QTY"].ToString()},'入庫但無對應倉')");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return meg;
        }
    }
}
