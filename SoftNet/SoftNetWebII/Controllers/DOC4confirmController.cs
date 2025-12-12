using Base.Models;
using Base.Services;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using System;
using BaseApi.Controllers;
using Base;

namespace SoftNetWebII.Controllers
{
    public class DOC4confirmController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public DOC4confirmController(SNWebSocketService websocket, SFC_Common sfc_Common)
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

        private DOC4confirmService EditService()
        {
            return new DOC4confirmService(Ctrl);
        }

        [HttpPost]
        public string ConfirmDOC1Buy(string keys) //委外確認  0=ipport,1=單號子項的Id,2=單號子項的Id,,,,,
        {
            string meg = "";
            var br = _Fun.GetBaseUser();
            if (keys == null || br == null || !br.IsLogin || br.UserNO.Trim() == "")
            { meg = "作業失敗, 畫面已逾時, 請登出餅重新登入."; return meg; }
            string top_flag = $" TOP {_Fun.Config.AdminKey03} ";

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
                    DataRow dr_DOC = null;
                    DataRow dr_tmp = null;
                    DataRow dr_tag = null;
                    foreach (string s in lists)
                    {
                        DataRow dr_DOC4ProductionII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC4ProductionII] where Id='{s}' and IsOK='0'");
                        dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_DOC4ProductionII["StoreNO"].ToString()}'");
                        if (dr_tag["Class"].ToString() == "虛擬倉")
                        {
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

                                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set Detail_QTY+={dr_DOC["QTY"].ToString()} where SimulationId='{dr_DOC["SimulationId"].ToString()}'"))
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
                                                        #region 查找適合入庫儲別
                                                        _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                                        #endregion
                                                        #region 無倉紀錄, 加空倉
                                                        _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, tmp_int, "", "", $"工單:{d["DOCNumberNO"].ToString()} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO, true);
                                                        #endregion
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
                        else
                        {
                            int docQTY = int.Parse(dr_DOC4ProductionII["QTY"].ToString());
                            int out_qty = docQTY;
                            dr_tag = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr_DOC4ProductionII["StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC4ProductionII["StoreSpacesNO"].ToString()}'");

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
            }

            return meg;
        }



    }
}
