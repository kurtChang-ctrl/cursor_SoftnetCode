using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using DocumentFormat.OpenXml.Office2010.Excel;
//using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Models;
using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    //[XgProgAuth]
    public class STViewController : ApiCtrl
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public STViewController(SNWebSocketService websocket, SFC_Common sfc_Common)
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
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData("SELECT FactoryName FROM [dbo].[Factory]");
                if (dt != null && dt.Rows.Count > 0)
                {
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
        public string Update_PP_WorkOrder_Settlement(string id) //報工 StationNO,OrderNO,IndexSN,outQTY,failQTY,OPNO,ipport
        {
            var br = _Fun.GetBaseUser();
            if (id == null || br == null || !br.IsLogin || br.UserNO.Trim() == "")
            { return $"作業失敗, 畫面已逾時, 請關閉網頁瀏覽器,  並重新操作.";  }
            string key_OPNO = br.UserNO;
            //###???若此處有改 TMM Service 55ㄝ要改
            string sql = "";
            DateTime startTime = DateTime.Now;
            string meg = "";
            {
                //###???若此處有改 TMM Service 55 ㄝ要改
                //###???若此處有改 TagResult_EnterKeyController ㄝ要改
                string[] data_id = id.Split(',');
                LabelWork keys = new LabelWork();
                keys.Station = data_id[0];
                keys.OrderNO = data_id[1];
                keys.IndexSN = data_id[2];
                keys.LocalIPPort = data_id[6];
                int tryINT = 0;
                decimal defaultCT = 0;
                if (int.TryParse(data_id[3], out tryINT)) { keys.OKQTY = tryINT; }
                else { return "數量非數字,無法報工."; }
                if (int.TryParse(data_id[4], out tryINT)) { keys.FailQTY = tryINT; }
                else { return "數量非數字,無法報工."; }
                if (int.TryParse(data_id[7], out tryINT)) { defaultCT = tryINT; }
                keys.OPNO= data_id[5];
                try
                {
                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                    {
                        DataRow dr_StationNO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}'");
                        string[] data = new string[7] { keys.Station, keys.OrderNO, keys.IndexSN, keys.OKQTY.ToString(), keys.FailQTY.ToString(), keys.OPNO, keys.LocalIPPort };
                        int outQTY = int.Parse(data[3]);//報工良品數量
                        int failQTY = int.Parse(data[4]);//報工不良品數量
                        DataRow dr_Manufacture = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[0]}'");
                        keys.SimulationId = dr_Manufacture["SimulationId"].ToString();
                        #region 檢查網頁來源資料
                        if (keys.SimulationId=="") { return $"查無計畫碼,無法報工."; }
                        if (dr_Manufacture.IsNull("StartTime")) { return $"{data[1]} 工站沒有開工紀錄, 請先執行開工, 才能報工."; }

                        DataRow dr_WO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}'");
                        if (dr_WO == null) { return  $"查無 {data[0]} 工單資料紀錄, 請聯繫系統管理者."; }
                        sql = $"SELECT * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}' AND StationNO='{data[0]}' AND IndexSN={data[2]}";
                        DataRow dr = db.DB_GetFirstDataByDataRow(sql);
                        if (dr == null) { return $"查無 PP_WorkOrder_Settlement 資料紀錄, 請聯繫系統管理者.";  }
                        if (!dr.IsNull("StartTime") && dr["StartTime"].ToString().Trim() != "")
                        { startTime = Convert.ToDateTime(dr["StartTime"]); }
                        #endregion

                        #region 計算CT
                        decimal ct = 0;
                        if (defaultCT != 0)
                        { ct = defaultCT; }
                        else
                        {
                            DateTime rRemarkTimeS = startTime;
                            if (dr_Manufacture.IsNull("RemarkTimeS"))
                            { rRemarkTimeS = Convert.ToDateTime(dr_Manufacture["StartTime"]); }
                            else
                            { rRemarkTimeS = Convert.ToDateTime(dr_Manufacture["RemarkTimeS"]); }
                            #region 先查相同人員與站與工單, 是否報工過
                            DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and IndexSN={keys.IndexSN} and OrderNO='{keys.OrderNO}' and OperateType like '%報工%' and OP_NO like '%{keys.OPNO}%'");
                            if (tmp_dr != null && Convert.ToDateTime(tmp_dr["LOGDateTime"]) > rRemarkTimeS) { rRemarkTimeS = Convert.ToDateTime(tmp_dr["LOGDateTime"]); }
                            #endregion
                            if (dr_Manufacture.IsNull("RemarkTimeE") || rRemarkTimeS >= Convert.ToDateTime(dr_Manufacture["RemarkTimeE"]))
                            {
                                ct = _SFC_Common.TimeCompute2Seconds(rRemarkTimeS, DateTime.Now) / (keys.OKQTY + keys.FailQTY);
                                if (ct <= 0) { ct = 1; }
                            }
                            else
                            {
                                ct = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), rRemarkTimeS, Convert.ToDateTime(dr_Manufacture["RemarkTimeE"])) / (keys.OKQTY + keys.FailQTY);
                            }
                        }
                        decimal ct_log = ct < 1 ? 0 : ct;

                        int ops = dr_Manufacture["OP_NO"].ToString().Split(';').Length;
                        if (ops > 1) { ct = ct / ops; }
                        #endregion

                        #region 寫SFC_StationDetail
                        string partNO = dr_Manufacture["PartNO"].ToString();
                        string old_InTime = "";
                        string old_OutTime = "";
                        string old_ProductFinishedQty = "0";
                        string old_ProductFailedQty = "0";
                        int OP_Count = 1;
                        string logTime = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff");

                        DataRow dr_StationDetail = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail] WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}' AND StationNO='{data[0]}' AND IndexSN={data[2]}");
                        if (dr_StationDetail == null)
                        {
                            //###???PP_Name暫時
                            sql = string.Format(
                                @"INSERT INTO SoftNetLogDB.[dbo].[SFC_StationDetail] (
                                                [LOGDateTime],[Master_PP_Name],[PP_Name],[OP_NO],[StationNO],
                                                [IndexSN],[IndexSN_Merge],[OrderNO],[PartNO],[InTime],
                                                [OutTime],[CycleTime],[InputFlag],[OutputFlag],[FailFlag],
                                                [Station_Type],[ProductFinishedQty],[ProductFailedQty],[SerialNO],[RMSName],SimulationId,ServerId) VALUES (
                                                '{0}','{1}','{2}','{3}','{4}',{5},'{6}','{7}','{8}','{9}',
                                                '{10}',{11},'{12}','{13}','{14}','{15}',{16},{17},'','{18}','{19}','{20}')",
                                logTime,
                                dr["PP_Name"].ToString(),
                                dr["PP_Name"].ToString(),//###???暫時換dr["Sub_PP_Name"].ToString(),
                                data[5],//OPNO
                                data[0],
                                data[2],
                                dr["IndexSN_Merge"].ToString(),
                                data[1],
                                dr_WO["PartNO"].ToString(),
                                startTime.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                ct.ToString(),
                                (outQTY + failQTY) > 0 ? 1 : 0,
                                outQTY > 0 ? 1 : 0,
                                failQTY > 0 ? 1 : 0,
                                dr_StationNO["Station_Type"].ToString(),
                                outQTY.ToString(),
                                failQTY.ToString(), dr_StationNO["RMSName"].ToString(), keys.SimulationId, _Fun.Config.ServerId);
                        }
                        else
                        {
                            logTime = Convert.ToDateTime(dr_StationDetail["LOGDateTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff");
                            old_InTime = Convert.ToDateTime(dr_StationDetail["InTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff");
                            old_OutTime = Convert.ToDateTime(dr_StationDetail["OutTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff");
                            old_ProductFinishedQty = dr_StationDetail["ProductFinishedQty"].ToString();
                            old_ProductFailedQty = dr_StationDetail["ProductFailedQty"].ToString();
                            if (int.Parse(dr_StationDetail["CycleTime"].ToString()) != 0) { ct = (ct + int.Parse(dr_StationDetail["CycleTime"].ToString())) > 0 ? Math.Round((ct + int.Parse(dr_StationDetail["CycleTime"].ToString())) / 2) : ct; }
                            if (ct < 1) { ct = 0; }
                            string sId = "";
                            if (keys.SimulationId == "")
                            { sId = ""; }
                            else
                            { sId = $"and SimulationId='{keys.SimulationId}'"; }
                            sql = string.Format(
                                    @"UPDATE SoftNetLogDB.[dbo].[SFC_StationDetail] 
                                                SET [ProductFinishedQty]+={0}, [ProductFailedQty]+={1},
                                                [InTime]='{2}',[OutTime]='{3}',[CycleTime]={4} 
                                                WHERE ServerId='{9}' and OrderNO = '{5}' AND StationNO = '{6}' AND IndexSN={7} {8}",
                                    outQTY,
                                    failQTY,
                                    startTime.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                    ct.ToString(),
                                    data[1],
                                    data[0], data[2], sId, _Fun.Config.ServerId);
                        }
                        #endregion

                        if (db.DB_SetData(sql))
                        {
                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView','智慧報工','{dr_Manufacture["PP_Name"].ToString()}','{keys.Station}','{dr_Manufacture["PartNO"].ToString()}','{dr_Manufacture["OrderNO"].ToString()}','{key_OPNO}',{dr_Manufacture["IndexSN"].ToString()})");

                            #region 更新標籤累計量
                            if (outQTY != 0)
                            {
                                DataRow totalData = _Fun.GetAvgCTWTandTotalOutput(db, false, data[1], data[0], data[2]);
                                string dis_DetailQTY = "0";
                                if (totalData != null)
                                {
                                    dis_DetailQTY = totalData["TotalOutput"].ToString().Trim();
                                    totalData = db.DB_GetFirstDataByDataRow($"select b.* from SoftNetMainDB.[dbo].[Manufacture] as a join SoftNetMainDB.[dbo].[LabelStateINFO] as b on b.macID=a.Config_macID where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{data[0]}'");
                                    if (totalData != null && dis_DetailQTY != totalData["QTY"].ToString().Trim())
                                    {
                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set QTY='{dis_DetailQTY}',IsUpdate='0' where ServerId='{_Fun.Config.ServerId}' and macID='{totalData["macID"].ToString()}'");
                                    }
                                }
                            }
                            #endregion

                            #region 計算效能 PP_EfficientDetail處理
                            {
                                List<double> allCT = new List<double>();//list for all avg value
                                string top_flag = "";
                                try
                                {
                                    if (_Fun.Config.AdminKey03 != 0) { top_flag = $" TOP {_Fun.Config.AdminKey03} "; }
                                    DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} PP_Name,StationNO,PartNO as Sub_PartNO,CycleTime,WaitTime,(EditFinishedQty+EditFailedQty) as QTY from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG]
                                                    where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and PartNO='{dr_Manufacture["PartNO"].ToString()}' and PP_Name='{dr_Manufacture["PP_Name"].ToString()}' and IndexSN={dr_Manufacture["IndexSN"].ToString()} and EditFinishedQty!=0 and CycleTime!=0");
                                    if (dt_Efficient != null && dt_Efficient.Rows.Count > 0)
                                    {
                                        double efficient_CT = 0;
                                        foreach (DataRow dr_tmp in dt_Efficient.Rows)
                                        {
                                            if (_Fun.Config.AdminKey14)
                                            { efficient_CT = double.Parse(dr_tmp["CycleTime"].ToString()) + double.Parse(dr_tmp["WaitTime"].ToString()); }
                                            else
                                            { efficient_CT = double.Parse(dr_tmp["CycleTime"].ToString()); }
                                            for (int tmp01 = 1; tmp01 <= (int)dr_tmp["QTY"]; tmp01++)//工單數目若為2 需算作兩筆
                                            {
                                                allCT.Add(efficient_CT);
                                            }
                                        }
                                        if (allCT.Count > 0)
                                        {
                                            _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, keys.Station, dr_WO["PP_Name"].ToString(), dr_Manufacture["PP_Name"].ToString(), dr_Manufacture["IndexSN"].ToString(), dr_WO["PartNO"].ToString(), dr_Manufacture["PartNO"].ToString(), "");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs 計算效能PP_EfficientDetail處理 Exception: {ex.Message} {ex.StackTrace}", true);
                                }
                            }
                            #endregion

                            #region log SFC_StationDetail_ChangeLOG紀錄
                            int reportTime = 0;
                            if (dr_StationDetail != null)
                            {
                                #region 計算上一次與現在時間差
                                DataRow d2 = db.DB_GetFirstDataByDataRow($"SELECT LOGDateTimeID FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(dr_StationDetail["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by LOGDateTime,LOGDateTimeID desc");
                                if (d2 != null)
                                { reportTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), Convert.ToDateTime(d2["LOGDateTimeID"]), DateTime.Now); }
                                #endregion
                            }
                            string wsid = "NULL";
                            if (keys.SimulationId != "") { wsid = $"'{keys.SimulationId}'"; }
                            OP_Count = data[5].Split(";").Count();
                            if (OP_Count <= 0) { OP_Count = 1; }
                            string LOGDateTimeID = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                            #region 查詢PP_EfficientDetail
                            string eCT = "0";
                            string upperCT = "0";
                            string lowerCT = "0";
                            DataRow dr_tmp_ct = db.DB_GetFirstDataByDataRow($"select AVG(EfficientCycleTime) as ECT,AVG(SD_UpperLimit) as UpperCT,AVG(SD_LowerLimit) as LowerCT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_Manufacture["PartNO"]}' and StationNO='{dr_Manufacture["StationNO"]}' and PP_Name='{dr_Manufacture["PP_Name"].ToString()}' and IndexSN={dr_Manufacture["IndexSN"].ToString()} and DOCNO=''");
                            if (dr_tmp_ct != null)
                            {
                                if (!dr_tmp_ct.IsNull("ECT") && dr_tmp_ct["ECT"].ToString() != "") { eCT = dr_tmp_ct["ECT"].ToString(); }
                                if (!dr_tmp_ct.IsNull("UpperCT") && dr_tmp_ct["UpperCT"].ToString() != "") { upperCT = dr_tmp_ct["UpperCT"].ToString(); }
                                if (!dr_tmp_ct.IsNull("LowerCT") && dr_tmp_ct["LowerCT"].ToString() != "") { lowerCT = dr_tmp_ct["LowerCT"].ToString(); }
                            }
                            #endregion
                            sql = $@"INSERT INTO SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] (LOGDateTime,LOGDateTimeID,OLD_InTime,OLD_OutTime,EditFinishedQty,EditFailedQty,OLD_ProductFinishedQty,OLD_ProductFailedQty,OP_Count,OP_NO,ServerId,StationNO,ReportTime,PartNO,SimulationId,PP_Name,IndexSN,CycleTime,ECT,LowerCT,UpperCT)
                                            VALUES ('{logTime}',
                                            '{LOGDateTimeID}',
                                            '{old_InTime}',
                                            '{old_OutTime}',
                                            {outQTY},
                                            {failQTY},
                                            {old_ProductFinishedQty},
                                            {old_ProductFailedQty},
                                            {OP_Count.ToString()},
                                            '{data[5]}','{_Fun.Config.ServerId}','{data[0]}',{reportTime.ToString()},'{partNO}',{wsid},'{dr["PP_Name"].ToString()}',{dr["IndexSN"].ToString()},{ct_log.ToString()},{eCT},{lowerCT},{upperCT})";
                            if (db.DB_SetData(sql))
                            {
                                //###???
                            }
                            #endregion

                            #region 修正工站開始日期
                            if (dr_Manufacture["State"].ToString() == "1")
                            { db.DB_SetData($"update SoftNetMainDB.[dbo].[Manufacture] set RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[0]}'"); }
                            #endregion

                            _SFC_Common.Update_PP_WorkOrder_Settlement(db, data[1], keys.SimulationId);

                            //###??? 不良數量尚未處裡
                            if (keys.SimulationId != "")
                            {
                                bool isNeedQTY_OK = false;//判斷本站數量已足夠
                                DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{keys.SimulationId}'");
                                string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();

                                string in_NO = "AC01";//###??? 暫時寫死領料單別
                                string inOK_NO = "BC01";//###??? 暫時寫死入庫單別

                                DataRow dr_APS_PartNOTimeNote = null;
                                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{data[1].Trim()}',Detail_QTY+={data[3]} where SimulationId='{keys.SimulationId}'"))
                                {
                                    dr_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SimulationId}'");
                                    if ((int.Parse(dr_APS_PartNOTimeNote["Detail_QTY"].ToString()) + int.Parse(dr_APS_PartNOTimeNote["Detail_Fail_QTY"].ToString()) - int.Parse(dr_APS_PartNOTimeNote["NeedQTY"].ToString())) >= 0)
                                    { isNeedQTY_OK = true; }
                                }

                                //尋找相關BOM原物料
                                #region 扣Keep量 與 處理領料單單據
                                DataTable dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_APS_Simulation["NeedId"].ToString()}' and Apply_PP_Name='{dr_Manufacture["PP_Name"].ToString()}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and IndexSN={data[2]} order by PartSN desc");
                                if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                                {
                                    string docNumberNO = "";
                                    foreach (DataRow d in dt_APS_Simulation.Rows)
                                    {
                                        #region 處裡移轉量 APS_PartNOTimeNote
                                        if (!d.IsNull("Source_StationNO") && (d["Class"].ToString() == "4" || d["Class"].ToString() == "5"))
                                        {
                                            if (d["PartSN"].ToString() == "0")
                                            {
                                                #region 工單最後一站預開入庫單  
                                                string tmp_no = "";
                                                string in_StoreNO = "";
                                                string in_StoreSpacesNO = "";
                                                DataTable tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 Order by b.StoreOrder");
                                                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                                {
                                                    int tmp_int = outQTY;
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
                                                                _SFC_Common.Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO);
                                                                tmp_int = 0;
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                _SFC_Common.Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_01, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO);
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
                                                                _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO);
                                                            }
                                                            else
                                                            {
                                                                #region 查找適合入庫儲別
                                                                _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                                                #endregion
                                                                #region 無倉紀錄, 加空倉
                                                                _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO);
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
                                                    _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, int.Parse(data[3]), "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, br.UserNO);
                                                    #endregion
                                                }

                                                sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{tmp_no}',Next_StoreQTY+={data[3]} where SimulationId='{d["SimulationId"].ToString()}'";

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
                                            else
                                            {
                                                #region 非最後一站
                                                int tmp_int = (int.Parse(d["BOMQTY"].ToString()) * outQTY);
                                                DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d["SimulationId"].ToString()}'");
                                                int wr_next_StationQTY = 0;
                                                #region 檢查上一階之前是否有入退庫數量(BC01)
                                                if (int.Parse(tmp["Next_StoreQTY"].ToString()) > 0)
                                                {
                                                    int store_tmp = tmp_int;
                                                    DataTable dt_Store = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{tmp["Store_DOCNumberNO"].ToString()}' order by ArrivalDate");
                                                    if (dt_Store != null && dt_Store.Rows.Count > 0)
                                                    {
                                                        int wrQTY = 0;
                                                        foreach (DataRow dr_DOC3stockII in dt_Store.Rows)
                                                        {
                                                            if (store_tmp <= 0) { break; }
                                                            if (store_tmp >= int.Parse(dr_DOC3stockII["QTY"].ToString()))
                                                            {
                                                                store_tmp -= int.Parse(dr_DOC3stockII["QTY"].ToString());
                                                                wrQTY = int.Parse(dr_DOC3stockII["QTY"].ToString());
                                                                wr_next_StationQTY += wrQTY;
                                                            }
                                                            else
                                                            {
                                                                //拆單處置
                                                                wrQTY = store_tmp;
                                                                wr_next_StationQTY += store_tmp;
                                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set QTY={store_tmp} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}'");
                                                                db.DB_SetData($@"INSERT INTO [dbo].[DOC3stockII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],ArrivalDate) VALUES 
                                                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{dr_DOC3stockII["DOCNumberNO"].ToString()}','{dr_DOC3stockII["PartNO"].ToString()}',{dr_DOC3stockII["Price"].ToString()},'{dr_DOC3stockII["Unit"].ToString()}',{(int.Parse(dr_DOC3stockII["QTY"].ToString()) - store_tmp).ToString()}
                                                                                                ,'{dr_DOC3stockII["Remark"].ToString()}','{dr_DOC3stockII["SimulationId"].ToString()}','{dr_DOC3stockII["IsOK"].ToString()}','{dr_DOC3stockII["IN_StoreNO"].ToString()}','{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}','{dr_DOC3stockII["OUT_StoreNO"].ToString()}','{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}','{Convert.ToDateTime(dr_DOC3stockII["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                                store_tmp = 0;
                                                            }
                                                            if (!bool.Parse(dr_DOC3stockII["IsOK"].ToString()))
                                                            {
                                                                #region 寫入庫存
                                                                if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                                                                { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={wrQTY} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["IN_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}'"); }
                                                                else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={wrQTY} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'"); }
                                                                #endregion
                                                                #region 計算單據CT,平均,有效, 寫SFC_StationProjectDetail
                                                                int typeTotalTime = 0;
                                                                string writeSQL = "";
                                                                if (!dr_DOC3stockII.IsNull("StartTime")) { typeTotalTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC3stockII["StartTime"].ToString()), DateTime.Now); }
                                                                else { writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime}{writeSQL} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}' and IsOK='0'");
                                                                string efficient_partNO = dr_DOC3stockII["PartNO"].ToString();
                                                                string efficient_pp_Name = "";
                                                                string E_stationNO = "";
                                                                if (dr_DOC3stockII["SimulationId"].ToString() != "")
                                                                {
                                                                    DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC3stockII["SimulationId"].ToString()}'");
                                                                    efficient_pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                                                                    if (!dr_tmp.IsNull("Source_StationNO") && dr_tmp["Source_StationNO"].ToString() != "")
                                                                    { E_stationNO = dr_tmp["Source_StationNO"].ToString(); }
                                                                    else { E_stationNO = dr_tmp["Apply_StationNO"].ToString(); }
                                                                }
                                                                DataTable dt_Efficient = db.DB_GetData($@"select TOP {_Fun.Config.AdminKey03} CT from SoftNetMainDB.[dbo].[DOC3stockII] where SUBSTRING(Id,2,2)='{_Fun.Config.ServerId}' and SUBSTRING(DOCNumberNO,1,4)='{dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4)}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and CT>0");
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
                                                                    _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, efficient_pp_Name, efficient_pp_Name,"0", efficient_partNO, efficient_partNO, dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4));
                                                                }
                                                                #endregion
                                                            }
                                                            //開領料單
                                                            string tmpDOCNO = dr_DOC3stockII["DOCNumberNO"].ToString();
                                                            DataRow tmp_store = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and IN_StoreNO!=''");
                                                            if (tmp_store != null)
                                                            {
                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                _SFC_Common.Create_DOC3stock(db, d, tmp_store["IN_StoreNO"].ToString(), tmp_store["IN_StoreSpacesNO"].ToString(), "", "", in_NO, wrQTY, "", "", $"{stationno}Id:{d["SimulationId"].ToString()} 入庫後再領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmpDOCNO, br.UserNO);
                                                            }
                                                        }
                                                        if (wr_next_StationQTY > 0)
                                                        {
                                                            //將入庫數量Next_StoreQTY扣除回寫Next_StationQTY
                                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set Next_StoreQTY-={wr_next_StationQTY} where SimulationId='{d["SimulationId"].ToString()}'");
                                                        }
                                                    }
                                                }
                                                #endregion
                                                int detail_QTY = int.Parse(tmp["Detail_QTY"].ToString());
                                                int next_StationQTY = int.Parse(tmp["Next_StationQTY"].ToString());

                                                int next_next_detail_QTY = 0;
                                                #region 檢查下一階是否有偷先報工未移轉
                                                DataRow dr_next_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SimulationId}'");
                                                next_next_detail_QTY = int.Parse(dr_next_APS_PartNOTimeNote["Detail_QTY"].ToString());
                                                if (next_next_detail_QTY > 0)
                                                {
                                                    //detail_QTY - next_StationQTY=上一階數量
                                                    if ((next_StationQTY + tmp_int) < next_next_detail_QTY) { next_next_detail_QTY -= (next_StationQTY + tmp_int); }
                                                    else { next_next_detail_QTY = 0; }
                                                }
                                                #endregion

                                                if ((detail_QTY - next_StationQTY) >= tmp_int)
                                                {
                                                    //在製移轉 
                                                    sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Next_APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}',Next_StationQTY+={tmp_int + next_next_detail_QTY} where SimulationId='{d["SimulationId"].ToString()}'";
                                                    if (db.DB_SetData(sql))
                                                    {
                                                        #region 處理工站移轉時間
                                                        /*
                                                        tmp = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetLogDB.[dbo].[SFC_StationDetail] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["Source_StationNO"].ToString()}' and IndexSN='{d["Source_StationNO_IndexSN"].ToString()}' order by LOGDateTime desc");
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
                                                }
                                                else
                                                {
                                                    bool is_run = true;
                                                    //將在製移轉 剩餘
                                                    tmp_int -= (detail_QTY - next_StationQTY);
                                                    sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Next_APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}',Next_StationQTY={tmp["Detail_QTY"].ToString()} where SimulationId='{d["SimulationId"].ToString()}'";
                                                    if (db.DB_SetData(sql) && int.Parse(tmp["Detail_QTY"].ToString()) > 0)
                                                    {
                                                        #region 處理工站移轉時間
                                                        /*
                                                        tmp = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetLogDB.[dbo].[SFC_StationDetail] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["Source_StationNO"].ToString()}' and IndexSN='{d["Source_StationNO_IndexSN"].ToString()}' order by LOGDateTime desc");
                                                        if (tmp != null)
                                                        {
                                                            int NextStationTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), Convert.ToDateTime(tmp["LOGDateTime"]), DateTime.Now);
                                                            if (NextStationTime > 0)
                                                            {
                                                                #region 回寫上一站報工移轉時間
                                                                DataTable dt_SFC_StationDetail_ChangeLOG = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(tmp["LOGDateTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}' and StationNO='{tmp["StationNO"].ToString()}' and PartNO='{tmp["PartNO"].ToString()}' order by LOGDateTime");
                                                                if (dt_SFC_StationDetail_ChangeLOG != null && dt_SFC_StationDetail_ChangeLOG.Rows.Count > 0)
                                                                {
                                                                    for (int i = 1; i <= dt_SFC_StationDetail_ChangeLOG.Rows.Count; i++)
                                                                    {
                                                                        DataRow dLOG = dt_SFC_StationDetail_ChangeLOG.Rows[i];
                                                                        if (i == dt_SFC_StationDetail_ChangeLOG.Rows.Count && int.Parse(dLOG["NextStationTime"].ToString()) != 0)
                                                                        {

                                                                            db.DB_SetData($@"UPDATE SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] SET NextStationTime={((NextStationTime + int.Parse(dLOG["NextStationTime"].ToString())) / 2)} where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(dLOG["NextStationTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' 
                                                                                                and LOGDateTimeID='{Convert.ToDateTime(dLOG["NextStationTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and StationNO='{dLOG["StationNO"].ToString()}' and PartNO='{dLOG["PartNO"].ToString()}'");
                                                                        }
                                                                        else
                                                                        {
                                                                            if (int.Parse(dLOG["NextStationTime"].ToString()) == 0)
                                                                            {
                                                                                db.DB_SetData($@"UPDATE SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] SET NextStationTime={NextStationTime} where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(dLOG["NextStationTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' 
                                                                                                and LOGDateTimeID='{Convert.ToDateTime(dLOG["NextStationTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and StationNO='{dLOG["StationNO"].ToString()}' and PartNO='{dLOG["PartNO"].ToString()}'");
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                #endregion

                                                                #region 統計 PP_EfficientDetail 移轉時間
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
                                                                        _WebSocket.SfcTimerloopthread_Tick_Efficient(db, allCT, d["StationNO"].ToString(), d["Apply_PP_Name"].ToString(), d["Apply_PP_Name"].ToString(), d["PartNO"].ToString(), d["PartNO"].ToString(), "", data[0]);
                                                                    }
                                                                }
                                                                #endregion
                                                            }
                                                        }
                                                        */
                                                        #endregion

                                                    }

                                                    #region 先檢查是否已有單據, 且已移轉多少量
                                                    int stockQTY = 0;
                                                    tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) as qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}'");
                                                    if (tmp != null && !tmp.IsNull("qty"))
                                                    {
                                                        stockQTY = int.Parse(tmp["qty"].ToString());
                                                        if (stockQTY <= (detail_QTY + next_StationQTY)) { stockQTY = 0; }
                                                    }
                                                    if ((stockQTY - tmp_int) >= 0)
                                                    {
                                                        is_run = false;
                                                    }
                                                    else
                                                    {
                                                        tmp_int = tmp_int - stockQTY;
                                                    }
                                                    #endregion

                                                    if (tmp_int <= 0) { is_run = false; }
                                                    if (is_run)
                                                    {
                                                        DataTable tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 order by b.StoreOrder");
                                                        if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                                        {
                                                            #region 有計畫Keep量  by StoreOrder順序扣
                                                            foreach (DataRow d2 in tmp_dt.Rows)
                                                            {
                                                                if (tmp_int > 0)
                                                                {
                                                                    if (int.Parse(d2["KeepQTY"].ToString()) >= tmp_int)
                                                                    {
                                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY-={tmp_int} where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                        tmp_int = 0;
                                                                        break;
                                                                    }
                                                                    else
                                                                    {
                                                                        int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                                        //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_01} where Id='{d2["Id"].ToString()}'");
                                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                        tmp_int -= tmp_01;
                                                                    }
                                                                }
                                                            }
                                                            if (tmp_int > 0)
                                                            {
                                                                #region 計畫量不夠扣, 扣實體倉
                                                                DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                                {
                                                                    foreach (DataRow d2 in tmp_dt2.Rows)
                                                                    {
                                                                        if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                        {
                                                                            //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_int} where Id='{d2["Id"].ToString()}'");
                                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                            _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                            tmp_int = 0;
                                                                            break;
                                                                        }
                                                                        else
                                                                        {
                                                                            int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                            if (tmp_01 != 0)
                                                                            {
                                                                                //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_01} where Id='{d2["Id"].ToString()}'");
                                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                                _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                                tmp_int -= tmp_01;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                #endregion

                                                                #region 實體倉不購扣, 扣空倉
                                                                if (tmp_int > 0)
                                                                {
                                                                    #region 查找適合出庫儲別
                                                                    string out_StoreNO = "";
                                                                    string out_StoreSpacesNO = "";
                                                                    _SFC_Common.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                                    #endregion
                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                    _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                }
                                                                #endregion
                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            #region 沒計畫量, 扣實體倉
                                                            DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow d2 in tmp_dt2.Rows)
                                                                {
                                                                    if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                    {
                                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                        tmp_int = 0;
                                                                        break;
                                                                    }
                                                                    else
                                                                    {
                                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                        if (tmp_01 != 0)
                                                                        {
                                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                            _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                            tmp_int -= tmp_01;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            if (tmp_int > 0)
                                                            {
                                                                #region 查找適合庫儲別
                                                                string out_StoreNO = "";
                                                                string out_StoreSpacesNO = "";
                                                                _SFC_Common.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                                #endregion

                                                                #region 實體倉不購扣, 扣空倉
                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                #endregion
                                                            }
                                                            #endregion
                                                        }
                                                    }
                                                }
                                                #endregion
                                            }
                                        }
                                        else
                                        {
                                            bool is_run = true;
                                            #region 原物料 扣庫存帳 TotalStock,TotalStockII
                                            int tmp_int = (int.Parse(d["BOMQTY"].ToString()) * outQTY);
                                            #region 先檢查是否已有單據, 且已移轉多少量
                                            int detailQTY = tmp_int;
                                            int stockQTY = 0;
                                            DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO!=''");
                                            if (tmp != null)
                                            {
                                                docNumberNO = tmp["DOCNumberNO"].ToString();
                                                detailQTY += (int.Parse(tmp["Detail_QTY"].ToString()) + int.Parse(tmp["Next_StationQTY"].ToString()) + int.Parse(tmp["Next_StoreQTY"].ToString()));
                                                tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) as qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{docNumberNO}'");
                                                if (tmp != null && !tmp.IsNull("qty"))
                                                {
                                                    stockQTY = int.Parse(tmp["qty"].ToString());
                                                }
                                                if ((stockQTY - detailQTY) >= 0)
                                                {
                                                    is_run = false;
                                                }
                                                else
                                                {
                                                    tmp_int = detailQTY - stockQTY;
                                                }
                                            }
                                            else
                                            {
                                                tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) as qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}'");
                                                if (tmp != null && !tmp.IsNull("qty"))
                                                {
                                                    stockQTY = int.Parse(tmp["qty"].ToString());
                                                }
                                                if ((stockQTY - detailQTY) >= 0)
                                                {
                                                    is_run = false;
                                                    tmp = db.DB_GetFirstDataByDataRow($"select top 1 DOCNumberNO from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}'");
                                                    if (tmp != null) { docNumberNO = tmp["DOCNumberNO"].ToString(); }
                                                }
                                                else
                                                {
                                                    tmp_int = detailQTY - stockQTY;
                                                }
                                            }
                                            #endregion

                                            if (tmp_int <= 0) { is_run = false; }
                                            int in_APS_PartNOTimeNote_QTY = tmp_int;

                                            if (is_run)
                                            {
                                                DataTable tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 order by b.StoreOrder");
                                                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                                {
                                                    #region 有計畫Keep量  by StoreOrder順序扣
                                                    foreach (DataRow d2 in tmp_dt.Rows)
                                                    {
                                                        if (tmp_int > 0)
                                                        {
                                                            if (int.Parse(d2["KeepQTY"].ToString()) >= tmp_int)
                                                            {
                                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY-={tmp_int} where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                tmp_int = 0;
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                tmp_int -= tmp_01;
                                                            }
                                                        }
                                                    }
                                                    if (tmp_int > 0)
                                                    {
                                                        #region 有計畫量不夠扣, 扣實體倉
                                                        DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                        if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                        {
                                                            foreach (DataRow d2 in tmp_dt2.Rows)
                                                            {
                                                                if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                {
                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                    _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                    tmp_int = 0;
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                    if (tmp_01 != 0)
                                                                    {
                                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                        tmp_int -= tmp_01;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        #endregion

                                                        #region 實體倉不購扣, 扣空倉
                                                        if (tmp_int > 0)
                                                        {
                                                            #region 查找適合庫儲別
                                                            string out_StoreNO = "";
                                                            string out_StoreSpacesNO = "";
                                                            _SFC_Common.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                            #endregion
                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                            _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                        }
                                                        #endregion
                                                    }
                                                    #endregion
                                                }
                                                else
                                                {
                                                    #region 沒計畫量, 扣實體倉
                                                    if (tmp_int > 0)
                                                    {
                                                        DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                        if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                        {
                                                            foreach (DataRow d2 in tmp_dt2.Rows)
                                                            {
                                                                if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                {
                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                    _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                    tmp_int = 0;
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                    if (tmp_01 != 0)
                                                                    {
                                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                        _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                                        tmp_int -= tmp_01;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    if (tmp_int > 0)
                                                    {
                                                        #region 實體倉不購扣, 扣空倉
                                                        #region 查找適合庫儲別
                                                        string out_StoreNO = "";
                                                        string out_StoreSpacesNO = "";
                                                        _SFC_Common.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                        #endregion
                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                        _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno}Id:{d["SimulationId"].ToString()}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, br.UserNO);
                                                        #endregion
                                                    }
                                                    #endregion
                                                }
                                            }
                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Next_APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}',DOCNumberNO='{docNumberNO}',Next_StationQTY+={in_APS_PartNOTimeNote_QTY} where SimulationId='{d["SimulationId"].ToString()}'");
                                            if (d["DOCNumberNO"].ToString().Trim() == "" && docNumberNO != "")
                                            { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET DOCNumberNO='{docNumberNO}' where SimulationId='{d["SimulationId"].ToString()}'"); }
                                            #endregion
                                        }
                                        #endregion

                                        #region 修正上階子計畫完成
                                        if (isNeedQTY_OK && !bool.Parse(d["IsOK"].ToString()))
                                        {
                                            if (!d.IsNull("Source_StationNO") && (d["Class"].ToString() == "4" || d["Class"].ToString() == "5"))
                                            {
                                                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsOK='1' where SimulationId='{d["SimulationId"].ToString()}'"))
                                                {
                                                    db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set ArrivalDate='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where IsOK='0' and SimulationId='{d["SimulationId"].ToString()}'");
                                                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=1,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{d["NeedId"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                }
                                            }
                                            else
                                            {
                                                db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC3stockII] set ArrivalDate='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where IsOK='0' and SimulationId='{d["SimulationId"].ToString()}'");
                                            }
                                        }
                                        #endregion
                                    }
                                }
                                #endregion

                                #region 判斷下一站是否為委外加工
                                if (_Fun.Config.OutPackStationName == dr_APS_Simulation["Apply_StationNO"].ToString())
                                {
                                    DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr_APS_Simulation["SimulationId"].ToString()}'");
                                    int docQTY = int.Parse(tmp["Detail_QTY"].ToString());
                                    int changQTY = docQTY;
                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_APS_Simulation["NeedId"].ToString()}' and Source_StationNO='{dr_APS_Simulation["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN={(int.Parse(dr_APS_Simulation["Source_StationNO_IndexSN"].ToString()) + 1).ToString()}");
                                    if (tmp != null)//tmp為下一站的 APS_Simulation
                                    {
                                        string docNumberNO = tmp["DOCNumberNO"].ToString();
                                        #region 扣已有單據數量
                                        DataRow doc_DOC4II = db.DB_GetFirstDataByDataRow($"select sum(QTY) as okQTY from SoftNetMainDB.[dbo].[DOC4ProductionII] where SimulationId='{tmp["SimulationId"].ToString()}'");
                                        if (doc_DOC4II != null && !doc_DOC4II.IsNull("okQTY") && doc_DOC4II["okQTY"].ToString() != "") { docQTY -= int.Parse(doc_DOC4II["okQTY"].ToString()); }
                                        #endregion
                                        if (docQTY > 0)
                                        {
                                            string tmp_down_SID = "NULL";
                                            string tmp_down_Source_StationNO = "NULL";
                                            //用ArrivalDate網前計算StartTime
                                            string tmp_down_StartTime = Convert.ToDateTime(dr_APS_Simulation["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss");
                                            string tmp_down_ArrivalDate = Convert.ToDateTime(tmp["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss");
                                            DataRow tmp_up = dr_APS_Simulation;
                                            DataRow tmp_down = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{tmp["NeedId"].ToString()}' and Source_StationNO='{tmp["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN={(int.Parse(tmp["Source_StationNO_IndexSN"].ToString()) + 1).ToString()} and PartSN<{tmp["PartSN"].ToString()} order by PartSN desc");
                                            if (tmp_down != null)
                                            {
                                                tmp_down_SID = $"'{tmp_down["SimulationId"].ToString()}'";
                                                tmp_down_Source_StationNO = $"'{tmp_down["Source_StationNO"].ToString()}'";
                                                tmp_down_ArrivalDate = Convert.ToDateTime(tmp_down["StartDate"]).ToString("yyyy-MM-dd HH:mm:ss");
                                            }
                                            #region 查找適合廠商
                                            float price = 0;
                                            string mFNO = _SFC_Common.SelectDOC4ProductionMFNO(db, tmp["PartNO"].ToString(), tmp["SimulationId"].ToString(), in_NO, ref price);
                                            #endregion
                                            #region 查找適合入庫儲別
                                            string in_StoreNO = "";
                                            string in_StoreSpacesNO = "";
                                            _SFC_Common.SelectINStore(db, tmp["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "PA02", true);
                                            #endregion

                                            if (_SFC_Common.Create_DOC4stock(db, tmp, mFNO, price, in_StoreNO, in_StoreSpacesNO, "PA02", docQTY, "", "", "工站報工,開下一站委外加工", tmp_down_StartTime, tmp_down_ArrivalDate, br.UserNO.Trim(), ref docNumberNO))
                                            {
                                                DataRow tmp_9 = db.DB_GetFirstDataByDataRow($"SELECT PartNO,IS_WorkingPaper,IS_Store_Test FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{tmp["PartNO"].ToString()}'");
                                                if (bool.Parse(tmp_9["IS_WorkingPaper"].ToString()))
                                                {
                                                    tmp_9 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where PartNO='{tmp["PartNO"].ToString()}' and WorkType='2' and SimulationId='{tmp["SimulationId"].ToString()}' and MFNO='{mFNO}' and DOCNumberNO='{docNumberNO}'");
                                                    if (tmp_9 == null)
                                                    {
                                                        sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkingPaper] (ServerId,[Id],[WorkType],[PartNO],[Class],[IsOK],[NeedId],[SimulationId],[UP_SimulationId],[Down_SimulationId],[NeedQTY],[Price],[Unit],[MFNO],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],[APS_StationNO],[APS_StationNO_SID],[StartTime],[ArrivalDate],[EndTime],[UpdateTime],DOCNumberNO)
                                                                VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('P')}','2','{tmp["PartNO"].ToString()}','{tmp["Class"].ToString()}','0','{tmp["NeedId"].ToString()}','{tmp["SimulationId"].ToString()}','{tmp_up["SimulationId"].ToString()}',{tmp_down_SID},{docQTY},{price},'PCS','{mFNO}','{in_StoreNO}','{in_StoreSpacesNO}','','',
                                                                {tmp_down_Source_StationNO},{tmp_down_SID},'{tmp_down_StartTime}','{tmp_down_ArrivalDate}',NULL,'{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{docNumberNO}')";
                                                        db.DB_SetData(sql);
                                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsWPaper='1',DOCNumberNO='{docNumberNO}' where SimulationId='{tmp["SimulationId"].ToString()}'");
                                                    }
                                                }
                                            }
                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{docNumberNO}' where SimulationId='{tmp["SimulationId"].ToString()}'");
                                        }
                                    }
                                    if (isNeedQTY_OK)
                                    {
                                        DataTable dt_APS_WorkingPaper = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkingPaper] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{dr_APS_Simulation["SimulationId"].ToString()}' and WorkType='2' and IsOK='0'");
                                        if (dt_APS_WorkingPaper != null && dt_APS_WorkingPaper.Rows.Count > 0)
                                        {
                                            foreach (DataRow d2 in dt_APS_WorkingPaper.Rows)
                                            {
                                                if (d2.IsNull("StartTime") || (!d2.IsNull("StartTime") && Convert.ToDateTime(d2["StartTime"]) < DateTime.Now))
                                                {
                                                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkingPaper] set StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where Id='{d2["Id"].ToString()}'");
                                                }
                                            }
                                        }
                                    }

                                }
                                #endregion

                                #region 消除 APS_WorkTimeNote 工站負荷
                                DataRow tmp_del = null;
                                string stationNO_Merge = "";
                                int delMath_UseTime = 0; int tmp_ct = 0; int tmp_wt = 0; int tmp_st = 0; int tmp_1 = 0; int tmp_2 = 0; int tmp_3 = 0; int tmp_4 = 0;
                                tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SimulationId}'");
                                if (tmp_del != null)
                                {
                                    tmp_ct = int.Parse(tmp_del["Math_EfficientCT"].ToString());
                                    tmp_wt = int.Parse(tmp_del["Math_EfficientWT"].ToString());
                                    tmp_st = int.Parse(tmp_del["Math_StandardCT"].ToString());
                                    if ((tmp_ct + tmp_ct) != 0)
                                    { delMath_UseTime += (tmp_ct + tmp_ct) * (outQTY + failQTY); }
                                    else if (tmp_st != 0)
                                    { delMath_UseTime += tmp_st * (outQTY + failQTY); }
                                    else
                                    { delMath_UseTime += (int)ct * (outQTY + failQTY); }
                                    if (!tmp_del.IsNull("StationNO_Merge") && tmp_del["StationNO_Merge"].ToString().Trim() != "")
                                    {
                                        stationNO_Merge = tmp_del["StationNO_Merge"].ToString().Trim().Substring(0, tmp_del["StationNO_Merge"].ToString().Trim().Length - 1);
                                        stationNO_Merge = $" in ('{stationNO_Merge.Replace(",", "','")}')";
                                    }
                                }
                                #region 先消自己
                                dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SimulationId}' and StationNO='{data[0]}' order by CalendarDate");
                                if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                                {
                                    foreach (DataRow d in dt_APS_Simulation.Rows)
                                    {
                                        tmp_1 = int.Parse(d["Time1_C"].ToString()); tmp_2 = int.Parse(d["Time2_C"].ToString()); tmp_3 = int.Parse(d["Time3_C"].ToString()); tmp_4 = int.Parse(d["Time4_C"].ToString());
                                        if (delMath_UseTime > 0)
                                        {
                                            if (tmp_1 > 0)
                                            {
                                                if (delMath_UseTime > tmp_1) { delMath_UseTime -= tmp_1; tmp_1 = 0; } else { tmp_1 -= delMath_UseTime; delMath_UseTime = 0; }
                                            }
                                            if (delMath_UseTime > 0 && tmp_2 > 0)
                                            {
                                                if (delMath_UseTime > tmp_2) { delMath_UseTime -= tmp_2; tmp_2 = 0; } else { tmp_2 -= delMath_UseTime; delMath_UseTime = 0; }
                                            }
                                            if (delMath_UseTime > 0 && tmp_3 > 0)
                                            {
                                                if (delMath_UseTime > tmp_3) { delMath_UseTime -= tmp_3; tmp_3 = 0; } else { tmp_3 -= delMath_UseTime; delMath_UseTime = 0; }
                                            }
                                            if (delMath_UseTime > 0 && tmp_4 > 0)
                                            {
                                                if (delMath_UseTime > tmp_4) { delMath_UseTime -= tmp_4; tmp_4 = 0; } else { tmp_4 -= delMath_UseTime; delMath_UseTime = 0; }
                                            }
                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET Time1_C={tmp_1},Time2_C={tmp_2},Time3_C={tmp_3},Time4_C={tmp_4} where Id='{d["Id"].ToString()}'");
                                        }
                                        if (delMath_UseTime <= 0) { break; }
                                    }
                                }
                                #endregion

                                #region 不夠消, 消其他合併站
                                if (delMath_UseTime > 0 && stationNO_Merge != "")
                                {
                                    dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SimulationId}' and StationNO {stationNO_Merge} order by CalendarDate");
                                    if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                                    {
                                        foreach (DataRow d in dt_APS_Simulation.Rows)
                                        {
                                            tmp_1 = int.Parse(d["Time1_C"].ToString());
                                            tmp_2 = int.Parse(d["Time2_C"].ToString());
                                            tmp_3 = int.Parse(d["Time3_C"].ToString());
                                            tmp_4 = int.Parse(d["Time4_C"].ToString());
                                            if (delMath_UseTime > 0)
                                            {
                                                if (tmp_1 > 0)
                                                {
                                                    if (delMath_UseTime > tmp_1) { delMath_UseTime -= tmp_1; tmp_1 = 1; } else { tmp_1 -= delMath_UseTime; delMath_UseTime = 0; }
                                                }
                                                if (delMath_UseTime > 0 && tmp_2 > 0)
                                                {
                                                    if (delMath_UseTime > tmp_2) { delMath_UseTime -= tmp_2; tmp_2 = 1; } else { tmp_2 -= delMath_UseTime; delMath_UseTime = 0; }
                                                }
                                                if (delMath_UseTime > 0 && tmp_3 > 0)
                                                {
                                                    if (delMath_UseTime > tmp_3) { delMath_UseTime -= tmp_3; tmp_3 = 1; } else { tmp_3 -= delMath_UseTime; delMath_UseTime = 0; }
                                                }
                                                if (delMath_UseTime > 0 && tmp_4 > 0)
                                                {
                                                    if (delMath_UseTime > tmp_4) { delMath_UseTime -= tmp_4; tmp_4 = 1; } else { tmp_4 -= delMath_UseTime; delMath_UseTime = 0; }
                                                }
                                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET Time1_C={tmp_1},Time2_C={tmp_2},Time3_C={tmp_3},Time4_C={tmp_4} where Id='{d["Id"].ToString()}'");
                                            }
                                            if (delMath_UseTime <= 0) { break; }
                                        }
                                    }
                                }
                                #endregion
                                #endregion
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    meg = $"後臺錯誤: {ex.Message}";
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"後臺錯誤: {ex.Message} {ex.StackTrace}", true);
                }
            }



            return meg;
        }

        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(Ctrl, dt));
        }


        private STViewService EditService()
        {
            return new STViewService(Ctrl);
        }

        //讀取要修改的資料(Get Updated Json)
        [HttpPost]
        public async Task<ContentResult> GetUpdJson(string key)
        {
            return JsonToCnt(await EditService().GetUpdJsonAsync(key));
        }

        //新增(DB)
        public async Task<JsonResult> Create(string json)
        {
            return Json(await EditService().CreateAsync(_Str.ToJson(json)));
        }
        //修改(DB)
        public async Task<JsonResult> Update(string key, string json)
        {
            return Json(await EditService().UpdateAsync(key, _Str.ToJson(json)));
        }

        //刪除(DB)
        public async Task<JsonResult> Delete(string key)
        {
            return Json(await EditService().DeleteAsync(key));
        }


    }//class
}
