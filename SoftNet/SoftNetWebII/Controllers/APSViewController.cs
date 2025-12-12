//using AspNetCore;
using Base;
using Base.Enums;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion.Internal;
using Microsoft.VisualBasic;
using SoftNetWebII.Models;
using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Net.WebSockets;
using System.Threading.Tasks;
using System.Xml;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;

namespace SoftNetWebII.Controllers
{
    public class APSViewController : ApiCtrl
    {
        //private SocketClientService _SNsocket = null;
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;

        public APSViewController(SNWebSocketService websocket, SFC_Common sfc_Common)
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

        private void RecursiveBOM2(DBADO db, string needId, int sn, int sn2, DataTable dr_M,int indexSN, string pp_Name, string apply_PartNO, string mBOMId, ref string re)
        {
            if (dr_M == null || dr_M.Rows.Count <= 0) { return; }
            DataRow dr_tmp = null;
            DataTable dr_MII = null;

            if (sn != sn2) { re = $"{re}<ul>"; sn = sn2; }
            foreach (DataRow dr1 in dr_M.Rows)
            {
                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Apply_PartNO='{apply_PartNO}' and Apply_PP_Name='{pp_Name}' and IndexSN={(indexSN-1)}");
                if (dr_tmp != null)
                {
                    re = $"{re}   <li><button id='{dr_tmp["Id"].ToString()}' onclick='_me.onCreateBOM(\"1,{mBOMId},{pp_Name},{dr_tmp["Id"].ToString()}\")'>{dr1["PartNO"].ToString()}</button>";
                    dr_MII = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[BOMII] where BOMId='{dr_tmp["Id"].ToString()}' order by sn");
                    if (dr_MII != null && dr_MII.Rows.Count > 0) { RecursiveBOM2(db, needId, sn2, (sn2 + 1), dr_MII, (indexSN - 1), pp_Name, apply_PartNO, mBOMId, ref re); }
                }
                else
                {
                    re = $"{re}   <li><button id='{dr1["Id"].ToString()}' onclick='_me.onCreateBOM(\"2,{mBOMId},{pp_Name},{dr1["Id"].ToString()}\")'>{dr1["PartNO"].ToString()}</button>";
                }
                re = $"{re}</li>";
            }
            re = $"{re}</ul>";
        }

     

        public IActionResult DisplayLaout(string id)
        {
            string re = "";
            DBADO db = new DBADO("1", _Fun.Config.Db);
            DataRow dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{id}'");
            if (dr != null && dr["State"].ToString() == "3")
            {
                re = "<p>模擬的資料已逾時, 需重新模擬</p>";
            }
            else
            {

                //BOM結構 取得第一層結構
                DataTable dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{id}' and PartSN=0 order by PartSN");
                if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
                {
                    DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select b.* from SoftNetSYSDB.[dbo].[APS_Simulation] as a,SoftNetMainDB.[dbo].[Material] as b where a.NeedId='{id}' and a.PartSN=0 and a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}'");
                    re = $@"<figure>   <figcaption>需求代碼:{id}  料號:{dr_tmp["PartNO"].ToString()}  品名:{dr_tmp["PartName"].ToString()}  規格:{dr_tmp["Specification"].ToString()}</figcaption><hr><figcaption>用料結構表</figcaption><ul class='tree'>   <li><code>{dt_Simulation.Rows[0]["Master_PartNO"].ToString()}</code>";
                    dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{id}' and PartSN=1 order by PartSN");

                    RecursiveSimulation2(db, id, 0, 1, dt_Simulation, ref re);

                    re = $"{re}</ul></figure>";
                }
                else
                {
                    re = $@"<figure>   <figcaption>需求代碼:{id}  目前沒模擬資訊可以顯示</figcaption><figure>";
                }
                ViewBag.HtmlOutputBOM = re;

                re = "";
                //生產製程 取得所有階層
                string state = dr["State"].ToString();
                dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{id}' and PartSN>=0 order by PartSN desc");
                if (dr != null) { maxPartSN = int.Parse(dr["PartSN"].ToString()); }
                dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{id}' and PartSN>=0 order by PartSN");
                if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
                {
                    string show = "label";
                    string custom_DisplayName = dt_Simulation.Rows[0]["Source_StationNO_Custom_DisplayName"].ToString();//
                    if (custom_DisplayName != "") { custom_DisplayName = $"<br />{custom_DisplayName}"; }
                    if (db.DB_GetQueryCount($"SELECT SimulationId FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{dt_Simulation.Rows[0]["SimulationId"].ToString()}'") > 0) { show = "labelRed"; }
                    re = $"<hr><figcaption>生產流程圖</figcaption>\n<div id='wrapper'><button id='{dt_Simulation.Rows[0]["SimulationId"].ToString()}' class='{show}' onclick=\"_me.onShow('{dt_Simulation.Rows[0]["SimulationId"].ToString()}')\">{dt_Simulation.Rows[0]["PartNO"].ToString()}<br />[{dt_Simulation.Rows[0]["Source_StationNO"].ToString()}][{dt_Simulation.Rows[0]["Source_StationNO_IndexSN"].ToString()}]{custom_DisplayName}</button>\n";
                    if (dt_Simulation.Rows.Count > 1)
                    {
                        dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{id}' and PartSN=1 order by PartSN");
                        DataRow[] dr_WorkTimeNote = dt_Simulation.Select($"PartSN=2 and Master_PartNO='{dt_Simulation.Rows[0]["PartNO"].ToString()}' and Apply_StationNO='{dt_Simulation.Rows[0]["Source_StationNO"].ToString()}'");
                        RecursiveTree(db, id, 1, (dr_WorkTimeNote != null ? true : false), dt_Simulation, state, ref re);
                    }
                    re = $"{re}</div>\n";
                }
            }
            db.Dispose();

            ViewBag.HtmlOutputProcess = re;

            return View();
        }
        private void RecursiveSimulation2(DBADO db, string needId, int sn, int sn2, DataTable dr_M, ref string re)
        {
            if (dr_M == null || dr_M.Rows.Count <= 0) { return; }
            DataRow dr = null;
            DataTable dr_MII = null;
            if (!dr_M.Rows[0].IsNull("Source_StationNO") && dr_M.Rows.Count == 1)
            {
                #region 僅加工
                dr_MII = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and Master_PartNO='{dr_M.Rows[0]["PartNO"].ToString()}' and Apply_StationNO='{dr_M.Rows[0]["Source_StationNO"].ToString()}' and PartSN={(sn2 + 1)}");
                if (dr_MII == null || dr_MII.Rows.Count > 0)
                {
                    RecursiveSimulation2(db, needId, sn2, int.Parse(dr_MII.Rows[0]["PartSN"].ToString()), dr_MII, ref re);
                }
                #endregion
            }
            else
            {
                if (sn != sn2) { re = $"{re}<ul>"; sn = sn2; }

                foreach (DataRow dr1 in dr_M.Rows)
                {
                    if (!dr1.IsNull("Source_StationNO"))
                    {
                        re = $"{re}   <li><code>{dr1["PartNO"].ToString()}</code>";
                        dr_MII = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and Master_PartNO='{dr1["PartNO"].ToString()}' and PartSN={(sn2 + 1)}");
                        RecursiveSimulation2(db, needId, sn2, int.Parse(dr_MII.Rows[0]["PartSN"].ToString()), dr_MII, ref re);
                        re = $"{re}</li>";
                    }
                    else
                    { re = $"{re}<li><code>{dr1["PartNO"].ToString()}</code></li>"; }
                }
                re = $"{re}</ul>";
            }
        }
        private int maxPartSN = 0;
        private void RecursiveTree(DBADO db, string needId, int sn, bool hasNext, DataTable dr_M,string state, ref string re,bool isHasButton=true)
        {
            if (hasNext) { re = $"{re}<div class='branch lv{sn}'>\n"; }
            string show = "label";
            string entryType = "entry";
            DataRow[] dr_WorkTimeNote = null;
            DataRow dr_PartNOTimeNote = null;
            string s_date = "";
            string is_Disabled = "";
            if (!isHasButton) { is_Disabled = "disabled"; }
            foreach (DataRow dr1 in dr_M.Rows)
            {
                if (db.DB_GetQueryCount($"SELECT SimulationId FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{dr1["SimulationId"].ToString()}' and ActionType=''") > 0) { show = "labelRed"; }
                else { show = "label"; }
                if (!dr1.IsNull("Source_StationNO") && dr1["Source_StationNO"].ToString().Trim() != "")
                {
                    //工站
                    DataTable dr_MII = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartSN={(sn + 1)} and Master_PartNO='{dr1["PartNO"].ToString()}' and Apply_StationNO='{dr1["Source_StationNO"].ToString()}' order by PartSN,PartSN_Sub desc");
                    if (dr_M.Rows.Count == 1 && dr_MII != null && dr_MII.Rows.Count > 0) { entryType = "entry sole"; }
                    else if (dr_MII == null || dr_MII.Rows.Count <= 0) { continue; }
                    s_date = dr1.IsNull("StartDate") ? "" : Convert.ToDateTime(dr1["StartDate"]).ToString("yyyy-MM-dd HH:mm");
                    string qty = "";
                    if (!dr1.IsNull("Source_StationNO") && state == "6")
                    {
                        dr_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr1["SimulationId"].ToString()}'");
                        if (dr_PartNOTimeNote == null) { qty = $"<br /><label style='background-color:tomato;'>排程資料異常</label>"; }
                        else { if (dr1["DOCNumberNO"].ToString() != "") { qty = $"<br />{dr1["DOCNumberNO"].ToString()}&nbsp;&nbsp;已產:{dr_PartNOTimeNote["Detail_QTY"].ToString()}"; } }
                    }
                    string custom_DisplayName = dr1["Source_StationNO_Custom_DisplayName"].ToString();//
                    if (custom_DisplayName != "") { custom_DisplayName = $"{custom_DisplayName}<br />"; }

                    string otherINFO = "";
                    if (!dr1.IsNull("StationNO_Merge") && dr1["StationNO_Merge"].ToString() != "") { otherINFO = $"合併站:{dr1["StationNO_Merge"].ToString()}<br />"; }
                    re = $"{re}<div class='{entryType}'><button id='{dr1["SimulationId"].ToString()}' class='{show}' {is_Disabled} onclick=\"_me.onShow('{dr1["SimulationId"].ToString()}')\">{dr1["PartNO"].ToString()}[{dr1["Source_StationNO"].ToString()}][{dr1["Source_StationNO_IndexSN"].ToString()}]<br />{otherINFO}{custom_DisplayName}開始日:{s_date}<br />完成日:{Convert.ToDateTime(dr1["SimulationDate"]).ToString("yyyy-MM-dd HH:mm")}{qty}</button>\n";

                    if (dr_MII != null && dr_MII.Rows.Count > 0)
                    {
                        dr_WorkTimeNote = dr_M.Select($"PartSN={(sn + 1).ToString()} and Master_PartNO='{dr1["PartNO"].ToString()}' and Apply_StationNO='{dr1["Source_StationNO"].ToString()}'");
                        RecursiveTree(db, needId, (sn + 1), (dr_WorkTimeNote != null ? true : false), dr_MII, state, ref re, isHasButton);
                    }
                    else
                    {
                        re = $"{re}<div class='entry'><button id='{dr1["SimulationId"].ToString()}' class='{show}' {is_Disabled} onclick=\"_me.onShow('{dr1["SimulationId"].ToString()}')\">{dr1["PartNO"].ToString()}</button>";
                    }
                    re = $"{re}</div>\n";
                }
                else
                {
                    re = $"{re}<div class='entry'><button id='{dr1["SimulationId"].ToString()}' class='{show}' {is_Disabled} onclick=\"_me.onShow('{dr1["SimulationId"].ToString()}')\">{dr1["PartNO"].ToString()}</button></div>\n";
                }
            }
            if (hasNext) { re = $"{re}</div>\n"; }
        }
        [HttpPost]
        public ActionResult SetchangNeedQTY(string keys) //0=chang_NeedID 1=chang_NeedQTY 2=chang_SafeQTY 3=是否重新SimulationId
        {
            string re = "";
            string[] data = keys.Split(',');
            
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow dr_tmp = null;
                int c_NeedQTY = int.Parse(data[1]);
                int c_SafeQTY = int.Parse(data[2]);
                int tot_qty = c_NeedQTY + c_SafeQTY;
                //###??? 暫時把數量先寫入NeedQTY
                if (data[3] == "0")
                {
                    #region 只改負荷, 數量
                    DataTable dt_APS_NeedData = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{data[0]}'");
                    DataTable dt_APS_WorkTimeNote = null;
                    DataTable dt_tmp = null;
                    if (dt_APS_NeedData != null && dt_APS_NeedData.Rows.Count > 0)
                    {
                        int needQTY = 0;
                        int safeQTY = 0;
                        int qty = 0;//數量差

                        int next_StoreQTY = 0;
                        foreach (DataRow dr in dt_APS_NeedData.Rows)
                        {
                            if (tot_qty > int.Parse(dr["Detail_QTY"].ToString()))
                            {
                                needQTY = int.Parse(dr["NeedQTY"].ToString());
                                //safeQTY = int.Parse(dr["SafeQTY"].ToString());
                                qty = tot_qty - (needQTY + safeQTY);
                                if (dr["NoStation"].ToString() == "0")
                                {
                                    #region 修正已退庫IsOK=0的數量
                                    if (qty < 0)//變少
                                    {
                                        if (int.Parse(dr["Next_StoreQTY"].ToString()) >= Math.Abs(qty))
                                        {
                                            next_StoreQTY = Math.Abs(qty);
                                            dt_tmp = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where DOCNumberNO='{dr["Store_DOCNumberNO"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}' and IsOK='0'");
                                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                            {
                                                foreach (DataRow dr_DOC3 in dt_tmp.Rows)
                                                {
                                                    if (int.Parse(dr_DOC3["QTY"].ToString()) >= next_StoreQTY)
                                                    {
                                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set QTY-={next_StoreQTY} where Id='{dr_DOC3["Id"].ToString()}' and DOCNumberNO='{dr_DOC3["DOCNumberNO"].ToString()}'");
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        next_StoreQTY -= int.Parse(dr_DOC3["QTY"].ToString());
                                                        db.DB_SetData($"delete SoftNetMainDB.[dbo].[DOC3stockII] set QTY=0 where Id='{dr_DOC3["Id"].ToString()}' and DOCNumberNO='{dr_DOC3["DOCNumberNO"].ToString()}'");
                                                    }
                                                }
                                                #region 退庫不夠扣
                                                if (next_StoreQTY > 0)
                                                {
                                                    if (c_SafeQTY >= next_StoreQTY) { c_SafeQTY -= next_StoreQTY; }
                                                    else
                                                    {
                                                        next_StoreQTY -= c_SafeQTY;
                                                        c_SafeQTY = 0;
                                                        c_NeedQTY -= next_StoreQTY;
                                                        if (c_NeedQTY < 0) { c_NeedQTY = 0; }
                                                    }
                                                }
                                                #endregion
                                            }
                                        }
                                    }
                                    #endregion

                                    #region 修正TotalStockII Keep量
                                    if (qty < 0)//變少
                                    {
                                        next_StoreQTY = Math.Abs(qty);
                                        dt_tmp = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{data[0]}' and SimulationId='{dr["SimulationId"].ToString()}' and KeepQTY>0 order by ArrivalDate desc");
                                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                        {
                                            foreach (DataRow dr_StockII in dt_tmp.Rows)
                                            {
                                                if (int.Parse(dr_StockII["KeepQTY"].ToString()) >= next_StoreQTY)
                                                {
                                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY-={next_StoreQTY} where NeedId='{data[0]}' and Id='{dr_StockII["Id"].ToString()}' and SimulationId='{dr_StockII["SimulationId"].ToString()}'");
                                                    break;
                                                }
                                                else
                                                {
                                                    next_StoreQTY -= int.Parse(dr_StockII["QTY"].ToString());
                                                    db.DB_SetData($"delete SoftNetMainDB.[dbo].[DOC3stockII] set KeepQTY=0 where NeedId='{data[0]}' and Id='{dr_StockII["Id"].ToString()}' and SimulationId='{dr_StockII["SimulationId"].ToString()}'");
                                                    if (int.Parse(dr_StockII["OverQTY"].ToString()) > 0)
                                                    {
                                                        if (int.Parse(dr_StockII["OverQTY"].ToString()) >= next_StoreQTY)
                                                        {
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set OverQTY-={next_StoreQTY} where NeedId='{data[0]}' and Id='{dr_StockII["Id"].ToString()}' and SimulationId='{dr_StockII["SimulationId"].ToString()}'");
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            next_StoreQTY -= int.Parse(dr_StockII["OverQTY"].ToString());
                                                            db.DB_SetData($"delete SoftNetMainDB.[dbo].[DOC3stockII] set OverQTY=0 where NeedId='{data[0]}' and Id='{dr_StockII["Id"].ToString()}' and SimulationId='{dr_StockII["SimulationId"].ToString()}'");
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        next_StoreQTY = Math.Abs(qty);
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{data[0]}' and SimulationId='{dr["SimulationId"].ToString()}' order by ArrivalDate desc");
                                        if (dr_tmp != null)
                                        {
                                            db.DB_SetData($"delete from SoftNetMainDB.[dbo].[TotalStockII] set OverQTY+={next_StoreQTY} where NeedId='{data[0]}' and Id='{dr_tmp["Id"].ToString()}' and SimulationId='{dr_tmp["SimulationId"].ToString()}'");
                                        }
                                    }
                                    #endregion

                                    #region 先計算負荷基準值
                                    int delMath_UseTime = 0; int tmp_ct = 0; int tmp_wt = 0; int tmp_st = 0; int tmp_1 = 0; int tmp_2 = 0; int tmp_3 = 0; int tmp_4 = 0;int time_TOT = 0;
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{data[0]}' and SimulationId='{dr["SimulationId"].ToString()}'");
                                    if (dr_tmp != null)
                                    {
                                        tmp_ct = int.Parse(dr_tmp["Math_EfficientCT"].ToString());
                                        tmp_wt = int.Parse(dr_tmp["Math_EfficientWT"].ToString());
                                        tmp_st = int.Parse(dr_tmp["Math_StandardCT"].ToString());
                                        if ((tmp_ct + tmp_ct) != 0)
                                        { delMath_UseTime += (tmp_ct + tmp_ct) * qty; }
                                        else if (tmp_st != 0)
                                        { delMath_UseTime += tmp_st * qty; }
                                        else
                                        { delMath_UseTime = 60 * qty; }
                                    }
                                    #endregion

                                    #region 消除或增加 APS_WorkTimeNote 工站負荷
                                    if (delMath_UseTime < 0)
                                    {
                                        delMath_UseTime = Math.Abs(delMath_UseTime);
                                        dt_APS_WorkTimeNote = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{data[0]}' and SimulationId='{dr["SimulationId"].ToString()}' order by CalendarDate desc");
                                        if (dt_APS_WorkTimeNote != null && dt_APS_WorkTimeNote.Rows.Count > 0)
                                        {
                                            foreach (DataRow dr_Work in dt_APS_WorkTimeNote.Rows)
                                            {
                                                tmp_1 = int.Parse(dr_Work["Time1_C"].ToString()); tmp_2 = int.Parse(dr_Work["Time2_C"].ToString()); tmp_3 = int.Parse(dr_Work["Time3_C"].ToString()); tmp_4 = int.Parse(dr_Work["Time4_C"].ToString()); time_TOT= int.Parse(dr_Work["Time_TOT"].ToString());
                                                if (delMath_UseTime > 0)
                                                {
                                                    if (tmp_1 > 0)
                                                    {
                                                        if (delMath_UseTime > tmp_1) { delMath_UseTime -= tmp_1; tmp_1 = 1; } else { tmp_1 -= delMath_UseTime; delMath_UseTime = 0;  }
                                                    }
                                                    if (delMath_UseTime > 0 && tmp_2 > 0)
                                                    {
                                                        if (delMath_UseTime > tmp_2) { delMath_UseTime -= tmp_2; tmp_2 = 1; } else { tmp_2 -= delMath_UseTime; delMath_UseTime = 0;  }
                                                    }
                                                    if (delMath_UseTime > 0 && tmp_3 > 0)
                                                    {
                                                        if (delMath_UseTime > tmp_3) { delMath_UseTime -= tmp_3; tmp_3 = 1; } else { tmp_3 -= delMath_UseTime; delMath_UseTime = 0;  }
                                                    }
                                                    if (delMath_UseTime > 0 && tmp_4 > 0)
                                                    {
                                                        if (delMath_UseTime > tmp_4) { delMath_UseTime -= tmp_4; tmp_4 = 1; } else { tmp_4 -= delMath_UseTime; delMath_UseTime = 0;  }
                                                    }
                                                    time_TOT -= delMath_UseTime;
                                                    if (time_TOT < 0) { time_TOT = 0; }
                                                    db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET Time1_C={tmp_1.ToString()},Time2_C={tmp_2.ToString()},Time3_C={tmp_3.ToString()},Time4_C={tmp_4.ToString()},Time_TOT={time_TOT.ToString()} where Id='{dr_Work["Id"].ToString()}'");
                                                }
                                                if (delMath_UseTime <= 0) { break; }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{data[0]}' and SimulationId='{dr["SimulationId"].ToString()}' and StationNO='{dr["APS_StationNO"].ToString()}' order by CalendarDate desc");
                                        if (dr_tmp != null)
                                        {
                                            tmp_1 = int.Parse(dr_tmp["Time1_C"].ToString()); tmp_2 = int.Parse(dr_tmp["Time2_C"].ToString()); tmp_3 = int.Parse(dr_tmp["Time3_C"].ToString()); tmp_4 = int.Parse(dr_tmp["Time4_C"].ToString()); time_TOT = int.Parse(dr_tmp["Time_TOT"].ToString());
                                            if (delMath_UseTime > 0)
                                            {
                                                if (tmp_4 > 0) { tmp_4 += delMath_UseTime; }
                                                else if (tmp_3 > 0) { tmp_3 += delMath_UseTime;  }
                                                else if (tmp_2 > 0) { tmp_2 += delMath_UseTime;  }
                                                else if (tmp_1 > 0) { tmp_1 += delMath_UseTime;  }
                                                else { tmp_1 += delMath_UseTime; }
                                                time_TOT += delMath_UseTime;
                                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET Time1_C={tmp_1},Time2_C={tmp_2},Time3_C={tmp_3},Time4_C={tmp_4},Time_TOT={time_TOT.ToString()} where Id='{dr_tmp["Id"].ToString()}'");
                                            }
                                        }
                                    }
                                    #endregion

                                    //db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set NeedQTY={c_NeedQTY.ToString()},SafeQTY={c_SafeQTY.ToString()} where SimulationId='{dr["SimulationId"].ToString()}'");
                                    db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set NeedQTY={c_NeedQTY.ToString()} where SimulationId='{dr["SimulationId"].ToString()}'");
                                }
                            }
                        }
                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set NeedQTY={data[1]},SafeQTY={data[2]} where NeedId='{data[0]}' and PartSN=0");
                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] set NeedQTY={data[1]} where Id='{data[0]}'");
                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[PP_WorkOrder] set Quantity={c_NeedQTY.ToString()} where NeedId='{data[0]}'");
                        re = $"計畫數量已更新為: {c_NeedQTY.ToString()} 數量";
                    }
                    #endregion
                }
                else
                {
                    //###??? 重新計算排程時間
                    re = "重新計算排程時間, 程式未完成, 目前無作用";
                }
                return Content(re);
            }
        }


        [HttpPost]
        public ActionResult OnDeleteSimulation(string keys) //0=ipport,1.NeedId
        {
            string[] data = keys.Split(',');
            string meg = "";
            bool run = false;
            string sql = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string pp_html_DATA = "";
                #region 可用製程資訊
                DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{data[1]}' and DOCNumberNO='' and (Class='4' or Class='5') order by CalendarDate");
                DataRow dr_tmp = null;
                if (dt != null && dt.Rows.Count > 0)
                {
                    pp_html_DATA = $@"<input type='text' style='width: 400px;' id='sub_SelectPP_DATA' list='optionsPP_Name_Data'>
                        <datalist id='optionsPP_Name_Data' style='width: 100%;'>";
                    pp_html_DATA = $"{pp_html_DATA}<option style='width: 100%;' value=''></option>";
                    foreach (DataRow d in dt.Rows)
                    {
                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}'");
                        pp_html_DATA = $"{pp_html_DATA}<option style='width: 100%;' value='{d["SimulationId"].ToString()}'>{dr_tmp["PartNO"].ToString()} 工站:{dr_tmp["Source_StationNO"].ToString()}[{dr_tmp["Source_StationNO_IndexSN"].ToString()}][{dr_tmp["Source_StationNO_Custom_DisplayName"].ToString()}]</option>";
                    }
                    pp_html_DATA = $"{pp_html_DATA}</datalist>";
                }
                #endregion

                meg = $"選擇刪除行為:";
                meg = $"{meg}<input type='hidden' id='delNeedId' value='{data[1]}' />";
                meg = $"{meg}<p><input type='radio' name='delSIDradio' value='A'>刪除計畫,並刪除所有交易資料 (PS:帳會有不合理之風險)</input></p>";
                meg = $"{meg}<p><input type='radio' name='delSIDradio' value='B'>刪除計畫,確認無交易資料,才能刪除</input></p>";
                meg = $"{meg}<p><input type='radio' name='delSIDradio' value='C'>生產到下列指定的工站,並自動入庫</input></p>";
                meg = $"{meg}<label>選擇工站:</label><p>{pp_html_DATA}</p>";
                meg = $"{meg}<hr><p><button type='submit' onclick='this.style.display=\"none\";_me.ondeleteConfirm()'>確認更新</button></p>";

            }
            return Content(meg);
        }
        [HttpPost]
        public ActionResult OnDeleteSimulationII(APSViewBOM key)
        {
            string meg = "";
            bool run = false;
            string sql = "";
            string needId = "";
            if (key == null || key.MBOMId == "")
            {
                meg= $"作業失敗, 畫面已逾時, 請關閉網頁瀏覽器 並 重新操作."; return Content(meg);
            }
            else { needId = key.MBOMId; }
            if (key != null && key.Fun_S_List != "")
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataRow dr_tmp = null;
                    DataTable dt = null;
                    bool has_del_data = false;
                    switch (key.Fun_S_List)
                    {
                        case "A"://刪除計畫
                        case "B"://刪除計畫,但不交易才刪
                            #region 回朔已發生的單據倉儲交易
                            //###???檢驗單是否不扣倉
                            dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOC1BuyII] where SimulationId='{key.SimulationId}' and IsOK='1'");
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                if (key.Fun_S_List == "B") { has_del_data = true;break; }
                                foreach (DataRow dr in dt.Rows)
                                { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={dr["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}' and StoreNO='{dr["StoreNO"].ToString()}' and StoreSpacesNO='{dr["StoreSpacesNO"].ToString()}'"); }
                            }
                            dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOC2SalesII] where SimulationId='{key.SimulationId}' and IsOK='1'");
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                if (key.Fun_S_List == "B") { has_del_data = true; break; }
                                foreach (DataRow dr in dt.Rows)
                                { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={dr["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}' and StoreNO='{dr["StoreNO"].ToString()}' and StoreSpacesNO='{dr["StoreSpacesNO"].ToString()}'"); }
                            }
                            dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOC4ProductionII] where SimulationId='{key.SimulationId}' and IsOK='1'");
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                if (key.Fun_S_List == "B") { has_del_data = true; break; }
                                foreach (DataRow dr in dt.Rows)
                                { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={dr["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}' and StoreNO='{dr["StoreNO"].ToString()}' and StoreSpacesNO='{dr["StoreSpacesNO"].ToString()}'"); }
                            }
                            dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{key.SimulationId}' and IsOK='1'");
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                if (key.Fun_S_List == "B") { has_del_data = true; break; }
                                foreach (DataRow dr in dt.Rows)
                                {
                                    if (dr.IsNull("OUT_StoreNO") || dr["OUT_StoreNO"].ToString().Trim() == "")
                                    { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={dr["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}' and StoreNO='{dr["IN_StoreNO"].ToString()}' and StoreSpacesNO='{dr["IN_StoreSpacesNO"].ToString()}'"); }
                                    else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={dr["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}' and StoreNO='{dr["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr["OUT_StoreSpacesNO"].ToString()}'"); }
                                }
                            }
                            if (key.Fun_S_List == "B" && has_del_data) { meg = "無法刪除計畫, 因為倉儲已有異動交易."; return Content(meg); }
                            string sID = "";
                            List<string> all_WO = new List<string>();
                            dt = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}'");
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                foreach (DataRow dr in dt.Rows)
                                {
                                    if (sID == "") { sID = $"'{dr["SimulationId"].ToString()}'"; }
                                    else { sID = $"{sID},'{dr["SimulationId"].ToString()}'"; }
                                    if (dr["DOCNumberNO"].ToString() != "" && dr["DOCNumberNO"].ToString().Substring(0, 4) == "XX01")
                                    {
                                        if (!all_WO.Contains(dr["DOCNumberNO"].ToString())) { all_WO.Add(dr["DOCNumberNO"].ToString()); }
                                    }
                                }
                            }
                            //lock (_Fun.Lock_Simulation_Flag)
                            //{
                                if (sID != "")
                                {
                                    db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[DOC1BuyII] where SimulationId in ({sID})");
                                    db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[DOC2SalesII] where SimulationId in ({sID})");
                                    db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[DOC4ProductionII] where SimulationId in ({sID})");
                                    db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId in ({sID})");
                                    db.DB_SetData($"delete FROM SoftNetMainDB.[dbo].[TotalStock_Blank_LOG] where SimulationId in ({sID})");
                                    db.DB_SetData($"delete FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where SimulationId in ({sID})");
                                }
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_NeedData] where Id='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] where Trigger_NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where NeedId='{needId}'");
                                
                            //}
                            dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where SimulationId in ({sID})");
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                foreach (DataRow dr in dt.Rows)
                                {
                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Label_ProjectType='0',OrderNO='',OP_NO='',Master_PP_Name='',PP_Name='',IndexSN=0,Station_Custom_IndexSN='',StationNO_Custom_DisplayName='',State='4',PartNO='',StartTime=NULL,RemarkTimeS=NULL,RemarkTimeE=NULL,EndTime=NULL where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                                    #region 更新關閉的電子Tag
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr["Config_macID"].ToString()}'");
                                    if (dr_tmp != null)
                                    {
                                        string tmp_s = "";
                                        string tmp_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"\",\"Text2\":\"\",\"PartNO\":\"\",\"Text3\":\"\",\"PartName\":\"\",\"Text4\":\"\",\"QTY\":\"\",\"Text5\":\"\",\"EfficientCT\":\"\",\"Text6\":\"\",\"Rate\":\"\",\"Text7\":\"累計量:\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{dr["StationNO"].ToString()};0;;0\",\"outtime\":0";
                                        if (dr_tmp["Version"].ToString().Trim() != "" && dr_tmp["Version"].ToString().Trim().Substring(0, 2) == "42")
                                        {
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                                            tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr["StationNO"].ToString()}\",\"text17\":\"{dr_tmp["StationName"].ToString()}\"";
                                            tmp_s = $"\"mac\":\"{dr["Config_macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{tmp_ShowValue}";
                                        }
                                        else
                                        {
                                            tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"{dr["StationNO"].ToString()}\",\"text17\":\"\"";
                                            tmp_s = $"\"mac\":\"{dr["Config_macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{tmp_ShowValue},\"ledrgb\":\"0\",\"ledstate\":0";
                                        }
                                        if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                        {
                                            _Fun.Tag_Write(db,dr["Config_macID"].ToString(),"刪除計畫", tmp_s);
                                        }
                                        db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{tmp_ShowValue}',Ledrgb='0',Ledstate=0,StationNO='{dr["StationNO"].ToString()}',Type='1',OrderNO='',IndexSN='',StoreNO='',StoreSpacesNO='',QTY=0,IsUpdate='1' where ServerId='{_Fun.Config.ServerId}' and macID='{dr["Config_macID"].ToString()}'");
                                    }
                                    #endregion
                                }
                            }
                            dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[ManufactureII] where SimulationId in ({sID})");
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                List<string> run_HasWeb_Id_Change = new List<string>();
                                foreach (DataRow dr in dt.Rows)
                                {
                                    if (db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[ManufactureII] where Id='{dr["Id"].ToString()}'"))
                                    { run_HasWeb_Id_Change.Add(dr["StationNO"].ToString()); }
                                }
                                if (run_HasWeb_Id_Change.Count > 0)
                                {
                                    foreach (string sno in run_HasWeb_Id_Change)
                                    {
                                        #region 通知網頁更新
                                        try
                                        {
                                            lock (_WebSocket.lock__WebSocketList)
                                            {
                                                foreach (KeyValuePair<string, rmsConectUserData> r in _WebSocket._WebSocketList)
                                                {
                                                    if (r.Key != null && r.Value.socket != null)
                                                    {
                                                        _WebSocket.Send(r.Value.socket, $"HasWeb_Id_Change,STView2Work_PageReload,{sno}");
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            System.Threading.Tasks.Task task = _Log.ErrorAsync($"APSViewController.cs 工站:{sno} 發HasWeb_Id_Changeg失敗 {ex.Message} {ex.StackTrace}", true);
                                        }
                                        #endregion
                                    }
                                }
                                if (all_WO.Count>0)
                                {
                                    foreach (string s in all_WO)
                                    {
                                        db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[OperateLog] where OrderNO='{s}'");
                                        db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[SFC_StationDetail] where OrderNO='{s}'");
                                        db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where OrderNO='{s}'");
                                        db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] where OrderNO='{s}'");
                                    }
                                }
                            }
                            #endregion
                            break;
                        case "C"://刪除到指定工站
                            meg = "程式未寫";
                            
                            break;
                    }


                }
            }
            if (meg == "") 
            {
                meg = "程式正常完成, 但畫面須重新連結,才會顯示更新過後的資訊";
            }
            return Content(meg);
        }
        [HttpPost]
        public ActionResult OnAction(string keys) //0=ipport,1.SimulationId
        {
            string[] data = keys.Split(',');
            string meg = "";
            bool run = false;
            string sql = "";
            bool is_Button = false;
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                #region 異常處理
                DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{data[1]}' and ActionType='' order by LogDate,ErrorType");
                if (dt != null && dt.Rows.Count > 0)
                {
                    dataList.Clear();
                    string errINFO = "";
                    int i = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        switch (dr["ErrorType"].ToString())
                        {
                            case "01":
                            case "02": errINFO = $"{errINFO} {(++i).ToString()}.計畫保留量不足"; break;
                            case "03": errINFO = $"{errINFO} {(++i).ToString()}.工站未如期開工"; break;
                            case "04":
                            case "05": errINFO = $"{errINFO} {(++i).ToString()}.未開單據"; break;
                            case "06":
                            case "07":
                            case "08":
                            case "09":
                            case "10": errINFO = $"{errINFO} {(++i).ToString()}.單據應確認"; break;
                            case "11": errINFO = $"{errINFO} {(++i).ToString()}.工單未如期開工"; break;
                            case "12": errINFO = $"{errINFO} {(++i).ToString()}.工單未關帳"; break;
                            case "18": errINFO = $"{errINFO} {(++i).ToString()}.工站應停止"; break;
                            case "19": errINFO = $"{errINFO} {(++i).ToString()}.完成日未達標"; break;
                        }
                    }
                    meg = $"異常原因:{errINFO}<br />";
                    meg = $"{meg}干涉行為:<input id='ageID' name='ageDATA' list='ageAction' value='' /><datalist id='ageAction'><option value=''></option><option value='00,不作為'></option>";
                    foreach (DataRow dr in dt.Rows)
                    {
                        string tmp_s = dr["ErrorType"].ToString();
                        switch (tmp_s)
                        {
                            case "01"://TotalStockII keep量不足 (非開工單)
                                #region
                                AdddataList(tmp_s, "自動開立進貨單");
                                AdddataList(tmp_s, "自動調撥其他可用量");
                                AdddataList(tmp_s, "自動延後產出時間");
                                AdddataList(tmp_s, "自動採用替代料");
                                break;
                            #endregion
                            case "02"://TotalStockII keep量不足 (開工單)
                                #region
                                AdddataList(tmp_s, "自動調撥其他可用量");
                                AdddataList(tmp_s, "自動延後產出時間");
                                AdddataList(tmp_s, "自動追加生產數量");
                                AdddataList(tmp_s, "自動採用替代料");
                                AdddataList(tmp_s, "自動改委外生產");
                                break;
                            #endregion
                            case "03"://該開站未開(初次)
                                #region
                                AdddataList(tmp_s, "自動異動狀態為開站");
                                AdddataList(tmp_s, "自動延後產出時間");
                                AdddataList(tmp_s, "自動調撥其他可用工站");
                                break;
                            #endregion
                            case "04"://料應領未領
                                #region
                                AdddataList(tmp_s, "自動完成單據作業");
                                AdddataList(tmp_s, "自動延後產出時間");
                                AdddataList(tmp_s, "自動採用替代料");
                                AdddataList(tmp_s, "自動調撥其他可用量");
                                break;
                            #endregion
                            case "05"://APS_Simulation生產料件無單據發起
                                #region
                                AdddataList(tmp_s, "自動開立進貨單");
                                AdddataList(tmp_s, "自動延後產出時間");
                                break;
                            #endregion
                            case "06"://單據類(進貨類) 無預期IsOK=1 
                                #region
                                AdddataList(tmp_s, "警示進貨管理者追蹤");
                                AdddataList(tmp_s, "自動完成單據作業");
                                AdddataList(tmp_s, "自動延後產出時間");
                                break;
                            #endregion
                            case "07"://單據類(銷貨類)
                                #region
                                AdddataList(tmp_s, "警示銷貨管理者追蹤");
                                AdddataList(tmp_s, "自動完成單據作業");
                                AdddataList(tmp_s, "自動延後產出時間");
                                break;
                            #endregion
                            case "08"://單據類(存貨類)
                                #region
                                AdddataList(tmp_s, "警示倉儲管理者追蹤");
                                AdddataList(tmp_s, "自動延後產出時間");
                                break;
                            #endregion
                            case "09"://單據類(生產類)
                                #region
                                AdddataList(tmp_s, "警示生產管理者追蹤");
                                AdddataList(tmp_s, "自動完成單據作業");
                                AdddataList(tmp_s, "自動延後產出時間");
                                break;
                            #endregion
                            case "10"://單據類(委外類)
                                #region
                                AdddataList(tmp_s, "警示進貨管理者追蹤");
                                AdddataList(tmp_s, "自動完成單據作業");
                                AdddataList(tmp_s, "自動延後產出時間");
                                break;
                            #endregion
                            case "11"://工單應開未開
                                #region
                                AdddataList(tmp_s, "自動異動狀態為開站");
                                AdddataList(tmp_s, "自動延後產出時間");
                                AdddataList(tmp_s, "自動調撥其他可用工站");
                                break;
                            #endregion
                            case "12"://工單應關未關 (RUNTimeServer有安排 中午與晚上 自動檢查是否干涉關閉)
                                #region
                                AdddataList(tmp_s, "自動完成單據作業");
                                break;
                            #endregion
                            case "13"://CT or WT or UPH 未達有效值
                                #region
                                AdddataList(tmp_s, "警示生產管理者追蹤");
                                AdddataList(tmp_s, "自動延後產出時間");
                                break;
                            #endregion
                            case "14"://CT or WT or UPH 遠離目標值
                                #region
                                AdddataList(tmp_s, "警示生產管理者追蹤");
                                break;
                            #endregion
                            case "15"://監看人工報工是否延遲
                                #region
                                break;
                            #endregion
                            case "16"://監看工站工單是否如計畫
                                #region
                                AdddataList(tmp_s, "警示生產管理者追蹤");
                                break;
                            #endregion
                            case "17"://監看工站料件是否正常
                                #region
                                AdddataList(tmp_s, "自動調撥其他可用量");
                                AdddataList(tmp_s, "自動延後產出時間");
                                break;
                            #endregion
                            case "19"://APS_Simulation 的 預計完成日是否未達
                                #region
                                AdddataList(tmp_s, "自動延後產出時間");
                                break;
                                #endregion
                        }
                    }
                    if (dataList.Count > 0)
                    {
                        foreach (string s in dataList)
                        {
                            meg = $"{meg}<option value = '{s}'></option>";
                        }
                    }
                    meg = $"{meg}</datalist>";
                    is_Button = true;
                }
                else
                { meg = $"計畫編號:{data[1]} <p>無異常狀況<br /></p><input type='hidden' id='ageID' name='ageDATA' value=''>"; }
                #endregion

                #region 變更工站
                DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select a.*,b.Detail_QTY from SoftNetSYSDB.[dbo].[APS_Simulation] as a,SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as b where a.SimulationId='{data[1]}' and a.SimulationId=b.SimulationId");
                if (dr_tmp!=null && !dr_tmp.IsNull("Source_StationNO") && int.Parse(dr_tmp["Detail_QTY"].ToString()) ==0)
                {
                    if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{data[1]}' and IsOK='1'") <= 0 && db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[DOC4ProductionII] where SimulationId='{data[1]}' and IsOK='1'") <= 0)
                    {
                        is_Button = true;
                        string stationNO_Merge = dr_tmp.IsNull("StationNO_Merge") ? "" : $"共用站:{dr_tmp["StationNO_Merge"].ToString().Replace($"{dr_tmp["Source_StationNO"].ToString()},","")}";
                        if (dr_tmp["Source_StationNO"].ToString() == _Fun.Config.OutPackStationName)
                        {
                            meg = $"{meg}<p>原工作站:{dr_tmp["Source_StationNO"].ToString()}&emsp;{stationNO_Merge}<br />變更廠內工站生產:<input type='text' id='StationNO01' value=''><input type='hidden' id='StationNO02' value=''><input type='hidden' id='StationNO03' value=''></p>";
                        }
                        else
                        {
                            meg = $"{meg}<p>原工作站:{dr_tmp["Source_StationNO"].ToString()}&emsp;{stationNO_Merge}";
                            meg = $"{meg}<br />變更工作站:<input type='text' id='StationNO01' list='ageActionStationNO01' value=''><datalist id='ageActionStationNO01'><option value=''></option>";
                            DataTable dt_tmp02 = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}'");
                            if (dt_tmp02!=null && dt_tmp02.Rows.Count>0)
                            {
                                foreach(DataRow dr in dt_tmp02.Rows)
                                {
                                    meg = $"{meg}<option value='{dr["StationNO"].ToString()}'>{dr["StationNO"].ToString()}{dr["StationName"].ToString()}</option>";
                                }
                            }
                            meg = $"{meg}</datalist><br />增加共用站:<input type='text' id='StationNO02' value=''></p>";
                            meg = $"{meg} 或<br /><p>改委外加工:<input type='text' id='StationNO03' list='options_MFNO' value=''><datalist id='options_MFNO'> ";
                            dt_tmp02 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' order by MFNO");
                            if (dt_tmp02 != null && dt_tmp02.Rows.Count > 0)
                            {
                                foreach (DataRow d in dt_tmp02.Rows)
                                {
                                    meg = $"{meg}<option value='{d["MFNO"].ToString()}'>{d["MFNO"].ToString()};{d["MFName"].ToString()}</option>";
                                }
                            }
                            meg = $"{meg}</datalist><p>";
                        }
                    }
                    else { meg = $"{meg}<input type='hidden' id='StationNO01' value=''><input type='hidden' id='StationNO02' value=''><input type='hidden' id='StationNO03' value=''>"; }
                }
                else { meg = $"{meg}<input type='hidden' id='StationNO01' value=''><input type='hidden' id='StationNO02' value=''><input type='hidden' id='StationNO03' value=''>"; }
                #endregion

                if (is_Button)
                {
                    meg = $"{meg}<br /><button style= border: 2px blue none;' onclick=\"_me.SetAction('{data[1]}')\" >確認</button>";
                }
            }
            return Content(meg);
        }

        [HttpPost]
        public ActionResult OnActionII(string keys) //0=ipport,1.SimulationId,2=變更工作站,3=增加共用站,4=改委外加工廠商,5=.ActionID(type)6=動作名稱
        {
            string meg = "";
            string[] data = keys.Split(',');
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                try
                {
                    bool isOK = false;
                    string actionRemark = "";
                    string logTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    DataTable dt_tmp = null;
                    DataRow dr_tmp = null;

                    #region 處裡干涉異常
                    if (data.Length>=7 && data[6] != "")
                    {
                        if (data[6] == "不作為")
                        {
                            actionRemark = "不作為";
                            isOK = true;
                        }
                        else
                        {
                            //異動Action後的動作
                            string sql = "";

                            switch (data[6])
                            {
                                case "自動開立進貨單":
                                    #region
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($@"SELECT a.*,b.SimulationDate FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a
                                                                        join SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.SimulationId=b.SimulationId                             
                                                                        where a.SimulationId='{data[1]}'");
                                        if (dr_tmp != null)
                                        {
                                            string storeNO = "";
                                            string storeSpacesNO = "";
                                            string docNumberNO = "";

                                            #region 查找適合庫儲別
                                            _SFC_Common.SelectINStore(db, dr_tmp["PartNO"].ToString(), ref storeNO, ref storeSpacesNO, "AA02");
                                            #endregion
                                            int qty = int.Parse(dr_tmp["NeedQTY"].ToString());
                                            #region 查找適合數量
                                            DataRow dr_tmp2 = db.DB_GetFirstDataByDataRow($"select sum(a.QTY)-sum(b.KeepQTY + b.OverQTY) as Total from SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.PartNO = '{dr_tmp["PartNO"].ToString()}' and a.Id = b.Id and a.StoreNO!=''");
                                            if (dr_tmp2 != null && !dr_tmp2.IsNull("Total"))
                                            {
                                                qty = Math.Abs(int.Parse(dr_tmp2["Total"].ToString()));
                                            }
                                            #endregion
                                            #region 查找適合廠商
                                            float price = 0;
                                            string mFNO = _SFC_Common.SelectDOC1BuyMFNO(db, dr_tmp["PartNO"].ToString(), dr_tmp["SimulationId"].ToString(), "", ref price);
                                            #endregion

                                            //###???StartDate暫時, 完成日期要計算採購的CT
                                            if (!dr_tmp.IsNull("SimulationDate"))
                                            {
                                                isOK = _SFC_Common.Create_DOC1stock(db, dr_tmp, mFNO, price, storeNO, storeSpacesNO, "AA02", qty, "", "", "異常後人工介入干涉", logTime, Convert.ToDateTime(dr_tmp["StartDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "系統指派", ref docNumberNO);
                                                if (isOK) { actionRemark = $"加開採購單:{docNumberNO}"; }
                                            }
                                        }
                                    }
                                    #endregion
                                    break;
                                case "自動異動狀態為開站":
                                    #region
                                    {
                                        switch (data[5])
                                        {
                                            case "03":
                                                {
                                                    //###??? 來源要加工站,不然不知是主站,還是合併站 ?
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{data[1]}'");
                                                    if (dr_tmp != null)
                                                    {
                                                        DataRow dr_tmp2 = null;
                                                        DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT StationNO from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_tmp["Source_StationNO"].ToString()}'");
                                                        DataRow dr_M = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_tmp["Source_StationNO"].ToString()}'");
                                                        if (dr_PP_Station["Station_Type"].ToString() == "8")
                                                        {
                                                            #region 多工單模式的工站
                                                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_tmp["NeedId"].ToString()}','{dr_tmp["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','APSView','干涉開工','{dr_tmp["Apply_PP_Name"].ToString()}','{dr_tmp["Source_StationNO"].ToString()}','{dr_tmp["PartNO"].ToString()}','{dr_tmp["DOCNumberNO"].ToString()}','系統指派',{dr_tmp["Source_StationNO_IndexSN"].ToString()})");
                                                            DataRow dr_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"SELECT *,NeedQTY as PNQTY from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where PartNO='{dr_tmp["PartNO"].ToString()}' and SimulationId='{dr_tmp["SimulationId"].ToString()}' and NeedId='{dr_tmp["NeedId"].ToString().Trim()}'");
                                                            dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_tmp["Source_StationNO"].ToString()}' and OrderNO='{dr_tmp["DOCNumberNO"].ToString().Trim()}'");
                                                            if (dr_tmp2 != null)
                                                            { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[ManufactureII] SET PNQTY={dr_APS_PartNOTimeNote["PNQTY"].ToString()} and OrderNO='{dr_tmp["DOCNumberNO"].ToString().Trim()}',IndexSN={dr_tmp["Source_StationNO_IndexSN"].ToString()},Station_Custom_IndexSN='{dr_tmp["Source_StationNO_Custom_IndexSN"].ToString()}',StationNO_Custom_DisplayName='{dr_tmp["Source_StationNO_Custom_DisplayName"].ToString()}',Master_PP_Name='{dr_tmp["Apply_PP_Name"].ToString()}',PP_Name='{dr_tmp["Apply_PP_Name"].ToString()}',SimulationId='{dr_tmp["SimulationId"].ToString()}',StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}',RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}',EndTime=NULL,RemarkTimeE=NULL where Id='{dr_tmp2["Id"].ToString()}'"); }
                                                            else
                                                            {
                                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[ManufactureII] ([Id],[StationNO],[ServerId],[OrderNO],[Master_PP_Name],[PP_Name],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],[SimulationId],StartTime,RemarkTimeS,PNQTY)
                                                                            VALUES ('{_Str.NewId('C')}','{dr_tmp["Source_StationNO"].ToString()}','{_Fun.Config.ServerId}','{dr_tmp["DOCNumberNO"].ToString()}','{dr_tmp["Apply_PP_Name"].ToString()}','{dr_tmp["Apply_PP_Name"].ToString()}',{dr_tmp["Source_StationNO_IndexSN"].ToString()},'{dr_tmp["StationNO_Custom_DisplayName"].ToString()}','{dr_tmp["Source_StationNO_Custom_DisplayName"].ToString()}','{dr_tmp["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'),{dr_APS_PartNOTimeNote["PNQTY"].ToString()}");
                                                            }
                                                            actionRemark = "自動異動狀態為開站";
                                                            isOK = true;
                                                            #endregion

                                                        }
                                                        else
                                                        {
                                                            #region 單工單模式的工站
                                                            if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET State='1',OrderNO='{dr_tmp["DOCNumberNO"].ToString().Trim()}',IndexSN={dr_tmp["Source_StationNO_IndexSN"].ToString()},Station_Custom_IndexSN='{dr_tmp["Source_StationNO_Custom_IndexSN"].ToString()}',StationNO_Custom_DisplayName='{dr_tmp["Source_StationNO_Custom_DisplayName"].ToString()}',OP_NO='排程指派',Master_PP_Name='{dr_tmp["Apply_PP_Name"].ToString()}',PP_Name='{dr_tmp["Apply_PP_Name"].ToString()}',SimulationId='{data[1]}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_tmp["Source_StationNO"].ToString()}'"))//Master_PP_Name應該還要找有無上層
                                                            {
                                                                //db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[Manufacture_Log] ([Id],[logDate],[StationNO],[State],[OrderNO],[PartNO]) VALUES ('{_Str.NewId('C')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{dr_tmp["Source_StationNO"].ToString()}','1','{dr_tmp["DOCNumberNO"].ToString()}','{dr_tmp["PartNO"].ToString()}')");
                                                                db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_tmp["NeedId"].ToString()}','{dr_tmp["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','APSView','干涉開工','{dr_tmp["Apply_PP_Name"].ToString()}','{dr_tmp["Source_StationNO"].ToString()}','{dr_tmp["PartNO"].ToString()}','{dr_tmp["DOCNumberNO"].ToString()}','系統指派',{dr_tmp["Source_StationNO_IndexSN"].ToString()})");

                                                                actionRemark = "自動異動狀態為開站";
                                                                isOK = true;
                                                            }
                                                            #endregion
                                                        }
                                                        #region 送Service處理
                                                        //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                                        if (dr_tmp != null && _WebSocket.RmsSend(dr_PP_Station["RMSName"].ToString(), 1, $"WebChangeStationStatus,1,{dr_tmp["Source_StationNO"].ToString()},WEBProg,{dr_tmp["Source_StationNO"].ToString()},'排程指派',{dr_tmp["DOCNumberNO"].ToString()},{dr_tmp["Source_StationNO_IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
                                                        {
                                                            if (dr_PP_Station["Station_Type"].ToString() == "1")
                                                            {
                                                                #region 更新標籤
                                                                var ledstate = "0";
                                                                string ledrgb = "ff00";
                                                                DataRow totalData = _Fun.GetAvgCTWTandTotalOutput(db, false, dr_M["OrderNO"].ToString(), dr_M["StationNO"].ToString(), dr_M["IndexSN"].ToString());
                                                                string dis_DetailQTY = "0";
                                                                if (totalData != null)
                                                                {
                                                                    dis_DetailQTY = totalData["TotalOutput"].ToString().Trim();
                                                                }
                                                                int ct = 0;
                                                                int num = 0;
                                                                string simulationId = "";
                                                                string partNO = "";
                                                                string partName = "";
                                                                string typevalue = $"0;";
                                                                string isUpdate = "1";
                                                                DataRow dr_LabelStateINFO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                                                if (dr_LabelStateINFO != null)
                                                                {
                                                                    DataRow dr_Staion = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_M["StationNO"].ToString()}'");

                                                                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_tmp["NeedId"].ToString()}' and Source_StationNO='{dr_M["StationNO"].ToString()}' and Source_StationNO_IndexSN={dr_M["IndexSN".ToString()]} and DOCNumberNO='{dr_M["OrderNO"].ToString().Trim()}'");
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
                                                                            partName = tmp_dr["PartName"].ToString().Replace("\"","'");
                                                                        }

                                                                    }

                                                                    string tmp_s = $"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{dr_M["StationNO"].ToString()};{typevalue};{dr_M["IndexSN"].ToString()}";
                                                                    var json1 = "";
                                                                    var json_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{dr_M["OrderNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{partNO}\",\"Text3\":\"品名:\",\"PartName\":\"{partName}\",\"Text4\":\"需求量:\",\"DetailQTY\":\"{dis_DetailQTY}\",\"Text5\":\"CT:\",\"EfficientCT\":\"{ct.ToString()}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                                                                    var writeShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{dr_M["OrderNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{partNO}\",\"Text3\":\"品名:\",\"PartName\":\"{partName}\",\"Text4\":\"需求量:\",\"QTY\":\"{num.ToString()}\",\"Text5\":\"CT:\",\"EfficientCT\":\"{ct.ToString()}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                                                                    if (dr_LabelStateINFO["Version"].ToString().Trim() != "" && dr_LabelStateINFO["Version"].ToString().Trim().Substring(0, 2) == "42")
                                                                    {
                                                                        json_ShowValue = $"{json_ShowValue},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_Staion["StationNO"].ToString()}\",\"text17\":\"{dr_Staion["StationName"].ToString()}\"";
                                                                        writeShowValue = $"{writeShowValue},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_Staion["StationNO"].ToString()}\",\"text17\":\"{dr_Staion["StationName"].ToString()}\"";
                                                                        json1 = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{json_ShowValue}";
                                                                    }
                                                                    else
                                                                    {
                                                                        json_ShowValue = $"{json_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                                                        writeShowValue = $"{writeShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                                                        json1 = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{json_ShowValue}";
                                                                    }

                                                                    if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                                    {
                                                                        var json = $"{json1},\"QTY\":\"{num.ToString()}\",\"ledrgb\":\"{ledrgb}\",\"ledstate\":{ledstate}";
                                                                        _Fun.Tag_Write(db,dr_M["Config_macID"].ToString(),$"干涉開工", json);
                                                                    }
                                                                    else { isUpdate = "0"; }
                                                                    if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set Ledrgb='{ledrgb}',Ledstate=0,ShowValue='{writeShowValue}',StationNO='{dr_M["StationNO"].ToString()}',Type='1',OrderNO='{dr_M["OrderNO"].ToString()}',IndexSN='{dr_M["IndexSN"].ToString()}',StoreNO='',StoreSpacesNO='',QTY={dis_DetailQTY},IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'"))
                                                                    {

                                                                    }
                                                                }
                                                                #endregion
                                                            }
                                                            //###???通知單工網頁更新
                                                        }
                                                        else
                                                        {
                                                            isOK = false;
                                                            actionRemark = "";
                                                            meg = $"{meg}<br />工站:{dr_tmp["Source_StationNO"].ToString()} Service無作用,無法設定工站狀態";
                                                        }
                                                        #endregion

                                                    }
                                                }
                                                break;
                                            case "11":
                                                //###??? 沒寫
                                                break;
                                        }

                                    }
                                    #endregion
                                    break;
                                case "自動調撥其他可用量":
                                    #region
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select PartNO,NeedQTY from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{data[1]}'");
                                        dt_tmp = db.DB_GetData($"select Id,sum(QTY) as qty from SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_tmp["PartNO"].ToString()}' group by Id");
                                        if (dt_tmp != null)
                                        {
                                            int needQTY = int.Parse(dr_tmp["NeedQTY"].ToString());
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"select sum(NeedQTY) as h_qty from SoftNetMainDB.[dbo].[TotalStockII] where SimulationId='{data[1]}'");
                                            if (dr_tmp != null) { needQTY -= int.Parse(dr_tmp["h_qty"].ToString()); }
                                            if (needQTY > 0)
                                            {
                                                Dictionary<string, int> id_totQTY = new Dictionary<string, int>();
                                                foreach (DataRow d in dt_tmp.Rows)
                                                {
                                                    id_totQTY.Add(d["Id"].ToString(), int.Parse(d["qty"].ToString()));
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select sum(KeepQTY) as tot_qty from SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{d["Id"].ToString()}'");
                                                    if (dr_tmp != null && !dr_tmp.IsNull("tot_qty") && dr_tmp["tot_qty"].ToString().Trim() != "") { id_totQTY[d["Id"].ToString()] -= int.Parse(dr_tmp["tot_qty"].ToString()); }
                                                }
                                                //從數量較少的倉先扣起
                                                if (id_totQTY.Count > 0)
                                                {
                                                    var dictSort = from objDict in id_totQTY orderby objDict.Value select objDict;
                                                    foreach (KeyValuePair<string, int> kvp in dictSort)
                                                    {
                                                        if ((kvp.Value - needQTY) >= 0)
                                                        {
                                                            sql = $@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII] ([Id],[NeedId],[SimulationId],[KeepQTY],ArrivalDate) VALUES (
                                                            '{_Str.NewId('Z')}',
                                                            '{kvp.Key}',
                                                            '{data[1]}',
                                                            {(kvp.Value - needQTY)},'{logTime}')";
                                                            db.DB_SetData(sql);
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            if (kvp.Value > 0)
                                                            {
                                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[TotalStockII] where SimulationId='{data[1]}'");
                                                                if (dr_tmp != null)
                                                                { sql = $"update SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY+={kvp.Value} where Id='{dr_tmp["SimulationId"].ToString()}' and NeedId='{dr_tmp["NeedId"].ToString()}' and SimulationId='{dr_tmp["SimulationId"].ToString()}'"; }
                                                                else
                                                                {
                                                                    sql = $@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII] ([Id],[NeedId],[SimulationId],[KeepQTY],ArrivalDate) VALUES (
                                                                    '{_Str.NewId('Z')}',
                                                                    '{kvp.Key}',
                                                                    '{data[1]}',
                                                                    {kvp.Value},'{logTime}')";
                                                                }
                                                                db.DB_SetData(sql);
                                                            }
                                                        }
                                                        needQTY -= kvp.Value;
                                                    }
                                                }
                                                if (needQTY > 0)
                                                {
                                                    string docNumberNO = "";
                                                    string storeNO = "";
                                                    string storeSpacesNO = "";
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($@"SELECT *,b.StartDate FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a
                                                                        join SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.SimulationId=b.SimulationId                             
                                                                        where a.SimulationId='{data[1]}'");
                                                    DataRow tmp_dr2 = db.DB_GetFirstDataByDataRow($"SELECT a..*,b.StoreNO,b.StoreSpacesNO  FROM SoftNetMainDB.[dbo].[TotalStockII] as a, SoftNetMainDB.[dbo].[TotalStock] as b  where b.Class!='虛擬倉' and b.ServerId='{_Fun.Config.ServerId}' and a.SimulationId={dr_tmp["SimulationId"].ToString()} order by a.Id");
                                                    if (tmp_dr2 != null)
                                                    {
                                                        storeNO = tmp_dr2["StoreNO"].ToString();
                                                        storeNO = tmp_dr2["StoreSpacesNO"].ToString();
                                                    }
                                                    else
                                                    {

                                                        #region 查找適合庫儲別
                                                        _SFC_Common.SelectINStore(db, dr_tmp["PartNO"].ToString(), ref storeNO, ref storeSpacesNO, "AA02");
                                                        #endregion
                                                    }
                                                    #region 查找適合廠商
                                                    float price = 0;
                                                    string mFNO = _SFC_Common.SelectDOC1BuyMFNO(db, dr_tmp["PartNO"].ToString(), dr_tmp["SimulationId"].ToString(), "", ref price);
                                                    #endregion

                                                    isOK = _SFC_Common.Create_DOC1stock(db, dr_tmp, mFNO, price, storeNO, storeSpacesNO, "AA02", needQTY, "", "", "異常後人工介入干涉", logTime, Convert.ToDateTime(dr_tmp["StartDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "系統指派", ref docNumberNO);
                                                    if (isOK)
                                                    { meg = $"{meg}<br />料號:{dr_tmp["PartNO"].ToString()} 調撥後量還是不足,加開採購單:{docNumberNO}"; }
                                                    else
                                                    { meg = $"{meg}<br />料號:{dr_tmp["PartNO"].ToString()} 調撥後量還是不足"; }
                                                }
                                                actionRemark = "自動調撥其他可用量";
                                                isOK = true;
                                            }
                                            else
                                            {
                                                actionRemark = "不作為";
                                                isOK = true;
                                            }
                                        }
                                    }
                                    #endregion
                                    break;
                                case "自動完成單據作業": //04料應領未領(應該不會進入)   06,07,09,10單據類IsOK=0過期   12工單應開未開(應該不會進入)
                                    #region
                                    {
                                        switch (data[5])
                                        {
                                            case "06":
                                                #region
                                                //###???
                                                #endregion
                                                break;
                                            case "07":
                                                #region
                                                //###???
                                                #endregion
                                                break;
                                            case "08":
                                                /* //###????
                                                #region 寫入庫存
                                                string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
                                                DataRow dr_DOC3stockII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where Id='{d["ErrorKey"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}'");
                                                if (dr_DOC3stockII != null)
                                                {

                                                    if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                                                    { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={dr_DOC3stockII["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["IN_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}'"); }
                                                    else { db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={dr_DOC3stockII["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'"); }

                                                    #region 計算單據CT,平均,有效, 寫SFC_StationProjectDetail
                                                    int typeTotalTime = 0;
                                                    if (!dr_DOC3stockII.IsNull("StartTime")) { typeTotalTime = _WebSocket.TimeCompute2Seconds(Convert.ToDateTime(dr_DOC3stockII["StartTime"]), DateTime.Now); }
                                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime} where Id='{d["ErrorKey"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}'");

                                                    string partNO = dr_DOC3stockII["PartNO"].ToString();
                                                    string pp_Name = "";
                                                    string E_stationNO = "";
                                                    if (dr_DOC3stockII["SimulationId"].ToString() != "")
                                                    {
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC3stockII["SimulationId"].ToString()}'");
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
                                                        _WebSocket.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, pp_Name, pp_Name, partNO, partNO, dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4));
                                                    }
                                                    #endregion
                                                }
                                                #endregion
                                                */

                                                break;
                                            case "09":
                                                #region
                                                //###???
                                                #endregion
                                                break;
                                            case "10":
                                                #region
                                                //###???
                                                #endregion
                                                break;
                                        }
                                        if (isOK) { actionRemark = "自動完成單據作業"; }
                                    }
                                    #endregion
                                    break;
                                case "自動追加生產數量": //02
                                    #region
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select PartNO,(NeedQTY+SafeQTY) as Nqty from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{data[1]}'");
                                        int nQTY = int.Parse(dr_tmp["Nqty"].ToString());
                                        dt_tmp = db.DB_GetData($"select Id,sum(KeepQTY+OverQTY) as Kqty from SoftNetMainDB.[dbo].[TotalStockII] where SimulationId='{data[1]}' Group by Id");
                                        int tot = 0;//keep倉庫ID總量
                                        if (dt_tmp != null)
                                        {
                                            foreach (DataRow dr in dt_tmp.Rows)
                                            {
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select QTY from SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and Id='{dr["Id"].ToString()}'");
                                                if (dr_tmp != null)
                                                {
                                                    tot += int.Parse(dr_tmp["QTY"].ToString());
                                                }
                                            }
                                            foreach (DataRow dr in dt_tmp.Rows)
                                            {
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select a.Id,sum(a.QTY) as mQTY,sum(b.KeepQTY+b.OverQTY) as kQTY from SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and a.Id='{dr["Id"].ToString()}' and b.SimulationId!='{data[1]}' group by a.Id,b.Id");
                                                if (dr_tmp != null && !dr_tmp.IsNull("mQTY") && !dr_tmp.IsNull("kQTY"))
                                                {
                                                    int i1 = Math.Abs(int.Parse(dr_tmp["mQTY"].ToString()) - int.Parse(dr_tmp["kQTY"].ToString()));
                                                    tot -= i1;
                                                }
                                            }
                                        }
                                        tot = Math.Abs(tot);
                                        if (nQTY > tot)
                                        {
                                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set NeedQTY+={(nQTY - tot)} where SimulationId='{data[1]}'");
                                            actionRemark = "自動追加生產數量";
                                            isOK = true;
                                        }
                                    }
                                    #endregion
                                    break;
                                case "自動延後產出時間":
                                    {
                                        //lock (_Fun.Lock_Simulation_Flag)
                                        //{
                                            List<string> run_HasWeb_Id_Change = new List<string>();

                                            if (_SFC_Common.RefreshRunSetSimulation(db, data, ref run_HasWeb_Id_Change))
                                            {
                                                actionRemark = "自動延後產出時間";
                                                isOK = true;
                                            }
                                            else
                                            {
                                                string _s = "";//###???
                                            }
                                            foreach (string sno in run_HasWeb_Id_Change)
                                            {
                                                #region 通知網頁更新
                                                try
                                                {
                                                    lock (_WebSocket.lock__WebSocketList)
                                                    {
                                                        foreach (KeyValuePair<string, rmsConectUserData> r in _WebSocket._WebSocketList)
                                                        {
                                                            if (r.Key != null && r.Value.socket != null)
                                                            {
                                                                _WebSocket.Send(r.Value.socket, $"HasWeb_Id_Change,STView2Work_PageReload,{sno}");
                                                            }
                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"APSViewController.cs 發HasWeb_Id_Changeg失敗 {ex.Message} {ex.StackTrace}", true);
                                                }
                                                #endregion
                                            }
                                        //}
                                    }
                                    break;
                            }
                        }
                        if (isOK)
                        {
                            if (!db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] set ActionType='0',ActionRemark='{actionRemark}',ActionLogDate='{logTime}' where ServerId='{_Fun.Config.ServerId}' and SimulationId='{data[1]}'"))
                            {
                                //###???
                            }
                        }
                    }
                    #endregion

                    #region 處裡工站異動
                    if (data[2] != "" || data[3] != "" || data[4] != "")
                    {
                        if (data[4] != "" && (data[2] != "" || data[3] != "")) { meg = $"{meg}<br />改委外加工後, 不能同時異動工作站 或 增加共用站"; }
                        else
                        {
                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{data[1]}'");
                            if (dr_tmp != null)
                            {
                                if (data[4] != "")
                                {
                                    #region 改委外加工
                                    /*
                                    if (!dr_tmp.IsNull("StationNO_Merge"))
                                    {
                                        foreach(string s in dr_tmp["StationNO_Merge"].ToString().Split(','))
                                        {
                                            if (s.Trim()!="" && s!= dr_tmp["Source_StationNO"].ToString())
                                            {
                                                db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] where OrderNO='{dr_tmp["DOCNumberNO"].ToString()}' and StationNO='{s}' and IndexSN={dr_tmp["Source_StationNO_IndexSN"].ToString()} and IndexSN_Merge='1'");
                                            }
                                        }
                                    }
                                    */
                                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] set StationNO='{_Fun.Config.OutPackStationName}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_tmp["Source_StationNO"].ToString()}' and OrderNO='{dr_tmp["DOCNumberNO"].ToString()}' and IndexSN={dr_tmp["Source_StationNO_IndexSN"].ToString()}");
                                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set Apply_StationNO='{_Fun.Config.OutPackStationName}' where NeedId='{dr_tmp["NeedId"].ToString()}' and Apply_PP_Name='{dr_tmp["Apply_PP_Name"].ToString()}' and Master_PartNO='{dr_tmp["PartNO"].ToString()}' and Apply_StationNO='{dr_tmp["Source_StationNO"].ToString()}' and IndexSN='{dr_tmp["Source_StationNO_IndexSN"].ToString()}'");
                                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set Source_StationNO='{_Fun.Config.OutPackStationName}',OutPackType='1',StationNO_Merge=NLLL where SimulationId='{data[1]}'");
                                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set APS_StationNO='{_Fun.Config.OutPackStationName}',OutPackType='1',MFNO='{data[4]}',StationNO_Merge=NLLL where SimulationId='{data[1]}'");
                                    db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where SimulationId='{data[1]}'");
                                    db.DB_SetData($"delete from SoftNetMainDB.[dbo].[TotalStockII] where SimulationId='{data[1]}'");
                                    db.DB_SetData($"delete from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{data[1]}'");
                                    #endregion
                                }
                                else
                                {
                                    string StationNO_Merge = dr_tmp.IsNull("StationNO_Merge") ? "" : dr_tmp["StationNO_Merge"].ToString();
                                    if (data[3] != "")
                                    {
                                        #region 增加共用站
                                        foreach (string s in data[3].Split(';'))
                                        {
                                            DataRow dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{s}'");
                                            if (dr_tmp2 != null)
                                            {
                                                if (StationNO_Merge.IndexOf($"{s},") < 0)
                                                {
                                                    StationNO_Merge = $"{StationNO_Merge}{s},";
                                                }
                                            }
                                        }
										string tmp_StationNO_Merge = StationNO_Merge;
										if (tmp_StationNO_Merge == "") { tmp_StationNO_Merge = "NULL"; }
                                        else { tmp_StationNO_Merge = $"'{tmp_StationNO_Merge}'"; }
										db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set StationNO_Merge={tmp_StationNO_Merge} where SimulationId='{data[1]}'");
                                        db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set StationNO_Merge={tmp_StationNO_Merge} where SimulationId='{data[1]}'");
                                        db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set StationNO_Merge={tmp_StationNO_Merge} where SimulationId='{data[1]}'");
                                        #endregion
                                    }
                                    if (data[2] != "")
                                    {
                                        #region 變更工作站
                                        DataRow dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[2]}'");
                                        if (dr_tmp2 != null)
                                        {
                                            string name = dr_tmp2["StationName"].ToString();
                                            string orderNO = "";
                                            #region 找尋適合工單 PS:目前是錯的
                                            if (dr_tmp["Source_StationNO"].ToString() == _Fun.Config.OutPackStationName || dr_tmp["DOCNumberNO"].ToString().Trim().IndexOf("XX01") != 0 || dr_tmp["DOCNumberNO"].ToString() == "")
                                            {
                                                dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_tmp["NeedId"].ToString()}' and DOCNumberNO like 'XX01%'");
                                                if (dr_tmp2 != null)
                                                {
                                                    orderNO = dr_tmp2["DOCNumberNO"].ToString();
                                                }
                                            }
                                            else { orderNO = dr_tmp["DOCNumberNO"].ToString(); }
                                            #endregion
                                            if (orderNO != "")
                                            {
                                                if (StationNO_Merge == "") { StationNO_Merge = "NULL"; }
                                                else
                                                {
                                                    if (StationNO_Merge.IndexOf($"{data[2]},") < 0)
                                                    {
                                                        StationNO_Merge = $"{StationNO_Merge}{data[2]},";
                                                    }
                                                    StationNO_Merge = $"'{StationNO_Merge}'";
                                                }
                                                if (dr_tmp["Source_StationNO"].ToString() == _Fun.Config.OutPackStationName)
                                                {
                                                    db.DB_SetData($"delete from SoftNetMainDB.[dbo].[DOC4ProductionII] where SimulationId='{data[1]}' and IsOK='0'");
                                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type1,Time1_C,Time_TOT) VALUES 
                                                            ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{data[2]}','{Convert.ToDateTime(dr_tmp["StartDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}','{dr_tmp["NeedId"].ToString()}','{dr_tmp["SimulationId"].ToString()}','1',600,600)");
                                                }
                                                else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set StationNO='{data[2]}',StationNO_Merge={StationNO_Merge} where SimulationId='{data[1]}'"); }
                                                db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set Apply_StationNO='{data[2]}' where NeedId='{dr_tmp["NeedId"].ToString()}' and Apply_PP_Name='{dr_tmp["Apply_PP_Name"].ToString()}' and Master_PartNO='{dr_tmp["PartNO"].ToString()}' and Apply_StationNO='{dr_tmp["Source_StationNO"].ToString()}' and IndexSN='{dr_tmp["Source_StationNO_IndexSN"].ToString()}'");
                                                db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set Source_StationNO='{data[2]}',StationNO_Merge={StationNO_Merge} where SimulationId='{data[1]}'");
                                                db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set APS_StationNO='{data[2]}',StationNO_Merge={StationNO_Merge},OutPackType='0' where SimulationId='{data[1]}'");
                                                dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[2]}' and IndexSN={dr_tmp["Source_StationNO_IndexSN"].ToString()} and OrderNO='{orderNO}'");
                                                if (dr_tmp2 != null)
                                                {
                                                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] set StationNO='{data[2]}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_tmp["Source_StationNO"].ToString()}' and OrderNO='{orderNO}' and IndexSN={dr_tmp["Source_StationNO_IndexSN"].ToString()}");
                                                }
                                                else
                                                {
                                                    string sfc_sql = string.Format(
                                                                     @"INSERT INTO SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement (
                                                                                    [OrderNO],[StationNO],[StationName],[PP_Name],[IndexSN],[DisplaySN],
                                                                                    [IsLastStation],[Sub_PP_Name],[StatndCycleTime],[UpdateTime],[IndexSN_Merge],
                                                                                    [StartTime],[CumulativeTime],[AvarageCycleTime],[TotalCheckIn],[TotalCheckOut],
                                                                                    [TotalInput],[TotalOutput],[TotalFail],[TotalKeep],[FPY],
                                                                                    [YieldRate],[StationYieldRate],ServerId) VALUES 
                                                                                    ('{0}',N'{1}','{2}','{3}',{4},{5},'{6}','{7}',{8},'{9}','{10}',
                                                                                    null,0,0,0,0,0,0,0,0,0,0,0,'{11}')",
                                                                    orderNO, //0
                                                                    data[2], //1
                                                                    name, //2
                                                                    dr_tmp["Apply_PP_Name"].ToString(), //3
                                                                    dr_tmp["Source_StationNO_IndexSN"].ToString(),//4
                                                                    "0",//5 DisplaySN
                                                                    dr_tmp["PartSN"].ToString() != "0" ? "0" : "1", //6
                                                                    dr_tmp["Apply_PP_Name"].ToString(), //7 Sub_PP_Name
                                                                    "0", //8
                                                                    DateTime.Now.ToString("MM/dd/yyyy H:mm:ss"), //9
                                                                    StationNO_Merge== "NULL" ? "0":"1", _Fun.Config.ServerId);//10
                                                    db.DB_SetData(sfc_sql);
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                }
                            }
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    meg = $"{meg}<br />工站:干涉失敗, 原因:{ex.Message}";
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"{meg}<br />工站:干涉失敗, 原因: {ex.Message} {ex.StackTrace}", true);
                }
            }
            return Content(meg);
        }

        private List<string> dataList = new List<string>();
        private void AdddataList(string key, string meg)
        {
            if (!dataList.Contains($"{key},{meg}"))
            { dataList.Add($"{key},{meg}"); }
        }


        public IActionResult APSSETRead()
        {
            return View();
        }
        public IActionResult Read(APSViewData key)
        {
            string re = "";
            DBADO db = new DBADO("1", _Fun.Config.Db);
            if (key == null) { key = new APSViewData(); }
            selectSIDList(db, ref key);
            if (key.SelectNeedId == null)
            {
                re = $"<div>沒正確選擇適合的計畫編號</div>";
            }
            else
            {
                if (key.SelectNeedId.Trim() != "")
                {
                    //取得第一層結構
                    DataTable dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{key.SelectNeedId}' and PartSN=1 order by PartSN");
                    if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
                    {
                        string show = "label";
                        DataRow mast = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{key.SelectNeedId}' and PartSN=0");
                        if (db.DB_GetQueryCount($"SELECT SimulationId FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{mast["SimulationId"].ToString()}' and ActionType=''") > 0) { show = "labelRed"; }
                        string s_date = mast.IsNull("StartDate") ? "" : Convert.ToDateTime(mast["StartDate"]).ToString("yyyy-MM-dd HH:mm");
                        re = $"<div id='wrapper'><button id='{mast["SimulationId"].ToString()}' class='{show}' style='border: 2px blue none;' onclick=\"_me.onShow('{mast["SimulationId"].ToString()}')\">{mast["Master_PartNO"].ToString()} [{mast["Source_StationNO"].ToString()}]<br />開始日:{s_date}<br />計劃日:{Convert.ToDateTime(mast["SimulationDate"]).ToString("yyyy-MM-dd HH:mm")}</button>\n";
                        RecursiveSimulation(db, key.SelectNeedId, "0", dt_Simulation.Rows[0]["PartSN"].ToString().Trim(), dt_Simulation, ref re);
                        re = $"{re}</div>\n";
                    }
                }
            }
            db.Dispose();
            ViewBag.HtmlOutput = re;
            return View(key);
        }
        public IActionResult ProductCourse(APSViewData key)
        {
            string re = "";
            DBADO db = new DBADO("1", _Fun.Config.Db);
            if (key == null || key.SelectFun1 == "")
            {
                key = new APSViewData();
                key.SelectFun2 = "A3";
                key.SelectFun3 = "2";
                key.SelectFun4 = "D2";
            }
            old_selectSIDList(db, ref key, "69");
            if (key.SelectFun1 == "") { return View(key); }

            if (key.NeedIdList.Count <= 0)
            {
                re = $"<div>目前沒有任何生產歷程的計畫編號</div>";
            }
            else
            {
                DataTable dt = null;
                DataTable dt2 = null;
                DataRow dr_tmp = null;
                string sType1 = "";
                string sType2 = "";
                string tmp_s = "";
                string s_date = "";
                List<string> needID_List = new List<string>();
                if (key.SelectNeedId != null && key.SelectNeedId.Trim() != "")
                {
                    needID_List.Add(key.SelectNeedId.Trim());
                }
                else
                {
                    int betime = -1;
                    int.TryParse(key.SelectFun3, out betime);
                    if (key.SelectFun4 == "D1") //只針對排程異常警示
                    {
                        dt = db.DB_GetData($"SELECT NeedId from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId!='' and LogDate>='{DateTime.Now.AddDays(-betime).ToString("yyyy/MM/dd HH:mm:ss.fff")}' group by NeedId");
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dt.Rows)
                            {
                                if (!needID_List.Contains(dr["NeedId"].ToString())) { needID_List.Add(dr["NeedId"].ToString()); }
                            }
                        }
                        dt = db.DB_GetData($"SELECT NeedId from SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId!='' and WarningDate>='{DateTime.Now.AddDays(-betime).ToString("yyyy/MM/dd HH:mm:ss.fff")}' group by NeedId");
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dt.Rows)
                            {
                                if (!needID_List.Contains(dr["NeedId"].ToString())) { needID_List.Add(dr["NeedId"].ToString()); }
                            }
                        }
                    }
                    else if (key.SelectFun3 == "D2") //只針對數據疑似異常
                    {

                    }

                    dt = db.DB_GetData($"SELECT NeedId from SoftNetLogDB.[dbo].[OperateLog] where NeedId!='' and LOGDateTime>='{DateTime.Now.AddDays(-betime).ToString("yyyy/MM/dd HH:mm:ss.fff")}' group by NeedId");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            if (!needID_List.Contains(dr["NeedId"].ToString())) { needID_List.Add(dr["NeedId"].ToString()); }
                        }
                    }
                }
                if (needID_List.Count > 0)
                {
                    string ids = "";
                    foreach (string s in needID_List)
                    {
                        if (ids == "") { ids = $"'{s}'"; }
                        else { ids = $"{ids},'{s}'"; }
                    }

                    dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where Id in ({ids})");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        string needId = "";
                        foreach (DataRow dr_APS_NeedData in dt.Rows)
                        {
                            needId = dr_APS_NeedData["Id"].ToString();
                            #region 流程圖
                            tmp_s = "";
                            DataRow dr_Material = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where PartNO='{dr_APS_NeedData["PartNO"].ToString()}'");
                            if (key.SelectFun2!=null && key.SelectFun2 != "A3")
                            {
                                if (key.SelectFun2 == "A1" && dr_Material[""].ToString() != "4") { continue; }
                                if (key.SelectFun2 == "A2" && dr_Material[""].ToString() != "5") { continue; }
                            }
                            if (dr_APS_NeedData["State"].ToString() == "6") { sType1 = "已轉計畫"; }
                            else { sType1 = "已入庫"; }
                            if (dr_APS_NeedData["NeedType"].ToString() == "1") { sType2 = $"訂單需求"; }
                            else if (dr_APS_NeedData["NeedType"].ToString() == "2") { sType2 = $"客戶需求"; }
                            else if (dr_APS_NeedData["NeedType"].ToString() == "5") { sType2 = $"底稿發出"; }
                            else { sType2 = $"廠內需求"; }
                            re = $"{re}\n<div class='shadowbox'><fieldset id='{needId}'><legend>料號:{dr_APS_NeedData["PartNO"].ToString()} {sType2} {sType1} 預計入庫日:{Convert.ToDateTime(dr_APS_NeedData["NeedSimulationDate"].ToString()).ToString("yyyy-MM-dd HH:mm")}</legend>";
                            dr_tmp = db.DB_GetFirstDataByDataRow($"select *,(NeedQTY+SafeQTY) as QTY from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartSN=0");
                            if (dr_tmp != null)
                            {
                                string delbn = "";
                                string nqty = dr_tmp["SafeQTY"].ToString();
                                tmp_s = dr_tmp["QTY"].ToString();
                                re = $"{re}<p>計畫編號:{needId} 品名規格:{dr_Material["PartName"].ToString()}  {dr_Material["Specification"].ToString()}   {delbn}</p><p>需求量:{dr_APS_NeedData["NeedQTY"].ToString()} + 補安全量:{dr_tmp["SafeQTY"].ToString()}";
                                re = $"{re} = 總投產量:{tmp_s}";
                                if (dr_APS_NeedData["State"].ToString() == "9")
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr_tmp["SimulationId"].ToString()}'");
                                    if (dr_tmp != null)
                                    {
                                        if (int.Parse(dr_tmp["Next_StoreQTY"].ToString()) != 0)
                                        { re = $"{re}&nbsp;&nbsp;實際總入庫量:{dr_tmp["Next_StoreQTY"].ToString()}"; }
                                        else { re = $"{re}&nbsp;&nbsp;目前總入庫量:{dr_tmp["Detail_QTY"].ToString()}"; }
                                    }
                                }
                                re = $"{re}</p>";
                            }
                            //取得第一層結構
                            tmp_s = "";
                            //dr_tmp = db.DB_GetFirstDataByDataRow($"select NeedId from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{needId}' and ActionType=''");
                            //if (dr_tmp != null) { tmp_s = " open"; }
                            dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartSN>=0 order by PartSN desc");
                            if (dr_tmp != null) { maxPartSN = int.Parse(dr_tmp["PartSN"].ToString()); }
                            DataTable dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartSN>=0 order by PartSN");
                            if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
                            {
                                string show = "label";
                                DataRow mast = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartSN=0");
                                if (db.DB_GetQueryCount($"SELECT SimulationId FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{mast["SimulationId"].ToString()}' and ActionType=''") > 0) { show = "labelRed"; }
                                s_date = mast.IsNull("StartDate") ? "" : Convert.ToDateTime(mast["StartDate"]).ToString("yyyy-MM-dd HH:mm");
                                string qty = "";
                                if (!mast.IsNull("Source_StationNO") && dr_APS_NeedData["State"].ToString() == "6")
                                {
                                    DataRow dr_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{mast["SimulationId"].ToString()}'");
                                    if (dr_PartNOTimeNote != null)
                                    { qty = $"<br />{mast["DOCNumberNO"].ToString()}&nbsp;&nbsp;已產:{dr_PartNOTimeNote["Detail_QTY"].ToString()}"; }
                                }
                                string otherINFO = "";
                                if (!mast.IsNull("StationNO_Merge") && mast["StationNO_Merge"].ToString() != "") { otherINFO = $"合併站:{mast["StationNO_Merge"].ToString()}<br />"; }
                                re = $"{re}<details{tmp_s}><summary>生產流程樹狀圖</summary><div id='wrapper'><button id='{mast["SimulationId"].ToString()}' class='{show}' style='border: 2px blue none;' disabled onclick=\"_me.onShow('{mast["SimulationId"].ToString()}')\">{mast["Master_PartNO"].ToString()} [{mast["Source_StationNO"].ToString()}]<br />{otherINFO}開始日:{s_date}<br />完成日:{Convert.ToDateTime(mast["SimulationDate"]).ToString("yyyy-MM-dd HH:mm")}{qty}</button>\n";
                                if (dt_Simulation.Rows.Count > 1)
                                {
                                    dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartSN=1 order by PartSN");
                                    DataRow[] dr_WorkTimeNote = dt_Simulation.Select($"PartSN=2 and Master_PartNO='{dt_Simulation.Rows[0]["PartNO"].ToString()}' and Apply_StationNO='{dt_Simulation.Rows[0]["Source_StationNO"].ToString()}'");
                                    RecursiveTree(db, needId, 1, (dr_WorkTimeNote != null ? true : false), dt_Simulation, dr_APS_NeedData["State"].ToString(), ref re, false);
                                }
                                re = $"{re}</div></details></fieldset></div><br />\n";
                            }
                            #endregion

                            #region APS_PartNOTimeNote明細
                            re = $"{re}<p>生產流程實際明細</p><div>";
                            re = $"{re}<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                            re = $"{re}<tr><td>執行碼</td><td>計畫日期</td><td>工站</td><td>工序</td><td>單據編號</td><td>料號/品名/規格</td><td>已產量</td><td>不良</td><td>移轉量</td><td>移轉站</td><td>進倉單據</td><td>進倉量</td></tr>";
                            dt2 = db.DB_GetData($@"SELECT b.NeedQTY as P_NeedQTY,b.NoStation,b.DOCNumberNO,b.Detail_QTY,b.Detail_Fail_QTY,b.CalendarDate,b.[Next_APS_StationNO],b.[Next_StationQTY],b.[Store_DOCNumberNO],b.[Next_StoreQTY],a.*  FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                                    join SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as b on a.NeedId=b.NeedId and a.SimulationId=b.SimulationId
                                                    where a.ServerId='{_Fun.Config.ServerId}' and a.NeedId='{needId}' order by a.PartSN desc,b.SimulationId");
                            DataRow tmp = null;
                            string sID = "";
                            List<string> orderNO = new List<string>();
                            if (dt2 != null && dt2.Rows.Count > 0)
                            {
                                string rStationNO = "";
                                string rPartNO = "";
                                foreach (DataRow dr in dt2.Rows)
                                {
                                    if (dr["DOCNumberNO"].ToString() != "" && dr["DOCNumberNO"].ToString().Substring(0, 2) == "XX")
                                    {
                                        if (!orderNO.Contains(dr["DOCNumberNO"].ToString())) { orderNO.Add(dr["DOCNumberNO"].ToString()); }
                                    }
                                    if (sID == "") { sID = $"'{dr["SimulationId"].ToString()}'"; }
                                    else { sID = $"{sID},'{dr["SimulationId"].ToString()}'"; }
                                    if (!bool.Parse(dr["NoStation"].ToString()))
                                    {
                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT StationNO,StationName from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                                        if (tmp != null) { rStationNO = $"{tmp["StationNO"].ToString()}&nbsp;{tmp["StationName"].ToString()}"; } else { rStationNO = ""; }
                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT PartNO,PartName,Specification from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                                        if (tmp != null) { rPartNO = $"{tmp["PartNO"].ToString()}&nbsp;{tmp["PartName"].ToString()}&nbsp;{tmp["Specification"].ToString()}"; } else { rPartNO = ""; }
                                        re = $"{re}<tr><td>{dr["SimulationId"].ToString()}</td><td>{Convert.ToDateTime(dr["CalendarDate"]).ToString("yy-MM-dd HH:mm")}</td><td>{rStationNO}</td><td>{dr["Source_StationNO_IndexSN"].ToString()}{dr["Source_StationNO_Custom_DisplayName"].ToString()}</td><td>{dr["DOCNumberNO"].ToString()}</td><td>{rPartNO}</td><td>{dr["Detail_QTY"].ToString()}</td><td>{dr["Detail_Fail_QTY"].ToString()}</td><td>{dr["Next_StationQTY"].ToString()}</td><td>{dr["Next_APS_StationNO"].ToString()}</td><td>{dr["Store_DOCNumberNO"].ToString()}</td><td>{dr["Next_StoreQTY"].ToString()}</td></tr>";
                                    }
                                }
                            }
                            re = $"{re}</table>";
                            #endregion

                            #region 單據明細
                            if (sID != "")
                            {
                                DataTable dt_DOC = null;
                                string StartTime = "";
                                dt2 = db.DB_GetData($@"SELECT b.PartName,b.Specification,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[WeightQty],a.[Remark],a.[SimulationId],a.[IsOK],a.StoreNO as [IN_StoreNO],'' as OUT_StoreNO,a.StoreSpacesNO,a.[EndTime],a.[StartTime],a.[CT]  FROM SoftNetMainDB.[dbo].[DOC1BuyII] as a
                                                    join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO
                                                    where a.ServerId='{_Fun.Config.ServerId}' and a.SimulationId in ({sID})");
                                if (dt2 != null && dt2.Rows.Count > 0)
                                { if (dt_DOC == null) { dt_DOC = dt2; } else { dt_DOC.Merge(dt2); } }
                                dt2 = db.DB_GetData($@"SELECT b.PartName,b.Specification,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[WeightQty],a.[Remark],a.[SimulationId],a.[IsOK],a.StoreNO as [OUT_StoreNO],'' as IN_StoreNO,a.StoreSpacesNO,a.[EndTime],a.[StartTime],a.[CT]  FROM SoftNetMainDB.[dbo].[DOC2SalesII] as a
                                                    join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO
                                                    where a.ServerId='{_Fun.Config.ServerId}' and a.SimulationId in ({sID})");
                                if (dt2 != null && dt2.Rows.Count > 0)
                                { if (dt_DOC == null) { dt_DOC = dt2; } else { dt_DOC.Merge(dt2); } }
                                dt2 = db.DB_GetData($@"SELECT b.PartName,b.Specification,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[WeightQty],a.[Remark],a.[SimulationId],a.[IsOK],a.IN_StoreNO,a.OUT_StoreNO,a.IN_StoreSpacesNO as [StoreSpacesNO],a.[EndTime],a.[StartTime],a.[CT]  FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                    join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO
                                                    where a.ServerId='{_Fun.Config.ServerId}' and IN_StoreNO!='' and a.SimulationId in ({sID})");
                                if (dt2 != null && dt2.Rows.Count > 0)
                                { if (dt_DOC == null) { dt_DOC = dt2; } else { dt_DOC.Merge(dt2); } }
                                dt2 = db.DB_GetData($@"SELECT b.PartName,b.Specification,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[WeightQty],a.[Remark],a.[SimulationId],a.[IsOK],a.IN_StoreNO,a.OUT_StoreNO,a.OUT_StoreSpacesNO as [StoreSpacesNO],a.[EndTime],a.[StartTime],a.[CT]  FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                    join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO
                                                    where a.ServerId='{_Fun.Config.ServerId}' and OUT_StoreNO!='' and a.SimulationId in ({sID})");
                                if (dt2 != null && dt2.Rows.Count > 0)
                                { if (dt_DOC == null) { dt_DOC = dt2; } else { dt_DOC.Merge(dt2); } }
                                dt2 = db.DB_GetData($@"SELECT b.PartName,b.Specification,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[WeightQty],a.[Remark],a.[SimulationId],a.[IsOK],a.StoreNO as [IN_StoreNO],'' as OUT_StoreNO,a.StoreSpacesNO,a.[EndTime],a.[StartTime],a.[CT]  FROM SoftNetMainDB.[dbo].[DOC4ProductionII] as a
                                                    join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO
                                                    where a.ServerId='{_Fun.Config.ServerId}' and a.SimulationId in ({sID})");
                                if (dt2 != null && dt2.Rows.Count > 0)
                                { if (dt_DOC == null) { dt_DOC = dt2; } else { dt_DOC.Merge(dt2); } }
                                dt2 = db.DB_GetData($@"SELECT b.PartName,b.Specification,a.[Id],a.[DOCNumberNO],a.[ArrivalDate],a.[ServerId],a.[PartNO],a.[Price],a.[Unit],a.[QTY],a.[WeightQty],a.[Remark],a.[SimulationId],a.[IsOK],a.StoreNO as [OUT_StoreNO],'' as IN_StoreNO,a.StoreSpacesNO,a.[EndTime],a.[StartTime],a.[CT]  FROM SoftNetMainDB.[dbo].[DOC5OUTII] as a
                                                    join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO
                                                    where a.ServerId='{_Fun.Config.ServerId}' and a.SimulationId in ({sID})");
                                if (dt2 != null && dt2.Rows.Count > 0)
                                { if (dt_DOC == null) { dt_DOC = dt2; } else { dt_DOC.Merge(dt2); } }
                                if (dt_DOC != null && dt_DOC.Rows.Count > 0)
                                {
                                    if (key.SelectFun6 != null && key.SelectFun6 == "F2")
                                    { dt_DOC.DefaultView.Sort = "SimulationId,StartTime,DOCNumberNO,PartNO"; }
                                    else { dt_DOC.DefaultView.Sort = "StartTime,DOCNumberNO,PartNO"; }
                                    dt_DOC = dt_DOC.DefaultView.ToTable();
                                    re = $"{re}<p>生產單據明細</p><div>";
                                    re = $"{re}<table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'>";
                                    re = $"{re}<tr><td>執行碼</td><td>計畫日期</td><td>開始日期</td><td>單據編號</td><td>料號/品名/規格</td><td>備註</td><td>入庫編號</td><td>出庫編號</td><td>儲位</td><td>單位</td><td>數量</td><td>狀態</td></tr>";

                                    string rdocName = "";
                                    string rIN_StoreNO = "";
                                    string rOUT_StoreNO = "";
                                    string rPartNO = "";
                                    DataRow tmp02 = null;
                                    foreach (string s in orderNO)
                                    {
                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{s}'");
                                        if (tmp != null)
                                        {
                                            tmp02 = db.DB_GetFirstDataByDataRow($"SELECT DOCName from SoftNetMainDB.[dbo].[DOCRole] where ServerId='{_Fun.Config.ServerId}' and DOCNO='{s.Substring(0, 4)}'");
                                            if (tmp02 != null) { rdocName = $"{tmp02["DOCName"].ToString()}"; } else { rdocName = ""; }
                                            re = $"{re}<tr><td>{tmp["NeedId"].ToString()}</td><td></td><td>{Convert.ToDateTime(tmp["StartTime"]).ToString("yy-MM-dd HH:mm")}</td><td>{tmp["OrderNO"].ToString()}{rdocName}</td><td>{tmp["PartNO"].ToString()}&nbsp;{tmp["PartName"].ToString()}</td><td></td><td></td><td></td><td></td><td></td><td>{tmp["Quantity"].ToString()}</td><td></td></tr>";
                                        }
                                    }
                                    foreach (DataRow dr in dt_DOC.Rows)
                                    {
                                        rIN_StoreNO = "";
                                        rOUT_StoreNO = "";
                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT DOCName from SoftNetMainDB.[dbo].[DOCRole] where ServerId='{_Fun.Config.ServerId}' and DOCNO='{dr["DOCNumberNO"].ToString().Substring(0, 4)}'");
                                        if (tmp != null) { rdocName = $"{tmp["DOCName"].ToString()}"; } else { rdocName = ""; }
                                        if (dr["IN_StoreNO"].ToString() != "")
                                        {
                                            tmp = db.DB_GetFirstDataByDataRow($"SELECT StoreNO,StoreName from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr["IN_StoreNO"].ToString()}'");
                                            if (tmp != null) { rIN_StoreNO = $"{tmp["StoreNO"].ToString()}&nbsp;{tmp["StoreName"].ToString()}"; }
                                        }
                                        if (dr["OUT_StoreNO"].ToString() != "")
                                        {
                                            tmp = db.DB_GetFirstDataByDataRow($"SELECT StoreNO,StoreName from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{dr["OUT_StoreNO"].ToString()}'");
                                            if (tmp != null) { rOUT_StoreNO = $"{tmp["StoreNO"].ToString()}&nbsp;{tmp["StoreName"].ToString()}"; }
                                        }
                                        rPartNO = $"{dr["PartNO"].ToString()}&nbsp;{dr["PartName"].ToString()}&nbsp;{dr["Specification"].ToString()}";
                                        if (!dr.IsNull("StartTime") && dr["ArrivalDate"].ToString() != "") { StartTime = Convert.ToDateTime(dr["StartTime"]).ToString("yy-MM-dd HH:mm"); }
                                        else { StartTime = ""; }
                                        re = $"{re}<tr><td>{dr["SimulationId"].ToString()}</td><td>{Convert.ToDateTime(dr["ArrivalDate"]).ToString("yy-MM-dd HH:mm")}</td><td>{StartTime}</td><td>{dr["DOCNumberNO"].ToString()}{rdocName}</td><td>{rPartNO}</td><td>{dr["Remark"].ToString()}</td><td>{rIN_StoreNO}</td><td>{rOUT_StoreNO}</td><td>{dr["StoreSpacesNO"].ToString()}</td><td>{dr["Unit"].ToString()}</td><td>{dr["QTY"].ToString()}</td><td>{dr["IsOK"].ToString()}</td></tr>";
                                    }
                                    re = $"{re}</table>";
                                    re = $"{re}<div>";
                                }
                            }
                            #endregion
                        }
                    }
                    else { re = $"<div>目前條件中,無生產歷程的計畫編號</div>"; }
                }
                else { re = $"<div>目前條件中,無生產歷程的計畫編號</div>"; }
            }
            db.Dispose();
            ViewBag.HtmlOutput = re;
            return View(key);
        }

        private void old_selectSIDList(DBADO db, ref APSViewData key, string state = "")
        {
            if (key == null) { key = new APSViewData(); }
            List<string[]> list = new List<string[]>();
            if (state == "") { state = "a.State='9'"; }
            else if (state == "69") { state = "(a.State='6' or a.State='9')"; }
            DataTable dt_APS_NeedData = db.DB_GetData($@"select a.*,b.PartName,b.Specification,b.Class from SoftNetSYSDB.[dbo].[APS_NeedData] as a
                                                        join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                                                        where a.ServerId='{_Fun.Config.ServerId}' and {state} order by a.PartNO");
            if (dt_APS_NeedData != null && dt_APS_NeedData.Rows.Count > 0)
            {
                foreach (DataRow d in dt_APS_NeedData.Rows)
                {
                    list.Add(new string[] { d["State"].ToString(), d["NeedType"].ToString(), d["Id"].ToString(), d["PartNO"].ToString(), d["PartName"].ToString(), d["Specification"].ToString(), d["NeedSimulationDate"].ToString(), d["NeedQTY"].ToString(), d["Class"].ToString(), d["UpdateTime"].ToString() });
                }
            }
            key.NeedIdList = list;
        }
        private void selectSIDList(DBADO db,ref APSViewData key)
        {
            if (key == null) { key = new APSViewData(); }
            List<string[]> list = new List<string[]>();
            //DataTable dt_APS_NeedData = db.DB_GetData($"select a.*,b.PartName,b.Specification,b.Class from SoftNetSYSDB.[dbo].[APS_NeedData] as a,SoftNetMainDB.[dbo].[Material] as b where a.ServerId='{_Fun.Config.ServerId}' and b.ServerId='{_Fun.Config.ServerId}' and a.State!='0' and a.State!='1' and a.State!='9' and a.PartNO=b.PartNO order by a.PartNO");
            DataTable dt_APS_NeedData = db.DB_GetData($@"select a.*,b.PartName,b.Specification,b.Class from SoftNetSYSDB.[dbo].[APS_NeedData] as a
                                                        join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO and b.ServerId='{_Fun.Config.ServerId}'
                                                        where a.ServerId='{_Fun.Config.ServerId}' and a.State!='0' and a.State!='1' and a.State!='9' order by a.PartNO");
            if (dt_APS_NeedData != null && dt_APS_NeedData.Rows.Count > 0)
            {
                foreach (DataRow d in dt_APS_NeedData.Rows)
                {
                    list.Add(new string[] { d["State"].ToString(), d["NeedType"].ToString(), d["Id"].ToString(), d["PartNO"].ToString(), d["PartName"].ToString(), d["Specification"].ToString(), d["NeedSimulationDate"].ToString(), d["NeedQTY"].ToString(), d["Class"].ToString(), d["UpdateTime"].ToString() });
                }
            }
            key.NeedIdList = list;
        }

        public IActionResult OLDALLSimulation(APSViewData key)
        {
            string re = "";
            DBADO db = new DBADO("1", _Fun.Config.Db);
            if (key == null) { key = new APSViewData(); }
            old_selectSIDList(db, ref key);
            if (key.SelectFun1 == "") { return View(key); }

            if (key.NeedIdList.Count <= 0)
            {
                re = $"<div>目前線上沒有結案計畫編號</div>";
            }
            else
            {
                DataRow dr_tmp = null;
                string sType1 = "";
                string sType2 = "";
                string tmp_s = "";
                string s_date = "";
                foreach (string[] dr0 in key.NeedIdList)
                {
                    if (key.SelectNeedId != null && key.SelectNeedId.Trim() != "" && key.SelectNeedId != dr0[2]) { continue; }
                    else
                    {
                        if (key.SelectNeedId == null || key.SelectNeedId.Trim() == "")
                        {
                            if (key.SelectFun2 != "A3")
                            {
                                if (key.SelectFun2 == "A1" && dr0[8] != "4") { continue; }
                                if (key.SelectFun2 == "A2" && dr0[8] != "5") { continue; }
                            }
                            if (key.SelectFun3 != "B6")
                            {
                                DateTime time = Convert.ToDateTime(dr0[9]);
                                DateTime now1 = DateTime.Now.AddMonths(-1);
                                DateTime now2 = DateTime.Now.AddMonths(-2);
                                if (key.SelectFun3 == "B1" && DateTime.Now.AddMonths(-1) > time) { continue; }
                                if (key.SelectFun3 == "B2" && DateTime.Now.AddMonths(-2) > time) { continue; }
                                if (key.SelectFun3 == "B3" && DateTime.Now.AddMonths(-3) > time) { continue; }
                                if (key.SelectFun3 == "B4" && DateTime.Now.AddMonths(-4) > time) { continue; }
                                if (key.SelectFun3 == "B5" && DateTime.Now.AddYears(-1) > time) { continue; }
                            }
                        }
                    }
                    if (dr0[0] == "2") { sType1 = "模擬完成"; }
                    else if (dr0[0] == "3") { sType1 = $"模擬取消"; }
                    else if (dr0[0] == "6") { sType1 = $"已轉計畫"; }
                    else { sType1 = "結案"; }

                    if (dr0[1] == "1") { sType2 = $"訂單需求"; }
                    else if (dr0[1] == "2") { sType2 = $"客戶需求"; }
                    else if (dr0[1] == "5") { sType2 = $"底稿發出"; }
                    else { sType2 = $"廠內需求"; }
                    switch (dr0[0])
                    {
                        case "9":
                            {
                                tmp_s = "";
                                if (dr0[6].Trim() != "")
                                { re = $"{re}\n<div class='shadowbox'><fieldset id='{dr0[2]}'><legend>料號:{dr0[3]} {sType2} {sType1} 預計入庫日:{Convert.ToDateTime(dr0[6]).ToString("yyyy-MM-dd HH:mm")}</legend>"; }
                                else
                                { re = $"{re}\n<div class='shadowbox'><fieldset id='{dr0[2]}'><legend>料號:{dr0[3]} {sType2} {sType1}</legend>"; }
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select *,(NeedQTY+SafeQTY) as QTY from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}' and PartSN=0");
                                if (dr_tmp != null)
                                {
                                    string delbn = "";
                                    string nqty = dr_tmp["SafeQTY"].ToString();
                                    tmp_s = dr_tmp["QTY"].ToString();
                                    re = $"{re}<p>計畫編號:{dr0[2]} 品名規格:{dr0[4]}  {dr0[5]}   {delbn}</p><p>需求量:{dr0[7]} + 補安全量:{dr_tmp["SafeQTY"].ToString()}";
                                    re = $"{re} = 總投產量:{tmp_s}";
                                    re = $"{re}</p>";
                                }
                                //取得第一層結構
                                tmp_s = "";
                                //dr_tmp = db.DB_GetFirstDataByDataRow($"select NeedId from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{dr0[2]}' and ActionType=''");
                                //if (dr_tmp != null) { tmp_s = " open"; }
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}' and PartSN>=0 order by PartSN desc");
                                if (dr_tmp != null) { maxPartSN = int.Parse(dr_tmp["PartSN"].ToString()); }
                                DataTable dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}' and PartSN>=0 order by PartSN");
                                if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
                                {
                                    string show = "label";
                                    DataRow mast = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}' and PartSN=0");
                                    if (db.DB_GetQueryCount($"SELECT SimulationId FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{mast["SimulationId"].ToString()}' and ActionType=''") > 0) { show = "labelRed"; }
                                    s_date = mast.IsNull("StartDate") ? "" : Convert.ToDateTime(mast["StartDate"]).ToString("yyyy-MM-dd HH:mm");
                                    string qty = "";
                                    if (!mast.IsNull("Source_StationNO") && dr0[0] == "6")
                                    {
                                        DataRow dr_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{mast["SimulationId"].ToString()}'");
                                        if (dr_PartNOTimeNote != null)
                                        { qty = $"<br />{mast["DOCNumberNO"].ToString()}&nbsp;&nbsp;已產:{dr_PartNOTimeNote["Detail_QTY"].ToString()}"; }
                                    }
                                    string otherINFO = "";
                                    if (!mast.IsNull("StationNO_Merge") && mast["StationNO_Merge"].ToString() != "") { otherINFO = $"合併站:{mast["StationNO_Merge"].ToString()}<br />"; }
                                    re = $"{re}<details{tmp_s}><summary>生產流程樹狀圖</summary><div id='wrapper'><button id='{mast["SimulationId"].ToString()}' class='{show}' style='border: 2px blue none;' disabled onclick=\"_me.onShow('{mast["SimulationId"].ToString()}')\">{mast["Master_PartNO"].ToString()} [{mast["Source_StationNO"].ToString()}]<br />{otherINFO}開始日:{s_date}<br />完成日:{Convert.ToDateTime(mast["SimulationDate"]).ToString("yyyy-MM-dd HH:mm")}{qty}</button>\n";
                                    if (dt_Simulation.Rows.Count > 1)
                                    {
                                        dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}' and PartSN=1 order by PartSN");
                                        DataRow[] dr_WorkTimeNote = dt_Simulation.Select($"PartSN=2 and Master_PartNO='{dt_Simulation.Rows[0]["PartNO"].ToString()}' and Apply_StationNO='{dt_Simulation.Rows[0]["Source_StationNO"].ToString()}'");
                                        RecursiveTree(db, dr0[2], 1, (dr_WorkTimeNote != null ? true : false), dt_Simulation, dr0[0], ref re,false);
                                    }
                                    re = $"{re}</div></details></fieldset></div><br />\n";
                                }
                            }
                            break;
                    }
                }
            }
            db.Dispose();
            ViewBag.HtmlOutput = re;
            return View(key);
        }

        public IActionResult ALLSimulation(APSViewData key)
        {
            string re = "";
            DBADO db = new DBADO("1", _Fun.Config.Db);
            if (key == null) { key = new APSViewData(); }
            selectSIDList(db, ref key);
            if (key.SelectFun1 == "") { return View(key); }

            if (key.NeedIdList.Count <= 0)
            {
                re = $"<div>目前線上沒有計畫編號</div>";
            }
            else
            {
                DataRow dr_tmp = null;
                string sType1 = "";
                string sType2 = "";
                string tmp_s = "";
                string s_date = "";
                foreach (string[] dr0 in key.NeedIdList)
                {
                    if (key.SelectNeedId!=null && key.SelectNeedId.Trim() != "" && key.SelectNeedId!= dr0[2]) { continue; }
                    else
                    {
                        if (key.SelectNeedId == null || key.SelectNeedId.Trim() == "")
                        {
                            if (key.SelectFun2 != "A3")
                            {
                                if (key.SelectFun2 == "A1" && dr0[8] != "4") { continue; }
                                if (key.SelectFun2 == "A2" && dr0[8] != "5") { continue; }
                            }
                            if (key.SelectFun3 != "B6")
                            {
                                DateTime time = Convert.ToDateTime(dr0[9]);
                                DateTime now1 = DateTime.Now.AddMonths(-1);
                                DateTime now2 = DateTime.Now.AddMonths(-2);
                                if (key.SelectFun3 == "B1" && DateTime.Now.AddMonths(-1) > time) { continue; }
                                if (key.SelectFun3 == "B2" && DateTime.Now.AddMonths(-2) > time) { continue; }
                                if (key.SelectFun3 == "B3" && DateTime.Now.AddMonths(-3) > time) { continue; }
                                if (key.SelectFun3 == "B4" && DateTime.Now.AddMonths(-4) > time) { continue; }
                                if (key.SelectFun3 == "B5" && DateTime.Now.AddYears(-1) > time) { continue; }
                            }
                        }
                    }
                    if (dr0[0] == "2") { sType1 = "模擬完成"; }
                    else if (dr0[0] == "3") { sType1 = $"模擬取消"; }
                    else { sType1 = "已轉計畫"; }
                    if (dr0[1] == "1") { sType2 = $"訂單需求"; }
                    else if (dr0[1] == "2") { sType2 = $"客戶需求"; }
                    else if (dr0[1] == "5") { sType2 = $"底稿發出"; }
                    else { sType2 = $"廠內需求"; }
                    switch (dr0[0])
                    {
                        case "2"://模擬完成
                        case "6"://已轉生產
                            {
                                tmp_s = "";
                                re = $"{re}\n<div class='shadowbox'><fieldset id='{dr0[2]}'><legend>料號:{dr0[3]} {sType2} {sType1} 預計入庫日:{Convert.ToDateTime(dr0[6]).ToString("yyyy-MM-dd HH:mm")}</legend>";
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select *,(NeedQTY+SafeQTY) as QTY from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}' and PartSN=0");
                                if (dr_tmp != null)
                                {
                                    string delbn = "";
                                    if (dr0[0] == "6") { delbn = $"<button style='border: 2px blue none;background-color: #CDCD9A;' onclick='_me.deleteSimulation(\"{dr0[2]}\")'>刪除計畫</button>"; }
                                    string nqty = dr_tmp["SafeQTY"].ToString();
                                    tmp_s = dr_tmp["QTY"].ToString();
                                    re = $"{re}<p>計畫編號:{dr0[2]} 品名規格:{dr0[4]}  {dr0[5]}   {delbn}</p><p>需求量:{dr0[7]} + 補安全量:{dr_tmp["SafeQTY"].ToString()} - 倉庫可用:";
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select (NeedQTY+SafeQTY) as QTY from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}' and PartSN<0");
                                    if (dr_tmp != null) { re = $"{re}{dr_tmp["QTY"].ToString()}"; } else { re = $"{re}0"; }

                                    re = $"{re} = 總投產量:{tmp_s}";
                                    if (dr0[0] == "6") 
                                    { 
                                        re = $"{re}   <button style='border: 2px blue none;background-color: #CDCD9A;' onclick='_me.changNeedQTY(\"{dr0[2]},{dr0[7]},{nqty}\")'>調整投產數量</button>"; 
                                    }
                                    re = $"{re}</p>";
                                }
                                //取得第一層結構
                                tmp_s = "";
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select NeedId from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{dr0[2]}' and ActionType=''");
                                if (dr_tmp != null) { tmp_s = " open"; }
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}' and PartSN>=0 order by PartSN desc");
                                if (dr_tmp != null) { maxPartSN = int.Parse(dr_tmp["PartSN"].ToString()); }
                                DataTable dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}' and PartSN>=0 order by PartSN");
                                if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
                                {
                                    string show = "label";
                                    DataRow mast = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}' and PartSN=0");
                                    if (db.DB_GetQueryCount($"SELECT SimulationId FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{mast["SimulationId"].ToString()}' and ActionType=''") > 0) { show = "labelRed"; }
                                    s_date = mast.IsNull("StartDate") ? "" : Convert.ToDateTime(mast["StartDate"]).ToString("yyyy-MM-dd HH:mm");
                                    string qty = "";
                                    if (!mast.IsNull("Source_StationNO") && dr0[0] == "6")
                                    {
                                        DataRow dr_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{mast["SimulationId"].ToString()}'");
                                        if (dr_PartNOTimeNote != null)
                                        { qty = $"<br />{mast["DOCNumberNO"].ToString()}&nbsp;&nbsp;已產:{dr_PartNOTimeNote["Detail_QTY"].ToString()}"; }
                                    }
                                    string otherINFO = "";
                                    if (!mast.IsNull("StationNO_Merge") && mast["StationNO_Merge"].ToString() != "") { otherINFO = $"合併站:{mast["StationNO_Merge"].ToString()}<br />"; }
                                    re = $"{re}<details{tmp_s}><summary>生產流程樹狀圖</summary><div id='wrapper'><button id='{mast["SimulationId"].ToString()}' class='{show}' style='border: 2px blue none;' onclick=\"_me.onShow('{mast["SimulationId"].ToString()}')\">{mast["Master_PartNO"].ToString()} [{mast["Source_StationNO"].ToString()}]<br />{otherINFO}開始日:{s_date}<br />完成日:{Convert.ToDateTime(mast["SimulationDate"]).ToString("yyyy-MM-dd HH:mm")}{qty}</button>\n";
                                    if (dt_Simulation.Rows.Count > 1)
                                    {
                                        dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}' and PartSN=1 order by PartSN");
                                        DataRow[] dr_WorkTimeNote = dt_Simulation.Select($"PartSN=2 and Master_PartNO='{dt_Simulation.Rows[0]["PartNO"].ToString()}' and Apply_StationNO='{dt_Simulation.Rows[0]["Source_StationNO"].ToString()}'");
                                        RecursiveTree(db, dr0[2], 1, (dr_WorkTimeNote != null ? true : false), dt_Simulation, dr0[0], ref re);
                                    }
                                    re = $"{re}</div></details></fieldset></div><br />\n";
                                }
                                else 
                                { 
                                    re = $"{re}</div></fieldset></div><br />\n";
                                    db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='9' where Id='{dr0[2]}'");
                                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr0[2]}'");
                                    db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{dr0[2]}'");
                                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{dr0[2]}'");
                                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{dr0[2]}'");
                                }
                            }
                            break;
                        case "3"://模擬逾時
                            {
                                re = $"{re}\n<div class='shadowbox'><fieldset id='{dr0[2]}'><legend>料號:{dr0[3]} {sType2} {sType1} 模擬已過時,需重新模擬</legend>";
                                re = $"{re}</fieldset></div><br />\n";
                            }
                            break;
                        case "4"://庫存足夠
                        case "7"://庫存足夠
                        case "8"://庫存足夠
                            {
                                re = $"{re}\n<div class='shadowbox'><fieldset id='{dr0[2]}'><legend>料號:{dr0[3]} {sType2} {sType1} 預計入庫日:{Convert.ToDateTime(dr0[6]).ToString("yyyy-MM-dd HH:mm")}</legend>";
                                re = $"{re}</fieldset></div><br />\n";
                            }
                            break;
                    }
                }
            }
            db.Dispose();
            ViewBag.HtmlOutput = re;
            return View(key);
        }
        private void RecursiveSimulation(DBADO db, string needId, string sn, string sn2, DataTable dr_M, ref string re)
        {
            if (sn != sn2) { re = $"{re}<div class='branch lv{sn2}'>\n"; sn = sn2; }
            string show = "label";
            bool byPass = false;
            string s_date = "";
            #region code還是有問題 只加工無料
            if (dr_M != null && dr_M.Rows.Count > 0 && dr_M.Rows[0]["Master_PartNO"].ToString() == dr_M.Rows[0]["PartNO"].ToString())
            {
                s_date = dr_M.Rows[0].IsNull("StartDate") ? "" : Convert.ToDateTime(dr_M.Rows[0]["StartDate"]).ToString("yyyy-MM-dd HH:mm");
                if (db.DB_GetQueryCount($"SELECT SimulationId FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{dr_M.Rows[0]["SimulationId"].ToString()}'") > 0) { show = "labelRed"; }
                else { show = "label"; }
                re = $"{re}<div class='entry'><button id ='{dr_M.Rows[0]["SimulationId"].ToString()}' class='{show}' onclick=\"_me.onShow('{dr_M.Rows[0]["SimulationId"].ToString()}')\">{dr_M.Rows[0]["Apply_StationNO"].ToString()}<br />開始日:{s_date}<br />計劃日:{Convert.ToDateTime(dr_M.Rows[0]["SimulationDate"]).ToString("yyyy-MM-dd HH:mm")}</button></div>\n";

                byPass = true;
                //###???還是有問題
            }
            #endregion
            foreach (DataRow dr1 in dr_M.Rows)
            {
                //if (byPass) { byPass = false;continue; }
                if (db.DB_GetQueryCount($"SELECT SimulationId FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{dr1["SimulationId"].ToString()}' and ActionType=''") > 0) { show = "labelRed"; }
                else { show = "label"; }
                if (!dr1.IsNull("Source_StationNO") && dr1["Source_StationNO"].ToString().Trim() != "")
                {
                    DataTable dr_MII = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and Master_PartNO='{dr1["PartNO"].ToString()}' and Apply_StationNO='{dr1["Source_StationNO"].ToString()}' order by PartSN, PartSN_Sub desc");
                    if (dr_MII != null && dr_MII.Rows.Count > 0)
                    {
                        s_date = dr1.IsNull("StartDate") ? "" : Convert.ToDateTime(dr1["StartDate"]).ToString("yyyy-MM-dd HH:mm");
                        re = $"{re}<div class='entry'><button id='{dr1["SimulationId"].ToString()}' class='{show}' onclick=\"_me.onShow('{dr1["SimulationId"].ToString()}')\">{dr1["PartNO"].ToString()} [{dr1["Source_StationNO"].ToString()}]<br />開始日:{s_date}<br />計劃日:{Convert.ToDateTime(dr1["SimulationDate"]).ToString("yyyy-MM-dd HH:mm")}</button>\n";
                        //RecursiveSimulation(db, needId, sn2, dr_MII.Rows[0]["PartSN"].ToString().Trim(), dr_MII, ref re);
                        RecursiveSimulation(db, needId, sn2, (int.Parse(sn2)+1).ToString(), dr_MII, ref re);

                        re = $"{re}</div>\n";
                    }
                    else
                    {
                        re = $"{re}<div class='entry'><button id='{dr1["SimulationId"].ToString()}' class='{show}' onclick=\"_me.onShow('{dr1["SimulationId"].ToString()}')\">{dr1["PartNO"].ToString()}</button></div>\n";
                    }
                }else
                {
                    re = $"{re}<div class='entry'><button id='{dr1["SimulationId"].ToString()}' class='{show}' onclick=\"_me.onShow('{dr1["SimulationId"].ToString()}')\">{dr1["PartNO"].ToString()}</button></div>\n";
                }
            }
            re = $"{re}</div>\n";
        }









        public ActionResult DOCBuyRead()
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

        private APSService EditService()
        {
            return new APSService(Ctrl);
        }

        [HttpPost]
        public string CraeteDOCBuy(string keys) //採購確認  0=ipport,1.需求碼1,需求碼2....
        {
            //###???暫時寫死單別與廠商 AA02,單價
            string meg = "";
            string[] data = keys.Split(',');
            string id = "";

            #region 拆參數碼
            for (int i = 1; i < data.Length; i++)
            {
                if (id == "") { id = $"'{data[i]}'"; }
                else { id = $"{id},'{data[i]}'"; }
            }
            #endregion

            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                Dictionary<string, List<string[]>> _poolTEST_NO = new Dictionary<string, List<string[]>>();//暫存檢驗單
                Dictionary<string, List<string[]>> _poolStore_NO = new Dictionary<string, List<string[]>>();//暫存入庫單

                int no = 0;
                #region 取採購單號流水號
                //###???暫時寫死單別AA02 與寫死庫別 a1
                DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT DOCNumberNO FROM SoftNetMainDB.[dbo].[DOC1Buy] where ServerId='{_Fun.Config.ServerId}' and DOCNO='AA02' order by DOCNumberNO desc");
                if (dr != null)
                {
                    no = int.Parse(dr["DOCNumberNO"].ToString().Trim().Substring(12));
                }
                #endregion

                string sql = "";
                List<string[]> buys = new List<string[]>(); ///0="PartNO 1=CalendarDate 2=qty 3=Class 4=SimulationId 5=上筆關聯Id
                string tmp_Id = "";
                string tmp_mfno = "";
                string date = DateTime.Now.ToString("yyyyMMdd");
                string buyNO = "";
                DataTable dt = db.DB_GetData($"SELECT a.MFNO,a.CalendarDate,a.PartNO,sum(a.NeedQTY) as qty,b.Class,a.SimulationId  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a join SoftNetMainDB.[dbo].Material as b on a.PartNO=b.PartNO where b.ServerId='{_Fun.Config.ServerId}' and a.SimulationId in ({id}) group by a.MFNO,a.CalendarDate,a.PartNO,a.SimulationId,b.Class order by a.MFNO,a.CalendarDate,a.PartNO");
                if (dt != null && dt.Rows.Count > 0)
                {
                    tmp_mfno = dt.Rows[0]["MFNO"].ToString();
                    foreach (DataRow d in dt.Rows)
                    {
                        if (tmp_mfno != d["MFNO"].ToString())
                        {
                            buyNO = $"AA02{date}{(++no).ToString().PadLeft(4, '0')}";
                            sql = $"INSERT INTO SoftNetMainDB.[dbo].[DOC1Buy] (ServerId,DOCNO,DOCNumberNO,DOCDate,UserId,DOCType,MFNO,TotalMoney,TaxMoney,FlowLevel,FlowStatus) VALUES ('{_Fun.Config.ServerId}','AA02','{buyNO}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','Alex','1','{tmp_mfno}',100,5,0,'Y')";
                            tmp_mfno = d["MFNO"].ToString();
                            if (db.DB_SetData(sql))
                            {
                                foreach (string[] s in buys)
                                {
                                    tmp_Id = _Str.NewId('Z');
                                    sql = $"INSERT INTO SoftNetMainDB.[dbo].[DOC1BuyII] (Id,DOCNumberNO,PartNO,Price,Unit,QTY,SimulationId,ArrivalDate,StoreNO,StoreSpacesNO) VALUES ('{tmp_Id}','{buyNO}','{s[0]}',0,'PCS','{s[2]}','{s[4]}','{s[1]}','a1','')";
                                    if (db.DB_SetData(sql))
                                    {
                                        #region 檢驗,入庫單
                                        /*
                                        DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetMainDB.[dbo].[Material] where PartNO='{s[0]}'");
                                        if (tmp_dr != null)
                                        {
                                            if (bool.Parse(tmp_dr["IS_Store_Test"].ToString()))
                                            {
                                                //檢驗單
                                                if (_poolTEST_NO.ContainsKey(buyNO))
                                                { _poolTEST_NO[buyNO].Add(new string[] { s[0],  s[1], s[2], s[3], s[4], tmp_Id }); }
                                                else
                                                {
                                                    List<string[]> tmp = new List<string[]>();
                                                    tmp.Add(new string[] { s[0], s[1], s[2], s[3], s[4], tmp_Id });
                                                    _poolTEST_NO.Add(buyNO, tmp);
                                                }
                                                ////###??? 暫時改不連續發單
                                                ////入庫單
                                                //if (_poolStore_NO.ContainsKey($"????{buyNO}"))
                                                //{ _poolStore_NO[buyNO].Add(new string[] { s[0], s[1], s[2], s[3], s[4], tmp_Id }); }
                                                //else
                                                //{
                                                //    List<string[]> tmp = new List<string[]>();
                                                //    tmp.Add(new string[] { s[0], s[1], s[2], s[3], s[4], tmp_Id });
                                                //    _poolStore_NO.Add($"????{buyNO}", tmp);
                                                //}
                                                
                                            }
                                            else
                                            {
                                                //入庫單
                                                if (_poolStore_NO.ContainsKey(buyNO))
                                                { _poolStore_NO[buyNO].Add(new string[] { s[0], s[1], s[2], s[3], s[4], tmp_Id }); }
                                                else
                                                {
                                                    List<string[]> tmp = new List<string[]>();
                                                    tmp.Add(new string[] { s[0], s[1], s[2], s[3], s[4], tmp_Id });
                                                    _poolStore_NO.Add(buyNO, tmp);
                                                }
                                            }
                                        }
                                        */
                                        #endregion
                                        //回寫APS_PartNOTimeNote
                                        //db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO = '{buyNO}' where SimulationId='{s[4]}'");
                                        //回寫APS_Simulation
                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET DOCNumberNO = '{buyNO}' where SimulationId='{s[4]}'");
                                    }
                                }
                            }
                            else
                            {
                                string _s = "";
                            }
                            buys.Clear();
                        }
                        buys.Add(new string[] { d["PartNO"].ToString(), Convert.ToDateTime(d["CalendarDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff"), d["qty"].ToString(), d["Class"].ToString(), d["SimulationId"].ToString() ,""});
                    }
                    if (buys.Count > 0)
                    {
                        buyNO = $"AA02{date}{(++no).ToString().PadLeft(4, '0')}";
                        sql = $"INSERT INTO SoftNetMainDB.[dbo].[DOC1Buy] (ServerId,DOCNO,DOCNumberNO,DOCDate,UserId,DOCType,MFNO,TotalMoney,TaxMoney,FlowLevel,FlowStatus) VALUES ('{_Fun.Config.ServerId}','AA02','{buyNO}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','Alex','1','{tmp_mfno}',100,5,0,'Y')";
                        if (db.DB_SetData(sql))
                        {
                            //buys 0="PartNO 1=CalendarDate 2=qty 3=Class 4=SimulationId 5=上筆關聯Id
                            foreach (string[] s in buys)
                            {
                                tmp_Id = _Str.NewId('Z');
                                sql = $"INSERT INTO SoftNetMainDB.[dbo].[DOC1BuyII] (Id,DOCNumberNO,PartNO,Price,Unit,QTY,SimulationId,ArrivalDate,StoreNO,StoreSpacesNO) VALUES ('{tmp_Id}','{buyNO}','{s[0]}',0,'PCS','{s[2]}','{s[4]}','{s[1]}','a1','')";
                                if (db.DB_SetData(sql))
                                {
                                    #region 檢驗,入庫單
                                    /*
                                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetMainDB.[dbo].[Material] where PartNO='{s[0]}'");
                                    if (tmp_dr != null)
                                    {
                                        if (bool.Parse(tmp_dr["IS_Store_Test"].ToString()))
                                        {
                                            //檢驗單
                                            if (_poolTEST_NO.ContainsKey(buyNO))
                                            { _poolTEST_NO[buyNO].Add(new string[] { s[0], s[1], s[2], s[3], s[4], tmp_Id }); }
                                            else
                                            {
                                                List<string[]> tmp = new List<string[]>();
                                                tmp.Add(new string[] { s[0], s[1], s[2], s[3], s[4], tmp_Id });
                                                _poolTEST_NO.Add(buyNO, tmp);
                                            }
                                            ////###??? 暫時改不連續發單
                                            ////入庫單
                                            //if (_poolStore_NO.ContainsKey($"????{buyNO}"))
                                            //{ _poolStore_NO[$"????{buyNO}"].Add(new string[] { s[0], s[1], s[2], s[3], s[4], tmp_Id }); }
                                            //else
                                            //{
                                            //    List<string[]> tmp = new List<string[]>();
                                            //    tmp.Add(new string[] { s[0], s[1], s[2], s[3], s[4], tmp_Id });
                                            //    _poolStore_NO.Add($"????{buyNO}", tmp);
                                            //}
                                        }
                                        else
                                        {
                                            //入庫單
                                            if (_poolStore_NO.ContainsKey(buyNO))
                                            { _poolStore_NO[buyNO].Add(new string[] { s[0], s[1], s[2], s[3], s[4], tmp_Id }); }
                                            else
                                            {
                                                List<string[]> tmp = new List<string[]>();
                                                tmp.Add(new string[] { s[0], s[1], s[2], s[3], s[4], tmp_Id });
                                                _poolStore_NO.Add(buyNO, tmp);
                                            }
                                        }
                                    }
                                    */
                                    #endregion

                                    db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET DOCNumberNO = '{buyNO}' where SimulationId='{s[4]}'");
                                }
                            }
                        }
                        else
                        {
                            string _s = "";
                        }
                    }


                    if (_WebSocket._WebSocketList.ContainsKey(data[0]))
                    {
                        _WebSocket.Send(_WebSocket._WebSocketList[data[0]].socket, "StationStatusChange");
                    }
                }
                return meg;
            }

        }
    }
}
