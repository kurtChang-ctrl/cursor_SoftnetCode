using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Base.Services;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Base
{
    internal class CommLib
    {

    }

    public class LabelWork
    {
        public string Station { get; set; } = "";
        public string Station_Name { get; set; } = "";
        public string StationNO_Custom_DisplayName { get; set; } = "";
        public string Type { get; set; }
        public string TypeValue { get; set; }
        public string OrderNO { get; set; } = "";
        public string SimulationId { get; set; } = "";
        public string IndexSN { get; set; } = "";
        public string OPNO { get; set; }
        public string OPNO_Name { get; set; }
        public int OKQTY { get; set; }
        public int FailQTY { get; set; }
        public string State { get; set; } = "";
        public string Has_Knives { get; set; } = "0";
        public string LocalIPPort { get; set; } = "";
        public string ERRMsg { get; set; } = "";
        public List<string[]> MES_List { get; set; }
        public string MES_String { get; set; } = "";
        public string Station_State { get; set; } = "";
        public string RemarkTimeS { get; set; } = "";
        public string RemarkTimeE { get; set; } = "";
        public string WO_QTY { get; set; } = "0";
        public string TOT_OK_QTY { get; set; } = "0";
        public string TOT_Fail_QTY { get; set; } = "0";
        public string Transfer_QTY { get; set; } = "0";
        public string PartNO_INFO { get; set; } = "";
        public string WO_CT { get; set; } = "0";
        public string E_CT { get; set; } = "0";
        public string Station_Config_Store_Type { get; set; } = "1";
        public string BOOLonToggleMenu { get; set; } = "";

        public List<string[]> HasWO_List { get; set; }
        public List<string[]> HasPO_List { get; set; }



    }
    public class MUTIStationObj
    {
        public string OrderBY_CMD { get; set; } = "order by PartNO,IndexSN,OrderNO,RemarkTimeS desc";
        public string SI_PP_Name { get; set; } = "";
        public string SI_PartNO { get; set; } = "";
        public string SI_PartName { get; set; } = "";
        public string SI_OrderNO { get; set; } = "";
        public string SI_IndexSN { get; set; } = "";
        public int SI_OKQTY { get; set; }
        public int SI_FailQTY { get; set; }
        public string SI_OPNO { get; set; } = "";
        public string SI_Slect_OPNOs { get; set; } = "";
        public string SI_SimulationId { get; set; } = "";
        public string Has_Knives { get; set; } = "0";
        public string State { get; set; } = "";
        public string StationNO { get; set; } = "";
        public string StationName { get; set; }
        public string LocalIPPort { get; set; } = "";
        public List<string[]> HasPP_Name_List { get; set; }
        public List<string[]> HasPartNO_List { get; set; }
        public List<string[]> HasWorkData_List { get; set; }
        public List<string[]> HasPO_List { get; set; }
        public List<string[]> HasWO_List { get; set; }
        public string Station_Config_Store_Type { get; set; } = "1";

        public string Remark { get; set; } = "";
        public string MES_String { get; set; } = "";
        public string MES_Report { get; set; } = "";
        public string ERRMsg { get; set; } = "";
        public string Select_ID { get; set; } = "";
        public List<string[]> StationNOList { get; set; }
        public bool OutError { get; set; } = false;

    }
    public class SystemConfigObj
    {
        //監看Thead狀態  0=SfcTimerloopUpdateTagValue_Tick 1=SfcTimerloopautoRUN_DOC_Tick 2=SfcTimerloopautoRUN_Json_Tick 3=SfcTimerloopthread_Tick 4=DeviceConnectCheck_Tick

        public long _a01 { get; } = _Fun._a01;
        public long _a02 { get; } = _Fun._a02;
        public long _a03 { get; } = _Fun._a03;
        public long _a04 { get; } = _Fun._a04;
        public long _a05 { get; } = _Fun._a05;
        public long _a06 { get; } = _Fun._a06;
        public long _a07 { get; } = _Fun._a07;
        public long _a08 { get; } = _Fun._a08;
        public long _a09 { get; } = _Fun._a09;
        public long _a10 { get; } = _Fun._a10;
        public long _a11 { get; } = _Fun._a11;
        public long _a12 { get; } = _Fun._a12;
        public long _a13 { get; } = _Fun._a13;
        public long _a14 { get; } = _Fun._a14;
        public long _a15 { get; } = _Fun._a15;
        public long _a16 { get; } = _Fun._a16;

        public bool T0 { get; } = _Fun.Is_RUNTimeServer_Thread_State[0];
        public bool T1 { get; } = _Fun.Is_RUNTimeServer_Thread_State[1];
        public bool T2 { get; } = _Fun.Is_RUNTimeServer_Thread_State[2];
        public bool T3 { get; } = _Fun.Is_RUNTimeServer_Thread_State[3];
        public bool T4 { get; } = _Fun.Is_RUNTimeServer_Thread_State[4];
        public bool Thread_ForceClose { get; } = _Fun.Is_Thread_ForceClose;
        public bool Is_Thread_For_Test { get; } = _Fun.Is_Thread_For_Test;

        public string Send_RUNTimeServer_Thread_State { get; set; } = "";

        public string FunType { get; set; } = "";
    }
    public class OddCT_GetData
    {
        public string Key { get; set; } = "";
        public string Value { get; set; } = "";
    }
    public class LabelProject
    {
        public string PP_Name { get; set; }
        public string PartNO { get; set; } = "";
        public string PartName { get; set; } = "";
        public string Specification { get; set; } = "";
        public string StationNO { get; set; } = "";
        public string StationName { get; set; }
        public string Type { get; set; }
        public string OrderNO { get; set; }
        public string IndexSN { get; set; } = "0";
        public string OPNO { get; set; } = "";
        public string OPNO_Name { get; set; }
        public string State { get; set; } = "";

        public int TOTALOKQTY { get; set; } = 0;
        public int TOTALFailQTY { get; set; } = 0;

        public string TOTALWorkTime { get; set; } = "";
        public int OKQTY { get; set; } = 0;
        public int FailQTY { get; set; } = 0;
        public string LocalIPPort { get; set; } = "";
        public string ERRMsg { get; set; } = "";
        public List<string[]> MES_List { get; set; }
        public string MES_String { get; set; } = "";
        public string Station_State { get; set; } = "";

        public List<string[]> HasPP_Name_List { get; set; }
        public List<string[]> HasPO_List { get; set; }
        public List<string[]> HasPartNO_List { get; set; }
        public string Remark { get; set; } = "";

    }

    [Serializable]
    public class RMSProtocol
    {
        public int DataType = 0;
        public string Data = "";
        public ushort PoolID = 0;
        public bool IsReturnID = false;
        public uint TransferID = 0;
        public RMSProtocol()
        {
            DataType = 0;
            Data = "";
            PoolID = 0;
            IsReturnID = false;
        }
        public RMSProtocol(int datatype, string data, uint tid = 0)
        {
            DataType = datatype;
            Data = data;
            PoolID = 0;
            IsReturnID = false;
            TransferID = tid;
        }
        public RMSProtocol(int datatype, string data, bool rID, uint tid = 0)
        {
            DataType = datatype;
            Data = data;
            PoolID = 0;
            IsReturnID = rID;
            TransferID = tid;
        }
        public RMSProtocol(int datatype, string data, ushort id, uint tid = 0)
        {
            DataType = datatype;
            Data = data;
            PoolID = id;
            IsReturnID = false;
            TransferID = tid;
        }
        public RMSProtocol Clone()
        {
            return (RMSProtocol)this.MemberwiseClone();
        }
    }
    public class rmsMasterUserData : IDisposable
    {
        public string deviceName = "";
        public Socket socket;
        public IPEndPoint ipPoint;
        public byte Role = 0;    //0=None 1=RMS 5=Dashborad 10=TMRobot  15=Tools   20=gatwey 25=web
        public bool isWork = false;
        public bool ReSetFlag = false;
        public string AddrString = "";
        rmsMasterUserData(string deviceName)
        {
            this.deviceName = deviceName;
        }
        public rmsMasterUserData(string deviceName, Socket socket, IPEndPoint ipPoint)
        {
            this.deviceName = deviceName;
            this.socket = socket;
            this.ipPoint = ipPoint;
            this.AddrString = ipPoint.ToString();
        }
        public rmsMasterUserData(string deviceName, Socket socket, IPEndPoint ipPoint, byte role)
        {
            this.deviceName = deviceName;
            this.socket = socket;
            this.ipPoint = ipPoint;
            this.Role = role;
            this.AddrString = ipPoint.ToString();
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private bool disposed = false;
        protected virtual void Dispose(bool disposing)
        {
            if (disposed) { return; }
            if (disposing)
            {
                //清理CLR託管資源
                //system_eventLog.Dispose();
                if (socket != null) { socket.Close(); }
            }
            //清理非託管資源,寫在下方,如果有的話
            disposed = true;
        }
    }
    public class rmsConectUserData : IDisposable
    {
        public string deviceName;
        public string ipcRobotName;
        public Socket socket;
        public bool IsSupportUser;
        public byte Role;    //0=None 1=RMS 5=Dashborad 6=SFCClient 10=TMRobot  15=Tools   20=gatwey  25=Skymars   30=websocket
        public bool IsCheckOwner;//手臂端是否被取得控制權
        public bool IsAutoRemoteMode;//手臂端是否為AutoRemoteMode
        public bool IsSpeedAdjustment;//手臂端是否可調速度
        public string RobotBrandName;//紀錄手臂 "" "TM" "OMRON"
        public int Robot_Ver01;
        public int Robot_Ver02;
        public int Robot_Ver03;
        public string Robot_Status;
        public bool RobotAutoVarSync;
        public string Station_UI_StationNO_list;

        public rmsConectUserData(string deviceName, string ipcRobotName, Socket socket)
        {
            this.deviceName = deviceName;
            this.ipcRobotName = ipcRobotName;
            this.socket = socket;
            IsSupportUser = false;
            Role = 0;
            IsCheckOwner = true;//20210326 EDITOR KURT 寫死不看手臂控制權
            IsAutoRemoteMode = false;
            IsSpeedAdjustment = false;
            RobotBrandName = "";
            Robot_Ver01 = 1;
            Robot_Ver02 = 74;
            Robot_Ver03 = 0;
            Robot_Status = "";
            RobotAutoVarSync = true;
            Station_UI_StationNO_list = "";
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private bool disposed = false;
        protected virtual void Dispose(bool disposing)
        {
            if (disposed) { return; }
            if (disposing)
            {
                //清理CLR託管資源
                //system_eventLog.Dispose();
                if (socket != null) { socket.Close(); socket = null; }
            }
            //清理非託管資源,寫在下方,如果有的話
            disposed = true;

        }
    }


    public class SFC_Common : IDisposable
    {
        private string DBtype = "1";
        private string DBconnetstring = "";
        public SFC_Common(string dbtype,string dbconnetstring)
        {
            DBtype = dbtype;
            DBconnetstring = dbconnetstring;
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private bool disposed = false;
        protected virtual void Dispose(bool disposing)
        {
            if (disposed) { return; }
            if (disposing)
            {
                //清理CLR託管資源

            }
            //清理非託管資源,寫在下方,如果有的話
            disposed = true;

        }




        public string[] Get_Service_DB_TypeandConnectString(string dbType, string dbConnectString, string dbID)
        {
            string[] re = null;
            /*
            if (dbID != "")
            {
                DataSourceCollection _dataSourceCollection = DataSourceFactory.GetDataSources(dbType, dbConnectString);
                foreach (DataSource ds in _dataSourceCollection)
                {
                    if (ds.ID.ToString().Trim() == dbID)
                    {
                        re = new string[] { "", "" };
                        re[0] = ds.DBType2String();
                        re[1] = ds.ConnectionString;
                        break;
                    }
                }
            }
            */
            return re;
        }
        public List<string> Process_ALLSub_PP_Name(string pp_Name, string configFilePath, bool isRun_PP_ProductProcess_Item, string orderN)
        {
            /*
            if (pp_Name != "")
            {
                DataSourceCollection _dataSourceCollection = CommonLib.Database.DataSourceFactory.GetMasterDataSources();
                string dbType = "1";
                string dbConnectString = "";
                if (configFilePath == "")
                {
                    _dataSourceCollection = CommonLib.Database.DataSourceFactory.GetMasterDataSources();
                }
                else
                {
                    _dataSourceCollection = CommonLib.Database.DataSourceFactory.GetMasterDataSources(configFilePath);
                }
                foreach (CommonLib.Database.DataSource ds in _dataSourceCollection)
                {
                    if (ds.IsConnected)
                    {
                        dbType = ds.DBType2String();
                        dbConnectString = ds.ConnectionString;
                        break;
                    }
                }
                if (dbConnectString != "")
                {
                    return process_ALLSub_PP_NameII(dbType, dbConnectString, pp_Name, isRun_PP_ProductProcess_Item, orderN);
                }
            }
*/
            return null;
        }
        public DataTable Process_ALLSation(string ServerId, string pp_Name, string order_string, string configFilePath, bool isRun_PP_ProductProcess_Item, string orderN)
        {
            if (pp_Name != "")
            {
                /*
                DataSourceCollection _dataSourceCollection = CommonLib.Database.DataSourceFactory.GetMasterDataSources();
                string dbType = "1";
                string dbConnectString = "";
                if (configFilePath == "")
                {
                    _dataSourceCollection = CommonLib.Database.DataSourceFactory.GetMasterDataSources();
                }
                else
                {
                    _dataSourceCollection = CommonLib.Database.DataSourceFactory.GetMasterDataSources(configFilePath);
                }
                foreach (CommonLib.Database.DataSource ds in _dataSourceCollection)
                {
                    if (ds.IsConnected)
                    {
                        dbType = ds.DBType2String();
                        dbConnectString = ds.ConnectionString;
                        break;
                    }
                }
                */
                if (DBconnetstring != "")
                {
                    if (order_string == "") { order_string = " order by PP_Name,DisplaySN,IndexSN"; }
                    return process_ALLSationII(ServerId,DBtype, DBconnetstring, pp_Name, order_string, isRun_PP_ProductProcess_Item, orderN);
                }
            }
            return null;
        }
        public DataTable Process_ALLSation(string serverId, string dbType, string dbConnectString, string pp_Name, string order_string, bool isRun_PP_ProductProcess_Item, string orderNO)
        {
            if (pp_Name != "")
            {
                if (order_string == "") { order_string = " order by PP_Name,DisplaySN,IndexSN"; }
                return process_ALLSationII(serverId,dbType, dbConnectString, pp_Name, order_string, isRun_PP_ProductProcess_Item, orderNO);
            }
            return null;
        }
        public DataTable Process_ALLSation_RE_Custom(string serverId, string dbType, string dbConnectString, string pp_Name, string order_string, bool isRun_PP_ProductProcess_Item, string orderNO)
        {
            if (pp_Name != "")
            {
                if (order_string == "") { order_string = " order by DisplaySN,IndexSN"; }
                return process_ALLSationII(serverId, dbType, dbConnectString, pp_Name, order_string, isRun_PP_ProductProcess_Item, orderNO, true);
            }
            return null;
        }
        public Dictionary<string, List<string>> Process_Class_IndexSN(string dbType, string dbConnectString, string pp_Name, bool isRun_PP_ProductProcess_Item, string orderNO)//回傳製成階層
        {
            if (pp_Name != "")
            {
                return Process_Class_IndexSNII(dbType, dbConnectString, pp_Name, isRun_PP_ProductProcess_Item, orderNO);
            }
            return null;
        }
        private Dictionary<string, List<string>> Process_Class_IndexSNII(string dbType, string dbConnectString, string pp_Name, bool isRun_PP_ProductProcess_Item, string orderNO)
        {
            if (pp_Name != "")
            {
                using (DBADO db = new DBADO(dbType, dbConnectString))
                {
                    int tmp_befor_IndexSN = 0;
                    int tmp_count = 0;
                    Dictionary<string, List<string>> groupsItemList = new Dictionary<string, List<string>>();
                    if (db != null)
                    {
                        DataTable tmp = null;
                        if (isRun_PP_ProductProcess_Item)
                        { tmp = db.DB_GetData(string.Format("SELECT * FROM SoftNetSYSDB.[dbo].PP_ProductProcess_Item where PP_Name=N'{0}' order by IndexSN,DisplaySN", pp_Name)); }
                        else
                        { tmp = db.DB_GetData(string.Format("SELECT * FROM SoftNetSYSDB.[dbo].PP_WO_Process_Item where PP_Name=N'{0}' and OrderNO=N'{1}' order by IndexSN,DisplaySN", pp_Name, orderNO)); }
                        if (tmp != null && tmp.Rows.Count > 0)
                        {
                            string tmp_StationNO = "";
                            bool tmp_fristRUN = true;
                            string tmp_id = "";
                            foreach (DataRow dr in tmp.Rows)
                            {
                                tmp_StationNO = dr["StationNO"].ToString().Trim();
                                if (tmp_fristRUN || tmp_befor_IndexSN != int.Parse(dr["IndexSN"].ToString()))
                                {
                                    tmp_befor_IndexSN = int.Parse(dr["IndexSN"].ToString());
                                    tmp_count = 1;
                                    tmp_id = string.Format("{0},{1}", tmp_befor_IndexSN.ToString(), tmp_count.ToString());
                                    if (tmp_StationNO != null && tmp_StationNO != "")
                                    { groupsItemList.Add(tmp_id, new List<string>()); }
                                }
                                else
                                {
                                    ++tmp_count;
                                    tmp_id = string.Format("{0},{1}", tmp_befor_IndexSN.ToString(), tmp_count.ToString());
                                    if (tmp_StationNO != null && tmp_StationNO != "")
                                    { groupsItemList.Add(tmp_id, new List<string>()); }
                                }
                                tmp_fristRUN = false;
                                if (tmp_StationNO != "")
                                {
                                    if (!groupsItemList.ContainsKey(tmp_id)) { groupsItemList.Add(tmp_id, new List<string>()); }
                                    groupsItemList[tmp_id].Add(tmp_StationNO);
                                }
                                else if (!dr.IsNull("Sub_PP_Name") && dr["Sub_PP_Name"].ToString().Trim() != "")
                                {
                                    class_IndexSN_Recursively(db, ref groupsItemList, tmp_id, dr["Sub_PP_Name"].ToString().Trim(), isRun_PP_ProductProcess_Item, orderNO);
                                }
                            }
                        }
                    }
                    return groupsItemList;
                }
            }
            return null;
        }
        private void class_IndexSN_Recursively(DBADO db, ref Dictionary<string, List<string>> groupsItemList, string tmp_string_IndexSN, string sub_PP_Name, bool isRun_PP_ProductProcess_Item, string orderNO)
        {
            int tmp_befor_IndexSN = 0;
            int tmp_count = 0;
            DataTable tmp = null;
            if (isRun_PP_ProductProcess_Item)
            { tmp = db.DB_GetData(string.Format("SELECT * FROM SoftNetSYSDB.[dbo].PP_ProductProcess_Item where PP_Name=N'{0}' order by IndexSN,DisplaySN", sub_PP_Name)); }
            else
            { tmp = db.DB_GetData(string.Format("SELECT * FROM SoftNetSYSDB.[dbo].PP_WO_Process_Item where PP_Name=N'{0}' and OrderNO=N'{1}' order by IndexSN,DisplaySN", sub_PP_Name, orderNO)); }
            if (tmp != null && tmp.Rows.Count > 0)
            {
                string tmp_StationNO = "";
                bool tmp_fristRUN = true;
                string tmp_id = "";
                foreach (DataRow dr in tmp.Rows)
                {
                    tmp_StationNO = dr["StationNO"].ToString().Trim();
                    if (tmp_fristRUN || tmp_befor_IndexSN != int.Parse(dr["IndexSN"].ToString()))
                    {
                        tmp_befor_IndexSN = int.Parse(dr["IndexSN"].ToString());
                        tmp_count = 1;
                        tmp_id = string.Format("{0},{1},{2}", tmp_string_IndexSN, tmp_befor_IndexSN.ToString(), tmp_count.ToString());
                        if (tmp_StationNO != null && tmp_StationNO != "")
                        { groupsItemList.Add(tmp_id, new List<string>()); }
                    }
                    else
                    {
                        ++tmp_count;
                        tmp_id = string.Format("{0},{1},{2}", tmp_string_IndexSN, tmp_befor_IndexSN.ToString(), tmp_count.ToString());
                        if (tmp_StationNO != null && tmp_StationNO != "")
                        { groupsItemList.Add(tmp_id, new List<string>()); }
                    }
                    tmp_fristRUN = false;
                    if (tmp_StationNO != "")
                    {
                        if (!groupsItemList.ContainsKey(tmp_id)) { groupsItemList.Add(tmp_id, new List<string>()); }
                        groupsItemList[tmp_id].Add(tmp_StationNO);
                    }
                    else if (!dr.IsNull("Sub_PP_Name") && dr["Sub_PP_Name"].ToString().Trim() != "")
                    {
                        class_IndexSN_Recursively(db, ref groupsItemList, tmp_id, dr["Sub_PP_Name"].ToString().Trim(), isRun_PP_ProductProcess_Item, orderNO);
                    }
                }
            }
        }
        private List<string> process_ALLSub_PP_NameII(string dbType, string dbConnectString, string pp_Name, bool isRun_PP_ProductProcess_Item, string orderNO)
        {
            string serverId = "";
            List<string> ProcessList = new List<string>();
            using (DBADO db = new DBADO(dbType, dbConnectString))
            {
                //製程清單
                if (db != null)
                {
                    DataTable tmp = null;
                    if (isRun_PP_ProductProcess_Item)
                    { tmp = db.DB_GetData(string.Format("SELECT Sub_PP_Name FROM SoftNetSYSDB.[dbo].PP_ProductProcess_Item where PP_Name=N'{0}' and Sub_PP_Name != ''", pp_Name)); }
                    else
                    { tmp = db.DB_GetData(string.Format("SELECT Sub_PP_Name FROM SoftNetSYSDB.[dbo].PP_WO_Process_Item where PP_Name=N'{0}' and OrderNO=N'{1}' and Sub_PP_Name != ''", pp_Name, orderNO)); }
                    if (tmp != null && tmp.Rows.Count > 0)
                    {
                        foreach (DataRow dr in tmp.Rows)
                        {
                            ProcessList.Add(dr["Sub_PP_Name"].ToString());
                            sub_PP_Name_Recursively('1', db, dr["Sub_PP_Name"].ToString(), ref ProcessList, isRun_PP_ProductProcess_Item, orderNO, serverId);
                        }
                    }
                }
            }
            return ProcessList;
        }
        private void sub_PP_Name_Recursively(char type, DBADO db, string sub_PP_Name, ref List<string> ProcessList, bool isRun_PP_ProductProcess_Item, string orderNO,string serverId)
        {
            DataTable tmp = null;
            if (isRun_PP_ProductProcess_Item)
            { tmp = db.DB_GetData($"SELECT Sub_PP_Name FROM SoftNetSYSDB.[dbo].PP_ProductProcess_Item where ServerId='{serverId}' and PP_Name=N'{sub_PP_Name}' and Sub_PP_Name != ''"); }
            else
            { tmp = db.DB_GetData($"SELECT Sub_PP_Name FROM SoftNetSYSDB.[dbo].PP_WO_Process_Item where  ServerId='{serverId}' and PP_Name=N'{sub_PP_Name}' and OrderNO=N'{orderNO}' and Sub_PP_Name != ''"); }
            if (tmp != null && tmp.Rows.Count > 0)
            {
                foreach (DataRow dr in tmp.Rows)
                {
                    if (type == '1')
                    { ProcessList.Add(dr["Sub_PP_Name"].ToString()); }
                    else
                    { ProcessList.Add(string.Format("'{0}'", dr["Sub_PP_Name"].ToString())); }
                    sub_PP_Name_Recursively(type, db, dr["Sub_PP_Name"].ToString(), ref ProcessList, isRun_PP_ProductProcess_Item, orderNO, serverId);
                }
            }
        }
        private DataTable process_ALLSationII(string serverId, string dbType, string dbConnectString, string pp_Name, string order_string, bool isRun_PP_ProductProcess_Item, string orderNO, bool re_Custom = false)
        {
            DataTable dt = null;
            if (pp_Name != "")
            {
                using (DBADO db = new DBADO(dbType, dbConnectString))
                {
                    //製程清單
                    List<string> ProcessList = new List<string>() { string.Format("'{0}'", pp_Name) };
                    if (db != null)
                    {
                        #region 遞迴所有子製程清單
                        DataTable tmp = null;
                        if (isRun_PP_ProductProcess_Item)
                        { tmp = db.DB_GetData($"SELECT Sub_PP_Name FROM SoftNetSYSDB.[dbo].PP_ProductProcess_Item WHERE ServerId='{serverId}' and PP_Name=N'{pp_Name}' AND Sub_PP_Name != ''"); }
                        else
                        { tmp = db.DB_GetData($"SELECT Sub_PP_Name FROM SoftNetSYSDB.[dbo].PP_WO_Process_Item WHERE ServerId='{serverId}' and PP_Name=N'{pp_Name}' and OrderNO=N'{orderNO}' AND Sub_PP_Name != ''"); }
                        if (tmp != null && tmp.Rows.Count > 0)
                        {
                            foreach (DataRow dr in tmp.Rows)
                            {
                                ProcessList.Add(string.Format("'{0}'", dr["Sub_PP_Name"].ToString()));
                                sub_PP_Name_Recursively('2', db, dr["Sub_PP_Name"].ToString(), ref ProcessList, isRun_PP_ProductProcess_Item, orderNO, serverId);
                            }
                        }
                        #endregion
                    }
                    if (re_Custom)
                    {
                        string sql = "";
                        if (isRun_PP_ProductProcess_Item)
                        {
                            sql = string.Format(
                                         @"SELECT A.[Id],A.[FactoryName] as [Factory Name],A.[LineName] as [Line Name],A.[PP_Name] as [Process Name],
                                           B.[StationName] as [Station Name],A.[StationNO] as [Station NO], A.[Station_Custom_IndexSN] as Station_Custom_IndexSN, A.[DisplaySN] as [Display Index],
                                           A.[IndexSN] as [Process Index],A.[Sub_PP_Name] as [SubProcess Name],A.[IndexSN_Merge] as [Merge],A.E_CycleTime [Cycle Time],
                                           A.[BelowYield] as [Yield Warning Lower Bound],A.[BelowCycleTime] as [Cycle Time Warning Upper Bound],B.[RMSName] as [TM Service]
                                           FROM (SELECT * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{2}' and  IndexSN!=0 AND StationNO!='' and PP_Name in ({0})) as A left join 
                                           (SELECT [StationName], [StationNO], [RMSName] FROM SoftNetSYSDB.[dbo].[PP_Station]) as B on A.StationNO = B.StationNO {1}",
                                           string.Join(",", ProcessList.ToArray()), order_string, serverId);//多取一個Station_Custom_IndexSN
                        }
                        else
                        {
                            sql = string.Format(
                                         @"SELECT A.[sn],A.[FactoryName] as [Factory Name],A.[LineName] as [Line Name],A.[PP_Name] as [Process Name],
                                           B.[StationName] as [Station Name],A.[StationNO] as [Station NO], A.[Station_Custom_IndexSN] as Station_Custom_IndexSN, A.[DisplaySN] as [Display Index],
                                           A.[IndexSN] as [Process Index],A.[Sub_PP_Name] as [SubProcess Name],A.[IndexSN_Merge] as [Merge],A.E_CycleTime [Cycle Time],
                                           A.[BelowYield] as [Yield Warning Lower Bound],A.[BelowCycleTime] as [Cycle Time Warning Upper Bound],B.[RMSName] as [TM Service]
                                           FROM (SELECT * from SoftNetSYSDB.[dbo].[PP_WO_Process_Item] where IndexSN!=0 AND StationNO!='' and OrderNO=N'{2}' and PP_Name in ({0})) as A left join 
                                           (SELECT [StationName], [StationNO], [RMSName] FROM SoftNetSYSDB.[dbo].[PP_Station]) as B on A.StationNO = B.StationNO {1}",
                                           string.Join(",", ProcessList.ToArray()), order_string, orderNO);//多取一個Station_Custom_IndexSN

                        }
                        dt = db.DB_GetData(sql);
                    }
                    else
                    {
                        string sql = "";
                        if (isRun_PP_ProductProcess_Item)
                        {
                            sql = string.Format(
                            @"SELECT * FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{2}' and IndexSN!=0 AND StationNO!='' and PP_Name in ({0}) {1}",
                            string.Join(",", ProcessList.ToArray()), order_string, serverId);
                        }
                        else
                        {
                            sql = string.Format(
                            @"SELECT * FROM SoftNetSYSDB.[dbo].[PP_WO_Process_Item] where IndexSN!=0 AND StationNO!='' and OrderNO=N'{2}' and PP_Name in ({0}) {1}",
                            string.Join(",", ProcessList.ToArray()), order_string, orderNO);
                        }
                        dt = db.DB_GetData(sql);
                    }
                }
            }
            return dt;
        }



        private Dictionary<char, List<KeyAndValue>> _efficientConfig = new Dictionary<char, List<KeyAndValue>>();
        public void SfcTimerloopthread_Tick_Efficient(DBADO db, List<double> allCT, string stationNO, string master_PP_Name, string pp_Name, string indexSN, string partNO, string sub_PartNO, string docNO, string Next_StationNO = "", string MFNO = "")
        {
            try
            {
                if (_efficientConfig.Count == 0)
                {
                    //###???將來要改 TMService ㄝ要改
                    _efficientConfig.Add('A', new List<KeyAndValue>());
                    _efficientConfig.Add('B', new List<KeyAndValue>());
                    _efficientConfig.Add('C', new List<KeyAndValue>());
                    _efficientConfig.Add('D', new List<KeyAndValue>());
                    DataTable dt_Efficient = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_EfficientConfig where ServerId='{_Fun.Config.ServerId}' order by EfficiencyType,TypeKey");
                    if (dt_Efficient != null && dt_Efficient.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt_Efficient.Rows)
                        {
                            switch (dr["EfficiencyType"].ToString())
                            {
                                case "A": _efficientConfig['A'].Add(new KeyAndValue(dr["TypeKey"].ToString(), dr["TypeValue"].ToString())); break;
                                case "B": _efficientConfig['B'].Add(new KeyAndValue(dr["TypeKey"].ToString(), dr["TypeValue"].ToString())); break;
                                case "C": _efficientConfig['C'].Add(new KeyAndValue(dr["TypeKey"].ToString(), dr["TypeValue"].ToString())); break;
                                case "D": _efficientConfig['D'].Add(new KeyAndValue(dr["TypeKey"].ToString(), dr["TypeValue"].ToString())); break;
                            }
                        }
                    }
                }

                allCT.Sort();
                double totalP = 0;
                double sumsqrP = 0;
                double avg = 0;//有效CycleTime
                double std = 0;
                int Q1index = 0;//startIndex: first verified value index, endIndex: last verified value index
                int Q3index = 0;
                int startIndex = 0;
                int endIndex = 0;
                double IRQ;
                double upper = 0, lower = 0;
                double x = 1.5;//IRQ倍率
                double ct_avg = allCT.Average();//原平均值
                Q1index = (int)allCT.Count / 4; //4分後的第3線
                Q3index = (int)(allCT.Count * 3 / 4);//4分後的第1線
                IRQ = (allCT[Q3index] - allCT[Q1index] >= 0) ? allCT[Q3index] - allCT[Q1index] : 0;//第1線-第3線=差
                lower = allCT[Q1index] - (x * IRQ);
                upper = allCT[Q3index] + (x * IRQ);
                for (int i1 = 0; i1 < allCT.Count; i1++)
                    if (allCT[i1] >= lower)
                    {
                        startIndex = i1;
                        lower = allCT[i1];
                        break;
                    }
                for (int i1 = allCT.Count - 1; i1 >= 0; i1--)
                    if (allCT[i1] <= upper)
                    {
                        endIndex = i1;
                        upper = allCT[i1];
                        break;
                    }

                for (int i1 = startIndex; i1 <= endIndex; i1++)
                {
                    totalP += allCT[i1];//sum of CT
                    sumsqrP += allCT[i1] * allCT[i1];//sum of CT square
                }
                if (allCT.Count > 1)
                {
                    int tmp = endIndex - startIndex + 1;
                    avg = totalP / tmp;
                    std = Math.Sqrt(sumsqrP / tmp - avg * avg);
                    if (double.IsNaN(std)) { std = 0; }
                }
                else
                { avg = ct_avg; }
                double custom_SD_LowerLimit = lower;
                #region 客製效能
                if (avg > lower)
                {
                    KeyAndValue e_tmp = _efficientConfig['D'].Find(delegate (KeyAndValue t) { return t.Key == partNO; });
                    if (e_tmp.Key != null)
                    {
                        custom_SD_LowerLimit = avg - ((avg - lower) * (double.Parse(e_tmp.Value) / 100));
                    }
                    else
                    {
                        e_tmp = _efficientConfig['B'].Find(delegate (KeyAndValue t) { return t.Key == stationNO; });
                        if (e_tmp.Key != null)
                        {
                            custom_SD_LowerLimit = avg - ((avg - lower) * (double.Parse(e_tmp.Value) / 100));
                        }
                        else
                        {
                            //###??? 未完
                            /*
                            e_tmp = _efficientConfig['C'].Find(delegate (KeyAndValue t) { return t.Key == stationNO; });
                            if (e_tmp.Key != null)
                            {
                                custom_SD_LowerLimit = ((avg - lower) * (double.Parse(e_tmp.Value) / 100)) + avg;
                            }
                            */
                            if (_efficientConfig['A'].Count > 0)
                            {
                                //###??? 未完要能分哪廠
                                custom_SD_LowerLimit = avg - ((avg - lower) * (double.Parse(_efficientConfig['A'][0].Value) / 100));
                            }
                        }
                    }
                }
                #endregion
                //string stationNO, string op_NO, string master_PP_Name, string pp_Name, string partNO, string sub_PartNO, string docNO
                string tmp_Next_S = "";
                if (Next_StationNO != "") { tmp_Next_S = $" and Next_StationNO='{Next_StationNO}'"; }
                if (MFNO != "") { tmp_Next_S = $"{tmp_Next_S} and MFNO='{MFNO}'"; }
                if (db.DB_GetQueryCount($"select PartNO from SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}' and Apply_PP_Name='{master_PP_Name}' and PP_Name='{pp_Name}' and IndexSN={indexSN} and PartNO='{partNO}' and Sub_PartNO='{sub_PartNO}' and DOCNO='{docNO}' {tmp_Next_S}") <= 0)
                {
                    if (Next_StationNO != "") { tmp_Next_S = $"'{Next_StationNO}'"; }
                    else { tmp_Next_S = "NULL"; }
                    string mfNO = "NULL";
                    if (MFNO != "") { mfNO = $"'{MFNO}'"; }

                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_EfficientDetail] (ServerId,[Id],[StationNO],[Apply_PP_Name],[PP_Name],IndexSN,[PartNO],[Sub_PartNO],[DOCNO],[AverageCycleTime],[EfficientCycleTime],[SD_LowerLimit],[SD_UpperLimit],[Custom_SD_LowerLimit],[CountQTY],[STD],Next_StationNO,MFNO)
                                   VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('B')}','{stationNO}','{master_PP_Name}','{pp_Name}',{indexSN},'{partNO}','{sub_PartNO}','{docNO}',{ct_avg.ToString("0.000")},{avg.ToString("0.000")},{lower.ToString("0.000")},{upper.ToString("0.000")},{custom_SD_LowerLimit.ToString("0.000")},{endIndex - startIndex + 1},{std.ToString("0.000")},{tmp_Next_S},{mfNO})");
                }
                else
                {
                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[PP_EfficientDetail] set AverageCycleTime={ct_avg.ToString("0.000")},EfficientCycleTime={avg.ToString("0.000")},SD_LowerLimit={lower.ToString("0.000")},SD_UpperLimit={upper.ToString("0.000")},Custom_SD_LowerLimit={custom_SD_LowerLimit.ToString("0.000")},CountQTY={endIndex - startIndex + 1},STD={std.ToString("0.000")} where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}' and Apply_PP_Name='{master_PP_Name}' and PP_Name='{pp_Name}' and IndexSN={indexSN} and PartNO='{partNO}' and Sub_PartNO='{sub_PartNO}' and DOCNO='{docNO}' {tmp_Next_S}");
                }
            }
            catch (Exception ex)
            {
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"SNWebSocketService SfcTimerloopthread_Tick_Efficient Exception: {ex.Message} {ex.StackTrace}", true);
            }
        }
        public void Update_PP_WorkOrder_Settlement(DBADO db, string wo, string sid)
        {
            string tmp_indexSN = "";
            DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{sid}'");
            DataRow sfcdr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{wo}'");
            if (sfcdr != null)
            {
                #region 處理PP_WorkOrder_Settlement
                string isEND = int.Parse(dr_APS_Simulation["PartSN"].ToString()) == 0 ? "1" : "0";
                string indexSN_Merge = dr_APS_Simulation.IsNull("StationNO_Merge") ? "0" : "1";

                sfcdr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='{wo}' AND StationNO='{dr_APS_Simulation["Source_StationNO"].ToString()}' and IndexSN={dr_APS_Simulation["Source_StationNO_IndexSN"].ToString()}");
                if (sfcdr == null)
                {
                    string name = "";
                    int StatndCycleTime = int.Parse(dr_APS_Simulation["Math_EfficientCT"].ToString());
                    string StartTime = DateTime.Now.ToString("MM/dd/yyyy H:mm:ss");
                    DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT StationName FROM SoftNetSYSDB.[dbo].[PP_Station] WHERE ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_APS_Simulation["Source_StationNO"].ToString()}'");
                    if (tmp != null)
                    { name = tmp.IsNull("StationName") ? "" : tmp["StationName"].ToString(); }
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_EfficientSCT] WHERE ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_APS_Simulation["Source_StationNO"].ToString()}' and PP_Name='{dr_APS_Simulation["Apply_PP_Name"].ToString()}' and PartNO='{dr_APS_Simulation["PartNO"].ToString()}' and IndexSN={dr_APS_Simulation["Source_StationNO_IndexSN"].ToString()}");
                    if (tmp != null && int.Parse(tmp["SCT"].ToString()) > 0)
                    { StatndCycleTime = int.Parse(tmp["SCT"].ToString()); }

                    string sfc_sql = "";
                    if (StartTime == null)
                    {
                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_EfficientSCT] WHERE ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_APS_Simulation["Source_StationNO"].ToString()}' and PP_Name='{dr_APS_Simulation["Apply_PP_Name"].ToString()}' and PartNO='{dr_APS_Simulation["PartNO"].ToString()}' and IndexSN={dr_APS_Simulation["Source_StationNO_IndexSN"].ToString()}");
                        //type += "1";
                        sfc_sql = string.Format(
                                  @"INSERT INTO SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement (
                                                                                    [OrderNO],[StationNO],[StationName],[PP_Name],[IndexSN],[DisplaySN],
                                                                                    [IsLastStation],[Sub_PP_Name],[StatndCycleTime],[UpdateTime],[IndexSN_Merge],
                                                                                    [StartTime],[CumulativeTime],[AvarageCycleTime],[TotalCheckIn],[TotalCheckOut],
                                                                                    [TotalInput],[TotalOutput],[TotalFail],[TotalKeep],[FPY],
                                                                                    [YieldRate],[StationYieldRate],ServerId) VALUES 
                                                                                    ('{0}',N'{1}','{2}','{3}',{4},{5},'{6}','{7}',{8},'{9}','{10}',
                                                                                    null,0,0,0,0,0,0,0,0,0,0,0,'{11}')",
                                    wo, //0
                                    dr_APS_Simulation["Source_StationNO"].ToString(), //1
                                    name, //2
                                    dr_APS_Simulation["Apply_PP_Name"].ToString(), //3
                                    dr_APS_Simulation["Source_StationNO_IndexSN"].ToString(),//4
                                    0,//5 DisplaySN
                                    isEND, //6
                                    dr_APS_Simulation["Apply_PP_Name"].ToString(), //7 Sub_PP_Name
                                    StatndCycleTime.ToString(), //8
                                    DateTime.Now.ToString("MM/dd/yyyy H:mm:ss"), //9
                                    indexSN_Merge, _Fun.Config.ServerId);//10
                    }
                    else
                    {
                        //type += "2";
                        sfc_sql = string.Format(
                                  @"INSERT INTO SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement (
                                                                                    [OrderNO],[StationNO],[StationName],[PP_Name],[IndexSN],[DisplaySN],
                                                                                    [IsLastStation],[Sub_PP_Name],[StartTime],[StatndCycleTime],[UpdateTime],[IndexSN_Merge],
                                                                                    [CumulativeTime],[AvarageCycleTime],[TotalCheckIn],[TotalCheckOut],[TotalInput],
                                                                                    [TotalOutput],[TotalFail],[TotalKeep],[FPY],[YieldRate],[StationYieldRate],ServerId) VALUES                                                                                                                                                                         
                                                                                    ('{0}',N'{1}','{2}','{3}',{4},{5},'{6}','{7}','{8}',{9},'{10}','{11}',
                                                                                    0,0,0,0,0,0,0,0,0,0,0,'{12}')",
                                    wo, //0
                                    dr_APS_Simulation["Source_StationNO"].ToString(), //1
                                    name, //2
                                    dr_APS_Simulation["Apply_PP_Name"].ToString(), //3
                                    dr_APS_Simulation["Source_StationNO_IndexSN"].ToString(),//4
                                    0,//5 DisplaySN
                                    isEND, //6
                                    dr_APS_Simulation["Apply_PP_Name"].ToString(), //7 Sub_PP_Name
                                    StartTime, //8
                                    StatndCycleTime.ToString(), //9
                                    DateTime.Now.ToString("MM/dd/yyyy H:mm:ss"), //10                                                                                    
                                    indexSN_Merge, _Fun.Config.ServerId);//11
                    }
                    if (!db.DB_SetData(sfc_sql))
                    {

                    }
                }
                else
                {
                    if (sfcdr.IsNull("StartTime") || sfcdr["StartTime"].ToString().Trim() == "")
                    {
                        //type += "3";
                        db.DB_SetData(string.Format(
                            @"Update SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement SET 
                                                                                  StartTime='{0}',UpdateTime='{0}',IsLastStation='{3}',IndexSN_Merge='{5}' 
                                                                                  WHERE ServerId='{6}' and OrderNO='{1}' and StationNO='{2}' and IndexSN={4}",
                              DateTime.Now.ToString("MM/dd/yyyy H:mm:ss"),
                              wo,
                              dr_APS_Simulation["Source_StationNO"].ToString(),
                              isEND,
                              dr_APS_Simulation["Source_StationNO_IndexSN"].ToString(),
                              indexSN_Merge, _Fun.Config.ServerId));
                    }
                    else
                    {
                        //type += "4";
                        db.DB_SetData(string.Format(
                            @"Update SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement SET 
                                                                                  UpdateTime='{0}',IsLastStation='{3}',IndexSN_Merge='{5}' 
                                                                                  WHERE ServerId='{6}' and OrderNO='{1}' and StationNO='{2}' and IndexSN={4}",
                              DateTime.Now.ToString("MM/dd/yyyy H:mm:ss"),
                              wo,
                              dr_APS_Simulation["Source_StationNO"].ToString(),
                              isEND,
                              dr_APS_Simulation["Source_StationNO_IndexSN"].ToString(),
                              indexSN_Merge, _Fun.Config.ServerId));
                    }
                }
                #endregion

                #region 執行WorkOrderSettlementUpdate預存程序
                DataRow dataRow = db.DB_GetFirstDataByDataRow(string.Format(@"EXEC SoftNetLogDB.[dbo].WorkOrderSettlementUpdate '{0}','{1}',{2},'{3}'", wo, dr_APS_Simulation["Source_StationNO"].ToString(), dr_APS_Simulation["Source_StationNO_IndexSN"].ToString(), _Fun.Config.ServerId));

                if (dataRow != null)
                {
                    float avgCT = 0, avgWT = 0;
                    if (!dataRow.IsNull("_AVGWT")) { avgWT = float.Parse(dataRow["_AVGWT"].ToString()); }
                    if (!dataRow.IsNull("_AVGCT")) { avgCT = float.Parse(dataRow["_AVGCT"].ToString()); }

                    int Diffint = 0;

                    /*
                        KeyValuePair<NameSpaceDATA, PProcessValue> init_List = _PP_cts_Thread.FirstOrDefault(x => x.Value._PublicDATA.StationNO == dr_WO_Stations["Station NO"].ToString().Trim());
                        if (init_List.Value != null && init_List.Value._ProductProcessMaster != null)
                        {
                            ProductProcessMaster ttmp = init_List.Value._ProductProcessMaster.Find(delegate (ProductProcessMaster t) { return t.IsFree == false && t.OrderNO == dr_PP_WorkOrder_StartWO["OrderNO"].ToString().Trim(); });
                            if (ttmp != null && ttmp.OrderNO.Trim() != "")
                            {
                                TimeSpan Diff = new TimeSpan(DateTime.Now.Ticks - init_List.Value._PublicDATA.LogStartTime.Ticks);
                                if (Diff.TotalSeconds > 0)
                                {
                                    Diffint = Diff.Seconds;
                                }
                                init_List.Value._PublicDATA.LogStartTime = new TimeSpan(DateTime.Now.Ticks);
                            }
                        }
                    */

                    //###??? CumulativeTime+={2} 暫時拿掉

                    //###??? and Sub_PP_Name暫時沒加 , 等PP_WorkOrder_Settlement的Sub_PP_Name正確
                    string err = "";
                    db.DB_SetData(string.Format(
                            @"UPDATE SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement SET 
                                                                        AvarageCycleTime={0},AvarageWaitTime={1},CumulativeTime+={2},TotalCheckIn={3},TotalCheckOut={4},
                                                                        TotalInput={5},TotalOutput={6},TotalFail={7},TotalKeep={8},
                                                                        FPY={9},YieldRate={10},StationYieldRate={11},UpdateTime=N'{12}' 
                                                                        WHERE ServerId='{16}' and OrderNO=N'{13}' AND StationNO=N'{14}' AND IndexSN={15}",
                                avgCT.ToString(),//0 
                                avgWT.ToString(),//1
                                Diffint.ToString(), //2
                                dataRow["_TotalCheckIn"].ToString(), //3
                                dataRow["_TotalCheckOut"].ToString(), //4 
                                dataRow["_TotalInput"].ToString(),//5 
                                dataRow["_TotalOutput"].ToString(),//6
                                dataRow["_TotalFail"].ToString(),//7
                                dataRow["_TotalKeep"].ToString(),//8 
                                dataRow["_FPY"].ToString(),//9
                                dataRow["_YieldRate"].ToString(),//10
                                dataRow["_StationYieldRate"].ToString(),//11 
                                DateTime.Now.ToString("MM/dd/yyyy H:mm:ss"),//12 
                                wo,//13 
                                dr_APS_Simulation["Source_StationNO"].ToString(), dr_APS_Simulation["Source_StationNO_IndexSN"].ToString(), _Fun.Config.ServerId),//14 15
                                ref err);

                }
                #endregion
            }
        }

        public bool Reporting_LabelWork(DBADO db, DataRow dr_StationNO, string UserNO, LabelWork keys, bool is_Back, ref string Message, ref string StackTrace, ref string ViewBagERRMsg)//單工報工
        {
            {
                //###???若此處有改 TMM Service 55 ㄝ要改
                DateTime startTime = DateTime.Now;
                //string meg = "";
                string sql = "";
                try
                {
                    string[] data = new string[7] { keys.Station, keys.OrderNO, keys.IndexSN, keys.OKQTY.ToString(), keys.FailQTY.ToString(), keys.OPNO, keys.LocalIPPort };
                    int outQTY = int.Parse(data[3]);//報工良品數量
                    int failQTY = int.Parse(data[4]);//報工不良品數量
                    DataRow dr_Manufacture = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[0]}'");
                    DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{keys.SimulationId}'");
                    DataRow tmp_dr = null;
                    #region 檢查網頁來源資料
                    if (!is_Back)
                    {
                        if (dr_Manufacture["OrderNO"].ToString() != keys.OrderNO || dr_Manufacture["IndexSN"].ToString() != keys.IndexSN)
                        {
                            ViewBagERRMsg = $"作業失敗, 畫面已逾時, 請關閉網頁瀏覽器, 重新刷條碼 並 重新操作."; return false;
                        }
                        if (keys.OKQTY <= 0 && keys.FailQTY <= 0) { keys.ERRMsg = $"報工數量合計不能為0, 請重新報工."; return false; }
                        if (dr_Manufacture.IsNull("StartTime")) { keys.ERRMsg = $"{data[1]} 工站沒有開工紀錄, 請先執行開工, 才能報工."; return false; }
                    }
                    else
                    {
                        dr_Manufacture.BeginEdit();
                        dr_Manufacture["OrderNO"] = keys.OrderNO;
                        dr_Manufacture["IndexSN"] = keys.IndexSN;
                        dr_Manufacture["PP_Name"] = dr_APS_Simulation["Apply_PP_Name"].ToString();
                        dr_Manufacture["Master_PP_Name"] = dr_APS_Simulation["Apply_PP_Name"].ToString();
                        dr_Manufacture["OP_NO"] = UserNO;
                        dr_Manufacture["SimulationId"] = keys.SimulationId;
                        dr_Manufacture["StartTime"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        dr_Manufacture["RemarkTimeS"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        dr_Manufacture["RemarkTimeE"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        dr_Manufacture["PartNO"] = dr_APS_Simulation["PartNO"].ToString();
                        dr_Manufacture["State"] = "2";
                        dr_Manufacture.EndEdit();
                    }
                    DataRow dr_WO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}'");
                    if (dr_WO == null) { keys.ERRMsg = $"查無 {data[0]} 工單資料紀錄, 請聯繫系統管理者."; return false; }
                    sql = $"SELECT * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}' AND StationNO='{data[0]}' AND IndexSN={data[2]}";
                    DataRow dr = db.DB_GetFirstDataByDataRow(sql);
                    if (dr == null && dr_APS_Simulation != null)
                    {
                        string isend = dr_APS_Simulation["PartSN"].ToString().Trim() == "0" ? "1" : "0";
                        string is_IndexSN_Merge = dr_APS_Simulation.IsNull("StationNO_Merge") ? "1" : "0";
                        string sfc_sql = $@"INSERT INTO SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement (
                                                                                    [OrderNO],[StationNO],[StationName],[PP_Name],[IndexSN],[DisplaySN],
                                                                                    [IsLastStation],[Sub_PP_Name],[StatndCycleTime],[UpdateTime],[IndexSN_Merge],
                                                                                    [StartTime],[CumulativeTime],[AvarageCycleTime],[TotalCheckIn],[TotalCheckOut],
                                                                                    [TotalInput],[TotalOutput],[TotalFail],[TotalKeep],[FPY],
                                                                                    [YieldRate],[StationYieldRate],ServerId) VALUES 
                                                                                    ('{data[1]}','{data[0]}','{dr_StationNO["StationName"].ToString()}','{dr_Manufacture["PP_Name"].ToString()}',
                                                                                    {data[2]},0,'{isend}','{dr_Manufacture["PP_Name"].ToString()}',{dr_APS_Simulation["Math_StandardCT"].ToString()},'{DateTime.Now.ToString("MM/dd/yyyy H:mm:ss")}',
                                                                                    '{is_IndexSN_Merge}',null,0,0,0,0,0,0,0,0,0,0,0,'{_Fun.Config.ServerId}')";

                        if (!db.DB_SetData(sfc_sql))
                        {
                            keys.ERRMsg = $"查無 PP_WorkOrder_Settlement 資料紀錄, 請聯繫系統管理者."; return false;
                        }
                        else
                        {
                            dr = db.DB_GetFirstDataByDataRow(sql);
                            if (dr == null) { keys.ERRMsg = $"查無 PP_WorkOrder_Settlement 資料紀錄, 請聯繫系統管理者."; return false; }
                        }
                    }
                    if (dr == null) { keys.ERRMsg = $"查無 PP_WorkOrder_Settlement 資料紀錄, 請聯繫系統管理者."; return false; }

                    if (!dr.IsNull("StartTime") && dr["StartTime"].ToString().Trim() != "")
                    { startTime = Convert.ToDateTime(dr["StartTime"]); }
                    #endregion

                    #region 計算CT
                    decimal ct = 0;
                    decimal ct_log = 0;
                    if (!is_Back)
                    {
                        DateTime rRemarkTimeS = startTime;
                        if (dr_Manufacture.IsNull("RemarkTimeS"))
                        { rRemarkTimeS = Convert.ToDateTime(dr_Manufacture["StartTime"]); }
                        else
                        { rRemarkTimeS = Convert.ToDateTime(dr_Manufacture["RemarkTimeS"]); }
                        #region 先查相同人員與站與工單, 是否報工過
                        tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and IndexSN={keys.IndexSN} and OrderNO='{keys.OrderNO}' and OperateType like '%報工%' and OP_NO like '%{keys.OPNO}%'");
                        if (tmp_dr != null && Convert.ToDateTime(tmp_dr["LOGDateTime"]) > rRemarkTimeS) { rRemarkTimeS = Convert.ToDateTime(tmp_dr["LOGDateTime"]); }
                        #endregion
                        if (dr_Manufacture.IsNull("RemarkTimeE") || rRemarkTimeS >= Convert.ToDateTime(dr_Manufacture["RemarkTimeE"]))
                        {
                            ct = TimeCompute2Seconds(rRemarkTimeS, DateTime.Now) / (keys.OKQTY + keys.FailQTY);
                            if (ct <= 0 && !dr_Manufacture.IsNull("RemarkTimeE") && Convert.ToDateTime(dr_Manufacture["RemarkTimeE"]) > Convert.ToDateTime(dr_Manufacture["RemarkTimeS"]))
                            {
                                ct = TimeCompute2Seconds(Convert.ToDateTime(dr_Manufacture["RemarkTimeS"]), Convert.ToDateTime(dr_Manufacture["RemarkTimeE"])) / (keys.OKQTY + keys.FailQTY);
                            }
                        }
                        else
                        {
                            ct = TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), rRemarkTimeS, Convert.ToDateTime(dr_Manufacture["RemarkTimeE"])) / (keys.OKQTY + keys.FailQTY);
                            if (ct <= 0 && Convert.ToDateTime(dr_Manufacture["RemarkTimeE"]) > rRemarkTimeS)
                            {
                                ct = TimeCompute2Seconds(rRemarkTimeS, Convert.ToDateTime(dr_Manufacture["RemarkTimeE"])) / (keys.OKQTY + keys.FailQTY);
                            }
                        }
                        ct_log = ct < 1 ? 0 : ct;

                        int ops = dr_Manufacture["OP_NO"].ToString().Split(';').Length;
                        if (ops > 1) { ct = ct / ops; }
                    }
                    #endregion

                    #region 寫SFC_StationDetail
                    string partNO = dr_APS_Simulation["PartNO"].ToString();
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
                        if (!is_Back)
                        {
                            if (int.Parse(dr_StationDetail["CycleTime"].ToString()) != 0) { ct = (ct + int.Parse(dr_StationDetail["CycleTime"].ToString())) > 0 ? Math.Round((ct + int.Parse(dr_StationDetail["CycleTime"].ToString())) / 2) : ct; }
                        }
                        else { ct = int.Parse(dr_StationDetail["CycleTime"].ToString()); }
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
                        string logtype = "智慧報工";
                        if (is_Back) { logtype = "干涉修正"; }
                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_APS_Simulation["NeedId"].ToString()}','{dr_APS_Simulation["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelWork','{logtype}','{dr_Manufacture["PP_Name"].ToString()}','{keys.Station}','{dr_Manufacture["PartNO"].ToString()}','{dr_Manufacture["OrderNO"].ToString()}','{UserNO}',{dr_Manufacture["IndexSN"].ToString()})");

                        #region 更新標籤累計量
                        if (!is_Back)
                        {
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
                        }
                        #endregion

                        #region log SFC_StationDetail_ChangeLOG紀錄
                        int reportTime = 0;
                        DataRow SFC_StationDetail_ChangeLOG = null;
                        if (dr_StationDetail != null)
                        {
                            #region 計算上一次與現在時間差
                            SFC_StationDetail_ChangeLOG = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(dr_StationDetail["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by LOGDateTime,LOGDateTimeID desc");
                            if (SFC_StationDetail_ChangeLOG != null)
                            {
                                if (!is_Back)
                                { reportTime = TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), Convert.ToDateTime(SFC_StationDetail_ChangeLOG["LOGDateTimeID"]), DateTime.Now); }
                            }
                            #endregion
                        }
                        string wsid = "NULL";
                        if (keys.SimulationId != "") { wsid = $"'{keys.SimulationId}'"; }
                        OP_Count = data[5].Split(";").Count();
                        if (OP_Count <= 0) { OP_Count = 1; }
                        string LOGDateTimeID = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                        string eCT = "0";
                        string upperCT = "0";
                        string lowerCT = "0";
                        if (!is_Back)
                        {
                            #region 查詢PP_EfficientDetail
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
                        }
                        else
                        {
                            if (SFC_StationDetail_ChangeLOG != null)
                            {
                                eCT = SFC_StationDetail_ChangeLOG["ECT"].ToString();
                                upperCT = SFC_StationDetail_ChangeLOG["ECT"].ToString();
                                lowerCT = SFC_StationDetail_ChangeLOG["ECT"].ToString();

                            }
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
                        }
                        #endregion

                        #region 計算效能 PP_EfficientDetail處理
                        if (!is_Back)
                        {
                            List<double> allCT = new List<double>();//list for all avg value
                            string top_flag = "";
                            try
                            {
                                if (_Fun.Config.AdminKey03 != 0) { top_flag = $" TOP {_Fun.Config.AdminKey03} "; }
                                DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} PP_Name,StationNO,PartNO as Sub_PartNO,CycleTime,WaitTime,(EditFinishedQty+EditFailedQty) as QTY from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG]
                                                    where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and PartNO='{dr_Manufacture["PartNO"].ToString()}' and PP_Name='{dr_WO["PP_Name"].ToString()}' and IndexSN={dr_Manufacture["IndexSN"].ToString()} and EditFinishedQty!=0 and CycleTime!=0");
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
                                        SfcTimerloopthread_Tick_Efficient(db, allCT, keys.Station, dr_WO["PP_Name"].ToString(), dr_Manufacture["PP_Name"].ToString(), dr_Manufacture["IndexSN"].ToString(), dr_WO["PartNO"].ToString(), dr_Manufacture["PartNO"].ToString(), "");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs 計算效能PP_EfficientDetail處理 Exception: {ex.Message} {ex.StackTrace}", true);
                            }
                        }
                        #endregion

                        #region 記錄刀工具使用時數
                        if (!is_Back)
                        {
                            DataTable tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStockII_Knives] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and IsDel='0'");
                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                            {
                                int useTime = 0;
                                int useCount = outQTY + failQTY;
                                string k_stime = "";
                                foreach (DataRow d in tmp_dt.Rows)
                                {
                                    if (!dr_Manufacture.IsNull("RemarkTimeS"))
                                    {
                                        if (d.IsNull("StartTime")) { k_stime = $",StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}'"; } else { k_stime = ""; }
                                        if (!dr_Manufacture.IsNull("RemarkTimeE"))
                                        {
                                            tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.Station}' and PartNO='{dr_Manufacture["PartNO"].ToString()}' and LOGDateTime>'{Convert.ToDateTime(dr_Manufacture["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}' and OperateType like '%報工%'");
                                            if (tmp_dr != null)
                                            {
                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                            ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.Station}','{dr_Manufacture["PartNO"].ToString()}',{(useCount).ToString()},0)");
                                            }
                                            else
                                            {
                                                useTime = TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), Convert.ToDateTime(dr_Manufacture["RemarkTimeS"].ToString()), DateTime.Now);
                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                            ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.Station}','{dr_Manufacture["PartNO"].ToString()}',{(useCount).ToString()},{useTime.ToString()})");
                                            }
                                        }
                                        else
                                        {
                                            useTime = TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), Convert.ToDateTime(dr_Manufacture["RemarkTimeS"].ToString()), DateTime.Now);
                                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                        ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.Station}','{dr_Manufacture["PartNO"].ToString()}',{(useCount).ToString()},{useTime.ToString()})");
                                        }
                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStockII_Knives] set TOTWorkTime+={useTime.ToString()},TOTCount+={useCount.ToString()}{k_stime} where ServerId='{_Fun.Config.ServerId}' and KId='{d["KId"].ToString()}'");
                                    }
                                }
                            }
                        }
                        #endregion

                        if (!is_Back)
                        {
                            #region 修正工站開始日期
                            if (dr_Manufacture["State"].ToString() == "1")
                            { db.DB_SetData($"update SoftNetMainDB.[dbo].[Manufacture] set RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[0]}'"); }
                            #endregion
                            Update_PP_WorkOrder_Settlement(db, data[1], keys.SimulationId);
                        }
                        //###??? 不良數量尚未處裡
                        if (keys.SimulationId != "")
                        {
                            bool isNeedQTY_OK = false;//判斷本站數量已足夠
                            string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();

                            string in_tmp_NO = "AC03";//###??? 暫時寫死 生產件,前站移轉不足,先領倉補
                            string out_tmp_NO = "BC03";//###??? 暫時寫死 生產件,前站移轉不足,先領倉補, 之後補報工退料
                            string in_NO = "AC01";//###??? 暫時寫死領料單別
                            string inOK_NO = "BC01";//###??? 暫時寫死入庫單別

                            DataRow dr_APS_PartNOTimeNote = null;
                            if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{data[1].Trim()}',Detail_QTY+={data[3]},Detail_Fail_QTY+={data[4]} where SimulationId='{keys.SimulationId}'"))
                            {
                                dr_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SimulationId}'");
                                if ((int.Parse(dr_APS_PartNOTimeNote["Detail_QTY"].ToString()) + int.Parse(dr_APS_PartNOTimeNote["Detail_Fail_QTY"].ToString()) - int.Parse(dr_APS_PartNOTimeNote["NeedQTY"].ToString())) >= 0)
                                { isNeedQTY_OK = true; }
                            }

                            //尋找相關BOM原物料
                            #region 扣Keep量 與 處理領料單單據
                            DataTable tmp_dt = null;
                            DataTable dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_APS_Simulation["NeedId"].ToString()}' and Apply_PP_Name='{dr_Manufacture["PP_Name"].ToString()}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and IndexSN={data[2]} order by PartSN desc");
                            if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                            {
                                string docNumberNO = "";
                                foreach (DataRow d in dt_APS_Simulation.Rows)
                                {
                                    #region 處裡工站移轉量 APS_PartNOTimeNote
                                    if (!d.IsNull("Source_StationNO") && (d["Class"].ToString() == "4" || d["Class"].ToString() == "5"))
                                    {
                                        if (d["PartSN"].ToString() == "0")
                                        {
                                            #region 工單最後一站預開入庫單  
                                            string tmp_no = "";
                                            string in_StoreNO = "";
                                            string in_StoreSpacesNO = "";
                                            tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 Order by b.StoreOrder");
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
                                                            Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, UserNO, true);
                                                            tmp_int = 0;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                            Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_01, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, UserNO, true);
                                                            tmp_int -= tmp_01;
                                                        }
                                                    }
                                                }
                                                if (tmp_int > 0)
                                                {
                                                    #region 計畫量不夠扣, 入實體倉
                                                    SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                                    Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, UserNO, true);
                                                    #endregion

                                                }
                                                #endregion
                                            }
                                            else
                                            {
                                                #region 無Keep量
                                                SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                                Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, outQTY, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, UserNO, true);
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
                                                    int NextStationTime = TimeCompute2Seconds_BY_Calendar(db, dr_StationNO["CalendarName"].ToString(), Convert.ToDateTime(tmp["LOGDateTime"]), DateTime.Now);
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
                                            #region 檢查上一階之前是否有入退庫數量(BC01)並領出
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
                                                            #region 計算單據CT,平均,有效, 寫SFC_StationProjectDetail, IsOK='1'
                                                            int typeTotalTime = 0;
                                                            string writeSQL = "";
                                                            if (!dr_DOC3stockII.IsNull("StartTime")) { typeTotalTime = TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC3stockII["StartTime"].ToString()), DateTime.Now); }
                                                            else { writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime}{writeSQL} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}'");
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
                                                                SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, efficient_pp_Name, efficient_pp_Name, "0", efficient_partNO, efficient_partNO, dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4));
                                                            }
                                                            #endregion
                                                        }
                                                        //開領料單,IsOK='1'
                                                        string tmpDOCNO = "";
                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                        Create_DOC3stock(db, d, dr_DOC3stockII["IN_StoreNO"].ToString(), dr_DOC3stockII["IN_StoreSpacesNO"].ToString(), "", "", in_NO, wrQTY, "", "", $"{stationno}站 入庫後再領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmpDOCNO, UserNO, true, true);
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

                                                #region 若上站有先領半成品AC03,且之後又補報工 開BC03(報工多餘退料)
                                                DataRow dr_BC03 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d["SimulationId"].ToString()}'");
                                                if (dr_BC03 != null && int.Parse(dr_BC03["Next_StationQTY"].ToString()) > 0)
                                                {
                                                    int store_tmp = int.Parse(dr_BC03["Detail_QTY"].ToString());
                                                    tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_tmp_NO}'");
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
                                                                else { store_tmp = tmp_AC03 - store_tmp; }
                                                                if (store_tmp > 0)
                                                                {
                                                                    string doc = "";
                                                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_tmp_NO}'");
                                                                    Create_DOC3stock(db, dr_BC03, "", "", tmp["OUT_StoreNO"].ToString(), tmp["OUT_StoreSpacesNO"].ToString(), out_tmp_NO, store_tmp, "", "", $"報工後,生產先領倉量退回 {keys.Station}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref doc, "系統指派", true, true);
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (tmp_AC03 > 0)
                                                            {
                                                                if (store_tmp > tmp_AC03) { store_tmp = tmp_AC03; }
                                                                else { store_tmp = tmp_AC03 - store_tmp; }
                                                                if (store_tmp > 0)
                                                                {
                                                                    string doc = "";
                                                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_tmp_NO}'");
                                                                    Create_DOC3stock(db, dr_BC03, "", "", tmp["OUT_StoreNO"].ToString(), tmp["OUT_StoreSpacesNO"].ToString(), out_tmp_NO, store_tmp, "", "", $"報工後,生產先領倉量退回 {keys.Station}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref doc, "系統指派", true, true);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion
                                            }
                                            else
                                            {
                                                bool is_run = true;
                                                //將在製移轉 剩餘開領料單
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

                                                #region 先檢查是否已有領料單據, 且已移轉多少量
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
                                                    tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 order by b.StoreOrder");
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
                                                                    Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_int, "", "", $"前站生產不足移轉,先領倉量補 {keys.Station}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO.Trim());
                                                                    tmp_int = 0;
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                    Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_01, "", "", $"前站生產不足移轉,先領倉量補 {keys.Station}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO.Trim());
                                                                    tmp_int -= tmp_01;
                                                                }
                                                            }
                                                        }
                                                        if (tmp_int > 0)
                                                        {
                                                            #region 計畫量不夠扣, 扣實體倉
                                                            DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow d2 in tmp_dt2.Rows)
                                                                {
                                                                    if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                    {
                                                                        //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_int} where Id='{d2["Id"].ToString()}'");
                                                                        Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_int, "", "", $"前站生產不足移轉,先領倉量補 {keys.Station}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
                                                                        tmp_int = 0;
                                                                        break;
                                                                    }
                                                                    else
                                                                    {
                                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                        //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_01} where Id='{d2["Id"].ToString()}'");
                                                                        if (tmp_01 != 0)
                                                                        {
                                                                            Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_01, "", "", $"前站生產不足移轉,先領倉量補 {keys.Station}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
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
                                                                SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                                #endregion
                                                                Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_tmp_NO, tmp_int, "", "", $"前站生產不足移轉,先領倉量補 {keys.Station}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO.Trim());
                                                            }
                                                            #endregion
                                                        }
                                                        #endregion
                                                    }
                                                    else
                                                    {
                                                        #region 沒計畫量, 扣實體倉
                                                        DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                        if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                        {
                                                            foreach (DataRow d2 in tmp_dt2.Rows)
                                                            {
                                                                if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                {
                                                                    Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_int, "", "", $"前站生產不足移轉,先領倉量補 {keys.Station}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
                                                                    tmp_int = 0;
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                    if (tmp_01 != 0)
                                                                    {
                                                                        Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_01, "", "", $"前站生產不足移轉,先領倉量補 {keys.Station}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
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
                                                            SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                            #endregion

                                                            #region 實體倉不購扣, 扣空倉
                                                            Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_tmp_NO, tmp_int, "", "", $"前站生產不足移轉,先領倉量補 {keys.Station}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
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
                                            tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 order by b.StoreOrder");
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
                                                            Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
                                                            tmp_int = 0;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                            Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
                                                            tmp_int -= tmp_01;
                                                        }
                                                    }
                                                }
                                                if (tmp_int > 0)
                                                {
                                                    #region 有計畫量不夠扣, 扣實體倉
                                                    DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                    if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                    {
                                                        foreach (DataRow d2 in tmp_dt2.Rows)
                                                        {
                                                            if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                            {
                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno} 有計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
                                                                tmp_int = 0;
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                if (tmp_01 != 0)
                                                                {
                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                    Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno} 有計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
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
                                                        SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                        #endregion
                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                        Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno} 計畫量不夠扣,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
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
                                                    DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                    if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                    {
                                                        foreach (DataRow d2 in tmp_dt2.Rows)
                                                        {
                                                            if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                            {
                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno} 沒計畫量,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
                                                                tmp_int = 0;
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                if (tmp_01 != 0)
                                                                {
                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                    Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno} 沒計畫量,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
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
                                                    SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                    #endregion
                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                    Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno} 沒計畫量,扣實體倉", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref docNumberNO, UserNO, true);
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
                                        string mFNO = SelectDOC4ProductionMFNO(db, tmp["PartNO"].ToString(), tmp["SimulationId"].ToString(), in_NO, ref price);
                                        #endregion
                                        #region 查找適合入庫儲別
                                        string in_StoreNO = "";
                                        string in_StoreSpacesNO = "";
                                        SelectINStore(db, tmp["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "PA02", true);
                                        #endregion

                                        if (Create_DOC4stock(db, tmp, mFNO, price, in_StoreNO, in_StoreSpacesNO, "PA02", docQTY, "", "", "工站報工,開下一站委外加工", tmp_down_StartTime, tmp_down_ArrivalDate, UserNO.Trim(), ref docNumberNO))
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
                            DataRow tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where SimulationId='{keys.SimulationId}' order by CalendarDate desc");
                            if (tmp_del != null && int.Parse(tmp_del["Time1_C"].ToString()) <= 1 && int.Parse(tmp_del["Time2_C"].ToString()) == 0 && int.Parse(tmp_del["Time3_C"].ToString()) == 0 && int.Parse(tmp_del["Time4_C"].ToString()) == 0)
                            { }
                            else
                            {
                                if (tmp_del != null)
                                {
                                    string stationNO_Merge = "";
                                    int delMath_UseTime = 0; int tmp_ct = 0; int tmp_wt = 0; int tmp_st = 0; int tmp_1 = 0; int tmp_2 = 0; int tmp_3 = 0; int tmp_4 = 0;
                                    tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SimulationId}'");
                                    if (tmp_del != null)
                                    {
                                        tmp_ct = int.Parse(tmp_del["Math_EfficientCT"].ToString());
                                        tmp_wt = int.Parse(tmp_del["Math_EfficientWT"].ToString());
                                        tmp_st = int.Parse(tmp_del["Math_StandardCT"].ToString());
                                        if ((tmp_ct + tmp_wt) != 0)
                                        { delMath_UseTime += (tmp_ct + tmp_wt) * (outQTY + failQTY); }
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

                                    tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where SimulationId='{keys.SimulationId}' order by CalendarDate desc");
                                    if (tmp_del != null && int.Parse(tmp_del["Time1_C"].ToString()) == 0 && int.Parse(tmp_del["Time2_C"].ToString()) == 0 && int.Parse(tmp_del["Time3_C"].ToString()) == 0 && int.Parse(tmp_del["Time4_C"].ToString()) == 0 && int.Parse(tmp_del["NeedQTY"].ToString()) > (int.Parse(tmp_del["Detail_QTY"].ToString()) + int.Parse(tmp_del["Detail_Fail_QTY"].ToString())))
                                    {
                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET Time1_C=1 where SimulationId='{keys.SimulationId}'");
                                    }
                                }
                            }
                            #endregion
                        }
                    }
                }
                catch (Exception ex)
                {
                    //meg = $"後臺錯誤: {ex.Message}";
                    Message = ex.Message;
                    StackTrace = ex.StackTrace;
                    //System.Threading.Tasks.Task task = _Log.ErrorAsync($"後臺錯誤: {ex.Message} {ex.StackTrace}", true);
                    return false;
                }
            }


            return true;
        }
        public bool Reporting_STView2Work(DBADO db, DataRow dr_PP_Station, string UserNO, MUTIStationObj keys, bool is_Back, ref string Message, ref string StackTrace)//多工報工
        {
            string sql = "";
            DataRow tmp_dr = null;
            try
            {
                DataRow dr_WO = null;
                DataRow dr_MII = null;
                #region 檢查
                if (!is_Back)
                {
                    if (keys.Select_ID == null || keys.SI_Slect_OPNOs == null || keys.Select_ID == "" || keys.SI_Slect_OPNOs == "") { keys.ERRMsg = $"報工前, 需先選擇項目 或 操作員."; return false; }
                    if (keys.SI_OKQTY == 0 && keys.SI_FailQTY == 0) { keys.ERRMsg = $"數量均無值, 請重新輸入."; return false; }
                    string id = "";
                    foreach (string s in keys.Select_ID.Split(';'))
                    {
                        if (s.Trim() != "")
                        {
                            if (id == "") { id = s; }
                            else { keys.ERRMsg = $"報工 一次只能選擇一個項目, 請重新選擇."; return false; }
                        }
                    }

                    dr_MII = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and Id='{id}' and EndTime is NULL");
                    if (dr_MII == null) { keys.ERRMsg = $"該項目已關閉生產, 無法報工."; return false; }
                    if (dr_MII.IsNull("StartTime")) { keys.ERRMsg = $"項目未設定開工過, 無法報工, 請先按開工."; return false; }
                    if (dr_MII.IsNull("RemarkTimeS")) { keys.ERRMsg = $"項目未設定開工過, 無法報工, 請先按開工."; return false; }
                    keys.SI_SimulationId = dr_MII["SimulationId"].ToString();
                    keys.SI_IndexSN = dr_MII["IndexSN"].ToString();
                    keys.SI_OrderNO = dr_MII["OrderNO"].ToString();
                    keys.SI_PP_Name = dr_MII["PP_Name"].ToString();
                    keys.SI_PartNO = dr_MII["PartNO"].ToString();
                }
                else
                {
                    dr_MII = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and SimulationId='{keys.SI_SimulationId}'");
                    if (dr_MII == null)
                    {
                        string stationNO_Custom_IndexSN = "";
                        string stationNO_Custom_DisplayName = "";
                        string id = _Str.NewId('C');
                        DataRow dr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{keys.SI_PP_Name}' and StationNO='{keys.StationNO}' and IndexSN={keys.SI_IndexSN}");
                        if (dr == null)
                        { keys.ERRMsg = $"查無相關製程順序.製程={keys.SI_PP_Name} 順序={keys.SI_IndexSN}, 請通知管理者."; return false; }
                        else
                        {
                            stationNO_Custom_IndexSN = dr["Station_Custom_IndexSN"].ToString();
                            stationNO_Custom_DisplayName = dr["DisplayName"].ToString();
                        }
                        string tmpTime = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff");
                        db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[ManufactureII] ([Id],[StationNO],[ServerId],[OrderNO],[Master_PP_Name],[PP_Name],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],[PartNO],[SimulationId],StartTime,EndTime,RemarkTimeS,RemarkTimeE)
                                              VALUES ('{id}','{keys.StationNO}','{_Fun.Config.ServerId}','{keys.SI_OrderNO}','{keys.SI_PP_Name}','{keys.SI_PP_Name}',{keys.SI_IndexSN},'{stationNO_Custom_IndexSN}','{stationNO_Custom_DisplayName}','{keys.SI_PartNO}','{keys.SI_SimulationId}','{tmpTime}','{tmpTime}','{tmpTime}','{tmpTime}')");
                        dr_MII = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and SimulationId='{keys.SI_SimulationId}'");
                        if (dr_MII == null)
                        { keys.ERRMsg = $"查無相關製程順序.製程={keys.SI_PP_Name} 順序={keys.SI_IndexSN}, 請通知管理者."; return false; }
                    }
                }
                #endregion

                DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{keys.SI_SimulationId}'");
                DateTime startTime = DateTime.Now;
                if (keys.SI_OKQTY > 0 || keys.SI_FailQTY > 0)
                {
                    string[] data = new string[7] { keys.StationNO, keys.SI_OrderNO, keys.SI_IndexSN, keys.SI_OKQTY.ToString(), keys.SI_FailQTY.ToString(), keys.SI_Slect_OPNOs, keys.LocalIPPort };

                    #region 檢查網頁來源資料
                    dr_WO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}'");
                    if (dr_WO == null) { keys.ERRMsg = $"查無 {data[0]} 工單資料紀錄."; return false; }
                    sql = $"SELECT * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='{data[1]}' AND StationNO='{data[0]}' AND IndexSN={data[2]}";
                    DataRow dr = db.DB_GetFirstDataByDataRow(sql);
                    if (dr == null && dr_APS_Simulation != null)
                    {
                        string isend = dr_APS_Simulation["PartSN"].ToString().Trim() == "0" ? "1" : "0";
                        string is_IndexSN_Merge = dr_APS_Simulation.IsNull("StationNO_Merge") ? "1" : "0";
                        string sfc_sql = $@"INSERT INTO SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement (
                                                                                    [OrderNO],[StationNO],[StationName],[PP_Name],[IndexSN],[DisplaySN],
                                                                                    [IsLastStation],[Sub_PP_Name],[StatndCycleTime],[UpdateTime],[IndexSN_Merge],
                                                                                    [StartTime],[CumulativeTime],[AvarageCycleTime],[TotalCheckIn],[TotalCheckOut],
                                                                                    [TotalInput],[TotalOutput],[TotalFail],[TotalKeep],[FPY],
                                                                                    [YieldRate],[StationYieldRate],ServerId) VALUES 
                                                                                    ('{data[1]}','{data[0]}','{dr_PP_Station["StationName"].ToString()}','{dr_MII["PP_Name"].ToString()}',
                                                                                    {data[2]},0,'{isend}','{dr_MII["PP_Name"].ToString()}',{dr_APS_Simulation["Math_StandardCT"].ToString()},'{DateTime.Now.ToString("MM/dd/yyyy H:mm:ss")}',
                                                                                    '{is_IndexSN_Merge}',null,0,0,0,0,0,0,0,0,0,0,0,'{_Fun.Config.ServerId}')";

                        if (!db.DB_SetData(sfc_sql))
                        {
                            keys.ERRMsg = $"查無 PP_WorkOrder_Settlement 資料紀錄, 請聯繫系統管理者."; return false;
                        }
                        else
                        {
                            dr = db.DB_GetFirstDataByDataRow(sql);
                            if (dr == null) { keys.ERRMsg = $"查無 PP_WorkOrder_Settlement 資料紀錄, 請聯繫系統管理者."; return false; }
                        }
                    }
                    if (dr == null) { keys.ERRMsg = $"查無 PP_WorkOrder_Settlement 資料紀錄, 請聯繫系統管理者."; return false; }
                    if (dr != null && !dr.IsNull("StartTime") && dr["StartTime"].ToString().Trim() != "")
                    { startTime = Convert.ToDateTime(dr["StartTime"]); }
                    #endregion

                    int outQTY = keys.SI_OKQTY;
                    int failQTY = keys.SI_FailQTY;
                    decimal ct = 0;
                    decimal ct_log = 0;
                    DateTime rRemarkTimeS = Convert.ToDateTime(dr_MII["RemarkTimeS"]);

                    #region 計算CT

                    //檢查rRemarkTimeS是否需要更改適合的時間, 查上一次報工的時間
                    string select_OP = "";
                    if (!is_Back)
                    {
                        if (keys.SI_Slect_OPNOs.Split(';').Length > 1)
                        {
                            foreach (string s in keys.SI_Slect_OPNOs.Split(';'))
                            {
                                if (select_OP == "") { select_OP = $"OP_NO like '%{s}%'"; }
                                else { select_OP = $"{select_OP} or OP_NO like '%{s}%'"; }
                            }
                            select_OP = $"({select_OP})";
                        }
                        else { select_OP = $"OP_NO like '%{keys.SI_Slect_OPNOs}%'"; }
                        tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IndexSN={keys.SI_IndexSN} and OrderNO='{keys.SI_OrderNO}' and OperateType like '%報工%' and {select_OP}  order by LOGDateTime desc");
                        if (tmp_dr != null && Convert.ToDateTime(tmp_dr["LOGDateTime"]) > rRemarkTimeS)
                        {
                            rRemarkTimeS = Convert.ToDateTime(tmp_dr["LOGDateTime"]);
                        }
                        if (dr_MII.IsNull("RemarkTimeE") || rRemarkTimeS >= Convert.ToDateTime(dr_MII["RemarkTimeE"]))
                        {
                            ct = TimeCompute2Seconds(rRemarkTimeS, DateTime.Now) / (keys.SI_OKQTY + keys.SI_FailQTY);
                            if (ct <= 0 && !dr_MII.IsNull("RemarkTimeE") && Convert.ToDateTime(dr_MII["RemarkTimeE"]) > Convert.ToDateTime(dr_MII["RemarkTimeS"]))
                            {
                                ct = TimeCompute2Seconds(Convert.ToDateTime(dr_MII["RemarkTimeS"]), Convert.ToDateTime(dr_MII["RemarkTimeE"])) / (keys.SI_OKQTY + keys.SI_FailQTY);
                            }
                        }
                        else
                        {
                            ct = TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), rRemarkTimeS, Convert.ToDateTime(dr_MII["RemarkTimeE"])) / (keys.SI_OKQTY + keys.SI_FailQTY);
                            if (ct <= 0 && Convert.ToDateTime(dr_MII["RemarkTimeE"]) > rRemarkTimeS)
                            {
                                ct = TimeCompute2Seconds(rRemarkTimeS, Convert.ToDateTime(dr_MII["RemarkTimeE"])) / (keys.SI_OKQTY + keys.SI_FailQTY);
                            }
                        }
                        ct_log = ct < 1 ? 0 : ct;


                        int ops = keys.SI_Slect_OPNOs.Split(';').Length;
                        if (ops > 1) { ct = ct / ops; }
                    }
                    #endregion

                    #region 寫SFC_StationDetail
                    string logTime = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff");
                    string lOGDateTime = logTime;
                    string old_InTime = "";
                    string old_OutTime = "";
                    string old_ProductFinishedQty = "0";
                    string old_ProductFailedQty = "0";
                    int OP_Count = data[5].Split(';').Count() + 1;
                    if (OP_Count <= 0) { OP_Count = 1; }
                    DataRow dr_StationDetail = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail] WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO = '{data[1]}' AND StationNO = '{data[0]}' AND IndexSN={data[2]}");
                    if (dr_StationDetail == null)
                    {
                        //###???PP_Name暫時
                        old_InTime = startTime.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        old_OutTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
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
                            data[5],
                            data[0],
                            data[2],
                            dr["IndexSN_Merge"].ToString(),
                            data[1],
                            dr_MII["PartNO"].ToString(),
                            old_InTime,
                            old_OutTime,
                            ct.ToString(),
                            (outQTY + failQTY) > 0 ? 1 : 0,
                            outQTY > 0 ? 1 : 0,
                            failQTY > 0 ? 1 : 0,
                            dr_PP_Station["Station_Type"].ToString(),
                            outQTY.ToString(),
                            failQTY.ToString(), dr_PP_Station["RMSName"].ToString(), keys.SI_SimulationId, _Fun.Config.ServerId);
                    }
                    else
                    {
                        logTime = Convert.ToDateTime(dr_StationDetail["LOGDateTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff");
                        old_InTime = Convert.ToDateTime(dr_StationDetail["InTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff");
                        old_OutTime = Convert.ToDateTime(dr_StationDetail["OutTime"]).ToString("MM/dd/yyyy HH:mm:ss.fff");
                        old_ProductFinishedQty = dr_StationDetail["ProductFinishedQty"].ToString();
                        old_ProductFailedQty = dr_StationDetail["ProductFailedQty"].ToString();
                        if (!is_Back)
                        {
                            if (int.Parse(dr_StationDetail["CycleTime"].ToString()) != 0) { ct = (ct + int.Parse(dr_StationDetail["CycleTime"].ToString())) > 0 ? Math.Round((ct + int.Parse(dr_StationDetail["CycleTime"].ToString())) / 2) : ct; }
                        }
                        if (ct < 1) { ct = 0; }

                        sql = string.Format(
                                @"UPDATE SoftNetLogDB.[dbo].[SFC_StationDetail] 
                                                SET [ProductFinishedQty]+={0}, [ProductFailedQty]+={1},
                                                [InTime]='{2}',[OutTime]='{3}',[CycleTime]={4} 
                                                WHERE ServerId='{9}' and OrderNO = '{5}' AND StationNO = '{6}' AND IndexSN={7} AND LOGDateTime = '{8}'",
                                outQTY,
                                failQTY,
                                startTime.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                                ct.ToString(),
                                data[1],
                                data[0], data[2], logTime, _Fun.Config.ServerId);
                    }
                    #endregion

                    if (db.DB_SetData(sql))
                    {
                        #region log SFC_StationDetail_ChangeLOG紀錄
                        int reportTime = 0;
                        DataRow SFC_StationDetail_ChangeLOG = null;
                        if (dr_StationDetail != null)
                        {
                            #region 計算上一次報工與現在時間差
                            SFC_StationDetail_ChangeLOG = db.DB_GetFirstDataByDataRow($"SELECT LOGDateTimeID FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and LOGDateTime='{Convert.ToDateTime(dr_StationDetail["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by LOGDateTime,LOGDateTimeID desc");
                            if (SFC_StationDetail_ChangeLOG != null && !is_Back)
                            { reportTime = TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(SFC_StationDetail_ChangeLOG["LOGDateTimeID"]), DateTime.Now); }
                            #endregion
                        }
                        string wsid = "NULL";
                        if (keys.SI_SimulationId != "") { wsid = $"'{keys.SI_SimulationId}'"; }
                        string partNO = keys.SI_PartNO;
                        string LOGDateTimeID = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                        string eCT = "0";
                        string upperCT = "0";
                        string lowerCT = "0";
                        if (!is_Back)
                        {
                            #region 取得PP_EfficientDetail紀錄
                            DataRow dr_tmp_ct = db.DB_GetFirstDataByDataRow($"select AVG(EfficientCycleTime) as ECT,AVG(SD_UpperLimit) as UpperCT,AVG(SD_LowerLimit) as LowerCT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StationNO='{data[0]}' and PP_Name='{dr["PP_Name"].ToString()}' and IndexSN={dr["IndexSN"].ToString()} and DOCNO=''");
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
                                string _s = "";
                            }
                        }
                        else
                        {
                            if (SFC_StationDetail_ChangeLOG != null)
                            {
                                eCT = SFC_StationDetail_ChangeLOG["ECT"].ToString();
                                upperCT = SFC_StationDetail_ChangeLOG["ECT"].ToString();
                                lowerCT = SFC_StationDetail_ChangeLOG["ECT"].ToString();

                            }
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
                        }
                        #endregion

                        #region 計算效能 PP_EfficientDetail處理
                        if (!is_Back)
                        {
                            List<double> allCT = new List<double>();//list for all avg value
                            string top_flag = "";
                            try
                            {
                                if (_Fun.Config.AdminKey03 != 0) { top_flag = $" TOP {_Fun.Config.AdminKey03} "; }
                                DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} PP_Name,StationNO,PartNO as Sub_PartNO,CycleTime,WaitTime,(EditFinishedQty+EditFailedQty) as QTY from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG]
                                                    where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PartNO='{keys.SI_PartNO}' and PP_Name='{dr_WO["PP_Name"].ToString()}' and IndexSN={data[2]} and EditFinishedQty!=0 and CycleTime!=0");
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
                                        SfcTimerloopthread_Tick_Efficient(db, allCT, keys.StationNO, dr_WO["PP_Name"].ToString(), keys.SI_PP_Name, data[2], dr_WO["PartNO"].ToString(), keys.SI_PartNO, "");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Threading.Tasks.Task task = _Log.ErrorAsync($"STView2WorkController.cs 計算效能PP_EfficientDetail處理 Exception: {ex.Message} {ex.StackTrace}", true);
                            }
                        }
                        #endregion

                        #region 記錄刀工具使用時數
                        if (!is_Back)
                        {
                            DataTable tmp_dt = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStockII_Knives] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and IsDel='0'");
                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                            {
                                int useTime = 0;
                                int useCount = outQTY + failQTY;
                                string k_stime = "";
                                foreach (DataRow d in tmp_dt.Rows)
                                {
                                    if (!dr_MII.IsNull("RemarkTimeS"))
                                    {
                                        if (d.IsNull("StartTime")) { k_stime = $",StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}'"; } else { k_stime = ""; }
                                        if (!dr_MII.IsNull("RemarkTimeE"))
                                        {
                                            tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{keys.StationNO}' and PartNO='{dr_MII["PartNO"].ToString()}' and LOGDateTime>'{Convert.ToDateTime(dr_MII["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}' and OperateType like '%報工%'");
                                            if (tmp_dr != null)
                                            {
                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                            ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}',{(useCount).ToString()},0)");
                                            }
                                            else
                                            {
                                                useTime = TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(dr_MII["RemarkTimeS"].ToString()), DateTime.Now);
                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                            ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}',{(useCount).ToString()},{useTime.ToString()})");
                                            }
                                        }
                                        else
                                        {
                                            useTime = TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(dr_MII["RemarkTimeS"].ToString()), DateTime.Now);
                                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII_Knives_LOG] ([Id],[KId],[ServerId],[LOGDateTime],[StationNO],[PartNO],[WorkQTY],[WorkTime]) VALUES 
                                                        ('{_Str.NewId('K')}','{d["KId"].ToString()}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}',{(useCount).ToString()},{useTime.ToString()})");
                                        }
                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStockII_Knives] set TOTWorkTime+={useTime.ToString()},TOTCount+={useCount.ToString()}{k_stime} where ServerId='{_Fun.Config.ServerId}' and KId='{d["KId"].ToString()}'");
                                    }
                                }
                            }
                        }
                        #endregion

                        if (!is_Back)
                        {
                            Update_PP_WorkOrder_Settlement(db, data[1], keys.SI_SimulationId);
                        }
                        //###??? 不良數量尚未處裡
                        if (keys.SI_SimulationId != "")
                        {
                            string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();
                            bool isNeedQTY_OK = false;//判斷本站數量已足夠
                            string in_tmp_NO = "AC03";//###??? 暫時寫死 生產件,前站移轉不足,先領倉補
                            string out_tmp_NO = "BC03";//###??? 暫時寫死 生產件,前站移轉不足,先領倉補, 之後補報工退料
                            string in_NO = "AC01";//###??? 暫時寫死領料單別
                            string inOK_NO = "BC01";//###??? 暫時寫死入庫單別

                            //本階報工當筆資料寫入
                            DataRow dr_APS_PartNOTimeNote = null;
                            if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{data[1].Trim()}',Detail_QTY+={data[3]},Detail_Fail_QTY+={data[4]} where SimulationId='{keys.SI_SimulationId}'"))
                            {
                                dr_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SI_SimulationId}'");
                                if (dr_APS_PartNOTimeNote == null)
                                {
                                    keys.ERRMsg = $"程式有問題, 請聯繫系統管理員.";
                                    Message = $"查無此語法資料 select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SI_SimulationId}'";
                                    StackTrace = "";
                                    return false;
                                }
                                else
                                {
                                    if ((int.Parse(dr_APS_PartNOTimeNote["Detail_QTY"].ToString()) + int.Parse(dr_APS_PartNOTimeNote["Detail_Fail_QTY"].ToString()) - int.Parse(dr_APS_PartNOTimeNote["NeedQTY"].ToString())) >= 0)
                                    { isNeedQTY_OK = true; }
                                }
                            }

                            //尋找相關BOM原物料
                            #region 由本階工站,查上一階 扣Keep量 與 處理領料單單據 與若上階有先入庫,要先領出
                            DataTable tmp_dt = null;
                            DataTable dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_APS_Simulation["NeedId"].ToString()}' and Apply_PP_Name='{keys.SI_PP_Name}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and IndexSN={data[2]} order by PartSN desc");
                            if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
                            {
                                string docNumberNO = "";
                                foreach (DataRow d in dt_APS_Simulation.Rows)
                                {
                                    #region 上一階是工站, 處裡移轉量 APS_PartNOTimeNote
                                    if (!d.IsNull("Source_StationNO") && (d["Class"].ToString() == "4" || d["Class"].ToString() == "5"))
                                    {
                                        if (d["PartSN"].ToString() == "0")
                                        {
                                            #region 工單最後一站預開入庫單  
                                            string tmp_no = "";
                                            string in_StoreNO = "";
                                            string in_StoreSpacesNO = "";
                                            tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 Order by b.StoreOrder");
                                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                            {
                                                int tmp_int = outQTY;
                                                #region 有計畫先扣Keep量  先by StoreOrder順序扣, 在預開入庫單 
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
                                                            Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmp_no, UserNO, true);
                                                            tmp_int = 0;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                            Create_DOC3stock(db, d, "", "", d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), inOK_NO, tmp_01, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmp_no, UserNO, true);
                                                            tmp_int -= tmp_01;
                                                        }
                                                    }
                                                }
                                                if (tmp_int > 0)
                                                {
                                                    #region 計畫量不夠扣, 入實體倉
                                                    SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                                    Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, tmp_int, "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref tmp_no, UserNO, true, true);
                                                    #endregion
                                                }
                                                #endregion
                                            }
                                            else
                                            {
                                                #region 查找適合庫儲別
                                                SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, inOK_NO);
                                                #endregion
                                                #region 無倉紀錄, 加空倉
                                                Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, inOK_NO, int.Parse(data[3]), "", "", $"工單:{data[1]} 入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmp_no, UserNO, true);
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
                                                    int NextStationTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(tmp["LOGDateTime"]), DateTime.Now);
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
                                                            if (!dr_DOC3stockII.IsNull("StartTime")) { typeTotalTime = TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC3stockII["StartTime"].ToString()), DateTime.Now); }
                                                            else { writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime}{writeSQL} where Id='{dr_DOC3stockII["Id"].ToString()}' and DOCNumberNO='{dr_DOC3stockII["DOCNumberNO"].ToString()}'");
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
                                                                SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, efficient_pp_Name, efficient_pp_Name, "0", efficient_partNO, efficient_partNO, dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4));
                                                            }
                                                            #endregion
                                                        }
                                                        //開領料單,IsOK='1'
                                                        string tmpDOCNO = "";
                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                        Create_DOC3stock(db, d, dr_DOC3stockII["IN_StoreNO"].ToString(), dr_DOC3stockII["IN_StoreSpacesNO"].ToString(), "", "", in_NO, wrQTY, "", "", $"{stationno}站 入庫後再領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref tmpDOCNO, UserNO, true, true);
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
                                            DataRow dr_next_APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{keys.SI_SimulationId}'");
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
                                                //先在製移轉 
                                                sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Next_APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}',Next_StationQTY+={(tmp_int + next_next_detail_QTY).ToString()} where SimulationId='{d["SimulationId"].ToString()}'";
                                                if (db.DB_SetData(sql))
                                                {
                                                    #region 處理工站移轉時間
                                                    /*
                                                    tmp = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetLogDB.[dbo].[SFC_StationDetail] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["Source_StationNO"].ToString()}' and IndexSN='{d["Source_StationNO_IndexSN"].ToString()}' order by LOGDateTime desc");
                                                    if (tmp != null)
                                                    {
                                                        int NextStationTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(tmp["LOGDateTime"]), DateTime.Now);
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

                                                #region 若上站有先領半成品AC03,且之後又補報工 開BC03(報工多餘退料)
                                                DataRow dr_BC03 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d["SimulationId"].ToString()}'");
                                                if (dr_BC03 != null && int.Parse(dr_BC03["Next_StationQTY"].ToString()) > 0)
                                                {
                                                    int store_tmp = int.Parse(dr_BC03["Detail_QTY"].ToString());
                                                    tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_tmp_NO}'");
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
                                                                else { store_tmp = tmp_AC03 - store_tmp; }
                                                                if (store_tmp > 0)
                                                                {
                                                                    string doc = "";
                                                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_tmp_NO}'");
                                                                    Create_DOC3stock(db, dr_BC03, "", "", tmp["OUT_StoreNO"].ToString(), tmp["OUT_StoreSpacesNO"].ToString(), out_tmp_NO, store_tmp, "", "", $"報工後,生產先領倉量退回 {keys.StationNO}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref doc, "系統指派", true, true);
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (tmp_AC03 > 0)
                                                            {
                                                                if (store_tmp > tmp_AC03) { store_tmp = tmp_AC03; }
                                                                else { store_tmp = tmp_AC03 - store_tmp; }
                                                                if (store_tmp > 0)
                                                                {
                                                                    string doc = "";
                                                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_tmp_NO}'");
                                                                    Create_DOC3stock(db, dr_BC03, "", "", tmp["OUT_StoreNO"].ToString(), tmp["OUT_StoreSpacesNO"].ToString(), out_tmp_NO, store_tmp, "", "", $"報工後,生產先領倉量退回 {keys.StationNO}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;LabelWorkController", ref doc, "系統指派", true, true);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion

                                            }
                                            else
                                            {
                                                bool is_run = true;
                                                //將在製 剩餘移轉完
                                                tmp_int -= (detail_QTY - next_StationQTY);
                                                sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Next_APS_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}',Next_StationQTY={tmp["Detail_QTY"].ToString()} where SimulationId='{d["SimulationId"].ToString()}'";
                                                if (db.DB_SetData(sql) && int.Parse(tmp["Detail_QTY"].ToString()) > 0)
                                                {
                                                    #region 處理工站移轉時間
                                                    /*
                                                    tmp = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetLogDB.[dbo].[SFC_StationDetail] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["Source_StationNO"].ToString()}' and IndexSN='{d["Source_StationNO_IndexSN"].ToString()}' order by LOGDateTime desc");
                                                    if (tmp != null)
                                                    {
                                                        int NextStationTime = _WebSocket.TimeCompute2Seconds_BY_Calendar(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(tmp["LOGDateTime"]), DateTime.Now);
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

                                                #region 先檢查是否已有相關單據(已領過AC01), 且已移轉多少量
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
                                                    tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 order by b.StoreOrder");
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
                                                                    Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_int, "", "", $"前站生產不足移轉,先領倉量補 {keys.StationNO}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
                                                                    tmp_int = 0;
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                                    //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_01} where Id='{d2["Id"].ToString()}'");
                                                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                    Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_01, "", "", $"前站生產不足移轉,先領倉量補 {keys.StationNO}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
                                                                    tmp_int -= tmp_01;
                                                                }
                                                            }
                                                        }
                                                        if (tmp_int > 0)
                                                        {
                                                            #region 計畫量不夠扣, 扣實體倉
                                                            DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow d2 in tmp_dt2.Rows)
                                                                {
                                                                    if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                    {
                                                                        //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_int} where Id='{d2["Id"].ToString()}'");
                                                                        Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_int, "", "", $"前站生產不足移轉,先領倉量補 {keys.StationNO}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
                                                                        tmp_int = 0;
                                                                        break;
                                                                    }
                                                                    else
                                                                    {
                                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                        //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={tmp_01} where Id='{d2["Id"].ToString()}'");
                                                                        if (tmp_01 != 0)
                                                                        {
                                                                            Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_01, "", "", $"前站生產不足移轉,先領倉量補 {keys.StationNO}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
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
                                                                SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                                #endregion
                                                                Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_tmp_NO, tmp_int, "", "", $"前站生產不足移轉,先領倉量補 {keys.StationNO}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
                                                            }
                                                            #endregion
                                                        }
                                                        #endregion
                                                    }
                                                    else
                                                    {
                                                        #region 沒計畫量, 扣實體倉
                                                        DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                        if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                        {
                                                            foreach (DataRow d2 in tmp_dt2.Rows)
                                                            {
                                                                if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                                {
                                                                    Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_int, "", "", $"前站生產不足移轉,先領倉量補 {keys.StationNO}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
                                                                    tmp_int = 0;
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                    if (tmp_01 != 0)
                                                                    {
                                                                        Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_tmp_NO, tmp_01, "", "", $"前站生產不足移轉,先領倉量補 {keys.StationNO}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
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
                                                            SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                            #endregion

                                                            #region 實體倉不購扣, 扣空倉
                                                            Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_tmp_NO, tmp_int, "", "", $"前站生產不足移轉,先領倉量補 {keys.StationNO}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
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
                                        #region 上一階是原物料 扣庫存帳 TotalStock,TotalStockII
                                        int tmp_int = (int.Parse(d["BOMQTY"].ToString()) * outQTY);

                                        #region 先檢查是否已有單據, 且已移轉多少量
                                        int detailQTY = tmp_int;
                                        int stockQTY = 0;
                                        DataRow tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO!=''");
                                        if (tmp != null)
                                        {
                                            docNumberNO = tmp["DOCNumberNO"].ToString();
                                            detailQTY += (int.Parse(tmp["Detail_QTY"].ToString()) + int.Parse(tmp["Next_StationQTY"].ToString()) - int.Parse(tmp["Next_StoreQTY"].ToString()));
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
                                            tmp_dt = db.DB_GetData($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}' and b.KeepQTY>0 order by b.StoreOrder");
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
                                                            Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
                                                            tmp_int = 0;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                            db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                            string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                            Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
                                                            tmp_int -= tmp_01;
                                                        }
                                                    }
                                                }
                                                if (tmp_int > 0)
                                                {
                                                    #region 有計畫量不夠扣, 扣實體倉
                                                    DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                    if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                    {
                                                        foreach (DataRow d2 in tmp_dt2.Rows)
                                                        {
                                                            if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                            {
                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
                                                                tmp_int = 0;
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                if (tmp_01 != 0)
                                                                {
                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                    Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
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
                                                        SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                        #endregion
                                                        string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                        Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
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
                                                    DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                    if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                    {
                                                        foreach (DataRow d2 in tmp_dt2.Rows)
                                                        {
                                                            if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                            {
                                                                string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
                                                                tmp_int = 0;
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                                if (tmp_01 != 0)
                                                                {
                                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                                    Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", in_NO, tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
                                                                    tmp_int -= tmp_01;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                if (tmp_int > 0)
                                                {
                                                    #region 查找適合庫儲別
                                                    string out_StoreNO = "";
                                                    string out_StoreSpacesNO = "";
                                                    SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, in_NO);
                                                    #endregion
                                                    #region 實體倉不購扣, 扣空倉
                                                    string stationno = !d.IsNull("Source_StationNO") ? $"工站:{d["Source_StationNO"].ToString()} " : "";
                                                    Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", in_NO, tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SetAction;STView2WorkController", ref docNumberNO, UserNO, true);
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
                                        string mFNO = SelectDOC4ProductionMFNO(db, tmp["PartNO"].ToString(), tmp["SimulationId"].ToString(), in_NO, ref price);
                                        #endregion
                                        #region 查找適合入庫儲別
                                        string in_StoreNO = "";
                                        string in_StoreSpacesNO = "";
                                        SelectINStore(db, tmp["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "PA02", true);
                                        #endregion

                                        if (Create_DOC4stock(db, tmp, mFNO, price, in_StoreNO, in_StoreSpacesNO, "PA02", docQTY, "", "", "工站報工,開下一站委外加工", tmp_down_StartTime, tmp_down_ArrivalDate, UserNO, ref docNumberNO))
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
                            DataRow tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where SimulationId='{keys.SI_SimulationId}' order by CalendarDate desc");
                            if (tmp_del != null && int.Parse(tmp_del["Time1_C"].ToString()) <= 1 && int.Parse(tmp_del["Time2_C"].ToString()) == 0 && int.Parse(tmp_del["Time3_C"].ToString()) == 0 && int.Parse(tmp_del["Time4_C"].ToString()) == 0)
                            { }
                            else
                            {
                                if (tmp_del != null)
                                {
                                    string stationNO_Merge = "";
                                    int delMath_UseTime = 0; int tmp_ct = 0; int tmp_wt = 0; int tmp_st = 0; int tmp_1 = 0; int tmp_2 = 0; int tmp_3 = 0; int tmp_4 = 0;
                                    tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SI_SimulationId}'");
                                    if (tmp_del != null)
                                    {
                                        tmp_ct = int.Parse(tmp_del["Math_EfficientCT"].ToString());
                                        tmp_wt = int.Parse(tmp_del["Math_EfficientWT"].ToString());
                                        tmp_st = int.Parse(tmp_del["Math_StandardCT"].ToString());
                                        if ((tmp_ct + tmp_wt) != 0)
                                        { delMath_UseTime += (tmp_ct + tmp_wt) * (outQTY + failQTY); }
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
                                    dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SI_SimulationId}' and StationNO='{data[0]}' order by CalendarDate");
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
                                        dt_APS_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{dr_WO["NeedId"].ToString()}' and SimulationId='{keys.SI_SimulationId}' and StationNO {stationNO_Merge} order by CalendarDate");
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

                                    tmp_del = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where SimulationId='{keys.SI_SimulationId}' order by CalendarDate desc");
                                    if (tmp_del != null && int.Parse(tmp_del["Time1_C"].ToString()) == 0 && int.Parse(tmp_del["Time2_C"].ToString()) == 0 && int.Parse(tmp_del["Time3_C"].ToString()) == 0 && int.Parse(tmp_del["Time4_C"].ToString()) == 0 && int.Parse(dr_APS_PartNOTimeNote["NeedQTY"].ToString()) > (int.Parse(dr_APS_PartNOTimeNote["Detail_QTY"].ToString()) + int.Parse(dr_APS_PartNOTimeNote["Detail_Fail_QTY"].ToString())))
                                    {
                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET Time1_C=1 where SimulationId='{keys.SI_SimulationId}'");
                                    }
                                }
                            }
                            #endregion
                        }

                        keys.SI_FailQTY = 0;
                        keys.SI_OKQTY = 0;

                        keys.SI_PP_Name = "";
                        keys.SI_PartName = "";
                        keys.SI_IndexSN = "";
                        keys.SI_OrderNO = "";
                        keys.SI_PartNO = "";
                        keys.MES_String = "完成報工作業.";
                        string logtype = "智慧報工";
                        if (is_Back) { logtype = "干涉修正"; }
                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_APS_Simulation["NeedId"].ToString()}','{dr_APS_Simulation["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','STView2Work','{logtype}','{dr_MII["PP_Name"].ToString()}','{keys.StationNO}','{dr_MII["PartNO"].ToString()}','{dr_MII["OrderNO"].ToString()}','{keys.SI_Slect_OPNOs}',{dr_MII["IndexSN"].ToString()})");
                        keys.SI_OPNO = "";
                        keys.SI_Slect_OPNOs = "";
                    }
                }
            }
            catch (Exception ex)
            {
                keys.ERRMsg = $"程式有問題, 請聯繫系統管理員.";
                Message = ex.Message;
                StackTrace = ex.StackTrace;
                return false;
            }
            return true;
        }
        public bool Create_WorkOrder(DBADO db, string needID, string calendarName, ref string ERR)
        {
            DataRow tmp = null;
            DataRow tmp2 = null;
            #region 確認Manufacture資料存在
            string sql = @$"select A.Source_StationNO from SoftNetSYSDB.[dbo].[APS_Simulation] as A where A.NeedId='{needID}' group by A.Source_StationNO";
            DataTable dt_APS_Simulation = db.DB_GetData(sql);
            if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
            {
                foreach (DataRow d in dt_APS_Simulation.Rows)
                {
                    if (!d.IsNull("Source_StationNO") && d["Source_StationNO"].ToString() != "")
                    {
                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["Source_StationNO"].ToString()}'");
                        if (tmp == null)
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["Source_StationNO"].ToString()}'");
                            if (tmp != null)
                            {
                                if (tmp != null && tmp["Station_Type"].ToString() == "1")
                                { db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Manufacture] (ServerId,StationNO,Config_MutiWO) VALUES ('{_Fun.Config.ServerId}','{d["Source_StationNO"].ToString()}','0')"); }
                                else { db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Manufacture] (ServerId,StationNO,Config_MutiWO) VALUES ('{_Fun.Config.ServerId}','{d["Source_StationNO"].ToString()}','1')"); }
                            }
                            else
                            { ERR = $"無法正式轉排程, 因查無{d["Source_StationNO"].ToString()}工站資料, 請通知系統管理者."; return false; }
                        }
                    }
                }
            }
            #endregion

            #region 產生工單
            dt_APS_Simulation = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and IsOK='0' and PartSN>=0 order by NeedId,Apply_PP_Name,PartSN desc");
            if (dt_APS_Simulation != null && dt_APS_Simulation.Rows.Count > 0)
            {
                DataRow dr_MBOM = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where Id='{dt_APS_Simulation.Rows[0]["Apply_BOMId"].ToString()}'");
                string partNO = dr_MBOM["Apply_PartNO"].ToString();

                string pp_Name = dt_APS_Simulation.Rows[0]["Apply_PP_Name"].ToString();
                DateTime wo_StartDate = Convert.ToDateTime(dt_APS_Simulation.Rows[0]["SimulationDate"]);
                int partSN = int.Parse(dt_APS_Simulation.Rows[0]["PartSN"].ToString());
                bool changPP = false;
                foreach (DataRow d in dt_APS_Simulation.Rows)
                {
                    //單工單
                    if (changPP && pp_Name == d["Apply_PP_Name"].ToString()) { continue; }
                    else { changPP = false; }
                    if (pp_Name != d["Apply_PP_Name"].ToString())
                    {
                        pp_Name = d["Apply_PP_Name"].ToString();
                        partSN = int.Parse(d["PartSN"].ToString());
                        wo_StartDate = Convert.ToDateTime(d["SimulationDate"]);
                        continue;
                    }
                    else
                    {
                        if (partSN != int.Parse(d["PartSN"].ToString()))
                        {
                            //insert wo
                            string date = DateTime.Now.ToString("yyyyMMdd");
                            int sno = 1;
                            int qty = 0;
                            string partName = "";
                            tmp = db.DB_GetFirstDataByDataRow($"select OrderNO from SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO like 'XX01{date}%' order by OrderNO desc");
                            if (tmp != null)
                            {
                                try
                                {
                                    sno = int.Parse(tmp["OrderNO"].ToString().Substring(12)) + 1;
                                }
                                catch
                                { }
                            }
                            string orderNO = $"XX01{date}{sno.ToString().PadLeft(4, '0')}";
                            tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_ProductProcess where ServerId='{_Fun.Config.ServerId}' and PP_Name='{pp_Name}' and Apply_PartNO='{partNO}'");
                            if (tmp == null) { ERR = $"需求碼:{needID} 查無製程資料"; return false; }
                            tmp2 = db.DB_GetFirstDataByDataRow($"select top 1 a.Master_PartNO,b.PartName,a.PartSN from SoftNetSYSDB.[dbo].APS_Simulation as a,SoftNetMainDB.[dbo].[Material] as b where a.NeedId='{needID}' and a.Apply_PP_Name='{pp_Name}' and a.Master_PartNO=b.PartNO and a.PartSN>=0 order by a.PartNO");
                            if (int.Parse(tmp2["PartSN"].ToString()) != 1 && int.Parse(tmp2["PartSN"].ToString()) != 0)
                            {
                                tmp2 = db.DB_GetFirstDataByDataRow($"select a.PartNO,b.PartName,(a.NeedQTY+a.SafeQTY-a.Math_TotalStock_HasUseQTY-a.Math_Online_SurplusQTY) as qty from SoftNetSYSDB.[dbo].APS_Simulation as a,SoftNetMainDB.[dbo].[Material] as b where a.NeedId='{needID}' and a.PartNO='{tmp2["Master_PartNO"].ToString()}' and a.PartNO=b.PartNO");
                                qty = int.Parse(tmp2["qty"].ToString());
                                if (qty < 0) { qty = 0; }
                                partNO = tmp2["PartNO"].ToString();
                                partName = tmp2["PartName"].ToString();
                            }
                            else
                            {
                                tmp2 = db.DB_GetFirstDataByDataRow($"select a.PartNO,b.PartName,(c.NeedQTY+c.SafeQTY-c.Math_TotalStock_HasUseQTY-c.Math_Online_SurplusQTY) as qty,a.CalendarName from SoftNetSYSDB.[dbo].APS_NeedData as a,SoftNetMainDB.[dbo].[Material] as b,SoftNetSYSDB.[dbo].APS_Simulation as c where a.ServerId='{_Fun.Config.ServerId}' and a.Id='{needID}' and a.PartNO=b.PartNO and c.NeedId=a.Id and a.PartNO=c.PartNO and c.PartSN=0");
                                qty = int.Parse(tmp2["qty"].ToString());
                                partNO = tmp2["PartNO"].ToString();
                                partName = tmp2["PartName"].ToString();
                            }
                            sql = $"INSERT INTO SoftNetSYSDB.[dbo].[PP_WorkOrder] (ServerId,OrderNO,FactoryName,LineName,PP_Name,PartNO,PartName,Quantity,CalendarName,EstimatedStartTime,NeedId) VALUES ('{_Fun.Config.ServerId}','{orderNO}','{tmp["FactoryName"].ToString()}','{tmp["LineName"].ToString()}','{pp_Name}','{partNO}','{partName}','{qty.ToString()}','{calendarName}','{wo_StartDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needID}')";//
                            if (db.DB_SetData(sql))
                            {
                                sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET DOCNumberNO='{orderNO}' where NeedId='{needID}' and Apply_PP_Name='{pp_Name}' and IsOK='0' and (Class='4' or Class='5') and OutPackType='0' and Source_StationNO is not null and (NeedQTY+SafeQTY-Math_TotalStock_HasUseQTY-Math_Online_SurplusQTY)>0";
                                db.DB_SetData(sql);
                                DataTable dt_wn = db.DB_GetData($"SELECT SimulationId FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Apply_PP_Name='{pp_Name}' and Source_StationNO is NOT NULL and OutPackType='0'");
                                string wnid = "";
                                foreach (DataRow d2 in dt_wn.Rows)
                                {
                                    if (wnid == "") { wnid = $"'{d2["SimulationId"].ToString()}'"; }
                                    else { wnid = $"{wnid},'{d2["SimulationId"].ToString()}'"; }
                                }
                                sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET DOCNumberNO='{orderNO}' where SimulationId in ({wnid})";
                                db.DB_SetData(sql);
                            }
                            changPP = true;
                            continue;
                        }
                        if (wo_StartDate < Convert.ToDateTime(d["SimulationDate"]))
                        { wo_StartDate = Convert.ToDateTime(d["SimulationDate"]); }
                    }

                    //###???多工單 還沒寫
                    //sql = $"select a.FactoryName,a.LineName,a.PP_Name,b.PartNO,b.NeedQTY,b.SimulationDate,b.NeedId from SoftNetSYSDB.[dbo].PP_ProductProcess as a join SoftNetSYSDB.[dbo].APS_Simulation as b";
                    //sql = $"{sql} on a.PP_Name = b.Apply_PP_Name and a.PP_Name = '{d["Apply_PP_Name"].ToString()}' and b.Id = '{d["SimulationId"].ToString()}'";
                    //tmp = db.DB_GetFirstDataByDataRow("sql");
                }

            }
            else
            {
                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='9',StateINFO='無APS_Simulation資料,無生產行為',UpdateTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and Id='{needID}'"))
                {
                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] SET IsOK='1' where NeedId='{needID}'");
                    db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needID}'");
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needID}'");
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}'");
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{needID}' and ErrorType!='06' and ErrorType!='07' and ErrorType!='08' and ErrorType!='09' and ErrorType!='10'");
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{needID}'");
                    ERR = "無APS_Simulation資料,無生產行為";
                    return false;
                }
            }

            #endregion

            return true;
        }

        public bool Create_DOC1stock(DBADO db, DataRow d, string mfNO, float price, string in_StoreNO, string in_StoreSpacesNO, string docNO, int qty, string beforeDOCNumberNO, string beforeId, string remark, string startDate, string arrivalDate, string userID, ref string docNumberNO)
        {
            bool isrun = false;
            DataRow tmp = null;
            string partNO = d["PartNO"].ToString();
            string needId = d["NeedId"].ToString().Trim();
            string funName = "";
            #region 確認 TotalStock檔 有對應資料
            if (in_StoreNO != "")
            {
                if (in_StoreSpacesNO != "")
                { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{in_StoreNO}' and StoreSpacesNO='{in_StoreSpacesNO}'"); }
                else
                { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{in_StoreNO}'"); }
                if (tmp == null)
                {
                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{in_StoreNO}'");
                    if (tmp != null)
                    {
                        db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (Class,ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{tmp["Class"].ToString()}','{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{in_StoreNO}','{in_StoreSpacesNO}','{partNO}',0)");
                    }
                }
            }
            #endregion
            if (docNumberNO != "" && docNO != docNumberNO.Substring(0, 4)) { docNumberNO = ""; }
            string simulationID = d["SimulationId"].ToString().Trim();
            DataRow dr_DOCRole = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOCRole] where DOCNO='{docNO}' and ServerId='{_Fun.Config.ServerId}'");
            if (dr_DOCRole == null) { return false; }
            string docType = dr_DOCRole["DOCType"].ToString();
            //確認mfNO與單號與SID關析
            tmp = db.DB_GetFirstDataByDataRow($"SELECT a.* FROM SoftNetMainDB.[dbo].[DOC1BuyII] as a,SoftNetMainDB.[dbo].[DOC1Buy] as b where SUBSTRING(a.DOCNumberNO,1,4)='{docNO}' and a.SimulationId='{simulationID}' and a.DOCNumberNO=b.DOCNumberNO and b.MFNO='{mfNO}'");
            if (tmp == null)
            {

                //沒單據
                #region 取得單號
                if (docNumberNO == "")
                {
                    int sno = 0;
                    string date = DateTime.Now.ToString("yyyyMMdd");
                    tmp = db.DB_GetFirstDataByDataRow($"select top 1 DOCNumberNO from SoftNetMainDB.[dbo].[DOC1Buy] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO like '{docNO}{date}%' order by DOCNumberNO desc");
                    if (tmp != null)
                    {
                        try
                        {
                            sno = int.Parse(tmp["DOCNumberNO"].ToString().Substring(12));
                        }
                        catch
                        {
                            string _s = "";
                        }
                    }
                    docNumberNO = $"{docNO}{date}{(++sno).ToString().PadLeft(4, '0')}";
                    //###??? sql 暫時寫死
                    if (db.DB_SetData($"INSERT INTO [dbo].[DOC1Buy] (ServerId,[DOCNumberNO],[DOCNO],[DOCDate],[UserId],[DOCType],[SourceNO],[FlowLevel],[FlowStatus],[AIFalg],MFNO) VALUES ('{_Fun.Config.ServerId}','{docNumberNO}','{docNO}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}','{userID}','{docType}','{beforeDOCNumberNO}',0,'Y','0','{mfNO}')"))
                    {
                        if (beforeId == "") { beforeId = _Str.NewId('Z'); }//用來源單據明細的Id
                        if (db.DB_SetData($"INSERT INTO [dbo].[DOC1BuyII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[StoreNO],[StoreSpacesNO],ArrivalDate) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',0,'PCS',{qty.ToString()},'{remark}','{simulationID}','0','{in_StoreNO}','{in_StoreSpacesNO}','{arrivalDate}')"))
                        { isrun = true; }
                    }
                }
                else
                {
                    if (beforeId == "") { beforeId = _Str.NewId('Z'); }
                    if (db.DB_SetData($"INSERT INTO [dbo].[DOC1BuyII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[StoreNO],[StoreSpacesNO],ArrivalDate) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',0,'PCS',{qty.ToString()},'{remark}','{simulationID}','0','{in_StoreNO}','{in_StoreSpacesNO}','{arrivalDate}')"))
                    { isrun = true; }
                }
                #endregion
            }
            else
            {
                if (simulationID == "")
                {
                    string date = DateTime.Now.ToString("yyyyMMdd");
                    DataRow tmp02 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOC1Buy] where SUBSTRING(DOCNumberNO,1,4)='{docNO}' and SUBSTRING(DOCNumberNO,5,8)='{date}' and UserId='{userID}'");
                    if (tmp02 == null)
                    {
                        int sno = 0;
                        tmp02 = db.DB_GetFirstDataByDataRow($"select top 1 DOCNumberNO from SoftNetMainDB.[dbo].[DOC1Buy] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO like '{docNO}{date}%' order by DOCNumberNO desc");
                        if (tmp02 != null)
                        {
                            try
                            {
                                sno = int.Parse(tmp02["DOCNumberNO"].ToString().Substring(12));
                            }
                            catch
                            {
                                string _s = "";
                            }
                        }
                        docNumberNO = $"{docNO}{date}{(++sno).ToString().PadLeft(4, '0')}";

                        db.DB_SetData($"INSERT INTO [dbo].[DOC1Buy] (ServerId,[DOCNumberNO],[DOCNO],[DOCDate],[UserId],[DOCType],[SourceNO],[FlowLevel],[FlowStatus],[AIFalg]) VALUES ('{_Fun.Config.ServerId}','{docNumberNO}','{docNO}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}','{userID}','{docType}','{beforeDOCNumberNO}',0,'Y','0')");
                    }
                }

                //有單據
                if (docNumberNO == "") { docNumberNO = tmp["DOCNumberNO"].ToString(); }

                tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOC1BuyII] where IsOK='0' and SimulationId='{simulationID}' and SUBSTRING(DOCNumberNO,1,4)='{docNO}' and StoreNO='{in_StoreNO}' and StoreSpacesNO='{in_StoreSpacesNO}'");
                if (tmp != null)
                {
                    if (Convert.ToDateTime(tmp["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss") != arrivalDate)
                    {
                        if (beforeId == "") { beforeId = _Str.NewId('Z'); }
                        if (db.DB_SetData($"INSERT INTO [dbo].[DOC1BuyII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[StoreNO],[StoreSpacesNO],StartTime,ArrivalDate) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',{price},'PCS',{qty.ToString()},'{remark}','{simulationID}','0','{in_StoreNO}','{in_StoreSpacesNO}','{startDate}','{arrivalDate}')"))
                        { isrun = true; }
                    }
                    else
                    {
                        if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC1BuyII] set QTY+={qty.ToString()} where Id='{tmp["Id"].ToString()}'"))
                        { isrun = true; }
                    }
                }
                else
                {
                    if (beforeId == "") { beforeId = _Str.NewId('Z'); }
                    if (db.DB_SetData($"INSERT INTO [dbo].[DOC1BuyII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[StoreNO],[StoreSpacesNO],StartTime,ArrivalDate) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',{price},'PCS',{qty.ToString()},'{remark}','{simulationID}','0','{in_StoreNO}','{in_StoreSpacesNO}','{startDate}','{arrivalDate}')"))
                    { isrun = true; }
                }
            }
            if (isrun)
            {
                string[] tmp_array = funName.Split(";");
                string proName = "";
                if (tmp_array.Length >= 2) { proName = tmp_array[1]; }
                db.DB_SetData($"INSERT INTO [dbo].[DOCUpdateLog] ([Id],[LOGDateTime],[DOCId],[DOCNumberNO],[NeedId],[SimulationId],[PartNO],[QTY],[FunName],[ProName]) VALUES ('{_Str.NewId('0')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}','{beforeId}','{docNumberNO}','{needId}','{simulationID}','{partNO}',{qty.ToString()},'{tmp_array[0]}','{proName}')");
                if (remark != "")
                {
                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId,SimulationId) VALUES ('開立{dr_DOCRole["DOCName"].ToString()} {remark}  料號:{partNO} 數量:{qty.ToString()}','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationID}')");
                }
            }
            else
            { db.DB_SetData($"INSERT INTO [dbo].[DOCUpdateLog] ([Id],[LOGDateTime],[DOCId],[DOCNumberNO],[NeedId],[SimulationId],[PartNO],[QTY],[FunName],[ProName]) VALUES ('{_Str.NewId('0')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}','{beforeId}','{docNumberNO}','{needId}','{simulationID}','{partNO}',{qty.ToString()},'程式錯誤','{funName}')"); }

            return true;
        }

        public bool Create_DOC4stock(DBADO db, DataRow d, string mfNO, float price, string in_StoreNO, string in_StoreSpacesNO, string docNO, int qty, string beforeDOCNumberNO, string beforeId, string remark, string startDate, string arrivalDate, string userID, ref string docNumberNO)
        {
            bool isrun = false;
            DataRow tmp = null;
            string partNO = d["PartNO"].ToString();
            string needId = d["NeedId"].ToString().Trim();
            if (startDate == "") { startDate = "NULL"; }
            else { startDate = $"'{startDate}'"; }
            string funName = "";
            #region 確認 TotalStock檔 有對應資料
            if (in_StoreNO != "")
            {
                if (in_StoreSpacesNO != "")
                { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{in_StoreNO}' and StoreSpacesNO='{in_StoreSpacesNO}'"); }
                else
                { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{in_StoreNO}'"); }
                if (tmp == null)
                {
                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{in_StoreNO}'");
                    if (tmp != null)
                    {
                        db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (Class,ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{tmp["Class"].ToString()}','{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{in_StoreNO}','{in_StoreSpacesNO}','{partNO}',0)");
                    }
                }
            }
            #endregion
            if (docNumberNO != "" && docNO != docNumberNO.Substring(0, 4)) { docNumberNO = ""; }
            string simulationID = d["SimulationId"].ToString().Trim();
            DataRow dr_DOCRole = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOCRole] where DOCNO='{docNO}' and ServerId='{_Fun.Config.ServerId}'");
            if (dr_DOCRole == null) { return false; }
            string docType = dr_DOCRole["DOCType"].ToString();
            //確認mfNO與單號與SID關析
            tmp = db.DB_GetFirstDataByDataRow($"SELECT a.* FROM SoftNetMainDB.[dbo].[DOC4ProductionII] as a,SoftNetMainDB.[dbo].[DOC4Production] as b where SUBSTRING(a.DOCNumberNO,1,4)='{docNO}' and a.SimulationId='{simulationID}' and a.DOCNumberNO=b.DOCNumberNO and b.MFNO='{mfNO}'");
            if (tmp == null)
            {
                if (mfNO == "") { return false; }
                //沒單據
                #region 取得單號
                if (docNumberNO == "")
                {
                    int sno = 0;
                    string date = DateTime.Now.ToString("yyyyMMdd");
                    tmp = db.DB_GetFirstDataByDataRow($"select top 1 DOCNumberNO from SoftNetMainDB.[dbo].[DOC4Production] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO like '{docNO}{date}%' order by DOCNumberNO desc");
                    if (tmp != null)
                    {
                        try
                        {
                            sno = int.Parse(tmp["DOCNumberNO"].ToString().Substring(12));
                        }
                        catch
                        {
                            string _s = "";
                        }
                    }
                    docNumberNO = $"{docNO}{date}{(++sno).ToString().PadLeft(4, '0')}";
                    //###??? sql 暫時寫死
                    if (db.DB_SetData($"INSERT INTO [dbo].[DOC4Production] (ServerId,[DOCNumberNO],[DOCNO],[DOCDate],[UserId],[DOCType],[SourceNO],[FlowLevel],[FlowStatus],[AIFalg],MFNO) VALUES ('{_Fun.Config.ServerId}','{docNumberNO}','{docNO}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}','{userID}','{docType}','{beforeDOCNumberNO}',0,'Y','0','{mfNO}')"))
                    {
                        if (beforeId == "") { beforeId = _Str.NewId('Z'); }//用來源單據明細的Id
                        if (db.DB_SetData($"INSERT INTO [dbo].[DOC4ProductionII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[StoreNO],[StoreSpacesNO],ArrivalDate,StartTime) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',0,'PCS',{qty.ToString()},'{remark}','{simulationID}','0','{in_StoreNO}','{in_StoreSpacesNO}','{arrivalDate}',{startDate})"))
                        { isrun = true; }
                    }
                }
                else
                {
                    if (beforeId == "") { beforeId = _Str.NewId('Z'); }
                    if (db.DB_SetData($"INSERT INTO [dbo].[DOC4ProductionII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[StoreNO],[StoreSpacesNO],ArrivalDate,StartTime) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',0,'PCS',{qty.ToString()},'{remark}','{simulationID}','0','{in_StoreNO}','{in_StoreSpacesNO}','{arrivalDate}',{startDate})"))
                    { isrun = true; }
                }
                #endregion
            }
            else
            {
                if (simulationID == "")
                {
                    string date = DateTime.Now.ToString("yyyyMMdd");
                    DataRow tmp02 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOC4Production] where SUBSTRING(DOCNumberNO,1,4)='{docNO}' and SUBSTRING(DOCNumberNO,5,8)='{date}' and UserId='{userID}'");
                    if (tmp02 == null)
                    {
                        int sno = 0;
                        tmp02 = db.DB_GetFirstDataByDataRow($"select top 1 DOCNumberNO from SoftNetMainDB.[dbo].[DOC4Production] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO like '{docNO}{date}%' order by DOCNumberNO desc");
                        if (tmp02 != null)
                        {
                            try
                            {
                                sno = int.Parse(tmp02["DOCNumberNO"].ToString().Substring(12));
                            }
                            catch
                            {
                                string _s = "";
                            }
                        }
                        docNumberNO = $"{docNO}{date}{(++sno).ToString().PadLeft(4, '0')}";

                        db.DB_SetData($"INSERT INTO [dbo].[DOC4Production] (ServerId,[DOCNumberNO],[DOCNO],[DOCDate],[UserId],[DOCType],[SourceNO],[FlowLevel],[FlowStatus],[AIFalg]) VALUES ('{_Fun.Config.ServerId}','{docNumberNO}','{docNO}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}','{userID}','{docType}','{beforeDOCNumberNO}',0,'Y','0')");
                    }
                }
                //有單據
                if (docNumberNO == "") { docNumberNO = tmp["DOCNumberNO"].ToString(); }
                tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOC4ProductionII] where IsOK='0' and SimulationId='{simulationID}' and SUBSTRING(DOCNumberNO,1,4)='{docNO}' and StoreNO='{in_StoreNO}' and StoreSpacesNO='{in_StoreSpacesNO}'");
                if (tmp != null)
                {
                    if (Convert.ToDateTime(tmp["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss") != arrivalDate)
                    {
                        if (beforeId == "") { beforeId = _Str.NewId('Z'); }
                        if (db.DB_SetData($"INSERT INTO [dbo].[DOC4ProductionII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[StoreNO],[StoreSpacesNO],StartTime,ArrivalDate) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',{price},'PCS',{qty.ToString()},'{remark}','{simulationID}','0','{in_StoreNO}','{in_StoreSpacesNO}',{startDate},'{arrivalDate}')"))
                        { isrun = true; }
                    }
                    else
                    {
                        string startTime = "";
                        if (startDate.ToLower() != "null" && tmp.IsNull("StartTime"))
                        { startTime = $",StartTime={startDate}"; }
                        if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC4ProductionII] set QTY+={qty.ToString()}{startTime} where Id='{tmp["Id"].ToString()}'"))
                        { isrun = true; }
                    }
                }
                else
                {
                    if (beforeId == "") { beforeId = _Str.NewId('Z'); }
                    if (db.DB_SetData($"INSERT INTO [dbo].[DOC4ProductionII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[StoreNO],[StoreSpacesNO],StartTime,ArrivalDate) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',{price},'PCS',{qty.ToString()},'{remark}','{simulationID}','0','{in_StoreNO}','{in_StoreSpacesNO}',{startDate},'{arrivalDate}')"))
                    { isrun = true; }
                }
            }
            if (isrun)
            {
                string[] tmp_array = funName.Split(";");
                string proName = "";
                if (tmp_array.Length >= 2) { proName = tmp_array[1]; }
                db.DB_SetData($"INSERT INTO [dbo].[DOCUpdateLog] ([Id],[LOGDateTime],[DOCId],[DOCNumberNO],[NeedId],[SimulationId],[PartNO],[QTY],[FunName],[ProName]) VALUES ('{_Str.NewId('0')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}','{beforeId}','{docNumberNO}','{needId}','{simulationID}','{partNO}',{qty.ToString()},'{tmp_array[0]}','{proName}')");
                if (remark != "")
                {
                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId,SimulationId) VALUES ('開立{dr_DOCRole["DOCName"].ToString()} {remark} 料號:{partNO} 數量:{qty.ToString()}','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationID}')");
                }
            }
            else
            { db.DB_SetData($"INSERT INTO [dbo].[DOCUpdateLog] ([Id],[LOGDateTime],[DOCId],[DOCNumberNO],[NeedId],[SimulationId],[PartNO],[QTY],[FunName],[ProName]) VALUES ('{_Str.NewId('0')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}','{beforeId}','{docNumberNO}','{needId}','{simulationID}','{partNO}',{qty.ToString()},'程式錯誤','{funName}')"); }

            return true;
        }


        public bool Create_DOC3stock(DBADO db, DataRow d, string out_StoreNO, string out_StoreSpacesNO, string in_StoreNO, string in_StoreSpacesNO, string docNO, int qty, string beforeDOCNumberNO, ref string beforeId, string remark, string arrivalDate, string funName, ref string docNumberNO, string userID, bool write_StartTime = true, bool isOK = false)//處理領料單單據 DOC3stock,DOC3stockII
        {
            return Create_DOC3stock_II(db, d, out_StoreNO, out_StoreSpacesNO, in_StoreNO, in_StoreSpacesNO, docNO, qty, beforeDOCNumberNO, ref beforeId, remark, arrivalDate, funName, ref docNumberNO, write_StartTime, userID, isOK);
        }
        public bool Create_DOC3stock(DBADO db, DataRow d, string out_StoreNO, string out_StoreSpacesNO, string in_StoreNO, string in_StoreSpacesNO, string docNO, int qty, string beforeDOCNumberNO, string beforeId, string remark, string arrivalDate, string funName, ref string docNumberNO, string userID, bool write_StartTime = true, bool isOK = false)//處理領料單單據 DOC3stock,DOC3stockII
        {
            return Create_DOC3stock_II(db, d, out_StoreNO, out_StoreSpacesNO, in_StoreNO, in_StoreSpacesNO, docNO, qty, beforeDOCNumberNO, ref beforeId, remark, arrivalDate, funName, ref docNumberNO, write_StartTime, userID, isOK);
        }

        private bool Create_DOC3stock_II(DBADO db, DataRow d, string out_StoreNO, string out_StoreSpacesNO, string in_StoreNO, string in_StoreSpacesNO, string docNO, int qty, string beforeDOCNumberNO, ref string beforeId, string remark, string arrivalDate, string funName, ref string docNumberNO, bool write_StartTime, string userID, bool isOK = false)//處理領料單單據 DOC3stock,DOC3stockII
        {
            bool isrun = false;
            string needId = "";
            string simulationID = "";
            string partNO = "";
            try
            {
                partNO = d["PartNO"].ToString();
                DataRow tmp = null;

                #region 確認 TotalStock檔 有對應資料
                if (in_StoreNO != "")
                {
                    if (in_StoreSpacesNO != "")
                    { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{in_StoreNO}' and StoreSpacesNO='{in_StoreSpacesNO}'"); }
                    else
                    { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{in_StoreNO}'"); }
                    if (tmp == null)
                    {
                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{in_StoreNO}'");
                        if (tmp != null)
                        {
                            db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (Class,ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{tmp["Class"].ToString()}','{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{in_StoreNO}','{in_StoreSpacesNO}','{partNO}',0)");
                        }
                    }
                }
                else if (out_StoreNO != "")
                {
                    if (out_StoreSpacesNO != "")
                    { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{out_StoreNO}' and StoreSpacesNO='{out_StoreSpacesNO}'"); }
                    else
                    { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{out_StoreNO}'"); }
                    if (tmp == null)
                    {
                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{out_StoreNO}'");
                        if (tmp != null)
                        {
                            db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (Class,ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{tmp["Class"].ToString()}','{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{out_StoreNO}','{out_StoreSpacesNO}','{partNO}',0)");
                        }
                    }
                }
                #endregion

                needId = d["NeedId"].ToString().Trim();
                simulationID = d["SimulationId"].ToString().Trim();
                string startTime = "NULL";
                string endTime = "NULL";
                string docOK = "0";
                if (write_StartTime) { startTime = $"'{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}'"; }
                if (isOK)
                {
                    docOK = "1";
                    endTime = $"'{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}'";
                }
                if (docNumberNO != "" && docNO != docNumberNO.Substring(0, 4)) { docNumberNO = ""; }
                DataRow dr_DOCRole = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOCRole] where DOCNO='{docNO}' and ServerId='{_Fun.Config.ServerId}'");
                if (dr_DOCRole == null) { return false; }
                string docType = dr_DOCRole["DOCType"].ToString();
                tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where SUBSTRING(DOCNumberNO,1,4)='{docNO}' and SimulationId='{simulationID}'");
                if (tmp == null)
                {
                    //沒相同SID
                    #region 取得單號
                    if (docNumberNO == "")
                    {
                        int sno = 0;
                        string date = DateTime.Now.ToString("yyyyMMdd");
                        tmp = db.DB_GetFirstDataByDataRow($"select top 1 DOCNumberNO from SoftNetMainDB.[dbo].[DOC3stock] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO like '{docNO}{date}%' order by DOCNumberNO desc");
                        if (tmp != null)
                        {
                            try
                            {
                                sno = int.Parse(tmp["DOCNumberNO"].ToString().Substring(12));
                            }
                            catch
                            {
                                string _s = "";
                            }
                        }
                        docNumberNO = $"{docNO}{date}{(++sno).ToString().PadLeft(4, '0')}";
                        //###??? sql 暫時寫死
                        if (db.DB_SetData($"INSERT INTO [dbo].[DOC3stock] (ServerId,[DOCNumberNO],[DOCNO],[DOCDate],[UserId],[DOCType],[SourceNO],[FlowLevel],[FlowStatus],[AIFalg]) VALUES ('{_Fun.Config.ServerId}','{docNumberNO}','{docNO}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}','{userID}','{docType}','{beforeDOCNumberNO}',0,'Y','0')"))
                        {
                            if (beforeId == "") { beforeId = _Str.NewId('Z'); }//用來源單據明細的Id
                            if (db.DB_SetData($"INSERT INTO [dbo].[DOC3stockII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],ArrivalDate,StartTime,EndTime,IsOK) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',0,'PCS',{qty.ToString()},'{remark}','{simulationID}','{in_StoreNO}','{in_StoreSpacesNO}','{out_StoreNO}','{out_StoreSpacesNO}','{arrivalDate}',{startTime},{endTime},'{docOK}')"))
                            { isrun = true; }
                        }
                    }
                    else
                    {
                        if (beforeId == "") { beforeId = _Str.NewId('Z'); }
                        if (db.DB_SetData($"INSERT INTO [dbo].[DOC3stockII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],ArrivalDate,StartTime,EndTime,IsOK) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',0,'PCS',{qty.ToString()},'{remark}','{simulationID}','{in_StoreNO}','{in_StoreSpacesNO}','{out_StoreNO}','{out_StoreSpacesNO}','{arrivalDate}',{startTime},{endTime},'{docOK}')"))
                        { isrun = true; }
                    }
                    #endregion
                }
                else
                {
                    if (simulationID == "")
                    {
                        string date = DateTime.Now.ToString("yyyyMMdd");
                        DataRow tmp02 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stock] where SUBSTRING(DOCNumberNO,1,4)='{docNO}' and SUBSTRING(DOCNumberNO,5,8)='{date}' and UserId='{userID}'");
                        if (tmp02 == null)
                        {
                            int sno = 0;
                            tmp02 = db.DB_GetFirstDataByDataRow($"select top 1 DOCNumberNO from SoftNetMainDB.[dbo].[DOC3stock] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO like '{docNO}{date}%' order by DOCNumberNO desc");
                            if (tmp02 != null)
                            {
                                try
                                {
                                    sno = int.Parse(tmp02["DOCNumberNO"].ToString().Substring(12));
                                }
                                catch
                                {
                                    string _s = "";
                                }
                            }
                            docNumberNO = $"{docNO}{date}{(++sno).ToString().PadLeft(4, '0')}";

                            db.DB_SetData($"INSERT INTO [dbo].[DOC3stock] (ServerId,[DOCNumberNO],[DOCNO],[DOCDate],[UserId],[DOCType],[SourceNO],[FlowLevel],[FlowStatus],[AIFalg]) VALUES ('{_Fun.Config.ServerId}','{docNumberNO}','{docNO}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}','{userID}','{docType}','{beforeDOCNumberNO}',0,'Y','0')");
                        }
                    }
                    //有單據
                    if (docNumberNO == "") { docNumberNO = tmp["DOCNumberNO"].ToString(); }
                    string tmp_sql = $" and DOCNumberNO='{docNumberNO}'";
                    switch (dr_DOCRole["DOCType"].ToString())
                    {
                        case "3": tmp_sql = $" and OUT_StoreNO='{out_StoreNO}' and OUT_StoreSpacesNO='{out_StoreSpacesNO}'{tmp_sql}"; break; //出庫類
                        case "4": tmp_sql = $" and IN_StoreNO='{in_StoreNO}' and IN_StoreSpacesNO='{in_StoreSpacesNO}'{tmp_sql}"; break;     //入庫類
                    }
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and SUBSTRING(DOCNumberNO,1,4)='{docNO}' and SimulationId='{simulationID}' {tmp_sql}");
                    if (tmp != null)
                    {
                        if (Convert.ToDateTime(tmp["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss") != arrivalDate)
                        {
                            if (beforeId == "") { beforeId = _Str.NewId('Z'); }
                            if (db.DB_SetData($"INSERT INTO [dbo].[DOC3stockII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],ArrivalDate,StartTime,EndTime,IsOK) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',0,'PCS',{qty.ToString()},'{remark}','{simulationID}','{in_StoreNO}','{in_StoreSpacesNO}','{out_StoreNO}','{out_StoreSpacesNO}','{arrivalDate}',{startTime},{endTime},'{docOK}')"))
                            { isrun = true; }
                        }
                        else
                        {
                            beforeId = tmp["Id"].ToString();
                            partNO = d["PartNO"].ToString();
                            startTime = "";
                            if (write_StartTime)
                            {
                                if (tmp.IsNull("StartTime"))
                                { startTime = $",StartTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}'"; }
                            }
                            if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set QTY+={qty.ToString()}{startTime} where Id='{beforeId}' and DOCNumberNO='{tmp["DOCNumberNO"].ToString()}'"))
                            { isrun = true; }
                        }
                    }
                    else
                    {
                        if (beforeId == "") { beforeId = _Str.NewId('Z'); }
                        if (db.DB_SetData($"INSERT INTO [dbo].[DOC3stockII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],ArrivalDate,StartTime,EndTime,IsOK) VALUES ('{_Fun.Config.ServerId}','{beforeId}','{docNumberNO}','{partNO}',0,'PCS',{qty.ToString()},'{remark}','{simulationID}','{in_StoreNO}','{in_StoreSpacesNO}','{out_StoreNO}','{out_StoreSpacesNO}','{arrivalDate}',{startTime},{endTime},'{docOK}')"))
                        { isrun = true; }
                    }
                }
                if (isrun)
                {
                    string[] tmp_array = funName.Split(";");
                    string proName = "";
                    if (tmp_array.Length >= 2) { proName = tmp_array[1]; }
                    db.DB_SetData($"INSERT INTO [dbo].[DOCUpdateLog] ([Id],[LOGDateTime],[DOCId],[DOCNumberNO],[NeedId],[SimulationId],[PartNO],[QTY],[FunName],[ProName]) VALUES ('{_Str.NewId('0')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}','{beforeId}','{docNumberNO}','{needId}','{simulationID}','{partNO}',{qty.ToString()},'{tmp_array[0]}','{proName}')");
                    if (remark != "")
                    {
                        db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId,SimulationId) VALUES ('開立{dr_DOCRole["DOCName"].ToString()} {remark} 料號:{partNO} 數量:{qty.ToString()}','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationID}')");
                    }
                }
                else
                { db.DB_SetData($"INSERT INTO [dbo].[DOCUpdateLog] ([Id],[LOGDateTime],[DOCId],[DOCNumberNO],[NeedId],[SimulationId],[PartNO],[QTY],[FunName],[ProName]) VALUES ('{_Str.NewId('0')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}','{beforeId}','{docNumberNO}','{needId}','{simulationID}','{partNO}',{qty.ToString()},'程式錯誤','{funName}')"); }
            }
            catch (Exception ex)
            {
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"SNWebSocketService.cs Create_DOC3stock Exception: NId={needId} SID={simulationID} PartNO={partNO} {ex.Message} {ex.StackTrace}", true);
            }
            return isrun;
        }
        public void ConfirmHasTotalStock(DBADO db, string partNO, string storeNO, string storeSpacesNO)
        {
            #region 確認 TotalStock檔 有對應資料
            if (storeNO != "")
            {
                DataRow tmp = null;
                if (storeSpacesNO != "")
                { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{storeNO}' and StoreSpacesNO='{storeSpacesNO}'"); }
                else
                { tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and StoreNO='{storeNO}'"); }
                if (tmp == null)
                {
                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{storeNO}'");
                    if (tmp != null)
                    {
                        db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (Class,ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{tmp["Class"].ToString()}','{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{storeNO}','{storeSpacesNO}','{partNO}',0)");
                    }
                }
            }
            #endregion
        }

        public string SelectDOC1BuyMFNO(DBADO db, string partNO, string simulationId, string docNO, ref float price)
        {
            price = 0;
            string mfData = "";
            string tmp = "";
            DataRow dr_tmp = null;
            if (simulationId != "")
            {
                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{simulationId}'");
                if (bool.Parse(dr_tmp["OutPackType"].ToString()))
                {
                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_tmp["NeedId"].ToString()}' and Apply_StationNO='{dr_tmp["Source_StationNO"].ToString()}' and IndexSN='{dr_tmp["Source_StationNO_IndexSN"].ToString()}' and PartSN={(int.Parse(dr_tmp["PartSN"].ToString()) + 1).ToString()}");
                    if (dr_tmp != null)
                    {
                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where Id='{dr_tmp["Apply_BOMId"].ToString()}' and PartNO='{partNO}'");
                        if (dr_tmp != null)
                        {
                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT MFNO FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{dr_tmp["Apply_PP_Name"].ToString()}' and Apply_PartNO='{dr_tmp["Apply_PartNO"].ToString()}' and StationNO='{dr_tmp["Apply_StationNO"].ToString()}' and IndexSN={dr_tmp["IndexSN"].ToString()}");
                            if (dr_tmp != null && !dr_tmp.IsNull("MFNO") && dr_tmp["MFNO"].ToString() != "")
                            { mfData = dr_tmp["MFNO"].ToString(); }
                        }
                    }
                }
            }
            if (docNO != "") { tmp = $" and SUBSTRING(a.DOCNumberNO,1,4)='{docNO}'"; }
            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT a.*,b.MFNO FROM SoftNetMainDB.[dbo].[DOC1BuyII] as a,SoftNetMainDB.[dbo].[DOC1Buy] as b where a.PartNO='{partNO}' and a.DOCNumberNO=b.DOCNumberNO and a.Price!=0 {tmp} order by a.ArrivalDate desc,b.DOCNumberNO desc,a.Price ASC");
            if (dr_tmp != null)
            {
                if (mfData == "" && !dr_tmp.IsNull("MFNO")) { mfData = dr_tmp["MFNO"].ToString(); }
                price = float.Parse(dr_tmp["Price"].ToString());
            }
            if (mfData == "")
            {
                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT APS_Default_MFNO FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and APS_Default_MFNO is not null and APS_Default_MFNO!=''");
                if (dr_tmp != null) { mfData = dr_tmp["APS_Default_MFNO"].ToString(); }

            }
            return mfData;
        }
        public string SelectDOC4ProductionMFNO(DBADO db, string partNO, string simulationId, string docNO, ref float price)
        {
            price = 0;
            string mfData = "";
            string tmp = "";
            DataRow dr_tmp = null;
            if (docNO != "") { tmp = $" and SUBSTRING(a.DOCNumberNO,1,4)='{docNO}'"; }
            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT a.*,b.MFNO FROM SoftNetMainDB.[dbo].[DOC4ProductionII] as a,SoftNetMainDB.[dbo].[DOC4Production] as b where a.PartNO='{partNO}' and a.DOCNumberNO=b.DOCNumberNO {tmp} order by a.ArrivalDate desc,b.DOCNumberNO desc,a.Price ASC");
            if (dr_tmp != null)
            {
                if (mfData == "" && !dr_tmp.IsNull("MFNO") && dr_tmp["MFNO"].ToString() != "") { mfData = dr_tmp["MFNO"].ToString(); }
                price = float.Parse(dr_tmp["Price"].ToString());
            }
            if (mfData == "" && simulationId != "")
            {
                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{simulationId}'");
                if (bool.Parse(dr_tmp["OutPackType"].ToString()))
                {
                    DataRow dr_tmp1 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_tmp["NeedId"].ToString()}' and Apply_PP_Name='{dr_tmp["Apply_PP_Name"].ToString()}' and Apply_StationNO='{dr_tmp["Source_StationNO"].ToString()}' and IndexSN={dr_tmp["Source_StationNO_IndexSN"].ToString()}");
                    if (dr_tmp1 != null)
                    {
                        DataRow dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where Id='{dr_tmp1["Apply_BOMId"].ToString()}'");
                        if (dr_tmp2 != null && !dr_tmp2.IsNull("MFNO") && dr_tmp2["MFNO"].ToString() != "")
                        {
                            mfData = dr_tmp2["MFNO"].ToString();
                        }
                    }
                    else
                    {
                        if (dr_tmp["Source_StationNO_IndexSN"].ToString() == "1")
                        {
                            DataRow dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and IndexSN=1 and PartNO='{dr_tmp["PartNO"].ToString()}' and Apply_PP_Name='{dr_tmp["Apply_PP_Name"].ToString()}' and Apply_StationNO='{dr_tmp["Source_StationNO"].ToString()}'");
                            if (dr_tmp2 != null && !dr_tmp2.IsNull("MFNO") && dr_tmp2["MFNO"].ToString() != "")
                            {
                                mfData = dr_tmp2["MFNO"].ToString();
                            }
                        }
                    }
                }
            }
            if (mfData == "")
            {
                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT APS_Default_MFNO FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}' and APS_Default_MFNO is not null and APS_Default_MFNO!=''");
                if (dr_tmp != null) { mfData = dr_tmp["APS_Default_MFNO"].ToString(); }

            }
            return mfData;
        }

        public void SelectINStore(DBADO db, string partNO, ref string in_StoreNO, ref string in_StoreSpacesNO, string docNO, bool is_outProduction = false)
        {
            #region 查找適合入庫儲別
            DataRow tmp = null;
            if (is_outProduction && _Fun.Config.Default_WorkingPaper_AGE03 != "")
            {
                in_StoreNO = _Fun.Config.Default_WorkingPaper_AGE03;
                if (_Fun.Config.Default_WorkingPaper_AGE03_StoreSpacesNO != "")
                { in_StoreSpacesNO = _Fun.Config.Default_WorkingPaper_AGE03_StoreSpacesNO; }
            }
            else
            {
                DataRow dr_Material = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}'");
                if (dr_Material != null && !dr_Material.IsNull("APS_Default_StoreNO") && dr_Material["APS_Default_StoreNO"].ToString().Trim() != "")
                {
                    in_StoreNO = dr_Material["APS_Default_StoreNO"].ToString();
                    in_StoreSpacesNO = dr_Material["APS_Default_StoreSpacesNO"].ToString();
                }
                else
                {
                    #region 查歷史單據
                    if (docNO != "")
                    {
                        string sql = "";
                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[DOCRole] where ServerId='{_Fun.Config.ServerId}' and DOCNO='{docNO}'");
                        switch (tmp["DOCType"].ToString())
                        {
                            case "1":
                                sql = $"SELECT * from SoftNetMainDB.[dbo].[DOC1BuyII] where PartNO='{partNO}' and StoreNO!='' order by ArrivalDate desc";
                                tmp = db.DB_GetFirstDataByDataRow(sql);
                                if (tmp != null) { in_StoreNO = tmp["StoreNO"].ToString(); in_StoreSpacesNO = tmp["StoreNO"].ToString(); }
                                ConfirmHasTotalStock(db, partNO, in_StoreNO, in_StoreSpacesNO);
                                return;
                            case "2":
                                sql = $"SELECT * from SoftNetMainDB.[dbo].[DOC2SalesII] where PartNO='{partNO}' and StoreNO!='' order by ArrivalDate desc";
                                tmp = db.DB_GetFirstDataByDataRow(sql);
                                if (tmp != null) { in_StoreNO = tmp["StoreNO"].ToString(); in_StoreSpacesNO = tmp["StoreNO"].ToString(); }
                                ConfirmHasTotalStock(db, partNO, in_StoreNO, in_StoreSpacesNO);
                                return;
                            case "3":
                            case "4":
                            case "5":
                            case "8":
                                sql = $"SELECT * from SoftNetMainDB.[dbo].[DOC3stockII] where PartNO='{partNO}' and (IN_StoreNO!='' or IN_StoreNO!='') order by ArrivalDate desc";
                                tmp = db.DB_GetFirstDataByDataRow(sql);
                                if (tmp != null)
                                {
                                    if (tmp["IN_StoreNO"].ToString() != "")
                                    { tmp["IN_StoreNO"].ToString(); in_StoreSpacesNO = tmp["IN_StoreSpacesNO"].ToString(); ConfirmHasTotalStock(db, partNO, in_StoreNO, in_StoreSpacesNO); return; }
                                    if (tmp["OUT_StoreNO"].ToString() != "")
                                    { tmp["OUT_StoreNO"].ToString(); in_StoreSpacesNO = tmp["OUT_StoreNO"].ToString(); ConfirmHasTotalStock(db, partNO, in_StoreNO, in_StoreSpacesNO); return; }
                                }
                                break;
                            case "6": //###???工單
                                break;
                            case "7":
                                sql = $"SELECT * from SoftNetMainDB.[dbo].[DOC4ProductionII] where PartNO='{partNO}' and StoreNO!='' order by ArrivalDate desc";
                                tmp = db.DB_GetFirstDataByDataRow(sql);
                                if (tmp != null)
                                { tmp["StoreNO"].ToString(); in_StoreSpacesNO = tmp["StoreNO"].ToString(); }
                                ConfirmHasTotalStock(db, partNO, in_StoreNO, in_StoreSpacesNO);
                                return;
                            case "9": //###???財務傳票
                                break;
                        }
                    }
                    #endregion
                    if (in_StoreNO == "")
                    {
                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and Class!='虛擬倉' and PartNO='{partNO}' order by QTY desc");
                        if (tmp != null) { in_StoreNO = tmp["StoreNO"].ToString(); in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString(); }
                    }
                    else
                    {
                        if (dr_Material["Class"].ToString() == "6" || dr_Material["Class"].ToString() == "7")
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and (Default_IN_OUT='4' or Default_IN_OUT='3') order by Default_IN_OUT desc");
                            if (tmp != null) { in_StoreNO = tmp["StoreNO"].ToString(); ConfirmHasTotalStock(db, partNO, in_StoreNO, in_StoreSpacesNO); return; }
                        }
                        if (dr_Material["Class"].ToString() == "5")
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and (Default_IN_OUT='5' or Default_IN_OUT='3') order by Default_IN_OUT");
                            if (tmp != null) { in_StoreNO = tmp["StoreNO"].ToString(); ConfirmHasTotalStock(db, partNO, in_StoreNO, in_StoreSpacesNO); return; }
                        }
                        if (dr_Material["Class"].ToString() == "4")
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and (Default_IN_OUT='2' or Default_IN_OUT='3') order by Default_IN_OUT");
                            if (tmp != null) { in_StoreNO = tmp["StoreNO"].ToString(); ConfirmHasTotalStock(db, partNO, in_StoreNO, in_StoreSpacesNO); return; }
                        }
                        if (dr_Material["Class"].ToString() == "1" || dr_Material["Class"].ToString() == "2" || dr_Material["Class"].ToString() == "3")
                        {
                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and (Default_IN_OUT='1' or Default_IN_OUT='3') order by Default_IN_OUT");
                            if (tmp != null) { in_StoreNO = tmp["StoreNO"].ToString(); ConfirmHasTotalStock(db, partNO, in_StoreNO, in_StoreSpacesNO); return; }
                        }
                    }
                    if (in_StoreNO == "")
                    {
                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and Default_IN_OUT='0'");
                        if (tmp != null)
                        {
                            in_StoreNO = tmp["StoreNO"].ToString();
                        }
                        else
                        {
                            in_StoreNO = _Fun.Config.DefaultStoreNO;
                        }
                    }
                }
            }
            #endregion
            ConfirmHasTotalStock(db, partNO, in_StoreNO, in_StoreSpacesNO);

        }

        public void SelectOUTStore(DBADO db, string partNO, ref string out_StoreNO, ref string out_StoreSpacesNO, string docNO)
        {
            DataRow tmp = null;
            #region 查找適合出庫儲別
            DataRow dr_Material = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}'");
            if (dr_Material != null && !dr_Material.IsNull("APS_Default_StoreNO") && dr_Material["APS_Default_StoreNO"].ToString().Trim() != "")
            {
                out_StoreNO = dr_Material["APS_Default_StoreNO"].ToString();
                out_StoreSpacesNO = dr_Material["APS_Default_StoreSpacesNO"].ToString();
            }
            else
            {
                #region 查歷史單據
                if (docNO != "")
                {
                    string sql = "";
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[DOCRole] where ServerId='{_Fun.Config.ServerId}' and DOCNO='{docNO}'");
                    switch (tmp["DOCType"].ToString())
                    {
                        case "1":
                            sql = $"SELECT * from SoftNetMainDB.[dbo].[DOC1BuyII] where PartNO='{partNO}' and StoreNO!='' order by ArrivalDate desc";
                            tmp = db.DB_GetFirstDataByDataRow(sql);
                            if (tmp != null) { out_StoreNO = tmp["StoreNO"].ToString(); out_StoreSpacesNO = tmp["StoreSpacesNO"].ToString(); }
                            ConfirmHasTotalStock(db, partNO, out_StoreNO, out_StoreSpacesNO);
                            return;
                        case "2":
                            sql = $"SELECT * from SoftNetMainDB.[dbo].[DOC2SalesII] where PartNO='{partNO}' and StoreNO!='' order by ArrivalDate desc";
                            tmp = db.DB_GetFirstDataByDataRow(sql);
                            if (tmp != null) { out_StoreNO = tmp["StoreNO"].ToString(); out_StoreSpacesNO = tmp["StoreSpacesNO"].ToString(); }
                            ConfirmHasTotalStock(db, partNO, out_StoreNO, out_StoreSpacesNO);
                            return;
                        case "3":
                        case "4":
                        case "5":
                        case "8":
                            sql = $"SELECT * from SoftNetMainDB.[dbo].[DOC3stockII] where PartNO='{partNO}' and (IN_StoreNO!='' or OUT_StoreNO!='') order by ArrivalDate desc";
                            tmp = db.DB_GetFirstDataByDataRow(sql);
                            if (tmp != null)
                            {
                                if (tmp["OUT_StoreNO"].ToString() != "")
                                { out_StoreNO = tmp["OUT_StoreNO"].ToString(); out_StoreSpacesNO = tmp["OUT_StoreSpacesNO"].ToString(); ConfirmHasTotalStock(db, partNO, out_StoreNO, out_StoreSpacesNO); return; }
                                if (tmp["IN_StoreNO"].ToString() != "")
                                { out_StoreNO = tmp["IN_StoreNO"].ToString(); out_StoreSpacesNO = tmp["IN_StoreSpacesNO"].ToString(); ConfirmHasTotalStock(db, partNO, out_StoreNO, out_StoreSpacesNO); return; }
                            }
                            break;
                        case "6": //###???工單
                            break;
                        case "7":
                            sql = $"SELECT * from SoftNetMainDB.[dbo].[DOC4ProductionII] where PartNO='{partNO}' and StoreNO!='' order by ArrivalDate desc";
                            tmp = db.DB_GetFirstDataByDataRow(sql);
                            if (tmp != null)
                            { tmp["StoreNO"].ToString(); out_StoreSpacesNO = tmp["StoreSpacesNO"].ToString(); }
                            ConfirmHasTotalStock(db, partNO, out_StoreNO, out_StoreSpacesNO);
                            return;
                        case "9": //###???財務傳票
                            break;
                    }
                }
                #endregion

                if (dr_Material["Class"].ToString() == "6" || dr_Material["Class"].ToString() == "7")
                {
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and (Default_IN_OUT='4' or Default_IN_OUT='3') order by Default_IN_OUT desc");
                    if (tmp != null) { out_StoreNO = tmp["StoreNO"].ToString(); ConfirmHasTotalStock(db, partNO, out_StoreNO, out_StoreSpacesNO); return; }
                }
                if (dr_Material["Class"].ToString() == "5")
                {
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and (Default_IN_OUT='5' or Default_IN_OUT='3') order by Default_IN_OUT");
                    if (tmp != null) { out_StoreNO = tmp["StoreNO"].ToString(); ConfirmHasTotalStock(db, partNO, out_StoreNO, out_StoreSpacesNO); return; }
                }
                if (dr_Material["Class"].ToString() == "4")
                {
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and (Default_IN_OUT='2' or Default_IN_OUT='3') order by Default_IN_OUT");
                    if (tmp != null) { out_StoreNO = tmp["StoreNO"].ToString(); ConfirmHasTotalStock(db, partNO, out_StoreNO, out_StoreSpacesNO); return; }
                }
                if (dr_Material["Class"].ToString() == "1" || dr_Material["Class"].ToString() == "2" || dr_Material["Class"].ToString() == "3")
                {
                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and (Default_IN_OUT='1' or Default_IN_OUT='3') order by Default_IN_OUT");
                    if (tmp != null) { out_StoreNO = tmp["StoreNO"].ToString(); ConfirmHasTotalStock(db, partNO, out_StoreNO, out_StoreSpacesNO); return; }
                }

                tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and Default_IN_OUT='0'");
                if (tmp != null)
                {
                    out_StoreNO = tmp["StoreNO"].ToString();
                }
                else
                {
                    out_StoreNO = _Fun.Config.DefaultStoreNO;
                }
                ConfirmHasTotalStock(db, partNO, out_StoreNO, out_StoreSpacesNO);
            }
            #endregion
        }

        public void RunSetSimulation(RunSimulation_Arg args, string ipport, List<string> needId, char type)//type=全新模擬計算 2=變更計算 3=通知前端更新網頁 5=由工作底稿觸發
        {
            switch (type)
            {
                case '1':
                case '2':
                case '5':
                case '9':
                    Task ttr = new Task(() =>
                    {
                        RunSetSimulation_thread_0(args, ipport, needId, type);
                    });
                    ttr.Start();
                    break;
                case '3':
                    SendRMSSocketINFO(1, $"LIB_TO_WEB,{ipport},StationStatusChange");
                    //if (_WebSocketList.ContainsKey(ipport))
                    //{
                    //    Send(_WebSocketList[ipport].socket, "StationStatusChange");
                    //}
                    break;
            }
        }
        private bool SendRMSSocketINFO(byte type, string message)
        {
            try
            {
                IPEndPoint ipEndPoint = new IPEndPoint(IPAddress.Parse(_Fun.Config.MasterServiceIP), 5431); //###???
                using (Socket client = new(ipEndPoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp))
                {
                    client.Connect(ipEndPoint);
                    byte[] data = Encoding.UTF8.GetBytes(message);
                    if (data != null)
                    {
                        byte[] byteSend = new byte[data.Length + 5];
                        byte[] tmp2 = BitConverter.GetBytes(data.Length);
                        tmp2.CopyTo(byteSend, 0);
                        byteSend[4] = type;
                        data.CopyTo(byteSend, 5);
                        //###???暫改Send  user.BeginSend(byteSend, 0, byteSend.Length, SocketFlags.None, new AsyncCallback(rms_sendCallback), user);
                        if (client.Send(byteSend, 0, byteSend.Length, SocketFlags.None) > 0) { client.Close(); return true; }
                    }

                }
            }catch(Exception ex)
            {
                //###???? log
            }
            return false;
        }
        private void RunSetSimulation_thread_0(RunSimulation_Arg args, string ipport, List<string> needId, char wType) //wType 5=由工作底稿觸發
        {
            _Str.Set_Simulation_flag = true;
            DBADO db = new DBADO("1", _Fun.Config.Db);
            //lock (_Fun.Lock_Simulation_Flag)
            //{
                try
                {
                    if (ipport == null) { ipport = ""; }
                    int isARGs10_offset = 15;//###??? 10將來改參數
                    string err = "";
                    DataTable dt_tmp = null;
                    DataRow dr_tmp;
                    string first_M = "";
                    string today = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
                    for (int ai = 0; ai < needId.Count; ai++)
                    {
                        err = "";
                        DataRow dr_M = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{needId[ai]}'");
                        int tmp_int = int.Parse(dr_M["NeedQTY"].ToString());
                        DateTime intime = Convert.ToDateTime(dr_M["NeedDate"]);
                        if (wType == '5')
                        {
                            db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId) VALUES ('系統主動發出生產製造命令, 排程碼:{needId[ai]} 生產件:{dr_M["PartNO"].ToString()} 生產量:{dr_M["NeedQTY"].ToString()}','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId[ai]}')");
                        }
                        #region 若庫存足量跳過排程, 或減 NeedQTY 需求量
                        if (args.ARGs[11] && wType != '5')
                        {
                            //計算需求是否足量   庫存量 + 之前時間內已排產量 + IsOK=0時間內的單據
                            int tot_TotalStock_QTY = 0;//可用存量
                            int tot_DOC1BuyII_QTY = 0;//已keep量
                            dt_tmp = db.DB_GetData($"SELECT a.Id,a.PartNO, sum(a.QTY) as sQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{dr_M["PartNO"].ToString()}' group by a.Id, a.PartNO");
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                DataRow dr_tmp2 = null;
                                foreach (DataRow dr in dt_tmp.Rows)
                                {
                                    tot_TotalStock_QTY += dr.IsNull("sQTY") ? 0 : int.Parse(dr["sQTY"].ToString());
                                    DataTable dt_StockII = db.DB_GetData($"SELECT SimulationId,sum(KeepQTY+OverQTY) as kQTY FROM SoftNetMainDB.[dbo].[TotalStockII] where Id='{dr["Id"].ToString()}' group by SimulationId");
                                    if (dt_StockII != null && dt_StockII.Rows.Count > 0)
                                    {
                                        foreach (DataRow dr_tStockII in dt_StockII.Rows)
                                        {
                                            if (!dr_tStockII.IsNull("kQTY"))
                                            {
                                                dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_tStockII["SimulationId"].ToString()}'");
                                                if (dr_tmp2 != null && !dr_tmp2.IsNull("StartDate") && Convert.ToDateTime(dr_tmp2["StartDate"]) > intime)
                                                {
                                                    break;
                                                }
                                                tot_DOC1BuyII_QTY += dr_tStockII.IsNull("kQTY") ? 0 : int.Parse(dr_tStockII["kQTY"].ToString());
                                            }
                                        }
                                    }
                                }
                            }
                            tot_TotalStock_QTY -= tot_DOC1BuyII_QTY;
                            if (tot_TotalStock_QTY > 0)
                            {
                                #region 計算領料單計劃日, 得到 startDate 日期
                                int math_EfficientCT = 0;
                                //###??? DOCNO 暫時寫死 BC01  , 應該改成 針對領出類型的統計
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select ROUND(sum(EfficientCycleTime)*sum(CountQTY)/sum(CountQTY),0) as CT from SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and DOCNO='BC01' and Sub_PartNO='{dr_M["PartNO"].ToString()}' group by Sub_PartNO");
                                if (dr_tmp != null && !dr_tmp.IsNull("CT") && dr_tmp["CT"].ToString() != "" && dr_tmp["CT"].ToString() != "0")
                                {
                                    math_EfficientCT = int.Parse(dr_tmp["CT"].ToString());
                                }
                                intime = new DateTime(intime.Year, intime.Month, intime.Day, intime.Hour, intime.Minute, DateTime.Now.Second, DateTime.Now.Millisecond);
                                DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_M["CalendarName"].ToString()}' and Holiday<='{intime.ToString("MM/dd/yyyy HH:mm:ss.fff")}' order by Holiday desc");
                                DateTime etime = DateTime.Now;
                                DateTime stime2 = DateTime.Now;
                                int typeTotalTime = 0;
                                foreach (DataRow dr in dt.Rows)
                                {
                                    stime2 = Convert.ToDateTime(dr["Holiday"]);
                                    stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Millisecond);
                                    if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                                    {
                                        string[] comp = dr["Shift_Morning"].ToString().Trim().Split(',');
                                        etime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                                        stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                        typeTotalTime += TimeCompute2Seconds(stime2, etime);
                                        if (math_EfficientCT >= typeTotalTime)
                                        {
                                            etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                            stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                            typeTotalTime += TimeCompute2Seconds(stime2, etime);
                                            if (math_EfficientCT >= typeTotalTime) { intime = stime2; break; }
                                        }
                                        else { intime = stime2; break; }
                                    }
                                    if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                                    {
                                        string[] comp = dr["Shift_Afternoon"].ToString().Trim().Split(',');
                                        etime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                                        stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                        typeTotalTime += TimeCompute2Seconds(stime2, etime);
                                        if (math_EfficientCT >= typeTotalTime)
                                        {
                                            etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                            stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                            typeTotalTime += TimeCompute2Seconds(stime2, etime);
                                            if (math_EfficientCT >= typeTotalTime) { intime = stime2; break; }
                                        }
                                        else { intime = stime2; break; }
                                    }
                                    if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                                    {
                                        string[] comp = dr["Shift_Night"].ToString().Trim().Split(',');
                                        etime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                                        stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                        typeTotalTime += TimeCompute2Seconds(stime2, etime);
                                        if (math_EfficientCT >= typeTotalTime)
                                        {
                                            if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                            { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1); }
                                            else
                                            {
                                                etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                            }
                                            stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                            typeTotalTime += TimeCompute2Seconds(stime2, etime);
                                            if (math_EfficientCT >= typeTotalTime) { intime = stime2; break; }
                                        }
                                        else { intime = stime2; break; }
                                    }
                                    if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                                    {
                                        string[] comp = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                                        if (int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                                        { etime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0).AddDays(1); }
                                        else { etime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0); }
                                        stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                        typeTotalTime += TimeCompute2Seconds(stime2, etime);
                                        if (math_EfficientCT >= typeTotalTime)
                                        {
                                            if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                            { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1); }
                                            else
                                            {
                                                etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                            }
                                            stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                            typeTotalTime += TimeCompute2Seconds(stime2, etime);
                                            if (math_EfficientCT >= typeTotalTime) { intime = stime2; break; }
                                        }
                                        else { intime = stime2; break; }
                                    }
                                }
                                string startDate = intime.ToString("MM/dd/yyyy HH:mm:ss.fff");
                                #endregion

                                #region 改變 NeedQTY 需求量
                                if (tot_TotalStock_QTY < tmp_int)
                                {
                                    tmp_int = int.Parse(dr_M["NeedQTY"].ToString()) - tot_TotalStock_QTY;
                                }
                                else
                                {
                                    tot_TotalStock_QTY = int.Parse(dr_M["NeedQTY"].ToString());
                                    tmp_int = 0;
                                }
                                #endregion

                                string id = _Str.NewId('Y');
                                int tot_QTY = 0;
                                DataRow dr_Material = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_M["PartNO"].ToString()}'");
                                DataTable tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_M["PartNO"].ToString()}' and QTY>0 Order by QTY desc,StoreNO desc,StoreSpacesNO desc");
                                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                {
                                    foreach (DataRow d2 in tmp_dt2.Rows)
                                    {
                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                        if (tmp_01 >= tot_TotalStock_QTY)
                                        {
                                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII] ([Id],[NeedId],[SimulationId],[KeepQTY],ArrivalDate) VALUES ('{d2["Id"].ToString()}','{needId[ai]}','{id}',{tot_TotalStock_QTY},'{startDate}')");
                                            tot_QTY = tot_TotalStock_QTY;
                                            tot_TotalStock_QTY = 0;
                                            break;
                                        }
                                        else
                                        {
                                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII] ([Id],[NeedId],[SimulationId],[KeepQTY],ArrivalDate) VALUES ('{d2["Id"].ToString()}','{needId[ai]}','{id}',{tmp_01},'{startDate}')");
                                            tot_TotalStock_QTY -= tmp_01;
                                            tot_QTY += tmp_01;
                                        }
                                    }
                                    string sql = @$"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation] (ServerId,Apply_BOMId,SimulationId,NeedId,PartSN,Master_PartNO,Apply_PP_Name,Apply_StationNO,IndexSN,Station_Custom_IndexSN,PartNO,BOMQTY,NeedQTY,SafeQTY,Math_TotalStock_HasUseQTY,Master_Class,PartType,Class,Math_StandardCT,IsEnd,StartDate,SimulationDate)
                                        VALUES ('{_Fun.Config.ServerId}','','{id}','{needId[ai]}',-1,'','','',-1,'','{dr_M["PartNO"].ToString()}',1,{tot_QTY.ToString()},0,{tot_QTY.ToString()},'','0','{dr_Material["Class"].ToString()}',0,'0','{startDate}','{Convert.ToDateTime(dr_M["NeedDate"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}')";
                                    db.DB_SetData(sql);
                                    if (tmp_int == 0)
                                    {
                                        if (DateTime.Now.AddDays(1) > Convert.ToDateTime(startDate))
                                        { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='4',StateINFO=NULL,UpdateTime='{DateTime.Now.AddDays(1).ToString("MM/dd/yyyy HH:mm:ss.fff")}',NeedSimulationDate='{Convert.ToDateTime(dr_M["NeedDate"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and Id='{needId[ai]}'"); }
                                        else
                                        { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='4',StateINFO=NULL,UpdateTime='{DateTime.Now.AddMinutes(isARGs10_offset).ToString("MM/dd/yyyy HH:mm:ss.fff")}',NeedSimulationDate='{Convert.ToDateTime(dr_M["NeedDate"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and Id='{needId[ai]}'"); }
                                        continue;
                                    }
                                }
                            }
                        }
                        #endregion

                        string mainBOMID = "";//主BOM Id編號
                        #region 取得主BOM 編號
                        if (dr_M["BOMId"].ToString() != "")
                        {
                            mainBOMID = dr_M["BOMId"].ToString();
                            dr_tmp = db.DB_GetFirstDataByDataRow($"select a.Id,a.PartNO,a.Apply_PartNO,a.Apply_PP_Name,a.Apply_StationNO,a.IsEnd,b.Class from SoftNetMainDB.[dbo].[BOM] as a,dbo.[Material] as b where a.IsConfirm='1' and a.Main_Item='1' and b.ServerId='{_Fun.Config.ServerId}' and a.Apply_PartNO='{dr_M["PartNO"].ToString()}' and a.Apply_PartNO=b.PartNO and a.Id='{mainBOMID}' order by EffectiveDate desc");
                        }
                        else
                        { dr_tmp = db.DB_GetFirstDataByDataRow($"select a.Id,a.PartNO,a.Apply_PartNO,a.Apply_PP_Name,a.Apply_StationNO,a.IsEnd,b.Class from SoftNetMainDB.[dbo].[BOM] as a,dbo.[Material] as b where a.IsConfirm='1' and a.Main_Item='1' and b.ServerId='{_Fun.Config.ServerId}' and a.Apply_PartNO='{dr_M["PartNO"].ToString()}' and a.Apply_PartNO=b.PartNO and a.EffectiveDate<='{today}' and a.ExpiryDate>='{today}' order by EffectiveDate desc"); }
                        //查找BOM資料
                        if (dr_tmp == null)
                        {
                            if (wType == '5')
                            {
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{needId[ai]}'");
                                System.Threading.Tasks.Task task = _Log.ErrorAsync($"工作底稿發出自動生產失敗, 原因:料號:{dr_M["PartNO"].ToString()} 查無可用BOM表", true);
                            }
                            else
                            { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='5',StateINFO='查無BOM表',UpdateTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}',NeedSimulationDate=NULL where ServerId='{_Fun.Config.ServerId}' and Id='{needId[ai]}'"); }
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId[ai]}'");
                            continue;
                        }
                        else { mainBOMID = dr_tmp["Id"].ToString(); }
                        #endregion

                        first_M = dr_tmp["Apply_PP_Name"].ToString();
                        dt_tmp = db.DB_GetData($"select a.Id,a.M_PP_Name,b.PP_Name  FROM SoftNetMainDB.[dbo].[PP_ProductProcess_Index_M] as a join SoftNetMainDB.[dbo].[PP_ProductProcess_Index_S] as b on a.Id=b.IndexId where a.M_PP_Name='{dr_tmp["Apply_PP_Name"].ToString()}'");
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            //### code未檢查
                            #region 計算是否多線均可生產時  //###??? 此段code 未檢查過
                            //若PP_ProductProcess_Index_M檔有定義多線別時, 每一個線都會計算, 最後取得最短時間,再重新計算產生 APS_Simulation 資料
                            Dictionary<KeyAndValue, DateTime> dtime = new Dictionary<KeyAndValue, DateTime>();
                            Run_Index(args, db, needId[ai], tmp_int, wType);
                            dtime.Add(new KeyAndValue(dr_tmp["Apply_PP_Name"].ToString(), dr_tmp["Apply_PP_Name"].ToString()), RunSetSimulation_thread(args, db, ipport, needId[ai], mainBOMID, wType, ref err));

                            foreach (DataRow dr in dt_tmp.Rows)
                            {
                                #region 重新初始化
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId[ai]}'");
                                db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needId[ai]}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needId[ai]}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId[ai]}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{needId[ai]}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{needId[ai]}'");
                                db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] where Trigger_NeedId='{needId[ai]}'");

                                #endregion
                                Run_Index(args, db, needId[ai], tmp_int, wType, dr["M_PP_Name"].ToString(), dr["PP_Name"].ToString());
                                dtime.Add(new KeyAndValue(dr["M_PP_Name"].ToString(), dr["PP_Name"].ToString()), RunSetSimulation_thread(args, db, ipport, needId[ai], mainBOMID, wType, ref err));
                            }
                            string pp_M = "";
                            string pp_S = "";
                            DateTime pp_date = new DateTime();
                            foreach (KeyValuePair<KeyAndValue, DateTime> obj in dtime)
                            {
                                if (pp_S == "") { pp_M = obj.Key.Key; pp_S = obj.Key.Value; pp_date = obj.Value; continue; }
                                else
                                {
                                    if (pp_date.CompareTo(obj.Value) > 0)//###??? 應該 > 0
                                    {
                                        pp_M = obj.Key.Key;
                                        pp_S = obj.Key.Value;
                                        pp_date = obj.Value;
                                    }
                                }
                            }
                            #region 重新初始化
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] where Trigger_NeedId='{needId[ai]}'");
                            #endregion
                            if (pp_S != first_M)
                            { Run_Index(args, db, needId[ai], tmp_int, wType, pp_M, pp_S); }
                            else { Run_Index(args, db, needId[ai], tmp_int, wType); }
                            RunSetSimulation_thread(args, db, ipport, needId[ai], mainBOMID, wType, ref err);
                            #endregion
                        }
                        else
                        {
                            string bomid = Run_Index(args, db, needId[ai], tmp_int, wType);
                            if (bomid == mainBOMID)//建立 APS_Simulation 資料  生產量=NeedQTY+SafeQTY
                            { RunSetSimulation_thread(args, db, ipport, needId[ai], mainBOMID, wType, ref err); }
                            else
                            {
                                if (wType == '5')
                                {
                                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{needId[ai]}'");
                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId) VALUES ('系統主動發出生產製造命令,發生BOM表與製程有差異錯誤: 排程碼:{needId[ai]} 生產件:{dr_M["PartNO"].ToString()} 生產量:{dr_M["NeedQTY"].ToString()}','01','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId[ai]}')");

                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"工作底稿發出自動生產失敗, 原因:料號:{dr_M["PartNO"].ToString()} 參數BOM;{mainBOMID} 運算BOM:{bomid} BOM表與製程有差異", true);
                                }
                                else
                                { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='5',StateINFO='BOM表與製程有差異',UpdateTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}',NeedSimulationDate=NULL where ServerId='{_Fun.Config.ServerId}' and Id='{needId[ai]}'"); }
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId[ai]}'");
                                db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needId[ai]}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needId[ai]}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId[ai]}'");
                            }
                        }
                        if (err != "")
                        {
                            if (wType == '5')
                            {
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{needId[ai]}'");
                                System.Threading.Tasks.Task task = _Log.ErrorAsync($"工作底稿發出自動生產失敗, 原因:料號:{dr_M["PartNO"].ToString()} {err}", true);
                            }
                            else
                            { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='5',StateINFO='{err}',UpdateTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}',NeedSimulationDate=NULL where ServerId='{_Fun.Config.ServerId}' and Id='{needId[ai]}'"); }
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needId[ai]}'");
                            db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId[ai]}'");
                        }
                    }
                    //if (_WebSocketList.ContainsKey(ipport))
                    //{
                    //    Send(_WebSocketList[ipport].socket, "StationStatusChange");
                    //}
                    SendRMSSocketINFO(1, $"LIB_TO_WEB,{ipport},StationStatusChange");
                }
                catch (Exception ex)
                {
                    //if (_WebSocketList.ContainsKey(ipport))
                    //{
                        string sortID = "";
                        for (int i = 0; i < needId.Count; i++)
                        {
                            if (i == 0)
                            { sortID = $"('{needId[i]}'"; }
                            else { sortID = $"{sortID},'{needId[i]}'"; }
                        }
                        if (sortID != "") { sortID += ")"; }
                        if (wType == '5')
                        { db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id in {sortID}"); }
                        else
                        { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='5',StateINFO='程式異常Exception',UpdateTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}',NeedSimulationDate=NULL where ServerId='{_Fun.Config.ServerId}' and Id in {sortID}"); }
                        db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId in {sortID}");
                        db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId in {sortID}");
                        db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId in {sortID}");
                        db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId in {sortID}");
                    SendRMSSocketINFO(1, $"LIB_TO_WEB,{ipport},StationStatusChange");
                    //Send(_WebSocketList[ipport].socket, "StationStatusChange");
                    //}
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"SNWebSockeyService.cs RunSetSimulation_thread_0 wType={wType} {ex.Message} {ex.StackTrace}", true);
                }
            //}
            _Str.Set_Simulation_flag = false;
            db.Dispose();
            args.Dispose();
        }
        private string Run_Index(RunSimulation_Arg args, DBADO db, string needId, int need_QTY, char wType, string changeBOM_M = "", string changeBOM_S = "")
        {
            string mbomId = "";
            //changeBOM不等於空白時,要變更製程
            DataRow dr_M = db.DB_GetFirstDataByDataRow($"SELECT a.*,b.Class,b.PartType,b.StoreSTime from SoftNetSYSDB.[dbo].[APS_NeedData] as a,SoftNetMainDB.[dbo].[Material] as b where a.ServerId='{_Fun.Config.ServerId}' and b.ServerId='{_Fun.Config.ServerId}' and a.Id='{needId}' and a.PartNO=b.PartNO");
            string sql = "";
            int safeQTY = 0;
            #region 需求量+-調整量 = 主體計算量 
            if (int.Parse(dr_M["ChangeQTY"].ToString()) != 0) { need_QTY -= int.Parse(dr_M["ChangeQTY"].ToString()); }
            #endregion
            if (wType == '5') { safeQTY = need_QTY; need_QTY = 0; }


            #region 建立BOM結構,寫入模擬表 新增(APS_Simulation)
            string today = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
            if (dr_M["BOMId"].ToString() != "") { sql = $"select a.Id,a.PartNO,a.Apply_PartNO,a.Apply_PP_Name,a.Apply_StationNO,a.IsEnd,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName,a.OutPackType,b.Class from SoftNetMainDB.[dbo].[BOM] as a,dbo.[Material] as b where a.IsConfirm='1' and a.Main_Item='1' and b.ServerId='{_Fun.Config.ServerId}' and a.Apply_PartNO='{dr_M["PartNO"].ToString()}' and a.Id='{dr_M["BOMId"].ToString()}' and a.Apply_PartNO=b.PartNO order by EffectiveDate desc"; }
            else if (dr_M["Apply_PP_Name"].ToString() != "") { sql = $"select a.Id,a.PartNO,a.Apply_PartNO,a.Apply_PP_Name,a.Apply_StationNO,a.IsEnd,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName,a.OutPackType,b.Class from SoftNetMainDB.[dbo].[BOM] as a,dbo.[Material] as b where a.IsConfirm='1' and a.Main_Item='1' and b.ServerId='{_Fun.Config.ServerId}' and a.Apply_PartNO='{dr_M["PartNO"].ToString()}' and a.Apply_PP_Name='{dr_M["Apply_PP_Name"].ToString()}' and a.Apply_PartNO=b.PartNO order by EffectiveDate desc"; }
            else { sql = $"select a.Id,a.PartNO,a.Apply_PartNO,a.Apply_PP_Name,a.Apply_StationNO,a.IsEnd,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName,a.OutPackType,b.Class from SoftNetMainDB.[dbo].[BOM] as a,dbo.[Material] as b where a.IsConfirm='1' and a.Main_Item='1' and b.ServerId='{_Fun.Config.ServerId}' and a.Apply_PartNO='{dr_M["PartNO"].ToString()}' and a.Apply_PartNO=b.PartNO and a.EffectiveDate<='{today}' and a.ExpiryDate>='{today}' order by EffectiveDate desc"; }
            DataRow dr = db.DB_GetFirstDataByDataRow(sql);
            if (dr != null)
            {
                mbomId = dr["Id"].ToString();
                bool is_0SN = true;
                int indexSN = int.Parse(dr["IndexSN"].ToString());
                //取得第一層結構
                DataTable dt_BOMII = db.DB_GetData($"select a.BOMId,a.PartNO,a.BOMQTY,a.Class,b.PartType,b.StoreSTime,b.SafeQTY from SoftNetMainDB.[dbo].[BOMII] as a,dbo.[Material] as b where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO and BOMId='{dr["Id"].ToString()}' order by sn");
                if (dt_BOMII != null)
                {
                    string compare_Data = "";
                    bool isClass45 = false;
                    DataTable dt_S_Merge = null;
                    string apply_StationNO = dr["Apply_StationNO"].ToString();
                    foreach (DataRow dr0 in dt_BOMII.Rows)
                    {
                        isClass45 = false;
                        compare_Data = "";
                        if (changeBOM_S == "")
                        {
                            #region 寫第 0 階
                            if (is_0SN)
                            {
                                string stationNO_Merge = "NULL";
                                #region 寫合併站
                                if (!dr.IsNull("Apply_PP_Name") && dr["Apply_PP_Name"].ToString() != "" && !dr.IsNull("Apply_StationNO") && dr["Apply_StationNO"].ToString() != "")
                                {
                                    dt_S_Merge = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where IndexSN_Merge='1' and PP_Name='{dr["Apply_PP_Name"].ToString()}' and Apply_PartNO='{dr["Apply_PartNO"].ToString()}' and IndexSN={dr["IndexSN"].ToString()} order by DisplaySN");
                                    if (dt_S_Merge != null && dt_S_Merge.Rows.Count > 0)
                                    {
                                        foreach (DataRow d in dt_S_Merge.Rows)
                                        {
                                            if (stationNO_Merge == "NULL") { stationNO_Merge = $"{d["StationNO"].ToString()},"; }
                                            else { stationNO_Merge = $"{stationNO_Merge}{d["StationNO"].ToString()},"; }
                                        }
                                    }
                                }
                                if (stationNO_Merge != "NULL") { stationNO_Merge = $"'{stationNO_Merge}'"; }
                                #endregion
                                if (bool.Parse(dr["OutPackType"].ToString())) { apply_StationNO = _Fun.Config.OutPackStationName; }
                                sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation] (ServerId,OutPackType,Apply_BOMId,Source_StationNO_Custom_DisplayName,SimulationId,NeedId,PartSN,Master_PartNO,Apply_PP_Name,Apply_StationNO,IndexSN,Station_Custom_IndexSN,PartNO,BOMQTY,NeedQTY,SafeQTY,Master_Class,PartType,Class,Math_StandardCT,IsEnd,Source_StationNO,Source_StationNO_IndexSN,Source_StationNO_Custom_IndexSN,StationNO_Merge,PartNO_Replace) VALUES 
                                        ('{_Fun.Config.ServerId}','{dr["OutPackType"].ToString()}','{dr0["BOMId"].ToString()}','{dr["StationNO_Custom_DisplayName"].ToString()}','{_Str.NewId('Y')}','{needId}',0,'{dr_M["PartNO"].ToString()}','{dr["Apply_PP_Name"].ToString()}','{apply_StationNO}',{dr["IndexSN"].ToString()},'{dr["Station_Custom_IndexSN"].ToString()}','{dr_M["PartNO"].ToString()}',1,{need_QTY},{safeQTY},'{dr_M["Class"].ToString()}','{dr_M["PartType"].ToString()}','{dr_M["Class"].ToString()}',{dr_M["StoreSTime"].ToString()},'0','{apply_StationNO}',{dr["IndexSN"].ToString()},'{dr["Station_Custom_IndexSN"].ToString()}',{stationNO_Merge},NULL)";
                                db.DB_SetData(sql);
                                is_0SN = false;
                            }
                            #endregion

                            #region 安全量補足計算
                            if (wType != '5')
                            {
                                if (args.ARGs[1] && int.Parse(dr0["SafeQTY"].ToString()) > 0)
                                {
                                    safeQTY = int.Parse(dr0["SafeQTY"].ToString());
                                    int mQTY = 0;
                                    int kQTY = 0;
                                    int sQTY = 0;
                                    sql = @$"select sum(A.QTY) as mQTY,sum(B.KeepQTY+B.OverQTY) as kQTY from SoftNetMainDB.[dbo].[TotalStock] as A 
                                            join SoftNetMainDB.[dbo].[TotalStockII] as B on A.Id=B.Id and (B.KeepQTY!=0 or B.OverQTY!=0)
                                            where A.Class!='虛擬倉' and A.ServerId='{_Fun.Config.ServerId}' and A.PartNO='{dr0["PartNO"].ToString()}' group by A.PartNO";
                                    DataRow d = db.DB_GetFirstDataByDataRow(sql);
                                    if (d != null && !d.IsNull("mQTY") && d["mQTY"].ToString() != "")
                                    {
                                        mQTY = int.Parse(d["mQTY"].ToString());
                                        if (!d.IsNull("kQTY") && d["kQTY"].ToString() != "")
                                        { kQTY = int.Parse(d["kQTY"].ToString()); }
                                    }
                                    d = db.DB_GetFirstDataByDataRow($"SELECT sum(SafeQTY) as sQTY  FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartNO='{dr0["PartNO"].ToString()}' and IsOK='0' ");
                                    if (d != null && !d.IsNull("sQTY") && d["sQTY"].ToString() != "")
                                    { sQTY = int.Parse(d["sQTY"].ToString()); }
                                    safeQTY = safeQTY - mQTY - kQTY + sQTY;
                                    if (safeQTY < 0) { Math.Abs(safeQTY); }
                                }
                            }
                            #endregion
                            string custom_DisplayName = "";
                            if (dr0["Class"].ToString() == "4" || dr0["Class"].ToString() == "5")
                            {
                                isClass45 = true;
                                custom_DisplayName = dr["StationNO_Custom_DisplayName"].ToString();
                                compare_Data = $"{needId},1,{dr["PartNO"].ToString()},{dr["Apply_PP_Name"].ToString()},{apply_StationNO},{dr["IndexSN"].ToString()},{dr0["PartNO"].ToString()}";
                            }
                            sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation] (ServerId,OutPackType,Apply_BOMId,Source_StationNO_Custom_DisplayName,SimulationId,NeedId,PartSN,Master_PartNO,Apply_PP_Name,Apply_StationNO,IndexSN,Station_Custom_IndexSN,PartNO,BOMQTY,NeedQTY,SafeQTY,Master_Class,PartType,Class,Math_StandardCT,IsEnd) VALUES
                                ('{_Fun.Config.ServerId}','{dr["OutPackType"].ToString()}','{dr0["BOMId"].ToString()}','{custom_DisplayName}','{_Str.NewId('Y')}','{needId}',1,'{dr["PartNO"].ToString()}','{dr["Apply_PP_Name"].ToString()}','{apply_StationNO}',{dr["IndexSN"].ToString()},'{dr["Station_Custom_IndexSN"].ToString()}','{dr0["PartNO"].ToString()}',{dr0["BOMQTY"].ToString()},{(int.Parse(dr0["BOMQTY"].ToString()) * need_QTY)},{(int.Parse(dr0["BOMQTY"].ToString()) * safeQTY)},'{dr["Class"].ToString()}','{dr0["PartType"].ToString()}','{dr0["Class"].ToString()}',{dr0["StoreSTime"].ToString()},'{dr["IsEnd"].ToString()}')";
                        }
                        else
                        {
                            #region 改變生產線
                            sql = @$"select *  FROM SoftNetMainDB.[dbo].[PP_ProductProcess_Index_M] as a 
                                        join SoftNetMainDB.[dbo].[PP_ProductProcess_Index_S] as b on a.Id=b.IndexId and b.PP_Name='{changeBOM_S}'
                                        join SoftNetMainDB.[dbo].[PP_ProductProcess_Index_S_Index] as c on a.Id=c.IndexId_M and b.Id=c.IndexId_S and c.S_StationNO='{dr["Apply_StationNO"].ToString()}'
                                        where a.M_PP_Name='{changeBOM_M}'";
                            DataRow tmp = db.DB_GetFirstDataByDataRow(sql);
                            if (tmp != null)
                            {
                                #region 寫第 0 階
                                if (is_0SN)
                                {
                                    //###???寫合併站未處裡
                                    sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation] (ServerId,Apply_BOMId,Source_StationNO_Custom_DisplayName,SimulationId,NeedId,PartSN,Master_PartNO,Apply_PP_Name,Apply_StationNO,IndexSN,Station_Custom_IndexSN,PartNO,BOMQTY,NeedQTY,SafeQTY,Master_Class,PartType,Class,Math_StandardCT,IsEnd,Source_StationNO,Source_StationNO_IndexSN,Source_StationNO_Custom_IndexSN,StationNO_Merge,PartNO_Replace) VALUES
                                        ('{_Fun.Config.ServerId}','{dr0["BOMId"].ToString()}','{dr["StationNO_Custom_DisplayName"].ToString()}','{_Str.NewId('Y')}','{needId}',0,'{dr_M["PartNO"].ToString()}','{changeBOM_S}','{tmp["C_StationNO"].ToString()}',{tmp["C_IndexSN"].ToString()},'{tmp["C_Station_Custom_IndexSN"].ToString()}','{dr_M["PartNO"].ToString()}',1,{need_QTY},{safeQTY},'{dr_M["Class"].ToString()}','{dr_M["PartType"].ToString()}','{dr_M["Class"].ToString()}',{dr_M["StoreSTime"].ToString()},'0','{tmp["C_StationNO"].ToString()}',{tmp["C_IndexSN"].ToString()},'{tmp["C_Station_Custom_IndexSN"].ToString()}','','')";
                                    db.DB_SetData(sql);
                                    is_0SN = false;
                                }
                                #endregion

                                #region 安全量補足計算
                                if (wType != '5')
                                {
                                    if (args.ARGs[1] && int.Parse(dr0["SafeQTY"].ToString()) > 0)
                                    {
                                        safeQTY = int.Parse(dr0["SafeQTY"].ToString());
                                        int mQTY = 0;
                                        int kQTY = 0;
                                        int sQTY = 0;
                                        sql = @$"select sum(A.QTY) as mQTY,sum(B.KeepQTY+B.OverQTY) as kQTY from SoftNetMainDB.[dbo].[TotalStock] as A 
                                            join SoftNetMainDB.[dbo].[TotalStockII] as B on A.Id=B.Id and (B.KeepQTY!=0 or B.OverQTY!=0)
                                            where A.Class!='虛擬倉' and A.ServerId='{_Fun.Config.ServerId}' and A.PartNO='{dr0["PartNO"].ToString()}' group by A.PartNO";
                                        DataRow d = db.DB_GetFirstDataByDataRow(sql);
                                        if (d != null && !d.IsNull("mQTY") && d["mQTY"].ToString() != "")
                                        {
                                            mQTY = int.Parse(d["mQTY"].ToString());
                                            if (!d.IsNull("kQTY") && d["kQTY"].ToString() != "")
                                            { kQTY = int.Parse(d["kQTY"].ToString()); }
                                        }
                                        d = db.DB_GetFirstDataByDataRow($"SELECT sum(SafeQTY) as sQTY  FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartNO='{dr0["PartNO"].ToString()}' and IsOK='0' ");
                                        if (d != null && !d.IsNull("sQTY") && d["sQTY"].ToString() != "")
                                        { sQTY = int.Parse(d["sQTY"].ToString()); }
                                        safeQTY = safeQTY - mQTY - kQTY + sQTY;
                                        if (safeQTY < 0) { Math.Abs(safeQTY); }
                                    }
                                }
                                #endregion
                                string custom_DisplayName = "";
                                if (dr0["Class"].ToString() == "4" || dr0["Class"].ToString() == "5")
                                {
                                    isClass45 = true;
                                    custom_DisplayName = dr["StationNO_Custom_DisplayName"].ToString();
                                    compare_Data = $"{needId},1,{dr["PartNO"].ToString()},{changeBOM_S},{tmp["C_StationNO"].ToString()},{tmp["C_IndexSN"].ToString()},{dr0["PartNO"].ToString()}";
                                }
                                sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation] (ServerId,Apply_BOMId,Source_StationNO_Custom_DisplayName,SimulationId,NeedId,PartSN,Master_PartNO,Apply_PP_Name,Apply_StationNO,IndexSN,Station_Custom_IndexSN,PartNO,BOMQTY,NeedQTY,SafeQTY,Master_Class,PartType,Class,Math_StandardCT,IsEnd) VALUES 
                                        ('{_Fun.Config.ServerId}','{dr0["BOMId"].ToString()}','{custom_DisplayName}','{_Str.NewId('Y')}','{needId}',1,'{dr["PartNO"].ToString()}','{changeBOM_S}','{tmp["C_StationNO"].ToString()}',{tmp["C_IndexSN"].ToString()},'{tmp["C_Station_Custom_IndexSN"].ToString()}','{dr0["PartNO"].ToString()}',1,{need_QTY},{safeQTY},'{dr_M["Class"].ToString()}','{dr_M["PartType"].ToString()}','{dr_M["Class"].ToString()}',{dr_M["StoreSTime"].ToString()},'0')";
                            }
                            else
                            { continue; }
                            #endregion
                        }
                        if (db.DB_SetData(sql))
                        {
                            if (compare_Data != "")
                            {
                                string[] ssid_arry = compare_Data.Split(',');
                                if (ssid_arry.Length > 1)
                                {
                                    DataRow ssid_dr = db.DB_GetFirstDataByDataRow($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{ssid_arry[0]}' and PartSN={ssid_arry[1]} and Master_PartNO='{ssid_arry[2]}' and Apply_PP_Name='{ssid_arry[3]}' and Apply_StationNO='{ssid_arry[4]}' and IndexSN={ssid_arry[5]} and PartNO='{ssid_arry[6]}'");
                                    compare_Data = ssid_dr["SimulationId"].ToString().Trim();
                                }
                            }
                            if (isClass45)
                            { RecursiveBOMII(args, db, today, needId, dr["Apply_PP_Name"].ToString(), dr["Apply_PartNO"].ToString(), dr0, 1, need_QTY, safeQTY, compare_Data, indexSN, changeBOM_M, changeBOM_S); }
                        }
                    }
                }
            }
            #endregion

            return mbomId;
        }
        private void Run_IndexII(DBADO db, string needId, int sn, DataTable dr_MII)
        {
            db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set PartSN+=1 where NeedId='{needId}' and PartSN={sn.ToString()} and Master_PartNO='{dr_MII.Rows[0]["Master_PartNO"].ToString()}'");
            foreach (DataRow dr in dr_MII.Rows)
            {
                DataTable dr_MIII = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and Master_PartNO='{dr["PartNO"].ToString()}' order by PartSN, PartSN_Sub desc");
                if (dr_MIII != null && dr_MIII.Rows.Count > 0)
                { Run_IndexII(db, needId, (int.Parse(dr["PartSN"].ToString()) + 1), dr_MIII); }
            }
        }
        private void RecursiveBOMII(RunSimulation_Arg args, DBADO db, string today, string needId, string Apply_PP_Name, string apply_PartNO, DataRow dr0, int i, int need_qty, int need_Safeqty, string ssid, int indexSN, string changeBOM_M = "", string changeBOM_S = "")
        {
            string sql = "";
            DataRow dr = db.DB_GetFirstDataByDataRow($"select a.Id,a.PartNO,a.Apply_PartNO,a.Apply_PP_Name,a.Apply_StationNO,a.IsEnd,a.IndexSN,a.Station_Custom_IndexSN,a.StationNO_Custom_DisplayName,a.OutPackType,b.Class,IsChackQTY,IsChackIsOK from SoftNetMainDB.[dbo].[BOM] as a,dbo.[Material] as b where a.Main_Item='0' and a.IndexSN={(indexSN - 1)} and b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{dr0["PartNO"].ToString()}' and a.Apply_PP_Name='{Apply_PP_Name}' and a.Apply_PartNO='{apply_PartNO}' and a.PartNO=b.PartNO order by IndexSN desc");
            if (dr != null)
            {
                string source_StationNO = dr["Apply_StationNO"].ToString();
                indexSN -= 1;
                if (ssid.Split(',').Length == 1 && dr["Apply_PP_Name"].ToString().Trim() != "" && dr["Apply_StationNO"].ToString().Trim() != "")
                {
                    string stationNO_Merge = "NULL";
                    #region 寫合併站
                    if (!dr.IsNull("Apply_PP_Name") && dr["Apply_PP_Name"].ToString() != "" && !dr.IsNull("Apply_StationNO") && dr["Apply_StationNO"].ToString() != "")
                    {
                        DataTable dt_S_Merge = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where IndexSN_Merge='1' and PP_Name='{dr["Apply_PP_Name"].ToString()}' and Apply_PartNO='{dr["Apply_PartNO"].ToString()}' and IndexSN={dr["IndexSN"].ToString()} order by DisplaySN");
                        if (dt_S_Merge != null && dt_S_Merge.Rows.Count > 0)
                        {
                            foreach (DataRow d in dt_S_Merge.Rows)
                            {
                                if (stationNO_Merge == "NULL") { stationNO_Merge = $"{d["StationNO"].ToString()},"; }
                                else { stationNO_Merge = $"{stationNO_Merge}{d["StationNO"].ToString()},"; }
                            }
                        }
                    }
                    if (stationNO_Merge != "NULL") { stationNO_Merge = $"'{stationNO_Merge}'"; }
                    #endregion

                    if (bool.Parse(dr["OutPackType"].ToString())) { source_StationNO = _Fun.Config.OutPackStationName; }
                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set IsChackQTY='{bool.Parse(dr["IsChackQTY"].ToString())}',IsChackIsOK='{bool.Parse(dr["IsChackIsOK"].ToString())}',OutPackType='{dr["OutPackType"].ToString()}',StationNO_Merge={stationNO_Merge},Source_StationNO='{source_StationNO}',Source_StationNO_IndexSN={dr["IndexSN"].ToString()},Source_StationNO_Custom_IndexSN='{dr["Station_Custom_IndexSN"].ToString().Trim()}',Source_StationNO_Custom_DisplayName='{dr["StationNO_Custom_DisplayName"].ToString().Trim()}' where SimulationId='{ssid}'");
                }
                DataTable dt_BOMII = db.DB_GetData($"select a.BOMId,a.PartNO,a.BOMQTY,a.Class,b.PartType,b.StoreSTime,b.SafeQTY from SoftNetMainDB.[dbo].[BOMII] as a,dbo.[Material] as b where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO and BOMId='{dr["Id"].ToString()}' order by sn");
                if (dt_BOMII != null)
                {
                    string compare_Data = "";
                    bool isClass45 = false;
                    foreach (DataRow dr1 in dt_BOMII.Rows)
                    {
                        isClass45 = false;
                        if (changeBOM_S == "")
                        {
                            #region 安全量補足計算
                            if (need_Safeqty != 0)
                            {
                                if (args.ARGs[1] && int.Parse(dr0["SafeQTY"].ToString()) > 0)
                                {
                                    need_Safeqty = int.Parse(dr0["SafeQTY"].ToString());
                                    int mQTY = 0;
                                    int kQTY = 0;
                                    int sQTY = 0;
                                    sql = @$"select sum(A.QTY) as mQTY,sum(B.KeepQTY+B.OverQTY) as kQTY from SoftNetMainDB.[dbo].[TotalStock] as A 
                                            join SoftNetMainDB.[dbo].[TotalStockII] as B on A.Id=B.Id and (B.KeepQTY!=0 or B.OverQTY!=0)
                                            where A.Class!='虛擬倉' and A.ServerId='{_Fun.Config.ServerId}' and A.PartNO='{dr0["PartNO"].ToString()}' group by A.PartNO";
                                    DataRow d = db.DB_GetFirstDataByDataRow(sql);
                                    if (d != null && !d.IsNull("mQTY") && d["mQTY"].ToString() != "")
                                    {
                                        mQTY = int.Parse(d["mQTY"].ToString());
                                        if (!d.IsNull("kQTY") && d["kQTY"].ToString() != "")
                                        { kQTY = int.Parse(d["kQTY"].ToString()); }
                                    }
                                    d = db.DB_GetFirstDataByDataRow($"SELECT sum(SafeQTY) as sQTY  FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartNO='{dr0["PartNO"].ToString()}' and IsOK='0'");
                                    if (d != null && !d.IsNull("sQTY") && d["sQTY"].ToString() != "")
                                    { sQTY = int.Parse(d["sQTY"].ToString()); }
                                    need_Safeqty = need_Safeqty - mQTY - kQTY + sQTY;
                                    if (need_Safeqty < 0) { Math.Abs(need_Safeqty); }
                                }
                            }
                            #endregion

                            if (dr1["Class"].ToString() == "4" || dr1["Class"].ToString() == "5")
                            {
                                isClass45 = true;
                                compare_Data = $"{needId},{(i + 1)},{dr["PartNO"].ToString()},{dr["Apply_PP_Name"].ToString()},{source_StationNO},{dr["IndexSN"].ToString()},{dr1["PartNO"].ToString()}";
                            }
                            sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation] (ServerId,Apply_BOMId,SimulationId,NeedId,PartSN,Master_PartNO,Apply_PP_Name,Apply_StationNO,IndexSN,PartNO,BOMQTY,NeedQTY,SafeQTY,Master_Class,PartType,Class,Math_StandardCT,IsEnd) VALUES 
                                    ('{_Fun.Config.ServerId}','{dr1["BOMId"].ToString()}','{_Str.NewId('Y')}','{needId}',{(i + 1)},'{dr["PartNO"].ToString()}','{dr["Apply_PP_Name"].ToString()}','{source_StationNO}',{dr["IndexSN"].ToString()},'{dr1["PartNO"].ToString()}',{dr1["BOMQTY"].ToString()},{(int.Parse(dr1["BOMQTY"].ToString()) * need_qty)},{(int.Parse(dr1["BOMQTY"].ToString()) * need_Safeqty)},'{dr["Class"].ToString()}','{dr1["PartType"].ToString()}','{dr1["Class"].ToString()}',{dr1["StoreSTime"].ToString()},'{dr["IsEnd"].ToString()}')";
                        }
                        else
                        {
                            #region
                            sql = @$"select *  FROM SoftNetMainDB.[dbo].[PP_ProductProcess_Index_M] as a 
                                        join SoftNetMainDB.[dbo].[PP_ProductProcess_Index_S] as b on a.Id=b.IndexId and b.PP_Name='{changeBOM_S}'
                                        join SoftNetMainDB.[dbo].[PP_ProductProcess_Index_S_Index] as c on a.Id=c.IndexId_M and b.Id=c.IndexId_S and c.S_StationNO='{dr["Apply_StationNO"].ToString()}'
                                        where a.M_PP_Name='{changeBOM_M}'";
                            DataRow tmp = db.DB_GetFirstDataByDataRow(sql);
                            if (tmp != null)
                            {
                                #region 安全量補足計算
                                if (need_Safeqty != 0)
                                {
                                    if (args.ARGs[1] && int.Parse(dr0["SafeQTY"].ToString()) > 0)
                                    {
                                        need_Safeqty = int.Parse(dr0["SafeQTY"].ToString());
                                        int mQTY = 0;
                                        int kQTY = 0;
                                        int sQTY = 0;
                                        sql = @$"select sum(A.QTY) as mQTY,sum(B.KeepQTY+B.OverQTY) as kQTY from SoftNetMainDB.[dbo].[TotalStock] as A 
                                            join SoftNetMainDB.[dbo].[TotalStockII] as B on A.Id=B.Id and (B.KeepQTY!=0 or B.OverQTY!=0)
                                            where A.Class!='虛擬倉' and A.ServerId='{_Fun.Config.ServerId}' and A.PartNO='{dr0["PartNO"].ToString()}' group by A.PartNO";
                                        DataRow d = db.DB_GetFirstDataByDataRow(sql);
                                        if (d != null && !d.IsNull("mQTY") && d["mQTY"].ToString() != "")
                                        {
                                            mQTY = int.Parse(d["mQTY"].ToString());
                                            if (!d.IsNull("kQTY") && d["kQTY"].ToString() != "")
                                            { kQTY = int.Parse(d["kQTY"].ToString()); }
                                        }
                                        d = db.DB_GetFirstDataByDataRow($"SELECT sum(SafeQTY) as sQTY  FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartNO='{dr0["PartNO"].ToString()}' and IsOK='0' ");
                                        if (d != null && !d.IsNull("sQTY") && d["sQTY"].ToString() != "")
                                        { sQTY = int.Parse(d["sQTY"].ToString()); }
                                        need_Safeqty = need_Safeqty - mQTY - kQTY + sQTY;
                                        if (need_Safeqty < 0) { Math.Abs(need_Safeqty); }
                                    }
                                }
                                #endregion
                                if (dr1["Class"].ToString() == "4" || dr1["Class"].ToString() == "5")
                                {
                                    isClass45 = true;
                                    compare_Data = $"{needId},{(i + 1)},{dr["PartNO"].ToString()},{changeBOM_S},{tmp["C_StationNO"].ToString()},{tmp["C_IndexSN"].ToString()},{dr1["PartNO"].ToString()}";
                                }
                                sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation] (ServerId,Apply_BOMId,SimulationId,NeedId,PartSN,Master_PartNO,Apply_PP_Name,Apply_StationNO,IndexSN,PartNO,BOMQTY,NeedQTY,SafeQTY,Master_Class,PartType,Class,Math_StandardCT,IsEnd) VALUES 
                                        ('{_Fun.Config.ServerId}','{dr1["BOMId"].ToString()}','{_Str.NewId('Y')}','{needId}',{(i + 1)},'{dr["PartNO"].ToString()}','{changeBOM_S}','{tmp["C_StationNO"].ToString()}',{tmp["C_IndexSN"].ToString()},'{dr1["PartNO"].ToString()}',{dr1["BOMQTY"].ToString()},{(int.Parse(dr1["BOMQTY"].ToString()) * need_qty)},{(int.Parse(dr1["BOMQTY"].ToString()) * need_Safeqty)},'{dr["Class"].ToString()}','{dr1["PartType"].ToString()}','{dr1["Class"].ToString()}',{dr1["StoreSTime"].ToString()},'{dr["IsEnd"].ToString()}')";
                            }
                            else
                            {
                                continue;
                            }
                            #endregion
                        }
                        if (db.DB_SetData(sql))
                        {
                            if (compare_Data != "")
                            {
                                string[] ssid_arry = compare_Data.Split(',');
                                if (ssid_arry.Length > 2)
                                {
                                    DataRow ssid_dr = db.DB_GetFirstDataByDataRow($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{ssid_arry[0]}' and PartSN={ssid_arry[1]} and Master_PartNO='{ssid_arry[2]}' and Apply_PP_Name='{ssid_arry[3]}' and Apply_StationNO='{ssid_arry[4]}' and IndexSN={ssid_arry[5]} and PartNO='{ssid_arry[6]}'");
                                    compare_Data = ssid_dr["SimulationId"].ToString().Trim();
                                }
                            }
                            if (isClass45)
                            { RecursiveBOMII(args, db, today, needId, Apply_PP_Name, apply_PartNO, dr1, (i + 1), need_qty, need_Safeqty, compare_Data, indexSN, changeBOM_M, changeBOM_S); }
                        }
                    }
                }
            }
        }

        private void RecursiveUPSID_thread(DBADO db, string needId, DataTable dt_next, ref List<string> compare_Data)
        {
            if (dt_next == null || dt_next.Rows.Count <= 0) { return; }
            DataRow dr_tmp = null;
            DataTable dr_MII = null;
            foreach (DataRow dr1 in dt_next.Rows)
            {
                if (!compare_Data.Contains(dr1["SimulationId"].ToString())) { compare_Data.Add(dr1["SimulationId"].ToString()); }
                if (int.Parse(dr1["PartSN"].ToString()) > 0 && (!dr1.IsNull("Source_StationNO") && (dr1["Class"].ToString() == "4" || dr1["Class"].ToString() == "5")))
                {
                    DataTable dt_nextII = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and Apply_PP_Name='{dr1["Apply_PP_Name"].ToString()}' and Apply_StationNO='{dr1["Source_StationNO"].ToString()}' and IndexSN='{dr1["Source_StationNO_IndexSN"].ToString()}' and PartSN>=0 order by PartSN,Class desc");
                    if (dt_nextII != null && dt_nextII.Rows.Count > 0) { RecursiveUPSID_thread(db, needId, dt_nextII, ref compare_Data); }
                }
            }
        }
        private void RecursiveBOMIII_RunSetSimulation_thread_2(DBADO db, string today, string needId, DataRow dr, DataRow dr2, int needQTY, string pp_Name, string apply_PartNO)
        {
            DataRow dr_tmp = null;
            DataRow dr_M = db.DB_GetFirstDataByDataRow($"select a.Id,a.Apply_PartNO,a.PartNO,a.Apply_PP_Name,a.Apply_StationNO,a.IsEnd,a.IndexSN,a.Station_Custom_IndexSN,b.Class from SoftNetMainDB.[dbo].[BOM] as a,dbo.[Material] as b where a.Main_Item='0' and b.ServerId='{_Fun.Config.ServerId}' and a.Apply_PP_Name='{pp_Name}' and a.Apply_PartNO='{apply_PartNO}' and a.PartNO='{dr["PartNO"].ToString()}' and a.Apply_StationNO='{dr["Source_StationNO"].ToString()}' and a.IndexSN={dr["Source_StationNO_IndexSN"].ToString()} and a.Station_Custom_IndexSN='{dr["Source_StationNO_Custom_IndexSN"].ToString()}' and a.PartNO=b.PartNO and a.EffectiveDate<='{today}' and a.ExpiryDate>='{today}' order by EffectiveDate desc");
            if (dr_M != null)
            {
                bool is_0SN = true;
                //取得第一層結構
                DataTable dt_BOMII = db.DB_GetData($"select a.Id,a.Apply_PartNO,a.PartNO,a.BOMQTY,a.Class,b.PartType,b.StoreSTime,b.SafeQTY from SoftNetMainDB.[dbo].[BOMII] as a,dbo.[Material] as b where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO and BOMId='{dr_M["Id"].ToString()}' order by sn");
                if (dt_BOMII != null)
                {
                    string tmp_s = "";
                    int tmp_i = 0;
                    foreach (DataRow dr0 in dt_BOMII.Rows)
                    {
                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where IsOK='0' and  NeedId='{needId}' and PartNO='{dr0["PartNO"].ToString()}' and Master_PartNO='{dr["PartNO"].ToString()}' and Apply_StationNO='{dr["Source_StationNO"].ToString()}' and IndexSN={dr["Source_StationNO_IndexSN"].ToString()} and Station_Custom_IndexSN='{dr["Source_StationNO_Custom_IndexSN"].ToString()}'");
                        if (dr_tmp != null)
                        {
                            tmp_s = "";
                            tmp_i = int.Parse(dr0["BOMQTY"].ToString()) * needQTY;
                            if (int.Parse(dr_tmp["Math_TotalStock_HasUseQTY"].ToString()) >= tmp_i)
                            {
                                int sdfsd = int.Parse(dr_tmp["Math_TotalStock_HasUseQTY"].ToString());
                                tmp_s = $",Math_TotalStock_HasUseQTY-={tmp_i}";
                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] SET KeepQTY-={tmp_i} where SimulationId='{dr_tmp["SimulationId"].ToString()}'");
                            }
                            else if (int.Parse(dr_tmp["Math_TotalStock_HasUseQTY"].ToString()) != 0 && int.Parse(dr_tmp["Math_TotalStock_HasUseQTY"].ToString()) + int.Parse(dr_tmp["Math_Online_SurplusQTY"].ToString()) > int.Parse(dr_tmp["NeedQTY"].ToString()))
                            {
                                tmp_s = $",Math_TotalStock_HasUseQTY-={(int.Parse(dr_tmp["Math_TotalStock_HasUseQTY"].ToString()) + int.Parse(dr_tmp["Math_Online_SurplusQTY"].ToString()) - int.Parse(dr_tmp["NeedQTY"].ToString()))}";
                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] SET KeepQTY-={(int.Parse(dr_tmp["Math_TotalStock_HasUseQTY"].ToString()) + int.Parse(dr_tmp["Math_Online_SurplusQTY"].ToString()) - int.Parse(dr_tmp["NeedQTY"].ToString()))} where SimulationId='{dr_tmp["SimulationId"].ToString()}'");

                            }
                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET Math_Online_SurplusQTY+={tmp_i}{tmp_s} where SimulationId='{dr_tmp["SimulationId"].ToString()}'");
                            if (dr0["Class"].ToString() == "4" || dr0["Class"].ToString() == "5")
                            {
                                RecursiveBOMIII_RunSetSimulation_thread_2(db, today, needId, dr_tmp, dr0, needQTY, pp_Name, apply_PartNO);
                            }
                        }
                    }
                }
            }
        }

        private DateTime RunSetSimulation_thread(RunSimulation_Arg args, DBADO db, string ipport, string needId, string mainBOMID, char wType, ref string ERR)
        {
            DateTime sTimeLog = DateTime.Now;
            bool isError = false;
            string errINFO = "";
            string tmp_station = "";
            //try
            //{
            //###??? 將來要寫急件的 keep量 或 替代料 或 差單 的程式

            DataTable dt_Simulation = null;
            DataRow dr_APS_NeedData = null;

            dr_APS_NeedData = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{needId}'");
            if (dr_APS_NeedData["State"].ToString() != "1") { ERR = "異常加異常"; return sTimeLog; }
            DateTime need_dDayNoBufferTime = Convert.ToDateTime(dr_APS_NeedData["NeedDate"]);
            DateTime need_dDay = need_dDayNoBufferTime.AddHours(-int.Parse(dr_APS_NeedData["BufferTime"].ToString()));

            #region 檢查APS_Simulation資料是否完整
            bool iserr = false;
            dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartSN>=0 order by PartSN");
            if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
            {
                #region 檢查PartSN
                int sn = int.Parse(dt_Simulation.Rows[0]["PartSN"].ToString());
                if (sn != 0) { iserr = true; }
                else
                {
                    foreach (DataRow dr in dt_Simulation.Rows)
                    {
                        if (int.Parse(dr["PartSN"].ToString()) == sn || int.Parse(dr["PartSN"].ToString()) == (sn - 1))
                        { if (int.Parse(dr["PartSN"].ToString()) == sn) { sn += 1; } }
                        else
                        {
                            iserr = true;
                            break;
                        }
                    }
                }
                #endregion

                #region 檢查IndexSN
                if (!iserr)
                {
                    DataTable dt_tmp = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartSN>=0 order by Source_StationNO_IndexSN");
                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                    {
                        sn = int.Parse(dt_tmp.Rows[0]["Source_StationNO_IndexSN"].ToString());
                        if (sn != 0 && sn != 1) { iserr = true; }
                        else
                        {
                            foreach (DataRow dr in dt_tmp.Rows)
                            {
                                if (int.Parse(dr["Source_StationNO_IndexSN"].ToString()) == sn || int.Parse(dr["Source_StationNO_IndexSN"].ToString()) == (sn - 1))
                                { if (int.Parse(dr["Source_StationNO_IndexSN"].ToString()) == sn) { sn += 1; } }
                                else
                                {
                                    iserr = true;
                                    break;
                                }
                            }
                        }
                    }
                    else { iserr = true; }
                }
                #endregion

                #region Manufacture資料存在
                if (!iserr)
                {
                    DataRow dr_tmp = null;
                    DataTable dt_tmp = db.DB_GetData($"select Source_StationNO from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartSN>=0 and Source_StationNO is not NULL group by Source_StationNO");
                    foreach (DataRow dr in dt_tmp.Rows)
                    {
                        if (db.DB_GetQueryCount($"SELECT StationNO from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'") <= 0)
                        {
                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                            string Config_MutiWO = "0";
                            if (dr_tmp["Station_Type"].ToString().Trim() == "8") { Config_MutiWO = "1"; }
                            db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Manufacture] (ServerId,StationNO,Config_MutiWO) VALUES ('{_Fun.Config.ServerId}','{dr["Source_StationNO"].ToString()}','{Config_MutiWO}')");
                        }
                    }
                }
                #endregion
            }
            else { iserr = true; ERR = "BOM結構無法建立生產流程"; }
            if (iserr)
            {
                if (ERR == "") { ERR = "系統問題,造成不正確"; }
                if (wType == '5')
                {
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{needId}'");
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"工作底稿發出自動生產失敗, 原因:料號:{dr_APS_NeedData["PartNO"].ToString()} {ERR}", true);
                }
                else
                { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='5',StateINFO='{ERR}',UpdateTime=null where Id='{needId}'"); }
                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] where Trigger_NeedId='{needId}'");
                return sTimeLog;
            }
            #endregion

            #region 處理APS_Simulation替代料
            //###???未寫
            #endregion

            //###??? 判斷是否有共用製程(共用線), 選擇適用生產線 與判斷工站負荷選擇使用何工站

            #region 寫入計算每項歷史計算時間   寫入APS_Simulation 有效CT[有含WT] 或 AC01領料時間
            dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartType='0' order by PartSN desc");
            //###???要依Class抓取EfficientCT
            if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
            {
                DataRow dr_Material = null;
                DataRow dr_tmp = null;
                foreach (DataRow dr in dt_Simulation.Rows)
                {
                    if (int.Parse(dr["PartSN"].ToString()) >= 0 && (!dr.IsNull("Source_StationNO") && (dr["Class"].ToString() == "4" || dr["Class"].ToString() == "5")))
                    {
                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_EfficientSCT] WHERE ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}' and PP_Name='{dr["Apply_PP_Name"].ToString()}' and PartNO='{dr["PartNO"].ToString()}' and IndexSN={dr["Source_StationNO_IndexSN"].ToString()}");
                        if (dr_tmp != null && int.Parse(dr_tmp["SCT"].ToString()) > 0)
                        {
                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set Math_EfficientCT={dr_tmp["SCT"].ToString()} where SimulationId='{dr["SimulationId"].ToString()}'");
                        }
                        else
                        {
                            if (dr["Source_StationNO"].ToString() == _Fun.Config.OutPackStationName)
                            {
                                string mfData = "";
                                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr["NeedId"].ToString()}' and Apply_StationNO='{dr["Source_StationNO"].ToString()}' and IndexSN='{dr["Source_StationNO_IndexSN"].ToString()}' and PartSN={(int.Parse(dr["PartSN"].ToString()) + 1).ToString()}");
                                if (dr_tmp != null)
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where Id='{dr_tmp["Apply_BOMId"].ToString()}' and PartNO='{dr["PartNO"].ToString()}'");
                                    if (dr_tmp != null)
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT MFNO FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{dr_tmp["Apply_PP_Name"].ToString()}' and Apply_PartNO='{dr_tmp["Apply_PartNO"].ToString()}' and StationNO='{dr_tmp["Apply_StationNO"].ToString()}' and IndexSN={dr_tmp["IndexSN"].ToString()}");
                                        if (dr_tmp != null && !dr_tmp.IsNull("MFNO") && dr_tmp["MFNO"].ToString() != "")
                                        { mfData = dr_tmp["MFNO"].ToString(); }
                                    }
                                }
                                if (mfData != "") { mfData = $" and MFNO='{mfData}'"; }
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select ROUND(sum(EfficientCycleTime)*sum(CountQTY)/sum(CountQTY),0) as CT from SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{dr["Apply_PP_Name"].ToString()}' and StationNO='{_Fun.Config.OutPackStationName}' and IndexSN={dr["Source_StationNO_IndexSN"].ToString()} and Sub_PartNO='{dr["PartNO"].ToString()}' {mfData} group by Sub_PartNO");
                                if (dr_tmp != null && !dr_tmp.IsNull("CT") && dr_tmp["CT"].ToString() != "" && dr_tmp["CT"].ToString() != "0")
                                { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set Math_EfficientCT={dr_tmp["CT"].ToString()} where SimulationId='{dr["SimulationId"].ToString()}'"); }

                            }
                            else
                            {
                                //成品or半成品
                                tmp_station = $"='{dr["Source_StationNO"].ToString()}'";
                                if (!dr.IsNull("StationNO_Merge") && dr["StationNO_Merge"].ToString().Trim() != "")
                                {
                                    tmp_station = dr["StationNO_Merge"].ToString().Trim().Substring(0, dr["StationNO_Merge"].ToString().Trim().Length - 1);
                                    tmp_station = $" in ('{tmp_station.Replace(",", "','")}')";
                                }
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select ROUND(sum(EfficientCycleTime)*sum(CountQTY)/sum(CountQTY),0) as CT from SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO{tmp_station}  and DOCNO='' and Sub_PartNO='{dr["PartNO"].ToString()}' and IndexSN={dr["Source_StationNO_IndexSN"].ToString()} group by Sub_PartNO");
                                if (dr_tmp != null && !dr_tmp.IsNull("CT") && dr_tmp["CT"].ToString() != "" && dr_tmp["CT"].ToString() != "0")
                                { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set Math_EfficientCT={dr_tmp["CT"].ToString()} where SimulationId='{dr["SimulationId"].ToString()}'"); }
                            }
                        }
                    }
                    else
                    {
                        //###??? DOCNO 暫時寫死 AC01
                        dr_tmp = db.DB_GetFirstDataByDataRow($"select ROUND(sum(EfficientCycleTime)*sum(CountQTY)/sum(CountQTY),0) as CT from SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and DOCNO='AC01' and Sub_PartNO='{dr["PartNO"].ToString()}' group by Sub_PartNO");
                        if (dr_tmp != null && !dr_tmp.IsNull("CT") && dr_tmp["CT"].ToString() != "" && dr_tmp["CT"].ToString() != "0")
                        { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set Math_EfficientCT={dr_tmp["CT"].ToString()} where SimulationId='{dr["SimulationId"].ToString()}'"); }
                    }
                }
            }
            else { return DateTime.Now; }
            #endregion

            #region 計算倉庫現在 與 未來可用量 與 寫TotalStockII
            if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
            {
                DataTable dt_tmp = null;
                DataRow dr_tmp = null;
                int int_tmp = 0;//需求量
                int keepQTY = 0;//可用量
                int hasQTY = 0;

                foreach (DataRow dr in dt_Simulation.Rows)
                {
                    if (!bool.Parse(dr["OutPackType"].ToString()))
                    {
                        #region 若本站料與上一站生產料號相同則continue;
                        if (!dr.IsNull("Source_StationNO") && dr["Source_StationNO"].ToString() != "" && int.Parse(dr["PartSN"].ToString()) > 0)
                        {
                            dr_tmp = db.DB_GetFirstDataByDataRow($"select PartNO from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartSN={(int.Parse(dr["PartSN"].ToString()) - 1).ToString()} and PartNO='{dr["PartNO"].ToString()}' and (Class='4' or Class='5') and Source_StationNO is not null");
                            if (dr_tmp != null) { continue; }
                        }
                        #endregion

                        #region 目前線上在製有多出的量(產出>NeedQTY),且未開入庫單據
                        dr_tmp = db.DB_GetFirstDataByDataRow($"select sum(a.Detail_QTY-a.Next_StationQTY-a.NeedQTY) as surplusQTY from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a,SoftNetSYSDB.[dbo].[APS_Simulation] as b,SoftNetSYSDB.[dbo].[APS_NeedData] as c where c.ServerId='{_Fun.Config.ServerId}' and c.State='6' and a.NeedId=c.Id and a.PartNO='{dr["PartNO"].ToString()}' and a.SimulationId=b.SimulationId and b.IsOK=0 and (a.Store_DOCNumberNO=null or a.Store_DOCNumberNO='') and (a.Class='4' or a.Class='5') group by a.PartNO");
                        if (dr_tmp != null && dr_tmp["surplusQTY"].ToString().Trim() != "")
                        {

                            DataRow dr_tmp2 = db.DB_GetFirstDataByDataRow($"select Math_Online_SurplusQTY from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId!='{needId}' and PartNO='{dr["PartNO"].ToString()}' and (Class='4' or Class='5') and Source_StationNO is not null and IsOK='0'");
                            if (dr_tmp2 != null && dr_tmp2["Math_Online_SurplusQTY"].ToString().Trim() != "")
                            {
                                #region 回扣前期已計算過的數量
                                int tmp_i = int.Parse(dr_tmp["surplusQTY"].ToString()) - int.Parse(dr_tmp2["Math_Online_SurplusQTY"].ToString());
                                if (tmp_i > 0)
                                { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set Math_Online_SurplusQTY={tmp_i.ToString()} where SimulationId='{dr["SimulationId"].ToString()}'"); }
                                #endregion
                            }
                            else { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set Math_Online_SurplusQTY={dr_tmp["surplusQTY"].ToString()} where SimulationId='{dr["SimulationId"].ToString()}'"); }
                        }
                        #endregion

                        #region 計算庫存可用量, 並把(需求量+損耗量)先記錄下來 TotalStockII and 紀錄 目前倉處可用量Math_TotalStock_HasUseQTY
                        if (int.Parse(dr["PartSN"].ToString()) > 0)
                        {
                            hasQTY = 0;
                            dt_tmp = db.DB_GetData($"select a.StoreNO,a.StoreSpacesNO,a.PartNO,a.Id,b.Simulation_FirstNO,sum(a.QTY) as qty from SoftNetMainDB.[dbo].TotalStock as a,dbo.[Store] as b where a.Class!='虛擬倉' and b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{dr["PartNO"].ToString()}' and a.StoreNO=b.StoreNO group by a.StoreNO,a.StoreSpacesNO,a.PartNO,a.Id,b.Simulation_FirstNO order by b.Simulation_FirstNO");
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                int_tmp = int.Parse(dr["NeedQTY"].ToString()) + int.Parse(dr["SafeQTY"].ToString());

                                #region 加生產損耗率值
                                //###???計算損耗率加數量
                                #endregion

                                if (dt_tmp != null)
                                {
                                    foreach (DataRow d in dt_tmp.Rows)
                                    {
                                        #region 寫TotalStockII
                                        keepQTY = int.Parse(d["qty"].ToString());//針對目前 d 的庫存可用量
                                        if (keepQTY > 0)
                                        {
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"select sum(a.KeepQTY+a.OverQTY) as Kqty from SoftNetMainDB.[dbo].[TotalStockII] as a where a.Id='{d["Id"].ToString()}'");
                                            if (dr_tmp != null && !dr_tmp.IsNull("Kqty") && dr_tmp["Kqty"].ToString().Trim() != "") { keepQTY -= int.Parse(dr_tmp["Kqty"].ToString()); }
                                            if (keepQTY <= 0)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                if (int_tmp >= keepQTY) //需求>=可被Keep量
                                                {
                                                    int_tmp -= keepQTY;
                                                    hasQTY += keepQTY;
                                                    db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII] (Id,NeedId,SimulationId,KeepQTY) VALUES ('{d["Id"].ToString()}','{dr["NeedId"].ToString()}','{dr["SimulationId"].ToString()}',{keepQTY.ToString()})");
                                                }
                                                else
                                                {
                                                    hasQTY += int_tmp;
                                                    db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStockII] (Id,NeedId,SimulationId,KeepQTY) VALUES ('{d["Id"].ToString()}','{dr["NeedId"].ToString()}','{dr["SimulationId"].ToString()}',{int_tmp.ToString()})");
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        { continue; }
                                        #endregion
                                    }
                                }
                            }
                            db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set Math_TotalStock_HasUseQTY={hasQTY.ToString()} where SimulationId='{dr["SimulationId"].ToString()}'");
                        }
                        #endregion
                    }
                }
            }
            #endregion

            #region 計算其他單據IsOK欄位=0 (單據的空中數量)  //###??? Math_Online_SurplusQTY 程式斯乎沒用到, 日後要查一下, 目前這段程式有問題
            int surplusQTY = 0;
            //進貨 DOC1Buy
            DataTable dt_DOC = db.DB_GetData($@"select a.DOCNumberNO,a.Id,a.PartNO,a.Unit,a.QTY from SoftNetMainDB.[dbo].[DOC1BuyII] as a right join SoftNetMainDB.[dbo].[DOC1Buy] as b on a.DOCNumberNO=b.DOCNumberNO and b.FlowStatus='0' where b.ServerId='{_Fun.Config.ServerId}' and a.IsOK='0' and a.ArrivalDate<='{need_dDay.ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by a.DOCNumberNO,a.Id");
            if (dt_DOC != null && dt_DOC.Rows.Count > 0)
            {
                foreach (DataRow d in dt_DOC.Rows)
                {
                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC2SalesII] as a right join SoftNetMainDB.[dbo].[DOC2Sales] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC3stockII] as a right join SoftNetMainDB.[dbo].[DOC3stock] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC4ProductionII] as a right join SoftNetMainDB.[dbo].[DOC4Production] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC5OUTII] as a right join SoftNetMainDB.[dbo].[DOC5OUT] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    surplusQTY += int.Parse(d["QTY"].ToString());
                }
            }
            //出貨 DOC2Sales
            dt_DOC = db.DB_GetData($@"select a.DOCNumberNO,a.Id,a.PartNO,a.Unit,a.QTY from SoftNetMainDB.[dbo].[DOC2SalesII] as a right join SoftNetMainDB.[dbo].[DOC2Sales] as b on a.DOCNumberNO=b.DOCNumberNO and b.FlowStatus='0' where b.ServerId='{_Fun.Config.ServerId}' and a.IsOK='0' and a.ArrivalDate<='{need_dDay.ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by a.DOCNumberNO,a.Id");
            if (dt_DOC != null && dt_DOC.Rows.Count > 0)
            {
                foreach (DataRow d in dt_DOC.Rows)
                {
                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC1BuyII] as a right join SoftNetMainDB.[dbo].[DOC1Buy] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC3stockII] as a right join SoftNetMainDB.[dbo].[DOC3stock] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC4ProductionII] as a right join SoftNetMainDB.[dbo].[DOC4Production] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC5OUTII] as a right join SoftNetMainDB.[dbo].[DOC5OUT] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    surplusQTY -= int.Parse(d["QTY"].ToString());
                }
            }
            //存貨 DOC3stock
            dt_DOC = db.DB_GetData($@"select a.DOCNumberNO,a.Id,a.PartNO,a.Unit,a.QTY,b.DOCType from SoftNetMainDB.[dbo].[DOC3stockII] as a right join SoftNetMainDB.[dbo].[DOC3stock] as b on a.DOCNumberNO=b.DOCNumberNO and b.FlowStatus='0' where b.ServerId='{_Fun.Config.ServerId}' and a.IsOK='0' and a.ArrivalDate<='{need_dDay.ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by a.DOCNumberNO,a.Id");
            if (dt_DOC != null && dt_DOC.Rows.Count > 0)
            {
                foreach (DataRow d in dt_DOC.Rows)
                {
                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC1BuyII] as a right join SoftNetMainDB.[dbo].[DOC1Buy] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC3stockII] as a right join SoftNetMainDB.[dbo].[DOC3stock] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC4ProductionII] as a right join SoftNetMainDB.[dbo].[DOC4Production] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC5OUTII] as a right join SoftNetMainDB.[dbo].[DOC5OUT] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    switch (d["DOCType"].ToString())
                    {
                        case "3": surplusQTY -= int.Parse(d["QTY"].ToString()); break;
                        case "4": surplusQTY += int.Parse(d["QTY"].ToString()); break;
                    }

                }
            }
            //生產 DOC4Production
            dt_DOC = db.DB_GetData($@"select a.DOCNumberNO,a.Id,a.PartNO,a.Unit,a.QTY from SoftNetMainDB.[dbo].[DOC4ProductionII] as a right join SoftNetMainDB.[dbo].[DOC4Production] as b on a.DOCNumberNO=b.DOCNumberNO and b.FlowStatus='0' where b.ServerId='{_Fun.Config.ServerId}' and a.IsOK='0' and a.ArrivalDate<='{need_dDay.ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by a.DOCNumberNO,a.Id");
            if (dt_DOC != null && dt_DOC.Rows.Count > 0)
            {
                foreach (DataRow d in dt_DOC.Rows)
                {
                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC1BuyII] as a right join SoftNetMainDB.[dbo].[DOC1Buy] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC3stockII] as a right join SoftNetMainDB.[dbo].[DOC3stock] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC4ProductionII] as a right join SoftNetMainDB.[dbo].[DOC4Production] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC5OUTII] as a right join SoftNetMainDB.[dbo].[DOC5OUT] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    surplusQTY += int.Parse(d["QTY"].ToString());
                }
            }
            //生產 DOC5OUT
            dt_DOC = db.DB_GetData($@"select a.DOCNumberNO,a.Id,a.PartNO,a.Unit,a.QTY from SoftNetMainDB.[dbo].[DOC5OUTII] as a right join SoftNetMainDB.[dbo].[DOC5OUT] as b on a.DOCNumberNO=b.DOCNumberNO and b.FlowStatus='0' where b.ServerId='{_Fun.Config.ServerId}' and a.IsOK='0' and a.ArrivalDate<='{need_dDay.ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by a.DOCNumberNO,a.Id");
            if (dt_DOC != null && dt_DOC.Rows.Count > 0)
            {
                foreach (DataRow d in dt_DOC.Rows)
                {
                    DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC1BuyII] as a right join SoftNetMainDB.[dbo].[DOC1Buy] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC3stockII] as a right join SoftNetMainDB.[dbo].[DOC3stock] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC4ProductionII] as a right join SoftNetMainDB.[dbo].[DOC4Production] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    tmp_dr = db.DB_GetFirstDataByDataRow($"select a.* from SoftNetMainDB.[dbo].[DOC5OUTII] as a right join SoftNetMainDB.[dbo].[DOC5OUT] as b on a.DOCNumberNO=b.DOCNumberNO and b.SourceNO='{d["SourceNO"].ToString()}' where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO='{d["PartNO"].ToString()}' order by a.DOCNumberNO,a.Id");
                    if (tmp_dr != null) { continue; }
                    surplusQTY += int.Parse(d["QTY"].ToString());
                }
            }

            #endregion

            #region 計算每項需要時間  用 Math_EfficientCT 或 Math_StandardCT * 需求量  = 取得需求秒數 寫入 Math_UseTime (ps=Math_UseTime是總時間,沒除以合併站)
            {
                #region 寫 Math_UseTime 若有PP_EfficientSCT(人為定義的 Math_EfficientCT)
                dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartType='0' order by PartSN desc");
                if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
                {
                    //###??? 委外加工的 Math_EfficientCT 上面可能沒處裡
                    int ect = 0;
                    int wct = 0;
                    int sct = 0;
                    int needqty = 0;
                    int hasuseqty = 0;
                    foreach (DataRow dr in dt_Simulation.Rows)
                    {
                        ect = int.Parse(dr["Math_EfficientCT"].ToString());
                        sct = int.Parse(dr["Math_StandardCT"].ToString());
                        if (int.Parse(dr["PartSN"].ToString()) >= 0 && (!dr.IsNull("Source_StationNO") && (dr["Class"].ToString() == "4" || dr["Class"].ToString() == "5")))
                        {
                            needqty = int.Parse(dr["NeedQTY"].ToString()) + int.Parse(dr["SafeQTY"].ToString());
                            hasuseqty = int.Parse(dr["Math_TotalStock_HasUseQTY"].ToString()) + int.Parse(dr["Math_Online_SurplusQTY"].ToString());
                            if (hasuseqty >= needqty)
                            {
                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET Math_UseTime=600 where SimulationId='{dr["SimulationId"].ToString()}'");
                                continue;
                            }
                            wct = int.Parse(dr["Math_EfficientWT"].ToString());
                            if ((ect + wct) != 0)
                            { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET Math_UseTime={((ect + wct) * (needqty - hasuseqty)).ToString()} where SimulationId='{dr["SimulationId"].ToString()}'"); }
                            else if (sct != 0)
                            { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET Math_UseTime={(sct * (needqty - hasuseqty)).ToString()} where SimulationId='{dr["SimulationId"].ToString()}'"); }
                            else
                            {
                                string tmp = "";
                                if (ect == 0) { tmp = ",Math_EfficientCT=60"; }
                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET Math_UseTime={(600 * (needqty - hasuseqty)).ToString()}{tmp} where SimulationId='{dr["SimulationId"].ToString()}'");
                            }
                        }
                        else
                        {
                            if (ect > 0) { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET Math_UseTime={ect} where SimulationId='{dr["SimulationId"].ToString()}'"); }
                            else if (sct > 0) { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET Math_UseTime={sct} where SimulationId='{dr["SimulationId"].ToString()}'"); }
                            else { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET Math_UseTime=600 where SimulationId='{dr["SimulationId"].ToString()}'"); }

                        }
                    }
                }
                #endregion

                #region 計算多工站的工作可能人數 修改Math_UseOPCount
                int tmp_sTot = 0;
                DataRow dr_tmp = null;
                DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartType='0' and IsOK='0' and PartSN>=0 and (Class='4' or Class='5') and Source_StationNO is not null order by PartSN,Class");
                foreach (DataRow dr in dt.Rows)
                {
                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                    if (dr_tmp != null && dr_tmp["Station_Type"].ToString() == "8")
                    {
                        //tmp_sTot = int.Parse(dr["Math_UseTime"].ToString());
                        dr_tmp = db.DB_GetFirstDataByDataRow($"select FLOOR(sum(OP_Count)/COUNT(*)) as TOTData from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}' and PartNO='{dr["PartNO"].ToString()}' group by ServerId,StationNO,PartNO");
                        if (dr_tmp != null && !dr_tmp.IsNull("TOTData") && dr_tmp["TOTData"].ToString() != "" && int.Parse(dr_tmp["TOTData"].ToString()) != 0)
                        {
                            //tmp_sTot = Convert.ToInt32(Math.Floor(tmp_sTot / decimal.Parse(dr_tmp["TOTData"].ToString())));
                            db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set Math_UseOPCount={dr_tmp["TOTData"].ToString()} where NeedId='{needId}' and SimulationId='and {dr["SimulationId"].ToString()}'");
                        }
                    }
                }
                #endregion
            }
            #endregion

            #region 足量生產件扣除計畫 IsOK='1' 讓 IsDEL='1'
            {
                int needqty = 0;
                int hasuseqty = 0;
                string sID = "";
                DataRow dr_tmp = null;
                DataRow dr_M = null;
                List<string> compare_Data = new List<string>();
                dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartType='0' and PartSN>=0 order by PartSN,Class desc");
                if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt_Simulation.Rows)
                    {
                        needqty = int.Parse(dr["NeedQTY"].ToString()) + int.Parse(dr["SafeQTY"].ToString());
                        hasuseqty = int.Parse(dr["Math_TotalStock_HasUseQTY"].ToString()) + int.Parse(dr["Math_Online_SurplusQTY"].ToString());
                        if (hasuseqty >= needqty)
                        {
                            if (dr["PartSN"].ToString().Trim() == "0")
                            {
                                if (args.ARGs[12]) { break; }
                                else
                                {
                                    #region 完全足量
                                    sID = "";
                                    foreach (DataRow dr2 in dt_Simulation.Rows)
                                    {
                                        if (dr2["PartSN"].ToString().Trim() != "0")
                                        {
                                            if (sID == "") { sID = $"'{dr2["SimulationId"].ToString()}'"; }
                                            else { sID = $",'{dr2["SimulationId"].ToString()}'"; }
                                        }
                                    }
                                    if (sID != "")
                                    {
                                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needId}' and SimulationId in ({sID})");
                                        db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[APS_Simulation] set IsDEL='1' where NeedId='{needId}' and SimulationId in ({sID})");
                                    }
                                    db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='8',UpdateTime='{DateTime.Now.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss.fff")}' where NeedId='{needId}'");
                                    #endregion
                                    break;
                                }
                            }
                            else
                            {
                                if (int.Parse(dr["PartSN"].ToString()) > 0 && (!dr.IsNull("Source_StationNO") && (dr["Class"].ToString() == "4" || dr["Class"].ToString() == "5")))
                                {
                                    if (args.ARGs[12]) { continue; }
                                    if (!compare_Data.Contains(dr["SimulationId"].ToString())) { compare_Data.Add(dr["SimulationId"].ToString()); }

                                    #region 將所有下階 取得BOM結構 ㄝ一起 IsOK='1'
                                    DataTable dt_next = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and Apply_PP_Name='{dr["Apply_PP_Name"].ToString()}' and Apply_StationNO='{dr["Source_StationNO"].ToString()}' and IndexSN='{dr["Source_StationNO_IndexSN"].ToString()}' and PartSN>=0 order by PartSN,Class desc");
                                    if (dt_next != null && dt_next.Rows.Count > 0)
                                    {
                                        RecursiveUPSID_thread(db, needId, dt_next, ref compare_Data);
                                    }

                                    #endregion
                                }
                            }
                        }
                    }
                    if (compare_Data.Count > 0)
                    {
                        string tmp = string.Join("','", compare_Data);
                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET Math_UseTime='0',IsOK='1',IsDEL='1' where SimulationId in ('{tmp}')");
                        #region 清除不需要的TotalStockII KeepQTY
                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[TotalStockII] where SimulationId in ('{tmp}')");
                        #endregion
                    }
                }
            }
            #endregion

            #region 生產件有庫存,但量不足 修正往下階的Math_TotalStock_HasUseQTY, Math_Online_SurplusQTY值, TotalStockII的KeepQTY值
            {

                dt_Simulation = db.DB_GetData($"select *,((NeedQTY+SafeQTY)-(Math_TotalStock_HasUseQTY+Math_Online_SurplusQTY)) as sQTY from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartType='0' and IsOK='0' and PartSN>=0 and (Class='4' or Class='5') and Source_StationNO is not null and Math_TotalStock_HasUseQTY!=0 and NeedQTY>(Math_TotalStock_HasUseQTY+Math_Online_SurplusQTY) order by PartSN,Class desc");
                if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
                {
                    DataRow dr_tmp = null;
                    DataRow dr_bom = null;
                    DataRow dr_M = null;
                    int needQTY = 0;
                    string Apply_PartNO = "";
                    string today = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
                    foreach (DataRow dr in dt_Simulation.Rows)
                    {
                        needQTY = int.Parse(dr["sQTY"].ToString());//重新計算扣除量
                        if (needQTY <= 0) { continue; }
                        //取得有效BOM的ID
                        string Main_Item = "0"; if (int.Parse(dr["PartSN"].ToString()) == 0) { Main_Item = "1"; }
                        dr_bom = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB].[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and Id='{dr["Apply_BOMId"].ToString()}'");//這裡的Apply_BOMId,不管上下階,只為拿Apply_PartNO
                        Apply_PartNO = dr_bom["Apply_PartNO"].ToString();
                        dr_M = db.DB_GetFirstDataByDataRow($"select a.Id,a.PartNO,a.Apply_PP_Name,a.Apply_StationNO,a.IsEnd,a.IndexSN,a.Station_Custom_IndexSN,b.Class from SoftNetMainDB.[dbo].[BOM] as a,dbo.[Material] as b where a.Main_Item='{Main_Item}' and b.ServerId='{_Fun.Config.ServerId}' and a.Apply_PP_Name='{dr["Apply_PP_Name"].ToString()}' and a.Apply_PartNO='{Apply_PartNO}' and a.PartNO='{dr["PartNO"].ToString()}' and a.Apply_StationNO='{dr["Source_StationNO"].ToString()}' and a.IndexSN={dr["Source_StationNO_IndexSN"].ToString()} and a.Station_Custom_IndexSN='{dr["Source_StationNO_Custom_IndexSN"].ToString()}' and a.PartNO=b.PartNO and a.EffectiveDate<='{today}' and a.ExpiryDate>='{today}' order by EffectiveDate desc");
                        if (dr_M != null)
                        {
                            //取得第一層結構
                            DataTable dt_BOMII = db.DB_GetData($"select a.Id,a.PartNO,a.BOMQTY,a.Class,b.PartType,b.StoreSTime,b.SafeQTY from SoftNetMainDB.[dbo].[BOMII] as a,dbo.[Material] as b where b.ServerId='{_Fun.Config.ServerId}' and  a.PartNO=b.PartNO and BOMId='{dr_M["Id"].ToString()}' order by sn");
                            if (dt_BOMII != null)
                            {
                                string tmp_s = "";
                                int tmp_i = 0;
                                foreach (DataRow dr0 in dt_BOMII.Rows)
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where IsOK='0' and NeedId='{needId}' and PartNO='{dr0["PartNO"].ToString()}' and Master_PartNO='{dr["PartNO"].ToString()}' and Apply_StationNO='{dr["Source_StationNO"].ToString()}' and IndexSN={dr["Source_StationNO_IndexSN"].ToString()} and Station_Custom_IndexSN='{dr["Source_StationNO_Custom_IndexSN"].ToString()}'");
                                    if (dr_tmp != null)
                                    {
                                        tmp_s = "";
                                        tmp_i = int.Parse(dr0["BOMQTY"].ToString()) * needQTY;
                                        if (int.Parse(dr_tmp["Math_TotalStock_HasUseQTY"].ToString()) >= tmp_i)
                                        {
                                            tmp_s = $",Math_TotalStock_HasUseQTY-={tmp_i}";
                                        }
                                        else
                                        {
                                            tmp_s = $",Math_TotalStock_HasUseQTY=0";
                                            tmp_i -= int.Parse(dr_tmp["Math_TotalStock_HasUseQTY"].ToString());
                                        }
                                        DataTable dt_TotalStockII = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[TotalStockII] where SimulationId='{dr_tmp["SimulationId"].ToString()}' and KeepQTY>0 order by Id");
                                        if (dt_TotalStockII != null && dt_TotalStockII.Rows.Count > 0)
                                        {
                                            int tmp_i2 = tmp_i;
                                            foreach (DataRow dr2 in dt_TotalStockII.Rows)
                                            {
                                                if (tmp_i2 >= int.Parse(dr2["KeepQTY"].ToString()))
                                                {
                                                    tmp_i2 -= int.Parse(dr2["KeepQTY"].ToString());
                                                    db.DB_SetData($"delete from SoftNetMainDB.[dbo].[TotalStockII] where Id='{dr2["Id"].ToString()}' and NeedId='{dr2["NeedId"].ToString()}' and  SimulationId='{dr2["SimulationId"].ToString()}'");
                                                }
                                                else
                                                {
                                                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] SET KeepQTY-={tmp_i2} where Id='{dr2["Id"].ToString()}' and NeedId='{dr2["NeedId"].ToString()}' and  SimulationId='{dr2["SimulationId"].ToString()}'");
                                                }
                                            }
                                        }

                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET Math_Online_SurplusQTY+={tmp_i}{tmp_s} where SimulationId='{dr_tmp["SimulationId"].ToString()}'");
                                        if (!dr_tmp.IsNull("Source_StationNO") && (dr0["Class"].ToString() == "4" || dr0["Class"].ToString() == "5"))
                                        {
                                            RecursiveBOMIII_RunSetSimulation_thread_2(db, today, needId, dr_tmp, dr0, needQTY, dr["Apply_PP_Name"].ToString(), Apply_PartNO);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            #endregion

            #region 從最下階往上填模擬得到的日期  含搭配行事曆(可用工時) , 並記錄 APS_PartNOTimeNote 與 APS_WorkTimeNote 的資料
            {
                dt_Simulation = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartType='0' and IsOK='0' and PartSN>=0 order by PartSN desc,Class desc");
                {
                    //###???開始時間 BufferTime
                    if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
                    {
                        DataRow dr_PP_Station = null;
                        #region 排生產得到的第一個日期
                        DateTime simulationtime = DateTime.Now;
                        DateTime simulationtime_isARGs10 = DateTime.Now;
                        bool isARGs10 = false;
                        DataRow dr_tmp;
                        int isARGs10_offset = 15;//###??? 10將來改參數
                        if (args.ARGs[10])//將simulationtime的時間改為目前班別的起始時間,並不考慮第一階領料時間
                        {
                            #region 即時排定
                            isARGs10 = true;
                            DateTime intime = DateTime.Now.AddMinutes(isARGs10_offset);
                            DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_APS_NeedData["CalendarName"].ToString()}' and Holiday>='{intime.ToString("yyyy-MM-dd")}' order by Holiday");
                            if (dt == null || dt.Rows.Count <= 0)
                            {
                                errINFO = "行事曆日期不足計算.";
                                ERR = errINFO;
                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='5',StateINFO='{errINFO}',UpdateTime=null where Id='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{needId}'");
                                db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] where Trigger_NeedId='{needId}'");
                                System.Threading.Tasks.Task task = _Log.ErrorAsync($"工作底稿發出自動生產失敗, 原因:料號:{dr_APS_NeedData["PartNO"].ToString()} {errINFO}", true);
                                return sTimeLog;
                            }
                            DateTime etime = DateTime.Now;
                            DateTime stime2 = DateTime.Now;
                            foreach (DataRow dr in dt.Rows)
                            {
                                if (Convert.ToDateTime(dr["Holiday"]).ToString("yyyy-MM-dd") != intime.ToString("yyyy-MM-dd"))
                                {
                                    if (Convert.ToDateTime(dr["Holiday"]) < intime) { break; }
                                    else
                                    {
                                        stime2 = Convert.ToDateTime(dr["Holiday"]);
                                        stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, DateTime.Now.Hour, DateTime.Now.Minute, 0, 0);
                                        if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                                        { stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                                        else if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                                        { stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                                        else if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                                        { stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                                        else if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                                        { stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                                        intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, stime2.Hour, stime2.Minute, 0, 0);
                                    }
                                }
                                #region 確認時段
                                string[] comp = null;
                                if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                                {

                                    comp = dr["Shift_Morning"].ToString().Trim().Split(',');
                                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                                    if (etime >= intime && intime >= stime2)
                                    {
                                        simulationtime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                        simulationtime_isARGs10 = intime;
                                        if (simulationtime_isARGs10 > etime)
                                        { simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0); }
                                        break;
                                    }
                                    else
                                    {
                                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                        if (etime >= intime && intime >= stime2)
                                        {
                                            simulationtime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                            simulationtime_isARGs10 = intime;
                                            if (simulationtime_isARGs10 > etime)
                                            {
                                                if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                                                {
                                                    comp = dr["Shift_Afternoon"].ToString().Trim().Split(',');
                                                    simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                                }
                                                else if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                                                {
                                                    comp = dr["Shift_Night"].ToString().Trim().Split(',');
                                                    simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                                }
                                                else if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                                                {
                                                    int be_Hour = int.Parse(comp[3].Split(':')[0]);
                                                    comp = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                                                    if (be_Hour > int.Parse(comp[0].Split(':')[0]))
                                                    { simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0).AddDays(1); }
                                                    else
                                                    { simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0); }
                                                }
                                                else { continue; }
                                            }
                                            break;
                                        }
                                    }
                                }
                                if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                                {
                                    comp = dr["Shift_Afternoon"].ToString().Trim().Split(',');
                                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                                    if (etime >= intime && intime >= stime2)
                                    {
                                        simulationtime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                        simulationtime_isARGs10 = intime;
                                        if (simulationtime_isARGs10 > etime)
                                        { simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0); }
                                        break;
                                    }
                                    else
                                    {
                                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                        if (etime >= intime && intime >= stime2)
                                        {
                                            simulationtime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                            simulationtime_isARGs10 = intime;
                                            if (simulationtime_isARGs10 > etime)
                                            {
                                                if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                                                {
                                                    comp = dr["Shift_Night"].ToString().Trim().Split(',');
                                                    simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                                }
                                                else if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                                                {
                                                    int be_Hour = int.Parse(comp[3].Split(':')[0]);
                                                    comp = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                                                    if (be_Hour > int.Parse(comp[0].Split(':')[0]))
                                                    { simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0).AddDays(1); }
                                                    else
                                                    { simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0); }
                                                }
                                                else { continue; }
                                            }
                                            break;
                                        }
                                    }
                                }
                                if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                                {
                                    comp = dr["Shift_Night"].ToString().Trim().Split(',');
                                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                                    if (etime >= intime && intime >= stime2)
                                    {
                                        simulationtime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                        simulationtime_isARGs10 = intime;
                                        if (simulationtime_isARGs10 > etime)
                                        { simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0); }
                                        break;
                                    }
                                    else
                                    {
                                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                        if (etime >= intime && intime >= stime2)
                                        {
                                            simulationtime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                            simulationtime_isARGs10 = intime;
                                            if (simulationtime_isARGs10 > etime)
                                            {
                                                if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                                                {
                                                    int be_Hour = int.Parse(comp[3].Split(':')[0]);
                                                    comp = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                                                    if (be_Hour > int.Parse(comp[0].Split(':')[0]))
                                                    { simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0).AddDays(1); }
                                                    else
                                                    { simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0); }
                                                }
                                                else { continue; }
                                            }
                                            break;
                                        }
                                    }
                                }
                                if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                                {
                                    comp = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                                    if (stime2.Hour > etime.Hour) { etime = etime.AddDays(1); }
                                    if (etime >= intime && intime >= stime2)
                                    {
                                        simulationtime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                        simulationtime_isARGs10 = intime;
                                        if (simulationtime_isARGs10 > etime)
                                        { simulationtime_isARGs10 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0); }
                                        break;
                                    }
                                    else
                                    {
                                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                        if (etime >= intime && intime >= stime2)
                                        {
                                            simulationtime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                            simulationtime_isARGs10 = intime;
                                            if (simulationtime_isARGs10 > etime)
                                            {
                                                continue;
                                            }
                                        }
                                        break;
                                    }
                                }
                                #endregion
                            }
                            #endregion
                        }
                        else
                        {
                            int sTot = 0;//總需求工時
                            DataTable dt = db.DB_GetData($"select SimulationId,StationNO_Merge,sum(Math_UseTime/Math_UseOPCount) as sTot from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartType='0' and IsOK='0' and PartSN>=0 and  (Class='4' or Class='5') and Source_StationNO is not null group by SimulationId,StationNO_Merge");
                            foreach (DataRow dr in dt.Rows)
                            {
                                if (!dr.IsNull("sTot") && dr["sTot"].ToString().Trim() != "")
                                {
                                    if (!dr.IsNull("StationNO_Merge")) { sTot += (int.Parse(dr["sTot"].ToString()) / dr["StationNO_Merge"].ToString().Split(',').Length - 1); }
                                    else { sTot += int.Parse(dr["sTot"].ToString()); }
                                }
                            }

                            //###??? 要加站與站之間的移轉時間
                            //###??? 要考慮合併站的問題

                            #region 從需求日向前統計工站負荷 取得 開始排程日 與 修改Math_UseOPCount
                            if (DateTime.Now.AddSeconds(sTot) > need_dDay) { need_dDay = DateTime.Now.AddSeconds(sTot); }
                            ;
                            DateTime math_dDay = need_dDay;//###??? 將來要考慮是否要加2分法的日期計算
                            DateTime timeNow = DateTime.Now;
                            int tmp_sTot = 0;
                            string[] comp = null;
                            bool is_1 = false;

                            DateTime etime = DateTime.Now;
                            DateTime stime2 = DateTime.Now;
                            dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartType='0' and IsOK='0' and PartSN>=0 and (Class='4' or Class='5') and Source_StationNO is not null order by PartSN,Class");
                            foreach (DataRow dr in dt.Rows)
                            {
                                is_1 = false;
                                tmp_sTot = 0;
                                if (int.Parse(dr["Math_UseTime"].ToString()) > 0)
                                {
                                    tmp_sTot = Convert.ToInt32(Math.Floor(int.Parse(dr["Math_UseTime"].ToString()) / decimal.Parse(dr["Math_UseOPCount"].ToString())));
                                    if (!dr.IsNull("StationNO_Merge")) { tmp_sTot += (tmp_sTot / dr["StationNO_Merge"].ToString().Split(',').Length - 1); }
                                }
                                while (tmp_sTot > 0 && math_dDay > timeNow)
                                {
                                    dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                                    DataTable dt_Calendar = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_PP_Station["CalendarName"].ToString()}' and Holiday<='{math_dDay.ToString("yyyy-MM-dd")}' order by Holiday desc");
                                    if (dt_Calendar == null || dt_Calendar.Rows.Count <= 0) { math_dDay = DateTime.Now.AddMinutes(isARGs10_offset); ; break; }
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select sum(Time1_C) as beTot1,sum(Time2_C) as beTot2,sum(Time3_C) as beTot3,sum(Time4_C) as beTot4,sum(Time1_C+Time2_C+Time3_C+Time4_C) as ALLTot from SoftNetSYSDB.[dbo].APS_WorkTimeNote where StationNO='{dr["Source_StationNO"].ToString()}' and CalendarDate='{math_dDay.ToString("yyyy-MM-dd")}'");
                                    if (dr_tmp != null && !dr_tmp.IsNull("ALLTot") && dr_tmp["ALLTot"].ToString().Trim() != "" && int.Parse(dr_tmp["ALLTot"].ToString().Trim()) != 0)
                                    {
                                        DataRow dr_Calendar = dt_Calendar.Rows[0];
                                        #region 計算APS_WorkTimeNote可用時間
                                        comp = dr_Calendar["Shift_Graveyard"].ToString().Trim().Split(',');
                                        if (!dr_tmp.IsNull("beTot4") && dr_tmp["beTot4"].ToString().Trim() != "" && int.Parse(dr_tmp["beTot4"].ToString().Trim()) != 0)
                                        {
                                            tmp_sTot -= int.Parse(dr_tmp["beTot4"].ToString());
                                            if (tmp_sTot <= 0)
                                            {
                                                string[] comp_Night = dr_Calendar["Shift_Night"].ToString().Trim().Split(',');
                                                if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                                                { math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0).AddDays(1); }
                                                else
                                                { math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0); }
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            if (bool.Parse(dr_Calendar["Flag_Graveyard"].ToString()) == true)
                                            {
                                                if (Convert.ToDateTime(dr_Calendar["Holiday"]).ToString("yyyy-MM-dd") != math_dDay.ToString("yyyy-MM-dd"))
                                                {
                                                    stime2 = Convert.ToDateTime(dr_Calendar["Holiday"].ToString());
                                                    math_dDay = new DateTime(stime2.Year, stime2.Month, stime2.Day, math_dDay.Hour, math_dDay.Minute, 0);
                                                }
                                                stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                if (stime2.Hour > etime.Hour) { etime = etime.AddDays(1); }
                                                tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                if (tmp_sTot <= 0)
                                                {
                                                    string[] comp_Night = dr_Calendar["Shift_Night"].ToString().Trim().Split(',');
                                                    if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                                                    { math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0).AddDays(1); }
                                                    else
                                                    { math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0); }
                                                    break;
                                                }
                                                else
                                                {
                                                    stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                    etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                    if (stime2.Hour > etime.Hour) { etime = etime.AddDays(1); }
                                                    tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                    if (tmp_sTot <= 0)
                                                    {
                                                        string[] comp_Night = dr_Calendar["Shift_Night"].ToString().Trim().Split(',');
                                                        if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                                                        { math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0).AddDays(1); }
                                                        else
                                                        { math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0); }
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                        if (!dr_tmp.IsNull("beTot3") && dr_tmp["beTot3"].ToString().Trim() != "" && int.Parse(dr_tmp["beTot3"].ToString().Trim()) != 0)
                                        {
                                            tmp_sTot -= int.Parse(dr_tmp["beTot3"].ToString());
                                            if (tmp_sTot <= 0)
                                            {
                                                math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            if (bool.Parse(dr_Calendar["Flag_Night"].ToString()) == true)
                                            {
                                                comp = dr_Calendar["Shift_Night"].ToString().Trim().Split(',');
                                                if (Convert.ToDateTime(dr_Calendar["Holiday"]).ToString("yyyy-MM-dd") != math_dDay.ToString("yyyy-MM-dd"))
                                                {
                                                    stime2 = Convert.ToDateTime(dr_Calendar["Holiday"].ToString());
                                                    math_dDay = new DateTime(stime2.Year, stime2.Month, stime2.Day, math_dDay.Hour, math_dDay.Minute, 0);
                                                }
                                                stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                if (stime2.Hour > etime.Hour) { etime = etime.AddDays(1); }
                                                tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                if (tmp_sTot <= 0)
                                                {
                                                    if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                    { math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0).AddDays(1); }
                                                    else
                                                    { math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0); }
                                                    break;
                                                }
                                                else
                                                {
                                                    stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                    etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                    if (stime2.Hour > etime.Hour) { etime = etime.AddDays(1); }
                                                    tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                    if (tmp_sTot <= 0)
                                                    {
                                                        math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                        if (!dr_tmp.IsNull("beTot2") && dr_tmp["beTot2"].ToString().Trim() != "" && int.Parse(dr_tmp["beTot2"].ToString().Trim()) != 0)
                                        {
                                            tmp_sTot -= int.Parse(dr_tmp["beTot2"].ToString());
                                            if (tmp_sTot <= 0)
                                            {
                                                math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            if (bool.Parse(dr_Calendar["Flag_Afternoon"].ToString()) == true)
                                            {
                                                comp = dr_Calendar["Shift_Afternoon"].ToString().Trim().Split(',');
                                                if (Convert.ToDateTime(dr_Calendar["Holiday"]).ToString("yyyy-MM-dd") != math_dDay.ToString("yyyy-MM-dd"))
                                                {
                                                    stime2 = Convert.ToDateTime(dr_Calendar["Holiday"].ToString());
                                                    math_dDay = new DateTime(stime2.Year, stime2.Month, stime2.Day, math_dDay.Hour, math_dDay.Minute, 0);
                                                }
                                                stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                if (tmp_sTot <= 0)
                                                {
                                                    math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                    break;
                                                }
                                                else
                                                {
                                                    stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                    etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                    tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                    if (tmp_sTot <= 0)
                                                    {
                                                        math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                        if (!dr_tmp.IsNull("beTot1") && dr_tmp["beTot1"].ToString().Trim() != "" && int.Parse(dr_tmp["beTot1"].ToString().Trim()) != 0)
                                        {
                                            tmp_sTot -= int.Parse(dr_tmp["beTot1"].ToString());
                                            if (tmp_sTot <= 0)
                                            {
                                                comp = dr_Calendar["Shift_Morning"].ToString().Trim().Split(',');
                                                math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            if (bool.Parse(dr_Calendar["Flag_Morning"].ToString()) == true)
                                            {
                                                comp = dr_Calendar["Shift_Morning"].ToString().Trim().Split(',');
                                                if (Convert.ToDateTime(dr_Calendar["Holiday"]).ToString("yyyy-MM-dd") != math_dDay.ToString("yyyy-MM-dd"))
                                                {
                                                    stime2 = Convert.ToDateTime(dr_Calendar["Holiday"].ToString());
                                                    math_dDay = new DateTime(stime2.Year, stime2.Month, stime2.Day, math_dDay.Hour, math_dDay.Minute, 0);
                                                }
                                                stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                if (tmp_sTot <= 0)
                                                {
                                                    math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                    break;
                                                }
                                                else
                                                {
                                                    stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                    etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                    tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                    if (tmp_sTot <= 0)
                                                    {
                                                        math_dDay = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                        #endregion
                                        is_1 = true;
                                    }
                                    else
                                    {
                                        is_1 = false;
                                        foreach (DataRow dr_Calendar in dt_Calendar.Rows)
                                        {
                                            if (Convert.ToDateTime(dr_Calendar["Holiday"]).ToString("yyyy-MM-dd") != math_dDay.ToString("yyyy-MM-dd"))
                                            {
                                                if (Convert.ToDateTime(dr_Calendar["Holiday"].ToString()) <= timeNow) { math_dDay = timeNow; break; }
                                                else
                                                {
                                                    stime2 = Convert.ToDateTime(dr_Calendar["Holiday"]);
                                                    if (bool.Parse(dr_Calendar["Flag_Graveyard"].ToString()) == true)
                                                    { stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr_Calendar["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_Calendar["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), 0); }
                                                    else if (bool.Parse(dr_Calendar["Flag_Night"].ToString()) == true)
                                                    { stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr_Calendar["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_Calendar["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0); }
                                                    else if (bool.Parse(dr_Calendar["Flag_Afternoon"].ToString()) == true)
                                                    { stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr_Calendar["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_Calendar["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0); }
                                                    else if (bool.Parse(dr_Calendar["Flag_Morning"].ToString()) == true)
                                                    { stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr_Calendar["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_Calendar["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0); }
                                                    math_dDay = new DateTime(stime2.Year, stime2.Month, stime2.Day, stime2.Hour, stime2.Minute, stime2.Second);
                                                }
                                            }
                                            #region 確認時段
                                            if (bool.Parse(dr_Calendar["Flag_Graveyard"].ToString()) == true)
                                            {
                                                comp = dr_Calendar["Shift_Graveyard"].ToString().Trim().Split(',');
                                                stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                if (stime2.Hour > etime.Hour) { etime = etime.AddDays(1); }
                                                tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                if (tmp_sTot <= 0)
                                                {
                                                    if (math_dDay > stime2) { math_dDay = stime2; }
                                                    break;
                                                }
                                            }
                                            if (bool.Parse(dr_Calendar["Flag_Night"].ToString()) == true)
                                            {
                                                comp = dr_Calendar["Shift_Night"].ToString().Trim().Split(',');
                                                stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                if (stime2.Hour > etime.Hour) { etime = etime.AddDays(1); }
                                                tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                if (tmp_sTot <= 0) { if (math_dDay > stime2) { math_dDay = stime2; } break; }
                                            }
                                            if (bool.Parse(dr_Calendar["Flag_Afternoon"].ToString()) == true)
                                            {
                                                comp = dr_Calendar["Shift_Afternoon"].ToString().Trim().Split(',');
                                                stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                if (tmp_sTot <= 0) { if (math_dDay > stime2) { math_dDay = stime2; } break; }
                                            }
                                            if (bool.Parse(dr_Calendar["Flag_Morning"].ToString()) == true)
                                            {
                                                comp = dr_Calendar["Shift_Morning"].ToString().Trim().Split(',');
                                                stime2 = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                etime = new DateTime(math_dDay.Year, math_dDay.Month, math_dDay.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                tmp_sTot -= TimeCompute2Seconds(stime2, etime);
                                                if (tmp_sTot <= 0) { if (math_dDay > stime2) { math_dDay = stime2; } break; }
                                            }
                                            #endregion
                                        }
                                    }
                                    if (is_1) { math_dDay = math_dDay.AddDays(-1); }
                                }

                                #region 扣除領料時間
                                dr_tmp = db.DB_GetFirstDataByDataRow($@"select MAX(Math_UseTime) as MAXTime from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and ((Class!='4' and Class!='5') or Source_StationNO is null) and IsOK='0' and 
                                                                        Master_PartNO='{dr["PartNO"].ToString()}' and Apply_StationNO='{dr["Source_StationNO"].ToString()}' and IndexSN={dr["Source_StationNO_IndexSN"].ToString()}");
                                if (dr_tmp != null && !dr_tmp.IsNull("MAXTime") && dr_tmp["MAXTime"].ToString().Trim() != "" && int.Parse(dr_tmp["MAXTime"].ToString().Trim()) != 0)
                                {
                                    math_dDay = math_dDay.AddSeconds(-int.Parse(dr_tmp["MAXTime"].ToString()));
                                }
                                else { math_dDay = math_dDay.AddMinutes(-isARGs10_offset); }
                                #endregion

                                if (timeNow > math_dDay)
                                {
                                    math_dDay = DateTime.Now.AddMinutes(isARGs10_offset);
                                    break;//無法滿足預先牌計畫開始日期
                                }
                            }
                            simulationtime = math_dDay;
                            #endregion

                        }
                        #endregion

                        #region 排生產得到的日期, 並寫入APS_WorkTimeNote(工站負荷[計畫]檔), APS_PartNOTimeNote(料件負荷[計畫]檔)
                        DateTime tmp = DateTime.Now;
                        int sn = int.Parse(dt_Simulation.Rows[0]["PartSN"].ToString());
                        string partNOClass = "";
                        int real_NeedQTY = 0;
                        string calendarName = "";

                        try
                        {
                            string hasOutPackType_Station = "";
                            foreach (DataRow dr in dt_Simulation.Rows)
                            {
                                if (sn != int.Parse(dr["PartSN"].ToString()))
                                {
                                    sn = int.Parse(dr["PartSN"].ToString());
                                    simulationtime = sTimeLog;
                                }
                                partNOClass = dr["Class"].ToString();
                                if (dr["ChangeSimulationClass"].ToString().Trim() != "") { partNOClass = dr["ChangeSimulationClass"].ToString().Trim(); }
                                real_NeedQTY = int.Parse(dr["NeedQTY"].ToString()) + int.Parse(dr["SafeQTY"].ToString()) - int.Parse(dr["Math_TotalStock_HasUseQTY"].ToString()) - int.Parse(dr["Math_Online_SurplusQTY"].ToString());

                                if (dr.IsNull("Source_StationNO") || (partNOClass != "4" && partNOClass != "5") || bool.Parse(dr["OutPackType"].ToString()) || dr["Source_StationNO"].ToString() == _Fun.Config.OutPackStationName)
                                {
                                    #region 非生產件
                                    if (bool.Parse(dr["OutPackType"].ToString())) { hasOutPackType_Station = dr["Source_StationNO"].ToString(); }
                                    if (isARGs10)
                                    {
                                        errINFO = "";
                                        tmp = GetFinallyDate(args, db, needId, dr["SimulationId"].ToString(), "料", hasOutPackType_Station, dr["PartNO"].ToString(), partNOClass, real_NeedQTY.ToString(), dr_APS_NeedData["CalendarName"].ToString(), simulationtime_isARGs10, int.Parse(dr["Math_UseTime"].ToString()), int.Parse(dr["Math_UseOPCount"].ToString()), null, ref errINFO, isARGs10);
                                        if (errINFO != "") { isError = true; ERR = errINFO; break; }
                                    }
                                    else
                                    {
                                        if (bool.Parse(dr["OutPackType"].ToString()))
                                        {
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and Master_PartNO='{dr["PartNO"].ToString()}' and Apply_StationNO='{dr["Source_StationNO"].ToString()}' and IndexSN='{dr["Source_StationNO_IndexSN"].ToString()}' order by SimulationDate desc");
                                            if (dr_tmp != null) { simulationtime = dr_tmp.IsNull("SimulationDate") ? simulationtime : Convert.ToDateTime(dr_tmp["SimulationDate"]); }
                                        }
                                        errINFO = "";
                                        tmp = GetFinallyDate(args, db, needId, dr["SimulationId"].ToString(), "料", hasOutPackType_Station, dr["PartNO"].ToString(), partNOClass, real_NeedQTY.ToString(), dr_APS_NeedData["CalendarName"].ToString(), simulationtime, int.Parse(dr["Math_UseTime"].ToString()), int.Parse(dr["Math_UseOPCount"].ToString()), null, ref errINFO, isARGs10);
                                        if (errINFO != "") { isError = true; ERR = errINFO; break; }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region 生產件
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and Master_PartNO='{dr["PartNO"].ToString()}' and Apply_StationNO='{dr["Source_StationNO"].ToString()}' and IndexSN='{dr["Source_StationNO_IndexSN"].ToString()}' order by SimulationDate desc");
                                    //###???改多線時,會null
                                    if (dr_tmp != null)
                                    {
                                        DateTime indtime = dr_tmp.IsNull("SimulationDate") ? simulationtime : Convert.ToDateTime(dr_tmp["SimulationDate"]);
                                        if (isARGs10)
                                        {
                                            indtime = simulationtime_isARGs10;
                                        }
                                        object stationNO_Merge = null;
                                        if (int.Parse(dr["Math_UseTime"].ToString()) != 0)
                                        {
                                            calendarName = dr_APS_NeedData["CalendarName"].ToString();
                                            dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                                            if (dr_PP_Station != null && dr_PP_Station.IsNull("CalendarName") && dr_PP_Station["CalendarName"].ToString() != "")
                                            { calendarName = dr_PP_Station["CalendarName"].ToString(); }

                                            if (!dr.IsNull("StationNO_Merge")) { stationNO_Merge = dr["StationNO_Merge"]; }
                                            errINFO = "";
                                            tmp = GetFinallyDate(args, db, needId, dr["SimulationId"].ToString(), "站", dr["Source_StationNO"].ToString(), dr["PartNO"].ToString(), partNOClass, real_NeedQTY.ToString(), calendarName, indtime, int.Parse(dr["Math_UseTime"].ToString()), int.Parse(dr["Math_UseOPCount"].ToString()), stationNO_Merge, ref errINFO, isARGs10);
                                            if (errINFO != "") { isError = true; ERR = errINFO; break; }
                                        }
                                        else
                                        {
                                            string stationNO = "";
                                            if (!dr_tmp.IsNull("Apply_StationNO")) { stationNO = dr_tmp["Apply_StationNO"].ToString(); }
                                            if (!dr.IsNull("StationNO_Merge")) { stationNO_Merge = dr["StationNO_Merge"]; }

                                            GetFinallyDate_insert_APS_PartNOTimeNote(db, dr["PartNO"].ToString(), indtime, dr_tmp.IsNull("SimulationDate") ? simulationtime : Convert.ToDateTime(dr_tmp["SimulationDate"]), partNOClass, needId, dr["SimulationId"].ToString(), "0", stationNO, stationNO_Merge, isARGs10);
                                        }
                                        if (isARGs10)
                                        {
                                            isARGs10 = false;
                                        }
                                    }
                                    #endregion
                                }
                                if (tmp > sTimeLog) { sTimeLog = tmp; }
                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET SimulationDate='{tmp.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where SimulationId='{dr["SimulationId"].ToString()}'");
                            }

                            #region 回寫buffer紀錄
                            dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartSN=0");
                            if (dr_tmp != null)
                            {
                                int buf = TimeCompute2Seconds_BY_Calendar(db, dr_APS_NeedData["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp["SimulationDate"]), need_dDayNoBufferTime);
                                if (buf > 0)
                                {
                                    DataRow dr_01 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needId}' and StationNO='{dr_tmp["Source_StationNO"].ToString()}' and SimulationId='{dr_tmp["SimulationId"].ToString()}' order by CalendarDate desc");
                                    if (dr_01 != null)
                                    { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET BufferTime={buf.ToString()} where Id='{dr_01["Id"].ToString()}'"); }
                                }
                            }
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            errINFO = "模擬取得日期失敗,程式異常.";
                            isError = true;
                            System.Threading.Tasks.Task task = _Log.ErrorAsync($"排程模擬失敗,原因:模擬得到的日期. NeedId={needId} Exception:{ex.Message} {ex.StackTrace}", true);
                        }
                        #endregion
                    }
                }

            }
            #endregion

            if (!isError)
            {
                #region 修正APS_PartNOTimeNote非加工品的原物料量(依實際加工件數量的BOM)
                {
                    DataTable dt_PartNOTimeNote = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId}' and ((Class='4' or Class='5') and NoStation='0') order by CalendarDate");
                    if (dt_PartNOTimeNote != null && dt_PartNOTimeNote.Rows.Count > 0)
                    {
                        DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where Id='{needId}'");
                        string apply_PartNO = dr_tmp["PartNO"].ToString();
                        int tmp_i = 0;
                        foreach (DataRow dr in dt_PartNOTimeNote.Rows)
                        {
                            DataRow dt_BOM = db.DB_GetFirstDataByDataRow($@"SELECT a.*  FROM SoftNetMainDB.[dbo].[BOM] as a
                                                                    join SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.PartNO=b.PartNO and a.Apply_PP_Name=b.Apply_PP_Name and a.Apply_PartNO='{apply_PartNO}' and b.Source_StationNO=a.Apply_StationNO and b.Source_StationNO_IndexSN=a.IndexSN and b.Source_StationNO_Custom_IndexSN=a.Station_Custom_IndexSN
                                                                    where b.SimulationId='{dr["SimulationId"].ToString()}'");
                            if (dt_BOM != null)
                            {
                                DataTable dt_BOMII = db.DB_GetData($"select a.Id,a.PartNO,a.BOMQTY,a.Class,b.PartType,b.StoreSTime,b.SafeQTY from SoftNetMainDB.[dbo].[BOMII] as a,dbo.[Material] as b where b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO and BOMId='{dt_BOM["Id"].ToString()}' and a.Class!='4' and a.Class!='5' order by sn");
                                if (dt_BOMII != null && dt_BOMII.Rows.Count > 0)
                                {
                                    foreach (DataRow dr0 in dt_BOMII.Rows)
                                    {
                                        tmp_i = int.Parse(dr["NeedQTY"].ToString()) * int.Parse(dr0["BOMQTY"].ToString());
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and PartNO='{dr0["PartNO"].ToString()}' and Apply_PP_Name='{dt_BOM["Apply_PP_Name"].ToString()}' and Apply_StationNO='{dt_BOM["Apply_StationNO"].ToString()}' and IndexSN={dt_BOM["IndexSN"].ToString()} and Station_Custom_IndexSN='{dt_BOM["Station_Custom_IndexSN"].ToString()}'");
                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET NeedQTY={tmp_i} where SimulationId='{dr_tmp["SimulationId"].ToString()}'");
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                #region 修正 TotalStockII 的需求日ArrivalDate
                {
                    dt_Simulation = db.DB_GetData(@$"select a.SimulationId,a.StartDate,b.ArrivalDate,b.Id,b.NeedId from SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                                join SoftNetMainDB.[dbo].[TotalStockII] as b on a.SimulationId=b.SimulationId
                                                where a.NeedId='{needId}'");
                    if (dt_Simulation != null && dt_Simulation.Rows.Count > 0)
                    {
                        foreach (DataRow d in dt_Simulation.Rows)
                        {
                            if (!d.IsNull("StartDate") && d["StartDate"].ToString() != "")
                            {
                                db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] SET ArrivalDate='{Convert.ToDateTime(d["StartDate"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}' where Id='{d["Id"].ToString()}' and NeedId='{d["NeedId"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                            }
                        }
                    }
                }
                #endregion

                #region 計算多餘的量,回寫 APS_SimulationII
                #endregion

                if (wType == '5' || wType == '9')
                {
                    bool run = Create_WorkOrder(db, needId, dr_APS_NeedData["CalendarName"].ToString(), ref ERR);
                    if (ERR != "") { errINFO = ERR; }
                    if (run)
                    {
                        if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='6',StateINFO=NULL,UpdateTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}',NeedSimulationDate='{sTimeLog.ToString("yyyy-MM-dd HH:mm:ss")}' where ServerId='{_Fun.Config.ServerId}' and Id='{needId}'")) { run = true; }
                    }
                    else { isError = true; }
                }
                else
                {
                    db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='2',StateINFO=NULL,UpdateTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}',NeedSimulationDate='{sTimeLog.ToString("yyyy-MM-dd HH:mm:ss")}' where ServerId='{_Fun.Config.ServerId}' and Id='{needId}'");
                }
            }

            if (isError)
            {
                if (wType == '5')
                {
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{needId}'");
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"工作底稿發出自動生產失敗, 原因:料號:{dr_APS_NeedData["PartNO"].ToString()} {errINFO}", true);
                }
                else
                { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='5',StateINFO='{errINFO}',UpdateTime=null where Id='{needId}'"); }
                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{needId}'");
                db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] where Trigger_NeedId='{needId}'");
            }
            //}
            //catch(Exception ex)
            //{
            //    string _s = "";
            //    isError = true;
            //    //System.Threading.Tasks.Task task = _Log.ErrorAsync($"排程模擬失敗,原因:不明. NeedId={needId} Exception:{ex.Message} {ex.StackTrace}", true);
            //}
            return sTimeLog;

        }


        public bool RefreshRunSetSimulation(DBADO db, string[] data, ref List<string> run_HasWeb_Id_Change)
        {
            bool is_RefreshRunSetSimulation_OK = false;
            string sql = "";
            DataTable dt_tmp = null;
            DataRow dr_tmp = null;
            {
                DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select *,(NeedQTY+SafeQTY) as APS_SQTY from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{data[1]}' and IsOK='0'");
                DataRow dr_APS_NeedData = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{dr_APS_Simulation["NeedId"].ToString()}'");

                char type = '1';
                #region 先判斷 type = 1.有生產數據  2.未開工 與  紀錄觸發的 APS_Simulation 資料
                #region 變數
                int sQTY = 0;
                int pQTY = 0;
                DateTime startDate = DateTime.Now;
                DateTime simulationDate = DateTime.Now;
                string needID = "";
                string station = "";
                string indexSN = "";
                string partNO = "";
                string ppName = "";
                string simulationId = "";
                string orderNO = "";
                string partSN = "";

                List<string> log_Before_DATA = new List<string>();
                List<string> log_Change_DATA = new List<string>();
                int log_Dltime = 0;
                int log_Before_BufferTime = 0;
                string log_Change_NeedID_List = "";
                string log_Trigger_NeedId = "";
                string log_Trigger_SimulationId = "";
                string log_Trigger_Source_StationNO = "";
                DateTime log_Old_StartDate = DateTime.Now;
                DateTime log_Old_SimulationDate = DateTime.Now;
                DateTime log_New_SimulationDate = DateTime.Now;
                #endregion
                if (dr_APS_Simulation != null)
                {
                    #region 紀錄觸發的 APS_Simulation 資料
                    startDate = Convert.ToDateTime(dr_APS_Simulation["StartDate"]);
                    simulationDate = Convert.ToDateTime(dr_APS_Simulation["SimulationDate"]);
                    needID = dr_APS_Simulation["NeedId"].ToString();
                    station = dr_APS_Simulation["Source_StationNO"].ToString();
                    indexSN = dr_APS_Simulation["Source_StationNO_IndexSN"].ToString();
                    partNO = dr_APS_Simulation["PartNO"].ToString();
                    ppName = dr_APS_Simulation["Apply_PP_Name"].ToString();
                    simulationId = dr_APS_Simulation["SimulationId"].ToString();
                    orderNO = dr_APS_Simulation["DOCNumberNO"].ToString();
                    partSN = dr_APS_Simulation["PartSN"].ToString();

                    log_Trigger_NeedId = needID;
                    log_Trigger_SimulationId = simulationId;
                    log_Trigger_Source_StationNO = station;
                    log_Old_StartDate = startDate;
                    log_Old_SimulationDate = Convert.ToDateTime(dr_APS_NeedData["NeedSimulationDate"]);
                    #endregion

                    dr_tmp = db.DB_GetFirstDataByDataRow($"select *,(Detail_QTY+Detail_Fail_QTY) as PQTY,NeedQTY as SQTY from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and SimulationId='{simulationId}' and PartNO='{partNO}'");
                    if (dr_tmp != null)
                    {
                        sQTY = int.Parse(dr_tmp["SQTY"].ToString());
                        if (int.Parse(dr_tmp["PQTY"].ToString()) == 0)
                        {
                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{station}' and IndexSN={indexSN} and OrderNO='{orderNO}'  and PartNO='{partNO}' and LOGDateTime>='{startDate.ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTime<'{simulationDate.ToString("yyyy/MM/dd HH:mm:ss")}' and OperateType like '%開工%' order by LOGDateTime desc");
                            if (dr_tmp == null)
                            {
                                type = '2';
                            }
                        }
                        else
                        {
                            pQTY = int.Parse(dr_tmp["PQTY"].ToString());
                            if (pQTY >= sQTY) { return true; }
                        }
                    }
                    else { return true; }
                }
                else { return true; }
                #endregion

                int dltime = 0;
                DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_APS_Simulation["Source_StationNO"]}'");
                #region 計算要順延多小時間
                if (type == '1') { sQTY -= pQTY; }
                dltime = (int.Parse(dr_APS_Simulation["Math_UseTime"].ToString()) / sQTY) * int.Parse(dr_APS_Simulation["Math_UseOPCount"].ToString());
                if (!dr_APS_Simulation.IsNull("StationNO_Merge") && dr_APS_Simulation["StationNO_Merge"].ToString().Trim() != "")
                {
                    int count = dr_APS_Simulation["StationNO_Merge"].ToString().Split(',').Length - 1;
                    if (count > 0) { dltime = dltime / count; }

                }
                int isARGs10_offset = 15;//###??? 15將來改參數
                dltime += (isARGs10_offset * 60);
                DateTime dltime_startDate = TimeCompute2DateTime_BY_ReturnNextShift(db, dr_PP_Station["CalendarName"].ToString(), simulationDate, dltime);
                if (DateTime.Now > dltime_startDate)
                {
                    dltime_startDate = DateTime.Now;
                    if (sQTY > 0)
                    {
                        dltime = (int.Parse(dr_APS_Simulation["Math_UseTime"].ToString()) / sQTY) * int.Parse(dr_APS_Simulation["Math_UseOPCount"].ToString());
                        if (!dr_APS_Simulation.IsNull("StationNO_Merge") && dr_APS_Simulation["StationNO_Merge"].ToString().Trim() != "")
                        {
                            int count = dr_APS_Simulation["StationNO_Merge"].ToString().Split(',').Length - 1;
                            if (count > 0) { dltime = dltime / count; }

                        }
                    }
                    else { dltime = (isARGs10_offset * 60); }
                }
                else
                {
                    int dltime_TMP = TimeCompute2Seconds(simulationDate, dltime_startDate);
                    if (dltime_TMP > dltime) { dltime = dltime_TMP; }
                }
                log_Dltime = dltime;
                #endregion

                #region 先計算減自己buffer是否足夠
                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and PartSN=0");
                string tmp_station = $"StationNO='{dr_tmp["Source_StationNO"].ToString()}'";
                if (!dr_tmp.IsNull("StationNO_Merge") && dr_tmp["StationNO_Merge"].ToString().Trim() != "")
                {
                    string tmp_station02 = dr_tmp["StationNO_Merge"].ToString().Trim().Substring(0, dr_tmp["StationNO_Merge"].ToString().Trim().Length - 1);
                    tmp_station = $"({tmp_station} or StationNO in ('{tmp_station02.Replace(",", "','")}'))";
                }
                DataRow dr_01 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needID}' and {tmp_station} and SimulationId='{dr_tmp["SimulationId"].ToString()}' and BufferTime!=0 order by BufferTime,CalendarDate desc");
                if (dr_01 != null)
                {
                    log_Before_BufferTime = int.Parse(dr_01["BufferTime"].ToString());
                    if (log_Before_BufferTime >= dltime)
                    {
                        #region 要延滯的工站+後面工站 與原物料
                        dt_tmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and IsOK='0' and PartSN<={partSN} order by PartSN desc");
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            bool isRunOK = true;

                            //###???將來要做Log 若 Deferred_Calculation失敗, 要補救回來
                            string delStation = "";
                            List<string> noSID = new List<string>();
                            foreach (DataRow dr in dt_tmp.Rows)
                            {
                                #region 檢查已開站要排除
                                if (!dr.IsNull("Source_StationNO"))
                                {
                                    dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                                    if (dr_PP_Station["Station_Type"].ToString() == "1")
                                    {
                                        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}' and State='1'") > 0) { noSID.Add(dr["SimulationId"].ToString()); continue; }
                                    }
                                    else
                                    {
                                        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}' and OrderNO='{dr["DOCNumberNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeE is NULL") > 0) { noSID.Add(dr["SimulationId"].ToString()); continue; }
                                    }
                                }
                                #endregion

                                if (!dr.IsNull("Source_StationNO"))
                                {
                                    if (delStation == "") { delStation = $" and StationNO in ('{dr["Source_StationNO"].ToString()}'"; }
                                    else
                                    { delStation = $"{delStation},'{dr["Source_StationNO"].ToString()}'"; }
                                }
                            }
                            if (delStation != "") { delStation = $"{delStation})"; }
                            db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needID}' {delStation}");
                            string update_beforeEndTime_SID = "";
                            foreach (DataRow dr in dt_tmp.Rows)
                            {
                                if (noSID.Contains(dr["SimulationId"].ToString())) { update_beforeEndTime_SID = $"{update_beforeEndTime_SID},{dr["SimulationId"].ToString()}"; continue; }
                                if ((dr["Class"].ToString() == "4" || dr["Class"].ToString() == "5") && int.Parse(dr["PartSN"].ToString()) >= 0 && !dr.IsNull("Source_StationNO"))
                                { log_Before_DATA.Add($"StationNO:{dr["Source_StationNO"].ToString()},SimulationId:{dr["SimulationId"].ToString()},StartDate:{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")},SimulationDate:{dltime_startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}"); }
                                else { log_Before_DATA.Add($"PartNO:{dr["PartNO"]},SimulationId:{dr["SimulationId"].ToString()},StartDate:{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")},SimulationDate:{dltime_startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}"); }

                                isRunOK = Deferred_Calculation(db, dr, needID, dltime, ref dltime_startDate, ref log_Change_DATA);
                                if (!isRunOK)
                                {
                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"buffer足夠 RefreshRunSetSimulation延後失敗 NeedId={needID} SimulationId:{dr["SimulationId"].ToString()} PartSN={dr["PartSN"].ToString()} PartNO:{dr["PartNO"]}", true);
                                    goto break_FUN;
                                }
                                else
                                {
                                    #region 若上一站沒處理,則延沒處理SID的SimulationDate
                                    if (update_beforeEndTime_SID != "")
                                    {
                                        string[] b_SID = update_beforeEndTime_SID.Split(',');
                                        foreach (string s in b_SID)
                                        {
                                            if (s != "")
                                            {
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                                                db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set SimulationDate='{Convert.ToDateTime(dr_tmp["StartDate"]).AddMilliseconds(-_Fun.Config.RunTimeServerLoopTime).ToString("yyyy-MM-dd HH:mm:ss.fff")}' where NeedId='{needID}' and SimulationId='{s}'");
                                            }
                                        }
                                    }
                                    update_beforeEndTime_SID = "";
                                    #endregion

                                    #region 判斷多工站是否刪 ManufactureII
                                    if (!dr.IsNull("Source_StationNO"))
                                    {
                                        dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                                        if (dr_PP_Station != null && dr_PP_Station["Station_Type"].ToString() == "8")
                                        {
                                            if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[ManufactureII] where StationNO='{dr["Source_StationNO"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}' and RemarkTimeE is NULL") > 0)
                                            {
                                                db.DB_SetData($"delete from SoftNetMainDB.[dbo].[ManufactureII] where StationNO='{dr["Source_StationNO"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}' and RemarkTimeE is NULL");
                                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='' where SimulationId='{dr["SimulationId"].ToString()}'");
                                                if (!run_HasWeb_Id_Change.Contains(dr["Source_StationNO"].ToString())) { run_HasWeb_Id_Change.Add(dr["Source_StationNO"].ToString()); }
                                            }
                                        }
                                    }
                                    #endregion
                                }
                            }
                            if (isRunOK)
                            {
                                #region 消除之後的19,03,15,18,08警示
                                dt_tmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_Simulation] as a where NeedId='{needID}' and IsOK='0' and PartSN<={partSN} and (Class='4' or Class='5') and Source_StationNO is not null and PartSN>=0 order by PartSN desc");
                                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                {
                                    string sID = "";
                                    foreach (DataRow dr in dt_tmp.Rows)
                                    {
                                        if (sID == "") { sID = $"'{dr["SimulationId"].ToString()}'"; }
                                        else { sID = $"{sID},'{dr["SimulationId"].ToString()}'"; }
                                    }
                                    if (sID != "")
                                    {
                                        if (db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId in ({sID}) and (ErrorType='19' or ErrorType='03' or ErrorType='15' or ErrorType='18' or ErrorType='08')"))
                                        { db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_WarningData] where ServerId='{_Fun.Config.ServerId}' and SimulationId in ({sID}) and (ErrorType='19' or ErrorType='18') and IsDEL!='1'"); }
                                    }
                                }
                                #endregion
                                #region 回寫buffer紀錄
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and PartSN=0");
                                if (dr_tmp != null)
                                {
                                    log_New_SimulationDate = Convert.ToDateTime(dr_tmp["SimulationDate"]);
                                    int buf = TimeCompute2Seconds_BY_Calendar(db, dr_APS_NeedData["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp["SimulationDate"]), Convert.ToDateTime(dr_APS_NeedData["NeedDate"]));
                                    if (buf < 0) { buf = 0; }
                                    dr_01 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needID}' and StationNO='{dr_tmp["Source_StationNO"].ToString()}' and SimulationId='{dr_tmp["SimulationId"].ToString()}' order by CalendarDate desc");
                                    if (dr_01 != null)
                                    { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET BufferTime={buf.ToString()} where Id='{dr_01["Id"].ToString()}'"); }
                                }
                                #endregion
                            }
                        }
                        #endregion
                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] ([Id],[Trigger_NeedId],[Trigger_SimulationId],[Trigger_Source_StationNO],[Old_StartDate],[Old_SimulationDate],[Before_DATA],[Before_BufferTime],[Dltime],[Change_NeedID_List],[Change_DATA],[New_SimulationDate])
                                                             VALUES ('{_Str.NewId('Y')}','{log_Trigger_NeedId}','{log_Trigger_SimulationId}','{log_Trigger_Source_StationNO}','{log_Old_StartDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{log_Old_SimulationDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}',
                                                            '{string.Join(";", log_Before_DATA)}','{log_Before_BufferTime.ToString()}','{log_Dltime.ToString()}','{log_Change_NeedID_List}','{string.Join(";", log_Change_DATA)}','{log_New_SimulationDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                        is_RefreshRunSetSimulation_OK = true;
                        goto break_FUN;
                    }
                }
                #endregion

                if (dltime > 0)
                {
                    #region 查找期間其他工單
                    List<string> list = new List<string>();//區間內NeedID
                    List<string> stationNO_list = new List<string>();//區間內工站
                    DataRow dr_tmp01 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and PartSN=0");
                    DateTime workTimeNoteEndTime = Convert.ToDateTime(dr_tmp01["SimulationDate"]);

                    //###???區間的只是模擬的NeedID可能需要被標示被干涉過

                    dt_tmp = db.DB_GetData($@"select NeedId,StationNO from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where CalendarDate>='{simulationDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}' and CalendarDate<='{workTimeNoteEndTime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' 
                                                            and NeedId!='{needID}' and (Time1_C>0 or Time2_C>0 or Time3_C>0 or Time4_C>0) and DOCNumberNO!='' Group by NeedId,StationNO order by NeedId desc");
                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt_tmp.Rows)
                        {
                            if (!list.Contains(dr["NeedId"].ToString())) { list.Add(dr["NeedId"].ToString()); }
                            if (!stationNO_list.Contains(dr["StationNO"].ToString())) { stationNO_list.Add(dr["StationNO"].ToString()); }
                        }
                    }
                    #endregion

                    if (list.Count == 0)
                    {//直接延
                        #region 要延滯的工站+後面工站 與原物料
                        dt_tmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and IsOK='0' and PartSN<={partSN} order by PartSN desc");
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            bool isRunOK = true;

                            //###???將來要做Log 若 Deferred_Calculation失敗, 要補救回來



                            string delStation = "";
                            List<string> noSID = new List<string>();
                            foreach (DataRow dr in dt_tmp.Rows)
                            {
                                #region 檢查已開站要排除
                                if (!dr.IsNull("Source_StationNO"))
                                {
                                    dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                                    if (dr_PP_Station["Station_Type"].ToString() == "1")
                                    {
                                        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}' and State='1'") > 0) { noSID.Add(dr["SimulationId"].ToString()); continue; }
                                    }
                                    else
                                    {
                                        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}' and OrderNO='{dr["DOCNumberNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeE is NULL") > 0) { noSID.Add(dr["SimulationId"].ToString()); continue; }
                                    }
                                }
                                #endregion

                                if (!dr.IsNull("Source_StationNO"))
                                {
                                    if (delStation == "") { delStation = $" and StationNO in ('{dr["Source_StationNO"].ToString()}'"; }
                                    else
                                    { delStation = $"{delStation},'{dr["Source_StationNO"].ToString()}'"; }
                                }
                            }
                            if (delStation != "") { delStation = $"{delStation})"; }
                            db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needID}' {delStation}");
                            string update_beforeEndTime_SID = "";
                            foreach (DataRow dr in dt_tmp.Rows)
                            {
                                if (noSID.Contains(dr["SimulationId"].ToString())) { update_beforeEndTime_SID = $"{update_beforeEndTime_SID},{dr["SimulationId"].ToString()}"; continue; }
                                if ((dr["Class"].ToString() == "4" || dr["Class"].ToString() == "5") && int.Parse(dr["PartSN"].ToString()) >= 0 && !dr.IsNull("Source_StationNO"))
                                { log_Before_DATA.Add($"StationNO:{dr["Source_StationNO"].ToString()},SimulationId:{dr["SimulationId"].ToString()},StartDate:{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")},SimulationDate:{dltime_startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}"); }
                                else { log_Before_DATA.Add($"PartNO:{dr["PartNO"]},SimulationId:{dr["SimulationId"].ToString()},StartDate:{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")},SimulationDate:{dltime_startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}"); }
                                isRunOK = Deferred_Calculation(db, dr, needID, dltime, ref dltime_startDate, ref log_Change_DATA);
                                if (!isRunOK)
                                {
                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"直接延 RefreshRunSetSimulation延後失敗 NeedId={needID} SimulationId:{dr["SimulationId"].ToString()} PartSN={dr["PartSN"].ToString()} PartNO:{dr["PartNO"]}", true);
                                    goto break_FUN;
                                }
                                else
                                {
                                    #region 若上一站沒處理,則延沒處理SID的SimulationDate
                                    if (update_beforeEndTime_SID != "")
                                    {
                                        string[] b_SID = update_beforeEndTime_SID.Split(',');
                                        foreach (string s in b_SID)
                                        {
                                            if (s != "")
                                            {
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                                                db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set SimulationDate='{Convert.ToDateTime(dr_tmp["StartDate"]).AddMilliseconds(-_Fun.Config.RunTimeServerLoopTime).ToString("yyyy-MM-dd HH:mm:ss.fff")}' where NeedId='{needID}' and SimulationId='{s}'");
                                            }
                                        }
                                    }
                                    update_beforeEndTime_SID = "";
                                    #endregion

                                    #region 判斷多工站是否刪 ManufactureII
                                    if (!dr.IsNull("Source_StationNO"))
                                    {
                                        dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                                        if (dr_PP_Station != null && dr_PP_Station["Station_Type"].ToString() == "8")
                                        {
                                            if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[ManufactureII] where StationNO='{dr["Source_StationNO"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}' and RemarkTimeE is NULL") > 0)
                                            {
                                                db.DB_SetData($"delete from SoftNetMainDB.[dbo].[ManufactureII] where StationNO='{dr["Source_StationNO"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}' and RemarkTimeE is NULL");
                                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='' where SimulationId='{dr["SimulationId"].ToString()}'");
                                                if (!run_HasWeb_Id_Change.Contains(dr["Source_StationNO"].ToString())) { run_HasWeb_Id_Change.Add(dr["Source_StationNO"].ToString()); }
                                            }
                                        }
                                    }
                                    #endregion
                                }
                            }
                            if (isRunOK)
                            {
                                #region 消除之後的19,03,15,18,08警示
                                dt_tmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_Simulation] as a where NeedId='{needID}' and IsOK='0' and PartSN<={partSN} and (Class='4' or Class='5') and Source_StationNO is not null and PartSN>=0 order by PartSN desc");
                                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                {
                                    string sID = "";
                                    foreach (DataRow dr in dt_tmp.Rows)
                                    {
                                        if (sID == "") { sID = $"'{dr["SimulationId"].ToString()}'"; }
                                        else { sID = $"{sID},'{dr["SimulationId"].ToString()}'"; }
                                    }
                                    if (sID != "")
                                    {
                                        if (db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId in ({sID}) and (ErrorType='19' or ErrorType='03' or ErrorType='15' or ErrorType='18' or ErrorType='08')"))
                                        { db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_WarningData] where ServerId='{_Fun.Config.ServerId}' and SimulationId in ({sID}) and (ErrorType='19' or ErrorType='18') and IsDEL!='1'"); }
                                    }
                                }
                                #endregion
                                #region 回寫buffer紀錄
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and PartSN=0");
                                if (dr_tmp != null)
                                {
                                    log_New_SimulationDate = Convert.ToDateTime(dr_tmp["SimulationDate"]);
                                    int buf = TimeCompute2Seconds_BY_Calendar(db, dr_APS_NeedData["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp["SimulationDate"]), Convert.ToDateTime(dr_APS_NeedData["NeedDate"]));
                                    if (buf < 0) { buf = 0; }
                                    dr_01 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needID}' and StationNO='{dr_tmp["Source_StationNO"].ToString()}' and SimulationId='{dr_tmp["SimulationId"].ToString()}' order by CalendarDate desc");
                                    if (dr_01 != null)
                                    { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET BufferTime={buf.ToString()} where Id='{dr_01["Id"].ToString()}'"); }
                                }
                                db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] ([Id],[Trigger_NeedId],[Trigger_SimulationId],[Trigger_Source_StationNO],[Old_StartDate],[Old_SimulationDate],[Before_DATA],[Before_BufferTime],[Dltime],[Change_NeedID_List],[Change_DATA],[New_SimulationDate])
                                                             VALUES ('{_Str.NewId('Y')}','{log_Trigger_NeedId}','{log_Trigger_SimulationId}','{log_Trigger_Source_StationNO}','{log_Old_StartDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{log_Old_SimulationDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}',
                                                            '{string.Join(";", log_Before_DATA)}','{log_Before_BufferTime.ToString()}','{log_Dltime.ToString()}','{log_Change_NeedID_List}','{string.Join(";", log_Change_DATA)}','{log_New_SimulationDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");

                                #endregion
                                is_RefreshRunSetSimulation_OK = true;
                                goto break_FUN;
                            }
                        }
                        #endregion
                    }
                    else
                    {//從有bufferTime先延

                        Dictionary<string, string> needID_list = new Dictionary<string, string>();
                        #region needID 依bufferTime多排序 and 查找 partSN
                        dt_tmp = db.DB_GetData($@"select a.Id from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] as a where a.NeedId in ('{string.Join("','", list)}') and CalendarDate>='{simulationDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}' and CalendarDate<='{workTimeNoteEndTime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' group by a.Id");
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            string iddata = "";
                            foreach (DataRow dr in dt_tmp.Rows)
                            {
                                if (iddata == "") { iddata = $"'{dr["Id"].ToString()}'"; }
                                else { iddata = $"{iddata},'{dr["Id"].ToString()}'"; }
                            }
                            dt_tmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where Id in ({iddata}) order by BufferTime desc");
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                foreach (DataRow dr in dt_tmp.Rows)
                                {
                                    if (!needID_list.ContainsKey(dr["NeedId"].ToString()))
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{dr["NeedId"].ToString()}'");
                                        if (dr_tmp["KeyA"].ToString().Trim().Split(",")[4] == "1") { continue; }
                                        dr_tmp = db.DB_GetFirstDataByDataRow($@"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr["NeedId"].ToString()}' and IsOK='0' and Source_StationNO in ('{string.Join("','", stationNO_list)}') and StartDate>='{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by PartSN desc");
                                        if (dr_tmp != null) { needID_list.Add(dr["NeedId"].ToString(), dr_tmp["PartSN"].ToString()); }
                                        else { needID_list.Add(dr["NeedId"].ToString(), ""); }
                                    }
                                }
                            }
                            foreach (string s in list)
                            { if (!needID_list.ContainsKey(s)) { needID_list.Add(s, ""); } }
                        }
                        else
                        {
                            foreach (string s in list)
                            { if (!needID_list.ContainsKey(s)) { needID_list.Add(s, ""); } }
                        }
                        #endregion

                        #region 開始依needID_list順序依依延期,直到符合可以控制的日期
                        for (int i_list = 0; i_list < needID_list.Count; ++i_list)
                        {
                            KeyValuePair<string, string> kv = needID_list.ElementAt(i_list);
                            if (kv.Value == "") { continue; }
                            #region 檢查是否已有生產行為,則記錄下一階PartSN該預備NeedId的延期
                            bool isBreak = false;
                            DataRow dr_APS_NeedDataII = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{kv.Key}'");

                            dt_tmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{kv.Key}' and IsOK='0' and PartSN<={kv.Value} and Source_StationNO in ('{string.Join("','", stationNO_list)}') order by PartSN desc");
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                foreach (DataRow dr in dt_tmp.Rows)
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT TOP 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}' and IndexSN={dr["Source_StationNO_IndexSN"].ToString()} and OrderNO='{dr["DOCNumberNO"].ToString()}' and PartNO='{dr["PartNO"].ToString()}' and LOGDateTime>='{Convert.ToDateTime(dr["StartDate"]).ToString("MM/dd/yyyy HH:mm:ss.fff")}' and LOGDateTime<'{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}' and OperateType like '%開工%'");
                                    if (dr_tmp != null)
                                    {
                                        int i_tmp = int.Parse(dr["PartSN"].ToString()) - 1;
                                        if (i_tmp > 0)
                                        {
                                            needID_list[kv.Key] = i_tmp.ToString();
                                        }
                                        else { needID_list[kv.Key] = ""; continue; }
                                    }
                                }
                                if (log_Change_NeedID_List == "") { log_Change_NeedID_List = $"{kv.Key}:{kv.Value}"; }
                                else { log_Change_NeedID_List = $"{log_Change_NeedID_List},{kv.Key}:{kv.Value}"; }
                                #region 要延滯預備NeedId的工站+後面工站
                                isBreak = true;
                                dt_tmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{kv.Key}' and IsOK='0' and PartSN<={kv.Value} order by PartSN desc");
                                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                {
                                    //###???將來要做Log 若 Deferred_Calculation失敗, 要補救回來
                                    string delStation = "";
                                    List<string> noSID = new List<string>();
                                    foreach (DataRow dr in dt_tmp.Rows)
                                    {
                                        #region 檢查已開站要排除
                                        if (!dr.IsNull("Source_StationNO"))
                                        {
                                            dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                                            if (dr_PP_Station["Station_Type"].ToString() == "1")
                                            {
                                                if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}' and State='1'") > 0) { noSID.Add(dr["SimulationId"].ToString()); continue; }
                                            }
                                            else
                                            {
                                                if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}' and OrderNO='{dr["DOCNumberNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeE is NULL") > 0) { noSID.Add(dr["SimulationId"].ToString()); continue; }
                                            }
                                        }
                                        #endregion

                                        if (!dr.IsNull("Source_StationNO"))
                                        {
                                            if (delStation == "") { delStation = $" and StationNO in ('{dr["Source_StationNO"].ToString()}'"; }
                                            else
                                            { delStation = $"{delStation},'{dr["Source_StationNO"].ToString()}'"; }
                                        }
                                    }
                                    if (delStation != "") { delStation = $"{delStation})"; }
                                    db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needID}' {delStation}");
                                    string update_beforeEndTime_SID = "";
                                    foreach (DataRow dr in dt_tmp.Rows)
                                    {
                                        if (noSID.Contains(dr["SimulationId"].ToString())) { update_beforeEndTime_SID = $"{update_beforeEndTime_SID},{dr["SimulationId"].ToString()}"; continue; }
                                        if ((dr["Class"].ToString() == "4" || dr["Class"].ToString() == "5") && int.Parse(dr["PartSN"].ToString()) >= 0 && !dr.IsNull("Source_StationNO"))
                                        { log_Before_DATA.Add($"StationNO:{dr["Source_StationNO"].ToString()},SimulationId:{dr["SimulationId"].ToString()},StartDate:{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")},SimulationDate:{dltime_startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}"); }
                                        else { log_Before_DATA.Add($"PartNO:{dr["PartNO"]},SimulationId:{dr["SimulationId"].ToString()},StartDate:{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")},SimulationDate:{dltime_startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}"); }
                                        isBreak = Deferred_Calculation(db, dr, kv.Key, dltime, ref dltime_startDate, ref log_Change_DATA);
                                        if (!isBreak)
                                        {
                                            System.Threading.Tasks.Task task = _Log.ErrorAsync($"多NeedId RefreshRunSetSimulation延後失敗 NeedId={needID} SimulationId:{dr["SimulationId"].ToString()} PartSN={dr["PartSN"].ToString()} PartNO:{dr["PartNO"]}", true);
                                            goto break_FUN;
                                        }
                                        else
                                        {
                                            #region 若上一站沒處理,則延沒處理SID的SimulationDate
                                            if (update_beforeEndTime_SID != "")
                                            {
                                                string[] b_SID = update_beforeEndTime_SID.Split(',');
                                                foreach (string s in b_SID)
                                                {
                                                    if (s != "")
                                                    {
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                                                        db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set SimulationDate='{Convert.ToDateTime(dr_tmp["StartDate"]).AddMilliseconds(-_Fun.Config.RunTimeServerLoopTime).ToString("yyyy-MM-dd HH:mm:ss.fff")}' where NeedId='{needID}' and SimulationId='{s}'");
                                                    }
                                                }
                                            }
                                            update_beforeEndTime_SID = "";
                                            #endregion

                                            #region 判斷多工站是否刪 ManufactureII
                                            if (!dr.IsNull("Source_StationNO"))
                                            {
                                                dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                                                if (dr_PP_Station != null && dr_PP_Station["Station_Type"].ToString() == "8")
                                                {
                                                    if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[ManufactureII] where StationNO='{dr["Source_StationNO"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}' and RemarkTimeE is NULL") > 0)
                                                    {
                                                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[ManufactureII] where StationNO='{dr["Source_StationNO"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}' and RemarkTimeE is NULL");
                                                        db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='' where SimulationId='{dr["SimulationId"].ToString()}'");
                                                        if (!run_HasWeb_Id_Change.Contains(dr["Source_StationNO"].ToString())) { run_HasWeb_Id_Change.Add(dr["Source_StationNO"].ToString()); }
                                                    }
                                                }
                                            }
                                            #endregion
                                        }
                                    }
                                    if (isBreak)
                                    {
                                        #region 消除之後的19,03,15,18,08警示
                                        dt_tmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_Simulation] as a where NeedId='{kv.Key}' and IsOK='0' and PartSN<={kv.Value} and (Class='4' or Class='4') and PartSN>=0 order by PartSN desc");
                                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                        {
                                            string sID = "";
                                            foreach (DataRow dr in dt_tmp.Rows)
                                            {
                                                if (sID == "") { sID = $"'{dr["SimulationId"].ToString()}'"; }
                                                else { sID = $"{sID},'{dr["SimulationId"].ToString()}'"; }
                                            }
                                            if (sID != "")
                                            {
                                                if (db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId in ({sID}) and (ErrorType='19' or ErrorType='03' or ErrorType='15' or ErrorType='18' or ErrorType='08')"))
                                                { db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_WarningData] where ServerId='{_Fun.Config.ServerId}' and SimulationId in ({sID}) and (ErrorType='19' or ErrorType='18') and IsDEL!='1'"); }
                                            }
                                        }
                                        #endregion
                                        #region 回寫buffer紀錄
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{kv.Key}' and PartSN=0");
                                        if (dr_tmp != null)
                                        {
                                            int buf = TimeCompute2Seconds_BY_Calendar(db, dr_APS_NeedDataII["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp["SimulationDate"]), Convert.ToDateTime(dr_APS_NeedDataII["NeedDate"]));
                                            if (buf < 0) { buf = 0; }
                                            dr_01 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{kv.Key}' and StationNO='{dr_tmp["Source_StationNO"].ToString()}' and SimulationId='{dr_tmp["SimulationId"].ToString()}' order by CalendarDate desc");
                                            if (dr_01 != null)
                                            { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET BufferTime={buf.ToString()} where Id='{dr_01["Id"].ToString()}'"); }
                                        }
                                        #endregion

                                    }
                                    else
                                    {
                                        //###??? 發生錯誤 不應該進入
                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"RefreshRunSetSimulation延後失敗 發生錯誤 不應該進入 NeedId='{kv.Key} PartSN<={kv.Value}", true);

                                        string _s01 = "";
                                    }
                                }
                                #endregion

                                #region 判斷延期是否OK, 要離開foreach
                                if (isBreak)
                                {
                                    DateTime uu = TMP_Simulation_APS_WorkTimeNote(db, needID, partSN, dltime, simulationId);
                                    if (Convert.ToDateTime(dr_APS_NeedData["NeedDate"]) >= uu || i_list >= (needID_list.Count - 1))
                                    {
                                        bool isRunOK = true;

                                        #region 此次要延滯的工站+後面工站 與原物料
                                        dt_tmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and IsOK='0' and PartSN<={partSN} order by PartSN desc");
                                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                        {

                                            //###???將來要做Log 若 Deferred_Calculation失敗, 要補救回來
                                            string delStation = "";
                                            List<string> noSID = new List<string>();
                                            foreach (DataRow dr in dt_tmp.Rows)
                                            {
                                                #region 檢查已開站要排除
                                                if (!dr.IsNull("Source_StationNO"))
                                                {
                                                    dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                                                    if (dr_PP_Station["Station_Type"].ToString() == "1")
                                                    {
                                                        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}' and State='1'") > 0) { noSID.Add(dr["SimulationId"].ToString()); continue; }
                                                    }
                                                    else
                                                    {
                                                        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}' and OrderNO='{dr["DOCNumberNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeE is NULL") > 0) { noSID.Add(dr["SimulationId"].ToString()); continue; }
                                                    }
                                                }
                                                #endregion

                                                if (!dr.IsNull("Source_StationNO"))
                                                {
                                                    if (delStation == "") { delStation = $" and StationNO in ('{dr["Source_StationNO"].ToString()}'"; }
                                                    else
                                                    { delStation = $"{delStation},'{dr["Source_StationNO"].ToString()}'"; }
                                                }
                                            }
                                            if (delStation != "") { delStation = $"{delStation})"; }
                                            db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needID}' {delStation}");
                                            string update_beforeEndTime_SID = "";
                                            foreach (DataRow dr in dt_tmp.Rows)
                                            {
                                                if (noSID.Contains(dr["SimulationId"].ToString())) { update_beforeEndTime_SID = $"{update_beforeEndTime_SID},{dr["SimulationId"].ToString()}"; continue; }
                                                if ((dr["Class"].ToString() == "4" || dr["Class"].ToString() == "5") && int.Parse(dr["PartSN"].ToString()) >= 0 && !dr.IsNull("Source_StationNO"))
                                                { log_Before_DATA.Add($"StationNO:{dr["Source_StationNO"].ToString()},SimulationId:{dr["SimulationId"].ToString()},StartDate:{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")},SimulationDate:{dltime_startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}"); }
                                                else { log_Before_DATA.Add($"PartNO:{dr["PartNO"]},SimulationId:{dr["SimulationId"].ToString()},StartDate:{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")},SimulationDate:{dltime_startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}"); }
                                                isRunOK = Deferred_Calculation(db, dr, needID, dltime, ref dltime_startDate, ref log_Change_DATA);
                                                if (!isRunOK)
                                                {
                                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"判斷自己 RefreshRunSetSimulation延後失敗 發生錯誤 不應該進入 NeedId='{needID} PartSN<={partSN}", true);
                                                    break;
                                                }
                                                else
                                                {
                                                    #region 若上一站沒處理,則延沒處理SID的SimulationDate
                                                    if (update_beforeEndTime_SID != "")
                                                    {
                                                        string[] b_SID = update_beforeEndTime_SID.Split(',');
                                                        foreach (string s in b_SID)
                                                        {
                                                            if (s != "")
                                                            {
                                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                                                                db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set SimulationDate='{Convert.ToDateTime(dr_tmp["StartDate"]).AddMilliseconds(-_Fun.Config.RunTimeServerLoopTime).ToString("yyyy-MM-dd HH:mm:ss.fff")}' where NeedId='{needID}' and SimulationId='{s}'");
                                                            }
                                                        }
                                                    }
                                                    update_beforeEndTime_SID = "";
                                                    #endregion

                                                    #region 判斷多工站是否刪 ManufactureII
                                                    if (!dr.IsNull("Source_StationNO"))
                                                    {
                                                        dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");
                                                        if (dr_PP_Station != null && dr_PP_Station["Station_Type"].ToString() == "8")
                                                        {
                                                            if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[ManufactureII] where StationNO='{dr["Source_StationNO"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}' and RemarkTimeE is NULL") > 0)
                                                            {
                                                                db.DB_SetData($"delete from SoftNetMainDB.[dbo].[ManufactureII] where StationNO='{dr["Source_StationNO"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}' and RemarkTimeE is NULL");
                                                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='' where SimulationId='{dr["SimulationId"].ToString()}'");
                                                                if (!run_HasWeb_Id_Change.Contains(dr["Source_StationNO"].ToString())) { run_HasWeb_Id_Change.Add(dr["Source_StationNO"].ToString()); }
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                            }
                                            if (isRunOK)
                                            {
                                                #region 消除之後的19,03,15,18,08警示
                                                dt_tmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_Simulation] as a where NeedId='{needID}' and IsOK='0' and PartSN<={partSN} and (Class='4' or Class='5') and Source_StationNO is not null and PartSN>=0 order by PartSN desc");
                                                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                                {
                                                    string sID = "";
                                                    foreach (DataRow dr in dt_tmp.Rows)
                                                    {
                                                        if (dr["PartSN"].ToString() == "0")
                                                        {
                                                            log_New_SimulationDate = Convert.ToDateTime(dr_tmp["SimulationDate"]);
                                                        }
                                                        if (sID == "") { sID = $"'{dr["SimulationId"].ToString()}'"; }
                                                        else { sID = $"{sID},'{dr["SimulationId"].ToString()}'"; }
                                                    }
                                                    if (sID != "")
                                                    {
                                                        if (db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId in ({sID}) and (ErrorType='19' or ErrorType='03' or ErrorType='15' or ErrorType='18' or ErrorType='08')"))
                                                        { db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_WarningData] where ServerId='{_Fun.Config.ServerId}' and SimulationId in ({sID}) and (ErrorType='19' or ErrorType='18') and IsDEL!='1'"); }
                                                    }
                                                }
                                                #endregion
                                                #region 回寫buffer紀錄
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and PartSN=0");
                                                if (dr_tmp != null)
                                                {
                                                    int buf = TimeCompute2Seconds_BY_Calendar(db, dr_APS_NeedData["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp["SimulationDate"]), Convert.ToDateTime(dr_APS_NeedData["NeedDate"]));
                                                    if (buf < 0) { buf = 0; }
                                                    dr_01 = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{needID}' and StationNO='{dr_tmp["Source_StationNO"].ToString()}' and SimulationId='{dr_tmp["SimulationId"].ToString()}' order by CalendarDate desc");
                                                    if (dr_01 != null) { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET BufferTime={buf.ToString()} where Id='{dr_01["Id"].ToString()}'"); }
                                                    db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET NeedSimulationDate='{Convert.ToDateTime(dr_tmp["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss")}' where Id='{needID}'");

                                                }
                                                #endregion
                                                db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] ([Id],[Trigger_NeedId],[Trigger_SimulationId],[Trigger_Source_StationNO],[Old_StartDate],[Old_SimulationDate],[Before_DATA],[Before_BufferTime],[Dltime],[Change_NeedID_List],[Change_DATA],[New_SimulationDate])
                                                             VALUES ('{_Str.NewId('Y')}','{log_Trigger_NeedId}','{log_Trigger_SimulationId}','{log_Trigger_Source_StationNO}','{log_Old_StartDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{log_Old_SimulationDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}',
                                                            '{string.Join(";", log_Before_DATA)}','{log_Before_BufferTime.ToString()}','{log_Dltime.ToString()}','{log_Change_NeedID_List}','{string.Join(";", log_Change_DATA)}','{log_New_SimulationDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");

                                            }
                                        }
                                        #endregion
                                        if (isRunOK)
                                        {
                                            is_RefreshRunSetSimulation_OK = true;
                                            goto break_FUN;
                                        }
                                    }
                                }
                                #endregion
                            }
                            #endregion
                        }
                        #endregion

                    }

                }

            }

        break_FUN:
            if (!is_RefreshRunSetSimulation_OK)
            {
                string _s = "";
            }
            return is_RefreshRunSetSimulation_OK;
        }
        private bool Deferred_Calculation(DBADO db, DataRow dr, string needID, int dltime, ref DateTime dltime_startDate, ref List<string> log_list)
        {
            bool IsErr = true;
            string errINFO = "";
            DateTime startDate = dltime_startDate;
            DateTime arrivalDate = dltime_startDate;
            DataRow dr_tmp = null;
            RunSimulation_Arg args = new RunSimulation_Arg();
            try
            {
                DataRow dr_APS_NeedData = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{needID}'");
                DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["Source_StationNO"].ToString()}'");

                #region 讀排程設定參數 args
                List<string> cs = dr_APS_NeedData["KeyA"].ToString().Split(',').ToList();
                foreach (string c in cs)
                {
                    if (c == "0") { args.ARGs.Add(false); }
                    else { args.ARGs.Add(true); }
                }
                #endregion

                if ((dr["Class"].ToString() == "4" || dr["Class"].ToString() == "5") && int.Parse(dr["PartSN"].ToString()) >= 0 && !dr.IsNull("Source_StationNO"))
                {
                    #region for GetFinallyDate 前置處裡 目的修改 APS_WorkTimeNote 時間

                    dr_tmp = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Master_PartNO='{dr["PartNO"].ToString()}' order by SimulationDate desc");
                    string partNOClass = dr["Class"].ToString();
                    if (dr["ChangeSimulationClass"].ToString().Trim() != "") { partNOClass = dr["ChangeSimulationClass"].ToString().Trim(); }
                    int real_NeedQTY = int.Parse(dr["NeedQTY"].ToString()) + int.Parse(dr["SafeQTY"].ToString()) - int.Parse(dr["Math_TotalStock_HasUseQTY"].ToString()) - int.Parse(dr["Math_Online_SurplusQTY"].ToString());
                    #endregion
                    object stationNO_Merge = null;
                    if (!dr.IsNull("StationNO_Merge")) { stationNO_Merge = dr["StationNO_Merge"]; }
                    dltime_startDate = GetFinallyDate(args, db, needID, dr["SimulationId"].ToString(), "站", dr["Source_StationNO"].ToString(), dr["PartNO"].ToString(), partNOClass, real_NeedQTY.ToString(), dr_PP_Station["CalendarName"].ToString(), dltime_startDate, int.Parse(dr["Math_UseTime"].ToString()), int.Parse(dr["Math_UseOPCount"].ToString()), stationNO_Merge, ref errINFO, false, false);
                    if (errINFO != "") { return false; }
                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkTimeNote] SET DOCNumberNO='{dr["DOCNumberNO"].ToString()}' where NeedId='{needID}' and SimulationId='{dr["SimulationId"].ToString()}' and StationNO='{dr["Source_StationNO"].ToString()}'");

                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] SET StartDate='{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}',SimulationDate='{dltime_startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where SimulationId='{dr["SimulationId"].ToString()}'");
                    log_list.Add($"StationNO:{dr["Source_StationNO"].ToString()},SimulationId:{dr["SimulationId"].ToString()},StartDate:{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")},SimulationDate:{dltime_startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}");
                }
                else
                {
                    //###領料起始日ㄝ要改
                    #region 改TotalStockII keep的 ArrivalDate時間 
                    DataTable dt_tmp03 = db.DB_GetData($@"select * from SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{dr["NeedId"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}' and KeepQTY!=0 order by ArrivalDate");
                    if (dt_tmp03 != null && dt_tmp03.Rows.Count > 0)
                    {
                        foreach (DataRow dr3 in dt_tmp03.Rows)
                        {
                            arrivalDate = TimeCompute2DateTime_BY_ReturnNextShift(db, dr_APS_NeedData["CalendarName"].ToString(), Convert.ToDateTime(dr3["ArrivalDate"]), dltime);
                            db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStockII] SET ArrivalDate='{arrivalDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where NeedId='{dr3["NeedId"].ToString()}' and Id='{dr3["Id"].ToString()}' and SimulationId='{dr3["SimulationId"].ToString()}'");
                        }
                    }
                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{dr["PartNO"].ToString()}' and NeedId='{dr["NeedId"].ToString()}' and SimulationId='{dr["SimulationId"].ToString()}'");
                    db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] SET StartDate='{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}',SimulationDate='{arrivalDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where SimulationId='{dr["SimulationId"].ToString()}'");
                    log_list.Add($"PartNO:{dr["PartNO"]},SimulationId:{dr["SimulationId"].ToString()},StartDate:{startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")},SimulationDate:{arrivalDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}");
                    #endregion

                }
            }
            catch (Exception ex)
            {
                IsErr = false;
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"SNWebSockeyService.cs Deferred_Calculation {ex.Message} {ex.StackTrace}", true);

            }
            args.Dispose();
            return IsErr;
        }

        private DateTime TMP_Simulation_APS_WorkTimeNote(DBADO db, string needID, string partSN, int dltime, string simulationId)
        {
            DateTime intime = DateTime.Now;
            DataTable dt_tmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and (Class='4' or Class='5') and Source_StationNO is not null and IsOK='0' and PartSN>=0 order by PartSN desc");
            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
            {
                DataTable dttmp = db.DB_GetData($@"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId!='{needID}' and CalendarDate>='{Convert.ToDateTime(dt_tmp.Rows[0]["StartDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}'");
                DataTable dt_SimulationTMP = dttmp;
                DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dt_tmp.Rows[0]["Source_StationNO"]}'");
                intime = TimeCompute2DateTime_BY_ReturnNextShift(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(dt_tmp.Rows[0]["StartDate"]), dltime);
                DataRow workTime_TMP = null;
                bool fristRUN = false;
                bool IsARGs10 = false;
                foreach (DataRow dr_TMP in dt_tmp.Rows)
                {
                    string staionNO = dr_TMP["Source_StationNO"].ToString();
                    int times = int.Parse(dr_TMP["Math_UseTime"].ToString());
                    string err = "";
                    int op_Count = int.Parse(dr_TMP["Math_UseOPCount"].ToString());
                    DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_PP_Station["CalendarName"].ToString()}' and Holiday>='{intime.ToString("yyyy-MM-dd")}' order by Holiday");
                    DateTime etime1 = DateTime.Now;
                    DateTime etime2 = DateTime.Now;
                    DateTime stime1 = DateTime.Now;
                    DateTime stime2 = DateTime.Now;
                    string sql = "";
                    bool byPass = false;
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (Convert.ToDateTime(dr["Holiday"]).ToString("yyyy-MM-dd") != intime.ToString("yyyy-MM-dd"))
                        {
                            if (Convert.ToDateTime(dr["Holiday"].ToString()) < intime) { break; }
                            else
                            {
                                stime2 = Convert.ToDateTime(dr["Holiday"]);
                                stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Millisecond);
                                if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                                { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                                else if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                                { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                                else if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                                { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                                else if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                                { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                                else
                                { break; }
                            }
                        }
                        DataRow[] dr_WorkTimeNote = null;
                        #region Flag_Morning
                        if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                        {
                            #region 取得工作時間
                            int typeTotalTime = 0;
                            string[] comp = dr["Shift_Morning"].ToString().Trim().Split(',');
                            etime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                            stime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                            if (byPass)
                            {
                                typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                            }
                            else
                            {
                                if (IsARGs10)
                                {
                                    if (etime1 >= intime && intime >= stime1)
                                    {
                                        byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                    }
                                    else
                                    {
                                        if (stime1 >= intime) //前置時間
                                        {
                                            byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                        }
                                    }
                                }
                                else
                                {
                                    if (stime1 >= intime)//前置時間
                                    {
                                        byPass = true;
                                    }
                                }
                                if (!byPass)
                                {
                                    etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                    if (IsARGs10)
                                    {
                                        if (etime2 >= intime && intime >= stime2)
                                        {
                                            if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                        }
                                        else
                                        {
                                            if (stime2 >= intime) //前置時間
                                            {
                                                if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (stime2 >= intime)//前置時間
                                        {
                                            byPass = true;
                                        }
                                    }
                                }
                            }
                            #endregion


                            if (typeTotalTime != 0)
                            {
                                //###???要計算合併站的問題

                                #region 多工單, 增加可用工時
                                if (op_Count > 1) { typeTotalTime *= op_Count; }
                                #endregion
                                int TOT = 0;
                                //dr_WorkTimeNote = dt_SimulationTMP.Select($"CONVERT(varchar(100), CalendarDate, 23)='{intime.ToString("yyyy-MM-dd")}' and StationNO='{staionNO}'");
                                stime1 = new DateTime(intime.Year, intime.Month, intime.Day, 0, 0, 0, 1);
                                etime1 = new DateTime(intime.Year, intime.Month, intime.Day, 23, 59, 59, 999);
                                dr_WorkTimeNote = dt_SimulationTMP.Select($"CalendarDate>=#{stime1.ToString("yyyy-MM-dd HH:mm:ss.fff")}# and CalendarDate<=#{etime1.ToString("yyyy-MM-dd HH:mm:ss.fff")}# and StationNO='{staionNO}'");

                                if (dr_WorkTimeNote != null)
                                {
                                    foreach (DataRow d in dr_WorkTimeNote)
                                    {
                                        TOT += int.Parse(d["Time1_C"].ToString());
                                    }
                                }
                                if (TOT > 0)
                                {
                                    #region 已有其他排程, 與其他合計
                                    if (TOT < typeTotalTime)
                                    {
                                        if (times >= Math.Abs((typeTotalTime - TOT)))
                                        {
                                            times -= Math.Abs((typeTotalTime - TOT));
                                            workTime_TMP = dt_SimulationTMP.NewRow();
                                            workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = true; workTime_TMP[5] = false; workTime_TMP[6] = false; workTime_TMP[7] = false; workTime_TMP[8] = Math.Abs((typeTotalTime - TOT)); workTime_TMP[9] = 0; workTime_TMP[10] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                            dt_SimulationTMP.Rows.Add(workTime_TMP);
                                        }
                                        else
                                        {
                                            workTime_TMP = dt_SimulationTMP.NewRow();
                                            workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = true; workTime_TMP[5] = false; workTime_TMP[6] = false; workTime_TMP[7] = false; workTime_TMP[8] = times; workTime_TMP[9] = 0; workTime_TMP[10] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                            dt_SimulationTMP.Rows.Add(workTime_TMP);
                                            intime = intime.AddSeconds((Math.Abs(times) + 60));
                                            break;
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region INSERT
                                    if (times >= typeTotalTime)
                                    {
                                        times -= typeTotalTime;
                                        workTime_TMP = dt_SimulationTMP.NewRow();
                                        workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = true; workTime_TMP[5] = false; workTime_TMP[6] = false; workTime_TMP[7] = false; workTime_TMP[8] = times; workTime_TMP[9] = 0; workTime_TMP[10] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                        dt_SimulationTMP.Rows.Add(workTime_TMP);
                                    }
                                    else
                                    {
                                        workTime_TMP = dt_SimulationTMP.NewRow();
                                        workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = true; workTime_TMP[5] = false; workTime_TMP[6] = false; workTime_TMP[7] = false; workTime_TMP[8] = times; workTime_TMP[9] = 0; workTime_TMP[10] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                        dt_SimulationTMP.Rows.Add(workTime_TMP);

                                        intime = intime.AddSeconds((Math.Abs(times) + 60));
                                        break;
                                    }
                                    #endregion
                                }

                            }
                        }
                        #endregion

                        #region Flag_Afternoon
                        if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                        {
                            #region 取得工作時間
                            int typeTotalTime = 0;
                            string[] comp = dr["Shift_Afternoon"].ToString().Trim().Split(',');
                            etime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                            stime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                            if (byPass)
                            {
                                typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                            }
                            else
                            {
                                if (IsARGs10)
                                {
                                    if (etime1 >= intime && intime >= stime1)
                                    {
                                        byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                    }
                                    else
                                    {
                                        if (stime1 >= intime) //前置時間
                                        {
                                            byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                        }
                                    }
                                }
                                else
                                {
                                    if (stime1 >= intime)//前置時間
                                    {
                                        byPass = true;
                                    }
                                }
                                if (!byPass)
                                {
                                    etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                    if (IsARGs10)
                                    {
                                        if (etime2 >= intime && intime >= stime2)
                                        {
                                            if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                        }
                                        else
                                        {
                                            if (stime2 >= intime) //前置時間
                                            {
                                                if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (stime2 >= intime)//前置時間
                                        {
                                            byPass = true;
                                        }
                                    }
                                }
                            }

                            #endregion

                            if (typeTotalTime != 0)
                            {
                                #region 多工單, 增加可用工時
                                if (op_Count > 1) { typeTotalTime *= op_Count; }
                                #endregion
                                int TOT = 0;
                                stime1 = new DateTime(intime.Year, intime.Month, intime.Day, 0, 0, 0, 1);
                                etime1 = new DateTime(intime.Year, intime.Month, intime.Day, 23, 59, 59, 999);
                                dr_WorkTimeNote = dt_SimulationTMP.Select($"CalendarDate>=#{stime1.ToString("yyyy-MM-dd HH:mm:ss.fff")}# and CalendarDate<=#{etime1.ToString("yyyy-MM-dd HH:mm:ss.fff")}# and StationNO='{staionNO}'");
                                if (dr_WorkTimeNote != null)
                                {
                                    foreach (DataRow d in dr_WorkTimeNote)
                                    {
                                        TOT += int.Parse(d["Time2_C"].ToString());
                                    }
                                }
                                if (TOT > 0)
                                {
                                    #region 已有其他排程, 與其他合計
                                    if (TOT < typeTotalTime)
                                    {
                                        if (times >= Math.Abs((typeTotalTime - TOT)))
                                        {
                                            times -= Math.Abs((typeTotalTime - TOT));
                                            workTime_TMP = dt_SimulationTMP.NewRow();
                                            workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = true; workTime_TMP[6] = false; workTime_TMP[7] = false; workTime_TMP[9] = Math.Abs((typeTotalTime - TOT)); workTime_TMP[8] = 0; workTime_TMP[10] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                            dt_SimulationTMP.Rows.Add(workTime_TMP);
                                        }
                                        else
                                        {
                                            workTime_TMP = dt_SimulationTMP.NewRow();
                                            workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = true; workTime_TMP[6] = false; workTime_TMP[7] = false; workTime_TMP[9] = times; workTime_TMP[8] = 0; workTime_TMP[10] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                            dt_SimulationTMP.Rows.Add(workTime_TMP);
                                            intime = intime.AddSeconds((Math.Abs(times) + 60));
                                            break;
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region INSERT
                                    if (times >= typeTotalTime)
                                    {
                                        times -= typeTotalTime;
                                        workTime_TMP = dt_SimulationTMP.NewRow();
                                        workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = true; workTime_TMP[6] = false; workTime_TMP[7] = false; workTime_TMP[9] = typeTotalTime; workTime_TMP[8] = 0; workTime_TMP[10] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                        dt_SimulationTMP.Rows.Add(workTime_TMP);
                                    }
                                    else
                                    {
                                        workTime_TMP = dt_SimulationTMP.NewRow();
                                        workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = true; workTime_TMP[6] = false; workTime_TMP[7] = false; workTime_TMP[9] = times; workTime_TMP[8] = 0; workTime_TMP[10] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                        dt_SimulationTMP.Rows.Add(workTime_TMP);

                                        intime = intime.AddSeconds((Math.Abs(times) + 60));
                                        break;
                                    }
                                    #endregion
                                }
                            }
                        }
                        #endregion

                        #region Flag_Night
                        if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                        {
                            #region 取得工作時間
                            int typeTotalTime = 0;
                            string[] comp = dr["Shift_Night"].ToString().Trim().Split(',');
                            etime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                            stime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                            if (byPass)
                            {
                                typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;

                            }
                            else
                            {
                                if (IsARGs10)
                                {
                                    if (etime1 >= intime && intime >= stime1)
                                    {
                                        byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                    }
                                    else
                                    {
                                        if (stime1 >= intime) //前置時間
                                        {
                                            byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                        }
                                    }
                                }
                                else
                                {
                                    if (stime1 >= intime)//前置時間
                                    {
                                        byPass = true;
                                    }
                                }
                                if (!byPass)
                                {
                                    if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                    { etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1); }
                                    else
                                    {
                                        etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                    }
                                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                    if (IsARGs10)
                                    {
                                        if (etime2 >= intime && intime >= stime2)
                                        {
                                            if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                        }
                                        else
                                        {
                                            if (stime2 >= intime) //前置時間
                                            {
                                                if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (stime2 >= intime)//前置時間
                                        {
                                            byPass = true;
                                        }
                                    }
                                }
                            }
                            #endregion

                            if (typeTotalTime != 0)
                            {
                                #region 多工單, 增加可用工時
                                if (op_Count > 1) { typeTotalTime *= op_Count; }
                                #endregion
                                int TOT = 0;
                                stime1 = new DateTime(intime.Year, intime.Month, intime.Day, 0, 0, 0, 1);
                                etime1 = new DateTime(intime.Year, intime.Month, intime.Day, 23, 59, 59, 999);
                                dr_WorkTimeNote = dt_SimulationTMP.Select($"CalendarDate>=#{stime1.ToString("yyyy-MM-dd HH:mm:ss.fff")}# and CalendarDate<=#{etime1.ToString("yyyy-MM-dd HH:mm:ss.fff")}# and StationNO='{staionNO}'");
                                if (dr_WorkTimeNote != null)
                                {
                                    foreach (DataRow d in dr_WorkTimeNote)
                                    {
                                        TOT += int.Parse(d["Time3_C"].ToString());
                                    }
                                }
                                if (TOT > 0)
                                {
                                    #region 已有其他排程, 與其他合計
                                    if (TOT < typeTotalTime)
                                    {
                                        if (times >= Math.Abs((typeTotalTime - TOT)))
                                        {
                                            times -= Math.Abs((typeTotalTime - TOT));
                                            workTime_TMP = dt_SimulationTMP.NewRow();
                                            workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = false; workTime_TMP[6] = true; workTime_TMP[7] = false; workTime_TMP[10] = Math.Abs((typeTotalTime - TOT)); workTime_TMP[8] = 0; workTime_TMP[9] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                            dt_SimulationTMP.Rows.Add(workTime_TMP);
                                        }
                                        else
                                        {
                                            workTime_TMP = dt_SimulationTMP.NewRow();
                                            workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = false; workTime_TMP[6] = true; workTime_TMP[7] = false; workTime_TMP[10] = times; workTime_TMP[8] = 0; workTime_TMP[9] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                            dt_SimulationTMP.Rows.Add(workTime_TMP);
                                            intime = intime.AddSeconds((Math.Abs(times) + 60));
                                            break;
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region INSERT
                                    if (times >= typeTotalTime)
                                    {
                                        times -= typeTotalTime;
                                        workTime_TMP = dt_SimulationTMP.NewRow();
                                        workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = false; workTime_TMP[6] = true; workTime_TMP[7] = false; workTime_TMP[10] = typeTotalTime; workTime_TMP[8] = 0; workTime_TMP[9] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                        dt_SimulationTMP.Rows.Add(workTime_TMP);
                                    }
                                    else
                                    {
                                        workTime_TMP = dt_SimulationTMP.NewRow();
                                        workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = false; workTime_TMP[6] = true; workTime_TMP[7] = false; workTime_TMP[10] = times; workTime_TMP[8] = 0; workTime_TMP[9] = 0; workTime_TMP[11] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                        dt_SimulationTMP.Rows.Add(workTime_TMP);

                                        intime = intime.AddSeconds((Math.Abs(times) + 60));
                                        break;
                                    }
                                    #endregion
                                }
                            }
                        }
                        #endregion

                        #region Flag_Graveyard
                        if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                        {
                            #region 取得工作時間
                            bool be_addDay = false;
                            int typeTotalTime = 0;
                            string[] comp_Night = dr["Shift_Night"].ToString().Trim().Split(',');
                            string[] comp = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                            if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                            { be_addDay = true; }
                            if (be_addDay || int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                            {
                                etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1);
                                etime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0).AddDays(1);
                            }
                            else
                            {
                                etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                etime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                            }
                            if (be_addDay)
                            { stime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0).AddDays(1); }
                            else
                            { stime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0); }
                            if (byPass)
                            {
                                typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                            }
                            else
                            {
                                if (IsARGs10)
                                {
                                    if (etime1 >= intime && intime >= stime1)
                                    {
                                        byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                    }
                                    else
                                    {
                                        if (stime1 >= intime) //前置時間
                                        {
                                            byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                        }
                                    }
                                }
                                else
                                {
                                    if (stime1 >= intime)//前置時間
                                    {
                                        byPass = true;
                                    }
                                }
                                if (!byPass)
                                {
                                    if (be_addDay)
                                    { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0).AddDays(1); }
                                    else
                                    { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0); }
                                    if (IsARGs10)
                                    {
                                        if (etime2 >= intime && intime >= stime2)
                                        {
                                            if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                        }
                                        else
                                        {
                                            if (stime2 >= intime) //前置時間
                                            {
                                                if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (stime2 >= intime)//前置時間
                                        {
                                            byPass = true;
                                        }
                                    }
                                }
                            }
                            #endregion
                            if (typeTotalTime != 0)
                            {
                                #region 多工單, 增加可用工時
                                if (op_Count > 1) { typeTotalTime *= op_Count; }
                                #endregion
                                int TOT = 0;
                                stime1 = new DateTime(intime.Year, intime.Month, intime.Day, 0, 0, 0, 1);
                                etime1 = new DateTime(intime.Year, intime.Month, intime.Day, 23, 59, 59, 999);
                                dr_WorkTimeNote = dt_SimulationTMP.Select($"CalendarDate>=#{stime1.ToString("yyyy-MM-dd HH:mm:ss.fff")}# and CalendarDate<=#{etime1.ToString("yyyy-MM-dd HH:mm:ss.fff")}# and StationNO='{staionNO}'");
                                if (dr_WorkTimeNote != null)
                                {
                                    foreach (DataRow d in dr_WorkTimeNote)
                                    {
                                        TOT += int.Parse(d["Time4_C"].ToString());
                                    }
                                }
                                if (TOT > 0)
                                {
                                    #region 已有其他排程, 與其他合計
                                    if (TOT < typeTotalTime)
                                    {
                                        if (times >= Math.Abs((typeTotalTime - TOT)))
                                        {
                                            times -= Math.Abs((typeTotalTime - TOT));
                                            workTime_TMP = dt_SimulationTMP.NewRow();
                                            workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = false; workTime_TMP[6] = false; workTime_TMP[7] = true; workTime_TMP[11] = Math.Abs((typeTotalTime - TOT)); workTime_TMP[8] = 0; workTime_TMP[9] = 0; workTime_TMP[10] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                            dt_SimulationTMP.Rows.Add(workTime_TMP);
                                        }
                                        else
                                        {
                                            workTime_TMP = dt_SimulationTMP.NewRow();
                                            workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = false; workTime_TMP[6] = false; workTime_TMP[7] = true; workTime_TMP[11] = times; workTime_TMP[8] = 0; workTime_TMP[9] = 0; workTime_TMP[10] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                            dt_SimulationTMP.Rows.Add(workTime_TMP);
                                            intime = intime.AddSeconds((Math.Abs(times) + 60));
                                            break;
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region INSERT
                                    if (times >= typeTotalTime)
                                    {
                                        times -= typeTotalTime;
                                        workTime_TMP = dt_SimulationTMP.NewRow();
                                        workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = false; workTime_TMP[6] = false; workTime_TMP[7] = true; workTime_TMP[11] = typeTotalTime; workTime_TMP[8] = 0; workTime_TMP[9] = 0; workTime_TMP[10] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                        dt_SimulationTMP.Rows.Add(workTime_TMP);
                                    }
                                    else
                                    {
                                        workTime_TMP = dt_SimulationTMP.NewRow();
                                        workTime_TMP[0] = staionNO; workTime_TMP[1] = intime.ToString("yyyy-MM-dd HH:mm:ss.fff"); workTime_TMP[2] = needID; ; workTime_TMP[3] = dr_TMP["SimulationId"].ToString(); workTime_TMP[4] = false; workTime_TMP[5] = false; workTime_TMP[6] = false; workTime_TMP[7] = true; workTime_TMP[11] = times; workTime_TMP[8] = 0; workTime_TMP[9] = 0; workTime_TMP[10] = 0; ; workTime_TMP[12] = ""; workTime_TMP[13] = 0; workTime_TMP[14] = _Str.NewId('Y'); workTime_TMP[15] = dr_TMP["StationNO_Merge"].ToString();
                                        dt_SimulationTMP.Rows.Add(workTime_TMP);

                                        intime = intime.AddSeconds((Math.Abs(times) + 60));
                                        break;
                                    }
                                    #endregion
                                }
                            }
                        }
                        #endregion
                    }
                }
            }

            return intime;
        }



        public DateTime GetFinallyDate(RunSimulation_Arg args, DBADO db, string needId, string simulationId, string mathType, string staionNO, string partNO, string partClass, string needQTY, string calendarName, DateTime intime, int times, int op_Count, object StationNO_Merge, ref string err, bool IsARGs10 = false, bool IsRUN_insert_APS_PartNOTimeNote = true)
        {
            if (mathType == "站")
            {
                if (staionNO == "A01" || staionNO == "A02")
                {
                    string _s = "";
                }
            }
            else
            {
                if (partNO == "1_2_1_1採購1")
                {
                    string _s = "";
                }
            }


            //###??? 城市函式 GetFinallyDate 要加一個 急單向下一個時段插單 (急單用)
            string m_APS_StationNO = staionNO;
            List<string> S_MergeList = new List<string>();
            if (StationNO_Merge == null)
            {
                S_MergeList.Add(staionNO);
                StationNO_Merge = "NULL";
            }
            else
            {
                S_MergeList.AddRange(StationNO_Merge.ToString().Trim().Substring(0, StationNO_Merge.ToString().Trim().Length - 1).Split(',').ToList());
                StationNO_Merge = $"'{StationNO_Merge.ToString()}'";
            }

            if (times <= 0) { return intime.AddSeconds(60); }
            DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{calendarName}' and Holiday>='{intime.ToString("yyyy-MM-dd")}' order by Holiday");
            if (dt == null || dt.Rows.Count <= 0) { err = "行事曆日期不足計算."; goto break_FUN; }
            DateTime etime1 = DateTime.Now;
            DateTime etime2 = DateTime.Now;
            DateTime stime1 = DateTime.Now;
            DateTime stime2 = DateTime.Now;
            string sql = "";
            bool byPass = false;

            DateTime first_intime = intime;
            DateTime finish_MAX_intime = intime;
            int holiday_Count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                holiday_Count++;
                for (int i = 0; i < S_MergeList.Count; ++i)
                {
                    staionNO = S_MergeList[i];
                    if (i == 0) { first_intime = intime; }
                    else if (i > 0)
                    {
                        intime = first_intime;
                        if (holiday_Count == 1 && byPass) { byPass = false; }
                    }

                    #region 跨日
                    if (Convert.ToDateTime(dr["Holiday"]).ToString("yyyy-MM-dd") != intime.ToString("yyyy-MM-dd"))
                    {
                        if (Convert.ToDateTime(dr["Holiday"].ToString()) < intime) { break; }
                        else
                        {
                            stime2 = Convert.ToDateTime(dr["Holiday"]);
                            stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, DateTime.Now.Hour, DateTime.Now.Minute, 0, 0);
                            if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                            else if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                            else if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                            else if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                            else
                            { break; }
                        }
                    }
                    #endregion
                    DataRow dr_WorkTimeNote = null;

                    //###???若以下 有改 則 APSViewController 的 TMP_Simulation_APS_WorkTimeNote ㄝ要改

                    #region Flag_Morning
                    if (times > 0 && bool.Parse(dr["Flag_Morning"].ToString()) == true)
                    {
                        #region 取得工作時間
                        int typeTotalTime = 0;
                        string[] comp = dr["Shift_Morning"].ToString().Trim().Split(',');
                        etime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                        stime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                        if (byPass)
                        {
                            typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                        }
                        else
                        {
                            if (IsARGs10)
                            {
                                if (etime1 >= intime && intime >= stime1)
                                {
                                    byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                }
                                else
                                {
                                    if (stime1 >= intime) //前置時間
                                    {
                                        byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                    }
                                }
                            }
                            else
                            {
                                if (stime1 >= intime)//前置時間
                                {
                                    byPass = true;
                                }
                            }
                            if (!byPass)
                            {
                                etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                if (IsARGs10)
                                {
                                    if (etime2 >= intime && intime >= stime2)
                                    {
                                        if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                    }
                                    else
                                    {
                                        if (stime2 >= intime) //前置時間
                                        {
                                            if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                        }
                                    }
                                }
                                else
                                {
                                    if (stime2 >= intime)//前置時間
                                    {
                                        byPass = true;
                                    }
                                }
                            }
                        }
                        #endregion
                        if (typeTotalTime != 0)
                        {
                            if (mathType == "站")
                            {
                                //###???要計算合併站的問題

                                #region 多工單, 增加可用工時
                                if (op_Count > 1) { typeTotalTime *= op_Count; }
                                #endregion

                                dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time1_C) as TOT from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where CONVERT(varchar(100), CalendarDate, 23)='{intime.ToString("yyyy-MM-dd")}' and StationNO='{staionNO}'");
                                if (dr_WorkTimeNote != null && !dr_WorkTimeNote.IsNull("TOT") && int.Parse(dr_WorkTimeNote["TOT"].ToString()) > 0)
                                {
                                    #region 已有其他排程, 與其他合計
                                    if (int.Parse(dr_WorkTimeNote["TOT"].ToString()) < typeTotalTime)
                                    {
                                        intime = intime.AddSeconds(int.Parse(dr_WorkTimeNote["TOT"].ToString()));
                                        int tmp_time = typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString());
                                        if (times >= tmp_time)
                                        {
                                            times -= tmp_time;
                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type1,Time1_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{tmp_time.ToString()},{tmp_time.ToString()})"))
                                            {
                                                if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                                else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }
                                            }
                                            else
                                            {
                                                err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                            }
                                        }
                                        else
                                        {
                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type1,Time1_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{times.ToString()},{times.ToString()})"))
                                            {
                                                if (IsRUN_insert_APS_PartNOTimeNote)
                                                {
                                                    GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10);
                                                }
                                                else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }

                                            }
                                            else
                                            {
                                                err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                            }
                                            if (intime.AddSeconds(times) > finish_MAX_intime) { finish_MAX_intime = intime.AddSeconds(times); }
                                            times = 0;
                                            //return intime.AddSeconds(times + 60);
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region INSERT
                                    if (times >= typeTotalTime)
                                    {
                                        times -= typeTotalTime;
                                        if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type1,Time1_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{typeTotalTime.ToString()},{typeTotalTime.ToString()})"))
                                        {
                                            if (IsRUN_insert_APS_PartNOTimeNote)
                                            {
                                                GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10);
                                            }
                                            else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }

                                        }
                                        else
                                        {
                                            err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                        }
                                    }
                                    else
                                    {
                                        if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type1,Time1_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{times.ToString()},{times.ToString()})"))
                                        {
                                            if (IsRUN_insert_APS_PartNOTimeNote)
                                            {
                                                GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10);
                                            }
                                            else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }

                                        }
                                        else
                                        {
                                            err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                        }
                                        if (intime.AddSeconds(times) > finish_MAX_intime) { finish_MAX_intime = intime.AddSeconds(times); }
                                        times = 0;
                                        //return intime.AddSeconds((times) + 60);
                                    }
                                    #endregion
                                }
                            }
                            else
                            {
                                #region 料
                                //DateTime partMOTime = intime.AddSeconds(-times);//###???將來要參數化 + 或 - 或 同時 (同時有人機畫面參數)  default為 -
                                DateTime partMOTime = intime;
                                if (_Fun.Config.OutPackStationName == staionNO)
                                { partMOTime = TimeCompute2DateTime(db, calendarName, intime, times, true); }
                                else { partMOTime = TimeCompute2DateTime(db, calendarName, intime, times, false); }

                                if (IsARGs10) { partMOTime = intime; }
                                if (IsRUN_insert_APS_PartNOTimeNote)
                                {
                                    GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, partMOTime, partMOTime, partClass, needId, simulationId, needQTY, staionNO, null, IsARGs10);
                                }
                                return partMOTime.AddSeconds((Math.Abs(times) + 60));
                                #endregion
                            }
                        }
                    }
                    #endregion

                    #region Flag_Afternoon
                    if (times > 0 && bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                    {
                        #region 取得工作時間
                        int typeTotalTime = 0;
                        string[] comp = dr["Shift_Afternoon"].ToString().Trim().Split(',');
                        etime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                        stime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                        if (byPass)
                        {
                            typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                        }
                        else
                        {
                            if (IsARGs10)
                            {
                                if (etime1 >= intime && intime >= stime1)
                                {
                                    byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                }
                                else
                                {
                                    if (stime1 >= intime) //前置時間
                                    {
                                        byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                    }
                                }
                            }
                            else
                            {
                                if (stime1 >= intime)//前置時間
                                {
                                    byPass = true;
                                }
                            }
                            if (!byPass)
                            {
                                etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                if (IsARGs10)
                                {
                                    if (etime2 >= intime && intime >= stime2)
                                    {
                                        if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                    }
                                    else
                                    {
                                        if (stime2 >= intime) //前置時間
                                        {
                                            if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                        }
                                    }
                                }
                                else
                                {
                                    if (stime2 >= intime)//前置時間
                                    {
                                        byPass = true;
                                    }
                                }
                            }
                        }

                        #endregion
                        if (typeTotalTime != 0)
                        {
                            if (mathType == "站")
                            {
                                #region 多工單, 增加可用工時
                                if (op_Count > 1) { typeTotalTime *= op_Count; }
                                #endregion

                                dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time2_C) as TOT from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{staionNO}' and CONVERT(varchar(100), CalendarDate, 23)='{intime.ToString("yyyy-MM-dd")}'");
                                if (dr_WorkTimeNote != null && !dr_WorkTimeNote.IsNull("TOT") && int.Parse(dr_WorkTimeNote["TOT"].ToString()) > 0)
                                {
                                    #region 已有其他排程, 與其他合計
                                    if (int.Parse(dr_WorkTimeNote["TOT"].ToString()) < typeTotalTime)
                                    {
                                        intime = intime.AddSeconds(int.Parse(dr_WorkTimeNote["TOT"].ToString()));
                                        int tmp_time = typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString());
                                        if (times >= tmp_time)
                                        {
                                            times -= tmp_time;
                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type2,Time2_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{tmp_time.ToString()},{tmp_time.ToString()})"))
                                            {
                                                if (IsRUN_insert_APS_PartNOTimeNote)
                                                {
                                                    GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10);
                                                }
                                                else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }

                                            }
                                            else
                                            {
                                                err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                            }
                                        }
                                        else
                                        {
                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type2,Time2_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{times.ToString()},{times.ToString()})"))
                                            {
                                                if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                                else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }
                                            }
                                            else
                                            {
                                                err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                            }
                                            if (intime.AddSeconds(times) > finish_MAX_intime) { finish_MAX_intime = intime.AddSeconds(times); }
                                            times = 0;
                                            //return intime.AddSeconds(times + 60);
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region INSERT
                                    if (times >= typeTotalTime)
                                    {
                                        times -= typeTotalTime;
                                        if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type2,Time2_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{typeTotalTime.ToString()},{typeTotalTime.ToString()})"))
                                        {
                                            if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                            else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }
                                        }
                                        else
                                        {
                                            err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                        }
                                    }
                                    else
                                    {
                                        if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type2,Time2_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{times.ToString()},{times.ToString()})"))
                                        {
                                            if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                            else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }
                                        }
                                        else
                                        {
                                            err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                        }
                                        if (intime.AddSeconds(times) > finish_MAX_intime) { finish_MAX_intime = intime.AddSeconds(times); }
                                        times = 0;
                                        //return intime.AddSeconds((times + 60));
                                    }
                                    #endregion
                                }
                            }
                            else
                            {
                                #region 料
                                //DateTime partMOTime = intime.AddSeconds(-times);//###???將來要參數化 + 或 - 或 同時 (同時有人機畫面參數)
                                DateTime partMOTime = intime;
                                if (_Fun.Config.OutPackStationName == staionNO)
                                { partMOTime = TimeCompute2DateTime(db, calendarName, intime, times, true); }
                                else { partMOTime = TimeCompute2DateTime(db, calendarName, intime, times, false); }

                                if (IsARGs10) { partMOTime = intime; }
                                if (IsRUN_insert_APS_PartNOTimeNote)
                                {
                                    GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, partMOTime, partMOTime, partClass, needId, simulationId, needQTY, staionNO, null, IsARGs10);
                                }
                                return partMOTime.AddSeconds((Math.Abs(times) + 60));
                                #endregion
                            }
                        }
                    }
                    #endregion

                    #region Flag_Night
                    if (times > 0 && bool.Parse(dr["Flag_Night"].ToString()) == true)
                    {
                        #region 取得工作時間
                        int typeTotalTime = 0;
                        string[] comp = dr["Shift_Night"].ToString().Trim().Split(',');
                        etime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                        stime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                        if (byPass)
                        {
                            typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;

                        }
                        else
                        {
                            if (IsARGs10)
                            {
                                if (etime1 >= intime && intime >= stime1)
                                {
                                    byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                }
                                else
                                {
                                    if (stime1 >= intime) //前置時間
                                    {
                                        byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                    }
                                }
                            }
                            else
                            {
                                if (stime1 >= intime)//前置時間
                                {
                                    byPass = true;
                                }
                            }
                            if (!byPass)
                            {
                                if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                { etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1); }
                                else
                                {
                                    etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                                }
                                stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                                if (IsARGs10)
                                {
                                    if (etime2 >= intime && intime >= stime2)
                                    {
                                        if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                    }
                                    else
                                    {
                                        if (stime2 >= intime) //前置時間
                                        {
                                            if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                        }
                                    }
                                }
                                else
                                {
                                    if (stime2 >= intime)//前置時間
                                    {
                                        byPass = true;
                                    }
                                }
                            }
                        }
                        #endregion
                        if (typeTotalTime != 0)
                        {
                            if (mathType == "站")
                            {
                                #region 多工單, 增加可用工時
                                if (op_Count > 1) { typeTotalTime *= op_Count; }
                                #endregion
                                //###???要計算合併站的問題
                                dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time3_C) as TOT from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{staionNO}' and CONVERT(varchar(100), CalendarDate, 23)='{intime.ToString("yyyy-MM-dd")}'");
                                if (dr_WorkTimeNote != null && !dr_WorkTimeNote.IsNull("TOT") && int.Parse(dr_WorkTimeNote["TOT"].ToString()) > 0)
                                {
                                    #region 已有其他排程, 與其他合計
                                    if (int.Parse(dr_WorkTimeNote["TOT"].ToString()) < typeTotalTime)
                                    {
                                        intime = intime.AddSeconds(int.Parse(dr_WorkTimeNote["TOT"].ToString()));
                                        int tmp_time = typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString());
                                        if (times >= tmp_time)
                                        {
                                            times -= tmp_time;
                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type3,Time3_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{tmp_time.ToString()},{tmp_time.ToString()})"))
                                            {
                                                if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                                else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }

                                            }
                                            else
                                            {
                                                err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                            }
                                        }
                                        else
                                        {
                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type3,Time3_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{times.ToString()},{times.ToString()})"))
                                            {
                                                if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                                else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }

                                            }
                                            else
                                            {
                                                err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                            }
                                            if (intime.AddSeconds(times) > finish_MAX_intime) { finish_MAX_intime = intime.AddSeconds(times); }
                                            times = 0;
                                            //return intime.AddSeconds((times + 60));
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region INSERT
                                    if (times >= typeTotalTime)
                                    {
                                        times -= typeTotalTime;
                                        if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type3,Time3_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{typeTotalTime.ToString()},{typeTotalTime.ToString()})"))
                                        {
                                            if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                            else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }

                                        }
                                        else
                                        {
                                            err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                        }
                                    }
                                    else
                                    {
                                        if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type3,Time3_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{times.ToString()},{times.ToString()})"))
                                        {
                                            if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                            else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }

                                        }
                                        else
                                        {
                                            err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                        }
                                        if (intime.AddSeconds(times) > finish_MAX_intime) { finish_MAX_intime = intime.AddSeconds(times); }
                                        times = 0;
                                        return intime.AddSeconds((times + 60));
                                    }
                                    #endregion
                                }
                            }
                            else
                            {
                                #region 料
                                //DateTime partMOTime = intime.AddSeconds(-times);//###???將來要參數化 + 或 - 或 同時 (同時有人機畫面參數)
                                DateTime partMOTime = intime;
                                if (_Fun.Config.OutPackStationName == staionNO)
                                { partMOTime = TimeCompute2DateTime(db, calendarName, intime, times, true); }
                                else { partMOTime = TimeCompute2DateTime(db, calendarName, intime, times, false); }

                                if (IsARGs10) { partMOTime = intime; }
                                if (IsRUN_insert_APS_PartNOTimeNote)
                                {
                                    GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, partMOTime, partMOTime, partClass, needId, simulationId, needQTY, staionNO, null, IsARGs10);
                                }
                                return partMOTime.AddSeconds((Math.Abs(times) + 60));
                                #endregion
                            }
                        }
                    }
                    #endregion

                    #region Flag_Graveyard
                    if (times > 0 && bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                    {
                        #region 取得工作時間
                        bool be_addDay = false;
                        int typeTotalTime = 0;
                        string[] comp_Night = dr["Shift_Night"].ToString().Trim().Split(',');
                        string[] comp = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                        if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                        { be_addDay = true; }
                        if (be_addDay || int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                        {
                            etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1);
                            etime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0).AddDays(1);
                        }
                        else
                        {
                            etime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                            etime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                        }
                        if (be_addDay)
                        { stime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0).AddDays(1); }
                        else
                        { stime1 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0); }
                        if (byPass)
                        {
                            typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                        }
                        else
                        {
                            if (IsARGs10)
                            {
                                if (etime1 >= intime && intime >= stime1)
                                {
                                    byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                }
                                else
                                {
                                    if (stime1 >= intime) //前置時間
                                    {
                                        byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1;
                                    }
                                }
                            }
                            else
                            {
                                if (stime1 >= intime)//前置時間
                                {
                                    byPass = true;
                                }
                            }
                            if (!byPass)
                            {
                                if (be_addDay)
                                { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0).AddDays(1); }
                                else
                                { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0); }
                                if (IsARGs10)
                                {
                                    if (etime2 >= intime && intime >= stime2)
                                    {
                                        if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                    }
                                    else
                                    {
                                        if (stime2 >= intime) //前置時間
                                        {
                                            if (typeTotalTime == 0) { byPass = true; typeTotalTime = TimeCompute2Seconds_BY_Class(comp); intime = stime1; }
                                        }
                                    }
                                }
                                else
                                {
                                    if (stime2 >= intime)//前置時間
                                    {
                                        byPass = true;
                                    }
                                }
                            }
                        }
                        #endregion
                        if (typeTotalTime != 0)
                        {
                            if (mathType == "站")
                            {
                                #region 多工單, 增加可用工時
                                if (op_Count > 1) { typeTotalTime *= op_Count; }
                                #endregion
                                dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"select sum(Time4_C) as TOT from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{staionNO}' and CONVERT(varchar(100), CalendarDate, 23)='{intime.ToString("yyyy-MM-dd")}'");
                                if (dr_WorkTimeNote != null && !dr_WorkTimeNote.IsNull("TOT") && int.Parse(dr_WorkTimeNote["TOT"].ToString()) > 0)
                                {
                                    #region 已有其他排程, 與其他合計
                                    if (int.Parse(dr_WorkTimeNote["TOT"].ToString()) < typeTotalTime)
                                    {
                                        intime = intime.AddSeconds(int.Parse(dr_WorkTimeNote["TOT"].ToString()));
                                        int tmp_time = typeTotalTime - int.Parse(dr_WorkTimeNote["TOT"].ToString());
                                        if (times >= tmp_time)
                                        {
                                            times -= tmp_time;
                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type4,Time4_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{tmp_time.ToString()},{tmp_time.ToString()})"))
                                            {
                                                if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                                else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }

                                            }
                                            else
                                            {
                                                err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                            }
                                        }
                                        else
                                        {
                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type4,Time4_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{times.ToString()},{times.ToString()})"))
                                            {
                                                if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                                else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }
                                            }
                                            else
                                            {
                                                err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                            }
                                            if (intime.AddSeconds(times) > finish_MAX_intime) { finish_MAX_intime = intime.AddSeconds(times); }
                                            times = 0;
                                            //return intime.AddSeconds((times + 60));
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region INSERT
                                    if (times >= typeTotalTime)
                                    {
                                        times -= typeTotalTime;
                                        if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type4,Time4_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{typeTotalTime.ToString()},{typeTotalTime.ToString()})"))
                                        {
                                            if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                            else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }

                                        }
                                        else
                                        {
                                            err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                        }
                                    }
                                    else
                                    {
                                        if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (ServerId,StationNO_Merge,Id,StationNO,CalendarDate,NeedId,SimulationId,Type4,Time4_C,Time_TOT) VALUES ('{_Fun.Config.ServerId}',{StationNO_Merge},'{_Str.NewId('Y')}','{staionNO}','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','{simulationId}','1',{times.ToString()},{times.ToString()})"))
                                        {
                                            if (IsRUN_insert_APS_PartNOTimeNote) { GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, intime, intime, partClass, needId, simulationId, needQTY, m_APS_StationNO, StationNO_Merge, IsARGs10); }
                                            else { db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET CalendarDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where PartNO='{partNO}' and NeedId='{needId}' and SimulationId='{simulationId}'"); }

                                        }
                                        else
                                        {
                                            err = $"WorkTimeNote寫入失敗."; goto break_FUN;
                                        }
                                        if (intime.AddSeconds(times) > finish_MAX_intime) { finish_MAX_intime = intime.AddSeconds(times); }
                                        times = 0;
                                        //return intime.AddSeconds((times + 60));
                                    }
                                    #endregion
                                }
                            }
                            else
                            {
                                #region 料
                                //DateTime partMOTime = intime.AddSeconds(-times);//###???將來要參數化 + 或 - 或 同時 (同時有人機畫面參數)
                                DateTime partMOTime = intime;
                                if (_Fun.Config.OutPackStationName == staionNO)
                                { partMOTime = TimeCompute2DateTime(db, calendarName, intime, times, true); }
                                else { partMOTime = TimeCompute2DateTime(db, calendarName, intime, times, false); }

                                if (IsARGs10) { partMOTime = intime; }
                                if (IsRUN_insert_APS_PartNOTimeNote)
                                {
                                    GetFinallyDate_insert_APS_PartNOTimeNote(db, partNO, partMOTime, partMOTime, partClass, needId, simulationId, needQTY, staionNO, null, IsARGs10);
                                }
                                return partMOTime.AddSeconds((Math.Abs(times) + 60));
                                #endregion
                            }
                        }
                    }
                    #endregion
                    if (times == 0) { return finish_MAX_intime.AddSeconds(120); }
                }
            }
        break_FUN:

            return intime.AddSeconds((Math.Abs(times) + 120));
        }

        private void GetFinallyDate_insert_APS_PartNOTimeNote(DBADO db, string partNO, DateTime intime, DateTime stime, string partClass, string needId, string simulationId, string needQTY, string stationNO, object StationNO_Merge, bool IsARGs = false)
        {
            if (partNO != "")
            {
                DateTime tmp_date = new DateTime(intime.Year, intime.Month, intime.Day, stime.Hour, stime.Minute, stime.Second, stime.Millisecond);
                //if (!IsARGs)
                //{
                //    tmp_date = tmp_date.AddMinutes(store_OpenDOC_AdvanceTime);
                //}

                string outPackType = "0";
                string macID = "";
                string NoStation = "1";
                if (stationNO != "")
                {
                    NoStation = "0";
                    DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT Config_macID FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}'");
                    if (tmp != null) { macID = tmp["Config_macID"].ToString(); }
                }
                if (db.DB_GetQueryCount($"SELECT PartNO from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where PartNO='{partNO}' and SimulationId='{simulationId}' and NeedId='{needId}'") <= 0)
                {
                    if (StationNO_Merge == null) { StationNO_Merge = "NULL"; }
                    DataRow dr_WorkTimeNote = db.DB_GetFirstDataByDataRow($"SELECT OutPackType FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{simulationId}'");
                    if (bool.Parse(dr_WorkTimeNote["OutPackType"].ToString()))
                    { outPackType = "1"; }

                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] (ServerId,OutPackType,NoStation,StationNO_Merge,APS_StationNO,PartNO,CalendarDate,Class,NeedId,SimulationId,NeedQTY,macID) VALUES ('{_Fun.Config.ServerId}','{outPackType}','{NoStation}',{StationNO_Merge},'{stationNO}','{partNO}','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{partClass}','{needId}','{simulationId}',{needQTY},'{macID}')"))
                    { db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET StartDate='{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where SimulationId='{simulationId}'"); }
                }
            }
        }
        public int TimeCompute2Seconds(DateTime start, DateTime end)
        {
            int cycleTime = 0;
            if (start >= end)
            {
                return cycleTime;
            }
            TimeSpan ts = new TimeSpan(end.Ticks - start.Ticks);
            if (ts.TotalSeconds > 0)
            { cycleTime = (int)ts.TotalSeconds; }
            return cycleTime;
        }
        public int TimeCompute2Seconds_BY_Class(string[] comp)
        {
            if (comp == null) { return 0; }
            int cycleTime = 0;
            DateTime intime = DateTime.Now;
            DateTime etime = DateTime.Now;
            DateTime stime = DateTime.Now;
            if (int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
            { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0).AddDays(1); }
            else { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0); }
            stime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
            cycleTime = TimeCompute2Seconds(stime, etime);
            if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
            { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1); }
            else { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0); }
            stime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
            cycleTime += TimeCompute2Seconds(stime, etime);
            return cycleTime;
        }
        public DateTime TimeCompute2DateTime(DBADO db, string calendarName, DateTime start, int changSecond, bool IsAdd = true)
        {
            if (changSecond == 0) { return start; }
            if (IsAdd)
            {
                int cycleTime = 0;
                DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{calendarName}' and Holiday>='{start.ToString("yyyy-MM-dd")}' order by Holiday");
                DateTime intime = start;
                DateTime etime = DateTime.Now;
                DateTime stime2 = DateTime.Now;
                bool fristRUN = false;
                foreach (DataRow dr in dt.Rows)
                {
                    if (Convert.ToDateTime(dr["Holiday"]).ToString("yyyy-MM-dd") != intime.ToString("yyyy-MM-dd"))
                    {
                        if (Convert.ToDateTime(dr["Holiday"].ToString()) < intime) { break; }
                        else
                        {
                            stime2 = Convert.ToDateTime(dr["Holiday"]);
                            stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Millisecond);
                            if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                            else if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                            else if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                            else if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                            else
                            { break; }
                        }
                    }
                    #region Flag_Morning
                    if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                    {
                        string[] comp = dr["Shift_Morning"].ToString().Trim().Split(',');
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(changSecond); }
                            }
                        }
                        else
                        {
                            if (stime2 > intime) //中間休息時間
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                        }
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(changSecond); }
                            }
                        }
                        else
                        {
                            if (stime2 > intime)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                        }
                    }
                    #endregion

                    #region Flag_Afternoon
                    if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                    {
                        string[] comp = dr["Shift_Afternoon"].ToString().Trim().Split(',');
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(changSecond); }
                            }
                        }
                        else
                        {
                            if (stime2 > intime)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                        }
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(changSecond); }
                            }
                        }
                        else
                        {
                            if (stime2 > intime)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                        }
                    }
                    #endregion

                    #region Flag_Night
                    if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                    {
                        string[] comp = dr["Shift_Night"].ToString().Trim().Split(',');
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(changSecond); }
                            }
                        }
                        else
                        {
                            if (stime2 > intime)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                        }
                        if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                        { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), intime.Second, intime.Millisecond).AddDays(1); }
                        else { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), intime.Second, intime.Millisecond); }
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(changSecond); }
                            }
                        }
                        else
                        {
                            if (stime2 > intime)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                        }
                    }
                    #endregion

                    #region Flag_Graveyard
                    if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                    {
                        bool be_addDay = false;
                        string[] comp_Night = dr["Shift_Night"].ToString().Trim().Split(',');
                        string[] comp = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                        if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                        { be_addDay = true; }
                        if (be_addDay || int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                        { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), intime.Second, intime.Millisecond).AddDays(1); }
                        else { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), intime.Second, intime.Millisecond); }
                        if (be_addDay)
                        { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), intime.Second, intime.Millisecond).AddDays(1); }
                        else
                        { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), intime.Second, intime.Millisecond); }
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(changSecond); }
                            }
                        }
                        else
                        {
                            if (stime2 > intime)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                        }
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return stime2.AddSeconds(changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(changSecond); }
                            }
                        }
                    }
                    #endregion

                }
            }
            else
            {
                int cycleTime = 0;
                DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{calendarName}' and Holiday<='{start.ToString("yyyy-MM-dd")}' order by Holiday desc");
                DateTime intime = start;
                DateTime etime = DateTime.Now;
                DateTime stime2 = DateTime.Now;
                bool fristRUN = false;
                foreach (DataRow dr in dt.Rows)
                {
                    if (Convert.ToDateTime(dr["Holiday"]).ToString("yyyy-MM-dd") != intime.ToString("yyyy-MM-dd"))
                    {
                        if (Convert.ToDateTime(dr["Holiday"].ToString()) > intime) { break; }
                        else
                        {
                            stime2 = Convert.ToDateTime(dr["Holiday"]);
                            stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Millisecond);
                            if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                            else if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                            else if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                            else if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                            { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                            else
                            { break; }
                        }
                    }
                    #region Flag_Graveyard
                    if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                    {
                        bool be_addDay = false;
                        string[] comp_Night = dr["Shift_Night"].ToString().Trim().Split(',');
                        string[] comp = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                        if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                        { be_addDay = true; }
                        if (be_addDay || int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                        { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), intime.Second, intime.Millisecond).AddDays(1); }
                        else { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), intime.Second, intime.Millisecond); }
                        if (be_addDay)
                        { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), intime.Second, intime.Millisecond).AddDays(1); }
                        else
                        { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), intime.Second, intime.Millisecond); }
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return etime.AddSeconds(-changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(-changSecond); }
                            }
                        }

                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return etime.AddSeconds(-changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(-changSecond); }
                            }
                        }
                    }
                    #endregion

                    #region Flag_Night
                    if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                    {
                        string[] comp = dr["Shift_Night"].ToString().Trim().Split(',');
                        if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                        { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), intime.Second, intime.Millisecond).AddDays(1); }
                        else { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), intime.Second, intime.Millisecond); }
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return etime.AddSeconds(-changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(-changSecond); }
                            }
                        }
                        else
                        {
                            if (intime > etime)//休息時間
                            {
                                cycleTime = TimeCompute2Seconds(etime, intime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return etime.AddSeconds(-changSecond); }
                            }
                        }
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return etime.AddSeconds(-changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(-changSecond); }
                            }
                        }
                    }
                    #endregion

                    #region Flag_Afternoon
                    if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                    {
                        string[] comp = dr["Shift_Afternoon"].ToString().Trim().Split(',');
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return etime.AddSeconds(-changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(-changSecond); }
                            }
                        }
                        else
                        {
                            if (intime > etime)//休息時間
                            {
                                cycleTime = TimeCompute2Seconds(etime, intime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return etime.AddSeconds(-changSecond); }
                            }
                        }
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return etime.AddSeconds(-changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(-changSecond); }
                            }
                        }
                    }
                    #endregion

                    #region Flag_Morning
                    if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                    {
                        string[] comp = dr["Shift_Morning"].ToString().Trim().Split(',');
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return etime.AddSeconds(-changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(-changSecond); }
                            }
                        }
                        else
                        {
                            if (intime > etime)//休息時間
                            {
                                cycleTime = TimeCompute2Seconds(etime, intime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return etime.AddSeconds(-changSecond); }
                            }
                        }

                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), intime.Second, intime.Millisecond);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), intime.Second, intime.Millisecond);
                        if (fristRUN || (etime >= intime && intime >= stime2))
                        {
                            if (fristRUN)
                            {
                                cycleTime = TimeCompute2Seconds(stime2, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; }
                                else { return etime.AddSeconds(-changSecond); }
                            }
                            else
                            {
                                cycleTime = TimeCompute2Seconds(intime, etime);
                                if (changSecond > cycleTime) { changSecond -= cycleTime; fristRUN = true; }
                                else { return intime.AddSeconds(-changSecond); }
                            }
                        }
                    }
                    #endregion

                }
            }
            return DateTime.Now;
        }
        public DateTime TimeCompute2DateTime_BY_ReturnNextShift(DBADO db, string calendarName, DateTime start, int addSS)
        {
            if (addSS == 0) { return start; }
            start = start.AddSeconds(addSS);
            DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{calendarName}' and Holiday>='{start.ToString("yyyy-MM-dd")}' order by Holiday");
            DateTime intime = start;
            DateTime etime = DateTime.Now;
            DateTime stime2 = DateTime.Now;
            foreach (DataRow dr in dt.Rows)
            {
                if (Convert.ToDateTime(dr["Holiday"]).ToString("yyyy-MM-dd") != intime.ToString("yyyy-MM-dd"))
                {
                    if (Convert.ToDateTime(dr["Holiday"].ToString()) < intime) { break; }
                    else
                    {
                        stime2 = Convert.ToDateTime(dr["Holiday"]);
                        stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Millisecond);
                        if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                        { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                        else if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                        { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                        else if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                        { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                        else if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                        { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), stime2.Second, stime2.Millisecond); }
                        else
                        { break; }
                    }
                }
                #region Flag_Morning
                if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                {
                    string[] comp = dr["Shift_Morning"].ToString().Trim().Split(',');
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                    if (etime >= intime && intime >= stime2)
                    {
                        return new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                    }
                    else
                    {
                        if (stime2 > intime) //中間休息時間
                        {
                            return stime2;
                        }
                    }
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                    if (etime >= intime && intime >= stime2)
                    {
                        comp = dr["Shift_Afternoon"].ToString().Trim().Split(',');
                        return new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                    }
                    else
                    {
                        if (stime2 > intime) //中間休息時間
                        {
                            return stime2;
                        }
                    }
                }
                #endregion

                #region Flag_Afternoon
                if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                {
                    string[] comp = dr["Shift_Afternoon"].ToString().Trim().Split(',');
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                    if (etime >= intime && intime >= stime2)
                    {
                        return new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                    }
                    else
                    {
                        if (stime2 > intime) //中間休息時間
                        {
                            return stime2;
                        }
                    }
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                    if (etime >= intime && intime >= stime2)
                    {
                        comp = dr["Shift_Night"].ToString().Trim().Split(',');
                        return new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                    }
                    else
                    {
                        if (stime2 > intime) //中間休息時間
                        {
                            return stime2;
                        }
                    }
                }
                #endregion

                #region Flag_Night
                if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                {
                    string[] comp = dr["Shift_Night"].ToString().Trim().Split(',');
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                    if (etime >= intime && intime >= stime2)
                    {
                        return new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                    }
                    else
                    {
                        if (stime2 > intime) //中間休息時間
                        {
                            return stime2;
                        }
                    }
                    if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                    { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1); }
                    else
                    {
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                    }
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                    if (etime >= intime && intime >= stime2)
                    {
                        bool be_addDay = false;
                        string[] comp_Graveyard = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                        if (int.Parse(comp[3].Split(':')[0]) > int.Parse(comp_Graveyard[0].Split(':')[0]))
                        { be_addDay = true; }
                        if (be_addDay)
                        { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0).AddDays(1); }
                        else
                        { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0); }
                        return stime2;
                    }
                    else
                    {
                        if (stime2 > intime) //中間休息時間
                        {
                            return stime2;
                        }
                    }
                }
                #endregion

                #region Flag_Graveyard
                if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                {
                    bool be_addDay = false;
                    string[] comp_Night = dr["Shift_Night"].ToString().Trim().Split(',');
                    string[] comp = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                    if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                    { be_addDay = true; }
                    if (be_addDay || int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                    { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0).AddDays(1); }
                    else { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0); }
                    if (be_addDay)
                    { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0).AddDays(1); }
                    else
                    { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0); }
                    if (etime >= intime && intime >= stime2)
                    {
                        if (be_addDay)
                        { return new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0).AddDays(1); }
                        else { return new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0); }
                    }
                    else
                    {
                        if (stime2 > intime) //中間休息時間
                        {
                            return stime2;
                        }
                    }
                    if (be_addDay)
                    {
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0).AddDays(1);
                    }
                    else
                    {
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                    }
                    if (etime >= intime && intime >= stime2)
                    {
                        comp = dr["Shift_Morning"].ToString().Trim().Split(',');
                        if (be_addDay)
                        {
                            return new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0).AddDays(1);
                        }
                        else
                        {
                            return new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                        }
                    }
                    else
                    {
                        if (stime2 > intime) //中間休息時間
                        {
                            return stime2;
                        }
                    }
                }
                #endregion

            }
            return DateTime.Now;
        }

        public int TimeCompute2Seconds_BY_Calendar(DBADO db, string calendarName, DateTime start, DateTime end)
        {
            int cycleTime = 0;
            if (start >= end) { return cycleTime; }
            DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{calendarName}' and Holiday>='{start.ToString("yyyy-MM-dd")}' and Holiday<='{end.AddDays(1).ToString("yyyy-MM-dd")}' order by Holiday");
            DateTime intime = start;
            DateTime etime = DateTime.Now;
            DateTime stime2 = DateTime.Now;
            bool fristRUN = false;
            foreach (DataRow dr in dt.Rows)
            {
                if (Convert.ToDateTime(dr["Holiday"]).ToString("yyyy-MM-dd") != intime.ToString("yyyy-MM-dd"))
                {
                    if (Convert.ToDateTime(dr["Holiday"].ToString()) < intime) { break; }
                    else
                    {
                        stime2 = Convert.ToDateTime(dr["Holiday"]);
                        stime2 = new DateTime(stime2.Year, stime2.Month, stime2.Day, DateTime.Now.Hour, DateTime.Now.Minute, 0, 0);
                        if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                        { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                        else if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                        { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                        else if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                        { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                        else if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                        { intime = new DateTime(stime2.Year, stime2.Month, stime2.Day, int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), 0, 0); }
                        else
                        { break; }
                    }
                }
                #region Flag_Morning
                if (bool.Parse(dr["Flag_Morning"].ToString()) == true)
                {
                    string[] comp = dr["Shift_Morning"].ToString().Trim().Split(',');
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                    if (fristRUN || (etime >= intime && intime >= stime2))
                    {
                        bool isreturn = false;
                        if (etime >= end) { etime = end; isreturn = true; }
                        if (fristRUN)
                        { cycleTime += TimeCompute2Seconds(stime2, etime); }
                        else { cycleTime += TimeCompute2Seconds(intime, etime); fristRUN = true; }
                        if (isreturn) { return cycleTime; }
                    }
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                    if (fristRUN || (etime >= intime && intime >= stime2))
                    {
                        bool isreturn = false;
                        if (etime >= end) { etime = end; isreturn = true; }
                        if (fristRUN)
                        { cycleTime += TimeCompute2Seconds(stime2, etime); }
                        else { cycleTime += TimeCompute2Seconds(intime, etime); fristRUN = true; }
                        if (isreturn) { return cycleTime; }
                    }
                }
                #endregion

                #region Flag_Afternoon
                if (bool.Parse(dr["Flag_Afternoon"].ToString()) == true)
                {
                    string[] comp = dr["Shift_Afternoon"].ToString().Trim().Split(',');
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                    if (fristRUN || (etime >= intime && intime >= stime2))
                    {
                        bool isreturn = false;
                        if (etime >= end) { etime = end; isreturn = true; }
                        if (fristRUN)
                        { cycleTime += TimeCompute2Seconds(stime2, etime); }
                        else { cycleTime += TimeCompute2Seconds(intime, etime); fristRUN = true; }
                        if (isreturn) { return cycleTime; }
                    }
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                    if (fristRUN || (etime >= intime && intime >= stime2))
                    {
                        bool isreturn = false;
                        if (etime >= end) { etime = end; isreturn = true; }
                        if (fristRUN)
                        { cycleTime += TimeCompute2Seconds(stime2, etime); }
                        else { cycleTime += TimeCompute2Seconds(intime, etime); fristRUN = true; }
                        if (isreturn) { return cycleTime; }
                    }
                }
                #endregion

                #region Flag_Night
                if (bool.Parse(dr["Flag_Night"].ToString()) == true)
                {
                    string[] comp = dr["Shift_Night"].ToString().Trim().Split(',');
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                    if (fristRUN || (etime >= intime && intime >= stime2))
                    {
                        bool isreturn = false;
                        if (etime >= end) { etime = end; isreturn = true; }
                        if (fristRUN)
                        { cycleTime += TimeCompute2Seconds(stime2, etime); }
                        else { cycleTime += TimeCompute2Seconds(intime, etime); fristRUN = true; }
                        if (isreturn) { return cycleTime; }
                    }
                    if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                    { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1); }
                    else
                    {
                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                    }
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                    if (fristRUN || (etime >= intime && intime >= stime2))
                    {
                        bool isreturn = false;
                        if (etime >= end) { etime = end; isreturn = true; }
                        if (fristRUN)
                        { cycleTime += TimeCompute2Seconds(stime2, etime); }
                        else { cycleTime += TimeCompute2Seconds(intime, etime); fristRUN = true; }
                        if (isreturn) { return cycleTime; }
                    }
                }
                #endregion

                #region Flag_Graveyard
                if (bool.Parse(dr["Flag_Graveyard"].ToString()) == true)
                {
                    bool be_addDay = false;
                    string[] comp_Night = dr["Shift_Night"].ToString().Trim().Split(',');
                    string[] comp = dr["Shift_Graveyard"].ToString().Trim().Split(',');
                    if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                    { be_addDay = true; }
                    if (be_addDay || int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                    { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0).AddDays(1); }
                    else { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0, 0); }
                    if (be_addDay)
                    { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0).AddDays(1); }
                    else
                    { stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0); }
                    if (fristRUN || (etime >= intime && intime >= stime2))
                    {
                        bool isreturn = false;
                        if (etime >= end) { etime = end; isreturn = true; }
                        if (fristRUN)
                        { cycleTime += TimeCompute2Seconds(stime2, etime); }
                        else { cycleTime += TimeCompute2Seconds(intime, etime); fristRUN = true; }
                        if (isreturn) { return cycleTime; }
                    }
                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0);
                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0, 0);
                    if (fristRUN || (etime >= intime && intime >= stime2))
                    {
                        bool isreturn = false;
                        if (etime >= end) { etime = end; isreturn = true; }
                        if (fristRUN)
                        { cycleTime += TimeCompute2Seconds(stime2, etime); }
                        else { cycleTime += TimeCompute2Seconds(intime, etime); fristRUN = true; }
                        if (isreturn) { return cycleTime; }
                    }
                }
                #endregion

            }
            return cycleTime;
        }
        public DateTime GetYearOfWeekSunDay(int weekINT)//回傳第N周的星期日
        {
            DateTime firstDayInWeek = DateTime.Now;
            //CultureInfo info = CultureInfo.CurrentCulture;
            //int totalWeekOfYear = info.Calendar.GetWeekOfYear
            //(
            //    DateTime.Now,
            //    CalendarWeekRule.FirstDay,
            //    DayOfWeek.Sunday
            //);
            //totalWeekOfYear -= 1;

            //先取得該年第一天的日期
            DateTime firstDateOfYear = new DateTime(DateTime.Now.Year, 1, 1);
            //該年第一天再加上周數乘以七
            DateTime dayInWeek = firstDateOfYear.AddDays(weekINT * 7);
            firstDayInWeek = dayInWeek.Date;
            //ISO 8601所制定的標準中，一週的第一天為週一
            while (firstDayInWeek.DayOfWeek != DayOfWeek.Monday)
            {
                firstDayInWeek = firstDayInWeek.AddDays(-1);
            }

            return firstDayInWeek.AddDays(-1);


        }

        public TimeSpan GetCT(DBADO db, string CalendarName, DateTime Comintime, DateTime Comouttime)
        {

            //// 計算CT 要排除行事曆上的休息時間 
            ///  1. 取得製程名稱
            ///  2. 由製程table 取得 行事曆名稱
            ///  由行事曆table 取得行事曆
            ///  比較 intime / outtime 是否中間有休息時間
            //// 分成四班 中間有休息時間 劃分為八區

            // 
            DateTime[] myDateEnd = new DateTime[8];
            DateTime[] myDateStart = new DateTime[8];
            bool[] AreaEnable = new bool[8];
            int inttimeArea = 1;
            int outtimeArea = 1;

            DateTime Intime = Comintime;
            DateTime Outtime = Comouttime;

            DataRow drcalen = db.DB_GetFirstDataByDataRow($"select A.Flag_Morning,A.Flag_Afternoon,A.Flag_Night,A.Flag_Graveyard,A.Shift_Morning,A.Shift_Afternoon,A.Shift_Night,A.Shift_Graveyard from SoftNetSYSDB.[dbo].PP_HolidayCalendar as A where A.ServerId='{_Fun.Config.ServerId}' and A.CalendarName='{CalendarName}' and A.Holiday='{Intime.ToString("yyyy-MM-dd")}'");
            if (drcalen == null)
            {

                TimeSpan k = Outtime - Intime;
                if (k.TotalSeconds > 0)
                {
                    return k;
                }
                else
                {
                    return TimeSpan.Zero;
                }
            }

            GetCT_timeArea(drcalen, Intime, Outtime, ref myDateEnd, ref myDateStart, ref AreaEnable, ref inttimeArea, ref outtimeArea, Intime);

            TimeSpan interval = TimeSpan.Zero;
            DateTime logend = new DateTime(Intime.Year, Intime.Month, Intime.Day, 23, 59, 59);

            if (Intime.Day == Outtime.Day)  // 同一天
            {
                #region
                if (inttimeArea == outtimeArea)
                {
                    //// 原本流程
                    interval = Outtime - Intime;
                }
                else
                {
                    for (int i = inttimeArea; i <= 8; i++)
                    {
                        if (i == inttimeArea)
                        {
                            interval = myDateEnd[i - 1] - Intime;
                            logend = myDateEnd[i - 1];
                        }
                        else if (i < outtimeArea)
                        {
                            if (AreaEnable[(i - 1)])
                            {
                                interval += myDateEnd[(i - 1)] - myDateStart[(i - 1)];
                                logend = myDateEnd[i - 1];
                            }
                        }
                        else if (i == outtimeArea)
                        {
                            if (AreaEnable[(i - 1)])
                            {
                                if (myDateStart[(i - 1)] > Outtime)//休息時間
                                { interval += Outtime - myDateEnd[(i - 2)]; }
                                else
                                { interval += Outtime - myDateStart[(i - 1)]; }
                            }
                            else
                            {
                                interval += Outtime - logend;
                            }
                            break;
                        }
                    }
                }
                #endregion
            }
            else // 不同天
            {
                TimeSpan intervalDay = Outtime - Intime;
                bool cday = false;
                #region 算當天intime
                for (int i = inttimeArea; i <= 8; i++)
                {
                    if (i == inttimeArea && AreaEnable[i - 1])
                    {
                        interval = myDateEnd[i - 1] - Intime;
                        logend = myDateEnd[i - 1];
                        if (!cday) { cday = true; }
                    }
                    else
                    {
                        if (AreaEnable[i - 1])
                        {
                            interval += myDateEnd[i - 1] - myDateStart[i - 1];
                            logend = myDateEnd[i - 1];
                            if (!cday) { cday = true; }
                        }
                    }
                }
                if (!cday)
                { interval = logend - Intime; }
                #endregion

                #region 算每天
                for (int i = 1; i < intervalDay.Days; i++)
                {
                    drcalen = db.DB_GetFirstDataByDataRow($"select A.Flag_Morning,A.Flag_Afternoon,A.Flag_Night,A.Flag_Graveyard,A.Shift_Morning,A.Shift_Afternoon,A.Shift_Night,A.Shift_Graveyard from SoftNetSYSDB.[dbo].PP_HolidayCalendar as A where A.ServerId='{_Fun.Config.ServerId}' and A.CalendarName='{CalendarName}' and A.Holiday='{Intime.AddDays(i).ToString("yyyy-MM-dd")}'");
                    if (drcalen == null)
                    {
                        continue;
                    }
                    GetCT_timeArea(drcalen, Intime, Outtime, ref myDateEnd, ref myDateStart, ref AreaEnable, ref inttimeArea, ref outtimeArea, Intime.AddDays(i));
                    for (int k = 0; k < 8; k++)
                    {
                        if (AreaEnable[k])
                        {
                            interval += myDateEnd[k] - myDateStart[k];
                        }
                    }
                }
                #endregion

                #region 算當天 outtime
                logend = new DateTime(Outtime.Year, Outtime.Month, Outtime.Day, 00, 00, 01);
                cday = false;
                drcalen = db.DB_GetFirstDataByDataRow($"select A.Flag_Morning,A.Flag_Afternoon,A.Flag_Night,A.Flag_Graveyard,A.Shift_Morning,A.Shift_Afternoon,A.Shift_Night,A.Shift_Graveyard from SoftNetSYSDB.[dbo].PP_HolidayCalendar as A where A.ServerId='{_Fun.Config.ServerId}' and  A.CalendarName='{CalendarName}' and A.Holiday='{Outtime.ToString("yyyy-MM-dd")}'");
                if (drcalen != null)
                {
                    GetCT_timeArea(drcalen, Intime, Outtime, ref myDateEnd, ref myDateStart, ref AreaEnable, ref inttimeArea, ref outtimeArea, Outtime);
                    for (int i = 1; i <= outtimeArea; i++)
                    {
                        if (i == outtimeArea)
                        {
                            if (AreaEnable[(i - 1)])
                            {
                                if (Outtime > myDateStart[(i - 1)])
                                {
                                    interval += (Outtime - myDateStart[(i - 1)]);
                                    logend = myDateEnd[i - 1];
                                    if (!cday) { cday = true; }
                                }
                            }
                            else
                            {
                                if (cday) { interval += (Outtime - logend); }
                            }
                            break;
                        }
                        else
                        {
                            if (AreaEnable[(i - 1)])
                            {
                                interval += (myDateEnd[(i - 1)] - myDateStart[(i - 1)]);
                                logend = myDateEnd[i - 1];
                                if (!cday) { cday = true; }
                            }
                        }

                    }
                }
                if (!cday)
                { interval += (Outtime - logend); }
                #endregion
            }
            return interval;
            //if (interval.TotalSeconds > 0)
            //{
            //    return (int)interval.TotalSeconds;
            //}
            //else
            //{
            //    return 0;
            //}
        }
        private void GetCT_timeArea(DataRow drcalen, DateTime intime, DateTime outtime, ref DateTime[] myDateEnd, ref DateTime[] myDateStart, ref bool[] AreaEnable, ref int inttimeArea, ref int outtimeArea, DateTime compdate)
        {
            DateTime Intime = intime;
            DateTime Outtime = outtime;
            for (int i = 0; i < AreaEnable.Length; i++)
            {
                AreaEnable[i] = false;
            }
            inttimeArea = 8;
            outtimeArea = 8;
            #region myDateStart每段開始時間, myDateEnd每段結束時間
            // DateTime[] myDateEnd = new DateTime[8];
            myDateEnd[0] = compdate;// Convert.ToDateTime("1:1");
            myDateEnd[1] = compdate;
            myDateEnd[2] = compdate;
            myDateEnd[3] = compdate;
            myDateEnd[4] = compdate;
            myDateEnd[5] = compdate;
            myDateEnd[6] = compdate;
            myDateEnd[7] = compdate;


            // DateTime[] myDateStart = new DateTime[8];
            myDateStart[0] = compdate;
            myDateStart[1] = compdate;
            myDateStart[2] = compdate;
            myDateStart[3] = compdate;
            myDateStart[4] = compdate;
            myDateStart[5] = compdate;
            myDateStart[6] = compdate;
            myDateStart[7] = compdate;
            // bool[] AreaEnable = new bool[8];

            //int inttimeArea = 0;
            //int outtimeArea = 0;

            if (bool.Parse(drcalen["Flag_Morning"].ToString()) == true)
            {
                string[] Shife_Morning = drcalen["Shift_Morning"].ToString().Trim().Split(','); AreaEnable[0] = true;
                //myDateEnd[0] = myDateEnd[0].ParseExact(drcalen["Shift_Morning"].ToString().Trim().Split(',')[1], "HH:mm", System.Globalization.CultureInfo.InvariantCulture);
                //myDateStart[0] = DateTime.ParseExact(drcalen["Shift_Morning"].ToString().Trim().Split(',')[0], "HH:mm", System.Globalization.CultureInfo.InvariantCulture);
                myDateEnd[0] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(drcalen["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                myDateStart[0] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(drcalen["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                if (Shife_Morning.Length >= 4)
                {
                    //myDateEnd[1] = DateTime.ParseExact(drcalen["Shift_Morning"].ToString().Trim().Split(',')[3], "HH:mm", System.Globalization.CultureInfo.InvariantCulture);
                    //myDateStart[1] = DateTime.ParseExact(drcalen["Shift_Morning"].ToString().Trim().Split(',')[2], "HH:mm", System.Globalization.CultureInfo.InvariantCulture);
                    myDateEnd[1] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(drcalen["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                    myDateStart[1] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(drcalen["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);

                    AreaEnable[1] = true;
                }
            }

            if (bool.Parse(drcalen["Flag_Afternoon"].ToString()) == true)
            {
                string[] Shife_Morning = drcalen["Shift_Afternoon"].ToString().Trim().Split(','); AreaEnable[2] = true;
                myDateEnd[2] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(drcalen["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                myDateStart[2] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(drcalen["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                if (Shife_Morning.Length >= 4)
                {
                    myDateEnd[3] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(drcalen["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                    myDateStart[3] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(drcalen["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                    AreaEnable[3] = true;
                }
            }
            //###??? 可能有跨日的問題
            if (bool.Parse(drcalen["Flag_Night"].ToString()) == true)
            {
                string[] Shife_Morning = drcalen["Shift_Night"].ToString().Trim().Split(','); AreaEnable[4] = true;
                myDateEnd[4] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(drcalen["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                myDateStart[4] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(drcalen["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                if (Shife_Morning.Length >= 4)
                {
                    myDateEnd[5] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(drcalen["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                    myDateStart[5] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(drcalen["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                    AreaEnable[5] = true;
                }
            }
            //###??? 可能有跨日的問題
            if (bool.Parse(drcalen["Flag_Graveyard"].ToString()) == true)
            {
                string[] Shife_Morning = drcalen["Shift_Graveyard"].ToString().Trim().Split(','); AreaEnable[6] = true;
                myDateEnd[6] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Graveyard"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(drcalen["Shift_Graveyard"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                myDateStart[6] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(drcalen["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                if (Shife_Morning.Length >= 4)
                {
                    myDateEnd[7] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(drcalen["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                    myDateStart[7] = new DateTime(compdate.Year, compdate.Month, compdate.Day, int.Parse(drcalen["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(drcalen["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                    AreaEnable[7] = true;
                }
            }
            #endregion

            #region 判斷in, out在哪個時段
            if (AreaEnable[0] && myDateEnd[0] > Intime)
            {
                inttimeArea = 1;
            }
            else if (AreaEnable[1] && myDateEnd[1] > Intime)
            {
                inttimeArea = 2;
            }
            else if (AreaEnable[2] && myDateEnd[2] > Intime)
            {
                inttimeArea = 3;
            }
            else if (AreaEnable[3] && myDateEnd[3] > Intime)
            {
                inttimeArea = 4;
            }
            else if (AreaEnable[4] && myDateEnd[4] > Intime)
            {
                inttimeArea = 5;
            }
            else if (AreaEnable[5] && myDateEnd[5] > Intime)
            {
                inttimeArea = 6;
            }
            else if (AreaEnable[6] && myDateEnd[6] > Intime)
            {
                inttimeArea = 7;
            }
            else if (AreaEnable[7] && myDateEnd[7] > Intime)
            {
                inttimeArea = 8;
            }

            if (AreaEnable[0] && myDateEnd[0] > Outtime)
            {
                outtimeArea = 1;
            }
            else if (AreaEnable[1] && myDateEnd[1] > Outtime)
            {
                outtimeArea = 2;
            }
            else if (AreaEnable[2] && myDateEnd[2] > Outtime)
            {
                outtimeArea = 3;
            }
            else if (AreaEnable[3] && myDateEnd[3] > Outtime)
            {
                outtimeArea = 4;
            }
            else if (AreaEnable[4] && myDateEnd[4] > Outtime)
            {
                outtimeArea = 5;
            }
            else if (AreaEnable[5] && myDateEnd[5] > Outtime)
            {
                outtimeArea = 6;
            }
            else if (AreaEnable[6] && myDateEnd[6] > Outtime)
            {
                outtimeArea = 7;
            }
            else if (AreaEnable[7] && myDateEnd[7] > Outtime)
            {
                outtimeArea = 8;
            }
            #endregion 

        }


        public string ChangeStatus(string ipport, string keys, string programName, bool is_RUNTimeSTOP = false) //改變工站狀態   1=開始,2=停止,3=暫停,4=關站
        {
            //###???若此處改 RUNTimeServer.csㄝ要改
            string meg = "";
            string[] data = keys.Split(',');
            string status = data[data.Length - 1];
            string sql = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow dr = null;
                DataRow dr_M = null;
                for (int i = 0; i < (data.Length - 1); i++)
                {
                    dr_M = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");
                    #region 檢查 State合理性
                    if (dr_M["State"].ToString() == status) { continue; }
                    else if (dr_M["State"].ToString() == "2" && status == "3") { meg = $"{meg}<br>{data[i]} 無法設定暫停"; continue; }
                    else if (dr_M["State"].ToString() != "1" && status == "3") { meg = $"{meg}<br>{data[i]} 無法設定暫停"; continue; }
                    else if (status == "1" && (dr_M["OrderNO"].ToString() == "" || dr_M["IndexSN"].ToString() == "" || dr_M["OP_NO"].ToString() == "")) { meg = $"{meg}<br>{data[i]} 工站無設定,無法設定啟動"; continue; }
                    DataRow dr_WO = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_M["OrderNO"].ToString()}'");
                    if (dr_WO == null)
                    {
                        meg = $"{meg}<br>{data[i]} 查無工單,無法設定工站狀態"; continue;
                    }
                    if (meg == "")
                    {
                        if (dr_M["State"].ToString() == "1" && status == "1") { continue; }
                        if (dr_M["State"].ToString() == "2" && status == "2") { continue; }
                    }
                    #endregion
                    DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_M["SimulationId"].ToString()}'");
                    string for_Apply_StationNO_BY_Main_Source_StationNO = dr_APS_Simulation["Source_StationNO"].ToString();
                    string needID = "";

                    #region 關站處理
                    bool isLastStation = false;
                    if (status == "4")
                    {
                        #region 關站處理
                        dr = db.DB_GetFirstDataByDataRow($"select * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_M["OrderNO"].ToString()}'");//###??? and WOStatus=0
                        if (dr != null)
                        {
                            needID = dr.IsNull("NeedId") ? "" : dr["NeedId"].ToString();
                            dr = db.DB_GetFirstDataByDataRow($"select a.*,b.IsLastStation FROM SoftNetMainDB.[dbo].[Manufacture] as a, SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] as b where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO='{data[i]}' and a.OrderNO='{dr_M["OrderNO"].ToString()}' and a.IndexSN='{dr_M["IndexSN"].ToString()}' and a.StationNO=b.StationNO and a.IndexSN=b.IndexSN and a.OrderNO=b.OrderNO and a.PP_Name=b.PP_Name");
                            if (dr != null)
                            {
                                isLastStation = bool.Parse(dr["IsLastStation"].ToString());
                                List<string> tmp_list = new List<string>();
                                #region 判斷是否最後一站
                                if (isLastStation)
                                {
                                    /*
                                    SFC_Common SFC_FUN = new SFC_Common("1", _Fun.Config.Db);
                                    bool isRun_PP_ProductProcess_Item = true;
                                    if (db.DB_GetQueryCount($"SELECT * FROM SoftNetSYSDB.[dbo].PP_WO_Process_Item WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='" + dr_M["OrderNO"].ToString() + "'") > 0)
                                    { isRun_PP_ProductProcess_Item = false; }
                                    DataTable dt_WO_Stations = SFC_FUN.Process_ALLSation_RE_Custom(_Fun.Config.ServerId, "1", _Fun.Config.Db, dr_M["PP_Name"].ToString(), "ORDER BY IndexSN, PP_Name ASC", isRun_PP_ProductProcess_Item, dr_M["OrderNO"].ToString());
                                    if (dt_WO_Stations != null && dt_WO_Stations.Rows.Count > 0)
                                    {
                                        foreach (DataRow d in dt_WO_Stations.Rows)
                                        {
                                            tmp_list.Add($"'{d["Station NO"].ToString()}'");
                                        }

                                    }
                                    SFC_FUN.Dispose();
                                    */
                                    DataTable dt_WO_Stations = db.DB_GetData($"select Apply_StationNO from SoftNetSYSDB.[dbo].APS_Simulation where NeedId='{needID}' and PartSN>=0 group by Apply_StationNO");
                                    if (dt_WO_Stations != null && dt_WO_Stations.Rows.Count > 0)
                                    {
                                        foreach (DataRow d in dt_WO_Stations.Rows)
                                        {
                                            tmp_list.Add($"'{d["Apply_StationNO"].ToString()}'");
                                        }
                                    }
                                }
                                else
                                {
                                    DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].APS_Simulation where  NeedId='{needID}' and SimulationId='{dr_M["SimulationId"].ToString()}'");
                                    if (dr_tmp != null)
                                    { tmp_list.Add($"'{dr_tmp["Apply_StationNO"].ToString()}'"); }
                                }
                                #endregion
                                if (tmp_list.Count > 0)
                                {
                                    #region 半成品 or 成品 入庫 與 餘料入庫
                                    DataTable tmp_dt = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Apply_StationNO in ({string.Join(",", tmp_list)}) and Apply_PP_Name='{dr_M["PP_Name"].ToString()}' and (Class='4' or Class='5') and Source_StationNO is not null");
                                    if (tmp_dt != null)
                                    {
                                        foreach (DataRow d in tmp_dt.Rows)
                                        {
                                            DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and SimulationId='{d["SimulationId"].ToString()}'");
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
                                                    DataRow tmp = db.DB_GetFirstDataByDataRow($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id,b.KeepQTY FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}'");
                                                    if (tmp == null)
                                                    {
                                                        #region 查找適合庫儲別
                                                        SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "BC01");
                                                        #endregion
                                                        Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref tmp_no, "系統指派");
                                                    }
                                                    else
                                                    {
                                                        in_StoreNO = tmp["StoreNO"].ToString();
                                                        in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                                        Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref tmp_no, "系統指派");
                                                    }
                                                    sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{tmp_no}',Next_StoreQTY+={qty} where SimulationId='{d["SimulationId"].ToString()}'";
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
                                    tmp_list.Clear();
                                    DataTable dt_tmp = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needID}' and Apply_StationNO='{for_Apply_StationNO_BY_Main_Source_StationNO}' and Apply_PP_Name='{dr_M["PP_Name"].ToString()}' and ((Class!='4' and Class!='5') or Source_StationNO is null)");
                                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                    {
                                        foreach (DataRow d in dt_tmp.Rows)
                                        { tmp_list.Add($"'{d["SimulationId"].ToString()}'"); }
                                        sql = $"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needID}' and SimulationId in ({string.Join(",", tmp_list)})";
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
                                                DataTable tmp_dt = db.DB_GetData($@"SELECT a.*,c.SimulationDate FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                                                            join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.DOCNumberNO,1,4) and b.DOCType='3' and b.ServerId='{_Fun.Config.ServerId}'
                                                                                            join SoftNetSYSDB.[dbo].[APS_Simulation] as c on c.SimulationId=a.SimulationId
                                                                                            where a.SimulationId='{d["SimulationId"].ToString()}' order by OUT_StoreNO,OUT_StoreSpacesNO,IsOK");
                                                string docNumberNO = "";
                                                foreach (DataRow d2 in tmp_dt.Rows)
                                                {
                                                    if ((int.Parse(d2["QTY"].ToString()) - useQYU) > 0)
                                                    {
                                                        Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", useQYU, "", d2["Id"].ToString(), $"工單結束退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref docNumberNO, "系統指派");
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", int.Parse(d2["QTY"].ToString()), "", d2["Id"].ToString(), $"工單結束退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref docNumberNO, "系統指派");
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

                                db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsOK='1' where NeedId='{needID}'");
                            }
                        }
                        //db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET OrderNO='',OP_NO='',Master_PP_Name='',PP_Name='',IndexSN=0,Station_Custom_IndexSN='',StationNO_Custom_DisplayName='' where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");
                        #endregion
                    }
                    #endregion

                    #region 送Service處理
                    dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");
                    if (isLastStation)
                    {
                        status = "5";//關站加關工單
                    }
                    //db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[Manufacture_Log] ([Id],[logDate],[StationNO],[State],[OrderNO],[PartNO]) VALUES ('{_Str.NewId('C')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}','{data[i]}','{status}','{dr_M["OrderNO"].ToString()}','{dr_M["PartNO"].ToString()}')");
                    string type = "";
                    string tmp_date = "";
                    switch (status)
                    {
                        case "1":
                            tmp_date = $",RemarkTimeE=NULL";
                            if (dr_M.IsNull("RemarkTimeS")) { tmp_date = $"{tmp_date},RemarkTimeS='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'"; }
                            if (dr_M.IsNull("StartTime")) { tmp_date = $"{tmp_date},StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                            type = "智慧開工"; break;
                        case "2":
                            if (is_RUNTimeSTOP)
                            {
                                DataRow dr_stop = db.DB_GetFirstDataByDataRow($"SELECT top 1 * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday<='{DateTime.Now.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr["CalendarName"].ToString()}' order by CalendarName,Holiday desc");
                                if (dr_stop != null)
                                {
                                    DateTime logTime = DateTime.Now;
                                    if (bool.Parse(dr_stop["Flag_Morning"].ToString()))
                                    { logTime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_stop["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_stop["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), 0); }
                                    if (bool.Parse(dr_stop["Flag_Afternoon"].ToString()))
                                    { logTime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_stop["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_stop["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), 0); }
                                    if (bool.Parse(dr_stop["Flag_Night"].ToString()))
                                    { logTime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_stop["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_stop["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0); }
                                    if (bool.Parse(dr_stop["Flag_Graveyard"].ToString()))
                                    { logTime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_stop["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_stop["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[1]), 0); }
                                    dr_stop = db.DB_GetFirstDataByDataRow($"SELECT top 1 * FROM SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_M["OrderNO"].ToString()}' and IndexSN={dr_M["IndexSN"].ToString()} and LOGDateTime>'{logTime.ToString("MM/dd/yyyy HH:mm:ss")}' and LOGDateTime='{logTime.ToString("yyyy/MM/dd")}' and OrderNO='{dr_M["OrderNO"].ToString()}' and OperateType like '%報工%' and OperateType not like '%網頁報工%' order by LOGDateTime desc");
                                    if (dr_stop != null)
                                    {
                                        logTime = Convert.ToDateTime(dr_stop["LOGDateTime"]);
                                    }
                                    tmp_date = $",RemarkTimeE='{logTime.ToString("MM/dd/yyyy HH:mm:ss")}'";
                                }
                                else
                                {
                                    tmp_date = $",RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'";
                                }
                                type = "干涉停工";
                            }
                            else
                            {
                                tmp_date = $",RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'";
                                type = "智慧停工";
                            }
                            break;
                        case "4":
                        case "5":
                            tmp_date = $",EndTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'";
                            if (dr_M.IsNull("StartTime")) { tmp_date = $"{tmp_date},StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                            if (dr_M.IsNull("RemarkTimeS")) { tmp_date = $"{tmp_date},RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                            if (dr_M.IsNull("RemarkTimeE")) { tmp_date = $"{tmp_date},RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}'"; }
                            type = "智慧關站"; break;
                    }
                    db.DB_SetData(@$"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES 
                                    ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_APS_Simulation["NeedId"].ToString()}','{dr_APS_Simulation["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{programName}','{type}','{dr_M["PP_Name"].ToString()}','{data[i]}','{dr_M["PartNO"].ToString()}','{dr_M["OrderNO"].ToString()}','{dr_M["OP_NO"].ToString()}',{dr_M["IndexSN"].ToString()})");
                    db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET State='{status}'{tmp_date} where ServerId='{_Fun.Config.ServerId}' and StationNO='{data[i]}'");

                    if (SendRMSSocketINFO(1, $"WebChangeStationStatus,{status},{data[i]},WEBProg,{data[i]},{dr_M["OP_NO"].ToString()},{dr_M["OrderNO"].ToString()},{dr_M["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))


                    //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
       //             if (dr != null && RmsSend(dr["RMSName"].ToString(), 1, $"WebChangeStationStatus,{status},{data[i]},WEBProg,{data[i]},{dr_M["OP_NO"].ToString()},{dr_M["OrderNO"].ToString()},{dr_M["IndexSN"].ToString()},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
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
                    else
                    {
                        meg = $"{meg}<br>{data[i]} Service無作用,無法設定工站狀態";
                    }
                    #endregion
                }
            }
            return meg;
        }

    }

    public class SFCFunction : IDisposable
    {
        private bool disposed = false;

        private double FaultTolerance = 1.0;

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                disposed = true;
            }
        }

        public string GetSqlWhereInString(string semicolonString, char splitChar)
        {
            string text = "";
            string[] array = semicolonString.Split(splitChar);
            for (int i = 0; i < array.Length; i++)
            {
                if (i == array.Length - 1)
                {
                    text = text + "N'" + array[i] + "'";
                    break;
                }

                text = text + "N'" + array[i] + "',";
            }

            return text;
        }

        public void GetStringListOfDT(DataTable dt, string fieldName, ref List<string> list)
        {
            if (dt == null)
            {
                return;
            }

            foreach (DataRow row in dt.Rows)
            {
                list.Add(row[fieldName].ToString());
            }
        }

        public int GetWeekNO(DateTime dateTime)
        {
            Calendar calendar = new CultureInfo("zh-tw").Calendar;
            return calendar.GetWeekOfYear(dateTime, CalendarWeekRule.FirstDay, DayOfWeek.Sunday);
        }

        public string GetWhereInString_DT(DataTable dataTable, string fieldName)
        {
            string text = "";
            foreach (DataRow row in dataTable.Rows)
            {
                text = text + "N'" + row[fieldName].ToString() + "',";
            }

            return (text == "") ? "" : text.Substring(0, text.Length - 1);
        }

        public string GetWhereInString_List(List<string> stringList)
        {
            string text = "";
            foreach (string @string in stringList)
            {
                text = text + "N'" + @string + "',";
            }

            return (text == "") ? "" : text.Substring(0, text.Length - 1);
        }

        public double[] MergeAvgSD(double oAvg, double oSD, double oQty, double T, double TSD = 0.0, double TQty = 1.0)
        {
            double num = 0.0;
            double num2 = 0.0;
            double num3 = oQty + TQty;
            double num4 = (oQty * oAvg + TQty * T) / num3;
            num = Math.Sqrt((oQty * (Math.Pow(oAvg, 2.0) + Math.Pow(oSD, 2.0)) + TQty * (Math.Pow(T, 2.0) + Math.Pow(TSD, 2.0)) - num3 * Math.Pow(num4, 2.0)) / num3);
            if (num == double.NaN)
            {
                num = 0.0;
                num2 = 0.0;
            }
            else if (num != 0.0 && FaultTolerance != 0.0 && num4 != 0.0)
            {
                num2 = 3.0 * num * FaultTolerance / num4 * 100.0;
                if (num2 == double.NaN)
                {
                    num2 = 0.0;
                }
            }

            return new double[4] { num4, num, num3, num2 };
        }
    }
    public class SFCSqlFunction : IDisposable
    {
        private bool disposed = false;

        private DBADO _SoftNetSYSDB = null;

        private DBADO _tmLogDB = null;

        private string MasterDBType = "";

        private string MasterDBConnectString = "";

        public SFCFunction SFCFunction = new SFCFunction();

        private bool _isIncludeWT = true;

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                disposed = true;
            }
        }

        public SFCSqlFunction(DBADO SoftNetSYSDB, DBADO TMLogDB = null)
        {
            _SoftNetSYSDB = SoftNetSYSDB;
            _tmLogDB = TMLogDB;
        }

        public SFCSqlFunction(DBADO SoftNetSYSDB, string MasterDBType, string MasterDBConnectString)
        {
            _SoftNetSYSDB = SoftNetSYSDB;
            this.MasterDBType = MasterDBType;
            this.MasterDBConnectString = MasterDBConnectString;
        }




        public string GetWhereSmallerOrderStationNOIn(string wo, string processName, string stationNO)
        {
            List<string> stationNOList = new List<string>();
            DataTable dataTable = _SoftNetSYSDB.DB_GetData("SELECT [StationNO],[Sub_PP_Name] FROM SoftNetSYSDB.[dbo].[PP_WO_Process_Item] \r\n                    WHERE  OrderNO = '" + wo + "' and PP_Name = '" + processName + "' \r\n                    AND IndexSN <= (SELECT TOP 1 IndexSN FROM SoftNetSYSDB.[dbo].[PP_WO_Process_Item] \r\n                    WHERE OrderNO = '" + wo + "' and StationNO = '" + stationNO + "')");
            if (dataTable == null || dataTable.Rows.Count <= 0)
            {
                dataTable = _SoftNetSYSDB.DB_GetData("SELECT [StationNO],[Sub_PP_Name] FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] \r\n                    WHERE PP_Name = '" + processName + "' \r\n                    AND IndexSN <= (SELECT TOP 1 IndexSN FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] \r\n                    WHERE StationNO = '" + stationNO + "')");
            }

            foreach (DataRow row in dataTable.Rows)
            {
                if (row["StationNO"].ToString() != "")
                {
                    stationNOList.Add(row["StationNO"].ToString());
                }

                if (row["Sub_PP_Name"].ToString() != "")
                {
                    GetWhereSmallerOrderStationNOIn_(wo, row["Sub_PP_Name"].ToString(), ref stationNOList);
                }
            }

            return SFCFunction.GetWhereInString_List(stationNOList);
        }

        public void GetWhereSmallerOrderStationNOIn_(string wo, string processName, ref List<string> stationNOList)
        {
            DataTable dataTable = _SoftNetSYSDB.DB_GetData("SELECT [StationNO],[Sub_PP_Name] FROM SoftNetSYSDB.[dbo].[PP_WO_Process_Item] WHERE  OrderNO = '" + wo + "' and PP_Name = '" + processName + "'");
            if (dataTable == null || dataTable.Rows.Count <= 0)
            {
                dataTable = _SoftNetSYSDB.DB_GetData("SELECT [StationNO],[Sub_PP_Name] FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] WHERE and PP_Name = '" + processName + "'");
            }

            foreach (DataRow row in dataTable.Rows)
            {
                if (row["StationNO"].ToString() != "")
                {
                    stationNOList.Add(row["StationNO"].ToString());
                }

                if (row["Sub_PP_Name"].ToString() != "")
                {
                    GetWhereSmallerOrderStationNOIn_(wo, row["Sub_PP_Name"].ToString(), ref stationNOList);
                }
            }
        }

        private List<string> GetLastStationNOList(string processName)
        {
            return new List<string>();
        }

        public List<string> GetProcessNameList()
        {
            List<string> list = new List<string>();
            DataTable dataTable = _SoftNetSYSDB.DB_GetData("SELECT DISTINCT [PP_Name] FROM SoftNetSYSDB.[dbo].[PP_ProductProcess] WHERE PP_Name != ''");
            foreach (DataRow row in dataTable.Rows)
            {
                list.Add(row["PP_Name"].ToString());
            }

            return list;
        }

        public void GetProcessAllStationNOList(string processName, ref List<string> list)
        {
            if (processName == "")
            {
                DataTable dt = _SoftNetSYSDB.DB_GetData("SELECT DISTINCT [StationNO] FROM SoftNetSYSDB.[dbo].[PP_Station] WHERE StationNO!=''");
                SFCFunction.GetStringListOfDT(dt, "StationNO", ref list);
                return;
            }

            DataTable dt2 = _SoftNetSYSDB.DB_GetData("SELECT [StationNO] FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] WHERE PP_Name='" + processName + "' AND StationNO!=''");
            SFCFunction.GetStringListOfDT(dt2, "StationNO", ref list);
            DataTable dataTable = _SoftNetSYSDB.DB_GetData("SELECT [Sub_PP_Name] FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] WHERE PP_Name='" + processName + "' AND Sub_PP_Name!=''");
            if (dataTable == null)
            {
                return;
            }

            foreach (DataRow row in dataTable.Rows)
            {
                GetProcessAllStationNOList(row["Sub_PP_Name"].ToString(), ref list);
            }
        }

        public string GetStationName(string stationNO)
        {
            DataRow dataRow = _SoftNetSYSDB.DB_GetFirstDataByDataRow("SELECT StationName FROM SoftNetSYSDB.[dbo].[PP_Station] WHERE StationNO = '" + stationNO + "' ");
            return (dataRow == null) ? "" : dataRow["StationName"].ToString();
        }


        public DataRow GetStationStandardCT(string processName, string partNO, string stationNO, string orderNO)
        {
            DataRow dataRow = _SoftNetSYSDB.DB_GetFirstDataByDataRow($"SELECT Process.PP_Name,Station.PartNO,Station.StationNO,AVG(Station.E_CycleTime) E_CycleTime \r\n                                FROM\r\n                                (\r\n                                    SELECT [PartNO],[StationNO],[E_CycleTime] \r\n                                    FROM SoftNetSYSDB.[dbo].[PP_Station_RobotConfig]\r\n                                ) Station\r\n                                RIGHT JOIN\r\n                                (\r\n                                    SELECT [StationNO],[PP_Name] \r\n                                    FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item]\r\n                                ) Process\r\n                                ON Station.StationNO = Process.StationNO\r\n                                WHERE [PP_Name] = '{processName}' AND [PartNO] = '{partNO}' AND Station.[StationNO] = '{stationNO}'\r\n                                GROUP BY Process.PP_Name, Station.PartNO, Station.StationNO");
            if (dataRow == null)
            {
                dataRow = _SoftNetSYSDB.DB_GetFirstDataByDataRow($"SELECT [OrderNO],[PP_Name],[StationNO],[E_CycleTime] \r\n                                        FROM SoftNetSYSDB.[dbo].[PP_WO_Process_Item]\r\n                                        WHERE [OrderNO] = '{orderNO}' AND [PP_Name] = '{processName}' AND [StationNO] = '{stationNO}'");
                if (dataRow == null)
                {
                    dataRow = _SoftNetSYSDB.DB_GetFirstDataByDataRow($"SELECT [PP_Name],[StationNO],[E_CycleTime] \r\n                                        FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] Process\r\n                                        WHERE [PP_Name]='{processName}' AND [StationNO]='{stationNO}'");
                }
            }

            return dataRow;
        }

        public Tuple<float, int> GetStationWarningValue(string orderNO, string stationNO)
        {
            float item = 0f;
            int item2 = 0;
            DataRow dataRow = _SoftNetSYSDB.DB_GetFirstDataByDataRow("SELECT [PP_Name],[PartNO] FROM SoftNetSYSDB.[dbo].PP_WorkOrder WHERE OrderNO = '" + orderNO + "'");
            if (dataRow != null && !dataRow.IsNull("PartNO"))
            {
                DataRow dataRow2 = _SoftNetSYSDB.DB_GetFirstDataByDataRow("SELECT [BelowYield],[BelowCycleTime] FROM SoftNetSYSDB.[dbo].PP_Station_RobotConfig WHERE StationNO='" + stationNO + "' AND PartNO=N'" + dataRow["PartNO"].ToString() + "'");
                if (dataRow2 != null)
                {
                    if (!dataRow2.IsNull("BelowYield"))
                    {
                        item = float.Parse(dataRow2["BelowYield"].ToString());
                    }

                    if (!dataRow2.IsNull("BelowCycleTime"))
                    {
                        item2 = int.Parse(dataRow2["BelowCycleTime"].ToString());
                    }
                }
                else
                {
                    DataRow dataRow3 = _SoftNetSYSDB.DB_GetFirstDataByDataRow("SELECT [BelowYield],[BelowCycleTime] FROM SoftNetSYSDB.[dbo].[PP_WO_Process_Item] WHERE OrderNO='" + orderNO + "' and StationNO='" + stationNO + "' AND PP_Name='" + dataRow["PP_Name"].ToString() + "'");
                    if (dataRow3 == null)
                    {
                        dataRow3 = _SoftNetSYSDB.DB_GetFirstDataByDataRow("SELECT [BelowYield],[BelowCycleTime] FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] WHERE StationNO='" + stationNO + "' AND PP_Name='" + dataRow["PP_Name"].ToString() + "'");
                    }

                    if (dataRow3 != null)
                    {
                        if (!dataRow3.IsNull("BelowYield"))
                        {
                            item = float.Parse(dataRow3["BelowYield"].ToString());
                        }

                        if (!dataRow3.IsNull("BelowCycleTime"))
                        {
                            item2 = int.Parse(dataRow3["BelowCycleTime"].ToString());
                        }
                    }
                }
            }

            return Tuple.Create(item, item2);
        }

        public bool IsStationHasSNSource(string stationNO)
        {
            int num = _SoftNetSYSDB.DB_GetQueryCount("SELECT 1 FROM SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] WHERE StationNO = '" + stationNO + "' AND SerialNOKey!=0");
            return num > 0;
        }

        public bool IsStationTrackMode(string stationNO)
        {
            DataRow dataRow = _SoftNetSYSDB.DB_GetFirstDataByDataRow("SELECT StationUI_type FROM SoftNetSYSDB.[dbo].[PP_Station] WHERE StationNO = '" + stationNO + "' ");
            return dataRow == null || dataRow["StationUI_type"].ToString() == "1";
        }

        public void SFC_StoredProcedure(string indexSN, string stationNO, string wo, bool isTrack, string serverId)
        {
            //IL_0284: Unknown result type (might be due to invalid IL or missing references)
            try
            {
                DataRow dataRow = _tmLogDB.DB_GetFirstDataByDataRow("EXEC SoftNetLogDB.[dbo].WorkOrderSettlementUpdate '" + wo + "','" + stationNO + "'," + indexSN + ",'" + serverId + "'");
                if (dataRow == null)
                {
                    return;
                }

                float num = 0f;
                if (!dataRow.IsNull("_AVGCT"))
                {
                    num = float.Parse(dataRow["_AVGCT"].ToString());
                }

                DataRow dataRow2 = _SoftNetSYSDB.DB_GetFirstDataByDataRow("SELECT [StartTime] FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] WHERE OrderNO = N'" + wo + "' AND StationNO = N'" + stationNO + "' and IndexSN=" + indexSN);
                if (dataRow2 != null && !dataRow2.IsNull("StartTime") && dataRow2["StartTime"].ToString().Trim() != "")
                {
                    TimeSpan timeSpan = new TimeSpan(DateTime.Now.Ticks - Convert.ToDateTime(dataRow2["StartTime"]).Ticks);
                    if (timeSpan.TotalSeconds > 0.0)
                    {
                        _SoftNetSYSDB.DB_SetData(string.Format("UPDATE SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement SET \r\n                                AvarageCycleTime={0},CumulativeTime+={1},TotalCheckIn={2},TotalCheckOut={3},\r\n                                TotalInput={4},TotalOutput={5},TotalFail={6},TotalKeep={7},\r\n                                FPY={8},YieldRate={9},StationYieldRate={10},UpdateTime=N'{11}' \r\n                                WHERE ServerId='{15}' and OrderNO=N'{12}' AND StationNO=N'{13}' and IndexSN={14}", num.ToString(), timeSpan.Seconds, dataRow["_TotalCheckIn"].ToString(), dataRow["_TotalCheckOut"].ToString(), dataRow["_TotalInput"].ToString(), dataRow["_TotalOutput"].ToString(), dataRow["_TotalFail"].ToString(), dataRow["_TotalKeep"].ToString(), dataRow["_FPY"].ToString(), dataRow["_YieldRate"].ToString(), dataRow["_StationYieldRate"].ToString(), DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"), wo, stationNO, indexSN, serverId));
                    }
                }
            }
            catch (Exception exception)
            {
                string _s = "";
            }
        }

        public void CapacityIncreaseRateFunction(string wo)
        {
            DataRow dataRow = _SoftNetSYSDB.DB_GetFirstDataByDataRow("SELECT [OrderNO],[PP_Name],[PartNO],[CalendarName],[FirstInTime],[LastOutTime],[ActualQuantity] \r\n                                FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] WHERE [OrderNO] = N'" + wo + "'");
            if (dataRow == null || string.IsNullOrEmpty(dataRow["ActualQuantity"].ToString()))
            {
                return;
            }

            double num = CountCalendarWorkingTime(dataRow);
            double result;
            double t = ((double.TryParse(dataRow["ActualQuantity"].ToString(), out result) && result > 0.0) ? (num / double.Parse(dataRow["ActualQuantity"].ToString())) : 0.0);
            UpdateCapacityIncreaseRateTable(dataRow, t);
            DataTable dataTable = _SoftNetSYSDB.DB_GetData("SELECT [StationNO],[AvarageCycleTime],[AvarageWaitTime] FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] WHERE OrderNO='" + wo + "' AND (TotalOutput>0 OR TotalFail>0) AND StationNO != ''");
            if (dataTable == null || dataTable.Rows.Count <= 0)
            {
                return;
            }

            foreach (DataRow row in dataTable.Rows)
            {
                double num2 = double.Parse(row["AvarageCycleTime"].ToString());
                if (_isIncludeWT)
                {
                    num2 += double.Parse(row["AvarageWaitTime"].ToString());
                }

                UpdateCapacityIncreaseRateTable(dataRow, num2, row["StationNO"].ToString());
            }
        }

        private double CountCalendarWorkingTime(DataRow woInfo)
        {
            double num = 0.0;
            if (!string.IsNullOrEmpty(woInfo["FirstInTime"].ToString()) && !string.IsNullOrEmpty(woInfo["LastOutTime"].ToString()))
            {
                DateTime dateTime = DateTime.Parse(woInfo["FirstInTime"].ToString());
                DateTime dateTime2 = DateTime.Parse(woInfo["LastOutTime"].ToString());
                string text = dateTime.ToString("yyyy/MM/dd");
                string text2 = dateTime2.ToString("yyyy/MM/dd");
                string s = dateTime.ToString("HH:mm:ss");
                string s2 = dateTime2.ToString("HH:mm:ss");
                if (woInfo["CalendarName"].ToString() != "")
                {
                    DataTable dataTable = _SoftNetSYSDB.DB_GetData("SELECT [CalendarName],[Holiday],[Flag_Morning],[Flag_Afternoon],[Flag_Night],[Flag_Graveyard],[Shift_Morning],[Shift_Afternoon],[Shift_Night],[Shift_Graveyard] \r\n                                  FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] \r\n                                  WHERE CalendarName=N'" + woInfo["CalendarName"].ToString() + "' AND Holiday BETWEEN '" + text + "' AND '" + text2 + "'");
                    bool flag = false;
                    foreach (DataRow row in dataTable.Rows)
                    {
                        List<string> list = new List<string>();
                        if (bool.Parse(row["Flag_Morning"].ToString()))
                        {
                            string[] array = row["Shift_Morning"].ToString().Split(',');
                            list.Add(array[0] + "," + array[1]);
                            list.Add(array[2] + "," + array[3]);
                        }

                        if (bool.Parse(row["Flag_Afternoon"].ToString()))
                        {
                            string[] array2 = row["Shift_Afternoon"].ToString().Split(',');
                            list.Add(array2[0] + "," + array2[1]);
                            list.Add(array2[2] + "," + array2[3]);
                        }

                        if (bool.Parse(row["Flag_Night"].ToString()))
                        {
                            string[] array3 = row["Shift_Night"].ToString().Split(',');
                            list.Add(array3[0] + "," + array3[1]);
                            list.Add(array3[2] + "," + array3[3]);
                        }

                        if (bool.Parse(row["Flag_Graveyard"].ToString()))
                        {
                            string[] array4 = row["Shift_Graveyard"].ToString().Split(',');
                            list.Add(array4[0] + "," + array4[1]);
                            list.Add(array4[2] + "," + array4[3]);
                        }

                        list.Sort();
                        DateTime dateTime3 = DateTime.Now;
                        foreach (string item in list)
                        {
                            string[] array5 = item.Split(',');
                            if (!flag)
                            {
                                if (DateTime.Parse(s) >= DateTime.Parse(array5[0]) && DateTime.Parse(s) < DateTime.Parse(array5[1]))
                                {
                                    num += (DateTime.Parse(array5[1]) - DateTime.Parse(s)).TotalSeconds;
                                    dateTime3 = DateTime.Parse(array5[1]);
                                    flag = true;
                                }

                                if (DateTime.Parse(s) < DateTime.Parse(array5[0]))
                                {
                                    num += (DateTime.Parse(array5[1]) - DateTime.Parse(s)).TotalSeconds;
                                    dateTime3 = DateTime.Parse(array5[1]);
                                    flag = true;
                                }

                                continue;
                            }

                            if (DateTime.Parse(row["Holiday"].ToString()) == DateTime.Parse(text2))
                            {
                                if (DateTime.Parse(s2) >= DateTime.Parse(array5[0]) && DateTime.Parse(s2) < DateTime.Parse(array5[1]))
                                {
                                    num += (DateTime.Parse(s2) - DateTime.Parse(array5[0])).TotalSeconds;
                                    break;
                                }

                                if (DateTime.Parse(s2) < DateTime.Parse(array5[0]))
                                {
                                    if (dateTime3 < DateTime.Parse(s2))
                                    {
                                        num += (DateTime.Parse(s2) - dateTime3).TotalSeconds;
                                    }

                                    break;
                                }
                            }

                            num += (DateTime.Parse(array5[1]) - DateTime.Parse(array5[0])).TotalSeconds;
                            dateTime3 = DateTime.Parse(array5[1]);
                        }
                    }
                }
                else
                {
                    num = (dateTime2 - dateTime).TotalSeconds;
                }
            }

            return num;
        }

        private void UpdateCapacityIncreaseRateTable(DataRow woInfo, double T, string stationNO = "")
        {
            int weekNO = SFCFunction.GetWeekNO(DateTime.Now);
            bool flag = !(stationNO != "") || _isIncludeWT;
            DataRow dataRow = _SoftNetSYSDB.DB_GetFirstDataByDataRow(string.Format("SELECT [AvgUseSecond],[SD],[WOQty] FROM SoftNetSYSDB.[dbo].[PP_CapacityIncreaseRate]\r\n                   WHERE [Year] = {0} AND [WeekNO] = {1} \r\n                   AND [ProcessName] = N'{2}'\r\n                   AND [PartNO] = N'{3}'\r\n                   AND [StationNO] = N'{4}' \r\n                   AND [IsIncludeWT] ='{5}'", DateTime.Now.Year, weekNO, woInfo["PP_Name"].ToString(), woInfo["PartNO"].ToString(), stationNO, flag));
            if (dataRow != null)
            {
                double[] array = SFCFunction.MergeAvgSD(double.Parse(dataRow["AvgUseSecond"].ToString()), double.Parse(dataRow["SD"].ToString()), double.Parse(dataRow["WOQty"].ToString()), T);
                _SoftNetSYSDB.DB_SetData(string.Format("UPDATE SoftNetSYSDB.[dbo].[PP_CapacityIncreaseRate] \r\n                       SET [AvgUseSecond] = {0}, [SD] = {1}, [WOQty] = {2}, [IncreaseRate] = {3}\r\n                       WHERE [Year] = {4} AND [WeekNO] = {5} AND [ProcessName] = N'{6}' AND [PartNO] = N'{7}' AND [StationNO] = N'{8}' AND [IsIncludeWT]='{9}'", array[0], array[1], array[2], array[3], DateTime.Now.Year, weekNO, woInfo["PP_Name"].ToString(), woInfo["PartNO"].ToString(), stationNO, flag));
            }
            else
            {
                _SoftNetSYSDB.DB_SetData(string.Format("INSERT INTO SoftNetSYSDB.[dbo].[PP_CapacityIncreaseRate] \r\n                       ([Year],[WeekNO],[ProcessName],[StationNO],[PartNO],[AvgUseSecond],[IsIncludeWT],[SD],[WOQty],[IncreaseRate])\r\n                       VALUES ({0},{1},N'{2}',N'{3}',N'{4}',{5},'{6}',0,1,0)", DateTime.Now.Year, weekNO, woInfo["PP_Name"].ToString(), stationNO, woInfo["PartNO"].ToString(), T, flag));
            }
        }

        public Tuple<List<string>, int> GetWOCloseInfo(string wo, string serverId)
        {
            DataRow dataRow = _tmLogDB.DB_GetFirstDataByDataRow("SELECT MIN([InTime]) FirstInTime FROM SoftNetLogDB.[dbo].[SFC_StationDetail] WHERE ServerId='" + serverId + "' and OrderNO=N'" + wo + "' AND InTime IS NOT NULL");
            DataRow dataRow2 = _tmLogDB.DB_GetFirstDataByDataRow("SELECT MAX([OutTime]) LastOutTime FROM SoftNetLogDB.[dbo].[SFC_StationDetail] WHERE ServerId='" + serverId + "' and OrderNO=N'" + wo + "' AND OutTime IS NOT NULL");
            DataTable dataTable = _SoftNetSYSDB.DB_GetData("SELECT [StationNO] FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] WHERE ServerId='" + serverId + "' and OrderNO=N'" + wo + "' AND IsLastStation=1");
            DataRow dataRow3 = _tmLogDB.DB_GetFirstDataByDataRow("SELECT IIF(SUM([ProductFinishedQty]) IS NULL,0,SUM([ProductFinishedQty])) FinishQty FROM SoftNetLogDB.[dbo].[SFC_StationDetail] WHERE ServerId='" + serverId + "' and OrderNO=N'" + wo + "' AND IsDel=0 AND StationNO IN (" + ((dataTable == null || dataTable.Rows.Count == 0) ? "''" : (SFCFunction.GetWhereInString_DT(dataTable, "StationNO") ?? "")) + ")");
            DataRow dataRow4 = _tmLogDB.DB_GetFirstDataByDataRow("SELECT IIF(SUM([ProductFailedQty]) IS NULL,0,SUM([ProductFailedQty])) FailQty FROM SoftNetLogDB.[dbo].[SFC_StationDetail] WHERE ServerId='" + serverId + "' and OrderNO=N'" + wo + "' AND IsDel=0");
            int num = ((dataRow3 != null) ? ((!string.IsNullOrEmpty(dataRow3["FinishQty"].ToString())) ? int.Parse(dataRow3["FinishQty"].ToString()) : 0) : 0);
            int num2 = ((dataRow4 != null) ? ((!string.IsNullOrEmpty(dataRow4["FailQty"].ToString())) ? int.Parse(dataRow4["FailQty"].ToString()) : 0) : 0);
            int item = num + num2;
            return Tuple.Create(new List<string>
        {
            (dataRow != null && !string.IsNullOrEmpty(dataRow["FirstInTime"].ToString())) ? ("'" + DateTime.Parse(dataRow["FirstInTime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'") : "NULL",
            (dataRow2 != null && !string.IsNullOrEmpty(dataRow2["LastOutTime"].ToString())) ? ("'" + DateTime.Parse(dataRow2["LastOutTime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss") + "'") : "NULL"
        }, item);
        }
    }
    public class RunSimulation_Arg : IDisposable
    {
        public List<bool> ARGs = new List<bool>();

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private bool disposed = false;
        protected virtual void Dispose(bool disposing)
        {
            if (disposed) { return; }
            if (disposing)
            {
                //清理CLR託管資源
                ARGs.Clear();
                ARGs = null;
            }
            //清理非託管資源,寫在下方,如果有的話
            disposed = true;

        }
        public RunSimulation_Arg()
        {

        }
        public RunSimulation_Arg(string arg)
        {
            //ARGs.AddRange(arg.ToArray().ToList());
        }
    }




    public struct KeyAndValue
    {
        public string Key;
        public string Value;
        public KeyAndValue(string key, string value)
        {
            Key = key;
            Value = value;
        }
        public override string ToString()
        {
            return Value;
        }
    }


}
