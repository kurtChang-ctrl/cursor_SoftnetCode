//using AspNetCore;
using Base;
using Base.Services;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office.Word;
using Microsoft.Extensions.Hosting;
using Newtonsoft.Json;

using SoftNetWebII.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static StackExchange.Redis.Role;

namespace SoftNetWebII.Services
{
    public class RUNTimeServer : BackgroundService
    {
        #region 變數宣告
        public bool IsWork = false;
        private readonly HttpClient httpClient;//for電子標籤主機連線
        private bool IsWait = false;
        //SNWebSocketService WebSocketServiceOJB = null;
        bool SocketReceiveUserThreadPool = false;
        //private Dictionary<string, rmsConectUserData> _ToolUserList = new Dictionary<string, rmsConectUserData>(); //key=ip:port  value=rmsConectUserData(deviceName,ipcRobotName,socket)
        //private object lock__ToolUserList = new object();
        private object lock_logID = new object();
        private uint logID = 1;
        //private List<rmsMasterUserData> _MasterRMSUserList = new List<rmsMasterUserData>();
        private Dictionary<string, rmsMasterUserData> _MasterRMSUserList = new Dictionary<string, rmsMasterUserData>(); //key=ip:port  value=rmsConectUserData(deviceName,ipcRobotName,socket)
        private object lock__MasterRMSUserList = new object();
        private CancellationToken _stoppingToken = CancellationToken.None;

        private TcpListener _MastertcpListener = null;
        private bool _Mastertcplistenerstate = false;
        int SFCStatisticalAnalysisLoopTime = 60000;
        private bool ck5431OK = false;
        DBADO DBMaster = null;
        private SFC_Common _SFC_Common = new SFC_Common("1", _Fun.Config.Db);


        int RMSDBErrorCount = 0;

        #endregion
        private void MasterTcpListenerThread(CancellationToken cancellationToken = default)
        {
            Socket _client = null;
            try
            {
                //SoftNetService.Program._NLogMain.Write_Record(0, "", LogTitle.Null, LogSourceName.Null, "", "=======5431 log========== Start");
                _MastertcpListener.Start();
                ck5431OK = true;
                //SoftNetService.Program._NLogMain.Write_Record(0, "", LogTitle.Null, LogSourceName.Null, "MasterTcpListenerThread", "" + Program.RMSIP + "Service TcpListener Start");
                while (_Mastertcplistenerstate && !cancellationToken.IsCancellationRequested)
                {
                    if (_MastertcpListener != null && !_MastertcpListener.Pending())
                    {
                        try
                        {
                            Task.Delay(100, cancellationToken).Wait(cancellationToken);
                        }
                        catch (OperationCanceledException) { break; }
                    }
                    else
                    {
                        _client = _MastertcpListener.AcceptSocket();
                        if (_Mastertcplistenerstate)
                        {
                            Thread clientThread = new Thread(new ParameterizedThreadStart(MasterProcessRequest))
                            { IsBackground = true };
                            clientThread.Start(_client);
                        }
                    }
                }
               // SoftNetService.Program._NLogMain.Write_Record(0, "", LogTitle.Null, LogSourceName.Null, "MasterTcpListenerThread", "==== " + Program.RMSIP + "Exit : Service TcpListener status is " + _Mastertcplistenerstate);
                //SoftNetService.Program._NLogMain.Write_Record(0, "", LogTitle.Null, LogSourceName.Null, "", "=======5431 log========== where 離開 " + _Mastertcplistenerstate.ToString());
                if (_MastertcpListener != null)
                {
                    _MastertcpListener.Stop();
                }
            }
            catch (SocketException)
            {
                ++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_ExceptionError(1, "MasterTcpListenerThread", RMSError.Service_Exception, LogSourceName.Null, "", 0,
                //    ToolFun.StringAdd("Service已停止管理用TcpListener,請檢察網路,或重啟Service Engine. errorcode:", (ex as SocketException)?.ErrorCode.ToString(), " Exception:", (ex as SocketException)?.Message), ex);
            }
            catch (Exception)
            {
                ++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_ExceptionError(1, "MasterTcpListenerThread", RMSError.Service_Exception,
                //    LogSourceName.Null, "", 0, ToolFun.StringAdd("Service已停止管理用TcpListener,請檢察網路,或重啟Service Engine. Exception:", ex.Message), ex);
            }
            finally
            {
                _Mastertcplistenerstate = false;
                if (_MastertcpListener != null)
                {
                    _MastertcpListener.Stop();
                }
            }
            //SoftNetService.Program._NLogMain.Write_Record(0, "", LogTitle.Null, LogSourceName.Null, "MasterTcpListenerThread", "==== " + Program.RMSIP + " Exit : Service TcpListener status is " + _Mastertcplistenerstate);
        }
        
        // Async accept loop to reduce blocking threads and improve scalability
        private async Task MasterTcpListenerLoopAsync(CancellationToken cancellationToken = default)
        {
            Socket _client = null;
            try
            {
                _MastertcpListener.Start();
                ck5431OK = true;
                while (_Mastertcplistenerstate && !cancellationToken.IsCancellationRequested)
                {
                    try
                    {
                        if (_MastertcpListener != null && !_MastertcpListener.Pending())
                        {
                            await Task.Delay(100, cancellationToken);
                            continue;
                        }

                        _client = await _MastertcpListener.AcceptSocketAsync();
                        if (_Mastertcplistenerstate && _client != null)
                        {
                            _ = Task.Run(() => MasterProcessRequest(_client), cancellationToken);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                }
                if (_MastertcpListener != null)
                {
                    _MastertcpListener.Stop();
                }
            }
            catch (SocketException)
            {
                ++RMSDBErrorCount;
            }
            catch (Exception)
            {
                ++RMSDBErrorCount;
            }
            finally
            {
                _Mastertcplistenerstate = false;
                if (_MastertcpListener != null)
                {
                    _MastertcpListener.Stop();
                }
            }
        }
        public class ToolFun
        {
            public static byte[] KeepAlive(int onOff, int keepAliveTime, int keepAliveInterval)
            {
                byte[] array = new byte[12];
                BitConverter.GetBytes(onOff).CopyTo(array, 0);
                BitConverter.GetBytes(keepAliveTime).CopyTo(array, 4);
                BitConverter.GetBytes(keepAliveInterval).CopyTo(array, 8);
                return array;
            }
        }
        public struct NameSpaceDATA
        {
            public string NameSpace;

            public string Key;

            public int Number;

            public NameSpaceDATA(string ns, string key, int no)
            {
                NameSpace = ns;
                Key = key;
                Number = no;
            }

            public NameSpaceDATA(string ns, string key)
            {
                NameSpace = ns;
                Key = key;
                Number = 0;
            }

            public NameSpaceDATA(string ns, int no)
            {
                NameSpace = ns;
                Key = "";
                Number = no;
            }
        }
        public class PublicDATA
        {
            public string StationNO = "";

            public int WaitTimeFormula = 0;

            public int CycleTimeFormula = 0;

            public int CalculatePauseTime = 0;

            public DateTime PauseStartTime = default(DateTime);

            public string Station_Type = "1";

            public List<string> UserName = new List<string>();

            public bool IsPause = false;

            public bool IsError = false;

            public string IsErrorMEG = "";

            public string UIControlName = "";

            public string DashboradName = "";

            public string DashboradIPPort = "";

            public string DashboradIP = "";

            public List<string> IndexSN_Merge_StationLists = new List<string>();

            public string IndexSN_Merge_StationLists_IN_SQL = "";

            public Dictionary<string, List<string>> AllClassLists = new Dictionary<string, List<string>>();

            public char StationUI_Type = '1';

            public string NoDetail_OutQtyTagName = "";

            public string NoDetail_FailQtyTagName = "";

            public bool IsTagValueCumulative = true;

            public string WaitTagValueDIO = "Null";

            public bool IsWaitTargetFinish = false;

            public bool IsError_Last_Finishing = false;

            public TimeSpan PlayStartTime = new TimeSpan(DateTime.Now.Ticks);

            public TimeSpan LogStartTime = new TimeSpan(DateTime.Now.Ticks);

            public bool IsFristPlay = true;

            public bool IsUIControlCurrentMark = false;

            public int UIControlGotoMark = -1;

            public bool LockStatus = false;

            public List<string> OrderNO_List = new List<string>();

            public bool HasPPConfig = true;

            public int MAX_INCome = 20;

            public string ErrorType = "1";

            public bool IsThraedWork = false;

            public PublicDATA(string stationNO, string uIControlName, string dashboradName, List<string> userName)
            {
                StationNO = stationNO;
                UIControlName = uIControlName;
                DashboradName = dashboradName;
                foreach (string item in userName)
                {
                    if (item.Trim() != "")
                    {
                        UserName.Add(item);
                    }
                }
            }
        }
        public class ProductProcessMaster
        {
            public string OrderNO = "";

            public string ProcessName = "";

            public string WO_Parent_ProcessName = "";

            public int OrderQty = 0;

            public string PartNO = "";

            public string PartName = "";

            public string MasterSerialNO = "";

            public string Intime = "";

            public string Outtime = "";

            public TimeSpan TotalPauseTime = TimeSpan.Zero;

            public string FailMEG = "";

            public int CycleTime = 0;

            public int GotoMark = -10;

            public int CurrentMark = 0;

            public string CurrentPPName = "";

            public string DB_LOGDT = "";

            public char SerialNOKey = '0';

            public string SerialNOFrom_ThingName = "";

            public string SerialNOFrom_StationNO = "";

            public ushort SerialNOFrom_StationNO_Buffer = 1;

            public bool SerialNOFrom_StationNO_ByOrder = false;

            public string SerialNOFrom_CodingName = "";

            public string SerialNOFrom_SQL_ConnectString = "";

            public string SerialNOFrom_SQL_Language = "";

            public bool hasCheckOK = false;

            public string IsKeepData_LOGDateTime = "";

            public string IsKeepData_StationNO = "";

            public bool IsFree = true;

            public bool IsFinishFirstAction = false;

            public uint RunWIP_Index = 0u;

            public bool Is_has54 = false;

            public bool Is_has55 = false;

            public bool Is_has56 = false;

            public bool WaitTargetFinish_Flag = false;

            public bool Is_ActionLock = false;

            public string Link_57_StationNO = "";

            public string Link_57_LOGDateTime = "";

            public int WaitTime = 0;

            public bool IsForcedLeave = false;

            public string Station_Custom_IndexSN = "";

            public int IndexSN_Class = -1;

            public bool IsIndexSN_Merge = false;

            public ProductProcessMaster()
            {
            }

            public ProductProcessMaster(bool isFree, bool isFinishFirstAction, uint runWIP_Index)
            {
                IsFree = isFree;
                IsFinishFirstAction = isFinishFirstAction;
                RunWIP_Index = runWIP_Index;
            }

            public ProductProcessMaster(string OrderNO, string PartNO)
            {
                this.OrderNO = OrderNO;
                this.PartNO = PartNO;
            }
        }
        public class ProductProcessConfig : IEquatable<ProductProcessConfig>
        {
            private string stationNO = "";

            private int class_Process = 0;

            private string process_Type = "";

            private string operationName = "";

            private bool processEnable = true;

            private int loopTime = 1000;

            private List<string[]> conditionList = new List<string[]>();

            private int actionCount = 1;

            private int actionNOCount = 1;

            private int actionTotal = 1;

            private int conditionCount = 1;

            private int conditionCountTotal = 1;

            private List<string[]> actionList = new List<string[]>();

            private List<string[]> no_actionList = new List<string[]>();

            private List<string[]> dtact_Exception = new List<string[]>();

            private bool isSupportRMSData;

            private int gotoMark = -10;

            public int GotoMark
            {
                get
                {
                    return gotoMark;
                }
                set
                {
                    gotoMark = value;
                }
            }

            public bool IsSupportRMSData
            {
                get
                {
                    return isSupportRMSData;
                }
                set
                {
                    isSupportRMSData = value;
                }
            }

            public string StationNO
            {
                get
                {
                    return stationNO;
                }
                set
                {
                    stationNO = value;
                }
            }

            public int Class_Process
            {
                get
                {
                    return class_Process;
                }
                set
                {
                    class_Process = value;
                }
            }

            public string OperationName
            {
                get
                {
                    return operationName;
                }
                set
                {
                    operationName = value;
                }
            }

            public int LoopTime
            {
                get
                {
                    return loopTime;
                }
                set
                {
                    loopTime = value;
                }
            }

            public int ConditionCount
            {
                get
                {
                    return conditionCount;
                }
                set
                {
                    conditionCount = value;
                }
            }

            public int ActionCount
            {
                get
                {
                    return actionCount;
                }
                set
                {
                    actionCount = value;
                }
            }

            public int ActionNOCount
            {
                get
                {
                    return actionNOCount;
                }
                set
                {
                    actionNOCount = value;
                }
            }

            public int ActionTotal
            {
                get
                {
                    return actionTotal;
                }
                set
                {
                    actionTotal = value;
                }
            }

            public int ConditionCountTotal
            {
                get
                {
                    return conditionCountTotal;
                }
                set
                {
                    conditionCountTotal = value;
                }
            }

            public bool ProcessEnable
            {
                get
                {
                    return processEnable;
                }
                set
                {
                    processEnable = value;
                }
            }

            public List<string[]> ConditionList
            {
                get
                {
                    return conditionList;
                }
                set
                {
                    conditionList = value;
                }
            }

            public List<string[]> ActionList
            {
                get
                {
                    return actionList;
                }
                set
                {
                    actionList = value;
                }
            }

            public List<string[]> NO_ActionList
            {
                get
                {
                    return no_actionList;
                }
                set
                {
                    no_actionList = value;
                }
            }

            public List<string[]> Exception_ActionList
            {
                get
                {
                    return dtact_Exception;
                }
                set
                {
                    dtact_Exception = value;
                }
            }

            public ProductProcessConfig(DataRow sdr, DataTable conditionList, DataTable actionList, DataTable no_actionList, DataTable dtact_Exception, bool isSupportRMSData)
            {
                operationName = (sdr.IsNull("Process_Name") ? "" : sdr["Process_Name"].ToString());
                processEnable = bool.Parse(sdr["ProcessEnable"].ToString());
                stationNO = sdr["StationNO"].ToString();
                class_Process = int.Parse(sdr["Class_Process"].ToString());
                loopTime = int.Parse(sdr["Process_LoopTime"].ToString());
                actionCount = int.Parse(sdr["ActionCount"].ToString());
                if (actionCount <= 0)
                {
                    actionCount = 1;
                }

                actionNOCount = int.Parse(sdr["ActionNOCount"].ToString());
                if (actionNOCount <= 0)
                {
                    actionNOCount = 1;
                }

                this.isSupportRMSData = isSupportRMSData;
                gotoMark = -10;
                if (conditionList != null)
                {
                    foreach (DataRow row in conditionList.Rows)
                    {
                        this.conditionList.Add(new string[7]
                        {
                    row["ThingName"].ToString(),
                    row["Operation"].ToString(),
                    row["MatchValue"].ToString(),
                    row["Complex"].ToString(),
                    row["sn"].ToString(),
                    row["StationNO"].ToString(),
                    row["Class_Process"].ToString()
                        });
                    }
                }

                if (actionList != null)
                {
                    foreach (DataRow row2 in actionList.Rows)
                    {
                        this.actionList.Add(new string[11]
                        {
                    row2["ActionEnable"].ToString(),
                    row2["ActionType"].ToString(),
                    row2["SetText"].ToString(),
                    row2["SetValue"].ToString(),
                    row2["DeviceName"].ToString(),
                    row2["FCode"].ToString(),
                    row2["MailTO"].ToString(),
                    row2["ConvertValue"].ToString(),
                    row2["sn"].ToString(),
                    row2["Class_Process"].ToString(),
                    row2["StationNO"].ToString()
                        });
                    }
                }

                if (no_actionList != null)
                {
                    foreach (DataRow row3 in no_actionList.Rows)
                    {
                        this.no_actionList.Add(new string[11]
                        {
                    row3["ActionEnable"].ToString(),
                    row3["ActionType"].ToString(),
                    row3["SetText"].ToString(),
                    row3["SetValue"].ToString(),
                    row3["DeviceName"].ToString(),
                    row3["FCode"].ToString(),
                    row3["MailTO"].ToString(),
                    row3["ConvertValue"].ToString(),
                    row3["sn"].ToString(),
                    row3["Class_Process"].ToString(),
                    row3["StationNO"].ToString()
                        });
                    }
                }

                if (dtact_Exception == null)
                {
                    return;
                }

                foreach (DataRow row4 in dtact_Exception.Rows)
                {
                    this.dtact_Exception.Add(new string[11]
                    {
                row4["ActionEnable"].ToString(),
                row4["ActionType"].ToString(),
                row4["SetText"].ToString(),
                row4["SetValue"].ToString(),
                row4["DeviceName"].ToString(),
                row4["FCode"].ToString(),
                row4["MailTO"].ToString(),
                row4["ConvertValue"].ToString(),
                row4["sn"].ToString(),
                row4["Class_Process"].ToString(),
                row4["StationNO"].ToString()
                    });
                }
            }

            public bool Equals(ProductProcessConfig other)
            {
                if (stationNO == other.StationNO && class_Process == other.Class_Process)
                {
                    return true;
                }

                return false;
            }
        }
        public class PProcessValue : IDisposable
        {
            public PublicDATA _PublicDATA;

            public List<ProductProcessMaster> _ProductProcessMaster;

            public CancellationTokenSource _CancellationTokenSource;

            public List<ProductProcessConfig> _PPConfig;

            private uint index = 0u;

            private bool disposed = false;

            public PProcessValue(PublicDATA publicDATA, List<ProductProcessConfig> ppConfig, CancellationTokenSource cancellationTokenSource)
            {
                _PublicDATA = publicDATA;
                _PPConfig = ppConfig;
                _ProductProcessMaster = new List<ProductProcessMaster>();
                _CancellationTokenSource = cancellationTokenSource;
            }

            public PProcessValue(PublicDATA publicDATA, List<ProductProcessConfig> ppConfig, ProductProcessMaster productProcessMaster, CancellationTokenSource cancellationTokenSource)
            {
                _PublicDATA = publicDATA;
                _PPConfig = ppConfig;
                _ProductProcessMaster = new List<ProductProcessMaster>();
                _ProductProcessMaster.Add(productProcessMaster);
                _CancellationTokenSource = cancellationTokenSource;
            }

            public void Create_ProductProcessMaster(int count)
            {
                for (int i = 0; i < count; i++)
                {
                    _ProductProcessMaster.Add(new ProductProcessMaster());
                }
            }

            public void Create_PPM_SerialNO_DATA(DataRow dr)
            {
            }

            public int ReturnUSEIndex()
            {
                for (int i = 0; i < _ProductProcessMaster.Count; i++)
                {
                    if (_ProductProcessMaster[i].IsFree)
                    {
                        _ProductProcessMaster[i].IsFree = false;
                        _ProductProcessMaster[i].IsFinishFirstAction = false;
                        _ProductProcessMaster[i].RunWIP_Index = ++index;
                        _ProductProcessMaster[i].IndexSN_Class = -1;
                        _ProductProcessMaster[i].TotalPauseTime = TimeSpan.Zero;
                        return i;
                    }
                }

                return -1;
            }

            public int ReturnUSEIndex_by_NOWIP()
            {
                try
                {
                    for (int i = 0; i < _ProductProcessMaster.Count; i++)
                    {
                        if (_ProductProcessMaster[i].IsFree)
                        {
                            _ProductProcessMaster[i].IsFinishFirstAction = false;
                            _ProductProcessMaster[i].RunWIP_Index = ++index;
                            _ProductProcessMaster[i].DB_LOGDT = "";
                            _ProductProcessMaster[i].CycleTime = 0;
                            _ProductProcessMaster[i].FailMEG = "";
                            _ProductProcessMaster[i].Intime = "";
                            _ProductProcessMaster[i].Outtime = "";
                            _ProductProcessMaster[i].IsKeepData_LOGDateTime = "";
                            _ProductProcessMaster[i].IsKeepData_StationNO = "";
                            _ProductProcessMaster[i].hasCheckOK = false;
                            _ProductProcessMaster[i].GotoMark = -10;
                            _ProductProcessMaster[i].Is_ActionLock = false;
                            _ProductProcessMaster[i].Is_has54 = false;
                            _ProductProcessMaster[i].Is_has55 = false;
                            _ProductProcessMaster[i].Is_has56 = false;
                            _ProductProcessMaster[i].IsFinishFirstAction = false;
                            _ProductProcessMaster[i].OrderNO = "";
                            _ProductProcessMaster[i].ProcessName = "";
                            _ProductProcessMaster[i].WO_Parent_ProcessName = "";
                            _ProductProcessMaster[i].MasterSerialNO = "";
                            _ProductProcessMaster[i].IndexSN_Class = -1;
                            _ProductProcessMaster[i].IsFree = false;
                            _ProductProcessMaster[i].TotalPauseTime = TimeSpan.Zero;
                            return i;
                        }
                    }

                    _ProductProcessMaster.Add(new ProductProcessMaster(isFree: false, isFinishFirstAction: false, ++index));
                    return _ProductProcessMaster.Count - 1;
                }
                catch
                {
                    return -1;
                }
            }

            public List<ProductProcessMaster> Get_IsWork_PPM()
            {
                if (_ProductProcessMaster == null)
                {
                    return null;
                }

                return _ProductProcessMaster.FindAll((ProductProcessMaster x) => !x.IsFree);
            }

            public void Dispose()
            {
                Dispose(disposing: true);
                GC.SuppressFinalize(this);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (disposed)
                {
                    return;
                }

                if (disposing)
                {
                    _CancellationTokenSource.Dispose();
                    _CancellationTokenSource = null;
                    _PublicDATA.UserName.Clear();
                    _PublicDATA.IndexSN_Merge_StationLists.Clear();
                    _PublicDATA.AllClassLists.Clear();
                    _PublicDATA.UserName = null;
                    _PublicDATA.IndexSN_Merge_StationLists = null;
                    _PublicDATA.AllClassLists = null;
                    _PublicDATA = null;
                    _ProductProcessMaster.Clear();
                    _ProductProcessMaster = null;
                    if (_PPConfig != null)
                    {
                        _PPConfig.Clear();
                    }

                    _PPConfig = null;
                }

                disposed = true;
            }
        }


        private void MasterProcessRequest(object socket)
        {
            if (socket == null) { return; }
            object lock_receiveBuffer = new object();
            Socket _sockinfo = (Socket)socket;
            byte[] recvBytes = new byte[4096];
            List<byte> receiveBuffer = new List<byte>();
            System.Net.EndPoint oldEP;
            string ipport = "";
            int blen = 0;
            byte[] tra;
            byte type = 0;

            try
            {
                oldEP = _sockinfo.RemoteEndPoint;
                ipport = oldEP.ToString();
                _ = _sockinfo.IOControl(IOControlCode.KeepAliveValues, ToolFun.KeepAlive(1, 5000, 1000), null);
                while (_Mastertcplistenerstate)
                {
                    blen = _sockinfo.ReceiveFrom(recvBytes, ref oldEP);
                    if (blen == 0)
                    {
                        break;
                    }
                    else
                    {
                        //###???要寫當收到非RMS的封包,與攻擊時的code
                        tra = new byte[blen];
                        Array.Copy(recvBytes, tra, blen);

                        lock (lock_receiveBuffer)
                        {
                            blen = 0;
                            receiveBuffer.AddRange(tra.ToList());
                            if (receiveBuffer.Count >= 5)
                            {
                                byte[] tmpAC1 = new byte[] { receiveBuffer[0], receiveBuffer[1], receiveBuffer[2], receiveBuffer[3] };
                                int len = BitConverter.ToInt32(tmpAC1, 0);
                                if (receiveBuffer.Count >= len + 5)
                                {
                                    do
                                    {
                                        type = receiveBuffer[4];
                                        List<byte> a = receiveBuffer.GetRange(5, len);
                                        receiveBuffer.RemoveRange(0, len + 5);


                                        switch (type)
                                        {
                                            case 0://ping
                                                break;
                                            case 1://string
                                                string[] cmd = Encoding.UTF8.GetString(a.ToArray(), 0, a.Count).Split(',');
                                                if (SocketReceiveUserThreadPool)
                                                {
                                                    _ = SNThreadScheduler.Instance.StartAction(() => { Master_ResolveData2String(_sockinfo, ipport, cmd); });
                                                }
                                                else
                                                {
                                                    Task ttr21 = new Task(() =>
                                                    {
                                                        Master_ResolveData2String(_sockinfo, ipport, cmd);
                                                    });
                                                    ttr21.Start();
                                                }
                                                break;
                                            case 252:
                                                Task ttr252 = new Task(() =>
                                                {
                                                    Master_ResolveData2252(ipport, a.ToArray());
                                                });
                                                ttr252.Start();
                                                break;
                                            case 253://file
                                                Task ttr253 = new Task(() =>
                                                {
                                                    Master_ResolveData2253(ipport, a.ToArray());
                                                });
                                                ttr253.Start();
                                                break;
                                            case 254://RMSProtocol
                                                RMSProtocol pp = JsonConvert.DeserializeObject<RMSProtocol>(Encoding.UTF8.GetString(a.ToArray(), 0, a.Count));
                                                if (SocketReceiveUserThreadPool)
                                                { _ = SNThreadScheduler.Instance.StartAction(() => { Master_ResolveData2RMSProtocol(_sockinfo, ipport, pp); }); }
                                                else
                                                {
                                                    Task ttr = new Task(() =>
                                                    {
                                                        Master_ResolveData2RMSProtocol(_sockinfo, ipport, pp);
                                                    });
                                                    ttr.Start();
                                                }
                                                break;
                                            default:
                                                ++RMSDBErrorCount;
                                                //SoftNetService.Program._NLogMain.Write_RunError(1, ipport, "Master接收封包程序", RMSError.Service_Protocol_NoDefine, LogSourceName.DeviceName, Program.RMSName, 181, ToolFun.StringAdd("系統異常: Service收到未定義的Type,請聯絡TMM. type=", type.ToString(), "來源IP=", ipport));
                                                //_sockinfo.Close();
                                                //RunError A18
                                                break;
                                        }

                                        if (receiveBuffer.Count >= 5)
                                        {
                                            tmpAC1 = new byte[] { receiveBuffer[0], receiveBuffer[1], receiveBuffer[2], receiveBuffer[3] };
                                            len = BitConverter.ToInt32(tmpAC1, 0);
                                            if (receiveBuffer.Count >= len + 5)
                                            { continue; }
                                            else { break; }
                                        }
                                    } while (receiveBuffer.Count >= 5 && _Mastertcplistenerstate);
                                }
                            }
                        }
                    }
                }
            }
            catch (SocketException sex)
            {
                if (sex.ErrorCode != 10053)
                {
                    ++RMSDBErrorCount;
                    //SoftNetService.Program._NLogMain.Write_ExceptionError(1, "MasterProcessRequest", RMSError.Service_Exception, LogSourceName.IPPort, ipport, 0,
                    //    ToolFun.StringAdd("已關閉網路連線,請檢察網路,或重啟Service Engine. errorcode:", sex.ErrorCode.ToString(), " Exception:", sex.Message), sex);
                }
            }
            catch (Exception ex)
            {
                if (lock_receiveBuffer == null) { }
                else
                {
                    ++RMSDBErrorCount;
                    //SoftNetService.Program._NLogMain.Write_ExceptionError(1, "MasterProcessRequest", RMSError.Service_Exception, LogSourceName.IPPort,
                    //    ipport, 0, ToolFun.StringAdd("已關閉網路連線,請檢察網路,或重啟Service Engine. Exception:", ex.Message), ex);
                }
            }
            lock_receiveBuffer = null;
            if (_sockinfo != null)
            { _sockinfo.Close(); }

            if (lock__MasterRMSUserList != null)
            {
                lock (lock__MasterRMSUserList)
                {
                    try
                    {
                        if (_MasterRMSUserList != null && _MasterRMSUserList.ContainsKey(ipport))
                        {
                            /*
                            if (_LicenseListenerThread == true)
                            {
                                var clientBuilder = new Client()
                                {
                                    IP = ipport.Split(':')[0],
                                    ComputerName = "",
                                    DomainName = ""
                                };
                                foreach (KeyValuePair<string, rmsConectUserData> rudList in _ToolUserList)
                                {
                                    if (rudList.Key == ipport && rudList.Value.Role.ToString() == "15")
                                    {
                                        var giveResult = accessManager.RevokeAuthorization(clientBuilder, ClientType.Builder);
                                    }
                                }
                                SoftNetService.Program._NLogMain.Write_Record(0, "", LogTitle.Null, LogSourceName.Null, "", "" + clientBuilder.IP + "Remove Builder License OK");
                            }
                            */
                            _MasterRMSUserList[ipport].Dispose();
                            _ = _MasterRMSUserList.Remove(ipport);
                        }
                    }
                    catch (Exception ex)
                    {
                        //###???沒定義error
                        ++RMSDBErrorCount;
                        //SoftNetService.Program._NLogMain.Write_ExceptionError(1, "MonitorloopThread_Tick", RMSError.Service_Monitor_Fail,
                        //    LogSourceName.MonitorName, "", 61, ToolFun.StringAdd("沒定義error. Exception:", ex.Message), ex);
                    }
                    //ToolLicenseRemove
                }
            }
        }
        private Dictionary<NameSpaceDATA, PProcessValue> _PP_cts_Thread = new Dictionary<NameSpaceDATA, PProcessValue>();
        private object lock__PP_cts_Thread = new object();
        private void Master_ResolveData2String(Socket sender, string ipport, string[] cmd)
        {

            uint logid = 1;
            for (int i = 1; i < cmd.Length; i++)
            {
                cmd[i] = cmd[i].Replace("\x03", ",");
            }
            string ip = ipport.Split(':')[0].Trim();
            //string disIP = IpportConvertIPandName(ipport);

            try
            {

                lock (lock_logID)
                {
                    if (logid > 4290000000)
                    {
                        //SoftNetService.Program._NLogMain.NewOtherLogFile();
                        logID = 2;
                    }
                    logid = ++logID;
                }
                switch (cmd[0])//通知service要做什麼事
                {
                    case "IIS_Login":
                        if (cmd.Length >= 2)
                        {
                            //rmsMasterUserData mrul = _MasterRMSUserList..Find(delegate (rmsMasterUserData t) { return t.deviceName == cmd[1]; });
                            KeyValuePair<string, rmsMasterUserData> mrul = _MasterRMSUserList.FirstOrDefault(x => x.Value.AddrString == ipport);
                            if (mrul.Key == null || mrul.Key=="")
                            {
                                lock (lock__MasterRMSUserList)
                                {
                                    IPEndPoint oldEP = (IPEndPoint)sender.RemoteEndPoint;
                                    _MasterRMSUserList.Add(ipport,new rmsMasterUserData(cmd[1], sender, oldEP,25));
                                }
                            }
                        }
                        break;
                    case "LIB_TO_WEB":
                        if (cmd.Length >= 3)
                        {
                            _ = SendWebSocketClent_INFO(cmd);
                        }
                        break;
                    case "WebChangeStationStatus": //cmd[1]=stationNO cmd[2]=change值1=開始,2=停止,3=暫停,4=關站,5=關站+關閉工單
                        //###???將來10 = IsTagValueCumulative要改 tag Name
                        if (cmd.Length >= 5)
                        {
                            /*單工單
                            * *1.bnName, 2.StationNO, 3.obj.Name(WEBProg), 4._projectWithoutExtension(ipport), 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                            * //說明:IsTagValueCumulative==false需搭配WaitTagValueDIO==Null
                            */
                            NameSpaceDATA tmpns = new NameSpaceDATA(cmd[2], cmd[3]);
                            NameSpaceDATA ts_Status = new NameSpaceDATA(cmd[2], "");
                            switch (cmd[1])
                            {
                                case "1"://"CMD_Start";
                                    #region
                                    /* 原則:
                                     * 1.StationNO與PP_Name在資料庫為主Key, 但TMService是以StationNO為Key, 因為不同製程不會同時用相同站 , 同時不能有相同的StationNO run
                                     * 2.StationNO一定要有製程
                                     */
                                    {
                                        //先判斷有無開站
                                        /*
                                        try
                                        {

                                            if (_PP_cts_Thread.ContainsKey(tmpns))
                                            {
                                                if (_PP_cts_Thread[tmpns]._PublicDATA.HasPPConfig)
                                                {
                                                    lock (lock__PP_cts_Thread)
                                                    { _PP_cts_Thread[tmpns]._CancellationTokenSource.Cancel(); }
                                                }
                                                else
                                                {
                                                    lock (lock__PP_cts_Thread)
                                                    {
                                                        _PP_cts_Thread[tmpns].Dispose();
                                                        _PP_cts_Thread.Remove(tmpns);
                                                    }
                                                    ProductProcess_Thread_TickII(cmd[2]);
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "ResolveData2String", RMSError.Service_Exception, LogSourceName.StationNO, cmd[2], 61, ToolFun.StringAdd("Station Close時,發生Exception: ", ex.Message), ex);
                                            //ExceptionError A6
                                        }
                                        */
                                    }

                                    try
                                    {
                                        int indexSN_Class = -1;////本身製程設定中的階層
                                        KeyValuePair<NameSpaceDATA, PProcessValue> rud = _PP_cts_Thread.FirstOrDefault(x => x.Value._PublicDATA.StationNO == cmd[2]);
                                        if (rud.Key.NameSpace == null)
                                        {
                                            bool isRun_PP_ProductProcess_Item = true;//是否run途程
                                            string wo_parent_ProcessName = "";//紀錄工單上的製程 PS:(可能是主製程, station可能是子製程的)
                                            Dictionary<string, List<string>> allClassLists = new Dictionary<string, List<string>>();
                                            string tmpIndexSN = "";
                                            if (cmd[7].Trim() != "") { tmpIndexSN = $" and IndexSN={cmd[7]}"; }

                                            #region 檢查StationNO是否run途程
                                            DataRow dr_WO = DBMaster.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{cmd[6]}'");
                                            if (dr_WO == null)
                                            {
                                                //RmsSend(sender, true, new RMSProtocol(3511, JsonConvert.SerializeObject(new string[] { tmpns.NameSpace, "Err_Stop", ToolFun.StringAdd("製程明細中,查無".ToLocalizedStringByKey(LanguageLocalizer.Keys.SvcProcessDetailsHasNo), cmd[2], " 與 ", cmd[10]), "", "" }), false), logid);
                                                return;
                                            }
                                            string wo_partNO = dr_WO["PartNO"].ToString().Trim();
                                            wo_parent_ProcessName = dr_WO["PP_Name"].ToString().Trim();
                                            string wo_PartNO = dr_WO["PartNO"].ToString().Trim();
                                            string wo_qty = dr_WO["Quantity"].ToString().Trim();
                                            string wo_PartName = dr_WO["PartName"].ToString().Trim();
                                            DataRow dr_PP_ProductProcess_Item = DBMaster.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WO_Process_Item where ServerId='{_Fun.Config.ServerId}' and StationNO='{cmd[2]}' and OrderNO='{cmd[6]}'{tmpIndexSN}");
                                            if (dr_PP_ProductProcess_Item == null)
                                            {
                                                dr_PP_ProductProcess_Item = DBMaster.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_ProductProcess_Item where ServerId='{_Fun.Config.ServerId}' and StationNO='{cmd[2]}' and PP_Name='{wo_parent_ProcessName}'");
                                                if (dr_PP_ProductProcess_Item == null)
                                                {
                                                    //RmsSend(sender, true, new RMSProtocol(3511, JsonConvert.SerializeObject(new string[] { tmpns.NameSpace, "Err_Stop", ToolFun.StringAdd("製程明細中,查無".ToLocalizedStringByKey(LanguageLocalizer.Keys.SvcProcessDetailsHasNo), cmd[2], " 與 ", cmd[10]), "", "" }), false), logid);
                                                    return;
                                                }
                                            }
                                            else { isRun_PP_ProductProcess_Item = false; }
                                            #endregion



                                            //###???將來要 檢查是否有APS_Simulation計畫排程, 可參考ManufactureController的code
                                            #region 檢查是否有APS_Simulation計畫排程
                                            #endregion

                                            #region 若第一次開站,會回寫工單開始時間
                                            if (dr_WO.IsNull("StartTime") || dr_WO["StartTime"].ToString().Trim() == "")
                                            { _ = DBMaster.DB_SetData(string.Format("Update SoftNetSYSDB.[dbo].PP_WorkOrder SET StartTime='{0}' where OrderNO=N'{1}'", DateTime.Now.ToString("MM/dd/yyyy H:mm:ss"), cmd[6].Trim())); }
                                            #endregion

                                        }
                                        else
                                        {
                                        }





                                    }
                                    catch (Exception ex)
                                    {
                                        if (_PP_cts_Thread.ContainsKey(tmpns))
                                        {
                                            lock (lock__PP_cts_Thread)
                                            {
                                                _PP_cts_Thread[tmpns]._CancellationTokenSource.Cancel();
                                            }
                                        }
                                        //SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "ResolveData2String", RMSError.Service_Exception, LogSourceName.StationNO,
                                        //    cmd[2], 61, ToolFun.StringAdd("Station啟動發生Exception: ", ex.Message), ex);
                                    }
                                    #endregion
                                    break;
                                case "2"://"CMD_Stop"; 
                                    break;
                                case "3"://"CMD_Pause"; 
                                    break;
                                case "4"://"Close"; 
                                case "5"://"Close+關閉工單"; 
                                    #region 關閉工單
                                    if (cmd[1] == "5")
                                    { CloseWO(cmd[6]); }
                                    #endregion
                                    break;
                            }
                        }
                        break;
                    default:
                        ++RMSDBErrorCount;
//                        SoftNetService.Program._NLogMain.Write_RunError(logid, disIP, "解析String封包程序", RMSError.Service_Protocol_NoDefine, LogSourceName.IP, ip, 3014, ToolFun.StringAdd("系統異常:接收到訊號,但程式無定義,請聯絡TMM data=", string.Join(",", cmd)));

                        break;
                }
            }
            catch (Exception ex)
            {
                //++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "Master_ResolveData2String", RMSError.Service_Exception, LogSourceName.IPPort, ipport, 2012, ToolFun.StringAdd("通訊資料異常,請檢察網路或重啟Service Engine. data:", string.Join(",", cmd), " Exception:", ex.Message));
                //ExceptionError 72
                string _s = "";
            }

        }
        private void Master_ResolveData2252(string ip, byte[] data)
        {
            //byte[] tmpbyte410 = new byte[] { data[0], data[1], data[2], data[3] };
            //int fileNameLen410 = BitConverter.ToInt32(tmpbyte410, 0);
            //string ReciverFileSize_FileName = Encoding.Unicode.GetString(data, 4, fileNameLen410);
            //lock (lockdirectory)
            //{
            //    if (!Directory.Exists(Path.GetDirectoryName(ReciverFileSize_FileName)))
            //    {
            //        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(ReciverFileSize_FileName));
            //    }
            //}

            //BinaryWriter bWrite = new BinaryWriter(File.Open(ReciverFileSize_FileName, FileMode.Create, FileAccess.Write));
            //bWrite.Write(data, 4 + fileNameLen410, data.Length - 4 - fileNameLen410);
            //bWrite.Close();
            //bWrite.Dispose();
        }
        private void Master_ResolveData2253(string ip, byte[] data, uint logid = 1)
        {

        }
        private void Master_ResolveData2RMSProtocol(Socket sender, string ipport, RMSProtocol obj)
        {
            //###???要檢查來源的條件,如91501的if (_ToolUserList.ContainsKey(ipport)), 5430ㄝ要

            //Console.WriteLine("connect " + sender.RemoteEndPoint);
            uint logid = 1;
            if (obj.DataType != 11501 && obj.DataType != 11502 && obj.DataType != 11503 && obj.DataType < 91300)
            {
                lock (lock_logID)
                {
                    if (logid > 4290000000)
                    {
                        //SoftNetService.Program._NLogMain.NewOtherLogFile();
                        logID = 2;
                    }
                    logid = ++logID;
                }
                //SoftNetService.Program._NLogMain.Write_Record(logid, ipport, LogTitle.M_Receive, LogSourceName.ProtocolType, obj.DataType.ToString(), obj.Data);
            }
            //ThingsConfig tBuffer;
            string[] cmd = null;
            try
            {



            }
            catch (Exception ex)
            {
                //if (lock__RMSUserList == null) { return; }
                //++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "Master_ResolveData2RMSProtocol", RMSError.Service_Exception,
                //    LogSourceName.IPPort, ipport, 61, ToolFun.StringAdd("通訊資料異常,請重啟Service Engine或聯絡TMM. data:", obj.Data, " Exception:", ex.Message), ex);
                //ExceptionError A6
            }
        }
        private bool Master_Send_String(Socket user, string cmd)
        {
            EndPoint oldEP = null;
            //ushort rID = 65535;
            string ipport = "";
            try
            {
                oldEP = user.RemoteEndPoint;
                ipport = oldEP.ToString();
                byte[] data = Encoding.UTF8.GetBytes(cmd);
                byte[] byteSend = new byte[data.Length + 5];
                byte[] tmp2 = BitConverter.GetBytes(data.Length);
                tmp2.CopyTo(byteSend, 0);
                byteSend[4] = 1;
                data.CopyTo(byteSend, 5);
                _ = user.BeginSend(byteSend, 0, byteSend.Length, SocketFlags.None, new AsyncCallback(Rms_sendCallback), user);
            
            }
            catch (Exception ex)
            {
                ++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "Master_Send_RMSProtocol", RMSError.Service_Exception,
                //    LogSourceName.IPPort, ipport, 61, ToolFun.StringAdd("發送訊號失敗,請檢察網路連線或重啟Service Engine. data:", objdata.Data, " Exception:", ex.Message), ex);
                //ExceptionError A6
                return false;
            }
            return true;
        }
        private void Rms_sendCallback(IAsyncResult result)
        {
            string ipport = "";
            if (result.IsCompleted == false)
            {
                try
                {
                    Socket handler = (Socket)result.AsyncState;
                    ipport = handler.RemoteEndPoint.ToString();
                }
                catch { }
                ++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_ExceptionError(1, "rms_sendCallback", RMSError.Service_Exception, LogSourceName.IPPort, ipport, 0, "偵測到非同步Send失敗,請重啟Service Engine或聯絡TMM.", null);
                return;
            }
            try
            {
                Socket handler = (Socket)result.AsyncState;
                ipport = handler.RemoteEndPoint.ToString();
                int bytesSent = handler.EndSend(result);
            }
            catch (Exception ex)
            {
                ++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_ExceptionError(1, "rms_sendCallback", RMSError.Service_Exception, LogSourceName.IPPort, ipport, 0, ex.Message, ex);
            }

        }


        private void CloseWO(string wo)
        {
            SFCSqlFunction sFCSqlFunction = new SFCSqlFunction(DBMaster);
            Tuple<List<string>, int> woCloseInfo = sFCSqlFunction.GetWOCloseInfo(wo, _Fun.Config.ServerId);

            //###???hasFalseWOClose  CloseType=3沒做工單特結
            if (DBMaster.DB_SetData(
                $@"UPDATE SoftNetSYSDB.[dbo].[PP_WorkOrder] SET [EndTime]='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}',[CloseType]='2',TmpClose='1',[ActualQuantity] = {woCloseInfo.Item2}
                {(woCloseInfo.Item1[0] == "NULL" ? "" : $",[FirstInTime] = {woCloseInfo.Item1[0]}")}
                {(woCloseInfo.Item1[1] == "NULL" ? "" : $",[LastOutTime] = {woCloseInfo.Item1[1]}")}
                WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO=N'{wo}' 
                DELETE FROM SoftNetSYSDB.[dbo].PP_WorkOrder_Dispatch_Item WHERE OrderNO=N'{wo}'"))//2.正常關工單3.異常關工單
            {
                DataTable dataTable = DBMaster.DB_GetData(
                    $@"SELECT DISTINCT A.IndexSN,B.StationNO,B.IsTrack FROM
                        (
                            SELECT [StationNO],[IndexSN] FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO = N'{wo}'
                        ) A LEFT JOIN
                        (
                            SELECT ServerId,[StationNO],IIF([StationUI_type]='1',1,0) [IsTrack] FROM SoftNetSYSDB.[dbo].[PP_Station]
                        ) B ON A.StationNO=B.StationNO and B.ServerId='{_Fun.Config.ServerId}'");
                if (dataTable != null)
                {
                    foreach (DataRow item in dataTable.Rows)
                    {
                        sFCSqlFunction.SFC_StoredProcedure(item["IndexSN"].ToString(), item["StationNO"].ToString(), wo, item["IsTrack"].ToString() == "1", _Fun.Config.ServerId);
                    }
                }
                _ = DBMaster.DB_SetData($"EXEC SoftNetSYSDB.[dbo].WorkOrderProductionUpdate '{wo}','{_Fun.Config.ServerId}'");
                DataRow dr = DBMaster.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[PP_WorkOrder] where OrderNO=N'{wo}'");
                if (dr != null && dr["NeedId"].ToString() != "")
                {

                    //###??? 要考慮 APS_NeedData與APS_Simulation , 同一個NeedId可能有多個 工單時, 以下Code會有問題

                    _ = DBMaster.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='9',StateINFO=NULL where Id='{dr["NeedId"].ToString()}'");
                    _ = DBMaster.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{dr["NeedId"].ToString()}'");
                    _ = DBMaster.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{dr["NeedId"].ToString()}'");

                    #region 將計畫的 APS_PartNOTimeNote 報工合理性檢查
                    if (DBMaster.DB_GetQueryCount($"select Id from SoftNetSYSDB.[dbo].[APS_WarningData] where ServerId='{_Fun.Config.ServerId}' and NeedId='{dr["NeedId"].ToString()}' and DOCNumberNO='{wo}' and ErrorType='21'") <= 0)
                    {
                        DateTime day = DateTime.Now.AddDays(1);
                        DateTime time = new DateTime(DateTime.Now.Year, DateTime.Now.Month, day.Day, 0, 1, 1);
                        _ = DBMaster.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','21','{dr["NeedId"].ToString()}','','','{wo}','','','{time.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");

                    }
                    #endregion
                }
            }
            sFCSqlFunction.Dispose();


        }


        private void DBBase_OnException(string sql, string MEG)
        {
            ++RMSDBErrorCount;
            IsWork = false;
            //###??? 要建立停止運作的程序,例如:Dispose
            System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs停止運作, DB資料庫物件 ERROR : SQL={sql} INFO={MEG}", true);
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _stoppingToken = stoppingToken;
            int count = 0;
            string err_MEG = "";
            DBADO db = new DBADO("1", _Fun.Config.Db, ref err_MEG);
            if (err_MEG == "")
            {
                db.Error += new DBADO.ERROR(DBBase_OnException);
            }
            // 將 BackgroundService 的停止訊號傳遞給 SNWebSocketService
            try
            {
                var service = _Fun.DiBox?.GetService(typeof(SNWebSocketService));


                if (service == null)
                {
                    while (service == null && ++count <= 30)
                    {
                        await Task.Delay(1000, stoppingToken);
                        service = _Fun.DiBox?.GetService(typeof(SNWebSocketService));
                    }
                }
                if (service != null)
                {
                    var webSocketService = (SNWebSocketService)service;
                    webSocketService.SetCancellationToken(stoppingToken);
                }
                else
                {
                    IsWork = false;
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs ExecuteAsync 找不到 SNWebSocketService 停止運作", true);
                    _Mastertcplistenerstate = false;
                    return;
                }
            }
            catch { }

            #region 確認 Server 5431, IsWork, DBMaster OK
            count = 0;
            while (++count <= 10)
            {
                await Task.Delay(1000, stoppingToken);
                if (ck5431OK && IsWork && DBMaster != null)
                {
                    break;
                }
            }
            if (count > 10)
            {
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs ExecuteAsync Timeout 停止運作", true);
            }
            #endregion

            #region 初始化
            string logdate = "";

            //###???將來要改 SNWebSocketService ㄝ要改
            _efficientConfig.Clear();
            _efficientConfig.Add('A', new List<KeyAndValue>());
            _efficientConfig.Add('B', new List<KeyAndValue>());
            _efficientConfig.Add('C', new List<KeyAndValue>());
            _efficientConfig.Add('D', new List<KeyAndValue>());
            DataTable dt_Efficient = DBMaster.DB_GetData($"select * from SoftNetSYSDB.[dbo].PP_EfficientConfig where ServerId='{_Fun.Config.ServerId}' order by EfficiencyType,TypeKey");
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
            Stopwatch threadLoopTime = new Stopwatch();
            int next_time = 0;
            List<string> stations = new List<string>();

            bool isFirstRun = true;
            //falg_EfficientDetail_S = DateTime.Now;
            //falg_EfficientDetail_E = DateTime.Now;
            #endregion

            try
            {

                //SoftNetService.Program._NLogMain.Write_Record(0, "", LogTitle.Null, LogSourceName.Null, "", "=== SfcTimerloopthread_Tick start ===");
                while (IsWork && stoppingToken.CanBeCanceled)
                {
                    threadLoopTime.Restart();
                    string err = "", sql = "";

                    logdate = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff");

                        if (!IsWork) { break; }

                        #region 執行預存程序 WorkOrderProductionUpdate) 與 處理工單 TmpClose 未結
                        DataTable dt_PP_WorkOrder_StartWO = DBMaster.DB_GetData($@"SELECT * FROM SoftNetSYSDB.[dbo].PP_WorkOrder WHERE ServerId='{_Fun.Config.ServerId}' and StartTime IS NOT NULL AND (EndTime is null OR TmpClose IS NOT NULL)");
                        if (dt_PP_WorkOrder_StartWO != null && dt_PP_WorkOrder_StartWO.Rows.Count > 0)
                        {
                            for (int i = 0; i < dt_PP_WorkOrder_StartWO.Rows.Count; i++)
                            {
                                if (!IsWork) { break; }
                                DataRow dr_PP_WorkOrder_StartWO = dt_PP_WorkOrder_StartWO.Rows[i];
                                try
                                {
                                    #region 執行 WorkOrderProductionUpdate 預存程序
                                    err = "";
                                    _ = DBMaster.DB_SetData($"EXEC SoftNetSYSDB.[dbo].WorkOrderProductionUpdate '{dr_PP_WorkOrder_StartWO["OrderNO"].ToString()}','{_Fun.Config.ServerId}'", ref err);
                                    if (err != "")
                                    {
                                        ++RMSDBErrorCount;
                                        //Program._NLogMain.Write_RunError(logid, Program.RMSIP, "sfcTimerloopthread_Tick", RMSError.DBBase_Fail, LogSourceName.SQL,
                                        //    "EXEC SoftNetSYSDB.[dbo].WorkOrderProductionUpdate '" + dr_PP_WorkOrder_StartWO["OrderNO"].ToString() + "'", 2020, ToolFun.StringAdd("WorkOrderProductionUpdate預存程序執行失敗. ", err)); //RunError 15
                                    }
                                    #endregion

                                    #region TmpClose
                                    if (!dr_PP_WorkOrder_StartWO.IsNull("TmpClose"))
                                    {
                                        bool tmprun = true;
                                        if (tmprun)
                                        {
                                            #region 清除TotalStockII Keep的量
                                            if (!dr_PP_WorkOrder_StartWO.IsNull("NeedId") && dr_PP_WorkOrder_StartWO["NeedId"].ToString().Trim() != "")
                                            {
                                                _ = DBMaster.DB_SetData($"delete from SoftNetMainDB.[dbo].[TotalStockII] WHERE NeedId='{dr_PP_WorkOrder_StartWO["NeedId"].ToString()}'");
                                            }
                                            #endregion

                                            _ = DBMaster.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].PP_WorkOrder_Settlement SET EndTime='{logdate}' where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_PP_WorkOrder_StartWO["OrderNO"].ToString()}' and EndTime is null");
                                            _ = DBMaster.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].PP_WorkOrder SET TmpClose=null,EndTime='{logdate}' where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_PP_WorkOrder_StartWO["OrderNO"].ToString()}'");
                                        }
                                    }
                                    #endregion
                                }
                                catch (Exception ex)
                                {
                                    ++RMSDBErrorCount;
                                    //Program._NLogMain.Write_ExceptionError(logid, "sfcTimerloopthread_Tick", RMSError.Service_Exception, LogSourceName.Null, "", 2012,
                                    //    ToolFun.StringAdd("SoftNetLogDB.[dbo].WorkOrderSettlementUpdate預存程序,無法處理. ", ex.Message), ex);
                                    //ExceptionError A6
                                }
                            }
                        }
                        #endregion

                        if (!IsWork) { break; }

                        #region  AutoPauseWO
                        sql = $@"SELECT [OrderNO],[PP_Name],[PauseTime],PartNO 
                                FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] 
                                WHERE ServerId='{_Fun.Config.ServerId}' and IsAutoPauseWO='1' AND StartTime IS NOT NULL AND EndTime is null";
                        DataTable WOandProcessandPausetime = DBMaster.DB_GetData(sql);
                        if (WOandProcessandPausetime != null && WOandProcessandPausetime.Rows.Count > 0)
                        {
                            for (int i = 0; i < WOandProcessandPausetime.Rows.Count; i++)
                            {
                                if (!IsWork) { break; }
                                string WO = WOandProcessandPausetime.Rows[i][0].ToString(), Process = WOandProcessandPausetime.Rows[i][1].ToString(), PauseTime = WOandProcessandPausetime.Rows[i][2].ToString();
                                bool isRun_PP_ProductProcess_Item = true;
                                if (DBMaster.DB_GetQueryCount($"SELECT * FROM SoftNetSYSDB.[dbo].PP_WO_Process_Item WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='" + WO + "'") > 0)
                                { isRun_PP_ProductProcess_Item = false; }
                                DataTable allStationTable = _SFC_Common.Process_ALLSation_RE_Custom(_Fun.Config.ServerId, "1", _Fun.Config.Db, Process, WOandProcessandPausetime.Rows[i][3].ToString(), isRun_PP_ProductProcess_Item, WO);
                                if (allStationTable != null && allStationTable.Rows.Count > 0)
                                {
                                    bool isAllStationPause = true;
                                    foreach (DataRow item2 in allStationTable.Rows)
                                    {
                                        if (!IsWork) { break; }
                                        KeyValuePair<NameSpaceDATA, PProcessValue> init_List = _PP_cts_Thread.FirstOrDefault(x => x.Value._PublicDATA.StationNO == item2["Station NO"].ToString());
                                        if (init_List.Value != null && init_List.Value._ProductProcessMaster != null)
                                        {
                                            ProductProcessMaster ttmp = init_List.Value._ProductProcessMaster.Find(delegate (ProductProcessMaster t) { return t.OrderNO == WO; });
                                            if (ttmp != null && ttmp.OrderNO.Trim() != "")
                                            {
                                                if (!init_List.Value._PublicDATA.IsPause && !init_List.Value._PublicDATA.IsError && !init_List.Value._CancellationTokenSource.IsCancellationRequested)
                                                {
                                                    isAllStationPause = false;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    if (isAllStationPause && PauseTime == "")
                                    {
                                        sql = string.Format(
                                              @"UPDATE SoftNetSYSDB.[dbo].[PP_WorkOrder] 
                                                SET [PauseTime]='{0}' WHERE ServerId='{2}' and OrderNO='{1}'",
                                                logdate, WO, _Fun.Config.ServerId);
                                        _ = DBMaster.DB_GetData(sql);
                                    }
                                    if (!isAllStationPause && PauseTime != "")
                                    {
                                        sql = $@"SELECT [PauseTime],[AccumulatePauseHours] 
                                                FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] 
                                                WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='" + WO + "'";
                                        DataRow pauseTable = DBMaster.DB_GetFirstDataByDataRow(sql);
                                        if (!pauseTable.IsNull("PauseTime"))
                                        {
                                            DateTime pauseTime = Convert.ToDateTime(pauseTable[0].ToString());
                                            float accumulatePauseHours = float.Parse(pauseTable[1].ToString());
                                            float new_accumulatePauseHours = float.Parse(DateTime.Now.Subtract(pauseTime).TotalHours.ToString()) + accumulatePauseHours;
                                            sql = string.Format(
                                                  @"UPDATE SoftNetSYSDB.[dbo].[PP_WorkOrder] 
                                                    SET [PauseTime]=NULL,[AccumulatePauseHours]='{0}' where ServerId='{2}' and OrderNO='{1}'",
                                                    new_accumulatePauseHours, WO, _Fun.Config.ServerId);
                                            _ = DBMaster.DB_SetData(sql);
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                   
                    threadLoopTime.Stop();

                    next_time = SFCStatisticalAnalysisLoopTime - (int)threadLoopTime.ElapsedMilliseconds;

                    if (next_time < 20000)
                    { await Task.Delay(20000, stoppingToken); }
                    else
                    { await Task.Delay(next_time, stoppingToken); }
                    if (isFirstRun) { isFirstRun = false; }
                    next_time = SFCStatisticalAnalysisLoopTime + (int)threadLoopTime.ElapsedMilliseconds;

                }
            }
            catch (OperationCanceledException)
            {
                //正常關閉
                ++RMSDBErrorCount;
            }
            catch (Exception ex)
            {
                ++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "sfcTimerloopthread_Tick", RMSError.Service_Exception, LogSourceName.Null,
                //    "", 61, ToolFun.StringAdd("SFC定時程序失敗,導致無法使用定時程序功能,請檢查Service是否正常. Exception:", ex.Message), ex);
                //ExceptionError A6
            }
            db.Dispose();
            _Mastertcplistenerstate = false;
            return;
        }

        private bool SendWebSocketClent_INFO(string[] cmd)
        {
            KeyValuePair<string, rmsMasterUserData> mrul = _MasterRMSUserList.FirstOrDefault(x => x.Value.Role == 25);
            if (mrul.Key != null && mrul.Key != "" && mrul.Value.socket != null)
            {
                _ = Master_Send_String(mrul.Value.socket, string.Join(',', cmd));
            }
            return true;
        }
        private bool SendWebSocketClent_INFO(string meg)
        {
            KeyValuePair<string, rmsMasterUserData> mrul = _MasterRMSUserList.FirstOrDefault(x => x.Value.Role == 25);
            if (mrul.Key != null && mrul.Key != "" && mrul.Value.socket != null)
            {
                _ = Master_Send_String(mrul.Value.socket, meg);
            }
            return true;
        }
        
        
        public RUNTimeServer()
        {
            IsWork = true;
            string err_MEG = "";
            #region 建立 DBMaster 資料庫物件
            DBMaster = new DBADO("1", _Fun.Config.Db,ref err_MEG);
            if (err_MEG == "")
            {
                DBMaster.Error += new DBADO.ERROR(DBBase_OnException);
            }
            #endregion

            #region 建立與本機Service 連線管理用 Server 5431
            _Mastertcplistenerstate = true;
            _MastertcpListener = new TcpListener(IPAddress.Parse(_Fun.Config.MasterServiceIP), 5431);//###????暫時寫死
            // use async accept loop instead of dedicated blocking thread
            _ = Task.Run(() => MasterTcpListenerLoopAsync(_stoppingToken));
            #endregion

            httpClient = new HttpClient();
            #region 電子標籤Socket
            if (_Fun.Config.ElectronicTagsURL != "")
            {
                _Fun.Has_Tag_httpClient = true;
                httpClient.BaseAddress = new Uri($"http://{_Fun.Config.ElectronicTagsURL}");
                httpClient.Timeout = TimeSpan.FromSeconds(6);
            }
            #endregion

            #region 電子標籤主機網路監測
            _Fun.Is_RUNTimeServer_Thread_State[4] = true;
            Thread deviceConnectCheck = new Thread(() =>
            {
                Check_ElectronicTagsServer_Tick(_stoppingToken);
            });
            deviceConnectCheck.IsBackground = true;
            deviceConnectCheck.Start();
            #endregion

            Thread.Sleep(5000);

            #region 監控MES系統主程序
            _Fun.Is_RUNTimeServer_Thread_State[3] = true;
            Thread controlDATA_Othread2 = new Thread(() =>
            {
                SfcTimerloopthread_Tick(_stoppingToken);
            });
            controlDATA_Othread2.IsBackground = true;
            controlDATA_Othread2.Start();
            #endregion
        }
        private async void SfcTimerloopthread_Tick(CancellationToken cancellationToken = default)
        {
            string sql = "";
            DataTable dt_tmp = null;
            DataRow dr_tmp = null;
            string err_MEG = "";

            // 檢查是否要停止運行
            if (cancellationToken.IsCancellationRequested) { return; }

            DBADO db = new DBADO("1", _Fun.Config.Db, ref err_MEG);
            if (err_MEG == "")
            {
                db.Error += new DBADO.ERROR(DBBase_OnException);
            }
            #region 檢查 確認資料庫有建立委外加工站
            if (_Fun.Config.OutPackStationName != "")
            {
                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{_Fun.Config.OutPackStationName}'");
                if (dr_tmp == null)
                {
                    _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_Station] (ServerId,FactoryName,LineName,StationNO,StationName,RMSName,CalendarName) VALUES 
                                ('{_Fun.Config.ServerId}','','','{_Fun.Config.OutPackStationName}','{_Fun.Config.OutPackStationName}','','{_Fun.Config.DefaultCalendarName}')");
                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{_Fun.Config.OutPackStationName}'");
                    if (dr_tmp == null)
                    {
                        _ = db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[Manufacture] (ServerId,StationNO) VALUES 
                                ('{_Fun.Config.ServerId}','{_Fun.Config.OutPackStationName}')");
                    }
                }
            }
            #endregion

            #region 檢查 Manufacture 基本設定
            dt_tmp = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}'");
            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
            {
                string label_ProjectType = "0";
                string config_MutiWO = "0";
                if (_Fun.Config.RUNMode == '1') { label_ProjectType = "1"; }
                DataRow dr2 = null;
                foreach (DataRow dr in dt_tmp.Rows)
                {
                    if (dr["Station_Type"].ToString() == "8") { config_MutiWO = "1"; }
                    else { config_MutiWO = "0"; }
                    dr2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'");
                    if (dr2 == null)
                    { _ = db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Manufacture] ([StationNO],[ServerId],[Config_MutiWO],[Label_ProjectType]) VALUES ('{dr["StationNO"].ToString()}','{_Fun.Config.ServerId}','{config_MutiWO}','{label_ProjectType}')"); }
                    else { _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Config_MutiWO='{config_MutiWO}',Label_ProjectType='{label_ProjectType}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}'"); }
                }
            }
            #endregion

            #region 更新電子標籤內容 Thread 目前每1秒處理一次
            _Fun.Is_RUNTimeServer_Thread_State[0] = true;
            Thread SfcTimerloopUpdateTagValue = new Thread(() =>
            {
                SfcTimerloopUpdateTagValue_Tick(cancellationToken);
            });
            SfcTimerloopUpdateTagValue.IsBackground = true;
            SfcTimerloopUpdateTagValue.Start();
            #endregion

            #region 取消下段程式
            /*
            #region 等SNWebSocketServiceOK訊號
            int timeout = 120;
            while (_Fun.Has_Tag_httpClient && (!_Fun.Is_SNWebSocketService_OK || !IsUpdateTagValue_OK) && timeout >= 0)
            {
                timeout -= 1;
                await Task.Delay(1000, cancellationToken);
            }
            if (timeout < 0)
            {
                IsWork = false;
                _Fun.Is_RUNTimeServer_Thread_State[3] = false;
                db.Dispose();
                await _Log.ErrorAsync($"RUNTimeServer.cs SfcTimerloopthread_Tick 建立Is_SNWebSocketService_OK Timeout. RUNTimeServer已停止作業", true);
                return;
            }
            #endregion
            */
            #endregion

            await Task.Delay(2000, cancellationToken);
            //WebSocketServiceOJB = (SNWebSocketService)_Fun.DiBox.GetService(typeof(SNWebSocketService));

            #region 自動派工 與 工站累計量 Thread 目前每20秒處理一次
            _Fun.Is_RUNTimeServer_Thread_State[2] = true;
            Thread autoRUN_Json_thread = new Thread(() =>
            {
                SfcTimerloopautoRUN_Json_Tick(cancellationToken);
            });
            autoRUN_Json_thread.IsBackground = true;
            autoRUN_Json_thread.Start();
            #endregion

            #region 變數宣告
            int next_time = 0;
            string err = "";
            string logdate = "";
            bool isFirstRun = true;
            int isARGs10_offset = 15;//###??? 15將來改參數
            DataRow tmp = null;
            DataTable tmp_dt2 = null;
            DataTable tmp_dt = null;
            DataTable dt_APS_PartNOTimeNote = null;
            #endregion

            Stopwatch threadLoopTime = new Stopwatch();
            await Task.Delay(_Fun.Config.RunTimeServerLoopTime, cancellationToken);

            try
            {
                while (IsWork || !cancellationToken.IsCancellationRequested)
                {
                    logdate = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff");
                    threadLoopTime.Restart();

                    #region 以下為測試時暫時的code
                    /*
                    //db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStock] set QTY=0 where QTY>0");
                    //###???檢查TotalStock
                    //DataTable dt_test = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and APS_Default_StoreNO!=''");
                    //if (dt_test != null && dt_test.Rows.Count > 0)
                    //{
                    //    foreach (DataRow d in dt_test.Rows)
                    //    {
                    //        if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{d["APS_Default_StoreNO"].ToString()}' and PartNO='{d["PartNO"].ToString()}'") <= 0)
                    //        {
                    //            db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] ([Id],[ServerId],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{d["APS_Default_StoreNO"].ToString()}','{d["APS_Default_StoreSpacesNO"].ToString()}','{d["PartNO"].ToString()}',0)");
                    //        }
                    //    }
                    //}
                    */
                    #endregion

                    //lock (_Fun.Lock_Simulation_Flag)
                    //{
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        if (!_Str.Get_Simulation_flag())//判斷是否計算排程中
                        {
                            int tmp_int = 0;
                            string docNumberNO = "";

                            #region 判斷APS_Simulation已模擬完成卻未移轉, 將刪除模擬資料
                            //  State='3'被主動取消 =4足量轉計畫免生產 
                            string logdate_offset = DateTime.Now.AddMinutes(-isARGs10_offset).ToString("yyyy-MM-dd HH:mm:ss.fff");
                            string logdate_offsetADD = DateTime.Now.AddMinutes(isARGs10_offset).ToString("yyyy-MM-dd HH:mm:ss.fff");
                            sql = $@"SELECT * FROM SoftNetSYSDB.[dbo].[APS_NeedData] where State='2' and UpdateTime>'{logdate_offsetADD}' order by NeedDate";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                bool isRUN = false;
                                foreach (DataRow d in dt_tmp.Rows)
                                {
                                    isRUN = false;
                                    #region 讀排程設定參數 args
                                    List<string> cs = d["KeyA"].ToString().Split(',').ToList();
                                    RunSimulation_Arg args = new RunSimulation_Arg();
                                    foreach (string c in cs)
                                    {
                                        if (c == "0") { args.ARGs.Add(false); }
                                        else { args.ARGs.Add(true); }
                                    }
                                    #endregion
                                    if (args.ARGs[10])//將simulationtime的時間改為目前班別的起始時間,並不考慮第一階領料時間
                                    {
                                        if (!d.IsNull("UpdateTime"))
                                        {
                                            if (Convert.ToDateTime(d["UpdateTime"]).AddMinutes(isARGs10_offset) < DateTime.Now)
                                            {
                                                sql = $@"SELECT top 1 a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                                            where a.NeedId='{d["Id"].ToString()}' and a.StartDate<='{logdate}' and (a.Class='4' or Class='5') and a.Source_StationNO is not null and PartSN>=0 order by StartDate";
                                                tmp = db.DB_GetFirstDataByDataRow(sql);
                                                if (tmp != null) { isRUN = true; }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        sql = $@"SELECT top 1 a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                                    where a.NeedId='{d["Id"].ToString()}' and a.StartDate<='{logdate}' and PartSN>=0 order by StartDate";
                                        tmp = db.DB_GetFirstDataByDataRow(sql);
                                        if (tmp != null)
                                        {
                                            DateTime time = Convert.ToDateTime(tmp["StartDate"]);
                                            sql = $@"SELECT top 1 a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                                        where a.NeedId='{d["Id"].ToString()}' and (a.Class='4' or a.Class='5') and a.Source_StationNO is not null and PartSN>=0 order by StartDate";
                                            tmp = db.DB_GetFirstDataByDataRow(sql);
                                            if (tmp != null)
                                            {
                                                int s = _SFC_Common.TimeCompute2Seconds(time, Convert.ToDateTime(tmp["StartDate"]));
                                                if (s > 0) { s = s / 3; time = Convert.ToDateTime(tmp["StartDate"]).AddSeconds(-s); }
                                                if (time < DateTime.Now)
                                                {
                                                    isRUN = true;
                                                }
                                            }
                                        }
                                    }
                                    if (isRUN)
                                    {
                                        _ = db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{d["Id"].ToString()}'");
                                        _ = db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{d["Id"].ToString()}'");
                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{d["Id"].ToString()}'");
                                        _ = db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{d["Id"].ToString()}'");
                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] SET State='3',StateINFO=NULL,UpdateTime=null where Id='{d["Id"].ToString()}'");
                                        _ = db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{d["Id"].ToString()}'");
                                        _ = db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{d["Id"].ToString()}'");
                                        _ = db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] where Trigger_NeedId='{d["Id"].ToString()}'");
                                        _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId) VALUES ('排程需求模擬完成卻未移轉,已刪除模擬資料,需重新模擬或刪除它. 料號:{d["PartNO"].ToString()} 需求量:{d["NeedQTY"].ToString()} {d["NeedSource"].ToString()}','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d["Id"].ToString()}')");
                                    }
                                }
                            }
                            #endregion

                            #region 判斷APS_Simulation State='7'是否完成領料, 將改為State='9'
                            sql = $@"SELECT * FROM SoftNetSYSDB.[dbo].[APS_NeedData] where State='7'";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                foreach (DataRow d in dt_tmp.Rows)
                                {
                                    sql = $@"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{d["Id"].ToString()}' and StartDate<='{logdate}' and DOCNumberNO!='' and IsOK='0'";
                                    tmp_dt2 = db.DB_GetData(sql);
                                    if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                    {
                                        foreach (DataRow d2 in tmp_dt2.Rows)
                                        {
                                            sql = $"SELECT sum(QTY) as qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d2["SimulationId"].ToString()}' and IsOK='1' and IN_StoreNO=''";
                                            tmp = db.DB_GetFirstDataByDataRow(sql);
                                            if (tmp != null && !tmp.IsNull("qty") && tmp["qty"].ToString().Trim() != "")
                                            {
                                                if (int.Parse(tmp["qty"].ToString()) >= int.Parse(d2["NeedQTY"].ToString()))
                                                { _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set IsOK='1' where SimulationId='{d2["SimulationId"].ToString()}'"); }
                                            }
                                        }
                                    }
                                    if (db.DB_GetQueryCount($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{d["Id"].ToString()}' and IsOK='0'") <= 0)
                                    {
                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] set State='9',StateINFO=NULL where Id='{d["Id"].ToString()}'");
                                        _ = db.DB_SetData($"delete from SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{d["Id"].ToString()}'");
                                        _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId) VALUES ('排程需求領料完成,已自動結案. 料號:{d["PartNO"].ToString()} 需求量:{d["NeedQTY"].ToString()} {d["NeedSource"].ToString()}','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d["Id"].ToString()}')");

                                    }
                                }
                            }
                            #endregion

                            #region 排程採購 與 排程委外加工 與 安全存量 自動開立, 當 Math_TotalStock_HasUseQTY + Math_Online_SurplusQTY 小於 NeedQTY + SafeQTY
                            sql = $@"SELECT a.*,c.IS_WorkingPaper,c.IS_Store_Test FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                    join SoftNetSYSDB.[dbo].[APS_NeedData] as b on b.Id=a.NeedId and b.State='6'
                                    join SoftNetMainDB.[dbo].[Material] as c on c.PartNO=a.PartNO and c.IS_WorkingPaper='1' and c.ServerId='{_Fun.Config.ServerId}'
                                    where a.IsWPaper='0' and a.StartDate<='{logdate_offset}' and (a.DOCNumberNO is null or a.DOCNumberNO='') and (a.NeedQTY+a.SafeQTY)>(a.Math_TotalStock_HasUseQTY+a.Math_Online_SurplusQTY) order by a.StartDate,a.NeedId,a.PartNO";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                int tmp_i = 0;
                                #region 採購需求
                                string in_NO = "AA02";//###???暫時寫死
                                string mFNO = "";
                                docNumberNO = "";
                                string in_StoreNO = "";
                                string in_StoreSpacesNO = "";
                                float price = 0;
                                DataRow tmp_02 = null;
                                foreach (DataRow dr in dt_tmp.Rows)
                                {
                                    if (db.DB_GetQueryCount($"select * from SoftNetMainDB.[dbo].[DOC1BuyII] where SimulationId='{dr["SimulationId"].ToString()}'") <= 0)
                                    {
                                        if (dr["Class"].ToString() != "4" && dr["Class"].ToString() != "5")
                                        {
                                            mFNO = "";
                                            tmp_i = (int.Parse(dr["NeedQTY"].ToString()) + int.Parse(dr["SafeQTY"].ToString())) - (int.Parse(dr["Math_TotalStock_HasUseQTY"].ToString()) + int.Parse(dr["Math_Online_SurplusQTY"].ToString()));
                                            if (tmp_i > 0)
                                            {
                                                tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr["NeedId"].ToString()}' and PartNO='{dr["Master_PartNO"].ToString()}' and Source_StationNO='{dr["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN='{dr["IndexSN"].ToString()}' and PartSN<{dr["PartSN"].ToString()} and (Class='4' or Class='5') order by PartSN desc");
                                                if (tmp == null) { continue; }
                                                #region 查找適合廠商
                                                mFNO = _SFC_Common.SelectDOC1BuyMFNO(db, dr["PartNO"].ToString(), dr["SimulationId"].ToString(), "", ref price);
                                                #endregion

                                                #region 查找適合入庫儲別
                                                _SFC_Common.SelectINStore(db, dr["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, in_NO);
                                                #endregion

                                                if (_SFC_Common.Create_DOC1stock(db, dr, mFNO, price, in_StoreNO, in_StoreSpacesNO, in_NO, tmp_i, "", "", "排程物料需求不足", logdate, Convert.ToDateTime(dr["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "系統指派", ref docNumberNO))
                                                {
                                                    tmp_02 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where PartNO='{dr["PartNO"].ToString()}' and WorkType='1' and SimulationId='{dr["SimulationId"].ToString()}' and MFNO='{mFNO}' and DOCNumberNO='{docNumberNO}'");
                                                    if (tmp_02 == null)
                                                    {
                                                        sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkingPaper] (ServerId,[Id],[WorkType],[PartNO],[Class],[IsOK],[NeedId],[SimulationId],[UP_SimulationId],[Down_SimulationId],[NeedQTY],[Price],[Unit],[MFNO],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],[APS_StationNO],[APS_StationNO_SID],[StartTime],[ArrivalDate],[EndTime],[UpdateTime],DOCNumberNO)
                                                    VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('P')}','1','{dr["PartNO"].ToString()}','{dr["Class"].ToString()}','0','{dr["NeedId"].ToString()}','{dr["SimulationId"].ToString()}','','{tmp["SimulationId"].ToString()}',{tmp_i},{price},'PCS','{mFNO}','{in_StoreNO}','{in_StoreSpacesNO}','','',
                                                    '{tmp["Source_StationNO"].ToString()}','{tmp["SimulationId"].ToString()}','{logdate}','{Convert.ToDateTime(dr["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}',NULL,'{logdate}','{docNumberNO}')";
                                                        _ = db.DB_SetData(sql);
                                                        if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsWPaper='1' where SimulationId='{dr["SimulationId"].ToString()}'"))
                                                        {
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion

                                #region 委外加工
                                in_NO = "PA02";//###???暫時寫死
                                DataRow tmp_down = null;
                                DataRow tmp_up = null;
                                string tmp_down_SID = "NULL";
                                string tmp_down_Source_StationNO = "NULL";
                                string tmp_down_StartTime = "";
                                string tmp_down_ArrivalDate = "";
                                docNumberNO = "";
                                foreach (DataRow dr in dt_tmp.Rows)
                                {
                                    if (_Fun.Config.OutPackStationName == dr["Source_StationNO"].ToString())
                                    {
                                        mFNO = "";
                                        tmp_down_SID = "NULL";
                                        tmp_down_Source_StationNO = "NULL";

                                        tmp_down_StartTime = Convert.ToDateTime(dr["StartDate"]).ToString("yyyy-MM-dd HH:mm:ss.fff");

                                        //用ArrivalDate網前計算StartTime
                                        //DataRow dr_tmp = db.DB_GetFirstDataByDataRow($"select ROUND(sum(EfficientCycleTime)*sum(CountQTY)/sum(CountQTY),0) as CT from SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and DOCNO='PA02' and Sub_PartNO='{dr["PartNO"].ToString()}' group by Sub_PartNO");
                                        //if (dr_tmp != null && !dr_tmp.IsNull("CT") && dr_tmp["CT"].ToString() != "" && dr_tmp["CT"].ToString() != "0")
                                        //{
                                        //    db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set Math_EfficientCT={dr_tmp["CT"].ToString()} where SimulationId='{dr["SimulationId"].ToString()}'");
                                        //    tmp_down_StartTime = WebSocketServiceOJB.TimeCompute2DateTime(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(tmp["SimulationDate"]), int.Parse(dr_tmp["CT"].ToString()), false).ToString("yyyy-MM-dd HH:mm:ss.fff");
                                        //}

                                        tmp_down_ArrivalDate = Convert.ToDateTime(dr["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss");
                                        tmp_i = int.Parse(dr["NeedQTY"].ToString()) + int.Parse(dr["SafeQTY"].ToString());
                                        //if (int.Parse(dr["Source_StationNO_IndexSN"].ToString()) == 1) { tmp_i = int.Parse(dr["NeedQTY"].ToString())+ int.Parse(dr["SafeQTY"].ToString()); }
                                        tmp_up = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr["NeedId"].ToString()}' and Master_PartNO='{dr["PartNO"].ToString()}' and Apply_StationNO='{dr["Source_StationNO"].ToString()}' and IndexSN='{dr["Source_StationNO_IndexSN"].ToString()}'");
                                        tmp_down = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr["NeedId"].ToString()}' and PartNO='{dr["Master_PartNO"].ToString()}' and Source_StationNO='{dr["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN='{dr["IndexSN"].ToString()}' and PartSN<{dr["PartSN"].ToString()} and (Class='4' or Class='5') order by PartSN desc");
                                        if (tmp_down != null)
                                        {
                                            tmp_down_SID = $"'{tmp_down["SimulationId"].ToString()}'";
                                            tmp_down_Source_StationNO = $"'{tmp_down["Source_StationNO"].ToString()}'";
                                            tmp_down_ArrivalDate = Convert.ToDateTime(tmp_down["StartDate"]).ToString("yyyy-MM-dd HH:mm:ss");
                                        }
                                        #region 查找適合廠商
                                        mFNO = _SFC_Common.SelectDOC4ProductionMFNO(db, dr["PartNO"].ToString(), dr["SimulationId"].ToString(), in_NO, ref price);
                                        #endregion
                                        #region 查找適合入庫儲別
                                        _SFC_Common.SelectINStore(db, dr["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, in_NO, true);
                                        #endregion
                                        if (_SFC_Common.Create_DOC4stock(db, dr, mFNO, price, in_StoreNO, in_StoreSpacesNO, in_NO, tmp_i, "", "", "監測到排程委外加工需求", tmp_down_StartTime, tmp_down_ArrivalDate, "系統指派", ref docNumberNO))
                                        {
                                            tmp_02 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where PartNO='{dr["PartNO"].ToString()}' and WorkType='2' and SimulationId='{dr["SimulationId"].ToString()}' and MFNO='{mFNO}' and DOCNumberNO='{docNumberNO}'");
                                            if (tmp_02 == null)
                                            {
                                                string tmp_up_SID = tmp_up != null ? tmp_up["SimulationId"].ToString() : "";
                                                sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkingPaper] (ServerId,[Id],[WorkType],[PartNO],[Class],[IsOK],[NeedId],[SimulationId],[UP_SimulationId],[Down_SimulationId],[NeedQTY],[Price],[Unit],[MFNO],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],[APS_StationNO],[APS_StationNO_SID],[StartTime],[ArrivalDate],[EndTime],[UpdateTime],DOCNumberNO)
                                            VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('P')}','2','{dr["PartNO"].ToString()}','{dr["Class"].ToString()}','0','{dr["NeedId"].ToString()}','{dr["SimulationId"].ToString()}','{tmp_up_SID}',{tmp_down_SID},{tmp_i},{price},'PCS','{mFNO}','{in_StoreNO}','{in_StoreSpacesNO}','','',
                                            {tmp_down_Source_StationNO},{tmp_down_SID},'{tmp_down_StartTime}','{tmp_down_ArrivalDate}',NULL,'{logdate}','{docNumberNO}')";
                                                _ = db.DB_SetData(sql);
                                                _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsWPaper='1',DOCNumberNO='{docNumberNO}' where SimulationId='{dr["SimulationId"].ToString()}'");
                                                _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{docNumberNO}' where SimulationId='{dr["SimulationId"].ToString()}'");
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }

                            #region 安全存量 for 非生產件
                            sql = $"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and IS_WorkingPaper='1' and Class!='4' and Class!='5'";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                int tmp_i = 0;//需求量
                                string in_StoreNO = "";
                                string in_StoreSpacesNO = "";
                                string mFNO = "";
                                float price = 0;
                                foreach (DataRow dr in dt_tmp.Rows)
                                {
                                    mFNO = "";
                                    tmp_i = 0;
                                    #region 在倉可用量
                                    sql = @$"select sum(A.QTY) as mQTY,sum(B.KeepQTY+B.OverQTY) as kQTY from SoftNetMainDB.[dbo].[TotalStock] as A 
                                            join SoftNetMainDB.[dbo].[TotalStockII] as B on A.Id=B.Id and (B.KeepQTY!=0 or B.OverQTY!=0)
                                            where A.Class!='虛擬倉' and A.ServerId='{_Fun.Config.ServerId}' and A.PartNO='{dr["PartNO"].ToString()}' group by A.PartNO";
                                    tmp = db.DB_GetFirstDataByDataRow(sql);
                                    if (tmp != null && !tmp.IsNull("mQTY") && tmp["mQTY"].ToString() != "")
                                    {
                                        tmp_i = int.Parse(tmp["mQTY"].ToString()) - (tmp.IsNull("kQTY") ? 0 : int.Parse(tmp["kQTY"].ToString()));
                                    }
                                    else
                                    {
                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(QTY) as mQTY FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                                        if (tmp != null && !tmp.IsNull("mQTY") && tmp["mQTY"].ToString() != "")
                                        { tmp_i = int.Parse(tmp["mQTY"].ToString()); }
                                    }
                                    #endregion
                                    #region 在途可用量
                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(QTY) as mQTY FROM SoftNetMainDB.[dbo].[DOC1BuyII] where ServerId='{_Fun.Config.ServerId}' and IsOK='0' and PartNO='{dr["PartNO"].ToString()}'");
                                    if (tmp != null && !tmp.IsNull("mQTY") && tmp["mQTY"].ToString() != "")
                                    {
                                        tmp_i += (tmp.IsNull("mQTY") ? 0 : int.Parse(tmp["mQTY"].ToString()));
                                    }
                                    //tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(a.SafeQTY) as sQTY FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a,SoftNetSYSDB.[dbo].[APS_NeedData] as b where a.SafeQTY!=0 and a.PartNO='{dr["PartNO"].ToString()}' and a.NeedId=b.Id and b.State='6'");
                                    //if (tmp != null && !tmp.IsNull("sQTY") && tmp["sQTY"].ToString() != "")
                                    //{ tmp_i += int.Parse(tmp["sQTY"].ToString()); }
                                    #endregion
                                    #region 已在底稿量
                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(NeedQTY) as pQTY FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where PartNO='{dr["PartNO"].ToString()}'");
                                    if (tmp != null && !tmp.IsNull("pQTY") && tmp["pQTY"].ToString() != "")
                                    { tmp_i += int.Parse(tmp["pQTY"].ToString()); }
                                    #endregion


                                    if (int.Parse(dr["SafeQTY"].ToString()) > tmp_i || (_Fun.Config.Default_WorkingPaper_NOT_SafeQTY_DOC1 && tmp_i < 0))
                                    {

                                        #region 查找適合入庫儲別
                                        _SFC_Common.SelectINStore(db, dr["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "");
                                        #endregion
                                        #region 查找適合廠商
                                        mFNO = _SFC_Common.SelectDOC1BuyMFNO(db, dr["PartNO"].ToString(), "", "", ref price);
                                        #endregion
                                        string ArrivalDate = "NULL";
                                        #region 查找PP_EfficientDetail日期
                                        tmp = db.DB_GetFirstDataByDataRow($"select ROUND(sum(EfficientCycleTime)*sum(CountQTY)/sum(CountQTY),0) as CT from SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}' and DOCNO!='' and PartNO='{dr["PartNO"].ToString()}' group by PartNO");
                                        if (tmp != null && !tmp.IsNull("CT") && tmp["CT"].ToString() != "" && tmp["CT"].ToString() != "0")
                                        {
                                            ArrivalDate = DateTime.Now.AddSeconds(int.Parse(tmp["CT"].ToString())).ToString("yyyy-MM-dd HH:mm:ss.fff");
                                        }
                                        else { ArrivalDate = DateTime.Now.AddDays(7).ToString("yyyy-MM-dd HH:mm:ss.fff"); }
                                        #endregion
                                        tmp_i = int.Parse(dr["SafeQTY"].ToString()) - tmp_i;
                                        tmp_i = Math.Abs(tmp_i);
                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where PartNO='{dr["PartNO"].ToString()}' and WorkType='3' and MFNO='{mFNO}' and IN_StoreNO='{in_StoreNO}' and IN_StoreSpacesNO='{in_StoreSpacesNO}' order by ArrivalDate desc");
                                        if (tmp != null)
                                        { _ = db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkingPaper] set NeedQTY+={tmp_i} where Id='{tmp["Id"].ToString()}'"); }
                                        else
                                        {
                                            sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkingPaper] (ServerId,[Id],[WorkType],[PartNO],[Class],[IsOK],[NeedId],[SimulationId],[UP_SimulationId],[Down_SimulationId],[NeedQTY],[Price],[Unit],[MFNO],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],[APS_StationNO],[APS_StationNO_SID],[StartTime],[ArrivalDate],[EndTime],[UpdateTime])
                                            VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('P')}','3','{dr["PartNO"].ToString()}','{dr["Class"].ToString()}','0','','','','',{tmp_i},0,'PCS','{mFNO}','{in_StoreNO}','{in_StoreSpacesNO}','',
                                            '','','','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{ArrivalDate}',NULL,'{logdate}')";
                                            _ = db.DB_SetData(sql);
                                        }
                                        _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime) VALUES ('安全存量非生產件發現不足,已列入工作底稿等待處理. 料號:{dr["PartNO"].ToString()}','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");

                                    }
                                }
                            }

                            #endregion

                            #region 安全存量 for 生產件
                            sql = $"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and IS_WorkingPaper='1' and SafeQTY!=0 and (Class='4' or Class='5')";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                int tmp_i = 0;
                                string in_StoreNO = "";
                                string in_StoreSpacesNO = "";
                                foreach (DataRow dr in dt_tmp.Rows)
                                {
                                    tmp_i = 0;
                                    #region 庫存可用量
                                    sql = @$"select sum(A.QTY) as mQTY,sum(B.KeepQTY+B.OverQTY) as kQTY from SoftNetMainDB.[dbo].[TotalStock] as A 
                                            left join SoftNetMainDB.[dbo].[TotalStockII] as B on A.Id=B.Id and (B.KeepQTY!=0 or B.OverQTY!=0)
                                            where A.Class!='虛擬倉' and A.ServerId='{_Fun.Config.ServerId}' and A.PartNO='{dr["PartNO"].ToString()}' group by A.PartNO";
                                    tmp = db.DB_GetFirstDataByDataRow(sql);
                                    if (tmp != null)
                                    {
                                        int m = (tmp.IsNull("mQTY") || tmp["mQTY"].ToString() == "") ? 0 : int.Parse(tmp["mQTY"].ToString());
                                        int k = (tmp.IsNull("kQTY") || tmp["kQTY"].ToString() == "") ? 0 : int.Parse(tmp["kQTY"].ToString());
                                        tmp_i = m - k;
                                    }
                                    else
                                    {
                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(QTY) as mQTY FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}'");
                                        if (tmp != null && !tmp.IsNull("mQTY") && tmp["mQTY"].ToString() != "")
                                        { tmp_i = int.Parse(tmp["mQTY"].ToString()); }
                                    }
                                    #endregion

                                    #region 多生產的量
                                    //再線安全計畫量
                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(a.SafeQTY) as sQTY FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a,SoftNetSYSDB.[dbo].[APS_NeedData] as b where a.PartSN=0 and a.SafeQTY!=0 and a.PartNO='{dr["PartNO"].ToString()}' and a.NeedId=b.Id and b.State='6'");
                                    if (tmp != null && !tmp.IsNull("sQTY") && tmp["sQTY"].ToString() != "")
                                    { tmp_i += int.Parse(tmp["sQTY"].ToString()); }
                                    //多報工量
                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(Detail_QTY-a.NeedQTY) as sQTY FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a,SoftNetSYSDB.[dbo].[APS_NeedData] as b where a.APS_StationNO!='' and a.NoStation='0' and a.PartNO='{dr["PartNO"].ToString()}' and a.NeedId=b.Id and b.State='6'");
                                    if (tmp != null && !tmp.IsNull("sQTY") && tmp["sQTY"].ToString() != "" && int.Parse(tmp["sQTY"].ToString()) > 0)
                                    { tmp_i += int.Parse(tmp["sQTY"].ToString()); }
                                    #endregion

                                    #region 已存在模擬量
                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(NeedQTY) as pQTY FROM SoftNetSYSDB.[dbo].[APS_NeedData] where PartNO='{dr["PartNO"].ToString()}' and NeedType='5' and State!='6' and State!='9'");
                                    if (tmp != null && !tmp.IsNull("pQTY") && tmp["pQTY"].ToString() != "")
                                    { tmp_i += int.Parse(tmp["pQTY"].ToString()); }
                                    #endregion

                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(NeedQTY) as pQTY FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where PartNO='{dr["PartNO"].ToString()}' and (WorkType='4' or WorkType='1')");
                                    if (tmp != null && !tmp.IsNull("pQTY") && tmp["pQTY"].ToString() != "")
                                    { tmp_i += int.Parse(tmp["pQTY"].ToString()); }

                                    if (int.Parse(dr["SafeQTY"].ToString()) > tmp_i || (_Fun.Config.Default_WorkingPaper_NOT_SafeQTY_Order && tmp_i < 0))
                                    {
                                        string mBOMId = "";
                                        string apply_PP_Name = "";

                                        #region 查找適合入庫儲別
                                        _SFC_Common.SelectINStore(db, dr["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "");
                                        #endregion

                                        #region 查找適合的MBOMId,Apply_PP_Name
                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where PartNO='{dr["PartNO"].ToString()}' and PartSN=0 order by SimulationDate desc");
                                        if (tmp != null)
                                        {
                                            mBOMId = tmp["Apply_BOMId"].ToString();
                                            apply_PP_Name = tmp["Apply_PP_Name"].ToString();
                                        }
                                        else
                                        {
                                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr["PartNO"].ToString()}' and IsConfirm='1' and Main_Item='1' order by ExpiryDate desc,Version desc");
                                            if (tmp != null)
                                            {
                                                mBOMId = tmp["Id"].ToString();
                                                apply_PP_Name = tmp["Apply_PP_Name"].ToString();
                                            }
                                        }
                                        #endregion

                                        tmp_i = int.Parse(dr["SafeQTY"].ToString()) - tmp_i;
                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where PartNO='{dr["PartNO"].ToString()}' and WorkType='4' and MBOMId='{mBOMId}' and Apply_PP_Name='{apply_PP_Name}' and IN_StoreNO='{in_StoreNO}' and IN_StoreSpacesNO='{in_StoreSpacesNO}' order by ArrivalDate desc");
                                        if (tmp != null)
                                        { _ = db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkingPaper] set NeedQTY+={tmp_i} where Id='{tmp["Id"].ToString()}'"); }
                                        else
                                        {
                                            sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkingPaper] (ServerId,[Id],[WorkType],[PartNO],[Class],[IsOK],[NeedId],[SimulationId],[UP_SimulationId],[Down_SimulationId],[NeedQTY],[Price],[Unit],[MFNO],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],[APS_StationNO],[APS_StationNO_SID],[StartTime],[ArrivalDate],[EndTime],[UpdateTime],MBOMId,Apply_PP_Name)
                                            VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('P')}','4','{dr["PartNO"].ToString()}','{dr["Class"].ToString()}','0','','','','',{tmp_i},0,'PCS','','{in_StoreNO}','{in_StoreSpacesNO}','',
                                            '','','','{logdate}','{DateTime.Now.AddDays(3).ToString("yyyy-MM-dd HH:mm:ss.fff")}',NULL,'{logdate}','{mBOMId}','{apply_PP_Name}')";
                                            _ = db.DB_SetData(sql);
                                        }
                                        _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime) VALUES ('安全存量生產件發現不足,已列入工作底稿等待處理. 料號:{dr["PartNO"].ToString()}','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");

                                    }
                                }

                            }
                            #endregion
                            #endregion

                            #region 工站生產用領料單自動開立 原物料   //###???AC01暫時寫死  //###???? 少自動開立 上一站生產先入庫, 下一站生產前, 應該開出領出,繼續生產
                            docNumberNO = "";
                            string in_NO02 = "AC01";
                            //先查有無計畫
                            sql = $@"SELECT a.*,c.SimulationDate FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a
                                join SoftNetSYSDB.[dbo].[APS_NeedData] as b on b.Id=a.NeedId and b.State='6'
                                join SoftNetSYSDB.[dbo].APS_Simulation as c on c.SimulationId=a.SimulationId
                                where a.CalendarDate<'{logdate}' and (a.DOCNumberNO is null or a.DOCNumberNO='') and (a.NoStation='1' or (a.Class!='4' and a.Class!='5')) and a.Class!='7'";
                            dt_APS_PartNOTimeNote = db.DB_GetData(sql);

                            if (dt_APS_PartNOTimeNote != null && dt_APS_PartNOTimeNote.Rows.Count > 0)
                            {
                                bool is_run = true;
                                foreach (DataRow d in dt_APS_PartNOTimeNote.Rows)
                                {
                                    is_run = true;
                                    tmp_int = int.Parse(d["NeedQTY"].ToString()); //需求數量

                                    #region 先檢查DOC3stockII是否已有單據
                                    int stockQTY = 0;
                                    tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) as qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO02}'");
                                    if (tmp != null && !tmp.IsNull("qty"))
                                    {
                                        stockQTY = int.Parse(tmp["qty"].ToString());
                                    }
                                    if ((stockQTY - tmp_int) >= 0)
                                    {
                                        is_run = false;
                                        tmp = db.DB_GetFirstDataByDataRow($"select top 1 DOCNumberNO from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO02}'");
                                        if (tmp != null) { docNumberNO = tmp["DOCNumberNO"].ToString(); }
                                    }
                                    else
                                    {
                                        tmp_int -= stockQTY;
                                    }
                                    #endregion
                                    if (is_run)
                                    {
                                        //查有無Keep量
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
                                                        _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY-={tmp_int} where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", "生產用領料", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                        _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"生產用領料", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                                        tmp_int -= tmp_01;
                                                    }
                                                }
                                            }
                                            #endregion

                                            if (tmp_int > 0)
                                            {
                                                #region 計畫量不夠扣, 扣實體倉
                                                tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                {
                                                    foreach (DataRow d2 in tmp_dt2.Rows)
                                                    {
                                                        if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                        {
                                                            string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                            _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                                            tmp_int = 0;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                            if (tmp_01 != 0)
                                                            {
                                                                string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                                _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
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
                                                    _SFC_Common.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, "AC01");
                                                    #endregion
                                                    string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                    _ = _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", "AC01", tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                                }
                                                #endregion
                                            }
                                        }
                                        else
                                        {
                                            #region 沒keep量, 扣實體倉
                                            tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                            {
                                                foreach (DataRow d2 in tmp_dt2.Rows)
                                                {
                                                    if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                    {
                                                        string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                        _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                        if (tmp_01 != 0)
                                                        {
                                                            string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                            _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
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
                                                _SFC_Common.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, "AC01");
                                                #endregion
                                                string stationno = !d.IsNull("APS_StationNO") ? $"工站:{d["APS_StationNO"].ToString()} " : "";
                                                _ = _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", "AC01", tmp_int, "", "", $"{stationno}", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                            }
                                            #endregion
                                        }
                                    }
                                    if (docNumberNO != "")
                                    {
                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set DOCNumberNO='{docNumberNO}' where SimulationId='{d["SimulationId"].ToString()}'");
                                        if (db.DB_GetQueryCount($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where DOCNumberNO='' and SimulationId='{d["SimulationId"].ToString()}'") > 0)
                                        {
                                            _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set DOCNumberNO='{docNumberNO}' where SimulationId='{d["SimulationId"].ToString()}'");
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region 計畫性需求state=7 與 APS_Simulation,PartSN=-1 自動開立領料單

                            in_NO02 = "AC01";//###??? AC01要改參數化
                                             //先查有無計畫
                            sql = $@"SELECT a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                join SoftNetSYSDB.[dbo].[APS_NeedData] as b on b.Id=a.NeedId and b.State='7'
                                where a.StartDate<'{logdate}' and (a.DOCNumberNO is null or a.DOCNumberNO='')";
                            dt_APS_PartNOTimeNote = db.DB_GetData(sql);
                            sql = $@"SELECT a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                join SoftNetSYSDB.[dbo].[APS_NeedData] as b on b.Id=a.NeedId and b.State='6'
                                where a.StartDate<'{logdate}' and (a.DOCNumberNO is null or a.DOCNumberNO='') and PartSN=-1";
                            tmp_dt = db.DB_GetData(sql);
                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                            {
                                dt_APS_PartNOTimeNote.Merge(tmp_dt);
                            }
                            if (dt_APS_PartNOTimeNote != null && dt_APS_PartNOTimeNote.Rows.Count > 0)
                            {
                                bool is_run = true;
                                foreach (DataRow d in dt_APS_PartNOTimeNote.Rows)
                                {
                                    is_run = true;
                                    tmp_int = int.Parse(d["NeedQTY"].ToString()); //需求數量

                                    #region 先檢查DOC3stockII是否已有單據
                                    int stockQTY = 0;
                                    tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) as qty from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO02}'");
                                    if (tmp != null && !tmp.IsNull("qty"))
                                    {
                                        stockQTY = int.Parse(tmp["qty"].ToString());
                                    }
                                    if ((stockQTY - tmp_int) >= 0)
                                    {
                                        is_run = false;
                                        tmp = db.DB_GetFirstDataByDataRow($"select top 1 DOCNumberNO from SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO02}'");
                                        if (tmp != null) { docNumberNO = tmp["DOCNumberNO"].ToString(); }
                                    }
                                    else
                                    {
                                        tmp_int -= stockQTY;
                                    }
                                    #endregion
                                    if (is_run)
                                    {
                                        //查有無Keep量
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
                                                        _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY-={tmp_int} where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", "計畫性需求領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["KeepQTY"].ToString());
                                                        _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStockII] set KeepQTY=0 where Id='{d2["Id"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                        _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", "計畫性需求領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                                        tmp_int -= tmp_01;
                                                    }
                                                }
                                            }
                                            #endregion

                                            if (tmp_int > 0)
                                            {
                                                #region 計畫量不夠扣, 扣實體倉
                                                tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                {
                                                    foreach (DataRow d2 in tmp_dt2.Rows)
                                                    {
                                                        if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                        {
                                                            _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"計畫性需求領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                                            tmp_int = 0;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                            if (tmp_01 != 0)
                                                            {
                                                                _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"I計畫性需求領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
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
                                                    _SFC_Common.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, "AC01");
                                                    #endregion
                                                    _ = _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", "AC01", tmp_int, "", "", $"I計畫性需求領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                                }
                                                #endregion
                                            }
                                        }
                                        else
                                        {
                                            #region 沒keep量, 扣實體倉
                                            string _s = $"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc";
                                            tmp_dt2 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where Class!='虛擬倉' and ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' order by QTY desc");
                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                            {
                                                foreach (DataRow d2 in tmp_dt2.Rows)
                                                {
                                                    if (int.Parse(d2["QTY"].ToString()) >= tmp_int)
                                                    {
                                                        _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_int, "", "", $"計畫性需求領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                                        tmp_int = 0;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        int tmp_01 = int.Parse(d2["QTY"].ToString());
                                                        if (tmp_01 != 0)
                                                        {
                                                            _ = _SFC_Common.Create_DOC3stock(db, d, d2["StoreNO"].ToString(), d2["StoreSpacesNO"].ToString(), "", "", "AC01", tmp_01, "", "", $"計畫性需求領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
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
                                                _SFC_Common.SelectOUTStore(db, d["PartNO"].ToString(), ref out_StoreNO, ref out_StoreSpacesNO, "AC01");
                                                #endregion

                                                _ = _SFC_Common.Create_DOC3stock(db, d, out_StoreNO, out_StoreSpacesNO, "", "", "AC01", tmp_int, "", "", $"計畫性需求領出", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "SfcTimerloopautoRUN_DOC_Tick;RUNTimeServer", ref docNumberNO, "系統指派", false);
                                            }
                                            #endregion
                                        }
                                    }
                                    if (docNumberNO != "")
                                    {
                                        if (db.DB_GetQueryCount($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where DOCNumberNO='' and SimulationId='{d["SimulationId"].ToString()}'") > 0)
                                        {
                                            _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] set DOCNumberNO='{docNumberNO}' where SimulationId='{d["SimulationId"].ToString()}'");
                                        }
                                    }
                                }
                            }
                            #endregion
                        }
                        #region NeedId已結案的 DOC3stockII 單據,if IsOK='0', 依APS_Simulation_ErrorData_Clear_Day的天數,自動領/入庫
                        if (_Fun.Config.APS_Simulation_ErrorData_Clear_Day != 0 && _beforeTime.Day != DateTime.Now.Day)
                        {
                            _beforeTime = DateTime.Now;
                            dt_tmp = db.DB_GetData($@"SELECT a.NeedId,a.SimulationId FROM [SoftNetSYSDB].[dbo].[APS_Simulation] as a
                                                            join SoftNetSYSDB.[dbo].[APS_NeedData] as b on a.NeedId=b.Id and b.State='9' and b.Is_Close_DOC3stock='0'
                                                            where a.ServerId='{_Fun.Config.ServerId}' group by a.NeedId,a.SimulationId");
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                string sID = "";
                                List<string> needId = new List<string>();
                                foreach (DataRow row in dt_tmp.Rows)
                                {
                                    if (!needId.Contains(row["NeedId"].ToString())) { needId.Add(row["NeedId"].ToString()); }
                                    if (sID == "") { sID = $"'{row["SimulationId"].ToString()}'"; }
                                    else { sID = $"{sID},'{row["SimulationId"].ToString()}'"; }

                                }

                                if (sID != "")
                                {
                                    dt_tmp = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[DOC3stockII] where ServerId='{_Fun.Config.ServerId}' and IsOK='0' and ArrivalDate<='{DateTime.Now.AddDays(-_Fun.Config.APS_Simulation_ErrorData_Clear_Day).ToString("MM/dd/yyyy HH:mm:ss.fff")}' and SimulationId in ({sID})");
                                    if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                    {
                                        foreach (DataRow row in dt_tmp.Rows)
                                        {
                                            #region 寫入庫存
                                            if (row.IsNull("OUT_StoreNO") || row["OUT_StoreNO"].ToString().Trim() == "")
                                            { _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={row["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{row["PartNO"].ToString()}' and StoreNO='{row["IN_StoreNO"].ToString()}' and StoreSpacesNO='{row["IN_StoreSpacesNO"].ToString()}'"); }
                                            else { _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={row["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{row["PartNO"].ToString()}' and StoreNO='{row["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{row["OUT_StoreSpacesNO"].ToString()}'"); }
                                            #endregion

                                            string writeSQL = "";
                                            if (row.IsNull("StartTime")) { writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                                            else { writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                                            _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT=0{writeSQL} where Id='{row["Id"].ToString()}' and DOCNumberNO='{row["DOCNumberNO"].ToString()}' and IsOK='0'");
                                        }
                                    }
                                    if (needId.Count > 0)
                                    {
                                        sID = "";
                                        foreach (string s in needId)
                                        {
                                            if (sID == "") { sID = $"'{s}'"; }
                                            else { sID = $"{sID},'{s}'"; }
                                        }
                                        if (sID != "")
                                        {
                                            _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_NeedData] set Is_Close_DOC3stock='1' where Id in ({sID})");
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        #region NeedId已結案,清除APS_Simulation_ErrorData 非單據 與 工單應關未關 的狀態
                        //ActionType的定義 空=建立異常 0=資料正確(已處理) 1=已逾時不處裡  2=NeedId已結束強制清除
                        sql = $@"SELECT a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                            join SoftNetSYSDB.[dbo].[APS_NeedData] as b on b.Id=a.NeedId and b.State='9'
                            where a.ActionType='' and a.ServerId='{_Fun.Config.ServerId}' and a.ErrorType!='06' and a.ErrorType!='07' and a.ErrorType!='08' and a.ErrorType!='09' and a.ErrorType!='10' and a.ErrorType!='12'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='2',ActionLogDate='{logdate}' where ServerId='{_Fun.Config.ServerId}' and NeedId='{dt_tmp.Rows[0]["NeedId"].ToString()}' and ActionType=''");
                        }
                        #endregion

                        #region NeedId已結案清除工作底稿的資料
                        sql = $@"select a.Id from SoftNetSYSDB.[dbo].[APS_WorkingPaper] as a 
                        join SoftNetSYSDB.[dbo].[APS_NeedData] as b on a.NeedId=b.Id and b.State='9'
                        where a.IsOK='0' and a.ServerId='{_Fun.Config.ServerId}'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            string id = "";
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (id == "") { id = $"'{d["Id"]}'"; }
                                else { id = $"{id},'{d["Id"]}'"; }

                            }
                            if (id != "")
                            {
                                _ = db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_WorkingPaper] where Id in ({id})");
                            }
                        }
                        #endregion

                        #region 清除APS_WorkingPaper底稿 的 ArrivalDate 已過期
                        _ = db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_WorkingPaper] where ServerId='{_Fun.Config.ServerId}' and IsOK='0' and WorkType!='2' and SendTime is NULL and ArrivalDate<='{logdate}'");
                        #endregion

                        #region 清除 BarCode_TMP 已過期
                        _ = db.DB_SetData($"delete SoftNetMainDB.[dbo].[BarCode_TMP] where ServerId='{_Fun.Config.ServerId}' and CONVERT(varchar(100), FailTime, 111)<='{DateTime.Now.ToString("yyyy/MM/dd")}'");
                        #endregion
                        _Fun._a01 = threadLoopTime.ElapsedMilliseconds;
                    }

                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 處理被記錄時間到的干涉 APS_WarningData 已寫 12工單未關,18工站未停止, 19計畫日期已超過, 21工單已關閉強制檢查(加 22 , 23 的資料寫入)
                        sql = $"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WarningData] where IsDEL='0' and ServerId='{_Fun.Config.ServerId}' and WarningDate<='{logdate}'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            string MailSub = $"{_Fun.Config.ServerId} 系統主動 干涉 通知";
                            string MailBody = "";
                            List<string> run_HasWeb_Id_Change = new List<string>();
                            foreach (DataRow d_Error in dt_tmp.Rows)
                            {
                                switch (d_Error["ErrorType"].ToString())
                                {
                                    case "21"://工單已關閉,強制檢查APS_PartNOTimeNote 報工的合理性
                                        {
                                            DataRow dr_PP_Station = null;
                                            #region 干涉修正報工數量
                                            tmp_dt = db.DB_GetData($@"SELECT b.PartNO,b.PartSN,b.BOMQTY,b.DOCNumberNO as wo,b.Source_StationNO_IndexSN,b.Apply_PP_Name,b.IsWPaper,b.Source_StationNO,b.Source_StationNO_IndexSN,b.Master_PartNO,b.Apply_StationNO,a.*  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a
                                                                                        join SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.NeedId=b.NeedId and a.SimulationId=b.SimulationId
                                                                                        where a.NeedId='{d_Error["NeedId"].ToString()}' and a.NoStation='0' and a.NeedQTY!=0 order by b.PartSN desc,b.Class desc");
                                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                            {
                                                int qty = 0;
                                                tmp = db.DB_GetFirstDataByDataRow($@"SELECT a.* FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a
                                                                                        join SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.NeedId=b.NeedId and a.SimulationId=b.SimulationId and b.PartSN=0
                                                                                        where a.NeedId='{d_Error["NeedId"].ToString()}' and a.NoStation='0' and a.NeedQTY!=0");
                                                if (tmp != null)
                                                {
                                                    qty = int.Parse(tmp["Detail_QTY"].ToString()) + int.Parse(tmp["Detail_Fail_QTY"].ToString());
                                                }
                                                if (qty > 0)
                                                {
                                                    int offset = int.Parse(tmp["NeedQTY"].ToString()) - qty;
                                                    if (offset < 0) { offset = 0; }
                                                    int detail_qty = 0;
                                                    foreach (DataRow d in tmp_dt.Rows)
                                                    {
                                                        if (int.Parse(d["PartSN"].ToString()) == 0) { continue; }
                                                        detail_qty = (int.Parse(d["NeedQTY"].ToString()) - offset) - (int.Parse(d["Detail_QTY"].ToString()) + int.Parse(d["Detail_Fail_QTY"].ToString()));
                                                        if (detail_qty > 0 || d["APS_StationNO"].ToString() == _Fun.Config.OutPackStationName)
                                                        {
                                                            if (d["APS_StationNO"].ToString() == _Fun.Config.OutPackStationName)
                                                            {
                                                                if (d["DOCNumberNO"].ToString() == "")
                                                                {
                                                                    #region 補開委外加工單
                                                                    string in_StoreNO = "";
                                                                    string in_StoreSpacesNO = "";
                                                                    string in_NO = "AA02";//###???暫時寫死
                                                                    float price = 0;
                                                                    string docNumberNO = "";
                                                                    string docdate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                                                                    string tmp_down_SID = "NULL";
                                                                    string tmp_down_Source_StationNO = "NULL";
                                                                    //string tmp_down_StartTime = "";
                                                                    string tmp_down_ArrivalDate = "";

                                                                    #region 查找適合廠商
                                                                    string mFNO = _SFC_Common.SelectDOC4ProductionMFNO(db, d["PartNO"].ToString(), d["SimulationId"].ToString(), in_NO, ref price);
                                                                    #endregion
                                                                    #region 查找適合入庫儲別
                                                                    _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, in_NO, true);
                                                                    #endregion
                                                                    if (_SFC_Common.Create_DOC4stock(db, d, mFNO, price, in_StoreNO, in_StoreSpacesNO, in_NO, detail_qty, "", "", "工單關閉後補單", docdate, docdate, "系統指派", ref docNumberNO))
                                                                    {
                                                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET DOCNumberNO='{docNumberNO}' where SimulationId='{d["SimulationId"].ToString()}'");
                                                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET DOCNumberNO='{docNumberNO}' where SimulationId='{d["SimulationId"].ToString()}'");

                                                                        DataRow tmp_up = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{d["NeedId"].ToString()}' and Master_PartNO='{d["PartNO"].ToString()}' and Apply_StationNO='{d["Source_StationNO"].ToString()}' and IndexSN='{d["Source_StationNO_IndexSN"].ToString()}'");
                                                                        DataRow tmp_down = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{d["NeedId"].ToString()}' and PartNO='{d["Master_PartNO"].ToString()}' and Source_StationNO='{d["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN='{d["IndexSN"].ToString()}' and PartSN<{d["PartSN"].ToString()} and (Class='4' or Class='5') order by PartSN desc");
                                                                        if (tmp_down != null)
                                                                        {
                                                                            tmp_down_SID = $"'{tmp_down["SimulationId"].ToString()}'";
                                                                            tmp_down_Source_StationNO = $"'{tmp_down["Source_StationNO"].ToString()}'";
                                                                            tmp_down_ArrivalDate = Convert.ToDateTime(tmp_down["StartDate"]).ToString("yyyy-MM-dd HH:mm:ss");
                                                                        }


                                                                        string tmp_up_SID = tmp_up != null ? tmp_up["SimulationId"].ToString() : "";
                                                                        sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkingPaper] (ServerId,[Id],[WorkType],[PartNO],[Class],[IsOK],[NeedId],[SimulationId],[UP_SimulationId],[Down_SimulationId],[NeedQTY],[Price],[Unit],[MFNO],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],[APS_StationNO],[APS_StationNO_SID],[StartTime],[ArrivalDate],[EndTime],[UpdateTime],DOCNumberNO)
                                                                            VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('P')}','2','{d["PartNO"].ToString()}','{d["Class"].ToString()}','0','{d["NeedId"].ToString()}','{d["SimulationId"].ToString()}','{tmp_up_SID}',{tmp_down_SID},{detail_qty},{price},'PCS','{mFNO}','{in_StoreNO}','{in_StoreSpacesNO}','','',
                                                                            {tmp_down_Source_StationNO},{tmp_down_SID},'{docdate}','{docdate}',NULL,'{logdate}','{docNumberNO}')";
                                                                        _ = db.DB_SetData(sql);
                                                                    }
                                                                    #endregion
                                                                }
                                                                else
                                                                {
                                                                    #region 判斷是否要減少委外加工單數量
                                                                    tmp = db.DB_GetFirstDataByDataRow($@"SELECT sum(QTY) as qty FROM SoftNetMainDB.[dbo].[DOC4ProductionII] where SimulationId='{d["SimulationId"].ToString()}'");
                                                                    if (offset > 0 && tmp != null && !tmp.IsNull("qty") && tmp["qty"].ToString() != "")
                                                                    {
                                                                        if (int.Parse(tmp["qty"].ToString()) > offset)
                                                                        {
                                                                            if (db.DB_GetQueryCount($"select Id from SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{d_Error["NeedId"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString().Trim()}' and ErrorType='23'") <= 0)
                                                                            {
                                                                                detail_qty = offset;
                                                                                if (db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate],DATA_Remark) VALUES
                                                                                    ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','23','{d_Error["NeedId"].ToString()}','{d["SimulationId"].ToString()}','{d["APS_StationNO"].ToString()}','{d["DOCNumberNO"].ToString().Trim()}','{d["PartNO"].ToString()}','','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{detail_qty.ToString()}')"))
                                                                                {
                                                                                    MailBody = $"{MailBody}<p>疑似委外加工單數量多下 委外加工單:{d["wo"].ToString()} 工站:{d["APS_StationNO"].ToString()} 數量多出:{Math.Abs(detail_qty).ToString()}</p>";
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                    #endregion
                                                                }
                                                                continue;
                                                            }
                                                            if (d["wo"].ToString().Trim() != "")
                                                            {
                                                                #region 補報工量
                                                                dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["APS_StationNO"].ToString()}'");
                                                                if (dr_PP_Station["Station_Type"].ToString() == "8")
                                                                {
                                                                    MUTIStationObj keys = new MUTIStationObj();
                                                                    keys.StationNO = d["APS_StationNO"].ToString();
                                                                    keys.SI_SimulationId = d["SimulationId"].ToString();
                                                                    keys.SI_IndexSN = d["Source_StationNO_IndexSN"].ToString();
                                                                    keys.SI_OrderNO = d["wo"].ToString();
                                                                    keys.SI_PP_Name = d["Apply_PP_Name"].ToString();
                                                                    keys.SI_PartNO = d["PartNO"].ToString();
                                                                    keys.SI_OKQTY = detail_qty;
                                                                    keys.SI_FailQTY = 0;
                                                                    keys.SI_Slect_OPNOs = "系統指派";
                                                                    string message = "";
                                                                    string stackTrace = "";
                                                                    bool is_reportOK = _SFC_Common.Reporting_STView2Work(db, dr_PP_Station, "系統指派", keys, true, ref message, ref stackTrace);
                                                                    if (!is_reportOK)
                                                                    {
                                                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs 工站:{keys.StationNO} 干涉修正報工數量失敗 {message} {stackTrace}", true);
                                                                    }
                                                                    else
                                                                    {
                                                                        MailBody = $"{MailBody}<p>工單:{d["wo"].ToString()} 工站:{d["APS_StationNO"].ToString()} 系統自動補報工, 數量補:{detail_qty.ToString()}</p>";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    LabelWork keys = new LabelWork();
                                                                    keys.Station = d["APS_StationNO"].ToString();
                                                                    keys.SimulationId = d["SimulationId"].ToString();
                                                                    keys.OrderNO = d["wo"].ToString();
                                                                    keys.IndexSN = d["Source_StationNO_IndexSN"].ToString();
                                                                    keys.OKQTY = detail_qty;
                                                                    keys.FailQTY = 0;
                                                                    keys.OPNO = "系統指派";
                                                                    string message = "";
                                                                    string stackTrace = "";
                                                                    string ViewBagERRMsg = "";
                                                                    bool is_reportOK = _SFC_Common.Reporting_LabelWork(db, dr_PP_Station, "系統指派", keys, true, ref message, ref stackTrace, ref ViewBagERRMsg);
                                                                    if (!is_reportOK)
                                                                    {
                                                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs 工站:{keys.Station} 干涉修正報工數量失敗 {message} {stackTrace}", true);
                                                                    }
                                                                    else
                                                                    {
                                                                        MailBody = $"{MailBody}<p>工單:{d["wo"].ToString()} 工站:{d["APS_StationNO"].ToString()} 系統自動補報工, 數量補:{detail_qty.ToString()} ,須至網頁功能自行修正</p>";
                                                                    }
                                                                }
                                                                #endregion
                                                            }
                                                        }
                                                        else
                                                        {
                                                            #region 判斷是否要減少報工數量
                                                            if (db.DB_GetQueryCount($"select Id from SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{d_Error["NeedId"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["wo"].ToString().Trim()}' and ErrorType='22'") <= 0)
                                                            {
                                                                if (db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate],DATA_Remark) VALUES
                                                                    ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','22','{d_Error["NeedId"].ToString()}','{d["SimulationId"].ToString()}','{d["APS_StationNO"].ToString()}','{d["wo"].ToString().Trim()}','{d["PartNO"].ToString()}','','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{Math.Abs(detail_qty).ToString()}')"))
                                                                {
                                                                    MailBody = $"{MailBody}<p>疑似重複報工 工單:{d["wo"].ToString()} 工站:{d["APS_StationNO"].ToString()} 數量多出:{Math.Abs(detail_qty).ToString()}</p>";
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                    }
                                                }
                                            }
                                            _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WarningData] SET IsDEL='1' where Id='{d_Error["Id"].ToString()}'");
                                            #endregion
                                        }
                                        break;
                                    case "12"://工單應關未關
                                        {
                                            bool isrun = false;
                                            DateTime intime = DateTime.Now;
                                            string needId = "";
                                            string pP_Name = "";
                                            string calendarName = _Fun.Config.DefaultCalendarName;

                                            string lastStation = "";

                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].PP_WorkOrder where (EndTime is null OR TmpClose IS NULL) and ServerId='{_Fun.Config.ServerId}' and OrderNO='{d_Error["DOCNumberNO"].ToString()}'");
                                            if (dr_tmp != null)
                                            {
                                                calendarName = dr_tmp["CalendarName"].ToString();
                                                needId = dr_tmp["NeedId"].ToString();
                                                pP_Name = dr_tmp["PP_Name"].ToString();

                                                #region 判動干涉條件是否成立
                                                //bool isRun_PP_ProductProcess_Item = true;
                                                //if (db.DB_GetQueryCount($"SELECT * FROM SoftNetSYSDB.[dbo].PP_WO_Process_Item WHERE ServerId='{_Fun.Config.ServerId}' and OrderNO='{d_Error["DOCNumberNO"].ToString()}'") > 0)
                                                //{ isRun_PP_ProductProcess_Item = false; }
                                                //DataTable dt_WO_Stations = SFC_FUN.Process_ALLSation_RE_Custom(_Fun.Config.ServerId, "1", _Fun.Config.Db, pP_Name, "ORDER BY IndexSN, PP_Name ASC", isRun_PP_ProductProcess_Item, d_Error["DOCNumberNO"].ToString());
                                                List<string> station_list = new List<string>();
                                                DataTable dt_WO_Stations = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where ServerId='{_Fun.Config.ServerId}' and NeedId='{needId}'");
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

                                                if (station_list.Count > 0)
                                                {
                                                    List<string> station_list_NO_StationNO_Merge = new List<string>();
                                                    #region 不含合併站
                                                    DataTable dt_station_list_NO_StationNO_Merge = db.DB_GetData($"select Apply_StationNO from SoftNetSYSDB.[dbo].APS_Simulation where NeedId='{needId}' and PartSN>=0 group by Apply_StationNO");
                                                    if (dt_station_list_NO_StationNO_Merge != null && dt_station_list_NO_StationNO_Merge.Rows.Count > 0)
                                                    {
                                                        foreach (DataRow d in dt_station_list_NO_StationNO_Merge.Rows)
                                                        {
                                                            station_list_NO_StationNO_Merge.Add($"'{d["Apply_StationNO"].ToString()}'");
                                                        }
                                                    }
                                                    #endregion

                                                    string stationSQL = "";
                                                    #region 含合併站, 所有站
                                                    foreach (string s in station_list)
                                                    {
                                                        if (stationSQL == "") { stationSQL = $"'{s}'"; }
                                                        else { stationSQL = $"{stationSQL},'{s}'"; }
                                                    }
                                                    #endregion

                                                    if (stationSQL != "")
                                                    {
                                                        DataTable dt_tmp01 = null;
                                                        #region 檢查APS_Simulation LastStation的狀態IsOK是否='1', 與 APS_PartNOTimeNote數量是否足夠
                                                        string indexSN = "";
                                                        if (needId != "")
                                                        {
                                                            dt_tmp01 = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{d_Error["DOCNumberNO"].ToString()}' and StationNO in ({stationSQL})");
                                                            if (dt_tmp01 != null && dt_tmp01.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow d3 in dt_tmp01.Rows)
                                                                {
                                                                    if (bool.Parse(d3["IsLastStation"].ToString())) { lastStation = d3["StationNO"].ToString(); indexSN = d3["IndexSN"].ToString(); }
                                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and Apply_PP_Name='{pP_Name}' and Source_StationNO='{d3["StationNO"].ToString()}' and Source_StationNO_IndexSN={d3["IndexSN"].ToString()} and DOCNumberNO='{d_Error["DOCNumberNO"].ToString()}'");
                                                                    if (dr_tmp != null)
                                                                    {
                                                                        if (bool.Parse(dr_tmp["IsChackQTY"].ToString()))
                                                                        {
                                                                            DataRow dr_tmp4 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId}' and SimulationId='{dr_tmp["SimulationId"].ToString()}' and NeedQTY>Detail_QTY");
                                                                            if (dr_tmp4 != null)
                                                                            {
                                                                                isrun = true; break;
                                                                            }
                                                                        }
                                                                        if (bool.Parse(dr_tmp["IsChackIsOK"].ToString()) && !bool.Parse(dr_tmp["IsOK"].ToString()))
                                                                        {
                                                                            isrun = true; break;
                                                                        }

                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            //###??? 非計畫性生產的公單判斷
                                                        }
                                                        #endregion

                                                        #region 檢查Manufacture 與 ManufactureII 狀態是否ok, 並異動
                                                        if (!isrun)
                                                        {
                                                            dt_tmp01 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{d_Error["DOCNumberNO"].ToString()}' and StationNO in ({stationSQL}) order by RemarkTimeE desc");
                                                            if (dt_tmp01 != null && dt_tmp01.Rows.Count > 0)
                                                            {
                                                                foreach (DataRow d3 in dt_tmp01.Rows)
                                                                {
                                                                    if (!d3.IsNull("RemarkTimeS") && d3.IsNull("RemarkTimeE"))
                                                                    {
                                                                        isrun = true; break;
                                                                    }
                                                                    else
                                                                    {
                                                                        if (db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[ManufactureII] where Id='{d3["Id"].ToString()}'"))
                                                                        {
                                                                            _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{needId}','{d3["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','RUNTimeServer','智慧關站 系統主動干涉','{pP_Name}','{lastStation}','','{d_Error["DOCNumberNO"].ToString()}','系統指派',{d3["IndexSN"].ToString()})");
                                                                            run_HasWeb_Id_Change.Add(d3["StationNO"].ToString());
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            if (!isrun)
                                                            {
                                                                dt_tmp01 = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{d_Error["DOCNumberNO"].ToString()}' and StationNO in ({stationSQL}) order by State");
                                                                if (dt_tmp01 != null && dt_tmp01.Rows.Count > 0)
                                                                {
                                                                    foreach (DataRow dr_M in dt_tmp01.Rows)
                                                                    {
                                                                        if (dr_M["State"].ToString() == "1")
                                                                        {
                                                                            isrun = true; break;
                                                                        }
                                                                        else
                                                                        {
                                                                            _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET Label_ProjectType='0',OrderNO='',OP_NO='',Master_PP_Name='',PP_Name='',PartNO='',IndexSN=0,Station_Custom_IndexSN='',StationNO_Custom_DisplayName='',State='4' where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_M["StationNO"].ToString()}'");
                                                                            _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{needId}','{dr_M["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','RUNTimeServer','智慧關站 系統主動干涉','{pP_Name}','{lastStation}','','{d_Error["DOCNumberNO"].ToString()}','系統指派',{dr_M["IndexSN"].ToString()})");

                                                                            #region 更新電子Tag
                                                                            if (dr_M["Config_macID"].ToString().Trim() != "")
                                                                            {
                                                                                string isUpdate = "1";
                                                                                if (!_Fun.Is_Tag_Connect) { isUpdate = "0"; }
                                                                                {
                                                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                                                                    string tmp_s = "";
                                                                                    string tmp_ShowValue = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"\",\"Text2\":\"\",\"PartNO\":\"\",\"Text3\":\"\",\"PartName\":\"\",\"Text4\":\"\",\"QTY\":\"\",\"Text5\":\"\",\"EfficientCT\":\"\",\"Text6\":\"\",\"Rate\":\"\",\"Text7\":\"累計量:\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{dr_M["StationNO"].ToString()};0;;0\",\"outtime\":0";
                                                                                    if (dr_tmp["Version"].ToString().Trim() != "" && dr_tmp["Version"].ToString().Trim().Substring(0, 2) == "42")
                                                                                    {
                                                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_M["StationNO"].ToString()}'");
                                                                                        tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_M["StationNO"].ToString()}\",\"text17\":\"{dr_tmp["StationName"].ToString()}\"";
                                                                                        tmp_s = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{tmp_ShowValue}";
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        tmp_ShowValue = $"{tmp_ShowValue},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"{dr_M["StationNO"].ToString()}\",\"text17\":\"\"";
                                                                                        tmp_s = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{tmp_ShowValue},\"ledrgb\":\"0\",\"ledstate\":0";
                                                                                    }
                                                                                    if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                                                    {
                                                                                        _Fun.Tag_Write(db, dr_M["Config_macID"].ToString(), "干涉關站", tmp_s);
                                                                                    }
                                                                                    _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{tmp_ShowValue}',Ledrgb='0',Ledstate=0,StationNO='{dr_M["StationNO"].ToString()}',Type='1',OrderNO='',IndexSN='',StoreNO='',StoreSpacesNO='',QTY=0,IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                                                                }
                                                                            }
                                                                            #endregion
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        #endregion

                                                        if (!isrun && lastStation != "")
                                                        {
                                                            DataRow dr_PP_Station = null;
                                                            #region 干涉修正報工數量
                                                            tmp_dt = db.DB_GetData($@"SELECT b.PartSN,b.BOMQTY,b.DOCNumberNO as wo,b.Source_StationNO_IndexSN,b.Apply_PP_Name,a.*  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a
                                                                                        join SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.NeedId=b.NeedId and a.SimulationId=b.SimulationId
                                                                                        where a.NeedId='{d_Error["NeedId"].ToString()}' and a.NoStation='0' and a.NeedQTY!=0 order by b.PartSN desc,b.Class desc");
                                                            if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                                                            {
                                                                int qty = 0;
                                                                tmp = db.DB_GetFirstDataByDataRow($@"SELECT a.* FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a
                                                                                        join SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.NeedId=b.NeedId and a.SimulationId=b.SimulationId and b.PartSN=0
                                                                                        where a.NeedId='{d_Error["NeedId"].ToString()}' and a.NoStation='0' and a.NeedQTY!=0");
                                                                if (tmp != null)
                                                                {
                                                                    qty = int.Parse(tmp["Detail_QTY"].ToString()) + int.Parse(tmp["Detail_Fail_QTY"].ToString());
                                                                }
                                                                if (qty > 0)
                                                                {
                                                                    int offset = int.Parse(tmp["NeedQTY"].ToString()) - qty;
                                                                    if (offset < 0) { offset = 0; }
                                                                    int detail_qty = 0;
                                                                    foreach (DataRow d in tmp_dt.Rows)
                                                                    {
                                                                        if (int.Parse(d["PartSN"].ToString()) == 0) { continue; }
                                                                        detail_qty = (int.Parse(d["NeedQTY"].ToString()) - offset) - (int.Parse(d["Detail_QTY"].ToString()) + int.Parse(d["Detail_Fail_QTY"].ToString()));
                                                                        if (detail_qty > 0)
                                                                        {
                                                                            if (d["APS_StationNO"].ToString() == _Fun.Config.OutPackStationName)
                                                                            {
                                                                                continue;
                                                                            }
                                                                            if (d["wo"].ToString().Trim() != "")
                                                                            {
                                                                                dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["APS_StationNO"].ToString()}'");
                                                                                if (dr_PP_Station["Station_Type"].ToString() == "8")
                                                                                {
                                                                                    MUTIStationObj keys = new MUTIStationObj();
                                                                                    keys.StationNO = d["APS_StationNO"].ToString();
                                                                                    keys.SI_SimulationId = d["SimulationId"].ToString();
                                                                                    keys.SI_IndexSN = d["Source_StationNO_IndexSN"].ToString();
                                                                                    keys.SI_OrderNO = d["wo"].ToString();
                                                                                    keys.SI_PP_Name = d["Apply_PP_Name"].ToString();
                                                                                    keys.SI_PartNO = d["PartNO"].ToString();
                                                                                    keys.SI_OKQTY = detail_qty;
                                                                                    keys.SI_FailQTY = 0;
                                                                                    keys.SI_Slect_OPNOs = "系統指派";
                                                                                    string message = "";
                                                                                    string stackTrace = "";
                                                                                    bool is_reportOK = _SFC_Common.Reporting_STView2Work(db, dr_PP_Station, "系統指派", keys, true, ref message, ref stackTrace);
                                                                                    if (!is_reportOK)
                                                                                    {
                                                                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs 工站:{keys.StationNO} 干涉修正報工數量失敗 {message} {stackTrace}", true);
                                                                                    }
                                                                                }
                                                                                else
                                                                                {
                                                                                    LabelWork keys = new LabelWork();
                                                                                    keys.Station = d["APS_StationNO"].ToString();
                                                                                    keys.SimulationId = d["SimulationId"].ToString();
                                                                                    keys.OrderNO = d["wo"].ToString();
                                                                                    keys.IndexSN = d["Source_StationNO_IndexSN"].ToString();
                                                                                    keys.OKQTY = detail_qty;
                                                                                    keys.FailQTY = 0;
                                                                                    keys.OPNO = "系統指派";
                                                                                    string message = "";
                                                                                    string stackTrace = "";
                                                                                    string ViewBagERRMsg = "";
                                                                                    bool is_reportOK = _SFC_Common.Reporting_LabelWork(db, dr_PP_Station, "系統指派", keys, true, ref message, ref stackTrace, ref ViewBagERRMsg);
                                                                                    if (!is_reportOK)
                                                                                    {
                                                                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs 工站:{keys.Station} 干涉修正報工數量失敗 {message} {stackTrace}", true);
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WarningData] SET IsDEL='1' where Id='{d_Error["Id"].ToString()}'");
                                                            #endregion

                                                            bool isLastStation = true;
                                                            dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{lastStation}'");
                                                            if (dr_PP_Station != null)
                                                            {
                                                                if (dr_PP_Station["Station_Type"].ToString() == "8")
                                                                {
                                                                    //多工單
                                                                    //###???DOCNO暫時寫死BC01
                                                                    //###???入庫要考慮倉庫最大安置量
                                                                    if (station_list_NO_StationNO_Merge.Count > 0)
                                                                    {
                                                                        #region 半成品 or 成品 入庫 與 餘料入庫
                                                                        tmp_dt = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and Apply_StationNO in ({string.Join(",", station_list_NO_StationNO_Merge)}) and Apply_PP_Name='{pP_Name}' and (Class='4' or Class='5') and Source_StationNO is not null");
                                                                        if (tmp_dt != null)
                                                                        {
                                                                            foreach (DataRow d in tmp_dt.Rows)
                                                                            {
                                                                                DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId}' and SimulationId='{d["SimulationId"].ToString()}'");
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
                                                                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}'");
                                                                                        if (tmp == null)
                                                                                        {
                                                                                            #region 查找適合庫儲別
                                                                                            _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "BC01");
                                                                                            #endregion
                                                                                            _ = _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref tmp_no, "系統指派");
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            in_StoreNO = tmp["StoreNO"].ToString();
                                                                                            in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                                                                            _ = _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref tmp_no, "系統指派");
                                                                                        }
                                                                                        sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{tmp_no}',Next_StoreQTY+={qty} where SimulationId='{d["SimulationId"].ToString()}'";
                                                                                        _ = db.DB_SetData(sql);
                                                                                        #endregion
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                        #endregion
                                                                    }

                                                                    #region 非半成品 or 成品 餘料入庫  //###???暫時寫死 EB01
                                                                    dt_APS_PartNOTimeNote = db.DB_GetData($"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId}' and ((Class!='4' and Class!='5') or NoStation='1')");
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
                                                                                    dt_tmp01 = db.DB_GetData($@"SELECT a.*,c.SimulationDate FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                                                            join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.DOCNumberNO,1,4) and b.DOCType='3' and b.ServerId='{_Fun.Config.ServerId}'
                                                                                            join SoftNetSYSDB.[dbo].[APS_Simulation] as c on c.SimulationId=a.SimulationId
                                                                                            where a.SimulationId='{d["SimulationId"].ToString()}' order by OUT_StoreNO,OUT_StoreSpacesNO,IsOK");
                                                                                    string docNumberNO = "";
                                                                                    foreach (DataRow d2 in dt_tmp01.Rows)
                                                                                    {
                                                                                        if ((int.Parse(d2["QTY"].ToString()) - useQYU) > 0)
                                                                                        {
                                                                                            _ = _SFC_Common.Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", useQYU, "", d2["Id"].ToString(), $"工單結束退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref docNumberNO, "系統指派");
                                                                                            break;
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            _ = _SFC_Common.Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", int.Parse(d2["QTY"].ToString()), "", d2["Id"].ToString(), $"工單結束退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "ChangeStatus;TagResult_EnterKeyController", ref docNumberNO, "系統指派");
                                                                                            useQYU -= int.Parse(d2["QTY"].ToString());
                                                                                            if (useQYU <= 0) { break; }
                                                                                        }
                                                                                    }
                                                                                    sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{docNumberNO}',Next_StoreQTY+={wQTY} where SimulationId='{d["SimulationId"].ToString()}'";
                                                                                    _ = db.DB_SetData(sql);
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                    #endregion

                                                                }
                                                                else
                                                                {
                                                                    //單工單
                                                                    {
                                                                        //###???DOCNO暫時寫死BC01
                                                                        //###???入庫要考慮倉庫最大安置量
                                                                        #region 半成品 or 成品入庫 與 餘料入庫 Class='4' or Class='5'
                                                                        tmp_dt = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}' and Apply_StationNO in ({string.Join(",", station_list_NO_StationNO_Merge)}) and Apply_PP_Name='{pP_Name}' and (Class='4' or Class='5') and Source_StationNO is not null");
                                                                        if (tmp_dt != null)
                                                                        {
                                                                            foreach (DataRow d in tmp_dt.Rows)
                                                                            {
                                                                                DataRow tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{pP_Name}' and SimulationId='{d["SimulationId"].ToString()}'");
                                                                                if (tmp_dr != null)
                                                                                {
                                                                                    #region 多生產或多領入庫
                                                                                    int qty = int.Parse(tmp_dr["Detail_QTY"].ToString()) - (int.Parse(tmp_dr["Next_StationQTY"].ToString()) + int.Parse(tmp_dr["Next_StoreQTY"].ToString()));
                                                                                    if (qty > 0)
                                                                                    {
                                                                                        string tmp_no = "";
                                                                                        string in_StoreNO = "";
                                                                                        string in_StoreSpacesNO = "";
                                                                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT a.StoreNO,a.StoreSpacesNO,a.Id FROM SoftNetMainDB.[dbo].[TotalStock] as a,SoftNetMainDB.[dbo].[TotalStockII] as b where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and a.Id=b.Id and b.SimulationId='{d["SimulationId"].ToString()}'");
                                                                                        if (tmp == null)
                                                                                        {

                                                                                            #region 查找適合庫儲別
                                                                                            _SFC_Common.SelectINStore(db, d["PartNO"].ToString(), ref in_StoreNO, ref in_StoreSpacesNO, "BC01");
                                                                                            #endregion
                                                                                            _ = _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "RUNTimeServer;SfcTimerloopthread_Tick", ref tmp_no, "系統指派");

                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            in_StoreNO = tmp["StoreNO"].ToString();
                                                                                            in_StoreSpacesNO = tmp["StoreSpacesNO"].ToString();
                                                                                            _ = _SFC_Common.Create_DOC3stock(db, d, "", "", in_StoreNO, in_StoreSpacesNO, "BC01", qty, "", "", "工站移轉加工品餘料入庫", Convert.ToDateTime(d["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "RUNTimeServer;SfcTimerloopthread_Tick", ref tmp_no, "系統指派");
                                                                                        }
                                                                                        sql = $"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET Store_DOCNumberNO='{tmp_no}',Next_StoreQTY+={qty} where SimulationId='{d["SimulationId"].ToString()}'";
                                                                                        _ = db.DB_SetData(sql);
                                                                                    }
                                                                                    #endregion
                                                                                }
                                                                            }
                                                                        }
                                                                        #endregion

                                                                        //###???暫時寫死 EB01
                                                                        #region 非半成品成品 原物料 餘料退回入庫 Class!='4' and Class!='5'
                                                                        dt_APS_PartNOTimeNote = db.DB_GetData($"SELECT *  FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{needId}' and ((Class!='4' and Class!='5') or NoStation='1')");
                                                                        if (dt_APS_PartNOTimeNote != null && dt_APS_PartNOTimeNote.Rows.Count > 0)
                                                                        {
                                                                            DataRow tmp_dr = null;
                                                                            int sQTY = 0;
                                                                            int useQYU = 0;
                                                                            foreach (DataRow d in dt_APS_PartNOTimeNote.Rows)
                                                                            {
                                                                                tmp_dr = db.DB_GetFirstDataByDataRow($@"SELECT sum(a.QTY) as Total FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                                                                            join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.DOCNumberNO,1,4) and b.DOCType='3'
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
                                                                                                { _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] SET QTY-={useQYU.ToString()} where Id='{d2["Id"].ToString()}' and DOCNumberNO='{d2["DOCNumberNO"].ToString()}'"); }
                                                                                                else
                                                                                                { _ = _SFC_Common.Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", useQYU, "", d2["Id"].ToString(), $"生產結束餘料退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "RUNTimeServer;SfcTimerloopthread_Tick", ref docNumberNO, "系統指派"); }
                                                                                                break;
                                                                                            }
                                                                                            else
                                                                                            {
                                                                                                if (!bool.Parse(d2["IsOK"].ToString()))
                                                                                                { _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] SET QTY=0,Remark='生產結束清除用量' where Id='{d2["Id"].ToString()}' and DOCNumberNO='{d2["DOCNumberNO"].ToString()}'"); }
                                                                                                else
                                                                                                { _ = _SFC_Common.Create_DOC3stock(db, d, "", "", d2["OUT_StoreNO"].ToString(), d2["OUT_StoreSpacesNO"].ToString(), "EB01", int.Parse(d2["QTY"].ToString()), "", d2["Id"].ToString(), $"生產結束餘料退回入庫", Convert.ToDateTime(d2["SimulationDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "RUNTimeServer;SfcTimerloopthread_Tick", ref docNumberNO, "系統指派"); }
                                                                                                useQYU -= int.Parse(d2["QTY"].ToString());
                                                                                                if (useQYU <= 0) { break; }
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                        #endregion

                                                                    }

                                                                }
                                                                #region 送Service處理後續
                                                                string status = "5";
                                                                if (isLastStation)
                                                                {
                                                                    status = "5";//關站加關工單

                                                                    //###??? 要考慮 APS_NeedData與APS_Simulation , 同一個NeedId可能有多個 工單時, 發到Softnet Service的_CloseWO Code會有問題

                                                                }
                                                                //發到Softnet Service      1.bnName, 2.StationNO, 3.obj.Name, 4._projectWithoutExtension, 5.obj.UserName(OP), 6.obj.OrderNO, 7.Station_IndexSN  8.OutQtyTagName, 9.FailQtyTagName, 10 = IsTagValueCumulative, 11 = WaitTagValueDIO 12.CalculatePauseTime,13.WaitTime_Formula ,14.CycleTime_Formula, 15.IsWaitTargetFinish
                                                                DataRow dr_M = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{lastStation}'");
                                                                if (dr_M != null)
                                                                {
                                                                    //if (WebSocketServiceOJB.RmsSend(dr_PP_Station["RMSName"].ToString(), 1, $"WebChangeStationStatus,{status},{lastStation},WEBProg,{lastStation},{dr_M["OP_NO"].ToString()},{d_Error["DOCNumberNO"].ToString()},{indexSN},{dr_M["Config_OutQtyTagName"].ToString()},{dr_M["Config_FailQtyTagName"].ToString()},{dr_M["Config_IsTagValueCumulative"].ToString()},{dr_M["Config_WaitTagValueDIO"].ToString()},{dr_M["Config_CalculatePauseTime"].ToString()},{dr_M["Config_WaitTime_Formula"].ToString()},{dr_M["Config_CycleTime_Formula"].ToString()},{dr_M["Config_IsWaitTargetFinish"].ToString()}"))
                                                                    //{
                                                                    if (needId != "" && status == "5")
                                                                    {
                                                                        CloseWO(dr_M["OP_NO"].ToString());

                                                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where NeedId='{needId}'");
                                                                        _ = db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{needId}'");
                                                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_Simulation] SET IsOK='1' where NeedId='{needId}'");
                                                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WarningData] SET IsDEL='1' where Id='{d_Error["Id"].ToString()}'");
                                                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='0',ActionLogDate='{logdate}' where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO='{d_Error["DOCNumberNO"].ToString()}' and ErrorType='12'");
                                                                        MailBody = $"{MailBody}<p>系統將 {d_Error["DOCNumberNO"].ToString()} 工單 除了退入庫單據之外,已自動結案</p>";
                                                                        _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId) VALUES ('已將排程編號:{needId} 工單:{d_Error["DOCNumberNO"].ToString()} 除了退入庫單據之外的工作,已自動結案','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}')");

                                                                        DataTable dt_ManufactureII = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{needId}'");
                                                                        if (dt_ManufactureII != null && dt_ManufactureII.Rows.Count > 0)
                                                                        {
                                                                            string ii_ID = "";
                                                                            foreach (DataRow dr_II in dt_ManufactureII.Rows)
                                                                            {
                                                                                if (ii_ID == "") { ii_ID = $"'{dr_II["SimulationId"].ToString()}'"; }
                                                                                else { ii_ID = $"{ii_ID},'{dr_II["SimulationId"].ToString()}'"; }
                                                                            }
                                                                            if (ii_ID != "")
                                                                            {
                                                                                _ = db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[ManufactureII] where SimulationId in ({ii_ID})");
                                                                            }
                                                                        }

                                                                        continue;
                                                                    }
                                                                    //}
                                                                    //else
                                                                    //{ isrun = true; }
                                                                }
                                                                else
                                                                { isrun = true; }
                                                                #endregion

                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion
                                            }
                                            if (isrun && lastStation != "")
                                            {
                                                #region 寫下次確認時段
                                                DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{lastStation}'");
                                                sql = $@"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_PP_Station["CalendarName"].ToString()}' and Holiday>='{intime.ToString("MM/dd/yyyy")}'";
                                                dr_tmp = db.DB_GetFirstDataByDataRow(sql);
                                                if (dr_tmp != null)
                                                {
                                                    isrun = true;
                                                    DateTime etime = DateTime.Now;
                                                    DateTime stime2 = DateTime.Now;
                                                    string[] comp = dr_tmp["Shift_Morning"].ToString().Trim().Split(',');
                                                    stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                    if (etime >= intime && intime >= stime2)
                                                    {
                                                        isrun = false;
                                                        intime = etime.AddMinutes(isARGs10_offset);
                                                    }
                                                    if (isrun)
                                                    {
                                                        comp = dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',');
                                                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        if (etime >= intime && intime >= stime2)
                                                        {
                                                            isrun = false;
                                                            intime = etime.AddMinutes(isARGs10_offset);
                                                        }
                                                    }
                                                    if (isrun)
                                                    {
                                                        comp = dr_tmp["Shift_Night"].ToString().Trim().Split(',');
                                                        if (int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                        {
                                                            etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1);
                                                        }
                                                        else
                                                        {
                                                            etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        }
                                                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        if (etime >= intime && intime >= stime2)
                                                        {
                                                            isrun = false;
                                                            intime = etime.AddMinutes(isARGs10_offset);
                                                        }
                                                    }
                                                    if (isrun)
                                                    {
                                                        comp = dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',');
                                                        if (int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                        { etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                                        else
                                                        {
                                                            etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        }
                                                        stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        if (etime >= intime && intime >= stime2)
                                                        {
                                                            isrun = false;
                                                            intime = etime.AddMinutes(isARGs10_offset);
                                                        }
                                                    }
                                                    else if (isrun)
                                                    {
                                                        comp = dr_tmp["Shift_Morning"].ToString().Trim().Split(',');
                                                        intime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1);
                                                    }
                                                    _ = db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WarningData] SET WarningDate='{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}' where Id='{d_Error["Id"].ToString()}'");
                                                    _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId) VALUES ('排程編號:{needId} 工單:{d_Error["DOCNumberNO"].ToString()}應該可以結案,但發現尚有工作未完成, 預計 {intime.ToString("yyyy-MM-dd HH:mm:ss.fff")} 後再檢查是否可結案','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}')");

                                                }
                                                #endregion

                                            }

                                        }
                                        break;
                                    case "18"://工站應停止未停止
                                        {
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d_Error["StationNO"].ToString()}'");
                                            if (dr_tmp != null)
                                            {
                                                string station_List = "";
                                                if (dr_tmp["Station_Type"].ToString() == "8")
                                                {
                                                    #region 多工單
                                                    DataTable dt_ManufactureII = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d_Error["StationNO"].ToString()}' and StartTime is not NULL and EndTime is NULL and RemarkTimeS is not NULL and RemarkTimeE is NULL");
                                                    if (dt_ManufactureII != null && dt_ManufactureII.Rows.Count > 0)
                                                    {
                                                        foreach (DataRow d2 in dt_ManufactureII.Rows)
                                                        {
                                                            string wedate = "";
                                                            DataRow dr_stop = db.DB_GetFirstDataByDataRow($"SELECT top 1 * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday<='{DateTime.Now.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_tmp["CalendarName"].ToString()}' order by CalendarName,Holiday desc");
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
                                                                dr_stop = db.DB_GetFirstDataByDataRow($"SELECT top 1 * FROM SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d2["OrderNO"].ToString()}' and IndexSN={d2["IndexSN"].ToString()} and LOGDateTime>'{logTime.ToString("MM/dd/yyyy HH:mm:ss")}' and LOGDateTime='{logTime.ToString("yyyy/MM/dd")}' and OrderNO='{d2["OrderNO"].ToString()}' and OperateType like '%報工%' and OperateType not like '%網頁報工%' order by LOGDateTime desc");
                                                                if (dr_stop != null)
                                                                {
                                                                    logTime = Convert.ToDateTime(dr_stop["LOGDateTime"]);
                                                                }
                                                                wedate = logTime.ToString("MM/dd/yyyy HH:mm:ss");
                                                                _ = db.DB_SetData($"update SoftNetMainDB.[dbo].[ManufactureII] set RemarkTimeE='{wedate}' where Id='{d2["Id"].ToString()}'");
                                                            }
                                                            else
                                                            {
                                                                wedate = logdate;
                                                                _ = db.DB_SetData($"update SoftNetMainDB.[dbo].[ManufactureII] set RemarkTimeE='{logdate}' where Id='{d2["Id"].ToString()}'");
                                                            }

                                                            _ = db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,OP_NO,IndexSN) VALUES 
                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{d_Error["NeedId"].ToString()}','{d2["SimulationId"].ToString()}','{wedate}','RUNTimeServer','干涉停工','{d2["PP_Name"].ToString()}','{d_Error["StationNO"].ToString()}','{d2["PartNO"].ToString()}','{d2["OrderNO"].ToString()}','',{d2["IndexSN"].ToString()})");
                                                            station_List = $"{station_List} {d_Error["StationNO"].ToString()}";
                                                            MailBody = $"{MailBody}<p>非工作時間已強制設定停工,  工站:{d_Error["StationNO"].ToString()}  工序編號:{d2["IndexSN"].ToString()} 製程:{d2["PP_Name"].ToString()}  料號:{d2["PartNO"].ToString()}</p>";
                                                            #region 通知網頁更新
                                                            _ = SendWebSocketClent_INFO($"SendALLClient,StationStateChangEvent,8,{d2["Id"].ToString()},2");
                                                            //if (WebSocketServiceOJB != null)
                                                            //{
                                                            //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                                            //    {
                                                            //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                                            //        {
                                                            //            if (r.Key != null && r.Value.socket != null)
                                                            //            {
                                                            //                WebSocketServiceOJB.Send(r.Value.socket, $"StationStateChangEvent,8,{d2["Id"].ToString()},2");// 參數 1=單/多工站 2=StationNO/Id 3=狀態
                                                            //            }
                                                            //        }
                                                            //    }
                                                            //}
                                                            #endregion
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                {
                                                    #region 單工單
                                                    var ledrgb = "0";
                                                    string meg = "";
                                                    DataRow dr_M = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d_Error["StationNO"].ToString()}'");
                                                    if (dr_M != null)
                                                    {
                                                        // var webSocketService = (SNWebSocketService)_Fun.DiBox.GetService(typeof(SNWebSocketService));
                                                        if (dr_M["State"].ToString() == "1")
                                                        {
                                                            if (_SFC_Common != null)
                                                            {
                                                                //###??? 將來ChangeStatus , LabelProject_Start_Stop 改呼叫webSocketService
                                                                if (_Fun.Config.RUNMode == '1')//學習模式
                                                                {
                                                                    LabelProject keys = new LabelProject();
                                                                    keys.StationNO = d_Error["StationNO"].ToString();
                                                                    keys.PartNO = dr_M["PartNO"].ToString();
                                                                    meg = LabelProject_Start_Stop(db, "2", d_Error["StationNO"].ToString(), dr_M, dr_M["OP_NO"].ToString(), ref keys);
                                                                }
                                                                else
                                                                {
                                                                    meg = _SFC_Common.ChangeStatus(d_Error["StationNO"].ToString(), $"{d_Error["StationNO"].ToString()},2", "RUNTimeServer", true);
                                                                    if (meg == "")
                                                                    {
                                                                        //###??? 這裡要是, 標籤是否換燈
                                                                        _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set Ledrgb='{ledrgb}',Ledstate=0 where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                                                                    }
                                                                }
                                                                if (meg == "")
                                                                {
                                                                    station_List = $"{station_List} {d_Error["StationNO"].ToString()}";
                                                                    MailBody = $"{MailBody}<p>非工作時間已強制設定停工,  工站:{d_Error["StationNO"].ToString()}  工序編號:{dr_M["IndexSN"].ToString()} 製程:{dr_M["PP_Name"].ToString()}  料號:{dr_M["PartNO"].ToString()}</p>";
                                                                }
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (station_List != "")
                                                {
                                                    _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId,SimulationId) VALUES ('判定為非工作時間,系統主動停止生產下列工站 {station_List}','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d_Error["NeedId"].ToString()}','{d_Error["SimulationId"].ToString()}')");
                                                }
                                            }
                                            _ = db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WarningData] SET IsDEL='1' where Id='{d_Error["Id"].ToString()}'");
                                        }
                                        break;
                                    case "19"://工站延後產出時間
                                        {
                                            //SendWebSocketClent_INFO("");
                                            //if (WebSocketServiceOJB != null)
                                            //{
                                            //    string[] data = new string[] { "", d_Error["SimulationId"].ToString(), "", "自動延後產出時間" };
                                            //    if (WebSocketServiceOJB.RefreshRunSetSimulation(db, data, ref run_HasWeb_Id_Change))
                                            //    {
                                            //        db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WarningData] SET IsDEL='1' where Id='{d_Error["Id"].ToString()}'");
                                            //    }
                                            //    else
                                            //    {
                                            //        string _s = "";//###???
                                            //    }
                                            //}
                                        }
                                        break;
                                        //###??? 其他未寫
                                }
                            }
                            if (MailBody != "") { _Fun.Mail_Send(_Fun.Config.SendMonitorMail00.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail00.Split(',')[1].Split(';'), MailSub, MailBody, null, false); }

                            //if (WebSocketServiceOJB != null && run_HasWeb_Id_Change.Count > 0)
                            //{
                            foreach (string sno in run_HasWeb_Id_Change)
                            {
                                _ = SendWebSocketClent_INFO($"SendALLClient,HasWeb_Id_Change,STView2Work_PageReload,{sno}");
                                //#region 通知網頁更新
                                //try
                                //{
                                //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                //    {
                                //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                //        {
                                //            if (r.Key != null && r.Value.socket != null)
                                //            {
                                //                WebSocketServiceOJB.Send(r.Value.socket, $"HasWeb_Id_Change,STView2Work_PageReload,{sno}");
                                //            }
                                //        }
                                //    }
                                //}
                                //catch (Exception ex)
                                //{
                                //    System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs 工站:{sno} 發HasWeb_Id_Changeg失敗 {ex.Message} {ex.StackTrace}", true);
                                //}
                                //#endregion
                            }
                            //}
                        }
                        #endregion

                        #region 處裡工作底稿的Mail前置通知
                        sql = $"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where IsOK='0' and ServerId='{_Fun.Config.ServerId}' and StartTime<='{logdate}' order by WorkType,MFNO,PartNO";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            string sendTime = DateTime.Now.AddMinutes(isARGs10_offset).ToString("MM/dd/yyyy HH:mm:ss.fff");
                            string in_NO = "";
                            string docNumberNO = "";
                            string MailSub = $"{_Fun.Config.ServerId} 工作底稿 干涉通知 PS:(若以下資訊需修改,請至[排程管理][需求工作底稿] 功能頁修改)";
                            string MailBody = "";
                            string tmp_MFNO = "";
                            List<string> run_HasWeb_Id_Change = new List<string>();
                            string id_AGE = "";
                            string writesql = "";
                            if (!_Fun.Config.Default_WorkingPaper_AGE01)
                            {
                                sendTime = DateTime.Now.AddMinutes(_Fun.Config.Default_WorkingPaper_AGE02).ToString("MM/dd/yyyy HH:mm:ss.fff");
                                writesql = $",IsOK='2',SendTime='{sendTime}'";
                            }
                            else
                            {
                                sendTime = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff");
                                writesql = $",IsOK='1',SendTime='{sendTime}'";
                            }
                            foreach (DataRow d_Paper in dt_tmp.Rows)
                            {
                                switch (d_Paper["WorkType"].ToString())
                                {
                                    case "1"://排程採購
                                    case "3"://補存貨量
                                        {
                                            if (d_Paper["MFNO"].ToString() != "")
                                            {
                                                in_NO = "AA02";//###???暫時寫死
                                                if (d_Paper["MFNO"].ToString() != tmp_MFNO) { docNumberNO = ""; tmp_MFNO = d_Paper["MFNO"].ToString(); }
                                                if (d_Paper.IsNull("DOCNumberNO") || d_Paper["DOCNumberNO"].ToString() == "")
                                                {
                                                    if (_SFC_Common.Create_DOC1stock(db, d_Paper, tmp_MFNO, float.Parse(d_Paper["Price"].ToString()), d_Paper["IN_StoreNO"].ToString(), d_Paper["IN_StoreSpacesNO"].ToString(), in_NO, int.Parse(d_Paper["NeedQTY"].ToString()), "", "", "底稿干涉採購", Convert.ToDateTime(d_Paper["StartTime"]).ToString("yyyy-MM-dd HH:mm:ss"), Convert.ToDateTime(d_Paper["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "系統指派", ref docNumberNO))
                                                    {
                                                    }
                                                }
                                                else { docNumberNO = d_Paper["DOCNumberNO"].ToString(); }
                                                if (docNumberNO != "" && tmp_MFNO != "")
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d_Paper["PartNO"].ToString()}'");
                                                    MailBody = $"{MailBody}<p>採購單: {docNumberNO} 料號:{d_Paper["PartNO"].ToString()} 品名:{dr_tmp["PartName"].ToString()} 規格:{dr_tmp["Specification"].ToString()} 數量:{d_Paper["NeedQTY"].ToString()} 單價:{d_Paper["Price"].ToString()} 供應商:{tmp_MFNO}</p>";
                                                    id_AGE = $"{id_AGE};{d_Paper["Id"].ToString()}";
                                                    _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkingPaper] SET MFNO='{tmp_MFNO}',DOCNumberNO='{docNumberNO}'{writesql} where ServerId='{_Fun.Config.ServerId}' and Id='{d_Paper["Id"].ToString()}'");
                                                }
                                            }
                                        }
                                        break;
                                    case "2"://排程委外
                                        {
                                            if (d_Paper["MFNO"].ToString() != "")
                                            {
                                                in_NO = "PA02";//###???暫時寫死
                                                if (d_Paper["MFNO"].ToString() != tmp_MFNO) { docNumberNO = ""; tmp_MFNO = d_Paper["MFNO"].ToString(); }
                                                if (d_Paper.IsNull("DOCNumberNO") || d_Paper["DOCNumberNO"].ToString() == "")
                                                {
                                                    if (!_SFC_Common.Create_DOC4stock(db, d_Paper, tmp_MFNO, float.Parse(d_Paper["Price"].ToString()), d_Paper["IN_StoreNO"].ToString(), d_Paper["IN_StoreSpacesNO"].ToString(), in_NO, int.Parse(d_Paper["NeedQTY"].ToString()), "", "", "底稿干涉委外加工", Convert.ToDateTime(d_Paper["StartTime"]).ToString("yyyy-MM-dd HH:mm:ss"), Convert.ToDateTime(d_Paper["ArrivalDate"]).ToString("yyyy-MM-dd HH:mm:ss"), "系統指派", ref docNumberNO))
                                                    {
                                                        continue;
                                                    }
                                                }
                                                else { docNumberNO = d_Paper["DOCNumberNO"].ToString(); }
                                                int shortQTY = int.Parse(d_Paper["NeedQTY"].ToString());
                                                #region 檢查上一站實際完成量,並修改單據量
                                                tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d_Paper["UP_SimulationId"].ToString()}'");
                                                if (tmp != null)
                                                {
                                                    if (db.DB_GetQueryCount($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d_Paper["SimulationId"].ToString()}' and Source_StationNO_IndexSN=1") <= 0)
                                                    {
                                                        shortQTY = int.Parse(tmp["Detail_QTY"].ToString());
                                                        if (shortQTY > 0)
                                                        {
                                                            tmp = db.DB_GetFirstDataByDataRow($"select sum(QTY) as editQTY from SoftNetMainDB.[dbo].[DOC4ProductionII] where SimulationId='{d_Paper["SimulationId"].ToString()}' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}'");
                                                            int editQTY = int.Parse(tmp["editQTY"].ToString());
                                                            if (shortQTY != editQTY)
                                                            {
                                                                if (shortQTY > editQTY)
                                                                {
                                                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC4ProductionII] where SimulationId='{d_Paper["SimulationId"].ToString()}' and IsOK='0' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}' order by ArrivalDate");
                                                                    if (tmp != null)
                                                                    { _ = db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC4ProductionII] set  QTY+={(shortQTY - editQTY).ToString()} where Id='{tmp["Id"].ToString()}'"); }
                                                                }
                                                                else
                                                                {
                                                                    DataTable dt_tmp2 = db.DB_GetData($"select * from SoftNetMainDB.[dbo].[DOC4ProductionII] where SimulationId='{d_Paper["SimulationId"].ToString()}' and IsOK='0' and SUBSTRING(DOCNumberNO,1,4)='{in_NO}'");
                                                                    if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                                                                    {
                                                                        int tmp_i = editQTY - shortQTY;
                                                                        if (tmp_i > 0)
                                                                        {
                                                                            foreach (DataRow d2 in dt_tmp2.Rows)
                                                                            {
                                                                                if (int.Parse(d2["QTY"].ToString()) > tmp_i)
                                                                                {
                                                                                    _ = db.DB_SetData($"update SoftNetMainDB.[dbo].[DOC4ProductionII] set  QTY-={tmp_i.ToString()} where Id='{d2["Id"].ToString()}'");
                                                                                    break;
                                                                                }
                                                                                else
                                                                                {
                                                                                    tmp_i -= int.Parse(d2["QTY"].ToString());
                                                                                    _ = db.DB_SetData($"delete from SoftNetMainDB.[dbo].[DOC4ProductionII] where Id='{d2["Id"].ToString()}'");
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion
                                                if (shortQTY > 0 && docNumberNO != "" && tmp_MFNO != "")
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d_Paper["PartNO"].ToString()}'");
                                                    MailBody = $"{MailBody}<p>委外加工單: {docNumberNO} 料號:{d_Paper["PartNO"].ToString()} 品名:{dr_tmp["PartName"].ToString()} 規格:{dr_tmp["Specification"].ToString()} 計畫數量:{d_Paper["NeedQTY"].ToString()} 實際可發量:{shortQTY.ToString()} 單價:{d_Paper["Price"].ToString()} 供應商:{tmp_MFNO}</p>";
                                                    id_AGE = $"{id_AGE};{d_Paper["Id"].ToString()}";
                                                    _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkingPaper] SET MFNO='{tmp_MFNO}',DOCNumberNO='{docNumberNO}'{writesql} where ServerId='{_Fun.Config.ServerId}' and Id='{d_Paper["Id"].ToString()}'");
                                                }
                                            }
                                        }
                                        break;
                                    case "4"://廠內生產
                                        {
                                            if (d_Paper["MBOMId"].ToString() != "" || d_Paper["Apply_PP_Name"].ToString() != "")
                                            {
                                                string a_pp_Name = d_Paper["Apply_PP_Name"].ToString().Trim();
                                                if (a_pp_Name == "")
                                                {
                                                    DataRow dr2 = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[BOM] where Id='{d_Paper["MBOMId"].ToString()}'");
                                                    if (dr2 != null) { a_pp_Name = d_Paper["Apply_PP_Name"].ToString().Trim(); }
                                                }
                                                if (a_pp_Name != "")
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d_Paper["PartNO"].ToString()}'");
                                                    MailBody = $"{MailBody}<p>廠內生產工單 料號:{d_Paper["PartNO"].ToString()} 品名:{dr_tmp["PartName"].ToString()} 規格:{dr_tmp["Specification"].ToString()} 數量:{d_Paper["NeedQTY"].ToString()} 母BOM編號:{d_Paper["MBOMId"].ToString()} 製程:{d_Paper["Apply_PP_Name"].ToString()}</p>";
                                                    id_AGE = $"{id_AGE};{d_Paper["Id"].ToString()}";
                                                    if (!_Fun.Config.Default_WorkingPaper_AGE01)
                                                    {
                                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkingPaper] SET IsOK='2',SendTime='{DateTime.Now.AddMinutes(_Fun.Config.Default_WorkingPaper_AGE02).ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and Id='{d_Paper["Id"].ToString()}'");
                                                    }
                                                    else
                                                    {
                                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_WorkingPaper] SET IsOK='1'{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")} where ServerId='{_Fun.Config.ServerId}' and Id='{d_Paper["Id"].ToString()}'");
                                                    }
                                                }
                                            }
                                        }
                                        break;
                                }
                            }
                            if (MailBody != "")
                            {
                                if (_Fun.Config.Default_WorkingPaper_AGE01)
                                { MailBody = $"<form action='http://{_Fun.Config.LocalWebURL}/WorkingPaper/MailAutoAction/{id_AGE}'>{MailBody}<hr /><input type=submit value='按此紐 將立即幫助您完成單據發送' style='width:100%';height='70px' /><hr /></form>"; }
                                else
                                { MailBody = $"<form action='http://{_Fun.Config.LocalWebURL}/WorkingPaper/MailAutoAction/{id_AGE}'>{MailBody}<hr />系統將於{_Fun.Config.Default_WorkingPaper_AGE02.ToString()}分鐘後,自動發出正式單據<hr /></form>"; }
                                _Fun.Mail_Send(_Fun.Config.SendMonitorMail00.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail00.Split(',')[1].Split(';'), MailSub, MailBody, null, false);
                                _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime) VALUES ('工作底稿準備主動干涉作業, 明細已Mail發出','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                            }
                            #region 通知網頁更新
                            //try
                            //{
                            //lock (WebSocketServiceOJB.lock__WebSocketList)
                            //{
                            //    foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                            //    {
                            //        if (r.Key != null && r.Value.socket != null)
                            //        {
                            //            WebSocketServiceOJB.Send(r.Value.socket, $"HasWeb_Id_Change,STView2Work_PageReload,{sno}");
                            //        }
                            //    }
                            //}
                            //}
                            //catch (Exception ex)
                            //{
                            //    System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs 工站:{sno} 發HasWeb_Id_Changeg失敗 {ex.Message} {ex.StackTrace}", true);
                            //}
                            #endregion
                        }
                        #endregion

                        #region 處裡工作底稿自動干涉
                        sql = $"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WorkingPaper] where IsOK='2' and ServerId='{_Fun.Config.ServerId}' and SendTime<='{logdate}' order by WorkType,MFNO,PartNO";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {

                            string MailSub4 = "";
                            string MailSub3 = "";
                            string MailSub2 = "";
                            string MailBody4 = "";
                            string MailBody3 = "";
                            string MailBody2 = "";
                            string tmp_MFNO = "";
                            List<string> needID_data = new List<string>();
                            #region 排程設定參數 args
                            RunSimulation_Arg args = new RunSimulation_Arg();
                            args.ARGs = new List<bool>() { false, false, false, false, false, false, false, false, true, true, false, false, false, false, false, false, false, false, false, true };
                            #endregion
                            foreach (DataRow d_Paper in dt_tmp.Rows)
                            {
                                switch (d_Paper["WorkType"].ToString())
                                {
                                    case "1"://排程採購
                                    case "3"://補存貨量
                                        {
                                            if (d_Paper["MFNO"].ToString() != tmp_MFNO)
                                            {
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' and MFNO='{d_Paper["MFNO"].ToString()}'");
                                                MailSub3 = $"{_Fun.Config.ServerId} 廠商別:{dr_tmp["MFName"].ToString()} 採購通知單 請依下列到貨日期出貨";
                                                if (MailBody3 != "")
                                                {
                                                    MailBody3 = $"{MailBody3}</tbody></table></div>";
                                                    _Fun.Mail_Send(_Fun.Config.SendMonitorMail00.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail00.Split(',')[1].Split(';'), MailSub3, MailBody3, null, false);

                                                }
                                                MailBody3 = "<div><table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'><thead><tr><th>採購單號</th><th>料件編號</th><th>品名</th><th>規格</th><th>數量</th><th>單位</th><th>單價</th><th>到貨需求日</th><th>廠內生產碼</th></tr></thead><tbody>";
                                                tmp_MFNO = d_Paper["MFNO"].ToString();
                                            }
                                            DataTable dt_tmp2 = db.DB_GetData($"select a.* from SoftNetMainDB.[dbo].[DOC1BuyII] as a,SoftNetMainDB.[dbo].[DOC1Buy] as b where b.ServerId='{_Fun.Config.ServerId}' and b.DOCNumberNO='{d_Paper["DOCNumberNO"].ToString()}' and a.DOCNumberNO=b.DOCNumberNO");
                                            if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                                            {
                                                foreach (DataRow d2 in dt_tmp2.Rows)
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d_Paper["PartNO"].ToString()}'");
                                                    MailBody3 = $@"{MailBody3}<tr><th>{d2["DOCNumberNO"].ToString()}</th><th>{d2["PartNO"].ToString()}</th><th>{dr_tmp["PartName"].ToString()}</th><th>{dr_tmp["Specification"].ToString()}</th>
                                                                    <th>{d2["QTY"].ToString()}</th><th>{d2["Unit"].ToString()}</th><th>{d2["Price"].ToString()}</th><th>{d2["ArrivalDate"].ToString()}</th><th>{d2["SimulationId"].ToString()}</th></tr>";
                                                    _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC1BuyII] set StartTime='{logdate}' where Id='{d2["Id"].ToString()}' and DOCNumberNO='{d_Paper["DOCNumberNO"].ToString()}'");

                                                }
                                            }
                                            _ = db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_WorkingPaper] where ServerId='{_Fun.Config.ServerId}' and Id='{d_Paper["Id"].ToString()}'");
                                        }
                                        break;
                                    case "2"://排程委外
                                        {
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{d_Paper["NeedId"].ToString()}' and SimulationId='{d_Paper["SimulationId"].ToString()}'");
                                            if (dr_tmp != null)
                                            {
                                                if (d_Paper["MFNO"].ToString() != tmp_MFNO)
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' and MFNO='{d_Paper["MFNO"].ToString()}'");
                                                    MailSub2 = $"{_Fun.Config.ServerId} 廠商別:{dr_tmp["MFName"].ToString()} 加工通知單 請依下列到貨日期出貨";
                                                    if (MailBody2 != "")
                                                    {
                                                        MailBody2 = $"{MailBody2}</tbody></table></div>";
                                                        _Fun.Mail_Send(_Fun.Config.SendMonitorMail00.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail00.Split(',')[1].Split(';'), MailSub2, MailBody2, null, false);
                                                    }
                                                    MailBody2 = "<div><table style='background-color:whitesmoke;border:1px solid black;width:100%' rules='all'><thead><tr><th>加工單號</th><th>順序</th><th>加工名稱</th><th>料件編號</th><th>品名</th><th>規格</th><th>數量</th><th>單位</th><th>單價</th><th>到貨需求日</th><th>廠內生產碼</th></tr></thead><tbody>";
                                                    tmp_MFNO = d_Paper["MFNO"].ToString();
                                                }
                                                DataTable dt_tmp2 = db.DB_GetData($"select a.* from SoftNetMainDB.[dbo].[DOC4ProductionII] as a,SoftNetMainDB.[dbo].[DOC4Production] as b where b.ServerId='{_Fun.Config.ServerId}' and b.DOCNumberNO='{d_Paper["DOCNumberNO"].ToString()}' and a.DOCNumberNO=b.DOCNumberNO");
                                                if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                                                {
                                                    foreach (DataRow d2 in dt_tmp2.Rows)
                                                    {
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d_Paper["PartNO"].ToString()}'");
                                                        MailBody2 = $@"{MailBody2}<tr><th>{d2["DOCNumberNO"].ToString()}</th><th>{d2["Id"].ToString()}</th><th>{d2["PartNO"].ToString()}</th><th>{dr_tmp["PartName"].ToString()}</th><th>{dr_tmp["Specification"].ToString()}</th>
                                                                    <th>{d2["QTY"].ToString()}</th><th>{d2["Unit"].ToString()}</th><th>{d2["Price"].ToString()}</th><th>{d2["ArrivalDate"].ToString()}</th><th>{d2["SimulationId"].ToString()}</th></tr>";
                                                        _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC4ProductionII] set StartTime='{logdate}' where Id='{d2["Id"].ToString()}' and DOCNumberNO='{d2["DOCNumberNO"].ToString()}'");
                                                        if (!d_Paper.IsNull("UP_SimulationId") && d_Paper["UP_SimulationId"].ToString() != "")
                                                        {
                                                            dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{d_Paper["NeedId"].ToString()}' and SimulationId='{d_Paper["SimulationId"].ToString()}'");
                                                            _ = db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] set Next_APS_StationNO='{dr_tmp["APS_StationNO"].ToString()}',Next_StationQTY+={d2["QTY"].ToString()} where NeedId='{d_Paper["NeedId"].ToString()}' and SimulationId='{d_Paper["UP_SimulationId"].ToString()}'");
                                                        }
                                                    }
                                                }
                                            }
                                            _ = db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_WorkingPaper] where ServerId='{_Fun.Config.ServerId}' and Id='{d_Paper["Id"].ToString()}'");
                                        }
                                        break;
                                    case "4"://廠內生產
                                        {
                                            string a_pp_Name = d_Paper["Apply_PP_Name"].ToString().Trim();
                                            if (a_pp_Name == "")
                                            {
                                                DataRow dr2 = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[BOM] where Id='{d_Paper["MBOMId"].ToString()}'");
                                                if (dr2 != null) { a_pp_Name = d_Paper["Apply_PP_Name"].ToString().Trim(); }
                                            }
                                            if (a_pp_Name != "")
                                            {
                                                string needDate = DateTime.Now.AddDays(3).ToString("MM/dd/yyyy HH:mm:ss.fff");//###??? 第3天後, 將來改參數化
                                                string needID = _Str.NewId('X');
                                                sql = $@"INSERT INTO SoftNetSYSDB.[dbo].[APS_NeedData] (UpdateTime,KeyA,State,ServerId,Id,IsAdd_SafeQTY,NeedType,NeedSource,NeedDate,PartNO,NeedQTY,BufferTime,CalendarName,BOMId,Apply_PP_Name,CTNO,CTName,FactoryName) VALUES 
                                                        ('{DateTime.Now.AddMinutes(isARGs10_offset).ToString("yyyy-MM-dd HH:mm:ss")}','0,0,0,0,0,0,0,0,1,1,0,0,0,0,0,0,0,0,0,1','1','{_Fun.Config.ServerId}','{needID}','1','5','','{needDate}','{d_Paper["PartNO"].ToString()}',{d_Paper["NeedQTY"].ToString()},48,'{_Fun.Config.DefaultCalendarName}','{d_Paper["MBOMId"].ToString()}','{d_Paper["Apply_PP_Name"].ToString()}','','補安全存貨','{_Fun.Config.DefaultFactoryName}')";
                                                if (db.DB_SetData(sql))
                                                {
                                                    needID_data.Add(needID);
                                                    MailSub4 = $"{_Fun.Config.ServerId} 補足安全存貨工單, 已自動轉廠內生產排程, 內容如下.";
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d_Paper["PartNO"].ToString()}'");
                                                    MailBody4 = $"{MailBody4}<p>計畫碼:{needID}  料號:{d_Paper["PartNO"].ToString()} 品名:{dr_tmp["PartName"].ToString()} 規格:{dr_tmp["Specification"].ToString()} 數量:{d_Paper["NeedQTY"].ToString()}</p>";
                                                    _ = db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_WorkingPaper] where ServerId='{_Fun.Config.ServerId}' and Id='{d_Paper["Id"].ToString()}'");
                                                }
                                            }
                                        }
                                        break;
                                }
                            }
                            bool isw = false;
                            if (MailBody4 != "") { _Fun.Mail_Send(_Fun.Config.SendMonitorMail00.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail00.Split(',')[1].Split(';'), MailSub4, MailBody4, null, false); isw = true; }
                            if (MailBody3 != "") { _Fun.Mail_Send(_Fun.Config.SendMonitorMail00.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail00.Split(',')[1].Split(';'), MailSub3, MailBody3, null, false); isw = true; }
                            if (MailBody2 != "") { _Fun.Mail_Send(_Fun.Config.SendMonitorMail00.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail00.Split(',')[1].Split(';'), MailSub2, MailBody2, null, false); isw = true; }
                            if (isw)
                            {
                                _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime) VALUES ('工作底稿主動干涉, 作業明細已Mail發出','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                            }
                            if (needID_data.Count > 0)
                            {
                                _SFC_Common.RunSetSimulation(args, "", needID_data, '5');
                            }
                            #region 通知網頁更新
                            //try
                            //{
                            //lock (WebSocketServiceOJB.lock__WebSocketList)
                            //{
                            //    foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                            //    {
                            //        if (r.Key != null && r.Value.socket != null)
                            //        {
                            //            WebSocketServiceOJB.Send(r.Value.socket, $"HasWeb_Id_Change,STView2Work_PageReload,{sno}");
                            //        }
                            //    }
                            //}
                            //}
                            //catch (Exception ex)
                            //{
                            //    System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs 工站:{sno} 發HasWeb_Id_Changeg失敗 {ex.Message} {ex.StackTrace}", true);
                            //}
                            #endregion
                        }
                        #endregion
                        _Fun._a02 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 檢查 生產用領料,入庫,退料單據 的用量 IsOK=0是否可以自動改IsOK=1  先將 and b.ArrivalDate<'{logdate}'從sql移出
                        sql = $@"SELECT b.* FROM SoftNetMainDB.[dbo].[DOC3stock] as a
                            join SoftNetMainDB.[dbo].[DOC3stockII] as b on b.DOCNumberNO=a.DOCNumberNO
                            where a.ServerId='{_Fun.Config.ServerId}' and b.IsOK='0'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            int IsOKQTY = 0;

                            string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                DataRow dr_tmp_Note = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where ((Class!='4' and Class!='5') or NoStation='1') and Next_StationQTY>=NeedQTY and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}' and PartNO='{d["PartNO"].ToString()}'");
                                if (dr_tmp_Note != null)
                                {
                                    IsOKQTY = 0;
                                    int next_StationQTY = int.Parse(dr_tmp_Note["Next_StationQTY"].ToString());
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT sum(QTY) as OKQTY FROM SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}' and PartNO='{d["PartNO"].ToString()}' and IsOK='1'");
                                    if (dr_tmp != null && !dr_tmp.IsNull("OKQTY")) { IsOKQTY = int.Parse(dr_tmp["OKQTY"].ToString()); }
                                    if ((next_StationQTY - IsOKQTY) > 0)
                                    {
                                        int ct = 0;
                                        IsOKQTY = next_StationQTY - IsOKQTY;
                                        if (int.Parse(d["QTY"].ToString()) > IsOKQTY)
                                        {
                                            #region 確認 TotalStock檔 有對應資料
                                            if (d["IN_StoreNO"].ToString() != "")
                                            {
                                                if (d["IN_StoreSpacesNO"].ToString() != "")
                                                { dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' and StoreNO='{d["IN_StoreNO"].ToString()}' and StoreSpacesNO='{d["IN_StoreSpacesNO"].ToString()}'"); }
                                                else
                                                { dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where  ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' and StoreNO='{d["IN_StoreNO"].ToString()}'"); }
                                                if (dr_tmp == null)
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and StoreNO='{d["IN_StoreNO"].ToString()}'");
                                                    if (dr_tmp != null)
                                                    {
                                                        _ = db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (Class,ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{dr_tmp["Class"].ToString()}','{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{d["IN_StoreNO"].ToString()}','{d["IN_StoreSpacesNO"].ToString()}','{d["PartNO"].ToString()}',0)");
                                                    }
                                                }
                                            }
                                            #endregion
                                            string startTime = "NULL";
                                            if (!d.IsNull("StartTime"))
                                            {
                                                startTime = $"'{Convert.ToDateTime(d["ArrivalDate"]).ToString("yyyy/MM/dd HH:mm:ss.fff")}'";
                                                ct = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(d["StartTime"]), DateTime.Now);
                                            }
                                            _ = db.DB_SetData($@"INSERT INTO [dbo].[DOC3stockII] (ServerId,[Id],[DOCNumberNO],[PartNO],[Price],[Unit],[QTY],[Remark],[SimulationId],[IsOK],[IN_StoreNO],[IN_StoreSpacesNO],[OUT_StoreNO],[OUT_StoreSpacesNO],StartTime,ArrivalDate,CT) VALUES 
                                                        ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{d["DOCNumberNO"].ToString()}','{d["PartNO"].ToString()}',0,'PCS',{IsOKQTY},'{d["Remark"].ToString()}','{d["SimulationId"].ToString()}','1','{d["IN_StoreNO"].ToString()}','{d["IN_StoreSpacesNO"].ToString()}','{d["OUT_StoreNO"].ToString()}','{d["OUT_StoreSpacesNO"].ToString()}',{startTime},'{Convert.ToDateTime(d["ArrivalDate"]).ToString("yyyy/MM/dd HH:mm:ss.fff")}',{ct.ToString()})");
                                            _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set QTY={(int.Parse(d["QTY"].ToString()) - IsOKQTY).ToString()} where Id='{d["Id"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}'");
                                        }
                                        else
                                        {
                                            string writeSQL = "";
                                            DataRow dr_tmp_DOC3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}' and PartNO='{d["PartNO"].ToString()}' and IsOK='1' order by StartTime desc");
                                            if (dr_tmp_DOC3 != null && !dr_tmp_DOC3.IsNull("StartTime")) { ct = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_tmp_DOC3["StartTime"]), DateTime.Now); }
                                            else { writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                                            _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',CT={ct},EndTime='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}'{writeSQL} where Id='{d["Id"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}' and IsOK='0'");
                                            IsOKQTY = int.Parse(d["QTY"].ToString());
                                        }
                                        string needID = "";
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}'");
                                        if (dr_tmp != null) { needID = dr_tmp_Note["NeedId"].ToString(); }
                                        _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId,SimulationId) VALUES ('將生產用領料/入庫/退料的單據主動確認,並異動存貨帳的數量 單據碼:{d["DOCNumberNO"].ToString()} 料號:{d["PartNO"].ToString()} 數量:{IsOKQTY.ToString()}','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy -MM-dd HH:mm:ss.fff")}','{needID}','{d["SimulationId"].ToString()}')");
                                    }
                                    if (IsOKQTY > 0)
                                    {
                                        #region 寫入庫存
                                        if (d.IsNull("OUT_StoreNO") || d["OUT_StoreNO"].ToString().Trim() == "")
                                        { _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={IsOKQTY} where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' and StoreNO='{d["IN_StoreNO"].ToString()}' and StoreSpacesNO='{d["IN_StoreSpacesNO"].ToString()}'"); }
                                        else { _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={IsOKQTY} where ServerId='{_Fun.Config.ServerId}' and PartNO='{d["PartNO"].ToString()}' and StoreNO='{d["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{d["OUT_StoreSpacesNO"].ToString()}'"); }
                                        #endregion

                                        #region 計算單據CT,平均,有效
                                        int typeTotalTime = 0;
                                        string StartTime = "";
                                        if (!d.IsNull("StartTime")) { typeTotalTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(d["StartTime"].ToString()), DateTime.Now); }
                                        else { StartTime = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                                        _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime}{StartTime} where Id='{d["Id"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}' and IsOK='0'");
                                        if (typeTotalTime > 0)
                                        {
                                            string partNO = d["PartNO"].ToString();
                                            string pp_Name = "";
                                            string E_stationNO = "";
                                            string indexSN = "";
                                            if (d["SimulationId"].ToString() != "")
                                            {
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}'");
                                                pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                                                if (!dr_tmp.IsNull("Source_StationNO") && dr_tmp["Source_StationNO"].ToString() != "")
                                                { E_stationNO = dr_tmp["Source_StationNO"].ToString(); indexSN = dr_tmp["Source_StationNO_IndexSN"].ToString(); }
                                                else { E_stationNO = dr_tmp["Apply_StationNO"].ToString(); indexSN = dr_tmp["IndexSN"].ToString(); }
                                            }
                                            DataTable dt_Efficient = db.DB_GetData($@"select {top_flag} CT from SoftNetMainDB.[dbo].[DOC3stockII] where SUBSTRING(Id,2,2)='{_Fun.Config.ServerId}' and SUBSTRING(DOCNumberNO,1,4)='{d["DOCNumberNO"].ToString().Substring(0, 4)}' and PartNO='{d["PartNO"].ToString()}' and CT>0");
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
                                                _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, pp_Name, pp_Name, indexSN, partNO, partNO, d["DOCNumberNO"].ToString().Substring(0, 4));
                                            }
                                        }
                                        #endregion
                                    }
                                }
                            }
                        }
                        #endregion
                        _Fun._a03 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 監測APS_Simulation, 生產備料是否無作為無單據發起)  05
                        string mailSub = $"{_Fun.Config.ServerId} 警示 工作站生產物,無單據發起.";
                        string mailBody05 = "";
                        if (db.DB_GetQueryCount($"select TOP 1 * from SoftNetSYSDB.[dbo].APS_Simulation_ErrorData where ServerId='{_Fun.Config.ServerId}' and ErrorType='05' and ActionType=''") > 0)
                        {
                            #region 已處理解除
                            sql = $@"SELECT a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                                        join SoftNetSYSDB.[dbo].[APS_Simulation] as b on b.SimulationId=a.SimulationId and b.DOCNumberNO!=''
                                        where a.ServerId='{_Fun.Config.ServerId}' and a.ActionType='' and a.ErrorType='05'";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                foreach (DataRow d in dt_tmp.Rows)
                                {
                                    if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='0',ActionLogDate='{logdate}' where Id='{d["Id"].ToString()}'"))
                                    {
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,05,{d["SimulationId"].ToString()},0");
                                        //#region 通知網頁更新
                                        ////var webSocketService = (SNWebSocketService)_Fun.DiBox.GetService(typeof(SNWebSocketService));
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,05,{d["SimulationId"].ToString()},0");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                    }
                                }
                            }
                            #endregion
                        }
                        sql = $@"SELECT a.*,b.CalendarName  FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                    join SoftNetSYSDB.[dbo].[APS_NeedData] as b on b.Id=a.NeedId and b.State='6' and b.ServerId='{_Fun.Config.ServerId}'
                                    where a.StartDate<'{logdate}' and ((a.Class!='4' and a.Class!='5') or a.Source_StationNO is null) and a.Source_StationNO!='{_Fun.Config.OutPackStationName}' and (a.DOCNumberNO is null or a.DOCNumberNO='')";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            #region 成立 及 已逾時解除
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and ErrorType='05'");
                                if (dr_tmp == null)
                                {
                                    // Use parameterized query to avoid string interpolation and SQL injection
                                    var paramDict = new Dictionary<string, object>
                                    {
                                        { "Id", _Str.NewId('E') },
                                        { "ServerId", _Fun.Config.ServerId },
                                        { "SimulationId", d["SimulationId"].ToString() },
                                        { "ErrorType", "05" },
                                        { "LogDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") },
                                        { "NeedId", d["NeedId"].ToString() }
                                    };
                                    if (db.DB_SetDataByParams("INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId) VALUES (@Id,@ServerId,@SimulationId,@ErrorType,@LogDate,@NeedId)", paramDict))
                                    {
                                        if (_Fun.Config.SendMonitorMail05 != "")
                                        {
                                            mailBody05 = $"{mailBody05}<p>排程碼:{d["NeedId"].ToString()} {d["SimulationId"].ToString()} 工站:{d["StationNO"].ToString()}</p>";
                                        }
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,05,{d["SimulationId"].ToString()},");
                                        //#region 通知網頁更新
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,05,{d["SimulationId"].ToString()},");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                    }
                                }
                                /* 只要有報工,系統就會開單, 此處就不用檢查
                                else
                                {
                                    if (dr_tmp["ActionType"].ToString() == "")
                                    {
                                        string eID = dr_tmp["Id"].ToString();
                                        #region 已逾時解除 (條件=工站已生產超過的數量)
                                        sql = $"SELECT SimulationId FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{d["NeedId"].ToString()}' and PartNO='{d["PartNO"].ToString()}' and Source_StationNO='{d["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN='{d["IndexSN"].ToString()}' and Class='{d["Master_Class"].ToString()}'";
                                        dr_tmp = db.DB_GetFirstDataByDataRow(sql);
                                        if (dr_tmp != null)
                                        {
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{dr_tmp["SimulationId"].ToString()}' and Detail_QTY>=NeedQTY");
                                            if (dr_tmp != null)
                                            {
                                                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='1',ActionLogDate='{logdate}' where Id='{eID}'"))
                                                {
                                                    #region 通知網頁更新
                                                    //var webSocketService = (SNWebSocketService)_Fun.DiBox.GetService(typeof(SNWebSocketService));
                                                    if (WebSocketServiceOJB != null)
                                                    {
                                                        lock (WebSocketServiceOJB.lock__WebSocketList)
                                                        {
                                                            foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                                            {
                                                                if (r.Key != null && r.Value.socket != null)
                                                                {
                                                                    WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,05,{d["SimulationId"].ToString()},1");
                                                                }
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                }
                                */
                            }
                            if (mailBody05 != "")
                            { _Fun.Mail_Send(_Fun.Config.SendMonitorMail05.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail05.Split(',')[1].Split(';'), mailSub, mailBody05, null, false); }

                            #endregion
                        }
                        #endregion
                        _Fun._a04 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 監測生產排程計畫Keep(保留)量是否足夠庫存量 TotalStockII  01(原物料) 02(自製件) 
                        string mailSub = $"{_Fun.Config.ServerId} 警示 排程計畫Keep保留量不足夠庫存量";
                        string mailBody01 = "";
                        string mailBody02 = "";
                        if (db.DB_GetQueryCount($"select top 1 * from SoftNetSYSDB.[dbo].APS_Simulation_ErrorData where ServerId='{_Fun.Config.ServerId}' and (ErrorType='01' or ErrorType='02') and ActionType=''") > 0)
                        {
                            #region 解除
                            sql = $@"select a.Id,b.SimulationId,c.Id as eId,c.ErrorType,((b.KeepQTY+b.OverQTY)-a.QTY) as qty FROM SoftNetMainDB.[dbo].[TotalStock] as a
                                        right join SoftNetMainDB.[dbo].[TotalStockII] as b on a.Id=b.Id 
                                        right join SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as c on c.ServerId='{_Fun.Config.ServerId}' and c.ActionType='' and (c.ErrorType='01' or c.ErrorType='02')
                                        where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and (b.KeepQTY+b.OverQTY)<=a.QTY";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                foreach (DataRow d in dt_tmp.Rows)
                                {
                                    if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='0',ActionLogDate='{logdate}' where Id='{d["eId"].ToString()}'"))
                                    {
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,{d["ErrorType"].ToString()},{d["SimulationId"].ToString()},0");
                                        //#region 通知網頁更新
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,{d["ErrorType"].ToString()},{d["SimulationId"].ToString()},0");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                    }
                                }
                            }
                            #endregion
                        }
                        //###??? ㄝ有可能有委外件有存貨 因為sql 有加OutPackType!='1'
                        sql = $@"select a.Id,a.PartNO,e.PartName,e.Specification,b.SimulationId,c.NeedId,c.Class,((b.KeepQTY+b.OverQTY)-a.QTY) as qty FROM SoftNetMainDB.[dbo].[TotalStock] as a
                                right join SoftNetMainDB.[dbo].[TotalStockII] as b on a.Id=b.Id 
                                left join SoftNetSYSDB.[dbo].[APS_Simulation] as c on b.SimulationId=c.SimulationId and c.OutPackType!='1' 
                                left join SoftNetSYSDB.[dbo].[APS_NeedData] as d on c.NeedId=d.Id
                                join SoftNetMainDB.[dbo].[Material] as e on e.PartNO=a.PartNO
                                where a.Class!='虛擬倉' and a.ServerId='{_Fun.Config.ServerId}' and (b.KeepQTY+b.OverQTY)>a.QTY and d.State='6'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (d["Class"].ToString() != "4" && d["Class"].ToString() != "5")
                                {
                                    #region 01(原物料)
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and ErrorType='01'");
                                    if (dr_tmp == null)
                                    {
                                        if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d["SimulationId"].ToString()}','01','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d["NeedId"].ToString()}')"))
                                        {
                                            if (_Fun.Config.SendMonitorMail01 != "")
                                            {
                                                mailBody01 = $"{mailBody01}<p>排程碼:{d["NeedId"].ToString()} {d["SimulationId"].ToString()} 料號:{d["PartNO"].ToString()} {d["PartName"].ToString()} {d["Specification"].ToString()} 不足數量:{d["qty"].ToString()}</p>";
                                            }
                                            _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,01,{d["SimulationId"].ToString()},");
                                            //#region 通知網頁更新
                                            //if (WebSocketServiceOJB != null)
                                            //{
                                            //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                            //    {
                                            //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                            //        {
                                            //            if (r.Key != null && r.Value.socket != null)
                                            //            {
                                            //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,01,{d["SimulationId"].ToString()},");
                                            //            }
                                            //        }
                                            //    }
                                            //}
                                            //#endregion
                                        }
                                    }
                                    else
                                    {
                                        if (dr_tmp["ActionType"].ToString() == "")
                                        {
                                            string eID = dr_tmp["Id"].ToString();
                                            #region 已逾時解除 (條件=工站已生產超過的數量)
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d["SimulationId"].ToString()}' and Detail_QTY>=NeedQTY");
                                            if (dr_tmp != null)
                                            {
                                                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='1',ActionLogDate='{logdate}' where Id='{eID}'"))
                                                {
                                                    _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,01,{d["SimulationId"].ToString()},1");
                                                    //#region 通知網頁更新
                                                    //if (WebSocketServiceOJB != null)
                                                    //{
                                                    //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                                    //    {
                                                    //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                                    //        {
                                                    //            if (r.Key != null && r.Value.socket != null)
                                                    //            {
                                                    //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,01,{d["SimulationId"].ToString()},1");
                                                    //            }
                                                    //        }
                                                    //    }
                                                    //}
                                                    //#endregion
                                                }
                                            }
                                            #endregion
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select Id,SimulationId,ActionType from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and ErrorType='02'");
                                    if (dr_tmp == null)
                                    {
                                        if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d["SimulationId"].ToString()}','02','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d["NeedId"].ToString()}')"))
                                        {
                                            if (_Fun.Config.SendMonitorMail02 != "")
                                            {
                                                mailBody02 = $"{mailBody02}<p>排程碼:{d["NeedId"].ToString()} {d["SimulationId"].ToString()} 料號:{d["PartNO"].ToString()} {d["PartName"].ToString()} {d["Specification"].ToString()} 不足數量:{d["qty"].ToString()}</p>";
                                            }
                                            _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,02,{d["SimulationId"].ToString()},");
                                            //#region 通知網頁更新
                                            //if (WebSocketServiceOJB != null)
                                            //{
                                            //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                            //    {
                                            //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                            //        {
                                            //            if (r.Key != null && r.Value.socket != null)
                                            //            {
                                            //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,02,{d["SimulationId"].ToString()},");
                                            //            }
                                            //        }
                                            //    }
                                            //}
                                            //#endregion
                                        }
                                    }
                                    else
                                    {
                                        if (dr_tmp["ActionType"].ToString() == "")
                                        {
                                            string eID = dr_tmp["Id"].ToString();
                                            #region 已逾時解除 (條件=工站已生產超過的數量)
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where SimulationId='{d["SimulationId"].ToString()}' and Detail_QTY>=NeedQTY");
                                            if (dr_tmp != null)
                                            {
                                                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='1',ActionLogDate='{logdate}' where Id='{eID}' and ErrorType='02'"))
                                                {
                                                    _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,02,{d["SimulationId"].ToString()},1");

                                                    //#region 通知網頁更新
                                                    //if (WebSocketServiceOJB != null)
                                                    //{
                                                    //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                                    //    {
                                                    //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                                    //        {
                                                    //            if (r.Key != null && r.Value.socket != null)
                                                    //            {
                                                    //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,02,{d["SimulationId"].ToString()},1");
                                                    //            }
                                                    //        }
                                                    //    }
                                                    //}
                                                    //#endregion
                                                }
                                            }
                                            #endregion
                                        }
                                    }
                                }
                            }
                            if (mailBody01 != "")
                            { _Fun.Mail_Send(_Fun.Config.SendMonitorMail01.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail01.Split(',')[1].Split(';'), mailSub, mailBody01, null, false); }
                            if (mailBody02 != "")
                            { _Fun.Mail_Send(_Fun.Config.SendMonitorMail02.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail02.Split(',')[1].Split(';'), mailSub, mailBody02, null, false); }
                        }
                        #endregion
                        _Fun._a05 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 監測監測工作站生產物料應領用,而未領用  04 
                        string mailSub = $"{_Fun.Config.ServerId} 警示 工作站生產物料應領料,而未領料.";
                        string mailBody04 = "";
                        if (db.DB_GetQueryCount($"select * from SoftNetSYSDB.[dbo].APS_Simulation_ErrorData where ErrorType='04' and ActionType=''") > 0)
                        {
                            #region 解除
                            sql = $@"select a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                                        right join SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as b on a.SimulationId=b.SimulationId
                                        where a.ServerId='{_Fun.Config.ServerId}' and a.ErrorType='04' and a.ActionType='' and b.DOCNumberNO!=''";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                foreach (DataRow d in dt_tmp.Rows)
                                {
                                    if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='0',ActionLogDate='{logdate}' where Id='{d["Id"].ToString()}'"))
                                    {
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,{d["ErrorType"].ToString()},{d["SimulationId"].ToString()},0");
                                        //#region 通知網頁更新
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,{d["ErrorType"].ToString()},{d["SimulationId"].ToString()},0");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                    }
                                }
                            }
                            #endregion
                        }

                        //###??? 有可能會領生產件 Class='4' and a.Class='5'
                        sql = $@"select a.* from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a
                                join SoftNetMainDB.[dbo].[DOC3stockII] as b on a.SimulationId=b.SimulationId
                                join SoftNetSYSDB.[dbo].[APS_NeedData] as c on c.State='6' and a.NeedId=c.Id and c.ServerId='{_Fun.Config.ServerId}' 
                                where a.CalendarDate>='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' and a.DOCNumberNO='' and ((a.Class!='4' and a.Class!='5') or a.NoStation='1')";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select SimulationId,ActionType from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["APS_StationNO"].ToString()}' and ErrorType='04'");
                                if (dr_tmp == null)
                                {
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d["SimulationId"].ToString()}','04','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d["NeedId"].ToString()}','{d["APS_StationNO"].ToString()}')"))
                                    {
                                        if (_Fun.Config.SendMonitorMail04 != "")
                                        {
                                            mailBody04 = $"{mailBody04}<p>排程碼:{d["NeedId"].ToString()} {d["SimulationId"].ToString()} 工站:{d["StationNO"].ToString()}</p>";
                                        }
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,'04',{d["SimulationId"].ToString()}");
                                        //#region 通知網頁更新
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                //###???errorType 暫時寫死
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,'04',{d["SimulationId"].ToString()}");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                    }
                                }
                                else
                                {
                                    if (dr_tmp["ActionType"].ToString() == "")
                                    {
                                        string eID = dr_tmp["Id"].ToString();
                                        #region 已逾時解除 (條件=DOC3stockII有 SimulationId紀錄)
                                        sql = $"SELECT SimulationId FROM SoftNetMainDB.[dbo].[DOC3stockII] where SimulationId='{d["SimulationId"].ToString()}'";
                                        dr_tmp = db.DB_GetFirstDataByDataRow(sql);
                                        if (dr_tmp != null)
                                        {
                                            if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='1',ActionLogDate='{logdate}' where Id='{eID}'"))
                                            {
                                                _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,04,{d["SimulationId"].ToString()},1");
                                                //#region 通知網頁更新
                                                //if (WebSocketServiceOJB != null)
                                                //{
                                                //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                                //    {
                                                //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                                //        {
                                                //            if (r.Key != null && r.Value.socket != null)
                                                //            {
                                                //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,04,{d["SimulationId"].ToString()},1");
                                                //            }
                                                //        }
                                                //    }
                                                //}
                                                //#endregion
                                            }
                                        }
                                        #endregion
                                    }
                                }
                            }
                            if (mailBody04 != "")
                            { _Fun.Mail_Send(_Fun.Config.SendMonitorMail04.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail04.Split(',')[1].Split(';'), mailSub, mailBody04, null, false); }
                        }
                        #endregion
                        _Fun._a06 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 監測所有單據類是否如預期從Wait狀態,改成OK狀態 IsOK=1  06 07 08 09 10
                        string mailSub = $"{_Fun.Config.ServerId} 警示 單據未如預期從Wait狀態,改成OK狀態.";
                        string mailBody06 = "";
                        string mailBody07 = "";
                        string mailBody08 = "";
                        string mailBody09 = "";
                        string mailBody10 = "";
                        List<string> docTMP = new List<string>();
                        if (db.DB_GetQueryCount($"select * from SoftNetSYSDB.[dbo].APS_Simulation_ErrorData where ServerId='{_Fun.Config.ServerId}' and (ErrorType='06' or ErrorType='07' or ErrorType='08' or ErrorType='09' or ErrorType='10') and ActionType=''") > 0)
                        {
                            #region IsOK='1' 自動解除
                            DataTable dt_tmpII = new DataTable();
                            sql = $@"select a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                                    join SoftNetMainDB.[dbo].[DOC1BuyII] as b on a.SimulationId=b.SimulationId and a.DOCNumberNO=b.DOCNumberNO and a.ErrorKey=b.Id and b.IsOK='1'
                                    where a.ServerId='{_Fun.Config.ServerId}' and a.ErrorType='06' and a.ActionType=''";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0) { dt_tmpII.Merge(dt_tmp); }
                            sql = $@"select a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                                    join SoftNetMainDB.[dbo].[DOC2SalesII] as b on a.SimulationId=b.SimulationId and a.DOCNumberNO=b.DOCNumberNO and a.ErrorKey=b.Id and b.IsOK='1'
                                    where a.ServerId='{_Fun.Config.ServerId}' and a.ErrorType='07' and a.ActionType=''";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0) { dt_tmpII.Merge(dt_tmp); }
                            sql = $@"select a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                                    join SoftNetMainDB.[dbo].[DOC3stockII] as b on a.SimulationId=b.SimulationId and a.DOCNumberNO=b.DOCNumberNO and a.ErrorKey=b.Id and b.IsOK='1'
                                    where a.ServerId='{_Fun.Config.ServerId}' and a.ErrorType='08' and a.ActionType=''";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0) { dt_tmpII.Merge(dt_tmp); }
                            sql = $@"select a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                                    join SoftNetMainDB.[dbo].[DOC4ProductionII] as b on a.SimulationId=b.SimulationId and a.DOCNumberNO=b.DOCNumberNO and a.ErrorKey=b.Id and b.IsOK='1'
                                    where a.ServerId='{_Fun.Config.ServerId}' and a.ErrorType='09' and a.ActionType=''";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0) { dt_tmpII.Merge(dt_tmp); }
                            sql = $@"select a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                                    join SoftNetMainDB.[dbo].[DOC5OUTII] as b on a.SimulationId=b.SimulationId and a.DOCNumberNO=b.DOCNumberNO and a.ErrorKey=b.Id and b.IsOK='1'
                                    where a.ServerId='{_Fun.Config.ServerId}' and a.ErrorType='10' and a.ActionType=''";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0) { dt_tmpII.Merge(dt_tmp); }
                            if (dt_tmpII != null && dt_tmpII.Rows.Count > 0)
                            {
                                foreach (DataRow d in dt_tmpII.Rows)
                                {
                                    if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='0',ActionLogDate='{logdate}' where Id='{d["Id"].ToString()}'"))
                                    {
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,{d["ErrorType"].ToString()},{d["SimulationId"].ToString()},0");
                                        //#region 通知網頁更新
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,{d["ErrorType"].ToString()},{d["SimulationId"].ToString()},0");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                    }
                                }
                            }
                            #endregion

                            #region 製程下一站已完成對應數量 主動干涉 另 定時通知
                            sql = "";
                            dt_tmp = db.DB_GetData($@"select a.NeedId,a.SimulationId,a.DOCNumberNO,a.ErrorKey,a.ErrorType,b.PartSN,b.Master_PartNO,b.Apply_PP_Name,b.Apply_StationNO,b.IndexSN from SoftNetSYSDB.[dbo].APS_Simulation_ErrorData as a
                                                            join SoftNetSYSDB.[dbo].[APS_Simulation] as b on a.SimulationId=b.SimulationId
                                                            where a.ServerId='{_Fun.Config.ServerId}' and a.ErrorType='08' and ActionType='' and a.SimulationId!=''
                                                            order by a.NeedId,b.PartSN desc");
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                bool isrun = false;
                                DataRow APS_PartNOTimeNote = null;
                                foreach (DataRow d in dt_tmp.Rows)
                                {
                                    isrun = false;
                                    //本階的工站
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{d["NeedId"].ToString()}' and PartNO='{d["Master_PartNO"].ToString()}' and Source_StationNO='{d["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN={d["IndexSN"].ToString()} and Source_StationNO is not null");
                                    if (dr_tmp != null)
                                    {
                                        //下一階的工站
                                        //dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{dr_tmp["NeedId"].ToString()}' and PartNO='{dr_tmp["Master_PartNO"].ToString()}' and Source_StationNO='{dr_tmp["Apply_StationNO"].ToString()}' and Source_StationNO_IndexSN={dr_tmp["IndexSN"].ToString()} and (Class='4' or Class='5') and Source_StationNO is not null");
                                        APS_PartNOTimeNote = db.DB_GetFirstDataByDataRow($@"select a.SimulationId,b.NeedQTY,(b.Detail_QTY+b.Detail_Fail_QTY) as okQTY from SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                                        join SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as b on a.SimulationId=b.SimulationId
                                                        where a.NeedId='{dr_tmp["NeedId"].ToString()}' and a.PartNO='{dr_tmp["Master_PartNO"].ToString()}' and a.Source_StationNO='{dr_tmp["Apply_StationNO"].ToString()}' and a.Source_StationNO_IndexSN={dr_tmp["IndexSN"].ToString()} and a.Source_StationNO is not null");
                                        if (APS_PartNOTimeNote != null && int.Parse(APS_PartNOTimeNote["okQTY"].ToString()) >= int.Parse(APS_PartNOTimeNote["NeedQTY"].ToString()))
                                        {
                                            if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='0',ActionLogDate='{logdate}' where DOCNumberNO='{d["DOCNumberNO"].ToString()}' and ErrorKey='{d["ErrorKey"].ToString()}' and ErrorType='08'"))
                                            {
                                                isrun = true;
                                                #region 寫入庫存
                                                string top_flag = $" TOP {_Fun.Config.AdminKey03} ";
                                                DataRow dr_DOC3stockII = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' and Id='{d["ErrorKey"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}'");
                                                if (dr_DOC3stockII != null)
                                                {
                                                    if (dr_DOC3stockII.IsNull("OUT_StoreNO") || dr_DOC3stockII["OUT_StoreNO"].ToString().Trim() == "")
                                                    { _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY+={dr_DOC3stockII["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["IN_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["IN_StoreSpacesNO"].ToString()}'"); }
                                                    else { _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[TotalStock] set QTY-={dr_DOC3stockII["QTY"].ToString()} where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_DOC3stockII["PartNO"].ToString()}' and StoreNO='{dr_DOC3stockII["OUT_StoreNO"].ToString()}' and StoreSpacesNO='{dr_DOC3stockII["OUT_StoreSpacesNO"].ToString()}'"); }

                                                    #region 計算單據CT,平均,有效, 寫SFC_StationProjectDetail
                                                    int typeTotalTime = 0;
                                                    string writeSQL = "";
                                                    if (!dr_DOC3stockII.IsNull("StartTime")) { typeTotalTime = _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, Convert.ToDateTime(dr_DOC3stockII["StartTime"]), DateTime.Now); }
                                                    else { writeSQL = $",StartTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}'"; }
                                                    _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[DOC3stockII] set IsOK='1',EndTime='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}',CT={typeTotalTime}{writeSQL} where Id='{d["ErrorKey"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}' and IsOK='0'");

                                                    string partNO = dr_DOC3stockII["PartNO"].ToString();
                                                    string pp_Name = "";
                                                    string E_stationNO = "";
                                                    string indexSN = "0";
                                                    if (dr_DOC3stockII["SimulationId"].ToString() != "")
                                                    {
                                                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr_DOC3stockII["SimulationId"].ToString()}'");
                                                        pp_Name = dr_tmp["Apply_PP_Name"].ToString();
                                                        if (!dr_tmp.IsNull("Source_StationNO") && dr_tmp["Source_StationNO"].ToString() != "")
                                                        { E_stationNO = dr_tmp["Source_StationNO"].ToString(); indexSN = dr_tmp["Source_StationNO_IndexSN"].ToString(); }
                                                        else { E_stationNO = dr_tmp["Apply_StationNO"].ToString(); indexSN = dr_tmp["IndexSN"].ToString(); }
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
                                                        _SFC_Common.SfcTimerloopthread_Tick_Efficient(db, allCT, E_stationNO, pp_Name, pp_Name, indexSN, partNO, partNO, dr_DOC3stockII["DOCNumberNO"].ToString().Substring(0, 4));
                                                    }
                                                    #endregion
                                                }
                                                #endregion
                                            }
                                        }
                                        if (isrun)
                                        {
                                            _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,{d["ErrorType"].ToString()},{d["SimulationId"].ToString()},0");
                                            //#region 通知網頁更新
                                            //if (WebSocketServiceOJB != null)
                                            //{
                                            //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                            //    {
                                            //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                            //        {
                                            //            if (r.Key != null && r.Value.socket != null)
                                            //            {
                                            //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,{d["ErrorType"].ToString()},{d["SimulationId"].ToString()},0");
                                            //            }
                                            //        }
                                            //    }
                                            //}
                                            //#endregion
                                        }
                                    }
                                }
                            }
                            #endregion
                        }

                        #region 1進貨類 06
                        sql = $@"select a.SimulationId,a.Id,a.DOCNumberNO,c.DOCName FROM SoftNetMainDB.[dbo].[DOC1BuyII] as a
                                    right join  SoftNetMainDB.[dbo].[DOC1Buy] as b on a.DOCNumberNO=b.DOCNumberNO
                                    join  SoftNetMainDB.[dbo].[DOCRole] as c on c.DOCNO=b.DOCNO
                                    where b.ServerId='{_Fun.Config.ServerId}' and b.FlowStatus='Y' and a.ArrivalDate<='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' and a.IsOK='0'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null)
                        {

                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (db.DB_GetQueryCount($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}' and ErrorType='06' and ErrorKey='{d["Id"].ToString()}' and ActionType=''") <= 0)
                                {
                                    DataRow d01 = db.DB_GetFirstDataByDataRow($"select NeedId FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}'");
                                    if (d01 != null && db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,DOCNumberNO,ErrorType,ErrorKey,LogDate,NeedId) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d["SimulationId"].ToString()}','{d["DOCNumberNO"].ToString()}','06','{d["Id"].ToString()}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d01["NeedId"].ToString()}')"))
                                    {
                                        if (_Fun.Config.SendMonitorMail06 != "")
                                        {
                                            if (!docTMP.Contains(d["DOCNumberNO"].ToString()))
                                            {
                                                docTMP.Add(d["DOCNumberNO"].ToString());
                                                mailBody06 = $"{mailBody06}<p>進貨類 單據編號:{d["DOCNumberNO"].ToString()} {d["DOCName"].ToString()}</p>";
                                            }
                                        }
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,'06',{d["SimulationId"].ToString()},");
                                        //#region 通知網頁更新
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                //###???errorType 暫時寫死
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,'06',{d["SimulationId"].ToString()},");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                    }
                                }
                            }
                            if (mailBody06 != "")
                            { _Fun.Mail_Send(_Fun.Config.SendMonitorMail06.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail06.Split(',')[1].Split(';'), mailSub, mailBody06, null, false); }
                        }
                        #endregion
                        #region 2銷貨類 07
                        sql = $@"select a.SimulationId,a.Id,a.DOCNumberNO,c.DOCName FROM SoftNetMainDB.[dbo].[DOC2SalesII] as a
                                    right join  SoftNetMainDB.[dbo].[DOC2Sales] as b on a.DOCNumberNO=b.DOCNumberNO
                                    join  SoftNetMainDB.[dbo].[DOCRole] as c on c.DOCNO=b.DOCNO
                                    where b.ServerId='{_Fun.Config.ServerId}' and b.FlowStatus='Y' and a.ArrivalDate<='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' and a.IsOK='0'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null)
                        {
                            docTMP.Clear();
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (db.DB_GetQueryCount($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}' and ErrorType='07' and ErrorKey='{d["Id"].ToString()}' and ActionType=''") <= 0)
                                {
                                    DataRow d01 = db.DB_GetFirstDataByDataRow($"select NeedId FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}'");
                                    //###???errorType 暫時寫死
                                    if (d01 != null && db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,DOCNumberNO,ErrorType,ErrorKey,LogDate,NeedId) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d["SimulationId"].ToString()}','{d["DOCNumberNO"].ToString()}','07','{d["Id"].ToString()}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d01["NeedId"].ToString()}')"))
                                    {
                                        if (_Fun.Config.SendMonitorMail07 != "")
                                        {
                                            if (!docTMP.Contains(d["DOCNumberNO"].ToString()))
                                            {
                                                docTMP.Add(d["DOCNumberNO"].ToString());
                                                mailBody07 = $"{mailBody07}<p>銷貨類 單據編號:{d["DOCNumberNO"].ToString()} {d["DOCName"].ToString()}</p>";
                                            }
                                        }
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,'07',{d["SimulationId"].ToString()},");
                                        //#region 通知網頁更新
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                //###???errorType 暫時寫死
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,'07',{d["SimulationId"].ToString()},");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                    }
                                }
                            }
                            if (mailBody07 != "")
                            { _Fun.Mail_Send(_Fun.Config.SendMonitorMail07.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail07.Split(',')[1].Split(';'), mailSub, mailBody07, null, false); }
                        }
                        #endregion
                        #region 3存貨類 08
                        sql = $@"select a.SimulationId,a.Id,a.DOCNumberNO,c.DOCName FROM SoftNetMainDB.[dbo].[DOC3stockII] as a
                                        right join  SoftNetMainDB.[dbo].[DOC3stock] as b on a.DOCNumberNO=b.DOCNumberNO
                                        join  SoftNetMainDB.[dbo].[DOCRole] as c on c.DOCNO=b.DOCNO
                                        where b.ServerId='{_Fun.Config.ServerId}' and b.FlowStatus='Y' and a.ArrivalDate<='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' and a.IsOK='0'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null)
                        {
                            docTMP.Clear();
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (db.DB_GetQueryCount($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}' and ErrorType='08' and ErrorKey='{d["Id"].ToString()}' and ActionType=''") <= 0)
                                {
                                    DataRow d01 = db.DB_GetFirstDataByDataRow($"select NeedId FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}'");
                                    if (d01 != null && db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,DOCNumberNO,ErrorType,ErrorKey,LogDate,NeedId) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d["SimulationId"].ToString()}','{d["DOCNumberNO"].ToString()}','08','{d["Id"].ToString()}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d01["NeedId"].ToString()}')"))
                                    {
                                        if (_Fun.Config.SendMonitorMail08 != "")
                                        {
                                            if (!docTMP.Contains(d["DOCNumberNO"].ToString()))
                                            {
                                                docTMP.Add(d["DOCNumberNO"].ToString());
                                                mailBody08 = $"{mailBody08}<p>存貨類 單據編號:{d["DOCNumberNO"].ToString()} {d["DOCName"].ToString()}</p>";
                                            }
                                        }
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,'08',{d["SimulationId"].ToString()},");
                                        //#region 通知網頁更新
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                //###???errorType 暫時寫死
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,'08',{d["SimulationId"].ToString()},");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                    }
                                }
                            }
                            if (mailBody08 != "")
                            { _Fun.Mail_Send(_Fun.Config.SendMonitorMail08.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail08.Split(',')[1].Split(';'), mailSub, mailBody08, null, false); }
                        }
                        #endregion
                        #region 4生產類 09
                        sql = $@"select a.SimulationId,a.Id,a.DOCNumberNO,c.DOCName FROM SoftNetMainDB.[dbo].[DOC4ProductionII] as a
                            right join  SoftNetMainDB.[dbo].[DOC4Production] as b on a.DOCNumberNO=b.DOCNumberNO
                            join  SoftNetMainDB.[dbo].[DOCRole] as c on c.DOCNO=b.DOCNO
                            where b.ServerId='{_Fun.Config.ServerId}' and b.FlowStatus='Y' and a.ArrivalDate<='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' and a.IsOK='0'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null)
                        {
                            docTMP.Clear();
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (db.DB_GetQueryCount($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}' and ErrorType='09' and ErrorKey='{d["Id"].ToString()}' and ActionType=''") <= 0)
                                {
                                    DataRow d01 = db.DB_GetFirstDataByDataRow($"select NeedId FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}'");
                                    if (d01 != null && db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,DOCNumberNO,ErrorType,ErrorKey,LogDate,NeedId) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d["SimulationId"].ToString()}','{d["DOCNumberNO"].ToString()}','09','{d["Id"].ToString()}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d01["NeedId"].ToString()}')"))
                                    {
                                        if (_Fun.Config.SendMonitorMail09 != "")
                                        {
                                            if (!docTMP.Contains(d["DOCNumberNO"].ToString()))
                                            {
                                                docTMP.Add(d["DOCNumberNO"].ToString());
                                                mailBody09 = $"{mailBody09}<p>生產類 單據編號:{d["DOCNumberNO"].ToString()} {d["DOCName"].ToString()}</p>";
                                            }
                                        }
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,'09',{d["SimulationId"].ToString()},");
                                        //#region 通知網頁更新
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                ###???errorType 暫時寫死
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,'09',{d["SimulationId"].ToString()},");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                    }
                                }
                            }
                            if (mailBody09 != "")
                            { _Fun.Mail_Send(_Fun.Config.SendMonitorMail09.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail09.Split(',')[1].Split(';'), mailSub, mailBody09, null, false); }
                        }
                        #endregion
                        #region 5委外類 10
                        sql = $@"select a.SimulationId,a.Id,a.DOCNumberNO,c.DOCName FROM SoftNetMainDB.[dbo].[DOC5OUTII] as a
                            right join  SoftNetMainDB.[dbo].[DOC5OUT] as b on a.DOCNumberNO=b.DOCNumberNO
                            join  SoftNetMainDB.[dbo].[DOCRole] as c on c.DOCNO=b.DOCNO
                            where b.ServerId='{_Fun.Config.ServerId}' and b.FlowStatus='Y' and a.ArrivalDate<='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' and a.IsOK='0'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null)
                        {
                            docTMP.Clear();
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (db.DB_GetQueryCount($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}' and ErrorType='10' and ErrorKey='{d["Id"].ToString()}' and ActionType=''") <= 0)
                                {
                                    DataRow d01 = db.DB_GetFirstDataByDataRow($"select NeedId FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}'");
                                    //###???errorType 暫時寫死
                                    if (d01 != null && db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,DOCNumberNO,ErrorType,ErrorKey,LogDate,NeedId) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d["SimulationId"].ToString()}','{d["DOCNumberNO"].ToString()}','10','{d["Id"].ToString()}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d01["NeedId"].ToString()}')"))
                                    {
                                        if (_Fun.Config.SendMonitorMail10 != "")
                                        {
                                            if (!docTMP.Contains(d["DOCNumberNO"].ToString()))
                                            {
                                                docTMP.Add(d["DOCNumberNO"].ToString());
                                                mailBody10 = $"{mailBody10}<p>委外類 單據編號:{d["DOCNumberNO"].ToString()} {d["DOCName"].ToString()}</p>";
                                            }
                                        }
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,'10',{d["SimulationId"].ToString()},");
                                        //#region 通知網頁更新
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                //###???errorType 暫時寫死
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,'10',{d["SimulationId"].ToString()},");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                    }
                                }
                            }
                            if (mailBody10 != "")
                            { _Fun.Mail_Send(_Fun.Config.SendMonitorMail10.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail10.Split(',')[1].Split(';'), mailSub, mailBody10, null, false); }
                        }
                        #endregion

                        #endregion
                        _Fun._a07 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 修正若工站生產量已達計畫量,清除工作站負荷,並主動從工站執行工單明細移除移除,每日首班前清除APS_WorkTimeNote負荷
                        dr_tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{_Fun.Config.DefaultCalendarName}' and Holiday='{DateTime.Now.ToString("yyyy-MM-dd")}'");
                        if (dr_tmp != null)
                        {
                            DateTime now = DateTime.Now;
                            string[] compe = dr_tmp["Shift_Morning"].ToString().Trim().Split(',');
                            string[] comps = dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',');
                            DateTime etime = new DateTime(now.Year, now.Month, now.Day, int.Parse(compe[0].Split(':')[0]), int.Parse(compe[0].Split(':')[1]), 0);
                            DateTime stime = new DateTime(now.Year, now.Month, now.Day, int.Parse(comps[3].Split(':')[0]), int.Parse(comps[3].Split(':')[1]), 0);
                            if (now >= stime && now < etime)
                            {
                                sql = $@"SELECT StationNO,SimulationId FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where CONVERT(varchar(100), CalendarDate, 23)='{DateTime.Now.ToString("yyyy-MM-dd")}' and DOCNumberNO!='' and (Time1_C!=0 or Time2_C!=0 or Time3_C!=0 or Time4_C!=0) 
                                    group by StationNO,SimulationId";
                                dt_tmp = db.DB_GetData(sql);
                                if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                                {
                                    foreach (DataRow d in dt_tmp.Rows)
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where APS_StationNO='{d["StationNO"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}' and Detail_QTY>=NeedQTY");
                                        if (dr_tmp != null)
                                        {
                                            _ = db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_WorkTimeNote] set Time1_C=0,Time2_C=0,Time3_C=0,Time4_C=0 where StationNO='{d["StationNO"].ToString()}' and SimulationId='{d["SimulationId"].ToString()}'");
                                            _ = db.DB_SetData($"update SoftNetSYSDB.[dbo].[APS_Simulation] set IsOK='1' where SimulationId='{d["SimulationId"].ToString()}'");
                                            //###??? 要通知電子紙 或 平板更新
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        #region 監測工站每日執行工單明細的開工狀態, 是否如預期  03
                        string mailSub = $"{_Fun.Config.ServerId} 警示 工單開工狀態,沒如預期.";
                        string mailBody03 = "";
                        sql = $@"SELECT a.*,c.PartNO,c.Source_StationNO_IndexSN,c.Source_StationNO_Custom_IndexSN FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] as a
                            join SoftNetSYSDB.[dbo].[PP_WorkOrder] as b on b.OrderNO=a.DOCNumberNO and (b.CloseType='0' or b.EndTime is NULL)
                            join SoftNetSYSDB.[dbo].[APS_Simulation] as c on c.SimulationId=a.SimulationId
                            join SoftNetSYSDB.[dbo].[APS_NeedData] as d on d.State='6' and a.NeedId=d.Id
                            join SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as e on a.SimulationId=e.SimulationId and e.DOCNumberNO!=''
                            where b.ServerId='{_Fun.Config.ServerId}' and a.DOCNumberNO!='' and CONVERT(varchar(100), a.CalendarDate, 23)='{DateTime.Now.ToString("yyyy-MM-dd")}' and (a.Time1_C>=1 or a.Time2_C>0 or a.Time3_C>0 or a.Time4_C>0) order by a.StationNO,a.CalendarDate";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null)
                        {
                            #region 變數
                            DateTime stime = DateTime.Now;
                            DateTime etime = DateTime.Now;
                            DateTime stime2 = DateTime.Now;
                            DateTime etime2 = DateTime.Now;
                            DateTime now = DateTime.Now;
                            int type_Time = 0;
                            bool isrun = false;
                            DataRow dr_ErrorData = null;
                            DataRow dr_tmp2 = null;
                            DataRow PP_Station = null;
                            DataTable dt_tmp3 = null;
                            double typeTotalTime = 0;
                            bool has_data = false;
                            bool is_INFO = false;
                            #endregion
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                is_INFO = false;
                                has_data = false;
                                #region 判斷工站狀態, 是否正常開工, 解除.
                                PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}'");
                                dr_ErrorData = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["StationNO"].ToString()}' and ErrorType='03' and ActionType=''");
                                if (dr_ErrorData != null)
                                {
                                    if (int.Parse(d["Time1_C"].ToString()) <= 1 && d["Time2_C"].ToString() == "0" && d["Time3_C"].ToString() == "0" && d["Time4_C"].ToString() == "0")
                                    {
                                        _ = db.DB_SetData($"delete from SoftNetSYSDB.[dbo].APS_Simulation_ErrorData where Id='{dr_ErrorData["Id"].ToString()}'");
                                        _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,'03',{d["SimulationId"].ToString()},0");
                                        //#region 通知網頁更新 解除
                                        //if (WebSocketServiceOJB != null)
                                        //{
                                        //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                        //    {
                                        //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                        //        {
                                        //            if (r.Key != null && r.Value.socket != null)
                                        //            {
                                        //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,'03',{d["SimulationId"].ToString()},0");
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        //#endregion
                                        continue;
                                    }
                                    else
                                    { has_data = true; }

                                    if (PP_Station["Station_Type"].ToString() == "8")
                                    {
                                        dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[ManufactureII] where SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["StationNO"].ToString()}' and OrderNO='{d["DOCNumberNO"].ToString()}' and PartNO='{d["PartNO"].ToString()}' and IndexSN={d["Source_StationNO_IndexSN"].ToString()}");
                                        if (dr_tmp2 == null) { continue; }
                                        else
                                        {
                                            if (!dr_tmp2.IsNull("EndTime"))
                                            {
                                                if (has_data)
                                                {
                                                    _ = db.DB_SetData($"delete from SoftNetSYSDB.[dbo].APS_Simulation_ErrorData where Id='{dr_ErrorData["Id"].ToString()}'"); is_INFO = true;
                                                }
                                                else { continue; }
                                            }
                                            if (!dr_tmp2.IsNull("RemarkTimeS") && dr_tmp2.IsNull("RemarkTimeE"))
                                            {
                                                if (has_data)
                                                {
                                                    _ = db.DB_SetData($"delete from SoftNetSYSDB.[dbo].APS_Simulation_ErrorData where Id='{dr_ErrorData["Id"].ToString()}'"); is_INFO = true;
                                                }
                                                else { continue; }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["StationNO"].ToString()}'");
                                        if (dr_tmp2 == null) { continue; }
                                        else
                                        {
                                            if (dr_tmp2["State"].ToString() != "2" && dr_tmp2["State"].ToString() != "3")
                                            {
                                                if (has_data)
                                                {
                                                    _ = db.DB_SetData($"delete from SoftNetSYSDB.[dbo].APS_Simulation_ErrorData where Id='{dr_ErrorData["Id"].ToString()}'"); is_INFO = true;
                                                }
                                                else { continue; }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (int.Parse(d["Time1_C"].ToString()) <= 1 && d["Time2_C"].ToString() == "0" && d["Time3_C"].ToString() == "0" && d["Time4_C"].ToString() == "0") { continue; }
                                }
                                #endregion

                                if (!is_INFO && !has_data && Convert.ToDateTime(d["CalendarDate"]) < now)
                                {
                                    isrun = false;

                                    #region 判斷工站負荷時間是否足夠
                                    DataRow dr_PP_HolidayCalendar = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_HolidayCalendar where ServerId='{_Fun.Config.ServerId}' and CalendarName='{PP_Station["CalendarName"].ToString()}' and Holiday='{DateTime.Now.ToString("yyyy-MM-dd")}'");
                                    if (dr_PP_HolidayCalendar != null)
                                    {
                                        type_Time = int.Parse(d["Time1_C"].ToString());
                                        if (bool.Parse(d["Type1"].ToString()) && type_Time != 0)
                                        {
                                            string[] comp = dr_PP_HolidayCalendar["Shift_Morning"].ToString().Trim().Split(',');
                                            etime = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                            stime = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                            if (now >= stime && now < etime)
                                            {
                                                dt_tmp3 = db.DB_GetData($"SELECT TOP 100 LOGDateTime from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and IndexSN={d["Source_StationNO_IndexSN"].ToString()} and OperateType like '%開工%' and LOGDateTime>='{stime.ToString("MM/dd/yyyy HH:mm:ss.fff")}' and LOGDateTime<'{etime.ToString("MM/dd/yyyy HH:mm:ss.fff")}'");
                                                if (dt_tmp3 != null && dt_tmp3.Rows.Count > 0)
                                                {
                                                    typeTotalTime = 0;
                                                    foreach (DataRow d3 in dt_tmp3.Rows)
                                                    {
                                                        typeTotalTime += _SFC_Common.TimeCompute2Seconds(stime, Convert.ToDateTime(d3["LOGDateTime"]));
                                                    }
                                                    if (typeTotalTime != 0)
                                                    {
                                                        stime = stime.AddSeconds(Math.Floor(typeTotalTime / dt_tmp3.Rows.Count));
                                                        if (now >= stime) { isrun = true; }
                                                    }
                                                }
                                                else
                                                {
                                                    //###???將來要跟RUNTimeServer的 isARGs10_offset = 15;將來改參數
                                                    stime = stime.AddMinutes(isARGs10_offset);
                                                    if (now > stime) { isrun = true; }
                                                }
                                            }
                                            else
                                            {
                                                etime2 = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                stime2 = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                if (now >= stime2 && now < etime2)
                                                {
                                                    dt_tmp3 = db.DB_GetData($"SELECT TOP 100 LOGDateTime from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and IndexSN={d["Source_StationNO_IndexSN"].ToString()} and OperateType like '%開工%' and LOGDateTime>='{stime.ToString("MM/dd/yyyy HH:mm:ss.fff")}' and LOGDateTime<'{etime.ToString("MM/dd/yyyy HH:mm:ss.fff")}'");
                                                    if (dt_tmp3 != null && dt_tmp3.Rows.Count > 0)
                                                    {
                                                        typeTotalTime = 0;
                                                        foreach (DataRow d3 in dt_tmp3.Rows)
                                                        {
                                                            typeTotalTime += _SFC_Common.TimeCompute2Seconds(stime, Convert.ToDateTime(d3["LOGDateTime"]));
                                                        }
                                                        if (typeTotalTime != 0)
                                                        {
                                                            stime = stime.AddSeconds(Math.Floor(typeTotalTime / dt_tmp3.Rows.Count));
                                                            if (now >= stime) { isrun = true; }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //###???將來要跟RUNTimeServer的 isARGs10_offset = 10;將來改參數
                                                        stime = stime.AddMinutes(isARGs10_offset);
                                                        if (now > stime) { isrun = true; }
                                                    }
                                                }
                                            }
                                        }
                                        type_Time = int.Parse(d["Time2_C"].ToString());
                                        if (!isrun && bool.Parse(d["Type2"].ToString()) && type_Time != 0)
                                        {
                                            string[] comp = dr_PP_HolidayCalendar["Shift_Afternoon"].ToString().Trim().Split(',');
                                            etime = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                            stime = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                            if (now >= stime && now < etime)
                                            {
                                                dt_tmp3 = db.DB_GetData($"SELECT TOP 100 LOGDateTime from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and IndexSN={d["Source_StationNO_IndexSN"].ToString()} and OperateType like '%開工%' and LOGDateTime>='{stime.ToString("MM/dd/yyyy HH:mm:ss.fff")}' and LOGDateTime<'{etime.ToString("MM/dd/yyyy HH:mm:ss.fff")}'");
                                                if (dt_tmp3 != null && dt_tmp3.Rows.Count > 0)
                                                {
                                                    typeTotalTime = 0;
                                                    foreach (DataRow d3 in dt_tmp3.Rows)
                                                    {
                                                        typeTotalTime += _SFC_Common.TimeCompute2Seconds(stime, Convert.ToDateTime(d3["LOGDateTime"]));
                                                    }
                                                    if (typeTotalTime != 0)
                                                    {
                                                        stime = stime.AddSeconds(Math.Floor(typeTotalTime / dt_tmp3.Rows.Count));
                                                        if (now >= stime) { isrun = true; }
                                                    }
                                                }
                                                else
                                                {
                                                    //###???將來要跟RUNTimeServer的 isARGs10_offset = 10;將來改參數
                                                    stime = stime.AddMinutes(isARGs10_offset);
                                                    if (now > stime) { isrun = true; }
                                                }
                                            }
                                            else
                                            {
                                                etime2 = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                stime2 = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                if (now >= stime2 && now < etime2)
                                                {
                                                    dt_tmp3 = db.DB_GetData($"SELECT TOP 100 LOGDateTime from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and IndexSN={d["Source_StationNO_IndexSN"].ToString()} and OperateType like '%開工%' and LOGDateTime>='{stime.ToString("MM/dd/yyyy HH:mm:ss.fff")}' and LOGDateTime<'{etime.ToString("MM/dd/yyyy HH:mm:ss.fff")}'");
                                                    if (dt_tmp3 != null && dt_tmp3.Rows.Count > 0)
                                                    {
                                                        typeTotalTime = 0;
                                                        foreach (DataRow d3 in dt_tmp3.Rows)
                                                        {
                                                            typeTotalTime += _SFC_Common.TimeCompute2Seconds(stime, Convert.ToDateTime(d3["LOGDateTime"]));
                                                        }
                                                        if (typeTotalTime != 0)
                                                        {
                                                            stime = stime.AddSeconds(Math.Floor(typeTotalTime / dt_tmp3.Rows.Count));
                                                            if (now >= stime) { isrun = true; }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //###???將來要跟RUNTimeServer的 isARGs10_offset = 10;將來改參數
                                                        stime = stime.AddMinutes(isARGs10_offset);
                                                        if (now > stime) { isrun = true; }
                                                    }
                                                }
                                            }

                                        }
                                        type_Time = int.Parse(d["Time3_C"].ToString());
                                        if (!isrun && bool.Parse(d["Type3"].ToString()) && type_Time != 0)
                                        {
                                            string[] comp = dr_PP_HolidayCalendar["Shift_Night"].ToString().Trim().Split(',');
                                            etime = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                            stime = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                            if (now >= stime && now < etime)
                                            {
                                                dt_tmp3 = db.DB_GetData($"SELECT TOP 100 LOGDateTime from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and IndexSN={d["Source_StationNO_IndexSN"].ToString()} and OperateType like '%開工%' and LOGDateTime>='{stime.ToString("MM/dd/yyyy HH:mm:ss.fff")}' and LOGDateTime<'{etime.ToString("MM/dd/yyyy HH:mm:ss.fff")}'");
                                                if (dt_tmp3 != null && dt_tmp3.Rows.Count > 0)
                                                {
                                                    typeTotalTime = 0;
                                                    foreach (DataRow d3 in dt_tmp3.Rows)
                                                    {
                                                        typeTotalTime += _SFC_Common.TimeCompute2Seconds(stime, Convert.ToDateTime(d3["LOGDateTime"]));
                                                    }
                                                    if (typeTotalTime != 0)
                                                    {
                                                        stime = stime.AddSeconds(Math.Floor(typeTotalTime / dt_tmp3.Rows.Count));
                                                        if (now >= stime) { isrun = true; }
                                                    }
                                                }
                                                else
                                                {
                                                    //###???將來要跟RUNTimeServer的 isARGs10_offset = 10;將來改參數
                                                    stime = stime.AddMinutes(isARGs10_offset);
                                                    if (now > stime) { isrun = true; }
                                                }
                                            }
                                            else
                                            {
                                                stime2 = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                { etime2 = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                                else
                                                { etime2 = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                                if (now >= stime2 && now < etime2)
                                                {
                                                    dt_tmp3 = db.DB_GetData($"SELECT TOP 100 LOGDateTime from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and IndexSN={d["Source_StationNO_IndexSN"].ToString()} and OperateType like '%開工%' and LOGDateTime>='{stime.ToString("MM/dd/yyyy HH:mm:ss.fff")}' and LOGDateTime<'{etime.ToString("MM/dd/yyyy HH:mm:ss.fff")}'");
                                                    if (dt_tmp3 != null && dt_tmp3.Rows.Count > 0)
                                                    {
                                                        typeTotalTime = 0;
                                                        foreach (DataRow d3 in dt_tmp3.Rows)
                                                        {
                                                            typeTotalTime += _SFC_Common.TimeCompute2Seconds(stime, Convert.ToDateTime(d3["LOGDateTime"]));
                                                        }
                                                        if (typeTotalTime != 0)
                                                        {
                                                            stime = stime.AddSeconds(Math.Floor(typeTotalTime / dt_tmp3.Rows.Count));
                                                            if (now >= stime) { isrun = true; }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //###???將來要跟RUNTimeServer的 isARGs10_offset = 10;將來改參數
                                                        stime = stime.AddMinutes(isARGs10_offset);
                                                        if (now > stime) { isrun = true; }
                                                    }
                                                }
                                            }
                                        }
                                        type_Time = int.Parse(d["Time4_C"].ToString());
                                        if (!isrun && bool.Parse(d["Type4"].ToString()) && type_Time != 0)
                                        {
                                            string[] comp = dr_PP_HolidayCalendar["Shift_Graveyard"].ToString().Trim().Split(',');
                                            stime = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                            if (int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                                            { etime = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0).AddDays(1); }
                                            else
                                            { etime = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0); }
                                            if (now >= stime && now < etime)
                                            {
                                                dt_tmp3 = db.DB_GetData($"SELECT TOP 100 LOGDateTime from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and IndexSN={d["Source_StationNO_IndexSN"].ToString()} and OperateType like '%開工%' and LOGDateTime>='{stime.ToString("MM/dd/yyyy HH:mm:ss.fff")}' and LOGDateTime<'{etime.ToString("MM/dd/yyyy HH:mm:ss.fff")}'");
                                                if (dt_tmp3 != null && dt_tmp3.Rows.Count > 0)
                                                {
                                                    typeTotalTime = 0;
                                                    foreach (DataRow d3 in dt_tmp3.Rows)
                                                    {
                                                        typeTotalTime += _SFC_Common.TimeCompute2Seconds(stime, Convert.ToDateTime(d3["LOGDateTime"]));
                                                    }
                                                    if (typeTotalTime != 0)
                                                    {
                                                        stime = stime.AddSeconds(Math.Floor(typeTotalTime / dt_tmp3.Rows.Count));
                                                        if (now >= stime) { isrun = true; }
                                                    }
                                                }
                                                else
                                                {
                                                    //###???將來要跟RUNTimeServer的 isARGs10_offset = 10;將來改參數
                                                    stime = stime.AddMinutes(isARGs10_offset);
                                                    if (now > stime) { isrun = true; }
                                                }
                                            }
                                            else
                                            {
                                                stime2 = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                { etime2 = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                                else
                                                { etime2 = new DateTime(now.Year, now.Month, now.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                                if (now >= stime2 && now < etime2)
                                                {
                                                    dt_tmp3 = db.DB_GetData($"SELECT TOP 100 LOGDateTime from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and IndexSN={d["Source_StationNO_IndexSN"].ToString()} and OperateType like '%開工%' and LOGDateTime>='{stime.ToString("MM/dd/yyyy HH:mm:ss.fff")}' and LOGDateTime<'{etime.ToString("MM/dd/yyyy HH:mm:ss.fff")}'");
                                                    if (dt_tmp3 != null && dt_tmp3.Rows.Count > 0)
                                                    {
                                                        typeTotalTime = 0;
                                                        foreach (DataRow d3 in dt_tmp3.Rows)
                                                        {
                                                            typeTotalTime += _SFC_Common.TimeCompute2Seconds(stime, Convert.ToDateTime(d3["LOGDateTime"]));
                                                        }
                                                        if (typeTotalTime != 0)
                                                        {
                                                            stime = stime.AddSeconds(Math.Floor(typeTotalTime / dt_tmp3.Rows.Count));
                                                            if (now >= stime) { isrun = true; }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //###???將來要跟RUNTimeServer的 isARGs10_offset = 10;將來改參數
                                                        stime = stime.AddMinutes(isARGs10_offset);
                                                        if (now > stime) { isrun = true; }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    #endregion

                                    if (isrun)
                                    {
                                        if (_Fun.Config.SendMonitorMail03 != "")
                                        {
                                            mailBody03 = $"{mailBody03}<p>排程碼:{d["NeedId"].ToString()} {d["SimulationId"].ToString()} 工站:{d["StationNO"].ToString()} 未開工</p>";
                                        }
                                        _ = db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Count,StationNO,DOCNumberNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d["SimulationId"].ToString()}','03','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d["NeedId"].ToString()}',1,'{d["StationNO"].ToString()}','{d["DOCNumberNO"].ToString()}')");
                                    }
                                    else
                                    {
                                        dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["StationNO"].ToString()}' and ErrorType='03' and ActionType=''");
                                        if (dr_tmp2 != null)
                                        {
                                            if (db.DB_SetData($"delete from SoftNetSYSDB.[dbo].APS_Simulation_ErrorData where Id='{dr_tmp2["Id"].ToString()}'")) { is_INFO = true; }
                                        }
                                    }
                                }
                                if (is_INFO)
                                {
                                    _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,'03',{d["SimulationId"].ToString()},0");
                                    //#region 通知網頁更新 解除
                                    //if (WebSocketServiceOJB != null)
                                    //{
                                    //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                    //    {
                                    //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                    //        {
                                    //            if (r.Key != null && r.Value.socket != null)
                                    //            {
                                    //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,'03',{d["SimulationId"].ToString()},0");
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                    //#endregion
                                }
                            }
                            if (mailBody03 != "")
                            { _Fun.Mail_Send(_Fun.Config.SendMonitorMail03.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail03.Split(',')[1].Split(';'), mailSub, mailBody03, null, false); }
                        }
                        #endregion

                        #region 隔日解除
                        sql = $@"SELECT a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                                    where a.ServerId='{_Fun.Config.ServerId}' and a.ActionType='' and a.ErrorType='03' and CONVERT(varchar(100), a.LogDate, 23)<='{DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd")}'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            DateTime stime = DateTime.Now;
                            DateTime now = DateTime.Now;
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='1',ActionLogDate='{logdate}' where Id='{d["Id"].ToString()}'"))
                                {
                                    _ = SendWebSocketClent_INFO($"SendALLClient,SimulatioStatusChange,03,{d["SimulationId"].ToString()},1");
                                    //#region 通知網頁更新
                                    //if (WebSocketServiceOJB != null)
                                    //{
                                    //    lock (WebSocketServiceOJB.lock__WebSocketList)
                                    //    {
                                    //        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                    //        {
                                    //            if (r.Key != null && r.Value.socket != null)
                                    //            {
                                    //                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,03,{d["SimulationId"].ToString()},1");
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                    //#endregion
                                }
                            }
                        }
                        #endregion
                        _Fun._a08 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 工單應開未開  11  PS:03會判斷到, 所以暫時取消
                        /*
                        string mailSub = $"{_Fun.Config.ServerId} 警示 工單應開未開狀態.";
                        string mailBody11 = "";

                        dt_tmp = db.DB_GetData($@"SELECT * FROM SoftNetSYSDB.[dbo].PP_WorkOrder WHERE ServerId='{_Fun.Config.ServerId}' and StartTime IS NULL AND EstimatedStartTime < GETDATE()");
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (db.DB_GetQueryCount($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO='{d["OrderNO"].ToString()}' and ErrorType='11'") <= 0)
                                {
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,DOCNumberNO,ErrorType,LogDate,NeedId) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','','{d["OrderNO"].ToString()}','11','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d["NeedId"].ToString()}')"))
                                    {
                                        //###???還不知要通知誰
                                    }
                                }
                            }
                        }
                        #region 解除 11
                        sql = $@"SELECT a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                            join SoftNetSYSDB.[dbo].PP_WorkOrder as b on b.ServerId='{_Fun.Config.ServerId}' and b.OrderNO=a.DOCNumberNO and b.StartTime IS NOT NULL
                            where a.ServerId='{_Fun.Config.ServerId}' and a.ActionType='' and a.ErrorType='11'";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='0',ActionLogDate='{logdate}' where Id='{d["Id"].ToString()}'"))
                                {
                                    //###???還不知要通知誰
                                }
                            }

                        }
                        #endregion
                        */
                        #endregion
                        _Fun._a09 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 監測工單應關而未關  12 
                        //###???將來要加參數定義那些WO單據別不處理
                        dt_tmp = db.DB_GetData($@"SELECT b.StationNO,a.* FROM SoftNetSYSDB.[dbo].PP_WorkOrder as a, SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] as b
                                            WHERE a.ServerId='{_Fun.Config.ServerId}' AND a.OrderNO=b.OrderNO and a.EndTime IS NULL and b.IsLastStation='1' and a.Quantity>0 and b.TotalInput>=a.Quantity");
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            DataRow dr_PP_Station = null;
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (db.DB_GetQueryCount($"select Id from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and DOCNumberNO='{d["OrderNO"].ToString()}' and ErrorType='12'") <= 0)
                                {
                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,DOCNumberNO,ErrorType,LogDate) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','','{d["OrderNO"].ToString()}','12','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}')"))
                                    {
                                        //###???還不知要通知誰
                                        #region 寫入 主動干涉檔,等時間到自動關閉
                                        DateTime intime = DateTime.Now;
                                        dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}'");
                                        sql = $@"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_PP_Station["CalendarName"].ToString()}' and Holiday>='{intime.ToString("MM/dd/yyyy")}'";
                                        dr_tmp = db.DB_GetFirstDataByDataRow(sql);
                                        if (dr_tmp != null)
                                        {
                                            bool isrun = true;
                                            DateTime etime = DateTime.Now;
                                            DateTime stime2 = DateTime.Now;
                                            string[] comp = dr_tmp["Shift_Morning"].ToString().Trim().Split(',');
                                            stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                            etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                            if (etime >= intime && intime >= stime2)
                                            {
                                                isrun = false;
                                                comp = dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',');
                                                etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                intime = etime.AddMinutes(isARGs10_offset);
                                            }
                                            if (isrun)
                                            {
                                                comp = dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',');
                                                stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                if (etime >= intime && intime >= stime2)
                                                {
                                                    isrun = false;
                                                    comp = dr_tmp["Shift_Night"].ToString().Trim().Split(',');
                                                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                    intime = etime.AddMinutes(isARGs10_offset);
                                                }
                                            }
                                            if (isrun)
                                            {
                                                comp = dr_tmp["Shift_Night"].ToString().Trim().Split(',');
                                                if (int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                {
                                                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1);
                                                }
                                                else
                                                {
                                                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                }
                                                stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                if (etime >= intime && intime >= stime2)
                                                {
                                                    isrun = false;
                                                    intime = etime.AddMinutes(isARGs10_offset);
                                                }
                                            }
                                            if (isrun)
                                            {
                                                comp = dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',');
                                                if (int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                {
                                                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1);

                                                }
                                                else
                                                {
                                                    etime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                }
                                                stime2 = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                if (etime >= intime && intime >= stime2)
                                                {
                                                    isrun = false;
                                                    intime = etime.AddMinutes(isARGs10_offset);
                                                }
                                            }
                                            else if (isrun)
                                            {
                                                comp = dr_tmp["Shift_Morning"].ToString().Trim().Split(',');
                                                intime = new DateTime(intime.Year, intime.Month, intime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1);
                                            }
                                            _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','12','{d["NeedId"].ToString()}','','','{d["OrderNO"].ToString()}','','','{intime.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                        }
                                        #endregion

                                    }
                                }
                            }
                        }
                        if (db.DB_GetQueryCount($"select * from SoftNetSYSDB.[dbo].APS_Simulation_ErrorData where ServerId='{_Fun.Config.ServerId}' and ErrorType='12' and ActionType=''") > 0)
                        {
                            #region 解除 12
                            sql = $@"SELECT a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                                       join SoftNetSYSDB.[dbo].PP_WorkOrder as b on b.ServerId='{_Fun.Config.ServerId}' and b.OrderNO=a.DOCNumberNO and b.EndTime IS NOT NULL
                                       where a.ServerId='{_Fun.Config.ServerId}' and a.ActionType='' and a.ErrorType='12'";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                foreach (DataRow d in dt_tmp.Rows)
                                {
                                    if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='0',ActionLogDate='{logdate}' where Id='{d["Id"].ToString()}'"))
                                    {
                                        //###???還不知要通知誰
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WarningData] where ErrorType='12' and IsDEL='0' and ServerId='{_Fun.Config.ServerId}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}'");
                                        if (dr_tmp != null)
                                        {
                                            _ = db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[APS_WarningData] where ErrorType='12' and IsDEL='0' and ServerId='{_Fun.Config.ServerId}' and DOCNumberNO='{d["DOCNumberNO"].ToString()}'");
                                        }
                                    }
                                }

                            }
                            #endregion
                        }
                        #endregion
                        _Fun._a10 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        //###???
                        #region 監看工單CT/UPH 是否未達有效CT值, 是否遠離偏離目標CT值 14

                        #endregion
                        _Fun._a11 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 監看工站人工報工是否延遲(影響有效值) 15 (發mail警示)
                        {
                            DateTime tmp_date = DateTime.Now;
                            dt_tmp = db.DB_GetData($"SELECT * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}'");
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                #region 變數
                                DataRow dr_tmp2 = null;
                                string MailSub1 = "";
                                string MailSub2 = "";
                                string MailSub3 = "";
                                string MailSub4 = "";
                                string MailBody1 = "";
                                string MailBody2 = "";
                                string MailBody3 = "";
                                string MailBody4 = "";

                                int tmo_i = 0;
                                int tmo_e = 0;
                                DateTime tmp_sdate = DateTime.Now;
                                DateTime tmp_edate = DateTime.Now;
                                #endregion
                                foreach (DataRow d in dt_tmp.Rows)
                                {
                                    #region 判斷該站是否有派工 
                                    if (d["Station_Type"].ToString() == "1")
                                    {
                                        dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PartNO!='' and OrderNO!='' and RemarkTimeS is not NULL");
                                        if (dr_tmp2 == null) { continue; }
                                    }
                                    else
                                    {
                                        dr_tmp2 = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PartNO!='' and OrderNO!='' and RemarkTimeS is not NULL and EndTime is NULL");
                                        if (dr_tmp2 == null) { continue; }
                                    }
                                    #endregion
                                    tmp_date = DateTime.Now;
                                    #region 先找 當天行事曆
                                    if (tmp_date.Hour >= 0 && tmp_date.Hour < 6)  //###???將來 6 要參數化
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and CalendarName='{d["CalendarName"].ToString()}' and Holiday='{tmp_date.AddDays(-1).ToString("MM/dd/yyyy")}'");
                                        if (dr_tmp == null) { continue; }
                                    }
                                    else
                                    {
                                        dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and CalendarName='{d["CalendarName"].ToString()}' and Holiday='{tmp_date.ToString("MM/dd/yyyy")}'");
                                        if (dr_tmp == null) { continue; }
                                    }
                                    #endregion
                                    if (dr_tmp != null)
                                    {
                                        if (bool.Parse(dr_tmp["Flag_Graveyard"].ToString()) == true)
                                        {
                                            #region 第四時段 抓中間休息時間
                                            if (tmp_date.Hour >= 0 && tmp_date.Hour <= int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]))
                                            { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[1]), 0); }
                                            else
                                            {
                                                if (int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                                { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[1]), 0).AddDays(1); }
                                                else
                                                { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[1]), 0); }
                                            }
                                            tmp_edate = new DateTime(tmp_sdate.Year, tmp_sdate.Month, tmp_sdate.Day, int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 59);
                                            if (tmp_edate >= tmp_date && tmp_date >= tmp_sdate)
                                            {
                                                tmo_i = (_SFC_Common.TimeCompute2Seconds(tmp_sdate, tmp_edate)) / 2;
                                                tmo_e = _SFC_Common.TimeCompute2Seconds(tmp_sdate, tmp_date);
                                                if (tmo_e >= tmo_i)
                                                {
                                                    if (_Fun.Config.RUNMode == '2' && db.DB_GetQueryCount($"select * from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and ErrorType='15' and CONVERT(varchar(100), LogDate, 23)='{tmp_date.ToString("yyyy-MM-dd")}' and Time_N_Type='4'") > 0)
                                                    { continue; }
                                                    if (MailSub4 == "") { MailSub4 = $"{_Fun.Config.ServerId} {dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[0]}到{dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[3]} 之間, 報工可能有 延遲 提示"; }
                                                    if (int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                                    { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), 0).AddDays(1); }
                                                    else
                                                    { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), 0); }
                                                    tmp_edate = new DateTime(tmp_sdate.Year, tmp_sdate.Month, tmp_sdate.Day, int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                    if (_Fun.Config.RUNMode == '1')
                                                    {
                                                        if (d["Station_Type"].ToString() == "8")
                                                        {
                                                            #region 多工單工站
                                                            DataTable dt_tmp2 = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeS<='{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and EndTime is NULL");
                                                            if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                                                            {
                                                                DataRow dr_tmp3 = null;
                                                                foreach (DataRow d2 in dt_tmp2.Rows)
                                                                {
                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()} and PartNO='{d2["PartNO"].ToString()}' and ReportTime>5");
                                                                    if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                                    {
                                                                        int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                        if (avgReportTime > 5)
                                                                        {
                                                                            if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), tmp_sdate, tmp_edate) >= avgReportTime)
                                                                            {
                                                                                dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTime>'{Convert.ToDateTime(d2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTime<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()}  and PartNO='{d2["PartNO"].ToString()}'");
                                                                                if (dr_tmp3 == null)
                                                                                {
                                                                                    MailBody4 = $"{MailBody4}<p>工站:{d["StationNO"].ToString()}  工序編號:{d2["IndexSN"].ToString()} 製程:{d2["PP_Name"].ToString()}  料號:{d2["PartNO"].ToString()}</p>";
                                                                                    string needId = "";
                                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{d2["OrderNO"].ToString()}'");
                                                                                    if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','4','{d["StationNO"].ToString()}')"))
                                                                                    {
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            #region 單工
                                                            DataRow dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()} and PartNO='{dr_tmp2["PartNO"].ToString()}' and ReportTime>5");
                                                            if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                            {
                                                                int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                if (avgReportTime > 5)
                                                                {
                                                                    if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp2["RemarkTimeS"]), tmp_date) >= avgReportTime)
                                                                    {
                                                                        dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(dr_tmp2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()}  and PartNO='{dr_tmp2["PartNO"].ToString()}'");
                                                                        if (dr_tmp3 == null)
                                                                        {
                                                                            MailBody4 = $"{MailBody4}<p>工站:{d["StationNO"].ToString()}  工序編號:{dr_tmp2["IndexSN"].ToString()} 製程:{dr_tmp2["PP_Name"].ToString()}  料號:{dr_tmp2["PartNO"].ToString()}</p>";
                                                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_TypeStationNO,) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{dr_tmp2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','','4','{d["StationNO"].ToString()}')"))
                                                                            {
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (d["Station_Type"].ToString() == "8")
                                                        {
                                                            #region 多工單工站
                                                            DataTable dt_tmp2 = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeS<='{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and EndTime is NULL");
                                                            if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                                                            {
                                                                DataRow dr_tmp3 = null;
                                                                foreach (DataRow d2 in dt_tmp2.Rows)
                                                                {
                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()} and PartNO='{d2["PartNO"].ToString()}' and ReportTime>5");
                                                                    if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                                    {
                                                                        int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                        if (avgReportTime > 5)
                                                                        {
                                                                            if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), tmp_sdate, tmp_edate) >= avgReportTime)
                                                                            {
                                                                                dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(d2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()}  and PartNO='{d2["PartNO"].ToString()}'");
                                                                                if (dr_tmp3 == null)
                                                                                {
                                                                                    MailBody4 = $"{MailBody4}<p>工站:{d["StationNO"].ToString()}  工序編號:{d2["IndexSN"].ToString()} 製程:{d2["PP_Name"].ToString()}  料號:{d2["PartNO"].ToString()}</p>";
                                                                                    string needId = "";
                                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{d2["OrderNO"].ToString()}'");
                                                                                    if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','4','{d["StationNO"].ToString()}')"))
                                                                                    {
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            #region 單工
                                                            DataRow dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()} and PartNO='{dr_tmp2["PartNO"].ToString()}' and ReportTime>5");
                                                            if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                            {
                                                                int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                if (avgReportTime > 5)
                                                                {
                                                                    if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp2["RemarkTimeS"]), tmp_date) >= avgReportTime)
                                                                    {
                                                                        dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(dr_tmp2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()}  and PartNO='{dr_tmp2["PartNO"].ToString()}'");
                                                                        if (dr_tmp3 == null)
                                                                        {
                                                                            MailBody4 = $"{MailBody4}<p>工站:{d["StationNO"].ToString()}  工序編號:{dr_tmp2["IndexSN"].ToString()} 製程:{dr_tmp2["PP_Name"].ToString()}  料號:{dr_tmp2["PartNO"].ToString()}</p>";
                                                                            string needId = "";
                                                                            dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_tmp2["OrderNO"].ToString()}'");
                                                                            if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{dr_tmp2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','4','{d["StationNO"].ToString()}')"))
                                                                            {
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                        }
                                        if (bool.Parse(dr_tmp["Flag_Night"].ToString()) == true)
                                        {
                                            #region 第三時段 抓中間休息時間
                                            if (int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                            { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0).AddDays(1); }
                                            else
                                            { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0); }

                                            if (int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]) > int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]))
                                            { tmp_edate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), 0).AddDays(1); }
                                            else
                                            { tmp_edate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[1]), 0); }
                                            if (tmp_edate >= tmp_date && tmp_date >= tmp_sdate)
                                            {
                                                tmo_i = (_SFC_Common.TimeCompute2Seconds(tmp_sdate, tmp_edate)) / 2;
                                                tmo_e = _SFC_Common.TimeCompute2Seconds(tmp_sdate, tmp_date);
                                                if (tmo_e >= tmo_i)
                                                {
                                                    if (_Fun.Config.RUNMode == '2' && db.DB_GetQueryCount($"select * from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and ErrorType='15' and CONVERT(varchar(100), LogDate, 23)='{tmp_date.ToString("yyyy-MM-dd")}' and Time_N_Type='3'") > 0)
                                                    { continue; }
                                                    if (MailSub3 == "") { MailSub3 = $"{_Fun.Config.ServerId} {dr_tmp["Shift_Night"].ToString().Trim().Split(',')[0]}到{dr_tmp["Shift_Night"].ToString().Trim().Split(',')[3]} 之間, 報工可能有 延遲 提示"; }
                                                    if (int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                                    { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0).AddDays(1); }
                                                    else
                                                    { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0); }
                                                    tmp_edate = new DateTime(tmp_sdate.Year, tmp_sdate.Month, tmp_sdate.Day, int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                    if (_Fun.Config.RUNMode == '1')
                                                    {
                                                        if (d["Station_Type"].ToString() == "8")
                                                        {
                                                            #region 多工單工站
                                                            DataTable dt_tmp2 = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeS<='{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and EndTime is NULL");
                                                            if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                                                            {
                                                                DataRow dr_tmp3 = null;
                                                                foreach (DataRow d2 in dt_tmp2.Rows)
                                                                {
                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()} and PartNO='{d2["PartNO"].ToString()}' and ReportTime>5");
                                                                    if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                                    {
                                                                        int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                        if (avgReportTime > 5)
                                                                        {
                                                                            if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), tmp_sdate, tmp_edate) >= avgReportTime)
                                                                            {
                                                                                dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTime>'{Convert.ToDateTime(d2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTime<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()}  and PartNO='{d2["PartNO"].ToString()}'");
                                                                                if (dr_tmp3 == null)
                                                                                {
                                                                                    MailBody3 = $"{MailBody3}<p>工站:{d["StationNO"].ToString()}  工序編號:{d2["IndexSN"].ToString()} 製程:{d2["PP_Name"].ToString()}  料號:{d2["PartNO"].ToString()}</p>";
                                                                                    string needId = "";
                                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{d2["OrderNO"].ToString()}'");
                                                                                    if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','3','{d["StationNO"].ToString()}')"))
                                                                                    {
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            #region 單工
                                                            DataRow dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()} and PartNO='{dr_tmp2["PartNO"].ToString()}' and ReportTime>5");
                                                            if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                            {
                                                                int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                if (avgReportTime > 5)
                                                                {
                                                                    if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp2["RemarkTimeS"]), tmp_date) >= avgReportTime)
                                                                    {
                                                                        dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(dr_tmp2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()}  and PartNO='{dr_tmp2["PartNO"].ToString()}'");
                                                                        if (dr_tmp3 == null)
                                                                        {
                                                                            MailBody3 = $"{MailBody3}<p>工站:{d["StationNO"].ToString()}  工序編號:{dr_tmp2["IndexSN"].ToString()} 製程:{dr_tmp2["PP_Name"].ToString()}  料號:{dr_tmp2["PartNO"].ToString()}</p>";
                                                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{dr_tmp2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','','3','{d["StationNO"].ToString()}')"))
                                                                            {
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (d["Station_Type"].ToString() == "8")
                                                        {
                                                            #region 多工單工站
                                                            DataTable dt_tmp2 = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeS<='{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and EndTime is NULL");
                                                            if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                                                            {
                                                                DataRow dr_tmp3 = null;
                                                                foreach (DataRow d2 in dt_tmp2.Rows)
                                                                {
                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()} and PartNO='{d2["PartNO"].ToString()}' and ReportTime>5");
                                                                    if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                                    {
                                                                        int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                        if (avgReportTime > 5)
                                                                        {
                                                                            if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), tmp_sdate, tmp_edate) >= avgReportTime)
                                                                            {
                                                                                dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(d2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()}  and PartNO='{d2["PartNO"].ToString()}'");
                                                                                if (dr_tmp3 == null)
                                                                                {
                                                                                    MailBody3 = $"{MailBody3}<p>工站:{d["StationNO"].ToString()}  工序編號:{d2["IndexSN"].ToString()} 製程:{d2["PP_Name"].ToString()}  料號:{d2["PartNO"].ToString()}</p>";
                                                                                    string needId = "";
                                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{d2["OrderNO"].ToString()}'");
                                                                                    if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','3','{d["StationNO"].ToString()}')"))
                                                                                    {
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            #region 單工
                                                            DataRow dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()} and PartNO='{dr_tmp2["PartNO"].ToString()}' and ReportTime>5");
                                                            if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                            {
                                                                int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                if (avgReportTime > 5)
                                                                {
                                                                    if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp2["RemarkTimeS"]), tmp_date) >= avgReportTime)
                                                                    {
                                                                        dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(dr_tmp2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()}  and PartNO='{dr_tmp2["PartNO"].ToString()}'");
                                                                        if (dr_tmp3 == null)
                                                                        {
                                                                            MailBody3 = $"{MailBody3}<p>工站:{d["StationNO"].ToString()}  工序編號:{dr_tmp2["IndexSN"].ToString()} 製程:{dr_tmp2["PP_Name"].ToString()}  料號:{dr_tmp2["PartNO"].ToString()}</p>";
                                                                            string needId = "";
                                                                            dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_tmp2["OrderNO"].ToString()}'");
                                                                            if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{dr_tmp2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','3','{d["StationNO"].ToString()}')"))
                                                                            {
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                        }
                                        if (bool.Parse(dr_tmp["Flag_Afternoon"].ToString()) == true)
                                        {
                                            #region 第二時段 抓中間休息時間
                                            tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                            tmp_edate = new DateTime(tmp_sdate.Year, tmp_sdate.Month, tmp_sdate.Day, int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                            if (tmp_edate >= tmp_date && tmp_date >= tmp_sdate)
                                            {
                                                tmo_i = (_SFC_Common.TimeCompute2Seconds(tmp_sdate, tmp_edate)) / 2;
                                                tmo_e = _SFC_Common.TimeCompute2Seconds(tmp_sdate, tmp_date);
                                                if (tmo_e >= tmo_i)
                                                {
                                                    if (_Fun.Config.RUNMode == '2' && db.DB_GetQueryCount($"select * from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and ErrorType='15' and CONVERT(varchar(100), LogDate, 23)='{tmp_date.ToString("yyyy-MM-dd")}' and Time_N_Type='2'") > 0)
                                                    { continue; }
                                                    if (MailSub2 == "") { MailSub2 = $"{_Fun.Config.ServerId} {dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[0]}到{dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[3]} 之間, 報工可能有 延遲 提示"; }
                                                    if (int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                                    { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0).AddDays(1); }
                                                    else
                                                    { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0); }
                                                    tmp_edate = new DateTime(tmp_sdate.Year, tmp_sdate.Month, tmp_sdate.Day, int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                    if (_Fun.Config.RUNMode == '1')
                                                    {
                                                        if (d["Station_Type"].ToString() == "8")
                                                        {
                                                            #region 多工單工站
                                                            DataTable dt_tmp2 = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeS<='{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and EndTime is NULL");
                                                            if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                                                            {
                                                                DataRow dr_tmp3 = null;
                                                                foreach (DataRow d2 in dt_tmp2.Rows)
                                                                {
                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()} and PartNO='{d2["PartNO"].ToString()}' and ReportTime>5");
                                                                    if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                                    {
                                                                        int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                        if (avgReportTime > 5)
                                                                        {
                                                                            if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), tmp_sdate, tmp_edate) >= avgReportTime)
                                                                            {
                                                                                dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTime>'{Convert.ToDateTime(d2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTime<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()}  and PartNO='{d2["PartNO"].ToString()}'");
                                                                                if (dr_tmp3 == null)
                                                                                {
                                                                                    MailBody2 = $"{MailBody2}<p>工站:{d["StationNO"].ToString()}  工序編號:{d2["IndexSN"].ToString()} 製程:{d2["PP_Name"].ToString()}  料號:{d2["PartNO"].ToString()}</p>";
                                                                                    string needId = "";
                                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{d2["OrderNO"].ToString()}'");
                                                                                    if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','2','{d["StationNO"].ToString()}')"))
                                                                                    {
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            #region 單工
                                                            DataRow dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()} and PartNO='{dr_tmp2["PartNO"].ToString()}' and ReportTime>5");
                                                            if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                            {
                                                                int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                if (avgReportTime > 5)
                                                                {
                                                                    if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp2["RemarkTimeS"]), tmp_date) >= avgReportTime)
                                                                    {
                                                                        dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(dr_tmp2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()}  and PartNO='{dr_tmp2["PartNO"].ToString()}'");
                                                                        if (dr_tmp3 == null)
                                                                        {
                                                                            MailBody2 = $"{MailBody2}<p>工站:{d["StationNO"].ToString()}  工序編號:{dr_tmp2["IndexSN"].ToString()} 製程:{dr_tmp2["PP_Name"].ToString()}  料號:{dr_tmp2["PartNO"].ToString()}</p>";
                                                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{dr_tmp2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','','2','{d["StationNO"].ToString()}')"))
                                                                            {
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (d["Station_Type"].ToString() == "8")
                                                        {
                                                            #region 多工單工站
                                                            DataTable dt_tmp2 = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeS<='{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and EndTime is NULL");
                                                            if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                                                            {
                                                                DataRow dr_tmp3 = null;
                                                                foreach (DataRow d2 in dt_tmp2.Rows)
                                                                {
                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()} and PartNO='{d2["PartNO"].ToString()}' and ReportTime>5");
                                                                    if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                                    {
                                                                        int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                        if (avgReportTime > 5)
                                                                        {
                                                                            if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), tmp_sdate, tmp_edate) >= avgReportTime)
                                                                            {
                                                                                dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(d2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()}  and PartNO='{d2["PartNO"].ToString()}'");
                                                                                if (dr_tmp3 == null)
                                                                                {
                                                                                    MailBody2 = $"{MailBody2}<p>工站:{d["StationNO"].ToString()}  工序編號:{d2["IndexSN"].ToString()} 製程:{d2["PP_Name"].ToString()}  料號:{d2["PartNO"].ToString()}</p>";
                                                                                    string needId = "";
                                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{d2["OrderNO"].ToString()}'");
                                                                                    if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','2','{d["StationNO"].ToString()}')"))
                                                                                    {
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            #region 單工
                                                            DataRow dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()} and PartNO='{dr_tmp2["PartNO"].ToString()}' and ReportTime>5");
                                                            if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                            {
                                                                int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                if (avgReportTime > 5)
                                                                {
                                                                    if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp2["RemarkTimeS"]), tmp_date) >= avgReportTime)
                                                                    {
                                                                        dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(dr_tmp2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()}  and PartNO='{dr_tmp2["PartNO"].ToString()}'");
                                                                        if (dr_tmp3 == null)
                                                                        {
                                                                            MailBody2 = $"{MailBody2}<p>工站:{d["StationNO"].ToString()}  工序編號:{dr_tmp2["IndexSN"].ToString()} 製程:{dr_tmp2["PP_Name"].ToString()}  料號:{dr_tmp2["PartNO"].ToString()}</p>";
                                                                            string needId = "";
                                                                            dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_tmp2["OrderNO"].ToString()}'");
                                                                            if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{dr_tmp2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','2','{d["StationNO"].ToString()}')"))
                                                                            {
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                        }
                                        if (bool.Parse(dr_tmp["Flag_Morning"].ToString()) == true)
                                        {
                                            #region 第一時段 抓中間休息時間
                                            tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                            tmp_edate = new DateTime(tmp_sdate.Year, tmp_sdate.Month, tmp_sdate.Day, int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                            if (tmp_edate >= tmp_date && tmp_date >= tmp_sdate)
                                            {
                                                tmo_i = (_SFC_Common.TimeCompute2Seconds(tmp_sdate, tmp_edate)) / 2;
                                                tmo_e = _SFC_Common.TimeCompute2Seconds(tmp_sdate, tmp_date);
                                                if (tmo_e >= tmo_i)
                                                {
                                                    if (_Fun.Config.RUNMode == '2' && db.DB_GetQueryCount($"select * from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and ErrorType='15' and CONVERT(varchar(100), LogDate, 23)='{tmp_date.ToString("yyyy-MM-dd")}' and Time_N_Type='1'") > 0)
                                                    { continue; }
                                                    if (MailSub1 == "") { MailSub1 = $"{_Fun.Config.ServerId} {dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[0]}到{dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[3]} 之間, 報工可能有 延遲 提示"; }
                                                    if (int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                                    { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0).AddDays(1); }
                                                    else
                                                    { tmp_sdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0); }
                                                    tmp_edate = new DateTime(tmp_sdate.Year, tmp_sdate.Month, tmp_sdate.Day, int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_tmp["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                    if (_Fun.Config.RUNMode == '1')
                                                    {
                                                        if (d["Station_Type"].ToString() == "8")
                                                        {
                                                            #region 多工單工站
                                                            DataTable dt_tmp2 = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeS<='{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and EndTime is NULL");
                                                            if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                                                            {
                                                                DataRow dr_tmp3 = null;
                                                                foreach (DataRow d2 in dt_tmp2.Rows)
                                                                {
                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()} and PartNO='{d2["PartNO"].ToString()}' and ReportTime>5");
                                                                    if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                                    {
                                                                        int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                        if (avgReportTime > 5)
                                                                        {
                                                                            if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), tmp_sdate, tmp_edate) >= avgReportTime)
                                                                            {
                                                                                dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTime>'{Convert.ToDateTime(d2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTime<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()}  and PartNO='{d2["PartNO"].ToString()}'");
                                                                                if (dr_tmp3 == null)
                                                                                {
                                                                                    MailBody1 = $"{MailBody1}<p>工站:{d["StationNO"].ToString()}  工序編號:{d2["IndexSN"].ToString()} 製程:{d2["PP_Name"].ToString()}  料號:{d2["PartNO"].ToString()}</p>";
                                                                                    string needId = "";
                                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{d2["OrderNO"].ToString()}'");
                                                                                    if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','1','{d["StationNO"].ToString()}')"))
                                                                                    {
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            #region 單工
                                                            DataRow dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()} and PartNO='{dr_tmp2["PartNO"].ToString()}' and ReportTime>5");
                                                            if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                            {
                                                                int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                if (avgReportTime > 5)
                                                                {
                                                                    if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp2["RemarkTimeS"]), tmp_date) >= avgReportTime)
                                                                    {
                                                                        dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(dr_tmp2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()}  and PartNO='{dr_tmp2["PartNO"].ToString()}'");
                                                                        if (dr_tmp3 == null)
                                                                        {
                                                                            MailBody1 = $"{MailBody1}<p>工站:{d["StationNO"].ToString()}  工序編號:{dr_tmp2["IndexSN"].ToString()} 製程:{dr_tmp2["PP_Name"].ToString()}  料號:{dr_tmp2["PartNO"].ToString()}</p>";
                                                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{dr_tmp2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','','1','{d["StationNO"].ToString()}')"))
                                                                            {
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (d["Station_Type"].ToString() == "8")
                                                        {
                                                            #region 多工單工站
                                                            DataTable dt_tmp2 = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and RemarkTimeS is not NULL and RemarkTimeS<='{tmp_edate.ToString("yyyy/MM/dd HH:mm:ss")}' and EndTime is NULL");
                                                            if (dt_tmp2 != null && dt_tmp2.Rows.Count > 0)
                                                            {
                                                                DataRow dr_tmp3 = null;
                                                                foreach (DataRow d2 in dt_tmp2.Rows)
                                                                {
                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()} and PartNO='{d2["PartNO"].ToString()}' and ReportTime>5");
                                                                    if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                                    {
                                                                        int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                        if (avgReportTime > 5)
                                                                        {
                                                                            if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), tmp_sdate, tmp_edate) >= avgReportTime)
                                                                            {
                                                                                dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(d2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{d2["PP_Name"].ToString()}' and IndexSN={d2["IndexSN"].ToString()}  and PartNO='{d2["PartNO"].ToString()}'");
                                                                                if (dr_tmp3 == null)
                                                                                {
                                                                                    MailBody1 = $"{MailBody1}<p>工站:{d["StationNO"].ToString()}  工序編號:{d2["IndexSN"].ToString()} 製程:{d2["PP_Name"].ToString()}  料號:{d2["PartNO"].ToString()}</p>";
                                                                                    string needId = "";
                                                                                    dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{d2["OrderNO"].ToString()}'");
                                                                                    if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                                    if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','1','{d["StationNO"].ToString()}')"))
                                                                                    {
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            #region 單工
                                                            DataRow dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT TOP {_Fun.Config.AdminKey03} avg(ReportTime) as AVGTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()} and PartNO='{dr_tmp2["PartNO"].ToString()}' and ReportTime>5");
                                                            if (dr_tmp3 != null && !dr_tmp3.IsNull("AVGTime") && dr_tmp3["AVGTime"].ToString().Trim() != "")
                                                            {
                                                                int avgReportTime = int.Parse(dr_tmp3["AVGTime"].ToString());
                                                                if (avgReportTime > 5)
                                                                {
                                                                    if (_SFC_Common.TimeCompute2Seconds_BY_Calendar(db, d["CalendarName"].ToString(), Convert.ToDateTime(dr_tmp2["RemarkTimeS"]), tmp_date) >= avgReportTime)
                                                                    {
                                                                        dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}' and LOGDateTimeID>'{Convert.ToDateTime(dr_tmp2["RemarkTimeS"]).ToString("yyyy/MM/dd HH:mm:ss")}' and LOGDateTimeID<='{tmp_date.ToString("yyyy/MM/dd HH:mm:ss")}' and PP_Name='{dr_tmp2["PP_Name"].ToString()}' and IndexSN={dr_tmp2["IndexSN"].ToString()}  and PartNO='{dr_tmp2["PartNO"].ToString()}'");
                                                                        if (dr_tmp3 == null)
                                                                        {
                                                                            MailBody1 = $"{MailBody1}<p>工站:{d["StationNO"].ToString()}  工序編號:{dr_tmp2["IndexSN"].ToString()} 製程:{dr_tmp2["PP_Name"].ToString()}  料號:{dr_tmp2["PartNO"].ToString()}</p>";
                                                                            string needId = "";
                                                                            dr_tmp3 = db.DB_GetFirstDataByDataRow($"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr_tmp2["OrderNO"].ToString()}'");
                                                                            if (dr_tmp3 != null) { needId = dr_tmp3["NeedId"].ToString(); }
                                                                            if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,Time_N_Type,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{dr_tmp2["SimulationId"].ToString()}','15','{tmp_date.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{needId}','1','{d["StationNO"].ToString()}')"))
                                                                            {
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                        }
                                    }
                                }
                                #region 發出mail
                                if (MailSub1 != "")
                                {
                                    if (MailBody1 != "") { _Fun.Mail_Send(_Fun.Config.SendMonitorMail15.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail15.Split(',')[1].Split(';'), MailSub1, MailBody1, null, false); }
                                }
                                else if (MailSub2 != "")
                                {
                                    if (MailBody2 != "") { _Fun.Mail_Send(_Fun.Config.SendMonitorMail15.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail15.Split(',')[1].Split(';'), MailSub2, MailBody2, null, false); }
                                }
                                else if (MailSub3 != "")
                                {
                                    if (MailBody3 != "") { _Fun.Mail_Send(_Fun.Config.SendMonitorMail15.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail15.Split(',')[1].Split(';'), MailSub3, MailBody3, null, false); }
                                }
                                else if (MailSub4 != "")
                                {
                                    if (MailBody4 != "") { _Fun.Mail_Send(_Fun.Config.SendMonitorMail15.Split(',')[0].Split(';'), _Fun.Config.SendMonitorMail15.Split(',')[1].Split(';'), MailSub4, MailBody4, null, false); }
                                }
                                #endregion

                            }
                            #region 隔日解除
                            sql = $@"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and ActionType='' and ErrorType='15' and CONVERT(varchar(100), LogDate, 23)<='{DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd")}'";
                            dr_tmp = db.DB_GetFirstDataByDataRow(sql);
                            if (dr_tmp != null)
                            {
                                _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='1' where ServerId='{_Fun.Config.ServerId}' and ActionType='' and ErrorType='15' and CONVERT(varchar(100), LogDate, 23)<='{DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd")}'");
                            }
                            #endregion
                        }
                        #endregion
                        _Fun._a12 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 監看工站是否如計畫 16  //###??? 同19 取消
                        /*
                        #region 解除
                        sql = $"select * from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and ErrorType='16' and ActionType=''";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                dr_tmp = db.DB_GetFirstDataByDataRow($"select a.*,b.NeedQTY as STOT,(b.Detail_QTY+b.Detail_Fail_QTY) as PTOT FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a,SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as b where a.SimulationId='{d["SimulationId"].ToString()}' and b.SimulationId=a.SimulationId and a.NeedId='{d["NeedId"].ToString()}'");
                                if (dr_tmp != null)
                                {
                                    if (dr_tmp["IsOK"].ToString() == "1" || int.Parse(dr_tmp["PTOT"].ToString()) >= int.Parse(dr_tmp["STOT"].ToString()))
                                    {
                                        if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='0',ActionLogDate='{logdate}' where Id='{d["Id"].ToString()}'"))
                                        {
                                            //###???還不知要通知誰
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        sql = $@"SELECT a.*,b.NeedQTY as STOT,(b.Detail_QTY+b.Detail_Fail_QTY) as PTOT FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                join SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as b on b.SimulationId=a.SimulationId
                                join SoftNetSYSDB.[dbo].[APS_NeedData] as c on c.Id=a.NeedId and c.State='6' and c.ServerId='{_Fun.Config.ServerId}'
                                where a.IsOK='0' and a.SimulationDate<'{logdate}' and a.Source_StationNO is not null";
                        dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (int.Parse(d["STOT"].ToString())> int.Parse(d["PTOT"].ToString()))
                                {
                                    if (db.DB_GetQueryCount($"select SimulationId,ActionType from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["Source_StationNO"].ToString()}' and ErrorType='16' and ActionType=''")<=0)
                                    {
                                        if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,ErrorType,LogDate,NeedId,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d["SimulationId"].ToString()}','16','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d["NeedId"].ToString()}','{d["Source_StationNO"].ToString()}')"))
                                        {
                                            #region 通知網頁更新
                                            var service = _Fun.DiBox?.GetService(typeof(SNWebSocketService));
                                            if (service != null)
                                            {
                                                var webSocketService = (SNWebSocketService)service;
                                                List<rmsConectUserData> snapshot;
                                                lock (webSocketService.lock__WebSocketList)
                                                {
                                                    snapshot = new List<rmsConectUserData>(webSocketService._WebSocketList.Values);
                                                }
                                                foreach (var r in snapshot)
                                                {
                                                    if (r != null && r.socket != null)
                                                    {
                                                        //###???errorType 暫時寫死
                                                        webSocketService.Send(r.socket, $"SimulatioStatusChange,'16',{d["SimulationId"].ToString()}");
                                                    }
                                                }
                                            }
                                            #endregion
                                        }
                                    }
                                }
                            }
                        }
                        */

                        #endregion
                        _Fun._a13 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        //###???
                        #region 監看料件安全量,在制量,再途量,工站移轉量所有數據是否正常與合理 17


                        #endregion
                        _Fun._a14 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        #region 工站應停止是否未停止(影響有效值) 18 判斷跨 Shift 1階  (寫入 APS_WarningData 等時間到處理)
                        {
                            DataRow dr_PP_Station = null;
                            DateTime tmp_date = DateTime.Now;
                            DateTime wdate = DateTime.Now;
                            DateTime etime = DateTime.Now;
                            DateTime stime = DateTime.Now;
                            string[] comp = null;
                            #region 單工站
                            sql = $@"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and State='1' and Config_MutiWO='0' and StationNO!='{_Fun.Config.OutPackStationName}'";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                string needId = "";
                                bool isRUN = true;
                                foreach (DataRow d in dt_tmp.Rows)
                                {
                                    //查目前是哪個Shift, 是否為不上班狀態
                                    isRUN = true;
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WarningData] where ErrorType='18' and IsDEL='0' and ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}'");
                                    if (dr_tmp == null)
                                    {
                                        dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}'");
                                        sql = $@"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_PP_Station["CalendarName"].ToString()}' and Holiday<='{tmp_date.ToString("MM/dd/yyyy")}' order by Holiday desc";
                                        DataTable dt = db.DB_GetData(sql);
                                        if (dt != null && dt.Rows.Count > 0)
                                        {
                                            DataRow d2 = dt.Rows[0];
                                            #region 判斷是否為前一天晚班
                                            if (int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                            {
                                                comp = d2["Shift_Night"].ToString().Trim().Split(',');
                                                etime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1);
                                                stime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                                if (etime >= tmp_date && tmp_date >= stime) { isRUN = false; }
                                            }
                                            if (isRUN && int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                            {
                                                comp = d2["Shift_Graveyard"].ToString().Trim().Split(',');
                                                etime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1);
                                                stime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                                if (etime >= tmp_date && tmp_date >= stime) { isRUN = false; }
                                            }
                                            if (!isRUN || Convert.ToDateTime(dt.Rows[0]["Holiday"]).ToString("yyyy-MM-dd") != tmp_date.ToString("yyyy-MM-dd"))
                                            {
                                                if (dt.Rows.Count > 1) { d2 = dt.Rows[1]; }
                                            }
                                            #endregion

                                            needId = "";
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}'");
                                            if (dr_tmp != null) { needId = dr_tmp["NeedId"].ToString(); }
                                            if (d2 != null)
                                            {
                                                #region Flag_Morning
                                                if (tmp_date > new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), 0))
                                                {
                                                    if (!bool.Parse(d2["Flag_Afternoon"].ToString()))
                                                    {
                                                        wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(d2["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[1]), 0).AddMinutes(-5);
                                                        _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','{d["PartNO"].ToString()}','{d["OP_NO"].ToString()}','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                        continue;
                                                    }

                                                }
                                                #endregion
                                                #region Flag_Afternoon
                                                if (tmp_date > new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(d2["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), 0))
                                                {
                                                    if (!bool.Parse(d2["Flag_Night"].ToString()))
                                                    {
                                                        wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddMinutes(-5);
                                                        _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','','{d["OP_NO"].ToString()}','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                        continue;
                                                    }

                                                }
                                                #endregion
                                                #region Shift_Night
                                                if (int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                                {
                                                    comp = d2["Shift_Night"].ToString().Trim().Split(',');
                                                    stime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                                    etime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1);
                                                    if (etime >= tmp_date && tmp_date >= stime)
                                                    {
                                                        if (!bool.Parse(d2["Flag_Graveyard"].ToString()))
                                                        {
                                                            wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddDays(1).AddMinutes(-5);
                                                            _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','','{d["OP_NO"].ToString()}','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                            continue;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                    if (tmp_date > wdate)
                                                    {
                                                        if (!bool.Parse(d2["Flag_Graveyard"].ToString()))
                                                        {
                                                            if (int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[0]))
                                                            { wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddDays(1).AddMinutes(-5); }
                                                            else
                                                            { wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddMinutes(-5); }
                                                            _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','','{d["OP_NO"].ToString()}','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                            continue;
                                                        }
                                                    }
                                                }
                                                #endregion
                                                #region Shift_Graveyard
                                                if (int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                                {
                                                    comp = d2["Shift_Graveyard"].ToString().Trim().Split(',');
                                                    stime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                                    etime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1);
                                                    if (etime >= tmp_date && tmp_date >= stime)
                                                    {
                                                        d2 = dt.Rows[0];
                                                        if (!bool.Parse(d2["Flag_Morning"].ToString()))
                                                        {
                                                            wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddDays(1).AddMinutes(-5);
                                                            _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','','{d["OP_NO"].ToString()}','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                            continue;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                    if (tmp_date > wdate)
                                                    {
                                                        d2 = dt.Rows[0];
                                                        if (!bool.Parse(d2["Flag_Morning"].ToString()))
                                                        {
                                                            wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddDays(1).AddMinutes(-5);
                                                            _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','','{d["OP_NO"].ToString()}','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                            continue;
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

                            #region 多工單工站
                            sql = $@"SELECT * FROM SoftNetMainDB.[dbo].[ManufactureII] where ServerId='{_Fun.Config.ServerId}' and StartTime is not NULL and RemarkTimeS is not NULL and RemarkTimeE is NULL and EndTime is NULL order by StationNO";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                string needId = "";
                                bool isRUN = true;
                                foreach (DataRow d in dt_tmp.Rows)
                                {
                                    //查目前是哪個Shift, 是否為不上班狀態
                                    isRUN = true;
                                    dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_WarningData] where ErrorType='18' and IsDEL='0' and ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}'");
                                    if (dr_tmp == null)
                                    {
                                        dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["StationNO"].ToString()}'");
                                        sql = $@"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr_PP_Station["CalendarName"].ToString()}' and Holiday<='{tmp_date.ToString("MM/dd/yyyy")}' order by Holiday desc";
                                        DataTable dt = db.DB_GetData(sql);
                                        if (dt != null && dt.Rows.Count > 0)
                                        {
                                            DataRow d2 = dt.Rows[0];
                                            #region 判斷是否為前一天晚班
                                            if (int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                            {
                                                comp = d2["Shift_Night"].ToString().Trim().Split(',');
                                                etime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1);
                                                stime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                                if (etime >= tmp_date && tmp_date >= stime) { isRUN = false; }
                                            }
                                            if (isRUN && int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                            {
                                                comp = d2["Shift_Graveyard"].ToString().Trim().Split(',');
                                                etime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1);
                                                stime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                                if (etime >= tmp_date && tmp_date >= stime) { isRUN = false; }
                                            }
                                            if (!isRUN || Convert.ToDateTime(dt.Rows[0]["Holiday"]).ToString("yyyy-MM-dd") != tmp_date.ToString("yyyy-MM-dd"))
                                            {
                                                d2 = dt.Rows[1];
                                            }
                                            #endregion

                                            needId = "";
                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}'");
                                            if (dr_tmp != null) { needId = dr_tmp["NeedId"].ToString(); }
                                            if (d2 != null)
                                            {
                                                #region Flag_Morning
                                                if (tmp_date > new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), 0))
                                                {
                                                    if (!bool.Parse(d2["Flag_Afternoon"].ToString()))
                                                    {
                                                        wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(d2["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[1]), 0).AddMinutes(-5);
                                                        _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','{d["PartNO"].ToString()}','系統指派','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                        continue;
                                                    }

                                                }
                                                #endregion
                                                #region Flag_Afternoon
                                                if (tmp_date > new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(d2["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), 0))
                                                {
                                                    if (!bool.Parse(d2["Flag_Night"].ToString()))
                                                    {
                                                        wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddMinutes(-5);
                                                        _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','{d["PartNO"].ToString()}','系統指派','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                        continue;
                                                    }

                                                }
                                                #endregion
                                                #region Shift_Night
                                                if (int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                                {
                                                    comp = d2["Shift_Night"].ToString().Trim().Split(',');
                                                    stime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                                    etime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1);
                                                    if (etime >= tmp_date && tmp_date >= stime)
                                                    {
                                                        if (!bool.Parse(d2["Flag_Graveyard"].ToString()))
                                                        {
                                                            wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddDays(1).AddMinutes(-5);
                                                            _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','{d["PartNO"].ToString()}','系統指派','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                            continue;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                    if (tmp_date > wdate)
                                                    {
                                                        if (!bool.Parse(d2["Flag_Graveyard"].ToString()))
                                                        {
                                                            if (int.Parse(d2["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[0]))
                                                            { wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddDays(1).AddMinutes(-5); }
                                                            else
                                                            { wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddMinutes(-5); }
                                                            _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','{d["PartNO"].ToString()}','系統指派','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                            continue;
                                                        }
                                                    }
                                                }
                                                #endregion
                                                #region Shift_Graveyard
                                                if (int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[0].Split(':')[0]) > int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]))
                                                {
                                                    comp = d2["Shift_Graveyard"].ToString().Trim().Split(',');
                                                    stime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0, 0);
                                                    etime = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0, 0).AddDays(1);
                                                    if (etime >= tmp_date && tmp_date >= stime)
                                                    {
                                                        d2 = dt.Rows[0];
                                                        if (!bool.Parse(d2["Flag_Morning"].ToString()))
                                                        {
                                                            wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddDays(1).AddMinutes(-5);
                                                            _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','{d["PartNO"].ToString()}','系統指派','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                            continue;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(d2["Shift_Graveyard"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                    if (tmp_date > wdate)
                                                    {
                                                        d2 = dt.Rows[0];
                                                        if (!bool.Parse(d2["Flag_Morning"].ToString()))
                                                        {
                                                            wdate = new DateTime(tmp_date.Year, tmp_date.Month, tmp_date.Day, int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(d2["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[1]), 0).AddDays(1).AddMinutes(-5);
                                                            _ = db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                                ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','18','{needId}','{d["SimulationId"].ToString()}','{d["StationNO"].ToString()}','{d["OrderNO"].ToString()}','{d["PartNO"].ToString()}','系統指派','{wdate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");
                                                            continue;
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

                        #endregion
                        _Fun._a15 = threadLoopTime.ElapsedMilliseconds;
                    }
                    if (!_Fun.Is_Thread_ForceClose)
                    {
                        //###???
                        #region 警示 APS_Simulation 的 生產排程計畫明細預計完成日是否遠離偏離完成日 19
                        {
                            /*
                            #region 解除
                            sql = $@"SELECT a.* FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] as a
                    join SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as b on b.SimulationId=a.SimulationId and b.APS_StationNO=a.StationNO and b.NeedQTY<=(b.Detail_QTY+b.Detail_Fail_QTY)
                    where a.ServerId='{_Fun.Config.ServerId}' and a.ActionType='' and a.ErrorType='19'";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                foreach (DataRow d in dt_tmp.Rows)
                                {
                                    if (db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].APS_Simulation_ErrorData SET ActionType='0',ActionLogDate='{logdate}' where Id='{d["Id"].ToString()}'"))
                                    {
                                        db.DB_SetData($"delete SoftNetSYSDB.[dbo].[APS_WarningData] where ErrorType='19' and ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["StationNO"].ToString()}'");
                                        //###???還不知要通知誰
                                    }
                                }

                            }
                            #endregion

                            #region 成立
                            sql = $@"SELECT a.*,b.NeedQTY as Sqty,(b.Detail_QTY+b.Detail_Fail_QTY) as Kqty FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                join SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as b on b.SimulationId=a.SimulationId and b.NoStation='0'
                                join SoftNetSYSDB.[dbo].[APS_NeedData] as c on c.Id=a.NeedId and c.ServerId='{_Fun.Config.ServerId}' and c.State='6'
                                where DATEADD(n,{isARGs10_offset},a.SimulationDate)<'{logdate}' and a.DOCNumberNO!='' and a.IsOK='0' and a.Source_StationNO  is not null and (a.Class='4' or a.Class='5') and a.Source_StationNO is not null order by a.SimulationDate";
                            dt_tmp = db.DB_GetData(sql);
                            if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                            {
                                try
                                {
                                    foreach (DataRow d in dt_tmp.Rows)
                                    {
                                        if (!d.IsNull("Sqty") && !d.IsNull("Kqty") && int.Parse(d["Sqty"].ToString()) > int.Parse(d["Kqty"].ToString()))
                                        {
                                            if (db.DB_GetQueryCount($"select SimulationId from SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["Source_StationNO"].ToString()}' and ErrorType='19'") <= 0)
                                            {
                                                if (db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] (Id,ServerId,SimulationId,DOCNumberNO,ErrorType,LogDate,NeedId,StationNO) VALUES ('{_Str.NewId('E')}','{_Fun.Config.ServerId}','{d["SimulationId"].ToString()}','{d["DOCNumberNO"].ToString()}','19','{logdate}','{d["NeedId"].ToString()}','{d["Source_StationNO"].ToString()}')"))
                                                {
                                                    //var webSocketService = (SNWebSocketService)_Fun.DiBox.GetService(typeof(SNWebSocketService));
                                                    if (WebSocketServiceOJB != null)
                                                    {
                                                        #region 安排主動干涉觸發事件
                                                        if (db.DB_GetQueryCount($"select Id from SoftNetSYSDB.[dbo].[APS_WarningData] where ErrorType='19' and ServerId='{_Fun.Config.ServerId}' and SimulationId='{d["SimulationId"].ToString()}' and StationNO='{d["Source_StationNO"].ToString()}'") <= 0)
                                                        {

                                                            DataRow dr_APS_Simulation = db.DB_GetFirstDataByDataRow($"select *,(NeedQTY+SafeQTY) as APS_SQTY from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{d["SimulationId"].ToString()}' and IsOK='0'");
                                                            DataRow dr_APS_NeedData = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{_Fun.Config.ServerId}' and Id='{dr_APS_Simulation["NeedId"].ToString()}'");
                                                            bool isOK = true;
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
                                                                        if (pQTY >= sQTY) { isOK = false; }
                                                                    }
                                                                }
                                                                else { isOK = false; }
                                                            }
                                                            else { isOK = false; }
                                                            #endregion
                                                            if (isOK)
                                                            {
                                                                int dltime = 0;
                                                                DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_APS_Simulation["Source_StationNO"]}'");
                                                                #region 計算要順延多小時間
                                                                if (type == '1') { sQTY -= pQTY; }
                                                                dltime = (int.Parse(dr_APS_Simulation["Math_UseTime"].ToString()) / sQTY) * int.Parse(dr_APS_Simulation["Math_UseOPCount"].ToString());
                                                                if (!d.IsNull("StationNO_Merge") && d["StationNO_Merge"].ToString() != "")
                                                                {
                                                                    int count = d["StationNO_Merge"].ToString().Split(',').Length-1;
                                                                    if (count > 0) { dltime = dltime / count; }

                                                                }
                                                                dltime += (isARGs10_offset * 60);
                                                                DateTime dltime_startDate = WebSocketServiceOJB.TimeCompute2DateTime_BY_ReturnNextShift(db, dr_PP_Station["CalendarName"].ToString(), simulationDate, dltime).AddMinutes(-5);
                                                                if (DateTime.Now > dltime_startDate)
                                                                {
                                                                    dltime_startDate = DateTime.Now;
                                                                    //if (sQTY > 0)
                                                                    //{ dltime = (int.Parse(dr_APS_Simulation["Math_UseTime"].ToString()) / sQTY) * int.Parse(dr_APS_Simulation["Math_UseOPCount"].ToString()); }
                                                                    //else { dltime = (isARGs10_offset * 60); }
                                                                }
                                                                #endregion

                                                                db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[APS_WarningData] ([Id],[ServerId],[ErrorType],[NeedId],[SimulationId],[StationNO],[DOCNumberNO],[PartNO],[OP_NO],[WarningDate]) VALUES
                                                            ('{_Str.NewId('W')}','{_Fun.Config.ServerId}','19','{d["NeedId"].ToString()}','{d["SimulationId"].ToString()}','{d["Source_StationNO"].ToString()}','{d["DOCNumberNO"].ToString()}','{d["PartNO"].ToString()}','','{dltime_startDate.ToString("yyyy-MM-dd HH:mm:ss.fff")}')");

                                                            }
                                                        }
                                                        #endregion

                                                                                                    lock (WebSocketServiceOJB.lock__WebSocketList)
                                                                {
                                                        foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                                        {
                                                            if (r.Key != null && r.Value.socket != null)
                                                            {
                                                                //###???errorType 暫時寫死
                                                                WebSocketServiceOJB.Send(r.Value.socket, $"SimulatioStatusChange,'19',{d["SimulationId"].ToString()},");
                                                            }
                                                        }
                            }
                                                    }
                                                }
                                            }

                                        }

                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"監看 APS_Simulation 的 預計完成日是否未達 19 {ex.Message} {ex.StackTrace}", true);
                                }
                            }
                            #endregion
                            */
                        }
                        #endregion
                        _Fun._a16 = threadLoopTime.ElapsedMilliseconds;
                    }
                    //}


                    threadLoopTime.Stop();

                    if (_Fun.Is_Thread_ForceClose) { IsWork = false; break; }

                    // 計算下次執行前的等待時間
                    int elapsedMs = (int)threadLoopTime.ElapsedMilliseconds;
                    next_time = Math.Max(_Fun.Config.RunTimeServerLoopTime - elapsedMs, 30000);
                    await Task.Delay(next_time, cancellationToken);
                    if (isFirstRun) { isFirstRun = false; }
                }
            }
            catch (OperationCanceledException)
            {
                // 正常停止，不需記錄錯誤
                ++RMSDBErrorCount;
            }
            catch (Exception ex)
            {
                ++RMSDBErrorCount;
                string _s = "";
                await _Log.ErrorAsync($"RUNTimeServer.cs SfcTimerloopthread_Tick Exception停止Thread: {ex.Message} {ex.StackTrace}", true);
            }
            _Fun.Is_RUNTimeServer_Thread_State[3] = false;
            //SFC_FUN.Dispose();
            db.Dispose();
        }
        private async void Check_ElectronicTagsServer_Tick(CancellationToken cancellationToken = default)//每10分鐘確認一次
        {
            //###??? 若多台發射器,目前只探測一台
            //###??? 將來改問state並記錄電池容量
            Ping pingSender = new Ping();
            try
            {
                if (_Fun.Config.ElectronicTagsURL.Trim() != "")
                {
                    string urlIP = _Fun.Config.ElectronicTagsURL.Split(':')[0];
                    if (urlIP.Trim() == "") { _Fun.Is_Tag_Connect = false; _Fun.Is_RUNTimeServer_Thread_State[4] = false; pingSender.Dispose(); return; }
                    while (IsWork && !cancellationToken.IsCancellationRequested)
                    {
                        if (_Fun.Has_Tag_httpClient)
                        {
                            PingReply reply = pingSender.Send(urlIP, 3000);
                            if (reply.Status == IPStatus.Success)
                            {
                                if (!_Fun.Is_Tag_Connect)
                                { await _Log.ErrorAsync($"網路訊號復歸 電子標籤主機. IP={urlIP}", true); }
                                _Fun.Is_Tag_Connect = true;
                            }
                            else 
                            { 
                                //###???暫時拿掉
                                //_Fun.Is_Tag_Connect = false;
                                //await _Log.ErrorAsync($"電子標籤主機網路中斷. IP={urlIP}", true);
                            }
                        }
                        await Task.Delay(600000, cancellationToken);
                    }
                }
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"RUNTimeServer.cs DeviceConnectCheck_Tick Exception停止Thread: {ex.Message} {ex.StackTrace}", true);
                _Fun.Is_Tag_Connect = false;
            }
            _Fun.Is_RUNTimeServer_Thread_State[4] = false;
            pingSender.Dispose();
        }

        private DateTime _beforeTime = DateTime.Now;//一天只做一次的變數

        private async void SfcTimerloopautoRUN_Json_Tick(CancellationToken cancellationToken = default)//自動更新工站計畫需求電子標籤
        {
            string err_MEG = "";
            DBADO db = new DBADO("1", _Fun.Config.Db, ref err_MEG);
            if (err_MEG == "")
            {
                db.Error += new DBADO.ERROR(DBBase_OnException);
            }
            try
            {
                Stopwatch threadLoopTime = new Stopwatch();
                int next_time = 0;
                string sql = "";
                string logdate = "";
                while (IsWork && !cancellationToken.IsCancellationRequested)
                {
                    if (_Fun.Is_Thread_For_Test)
                    {
                        //目前無作用
                    }
                    logdate = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff");
                    threadLoopTime.Restart();

                    #region 判斷是否更新 工站累計量
                    DataTable dt_LabelStateINFO = db.DB_GetData($"SELECT a.* FROM SoftNetMainDB.[dbo].[LabelStateINFO] as a,SoftNetMainDB.[dbo].[Manufacture] as b where b.ServerId='{_Fun.Config.ServerId}' and a.Type='1' and a.OrderNO!='' and b.State='1' and a.macID=b.Config_macID");
                    if (dt_LabelStateINFO != null && dt_LabelStateINFO.Rows.Count > 0)
                    {
                        DataRow totalData = null;
                        foreach (DataRow d in dt_LabelStateINFO.Rows)
                        {
                            totalData = _Fun.GetAvgCTWTandTotalOutput(db, false, d["OrderNO"].ToString(), d["StationNO"].ToString(), d["IndexSN"].ToString());
                            string dis_DetailQTY = "0";
                            if (totalData != null)
                            {
                                dis_DetailQTY = totalData["TotalOutput"].ToString().Trim();
                            }
                            if (dis_DetailQTY!= d["QTY"].ToString().Trim())
                            {
                                _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set QTY='{dis_DetailQTY}',IsUpdate='0' where ServerId='{_Fun.Config.ServerId}' and macID='{d["macID"].ToString()}'");
                            }
                        }
                    }
                    #endregion

                    #region 自動派工監測 工站的第一次自動派工, 只針對主站派工
                    if (_Fun.Config.IsAutoDispatch)
                    {
                        sql = $@"SELECT a.NeedId,a.PartNO,a.CalendarDate,a.SimulationId,a.macID,a.NeedQTY as QTY,b.Apply_PP_Name,b.Math_EfficientCT,b.Math_StandardCT,b.DOCNumberNO,b.Source_StationNO,b.Source_StationNO_IndexSN,b.Source_StationNO_Custom_IndexSN,b.Source_StationNO_Custom_DisplayName,c.PartName FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a
                            join SoftNetSYSDB.[dbo].[APS_Simulation] as b on b.SimulationId=a.SimulationId and b.DOCNumberNO!='' and b.Source_StationNO is not NULL and b.Source_StationNO!='{_Fun.Config.OutPackStationName}' 
                            join SoftNetMainDB.[dbo].[Material] as c on c.PartNO=a.PartNO
                            join SoftNetSYSDB.[dbo].[APS_NeedData] as d on d.State='6' and a.NeedId=d.Id
                            where d.ServerId='{_Fun.Config.ServerId}' and a.DOCNumberNO='' and (a.Class='4' or a.Class='5') and CONVERT(varchar(100), a.CalendarDate, 120)<='{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}' order by b.Source_StationNO,a.CalendarDate";
                        DataTable dt_tmp = db.DB_GetData(sql);
                        if (dt_tmp != null && dt_tmp.Rows.Count > 0)
                        {
                            DataRow dr_LabelStateINFO = null;
                            DataRow dr_Staion = null;
                            string verLab = "";
                            DataRow tmp_dr = null;
                            string tmp_logStaion = "";
                            foreach (DataRow d in dt_tmp.Rows)
                            {
                                if (tmp_logStaion != d["Source_StationNO"].ToString())
                                {
                                    tmp_logStaion = d["Source_StationNO"].ToString();
                                    tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["Source_StationNO"].ToString()}' and State!='1'");
                                    if (tmp_dr != null && _Fun.Config.IsAutoDispatch_IsAutoUpdate_WO)
                                    {
                                        if (tmp_dr["SimulationId"].ToString() != "")
                                        {
                                            if (db.DB_GetQueryCount($"SELECT * from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where (Detail_QTY+Detail_Fail_QTY)>=NeedQTY and SimulationId='{tmp_dr["SimulationId"].ToString()}'") < 0)
                                            { continue; }
                                        }
                                        _ = db.DB_SetData($"UPDATE SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] SET APS_StationNO='{d["Source_StationNO"].ToString()}',DOCNumberNO='{d["DOCNumberNO"].ToString()}' where SimulationId='{d["SimulationId"].ToString()}'");
                                        if (!bool.Parse(tmp_dr["Config_MutiWO"].ToString()))
                                        { _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET OrderNO='{d["DOCNumberNO"].ToString()}',IndexSN={d["Source_StationNO_IndexSN"].ToString()},OP_NO='排程指派',Station_Custom_IndexSN='{d["Source_StationNO_Custom_IndexSN"].ToString()}',StationNO_Custom_DisplayName='{d["Source_StationNO_Custom_DisplayName"].ToString()}',Master_PP_Name='{d["Apply_PP_Name"].ToString()}',PP_Name='{d["Apply_PP_Name"].ToString()}',PartNO='{d["PartNO"].ToString()}',SimulationId='{d["SimulationId"].ToString()}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{d["Source_StationNO"].ToString()}'"); }
                                        else
                                        {
                                            if (db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[ManufactureII] ([Id],[StationNO],[ServerId],[OrderNO],[Master_PP_Name],[PP_Name],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],[PartNO],[SimulationId],PNQTY)
                                                    VALUES ('{_Str.NewId('C')}','{d["Source_StationNO"].ToString()}','{_Fun.Config.ServerId}','{d["DOCNumberNO"].ToString()}','{d["Apply_PP_Name"].ToString()}','{d["Apply_PP_Name"].ToString()}',{d["Source_StationNO_IndexSN"].ToString()},'{d["Source_StationNO_Custom_IndexSN"].ToString()}','{d["Source_StationNO_Custom_DisplayName"].ToString()}','{d["PartNO"].ToString()}','{d["SimulationId"].ToString()}',{d["QTY"].ToString()})"))
                                            {
                                                _ = SendWebSocketClent_INFO($"SendALLClient,HasWeb_Id_Change,STView2Work_PageReload,{tmp_logStaion}");
                                                //#region 通知網頁更新
                                                //try
                                                //{

                                                //    //var webSocketService = (SNWebSocketService)_Fun.DiBox.GetService(typeof(SNWebSocketService));
                                                //    if (WebSocketServiceOJB != null)
                                                //    {
                                                //        lock (WebSocketServiceOJB.lock__WebSocketList)
                                                //        {
                                                //            foreach (KeyValuePair<string, rmsConectUserData> r in WebSocketServiceOJB._WebSocketList)
                                                //            {
                                                //                if (r.Key != null && r.Value.socket != null)
                                                //                {
                                                //                    WebSocketServiceOJB.Send(r.Value.socket, $"HasWeb_Id_Change,STView2Work_PageReload,{tmp_logStaion}");
                                                //                }
                                                //            }
                                                //        }
                                                //    }
                                                //}
                                                //catch (Exception ex)
                                                //{
                                                //    string _s = "";
                                                //    await _Log.ErrorAsync($"派工訊息 {ex.Message} {ex.StackTrace}", true);
                                                //}
                                                //#endregion

                                            }
                                        }
                                        _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[APS_EventLog] (Event,EventType,Id,ServerId,LOGDateTime,NeedId,SimulationId) VALUES ('依排程需求, 自動派工. 工單:{d["DOCNumberNO"].ToString()} 工站:{d["Source_StationNO"].ToString()} 需求量:{d["QTY"].ToString()}','99','{_Str.NewId('Z')}','{_Fun.Config.ServerId}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}','{d["NeedId"].ToString()}','{d["SimulationId"].ToString()}')");

                                        #region 更新電子Tag
                                        if (tmp_dr["Config_macID"].ToString().Trim() != "")
                                        {
                                            int tmp_int = int.Parse(d["Math_EfficientCT"].ToString());
                                            if (tmp_int <= 0) { tmp_int = int.Parse(d["Math_StandardCT"].ToString()); }
                                            string tmp_s = $"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{d["Source_StationNO"].ToString()};2;{d["SimulationId"].ToString()};{d["Source_StationNO_IndexSN"].ToString()}";


                                            //###???要改
                                            dr_LabelStateINFO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{d["macID"].ToString()}' and Type='1'");
                                            if (dr_LabelStateINFO != null)
                                            {
                                                dr_Staion = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr_LabelStateINFO["StationNO"].ToString()}'");
                                                verLab = dr_LabelStateINFO["Version"].ToString().Trim();
                                                var json1 = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{d["DOCNumberNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{d["PartNO"].ToString()}\",\"Text3\":\"品名:\",\"PartName\":\"{d["PartName"].ToString().Replace("\"", "＂").Replace("'", "’")}\",\"Text4\":\"需求量:\",\"DetailQTY\":\"0\",\"Text5\":\"CT:\",\"EfficientCT\":\"{tmp_int}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                                                var json9 = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"{d["DOCNumberNO"].ToString()}\",\"Text2\":\"料號:\",\"PartNO\":\"{d["PartNO"].ToString()}\",\"Text3\":\"品名:\",\"PartName\":\"{d["PartName"].ToString().Replace("\"", "＂").Replace("'", "’")}\",\"Text4\":\"需求量:\",\"DetailQTY\":\"0\",\"Text5\":\"CT:\",\"EfficientCT\":\"{tmp_int}\",\"Text6\":\"達成率:\",\"Rate\":\"0\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";

                                                string isUpdate = "1";
                                                if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                                                {
                                                    string json = "";
                                                    if (verLab != "" && verLab.Substring(0, 2) == "42")
                                                    {
                                                        json1 = $"{json1},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_Staion["StationNO"].ToString()}\",\"text17\":\"{dr_Staion["StationName"].ToString()}\"";
                                                        json9 = $"{json9},\"text18\":\"規格\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{dr_Staion["StationNO"].ToString()}\",\"text17\":\"{dr_Staion["StationName"].ToString()}\"";
                                                        json = $"\"mac\":\"{d["macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{json1}";
                                                    }
                                                    else
                                                    {
                                                        json1 = $"{json1},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                                        json9 = $"{json9},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"\",\"text17\":\"\"";
                                                        json = $"\"mac\":\"{d["macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{json1}";
                                                    }
                                                    json = $"{json},\"QTY\":{d["QTY"].ToString()},\"ledrgb\":\"0\",\"ledstate\":0";
                                                    _Fun.Tag_Write(db,d["macID"].ToString(),"干涉派工", json);
                                                }
                                                else { isUpdate = "0"; }
                                                if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{json9}',Ledrgb='0',Ledstate=0,StationNO='{d["Source_StationNO"].ToString()}',Type='1',OrderNO='{d["DOCNumberNO"].ToString()}',IndexSN='{d["Source_StationNO_IndexSN"].ToString()}',StoreNO='',StoreSpacesNO='',QTY=0,IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{d["macID"].ToString()}'"))
                                                {
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

                    threadLoopTime.Stop();
                    next_time = (int)threadLoopTime.ElapsedMilliseconds;//###???暫時
                    next_time = Math.Max(_Fun.Config.RunTimeServerLoopTime - next_time, 30000);
                    await Task.Delay(next_time, cancellationToken);
                }
            }
            catch (Exception ex)
            {
                await _Log.ErrorAsync($"RUNTimeServer.cs SfcTimerloopautoRUN_Json_Tick Exception停止Thread: {ex.Message} {ex.StackTrace}", true);
            }
            _Fun.Is_RUNTimeServer_Thread_State[2] = false;
            db.Dispose();
        }
        bool IsUpdateTagValue_OK = false;
        private async void SfcTimerloopUpdateTagValue_Tick(CancellationToken cancellationToken = default)
        {
            if (_Fun.Has_Tag_httpClient)
            {
                string err_MEG = "";
                DBADO db = new DBADO("1", _Fun.Config.Db, ref err_MEG);
                if (err_MEG == "")
                {
                    db.Error += new DBADO.ERROR(DBBase_OnException);
                }

                #region 電子標籤資料庫初始化
                await GetAPI_AllMACs(db, cancellationToken);
                #endregion

                try
                {
                    int re_i = 0;
                    DataTable dt = null;
                    DataRow tmp_dr = null;
                    HttpResponseMessage response = null;
                    StringContent content = null;
                    string url = "";
                    var json = "";
                    string verLab = "";
                    while (IsWork && !cancellationToken.IsCancellationRequested)
                    {
                        if (_Fun.Is_Tag_Connect)
                        {
                            #region 送buffer SendTAGDATA 標籤訊號
                            if (_Fun.SendTAGDATA.Count > 0)
                            {
                                json = "";
                                lock (_Fun.Lock_Send_macID)
                                {
                                    DateTime now = DateTime.Now;
                                    foreach (string s in _Fun.SendTAGDATA.Values)
                                    {
                                        if (json == "")
                                        { json = "[{" + s + "}"; }
                                        else { json = json + ",{" + s + "}"; }
                                    }
                                    if (json != "") { json += "]"; }
                                    content = new StringContent(json, Encoding.UTF8, "application/json");
                                    try
                                    {
                                        url = $"http://{_Fun.Config.ElectronicTagsURL}/wms/associate/updateScreen";
                                        response = httpClient.PostAsync(url, content).Result;
                                        if (!response.IsSuccessStatusCode)
                                        {
                                            _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType],[JSON]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','發送Fail','{json}')");
                                            System.Threading.Tasks.Task task = _Log.ErrorAsync($"後台 傳送電子訊號失敗,請通知管理者", false);    //false here, not mailRoot, or endless roop !!
                                        }
                                        else 
                                        {
                                            _ = db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType]) VALUES ('{_Str.NewId('L')}','','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','發送OK')");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer SfcTimerloopUpdateTagValue_Tick 送電子標籤 Exception: {ex.Message}  {ex.StackTrace}", true);    //false here, not mailRoot, or endless roop !!
                                    }
                                    _Fun.SendTAGDATA.Clear();
                                }
                            }
                            #endregion

                            #region 未更新Tag, 重新更新 與 燈檢查 與 警告
                            if (++re_i > 30)
                            {
                                re_i = 0;
                                //###??? 將來 電子燈警告 與 燈檢查(要發API問TAG目前實際的燈狀態) 寫在這裡
                                /*
                                dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and Config_macID!=''");
                                if (dt != null && dt.Rows.Count > 0)
                                {
                                    json = "";
                                    verLab = "";
                                    foreach (DataRow dr in dt.Rows)
                                    {
                                        #region 燈檢查
                                        tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr["Config_macID"].ToString()}'");
                                        if (tmp_dr!=null && dr["Ledrgb"].ToString()!="0" && tmp_dr["Ledrgb"].ToString() != "0")
                                        {

                                        }
                                        #endregion
                                    }
                                }
                                */

                                dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and IsUpdate='0'");
                                if (dt != null && dt.Rows.Count > 0)
                                {
                                    await Task.Delay(1000, cancellationToken);
                                    json = "";
                                    verLab = "";
                                    foreach (DataRow dr in dt.Rows)
                                    {
                                        if (dr["ShowValue"].ToString().Trim() != "")
                                        {
                                            verLab = dr["Version"].ToString().Trim();
                                            if (_Fun.Is_Tag_Connect)
                                            {
                                                json = "";
                                                switch (dr["Type"].ToString().Trim())
                                                {
                                                    case "1":
                                                    case "4":
                                                        if (_Fun.Config.RUNMode == '1')
                                                        {
                                                            if (verLab != "" && verLab.Substring(0, 2) == "42")
                                                            { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{dr["ShowValue"].ToString()}"; }
                                                            else
                                                            { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":45,\"styleid\":54,{dr["ShowValue"].ToString()}"; }
                                                        }
                                                        else
                                                        {
                                                            if (verLab != "" && verLab.Substring(0, 2) == "42")
                                                            { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{dr["ShowValue"].ToString()}"; }
                                                            else
                                                            { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{dr["ShowValue"].ToString()}"; }
                                                        }
                                                        json = $"{json},\"DetailQTY\":\"{dr["QTY"].ToString()}\",\"ledrgb\":\"{dr["Ledrgb"].ToString()}\",\"ledstate\":{dr["Ledstate"].ToString()}";
                                                        break;
                                                    case "2":
                                                        if (verLab != "" && verLab.Substring(0, 2) == "42")
                                                        { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":628,\"styleid\":53,{dr["ShowValue"].ToString()}"; }
                                                        else
                                                        { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":23,\"styleid\":49,{dr["ShowValue"].ToString()}"; }
                                                        json = $"{json},\"ledrgb\":\"{dr["Ledrgb"].ToString()}\",\"ledstate\":{dr["Ledstate"].ToString()}";
                                                        break;
                                                }
                                                if (json != "")
                                                {
                                                    _Fun.Tag_Write(db,dr["macID"].ToString(),"重新更新", json);
                                                    _ = db.DB_SetData($"update SoftNetMainDB.[dbo].[LabelStateINFO] SET IsUpdate='1' where macID='{dr["macID"].ToString()}'");
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            #endregion
                        }
                        if (!IsUpdateTagValue_OK) { IsUpdateTagValue_OK = true; }
                        await Task.Delay(1000, cancellationToken);
                    }
                }
                catch (Exception ex)
                {
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs SfcTimerloopUpdateTagValue_Tick Exception停止Thread (標籤傳送永久失敗): {ex.Message} {ex.StackTrace}", true);
                }
                db.Dispose();
            }
            _Fun.Is_RUNTimeServer_Thread_State[0] = false;
        }

        private async Task GetAPI_AllMACs(DBADO db, CancellationToken cancellationToken = default)
        {
            string uri = $"wms/associate/getTagsMsg";
            //string uri = $"http://{_Fun.Config.ElectronicTagsURL}/wms/associate/getTagsMsg";
            if (!_Fun.Is_Tag_Connect) { return; }
            try
            {
                #region 查詢所有的電子紙是否有註冊在資料庫
                HttpResponseMessage response = httpClient.GetAsync(uri).Result;
                _ = response.EnsureSuccessStatusCode();//如果 Status Code 不為 2xx，會丟出HttpRequestException
                string responseBody = response.Content.ReadAsStringAsync().Result;//取 值
                var data = JsonConvert.DeserializeObject<List<GetPI01>>(responseBody);
                if (data != null && data.Count > 0)
                {
                    DataRow tmp = null;
                    string verLab = "";
                    bool is_del_BarCode_TMP = false;
                    foreach (GetPI01 s in data)
                    {
                        if (s.mac != null && s.mac.Trim() != "")
                        {
                            verLab = s.showStyle.Split('_')[0];
                            DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT macID FROM SoftNetMainDB.[dbo].[LabelMACs] where ServerId='{_Fun.Config.ServerId}' and macID='{s.mac.Trim()}'");
                            if (dr == null)
                            {
                                if (db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[LabelMACs] (ServerId,[macID],[lastOpreateTime],[power],[routerid],[rssi],[showStyle]) VALUES ('{_Fun.Config.ServerId}','{s.mac.Trim()}','{s.lastOpreateTime}',{s.power},'{s.routerid}',{s.rssi},'{s.showStyle}')"))
                                {
                                    if (_Fun.Config.RUNMode == '1')
                                    {
                                        #region 寫專案管理站 
                                        tmp = db.DB_GetFirstDataByDataRow($"select a.*,b.StationName from SoftNetMainDB.[dbo].[Manufacture] as a,SoftNetSYSDB.[dbo].[PP_Station] as b where b.ServerId='{_Fun.Config.ServerId}' and a.Config_macID='{s.mac.Trim()}' and a.StationNO=b.StationNO");
                                        if (tmp != null)
                                        {
                                            string stationNO = tmp["StationNO"].ToString();
                                            string stationName = tmp["StationName"].ToString();
                                            tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{s.mac.Trim()}'");
                                            if (tmp == null)
                                            {
                                                string tmp_s = $"\"Text1\":\"工站編號:\",\"StationNO\":\"{stationNO}\",\"StationName\":\"{stationName}\",\"Text2\":\"製程名稱:\",\"PP_Name\":\"\",\"Text3\":\"\",\"OrderNO\":\"\",\"Text4\":\"\",\"PartNO\":\"\",\"Text5\":\"\",\"OPNO\":\"\",\"Text6\":\"\",\"WorkTime\":\"\",\"Text7\":\"\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelProject/Index/{stationNO}\",\"outtime\":0";
                                                if (db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[LabelStateINFO] (ServerId,[macID],[Version],[Type],[StationNO],[OrderNO],[StoreNO],[StoreSpacesNO],[ShowValue],[IsUpdate],Ledrgb,Ledstate)
                                                            VALUES ('{_Fun.Config.ServerId}','{s.mac.Trim()}','{verLab}','4','{stationNO}','','','','{tmp_s}','0','0',0)"))
                                                {
                                                    _ = db.DB_SetData($"update SoftNetMainDB.[dbo].[Manufacture] set State='2',OrderNO='',OP_NO='',Master_PP_Name='',PP_Name='',IndexSN=0,Station_Custom_IndexSN='',SimulationId='',PartNO='' where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}' and Config_macID='{s.mac.Trim()}'");
                                                }
                                            }
                                            continue;
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        #region 寫工站
                                        tmp = db.DB_GetFirstDataByDataRow($"select a.*,b.StationName from SoftNetMainDB.[dbo].[Manufacture] as a,SoftNetSYSDB.[dbo].[PP_Station] as b where b.ServerId='{_Fun.Config.ServerId}' and a.Config_macID='{s.mac.Trim()}' and a.StationNO=b.StationNO");
                                        if (tmp != null)
                                        {
                                            string stationNO = tmp["StationNO"].ToString();
                                            string stationName = tmp["StationName"].ToString();
                                            tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{s.mac.Trim()}'");
                                            if (tmp == null)
                                            {
                                                string tmp_s = $"\"Text1\":\"工單編號:\",\"OrderNO\":\"\",\"Text2\":\"\",\"PartNO\":\"\",\"Text3\":\"\",\"PartName\":\"\",\"Text4\":\"\",\"QTY\":\"\",\"Text5\":\"\",\"EfficientCT\":\"\",\"Text6\":\"\",\"Rate\":\"\",\"Text7\":\"累計量:\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{stationNO};0;;0\",\"outtime\":0";
                                                if (verLab.Trim() != "" && verLab.Substring(0, 2) == "42")
                                                { tmp_s = $"{tmp_s},\"text18\":\"\",\"text19\":\"\",\"text20\":\"備註:\",\"text15\":\"工站\",\"text16\":\"{stationNO}\",\"text17\":\"{stationName}\""; }
                                                else
                                                { tmp_s = $"{tmp_s},\"text18\":\"\",\"text19\":\"\",\"text20\":\"\",\"text15\":\"\",\"text16\":\"{stationNO}\",\"text17\":\"\""; }
                                                if (db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[LabelStateINFO] (ServerId,[macID],[Version],[Type],[StationNO],[OrderNO],[StoreNO],[StoreSpacesNO],[ShowValue],[IsUpdate],Ledrgb,Ledstate)
                                                            VALUES ('{_Fun.Config.ServerId}','{s.mac.Trim()}','{verLab}','1','{stationNO}','','','','{tmp_s}','0','0',0)"))
                                                {
                                                    _ = db.DB_SetData($"update SoftNetMainDB.[dbo].[Manufacture] set State='2',OrderNO='',OP_NO='',Master_PP_Name='',PP_Name='',IndexSN=0,Station_Custom_IndexSN='',SimulationId='',PartNO='' where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}' and Config_macID='{s.mac.Trim()}'");
                                                }
                                            }
                                            continue;
                                        }
                                        #endregion
                                    }

                                    #region 寫倉庫
                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' and Config_macID='{s.mac.Trim()}'");
                                    if (tmp != null)
                                    {
                                        string storeNO = tmp["StoreNO"].ToString();
                                        string storeName = tmp["StoreName"].ToString();
                                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{s.mac.Trim()}'");
                                        if (tmp == null)
                                        {
                                            string tmp_s = $"\"Text1\":\"倉庫編號:\",\"StoreNO\":\"{storeNO}\",\"Text2\":\"名稱:\",\"StoreName\":\"{storeName}\",\"Text3\":\"狀態:\",\"State\":\"等待工作中...\",\"text7\":\"\",\"text8\":\"\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelStroe/Index/{storeNO}\",\"outtime\":0";
                                            //if (verLab.Trim() != "" && verLab.Substring(0, 2) == "42")
                                            //{ tmp_s = $"\"mac\":\"{s.mac.Trim()}\",\"mappingtype\":628,\"styleid\":53,{tmp_s}"; }
                                            //else
                                            //{ tmp_s = $"\"mac\":\"{s.mac.Trim()}\",\"mappingtype\":23,\"styleid\":49,{tmp_s}"; }
                                            _ = db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[LabelStateINFO] (ServerId,[macID],[Version],[Type],[StationNO],[OrderNO],[StoreNO],[StoreSpacesNO],[ShowValue],[IsUpdate],Ledrgb,Ledstate)
                                                            VALUES ('{_Fun.Config.ServerId}','{s.mac.Trim()}','{verLab}','2','','','{storeNO}','','{tmp_s}','0','0',0)");
                                        }
                                        continue;
                                    }
                                    #endregion

                                    #region 寫儲位
                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[StoreII] where Config_macID='{s.mac.Trim()}'");
                                    if (tmp != null)
                                    {
                                        string storeNO = tmp["StoreNO"].ToString();
                                        string storeSpacesNO = tmp["StoreSpacesNO"].ToString();
                                        tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{s.mac.Trim()}'");
                                        if (tmp == null)
                                        {
                                            _ = db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[LabelStateINFO] (ServerId,[macID],[Version],[Type],[StationNO],[OrderNO],[StoreNO],[StoreSpacesNO],[ShowValue],[IsUpdate])
                                                            VALUES ('{_Fun.Config.ServerId}','{s.mac.Trim()}','{verLab}','3','','','{storeNO}','{storeSpacesNO}','','0')");
                                        }
                                        continue;
                                    }
                                    #endregion
                                }
                            }
                        }
                    }
                }
                #endregion


                #region 是否補更新Tag
                DataTable dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[LabelStateINFO] where IsUpdate='0'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    await Task.Delay(1000, cancellationToken);
                    string json = "";
                    string verLab = "";
                    foreach (DataRow dr in dt.Rows)
                    {

                        verLab = dr["Version"].ToString().Trim();
                        if (_Fun.Is_Tag_Connect)
                        {
                            json = "";
                            switch (dr["Type"].ToString().Trim())
                            {
                                case "1":
                                case "4":
                                    if (dr["ShowValue"].ToString().Trim() != "")
                                    {
                                        if (_Fun.Config.RUNMode == '1')
                                        {
                                            if (verLab != "" && verLab.Substring(0, 2) == "42")
                                            { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{dr["ShowValue"].ToString()}"; }
                                            else
                                            { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":45,\"styleid\":54,{dr["ShowValue"].ToString()}"; }
                                        }
                                        else
                                        {
                                            if (verLab != "" && verLab.Substring(0, 2) == "42")
                                            { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":744,\"styleid\":52,{dr["ShowValue"].ToString()}"; }
                                            else
                                            { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":71,\"styleid\":48,{dr["ShowValue"].ToString()}"; }
                                        }
                                        json = $"{json},\"QTY\":\"{dr["QTY"].ToString()}\",\"ledrgb\":\"{dr["Ledrgb"].ToString()}\",\"ledstate\":{dr["Ledstate"].ToString()}";
                                    }
                                    break;
                                case "2":
                                    if (dr["ShowValue"].ToString().Trim() != "")
                                    {
                                        if (verLab != "" && verLab.Substring(0, 2) == "42")
                                        { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":628,\"styleid\":53,{dr["ShowValue"].ToString()}"; }
                                        else
                                        { json = $"\"mac\":\"{dr["macID"].ToString()}\",\"mappingtype\":23,\"styleid\":49,{dr["ShowValue"].ToString()}"; }
                                        json = $"{json},\"ledrgb\":\"{dr["Ledrgb"].ToString()}\",\"ledstate\":{dr["Ledstate"].ToString()}";
                                    }
                                    break;
                                case "3":
                                    _ = db.DB_SetData($"update SoftNetMainDB.[dbo].[LabelStateINFO] SET IsUpdate='1' where macID='{dr["macID"].ToString()}'");

                                    json = "[{" + $"\"mac\":\"{dr["macID"].ToString()}\",\"outtime\":0,\"ledrgb\":\"{dr["Ledrgb"].ToString()}\",\"ledmode\":0,\"buzzer\":0,\"lednum\":255,\"reserve\":\"reserve\"" + "}]";
                                    var content = new StringContent(json, Encoding.UTF8, "application/json");
                                    try
                                    {
                                        string url = $"http://{_Fun.Config.ElectronicTagsURL}/wms/associate/lightTagsLed";
                                        response = httpClient.PostAsync(url, content).Result;
                                        if (!response.IsSuccessStatusCode)
                                        {
                                            System.Threading.Tasks.Task task = _Log.ErrorAsync($"後台 儲位亮燈 傳送電子訊號失敗", false);    //false here, not mailRoot, or endless roop !!
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Threading.Tasks.Task task = _Log.ErrorAsync($"後台 儲位亮燈 傳送電子訊號失敗 Exception: {ex.Message}", false);    //false here, not mailRoot, or endless roop !!
                                    }
                                    continue;
                            }
                            if (json != "")
                            {
                                _Fun.Tag_Write(db,dr["macID"].ToString(),"初始化", json);
                                _ = db.DB_SetData($"update SoftNetMainDB.[dbo].[LabelStateINFO] SET IsUpdate='1' where macID='{dr["macID"].ToString()}'");
                            }
                        }
                    }
                }
                #endregion

                //###??? 補打Manufacture State=2的標籤燈

            }
            catch (Exception ex)
            {
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"RUNTimeServer.cs GetAPI_AllMACs HttpClient 送電子標籤 failed: {ex.Message} {ex.StackTrace}", true);

            }


        }


        private Dictionary<char, List<KeyAndValue>> _efficientConfig = new Dictionary<char, List<KeyAndValue>>();

        private string LabelProject_Start_Stop(DBADO db, string status, string stationNO, DataRow dr_M, string opNO, ref LabelProject keys)
        {
            //###???若此處改 RUNTimeServer.csㄝ要改
            keys.TOTALOKQTY = 0;
            keys.TOTALFailQTY = 0;
            string meg = "";
            string ledrgb = "0";
            string Ledstate = "0";
            int dis_DetailQTY = 0;
            TimeSpan workTime = TimeSpan.Zero;
            int totalQTY = 0;
            var json1 = "";
            DataRow dr_PP_Station = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}'");
            string stationName = dr_PP_Station["StationName"].ToString();
            string tmp_s = "";
            DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}' and PP_Name='{dr_M["PP_Name"].ToString()}' and OP_NO='{opNO}' and OrderNO='{dr_M["OrderNO"].ToString()}' and PartNO='{dr_M["PartNO"].ToString()}' and IndexSN={dr_M["IndexSN"].ToString()} and EndTime is null");
            if (dr != null)
            {
                if (keys.PartNO != null && keys.PartNO.Trim() != "")
                {
                    if (keys.PartName == null || keys.PartName.Trim() == "")
                    {
                        DataRow dr_Material = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{keys.PartNO}'");
                        if (dr_Material != null)
                        {
                            keys.PartName = dr_Material["PartName"].ToString();
                            keys.Specification = dr_Material["Specification"].ToString();
                        }
                    }
                }
                #region 計算總數量,總工時,確認StartTime not NULL
                DataTable dt_SFC_StationProjectDetail = db.DB_GetData($"SELECT * FROM SoftNetLogDB.[dbo].[SFC_StationProjectDetail] where ServerId='{_Fun.Config.ServerId}' and Id='{dr["Id"].ToString()}' and StationNO='{stationNO}'");
                DateTime st = DateTime.Now;
                foreach (DataRow d in dt_SFC_StationProjectDetail.Rows)
                {
                    keys.TOTALOKQTY += int.Parse(d["ProductFinishedQty"].ToString());
                    keys.TOTALFailQTY += int.Parse(d["ProductFailedQty"].ToString());
                    totalQTY += (keys.TOTALOKQTY + keys.TOTALFailQTY);
                    if (d.IsNull("StartTime"))
                    {
                        _ = db.DB_SetData($"update SoftNetLogDB.[dbo].[SFC_StationProjectDetail] set StartTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}',RemarkTimeS='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and Id='{d["Id"].ToString()}' and StationNO='{stationNO}' and OP_NO='{d["OP_NO"].ToString()}'");
                    }
                    else
                    {
                        workTime = workTime.Add(GetCT(db, dr_PP_Station["CalendarName"].ToString(), Convert.ToDateTime(d["StartTime"].ToString()), DateTime.Now));
                    }
                }
                keys.TOTALWorkTime = $"{((int)workTime.TotalHours).ToString()}小時 {workTime.Minutes.ToString()}分";
                #endregion
            }
            if (status == "1")
            {
                #region 開始
                Ledstate = "2500";
                ledrgb = "ff00";
                if (dr == null)
                {
                    if (!db.DB_SetData(@$"INSERT INTO SoftNetLogDB.[dbo].[SFC_StationProjectDetail] (ServerId,Id,StationNO,Master_PP_Name,PP_Name,OP_NO,IndexSN,OrderNO,PartNO,RMSName,StartTime,RemarkTimeS) VALUES 
                                        ('{_Fun.Config.ServerId}','{_Str.NewId('A')}','{stationNO}','{dr_M["Master_PP_Name"].ToString()}','{dr_M["PP_Name"].ToString()}','{opNO}',{dr_M["IndexSN"].ToString()},'{dr_M["OrderNO"].ToString()}','{dr_M["PartNO"].ToString()}','{dr_PP_Station["RMSName"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}')"))
                    { meg = $"{meg}<br>{stationNO} 無法新增開始資料, 請通知管理者."; }
                }
                else
                {
                    _ = db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SFC_StationProjectDetail] SET RemarkTimeE=NULL where ServerId='{_Fun.Config.ServerId}' and Id='{dr["Id"].ToString()}' and StationNO='{stationNO}'");
                }
                #endregion
            }
            else if (status == "2")
            {
                #region 停止
                if (dr != null)
                {
                    _ = db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SFC_StationProjectDetail] SET RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and Id='{dr["Id"].ToString()}' and StationNO='{stationNO}'");
                }
                #endregion
            }
            else if (status == "4")
            {
                #region 關閉, 寫EndTime
                if (dr != null)
                {
                    Ledstate = "0";
                    _ = db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SFC_StationProjectDetail] SET EndTime='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}',RemarkTimeE='{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}' where ServerId='{_Fun.Config.ServerId}' and Id='{dr["Id"].ToString()}' and StationNO='{stationNO}'");
                    _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET OrderNO='',OP_NO='',Master_PP_Name='',PP_Name='',IndexSN=0,Station_Custom_IndexSN='',StationNO_Custom_DisplayName='',State='4' where ServerId='{_Fun.Config.ServerId}' and ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}'");
                }
                else
                { meg = $"{meg}<br>{stationNO} 查無起動過的紀錄資料, 無法關站."; }
                #endregion
            }
            if (meg == "")
            {
                _ = db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] SET State='{status}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}'");
                // db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[Manufacture_Log] ([Id],[logDate],[StationNO],[State],[OrderNO],[PartNO]) VALUES ('{_Str.NewId('C')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{stationNO}','{status}','{dr_M["OrderNO"].ToString()}','{dr_M["PartNO"].ToString()}')");
                string type = "";
                switch (status)
                {
                    case "1":
                        type = "干涉開工"; break;
                    case "2":
                        type = "干涉停工"; break;
                    case "4":
                        type = "干涉關站"; break;
                }
                _ = db.DB_SetData(@$"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,IndexSN) VALUES 
                                ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','RUNTimeServer','{type}','{dr_M["PP_Name"].ToString()} {opNO}','{stationNO}','{dr_M["PartNO"].ToString()}','{dr_M["OrderNO"].ToString()}',{dr_M["IndexSN"].ToString()})");

            }

            #region 更新Tag
            if (_Fun.Has_Tag_httpClient)
            {
                DataRow dr_LabelStateINFO = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'");
                if (dr_LabelStateINFO != null && dr_M["Config_macID"].ToString() != "")
                {
                    string sql = "";
                    tmp_s = $"http://{_Fun.Config.LocalWebURL}/LabelProject/Index/{stationNO}";
                    string isUpdate = "1";
                    if (!_Fun.Is_Tag_Connect) { isUpdate = "0"; }
                    if (status != "4")
                    {
                        string dis_opNO = dr_M["OP_NO"].ToString();
                        if (dr_M["OP_NO"].ToString().Split(';').Length > 1) { dis_opNO = "多人共同作業"; }
                        string json = $"\"Text1\":\"工站編號:\",\"StationNO\":\"{stationNO}\",\"StationName\":\"{stationName}\",\"Text2\":\"製程名稱:\",\"PP_Name\":\"{dr_M["PP_Name"].ToString()}\",\"Text3\":\"工單編號:\",\"OrderNO\":\"{dr_M["OrderNO"].ToString()}\",\"Text4\":\"料號:\",\"PartNO\":\"{dr_M["PartNO"].ToString()}\",\"Text5\":\"作業人員:\",\"OPNO\":\"{dis_opNO}\",\"Text6\":\"累計工時:\",\"WorkTime\":\"{((int)workTime.TotalHours).ToString()}.{workTime.Minutes.ToString()}\",\"Text7\":\"累計量:\",\"BarCode\":\"{tmp_s}\",\"outtime\":0";
                        json1 = $"{json},\"QTY\":\"{totalQTY.ToString()}\",\"ledrgb\":\"{ledrgb}\",\"ledstate\":{Ledstate.ToString()}";
                        sql = $"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{json}',Ledrgb='{ledrgb}',Ledstate={Ledstate},StationNO='{stationNO}',Type='4',OrderNO='{dr_M["OrderNO"].ToString()}',IndexSN='{dr_M["IndexSN"].ToString()}',StoreNO='',StoreSpacesNO='',QTY=0,IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'";
                    }
                    else
                    {
                        string json = $"\"Text1\":\"工站編號:\",\"StationNO\":\"{stationNO}\",\"StationName\":\"{stationName}\",\"Text2\":\"製程名稱:\",\"PP_Name\":\"\",\"Text3\":\"\",\"OrderNO\":\"\",\"Text4\":\"\",\"PartNO\":\"\",\"Text5\":\"\",\"OPNO\":\"\",\"Text6\":\"\",\"WorkTime\":\"\",\"Text7\":\"\",\"BarCode\":\"http://{_Fun.Config.LocalWebURL}/LabelProject/Index/{stationNO}\",\"outtime\":0";
                        json1 = $"{json},\"QTY\":\"\",\"ledrgb\":\"{ledrgb}\",\"ledstate\":{Ledstate}";
                        sql = $"UPDATE SoftNetMainDB.[dbo].[LabelStateINFO] set ShowValue='{json}',Ledrgb='{ledrgb}',Ledstate={Ledstate},StationNO='{stationNO}',Type='4',OrderNO='',IndexSN='',StoreNO='',StoreSpacesNO='',QTY=0,IsUpdate='{isUpdate}' where ServerId='{_Fun.Config.ServerId}' and macID='{dr_M["Config_macID"].ToString()}'";
                    }
                    if (_Fun.Has_Tag_httpClient && _Fun.Is_Tag_Connect)
                    {
                        json1 = $"\"mac\":\"{dr_M["Config_macID"].ToString()}\",\"mappingtype\":45,\"styleid\":54,{json1}";
                        _Fun.Tag_Write(db,dr_M["Config_macID"].ToString(),"", json1);
                    }
                    else { meg = $"{meg}<br>{stationNO} 工站傳送電子訊號失敗, 請通知管理者"; }
                    if (db.DB_SetData(sql))
                    {

                    }
                }
            }
            #endregion

            return meg;
        }

        private TimeSpan GetCT(DBADO db, string CalendarName, DateTime Comintime, DateTime Comouttime)
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

    }
}
