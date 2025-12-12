using Base;
using Base.Models;
using Base.Services;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;

using SoftNetWebII.Models;
using SoftNetWebII.Tables;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Net.Sockets;
using System.Net.WebSockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static SoftNetWebII.Services.RUNTimeServer;
using static StackExchange.Redis.Role;

namespace SoftNetWebII.Services
{
    public class SNWebSocketService
    {
        #region 變數宣告
        public bool IsWork = false;

        public static ManualResetEvent allDone = new ManualResetEvent(false);
        private Socket listener;
        public object lock__WebSocketList = new object();
        public Dictionary<string, rmsConectUserData> _WebSocketList = new Dictionary<string, rmsConectUserData>(); //key=ip:port  value=rmsConectUserData(deviceName,ipcRobotName,socket)
        private uint logID = 1;
        private object lock_WebSocket = new object();

        public List<rmsMasterUserData> _MasterRMSUserList = new List<rmsMasterUserData>(); //與Service(s) Socket連線列表
        private object lock__MasterRMSUserList = new object();
        private object lock_Sevice = new object();
        public char RMSLogMode = '1'; //LogMode 0:Normal Mode 1:Debug Mode
        #endregion

        public SNWebSocketService()
        {
            StartListening();
            SpinWait.SpinUntil(() => !IsWork, 2000);
            if (IsWork) { _Fun.Is_SNWebSocketService_OK = true; }
        }
        private void StartListening()
        {
            try
            {
                IsWork = true;

                #region 建立 WebSocket Server Thread
                listener = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.IP);
                listener.Bind(new IPEndPoint(IPAddress.Any, _Fun.Config.WesocketPort));//###???要改指地IP
                listener.Listen(1024);//###???將來要改參數化
                Thread controlDATA_Othread = new Thread(() =>
                {
                    startListening_Thread();
                });
                controlDATA_Othread.IsBackground = true;
                controlDATA_Othread.Start();
                #endregion

                #region 建立 _MasterRMSUserList list 與 所有Service 之間的連線 port 5431
                //###???將來要譨定義主Service與功能性Service, 分Service群主要是分散主Service運算執行負載
                DataTable dt;
                using (DBADO db = new("1", _Fun.Config.Db))
                {
                    dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[Device] where ServerId='{_Fun.Config.ServerId}' and DeviceEnable=1 and Device_Type='RMS'");
                    if (dt != null)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            rmsMasterUserData mrul = _MasterRMSUserList.Find(delegate (rmsMasterUserData t) { return t.deviceName == dr["DeviceName"].ToString(); });
                            if (mrul == null)
                            {
                                lock (lock__MasterRMSUserList)
                                {
                                    Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                                    IPEndPoint ipPoint = new IPEndPoint(IPAddress.Parse(dr["Device_IP"].ToString()), 5431);
                                    _MasterRMSUserList.Add(new rmsMasterUserData(dr["DeviceName"].ToString(), socket, ipPoint));
                                }
                            }
                        }
                    }
                }
                #endregion

                #region 建立 _MasterRMSUserList Socket Client 連線與斷線自動連線
                Thread MasterSocketInit_thread = new Thread(() =>
                {
                    CheckMasterSocketConnectIsOKthread_Tick();
                });
                MasterSocketInit_thread.IsBackground = true;
                MasterSocketInit_thread.Start();
                #endregion

                
            }
            catch (Exception ex)
            {
                IsWork = false;
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"SNWebSocketService.cs Start Exception停止Thread (SNWebSocketService永久失敗): {ex.Message} {ex.StackTrace}", true);
            }
        }

        private void CheckMasterSocketConnectIsOKthread_Tick()//輪巡 _MasterRMSUserList(5431)網路狀況 
        {
            DBADO db = new DBADO("1", _Fun.Config.Db);
            DataRow dr = null;
            do
            {
                for (int i = 0; i < _MasterRMSUserList.Count; i++)
                {
                    try
                    {
                        if (_MasterRMSUserList[i].socket == null)
                        {
                            #region reead DB 確認
                            dr = db.DB_GetFirstDataByDataRow($"select Device_IP from SoftNetSYSDB.[dbo].Device where ServerId='{_Fun.Config.ServerId}' and DeviceName='" + _MasterRMSUserList[i].deviceName + "'");
                            if (dr == null)
                            {
                                lock (lock__MasterRMSUserList)
                                {
                                    _MasterRMSUserList.RemoveAt(i);
                                    continue;
                                }
                            }
                            else
                            {
                                lock (lock__MasterRMSUserList)
                                {
                                    _MasterRMSUserList[i].socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                                    _MasterRMSUserList[i].ipPoint = new IPEndPoint(IPAddress.Parse(dr["Device_IP"].ToString()), 5431);
                                    _MasterRMSUserList[i].ReSetFlag = false;
                                }
                            }
                            #endregion
                        }
                        if (RmsSendPing(_MasterRMSUserList[i].socket))
                        {
                            #region 連線正常
                            if (!_MasterRMSUserList[i].isWork)
                            {
                                lock (lock__MasterRMSUserList) { _MasterRMSUserList[i].isWork = true; }
                            }
                            #endregion
                        }
                        else
                        {
                            #region 重新連線
                            if (RmsLoopConnectII(_MasterRMSUserList[i].socket, _MasterRMSUserList[i].ipPoint))
                            {
                                #region 重新連線,並送初始訊號
                                _MasterRMSUserList[i].ReSetFlag = false;
                                _MasterRMSUserList[i].socket.IOControl(IOControlCode.KeepAliveValues, ToolFunStatic.KeepAlive(1, 5000, 1000), null);
                                _MasterRMSUserList[i].socket.SetSocketOption(SocketOptionLevel.Tcp, SocketOptionName.NoDelay, 1);
                                if (!_MasterRMSUserList[i].isWork)
                                {
                                    lock (lock__MasterRMSUserList) { _MasterRMSUserList[i].isWork = true; }
                                }
                                Thread clientThread = new Thread(new ParameterizedThreadStart(MasterProcessRequest))
                                { IsBackground = true };
                                clientThread.Start(_MasterRMSUserList[i].socket);
                                #endregion
                                //###??? 將來要確認
                                RmsSend(_MasterRMSUserList[i].socket, 1, "IIS_Login,SoftNet_I,");
                            }
                            else
                            {
                                #region 重新new Socket
                                if (!_MasterRMSUserList[i].ReSetFlag)
                                {
                                    lock (lock__MasterRMSUserList)
                                    {
                                        if (_MasterRMSUserList[i].isWork) { _MasterRMSUserList[i].isWork = false; }
                                        if (_MasterRMSUserList[i].socket != null) { _MasterRMSUserList[i].socket.Close(); _MasterRMSUserList[i].socket = null; }
                                        _MasterRMSUserList[i].socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                                        _MasterRMSUserList[i].ReSetFlag = true;
                                    }
                                }
                                #endregion
                            }
                            #endregion
                        }
                    }
                    catch (Exception ee)
                    {
                        if (lock__MasterRMSUserList != null)
                        {
                            lock (lock__MasterRMSUserList)
                            {
                                if (_MasterRMSUserList[i].socket != null) { _MasterRMSUserList[i].socket.Close(); _MasterRMSUserList[i].socket = null; }
                                _MasterRMSUserList[i].socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                            }
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                if (_Fun.Is_Thread_ForceClose) { IsWork = false; break; }
                SpinWait.SpinUntil(() => !IsWork, 5000);
            }
            while (IsWork);
            db.Dispose();
        }
        private bool RmsSendPing(Socket sender)
        {
            if (sender == null) { return false; }
            try
            {
                byte[] byteSend = new byte[6] { 1, 0, 0, 0, 0, 1 };
                int i = sender.Send(byteSend, 0, byteSend.Length, SocketFlags.None);
                if (i <= 0)
                {
                    return false;
                }
                return true;
            }
            catch { }
            return false;
        }
        private bool RmsLoopConnectII(Socket rmsNet, IPEndPoint host)
        {
            try
            {
                IAsyncResult result = rmsNet.BeginConnect(host, null, null);
                bool flag = result.AsyncWaitHandle.WaitOne(2000, true);
                if (result.IsCompleted)
                {
                    if (rmsNet != null)
                    { rmsNet.EndConnect(result); }
                    if (IsSocketConnected(rmsNet)) { return true; }
                }
            }
            catch { }
            if (rmsNet != null) { rmsNet.Close(); rmsNet = null; }
            return false;
        }
        private bool IsSocketConnected(Socket sock)
        {
            if (sock == null) { return false; }
            if (!RmsSendPing(sock)) { return false; }
            return true;
        }
        public bool RmsSend(string deviceNmae, byte type, object objdata, uint logid = 1, bool writelog = true)
        {
            rmsMasterUserData mrul = _MasterRMSUserList.Find(delegate (rmsMasterUserData t) { return t.deviceName == deviceNmae; });
            if (mrul != null && mrul.socket != null)
            { return RmsSend(mrul.socket, type, objdata, logid, writelog).Result; }
            return false;
        }
        public async Task<bool> RmsSend(Socket user, byte type, object objdata, uint logid = 1, bool writelog = true)
        {
            EndPoint oldEP = null;
            string ipport = "";
            try
            {
                if (user == null || user.RemoteEndPoint == null) { return false; }

                oldEP = user.RemoteEndPoint;
                ipport = oldEP.ToString();

                byte[] data = null;
                switch (type)
                {
                    case 0://ping
                        data = new byte[] { 1 };
                        break;
                    case 1://string
                        //if (writelog)
                        //{
                        //    string disIP = IpportConvertIPandName(ipport);
                        //    if (RMSLogMode == '1')
                        //    { SoftNetService.Program._NLogMain.Write_Record(logid, disIP, LogTitle.Send, LogSourceName.ProtocolCMD, "", (string)objdata); }
                        //}
                        data = Encoding.UTF8.GetBytes((string)objdata);
                        await _Log.SocketLogAsync("Socket5431Log", $"S Type=1 Data={(string)objdata}");
                        break;
                    case 254://RMSProtocol
                        RMSProtocol obj = (RMSProtocol)objdata;
                        //if (writelog)
                        //{
                        //    string disIP = IpportConvertIPandName(ipport);
                        //    if (RMSLogMode == '1')
                        //    { SoftNetService.Program._NLogMain.Write_Record(logid, disIP, LogTitle.Send, LogSourceName.ProtocolType, obj.DataType.ToString(), obj.Data); }
                        //}
                        data = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(obj));
                        await _Log.SocketLogAsync("Socket5431Log", $"S Type=1 Data={obj.Data}");

                        break;
                }
                if (data != null)
                {
                    byte[] byteSend = new byte[data.Length + 5];
                    byte[] tmp2 = BitConverter.GetBytes(data.Length);
                    tmp2.CopyTo(byteSend, 0);
                    byteSend[4] = type;
                    data.CopyTo(byteSend, 5);
                    //###???暫改Send  user.BeginSend(byteSend, 0, byteSend.Length, SocketFlags.None, new AsyncCallback(rms_sendCallback), user);
                    if (user.Send(byteSend, 0, byteSend.Length, SocketFlags.None) > 0) { return true; }
                }
            }
            catch (Exception ex)
            {
                //++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "rmsSend", RMSError.Service_Exception, LogSourceName.IPPort, ipport, 61,
                //    ToolFun.StringAdd("訊號發送失敗,請檢察網路,或重啟Service Engine. data=", errMEG, " Exception:", ex.Message), ex);
                ////ExceptionError A6
                //if (user != null) { user.Close(); }
                await _Log.SocketLogAsync("Socket5431Log", $"Err RmsSend MEG={ex.Message}");

            }
            return false;
        }
        private void MasterProcessRequest(object socket)
        {
            if (socket == null) { return; }
            object lock_receiveBuffer = new object();
            Socket _sockinfo = (Socket)socket;
            byte[] recvBytes = new byte[4096];
            List<byte> receiveBuffer = new List<byte>();
            EndPoint oldEP;
            string ipport = "";
            int blen = 0;
            byte[] tra;
            byte type = 0;

            try
            {
                oldEP = _sockinfo.RemoteEndPoint;
                ipport = oldEP.ToString();
                while (IsWork)
                {
                    if (_Fun.Is_Thread_ForceClose) { IsWork = false; break; }
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
                                                if (false)//###???Program.SocketReceiveUserThreadPool
                                                {
                                                    SNThreadScheduler.Instance.StartAction(() => { Master_ResolveData2String(ipport, cmd); });
                                                }
                                                else
                                                {
                                                    Task ttr21 = new Task(() =>
                                                    {
                                                        Master_ResolveData2String(ipport, cmd);
                                                    });
                                                    ttr21.Start();
                                                }
                                                break;
                                            case 254://RMSProtocol
                                                RMSProtocol pp = JsonConvert.DeserializeObject<RMSProtocol>(Encoding.UTF8.GetString(a.ToArray(), 0, a.Count));
                                                if (false)//###???Program.SocketReceiveUserThreadPool
                                                { SNThreadScheduler.Instance.StartAction(() => { Master_ResolveData2RMSProtocol(_sockinfo, ipport, pp); }); }
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
                                                //++RMSDBErrorCount;
                                                //SoftNetService.Program._NLogMain.Write_RunError(1, ipport, "Master接收封包程序", RMSError.Service_Protocol_NoDefine, LogSourceName.DeviceName, Program.RMSName, 181, ToolFun.StringAdd("系統異常: Service收到未定義的Type,請聯絡TMM. type=", type.ToString(), "來源IP=", ipport));
                                                _sockinfo.Close();
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
                                    } while (receiveBuffer.Count >= 5 && IsWork);
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
                    //++RMSDBErrorCount;
                    //SoftNetService.Program._NLogMain.Write_ExceptionError(1, "MasterProcessRequest", RMSError.Service_Exception, LogSourceName.IPPort, ipport, 0,
                    //    ToolFun.StringAdd("已關閉網路連線,請檢察網路,或重啟Service Engine. errorcode:", sex.ErrorCode.ToString(), " Exception:", sex.Message), sex);
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"MasterProcessRequest1 {sex.Message} {sex.StackTrace}", true);
                }
            }
            catch (Exception ex)
            {
                if (lock_receiveBuffer == null) { }
                else
                {
                    //++RMSDBErrorCount;
                    //SoftNetService.Program._NLogMain.Write_ExceptionError(1, "MasterProcessRequest", RMSError.Service_Exception, LogSourceName.IPPort,
                    //    ipport, 0, ToolFun.StringAdd("已關閉網路連線,請檢察網路,或重啟Service Engine. Exception:", ex.Message), ex);
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"MasterProcessRequest2 {ex.Message} {ex.StackTrace}", true);
                }
            }
            lock_receiveBuffer = null;
            if (_sockinfo != null)
            {
                _sockinfo.Close();
                _sockinfo = null;
            }

        }
        private void Master_ResolveData2String(string ipport, string[] cmd)
        {
            uint logid = 1;
            for (int i = 1; i < cmd.Length; i++)
            {
                cmd[i] = cmd[i].Replace("\x03", ",");
            }
            string ip = ipport.Split(':')[0].Trim();
            string disIP = "";
            //string disIP = IpportConvertIPandName(ipport);

            try
            {
                switch (cmd[0])//通知service要做什麼事
                {
                    case "LIB_TO_WEB":
                        if (cmd.Length >= 3)
                        {
                            switch (cmd[2])
                            {
                                case "StationStatusChange":
                                    string _s = "";
                                    if (_WebSocketList.ContainsKey(cmd[1]))
                                    {
                                        Send(_WebSocketList[cmd[1]].socket, cmd[2]);
                                    }
                                    break;
                            }
                        }
                        break;
                    case "SendALLClient":
                        if (cmd.Length >2)
                        {
                            string tmp = string.Join(',', cmd).Replace("SendALLClient,","");
                            foreach (KeyValuePair<string, rmsConectUserData> r in _WebSocketList)
                            {
                                if (r.Key != null && r.Value.socket != null)
                                {
                                    Send(r.Value.socket, tmp);
                                }
                            }
                        }
                        break;
                    default:
                        //++RMSDBErrorCount;
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
        private async void Master_ResolveData2RMSProtocol(Socket sender, string ipport, RMSProtocol obj)
        {
            //###???要檢查來源的條件,如91501的if (_ToolUserList.ContainsKey(ipport)), 5430ㄝ要

            //Console.WriteLine("connect " + sender.RemoteEndPoint);
            uint logid = 1;
            bool writeLog = false;
            if (obj.DataType != 11501 && obj.DataType != 11502 && obj.DataType != 11503 && obj.DataType < 91300)
            {
                writeLog = true;
                lock (lock_Sevice)
                {
                    if (logid > 4290000000)
                    {
                        //SoftNetService.Program._NLogMain.NewOtherLogFile();
                        logID = 2;
                    }
                    logid = ++logID;
                }
            }
            if (RMSLogMode == '0')
            {
                writeLog = false;
            }
            string[] cmd = null;
            string errMEG = "";
            try
            {
                DataRow dr = null;
                await _Log.SocketLogAsync("Socket5431Log", $"R Type={obj.DataType.ToString()} Data={obj.Data}");

                switch (obj.DataType)
                {
                    case 3511://SNService引發的工站狀態改變 cmd=staionNO,orderNO,state,op_no,master_PP_namete,pp_Name,indexSN,custom_Index
                        cmd = JsonConvert.DeserializeObject<string[]>(obj.Data);
                        DBADO db = new DBADO("1", _Fun.Config.Db);
                        switch (cmd[2])
                        {
                            case "0"://CMD_Close
                            case "1"://CMD_Start
                            case "2"://CMD_Stop
                            case "3"://CMD_Pause
                            case "4"://關工站 //###???未完
                                string simulationId = "";
                                string partNO = "";
                                #region 查有無SID
                                DataRow sfcdr = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].PP_WorkOrder where ServerId='{_Fun.Config.ServerId}' and OrderNO='{cmd[1]}'");
                                partNO = sfcdr["PartNO"].ToString();
                                if (!sfcdr.IsNull("needId") && sfcdr["needId"].ToString().Trim() != "")
                                {
                                    DataRow sId_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] WHERE NeedId='{sfcdr["needId"].ToString()}' and DOCNumberNO = '{cmd[1]}' AND Source_StationNO = '{cmd[0]}' AND (Source_StationNO_IndexSN={cmd[6]} or Source_StationNO_Custom_IndexSN='{cmd[7]}')");
                                    if (sId_dr != null)
                                    {
                                        simulationId = sId_dr["SimulationId"].ToString();
                                        partNO = sId_dr["PartNO"].ToString();
                                    }
                                }
                                #endregion
                                if (db.DB_SetData($"UPDATE SoftNetMainDB.[dbo].[Manufacture] set State='{cmd[2]}',OrderNO='{cmd[1]}',OP_NO='{cmd[3]}',Master_PP_Name='{cmd[4]}',PP_Name='{cmd[5]}',IndexSN={cmd[6]},Station_Custom_IndexSN='{cmd[7]}',SimulationId='{simulationId}' where ServerId='{_Fun.Config.ServerId}' and StationNO='{cmd[0]}'"))
                                {
                                    //db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[Manufacture_Log] ([Id],[logDate],[StationNO],[State],[OrderNO],[PartNO]) VALUES ('{_Str.NewId('C')}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','{cmd[0]}','{cmd[2]}','{cmd[1]}','{partNO}')");
                                    string type = "";
                                    switch (cmd[2])
                                    {
                                        case "1":
                                            type = "智慧開工"; break;
                                        case "2":
                                            type = "智慧停工"; break;
                                        case "4":
                                        case "0":
                                            type = "智慧關站"; break;
                                    }
                                    DataRow dr_SimulationId = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{simulationId}'");
                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{dr_SimulationId["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','RUNTimeServer','{type}','{cmd[5]} {cmd[3]}','{cmd[0]}','','{cmd[1]}',{cmd[6]})");//###??? dr_M["PartNO"].ToString() 暫時放空

                                    //通知webSocket send
                                    //var webSocketService = (SNWebSocketService)_Fun.DiBox.GetService(typeof(SNWebSocketService));
                                    foreach (KeyValuePair<string, rmsConectUserData> r in _WebSocketList)
                                    {
                                        if (r.Key != null && r.Value.socket != null)
                                        {
                                            Send(r.Value.socket, "StationStatusChange");
                                        }
                                    }

                                }
                                break;
                            case "5"://通知
                                {
                                    //通知webSocket send
                                    //var webSocketService = (SNWebSocketService)_Fun.DiBox.GetService(typeof(SNWebSocketService));
                                    foreach (KeyValuePair<string, rmsConectUserData> r in _WebSocketList)
                                    {
                                        if (r.Key != null && r.Value.socket != null)
                                        {
                                            Send(r.Value.socket, "StationStatusChange");
                                        }
                                    }
                                }
                                break;
                        }
                        db.Dispose();
                        break;
                    case 11100://共用成功與否訊號  [1]=原cmd [2]=訊號
                        break;
                    default:
                        //if (writeLog) { SoftNetService.Program._NLogMain.Write_Record(logid, ipport, LogTitle.M_Receive, LogSourceName.ProtocolType, obj.DataType.ToString(), obj.Data); }
                        //++RMSDBErrorCount;
                        //SoftNetService.Program._NLogMain.Write_RunError(logid, ipport, "Master Parsing Network Packets", RMSError.Service_Protocol_NoDefine, LogSourceName.DeviceName, Program.RMSName, 181, ToolFun.StringAdd("系統異常: 接收到訊號,但程式無定義,請聯絡TMM. type=", obj.DataType.ToString()));
                        //RunError A18
                        sender.Close();
                        return;
                }
            }
            catch (Exception ex)
            {
                await _Log.SocketLogAsync("Socket5431Log", $"Err Master_ResolveData2RMSProtocol MEG={ex.Message}");
                //if (lock__RMSUserList == null) { return; }
                //++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "Master_ResolveData2RMSProtocol", RMSError.Service_Exception,
                //    LogSourceName.IPPort, ipport, 61, ToolFun.StringAdd("通訊資料異常,請重啟Service Engine或聯絡TMM. data:", obj.Data, " Exception:", ex.Message), ex);
                //ExceptionError A6
            }
        }






        private void startListening_Thread()
        {
            while (IsWork)
            {
                allDone.Reset();
                listener.BeginAccept(new AsyncCallback(AcceptCallback), listener);
                allDone.WaitOne();
            }
        }
        private void AcceptCallback(IAsyncResult ar)
        {
            // Signal the main thread to continue.
            allDone.Set();


            // Get the socket that handles the client request.
            Socket client = (Socket)ar.AsyncState;
            Socket handler = client.EndAccept(ar);
            string ipport = "";// = handler.RemoteEndPoint.ToString();
            try
            {
                ipport = handler.RemoteEndPoint.ToString();
                string[] deviceName = GetUserList_DeviceName(ipport.Split(':')[0], 30);
                lock (lock__WebSocketList)
                {
                    if (_WebSocketList.ContainsKey(ipport))
                    {
                        _WebSocketList[ipport].deviceName = deviceName[0];
                        _WebSocketList[ipport].ipcRobotName = deviceName[1];
                        _WebSocketList[ipport].socket = handler;
                    }
                    else
                    { _WebSocketList.Add(ipport, new rmsConectUserData(deviceName[0], deviceName[1], handler)); }
                }


                byte[] buffer = new byte[1024];
                string headerResponse = string.Empty;
                byte[] sendBuffer = null;

                if (listener != null && listener.IsBound)
                {
                    int rcvBytes = handler.Receive(buffer);
                    headerResponse = getStrFromByte(buffer, rcvBytes);
                    if (handler != null)
                    {
                        prepareClientResponse(headerResponse, ref sendBuffer);
                        handler.Send(sendBuffer);
                    }
                }
                if (IsWork)
                {
                    handler.IOControl(IOControlCode.KeepAliveValues, ToolFunStatic.KeepAlive(1, 5000, 1000), null);

                    // Create the state object.
                    StateObject state = new StateObject();
                    state.workSocket = handler;
                    handler.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(ReadCallback), state);
                }
            }
            catch
            {
                handler.Close();
                if (ipport != "")
                {
                    lock (lock__WebSocketList)
                    {
                        _WebSocketList.Remove(ipport);
                    }
                }
            }
        }
        private void ReadCallback(IAsyncResult ar)
        {
            string content = string.Empty;
            StateObject state = (StateObject)ar.AsyncState;
            Socket handler = state.workSocket;
            string ipport = "";// = handler.RemoteEndPoint.ToString();
            try
            {
                ipport = handler.RemoteEndPoint.ToString();
                if (!handler.Connected)
                {
                    lock (lock__WebSocketList)
                    {
                        _WebSocketList.Remove(ipport);
                    }
                    return;
                }
                // Read data from the client socket. 
                int bytesRead = 0;

                bytesRead = handler.EndReceive(ar);

                if (IsWork && bytesRead > 0)
                {
                    //DeserializeObject
                    content = Read(ipport, state.buffer);
                    //###???當buffer超過4096會 Error
                    if (content != "")
                    {
                        RMSProtocol pp = JsonConvert.DeserializeObject<RMSProtocol>(content);

                        //List<byte> c = new List<byte>();

                        Rms_ResolveData2RMSProtocol_Web(ipport, pp);
                    }
                    if (IsWork)
                    {
                        handler.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(ReadCallback), state);
                    }
                }
                else
                {
                    handler.Close();
                    //###???_proxyServer.DeleteUserList(user);
                }
            }
            catch (Exception ex)// ex)
            {
                handler.Close();
                lock (lock__WebSocketList)
                {
                    _WebSocketList.Remove(ipport);
                }
                //###???_proxyServer.DeleteUserList(user);
            }
        }
        private void Rms_ResolveData2RMSProtocol_Web(string ipport, RMSProtocol obj)
        {
            if (obj == null) { return; }
            string ipDis = ipport;// IpportConvertIPandName(ipport);
            if (!_WebSocketList.ContainsKey(ipport))
            {
                //++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_RunError(1, ipDis, "Parsing Network Packets", RMSError.Socket_ConnectLose, LogSourceName.IPPort, ipport, 261, ToolFun.StringAdd("接收到訊號但來源socket無紀錄,會導致資料無法正常執行,請檢察網路或重啟Service Engine. Type=", obj.DataType.ToString(), " Data=", obj.Data));
                //RunError A26
                return;
            }
            uint logid = 1;
            string[] cmd = null;
            try
            {
                rmsConectUserData sender = _WebSocketList[ipport];
                lock (lock_WebSocket)
                {
                    if (logid > 4290000000)
                    {
                        //SoftNetService.Program._NLogMain.NewOtherLogFile();
                        logID = 2;
                    }
                    logid = ++logID;
                }
                switch (obj.DataType)
                {
                    case 0:
                        //RmsSend(sender.socket, 1, "GetRMSLicenseList", 1);
                        Send(sender.socket, "ABV");
                        //RmsSend(sender.socket, false, new RMSProtocol(3, JsonConvert.SerializeObject("ABC"), obj.PoolID), logid);
                        break;
                    case 1:
                        if (obj.Data != null && obj.Data.Length > 0)
                        {
                            cmd = obj.Data.Split(',');
                            switch (cmd[0])
                            {
                                case "WebSocket_Login":
                                    if (sender.deviceName == "" && cmd.Length >= 2)
                                    {
                                        lock (lock__WebSocketList)
                                        {
                                            sender.deviceName = cmd[1];
                                            sender.ipcRobotName = cmd[2];
                                        }
                                    }
                                    Send(sender.socket, $"ReturnIPPort,{ipport}");
                                    break;
                            }
                        }
                        break;
                    default:
                        //++RMSDBErrorCount;
                        //SoftNetService.Program._NLogMain.Write_RunError(logid, ipDis, "Parsing Network Packets", RMSError.Service_Protocol_NoDefine, LogSourceName.ProtocolType, obj.DataType.ToString(), 181, "系統異常:接收到訊號,但程式無定義,請聯絡TMM.");
                        //RunError A18
                        sender.socket.Close();
                        lock (lock__WebSocketList)
                        {
                            _WebSocketList.Remove(ipport);
                        }
                        return;
                }
            }
            catch (Exception ex)
            {
                //if (lock__WebSocketList == null || lock_ThingsConfig == null || lock_MonitorConfig == null || lock_groupsItemList == null || lock_dashboard_thingNameData == null) { return; }
                //++RMSDBErrorCount;
                //SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "rms_ResolveData2RMSProtocol", RMSError.Service_Exception,
                //    LogSourceName.IPPort, ipDis, 61, ToolFun.StringAdd("通訊資料異常,請重啟Service Engine或聯絡TMM. data:", obj.Data, " Exception:", ex.Message), ex);
                //ExceptionError A6
                string _s = "";
            }

        }
        private string Read(string user, byte[] _dataBuffer)
        {

            // 判斷是否為最後一個Frame(第一個bit為FIN若為1代表此Frame為最後一個Frame)，超過一個Frame暫不處理
            if (!((_dataBuffer[0] & 0x80) == 0x80))
            {
                //Debug.WriteLine("Exceed 1 Frame. Not Handle");
                return "";
            }
            // 是否包含Mask(第一個bit為1代表有Mask)，沒有Mask則不處理
            if (!((_dataBuffer[1] & 0x80) == 0x80))
            {
                //Debug.WriteLine("Exception: No Mask");
                return "";
            }
            // 資料長度 = dataBuffer[1] - 127
            var payloadLen = _dataBuffer[1] & 0x7F;
            var masks = new Byte[4];
            var payloadData = filterPayloadData(ref _dataBuffer, ref payloadLen, ref masks);
            // 使用WebSocket Protocol中的公式解析資料
            for (var i = 0; i < payloadLen; i++)
                payloadData[i] = (Byte)(payloadData[i] ^ masks[i % 4]);

            return Encoding.UTF8.GetString(payloadData);
        }
        private Byte[] filterPayloadData(ref byte[] _dataBuffer, ref int length, ref Byte[] masks)
        {
            Byte[] payloadData;
            switch (length)
            {
                // 包含16 bit Extend Payload Length
                case 126:
                    Array.Copy(_dataBuffer, 4, masks, 0, 4);
                    length = (UInt16)(_dataBuffer[2] << 8 | _dataBuffer[3]);
                    payloadData = new Byte[length];
                    Array.Copy(_dataBuffer, 8, payloadData, 0, length);
                    break;
                // 包含 64 bit Extend Payload Length
                case 127:
                    Array.Copy(_dataBuffer, 10, masks, 0, 4);
                    var uInt64Bytes = new Byte[8];
                    for (int i = 0; i < 8; i++)
                    {
                        uInt64Bytes[i] = _dataBuffer[9 - i];
                    }
                    UInt64 len = BitConverter.ToUInt64(uInt64Bytes, 0);

                    payloadData = new Byte[len];
                    for (UInt64 i = 0; i < len; i++)
                        payloadData[i] = _dataBuffer[i + 14];
                    break;
                // 沒有 Extend Payload Length
                default:
                    Array.Copy(_dataBuffer, 2, masks, 0, 4);
                    payloadData = new Byte[length];
                    Array.Copy(_dataBuffer, 6, payloadData, 0, length);
                    break;
            }
            return payloadData;
        }
        public void Send2(Socket handler, byte[] data)
        {
            object user = null;
            try
            {
                user = handler.RemoteEndPoint.ToString();
                handler.Send(data);
                //handler.BeginSend(data, 0, data.Length, 0, new AsyncCallback(SendCallback), handler);
            }
            catch
            {
                handler.Close();
                if (user != null)
                {
                    //###???_proxyServer.DeleteUserList(user);
                }
            }
        }
        public void Send(Socket handler, string data)
        {
            object user = null;
            try
            {
                user = handler.RemoteEndPoint.ToString();
                // 將資料字串轉成Byte
                var contentByte = Encoding.UTF8.GetBytes(data);
                var dataBytes = new List<byte>();

                if (contentByte.Length < 126)   // 資料長度小於126，Type1格式
                {
                    // 未切割的Data Frame開頭
                    dataBytes.Add((Byte)0x81);
                    dataBytes.Add((Byte)contentByte.Length);
                    dataBytes.AddRange(contentByte);
                }
                else if (contentByte.Length <= 65535)       // 長度介於126與65535(0xFFFF)之間，Type2格式
                {
                    dataBytes.Add((Byte)0x81);
                    dataBytes.Add((Byte)0x7E);              // 126
                                                            // Extend Data 加長至2Byte
                    dataBytes.Add((Byte)((contentByte.Length >> 8) & 0xFF));
                    dataBytes.Add((Byte)((contentByte.Length) & 0xFF));
                    dataBytes.AddRange(contentByte);
                }
                else                 // 長度大於65535，Type3格式
                {
                    dataBytes.Add((Byte)0x81);
                    dataBytes.Add((Byte)0x7F);              // 127
                                                            // Extned Data 加長至8Byte
                    /*
                                        dataBytes.Add((Byte)((contentByte.Length >> 56) & 0xFF));
                                        dataBytes.Add((Byte)((contentByte.Length >> 48) & 0xFF));
                                        dataBytes.Add((Byte)((contentByte.Length >> 40) & 0xFF));
                                        dataBytes.Add((Byte)((contentByte.Length >> 32) & 0xFF));
                                        dataBytes.Add((Byte)((contentByte.Length >> 24) & 0xFF));
                                        dataBytes.Add((Byte)((contentByte.Length >> 16) & 0xFF));
                                        dataBytes.Add((Byte)((contentByte.Length >> 8) & 0xFF));
                                        dataBytes.Add((Byte)((contentByte.Length) & 0xFF));
                     */
                    dataBytes.Add((Byte)0x00);
                    dataBytes.Add((Byte)0x00);
                    dataBytes.Add((Byte)0x00);
                    dataBytes.Add((Byte)0x00);
                    dataBytes.Add((Byte)((contentByte.Length >> 24)));
                    dataBytes.Add((Byte)((contentByte.Length >> 16)));
                    dataBytes.Add((Byte)((contentByte.Length >> 8)));
                    dataBytes.Add((Byte)((contentByte.Length) & 0xFF));
                    dataBytes.AddRange(contentByte);
                }
                //handler.Send(dataBytes.ToArray());
                byte[] byteData = dataBytes.ToArray();
                handler.BeginSend(byteData, 0, byteData.Length, 0, new AsyncCallback(SendCallback), handler);

            }
            catch
            {
                handler.Close();
                if (user != null)
                {
                    //###??? _proxyServer.DeleteUserList(user);
                }
            }
        }
        private void SendCallback(IAsyncResult ar)
        {
            Socket handler = (Socket)ar.AsyncState;
            object user = null;// = handler.RemoteEndPoint.ToString();
            try
            {
                user = handler.RemoteEndPoint.ToString();
                // Retrieve the socket from the state object.

                // Complete sending the data to the remote device.
                int bytesSent = handler.EndSend(ar);
                //Debug.WriteLine("Sent {0} bytes to client.", bytesSent);

            }
            catch
            {
                //Debug.WriteLine(e.ToString());
                //###???移除client
                //string _s = "";
                handler.Close();
                if (user != null)
                {
                    //###???_proxyServer.DeleteUserList(user);
                }
            }
        }
        private string uniqueID = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11";
        private SHA1 sha1Crypto = SHA1CryptoServiceProvider.Create();
        private void parseReceiveData(List<byte> data)
        {
            var mask = data.Skip(2).Take(4).ToArray();
            var pldata =
                data.Skip(6).Take(data[1] & 127).Select(
                (d, i) => (byte)(d ^ mask[i % 4]));
            List<byte> readdata = new List<byte>();
            readdata.AddRange(pldata);
            //Debug.WriteLine("Receive=" + Encoding.UTF8.GetString(readdata.ToArray()));
        }
        private string getStrFromByte(byte[] bytes, int sizeLen)
        {
            return Encoding.UTF8.GetString(bytes, 0, sizeLen);
        }
        private void prepareClientResponse(string header, ref byte[] retVal)
        {

            string key = getKeyFromHeader(header);
            string newKey = key + uniqueID;

            byte[] hashBytesToSend = sha1Crypto.ComputeHash(Encoding.ASCII.GetBytes(newKey));

            string respkey = Convert.ToBase64String(hashBytesToSend);

            string clientResp = "HTTP/1.1 101 Switching Protocols\r\n"
                               + "Upgrade: websocket\r\n"
                               + "Connection: Upgrade\r\n"
                               + "Sec-WebSocket-Accept: " + respkey + "\r\n\r\n";
            //+"Sec-WebSocket-Protocol: chat, superchat\r\n" 
            //+ "Sec-WebSocket-Version: 13\r\n";

            Trace.TraceInformation("prepareClientResponse -- Sending Response to Client: {0}", clientResp);

            retVal = Encoding.UTF8.GetBytes(clientResp);

        }
        private string getKeyFromHeader(string header)
        {
            string keyStr = header.Replace("ey:", "`")
                                   .Split('`')[1]
                                   .Replace("\r", "").Split('\n')[0]
                                   .Trim();
            return keyStr;
        }
        private string[] GetUserList_DeviceName(string ip, byte role, uint logid = 1)
        {
            string[] re = new string[] { "", "" };
            string ipport = "";
            if (role != 0)
            {
                #region
                KeyValuePair<string, rmsConectUserData> rud;
                try
                {
                    lock (lock__WebSocketList)
                    { rud = _WebSocketList.FirstOrDefault(x => x.Key.Split(':')[0] == ip && x.Value.Role == role); }
                    if (rud.Key != null)
                    {
                        re[0] = rud.Value.deviceName;
                        re[1] = rud.Value.ipcRobotName;
                        ipport = rud.Key;
                    }
                }
                catch (Exception ex)
                {
                    //if (lock__WebSocketList == null) { }
                    //else
                    //{
                    //    ++RMSDBErrorCount;
                    //    SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "getUserList_DeviceName", RMSError.Service_Exception,
                    //        LogSourceName.IP, ip, 61, ToolFun.StringAdd("進入Service的Socket身分認證失敗,會導致該IP身分功能異常,請檢察網路,或重啟Service Engine. Exception:", ex.Message), ex);
                    //    //ExceptionError A6
                    //}
                    return re;
                }
                #endregion
            }
            else
            {
                #region
                /*
                var rud = _WebSocketList.Where(x => x.Key.Split(':')[0] == ip);
                foreach (KeyValuePair<string, rmsConectUserData> d in rud)
                {
                    re[0] = d.Value.deviceName;
                    re[1] = d.Value.ipcRobotName;
                    ipport = d.Key;
                    break;
                }
                 */
                KeyValuePair<string, rmsConectUserData> rud;
                try
                {
                    lock (lock__WebSocketList)
                    {
                        rud = _WebSocketList.FirstOrDefault(x => x.Key.Split(':')[0] == ip);
                    }
                    if (rud.Key != null)
                    {
                        re[0] = rud.Value.deviceName;
                        re[1] = rud.Value.ipcRobotName;
                        ipport = rud.Key;
                    }
                }
                catch (Exception ex)
                {
                    if (lock__WebSocketList == null) { }
                    else
                    {
                        //++RMSDBErrorCount;
                        //SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "getUserList_DeviceName find _WebSocketList.FirstOrDefault", RMSError.Service_Exception,
                        //    LogSourceName.IP, ip, 61, ToolFun.StringAdd("連入Service的Socket身分認證失敗,會導致該IP身分功能異常,請檢察網路,或重啟Service Engine. Exception:", ex.Message), ex);
                        //ExceptionError A6
                    }
                    return re;
                }
                #endregion
            }
            if (re[0] == "")
            {
                #region
                string type = "";
                switch (role)
                {
                    case 1: type = " and Device_Type='RMS'"; break;
                    case 10: type = " and Device_Type='TMRobot'"; break;
                }
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    DataRow dr = db.DB_GetFirstDataByDataRow($"select DeviceName,IPCRobotName,RobotAutoVarSync from SoftNetSYSDB.[dbo].Device where ServerId='{_Fun.Config.ServerId}' and Device_IP='" + ip + "'" + type);
                    if (dr != null)
                    {
                        re[0] = dr["DeviceName"].ToString();
                        re[1] = dr["IPCRobotName"].ToString();
                        if (ipport != "")
                        {
                            try
                            {
                                lock (lock__WebSocketList)
                                {
                                    //rmsConectUserData rud = _WebSocketList[ipport];
                                    if (!dr.IsNull("RobotAutoVarSync")) { _WebSocketList[ipport].RobotAutoVarSync = bool.Parse(dr["RobotAutoVarSync"].ToString()); }
                                    _WebSocketList[ipport].deviceName = re[0];
                                    _WebSocketList[ipport].ipcRobotName = re[1];
                                    //_WebSocketList[ipport] = rud;
                                }
                            }
                            catch (Exception ex)
                            {
                                if (lock__WebSocketList == null) { }
                                else
                                {
                                    //++RMSDBErrorCount;
                                    //SoftNetService.Program._NLogMain.Write_ExceptionError(logid, "getUserList_DeviceName set _WebSocketList[ipport]", RMSError.Service_Exception,
                                    //    LogSourceName.IP, ip, 61, ToolFun.StringAdd("連入Service的Socket身分認證失敗,會導致該IP身分功能異常,請檢察網路,或重啟Service Engine. Exception:", ex.Message), ex);
                                    //ExceptionError A6
                                }
                                return re;
                            }
                        }
                    }
                }
                #endregion
            }
            return re;
        }
        public class StateObject
        {
            // Client socket.
            public Socket workSocket = null;
            // Size of receive buffer.
            public const int BufferSize = 65535;//v   65535
                                                // Receive buffer.
            public byte[] buffer = new byte[BufferSize];
            // Received data string.
            public StringBuilder sb = new StringBuilder();
            public string Text()
            {
                List<byte> data = buffer.ToList<byte>();
                var mask = data.Skip(2).Take(4).ToArray();
                var pldata =
                    data.Skip(6).Take(data[1] & 127).Select(
                    (d, i) => (byte)(d ^ mask[i % 4]));
                List<byte> readdata = new List<byte>();
                readdata.AddRange(pldata);
                return Encoding.UTF8.GetString(readdata.ToArray());
            }
        }
















    }



}
