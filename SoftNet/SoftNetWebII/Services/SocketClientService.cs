using Base;
using Base.Services;
using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{

    public class SocketClientService
    {
        public List<rmsMasterUserData> _MasterRMSUserList = new List<rmsMasterUserData>();
        private object lock__MasterRMSUserList = new object();
        private bool IsWork = false;
        private object lock_Sevice = new object();
        private uint logID = 1;
        public char RMSLogMode = '1'; //LogMode 0:Normal Mode 1:Debug Mode
        public SocketClientService()
        {
            #region 建立_MasterRMSUserList 與 RMS service之間的連線 5431
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
            IsWork = true;
            #endregion

            #region 建立管理用 5431 Socket Client自動連線 _MasterRMSUserList是有條件的控管
            Thread MasterSocketInit_thread = new Thread(() =>
            {
                CheckMasterSocketConnectIsOKthread_Tick();
            })
            {
                IsBackground = true
            };
            MasterSocketInit_thread.Start();
            #endregion

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
                            if (RmsLoopConnectII(_MasterRMSUserList[i].socket, _MasterRMSUserList[i].ipPoint))
                            {
                                #region 重新連線,並開Thraed讀資料
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
                                //###???
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
                if (_Fun.Is_Thread_ForceClose) { IsWork = false; }
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
                await _Log.SocketLogAsync("Socket5431Log",$"R Type={obj.DataType.ToString()} Data={obj.Data}");

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
                                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[OperateLog] (ServerId,[Id],[NeedId],[SimulationId],[LOGDateTime],[ProgramName],[OperateType],Remark,StationNO,PartNO,OrderNO,IndexSN) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('D')}','{dr_SimulationId["NeedId"].ToString()}','{dr_SimulationId["SimulationId"].ToString()}','{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff")}','LabelWork','{type}','{cmd[5]} {cmd[3]}','{cmd[0]}','','{cmd[1]}',{cmd[6]})");//###??? dr_M["PartNO"].ToString() 暫時放空

                                    //通知webSocket send
                                    var webSocketService = (SNWebSocketService)_Fun.DiBox.GetService(typeof(SNWebSocketService));
                                    foreach (KeyValuePair<string, rmsConectUserData> r in webSocketService._WebSocketList)
                                    {
                                        if (r.Key!=null && r.Value.socket!=null)
                                        {
                                             webSocketService.Send(r.Value.socket, "StationStatusChange");
                                        }
                                    }

                                }
                                break;
                            case "5"://通知
                                {
                                    //通知webSocket send
                                    var webSocketService = (SNWebSocketService)_Fun.DiBox.GetService(typeof(SNWebSocketService));
                                    foreach (KeyValuePair<string, rmsConectUserData> r in webSocketService._WebSocketList)
                                    {
                                        if (r.Key != null && r.Value.socket != null)
                                        {
                                            webSocketService.Send(r.Value.socket, "StationStatusChange");
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
    }

}
