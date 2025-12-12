using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Net.Sockets;
using Modbus.IO;
using Modbus.Utility;
using System.Xml;
using System.Net;
using System.Threading;

namespace Modbus.Device
{
    /// <summary>
    /// Modbus TCP slave device.
    /// </summary>
    public class ModbusTcpSlave : ModbusSlave
    {
        private readonly object _mastersLock = new object();
        private readonly object _serverLock = new object();
        //private readonly ILog _logger = LogManager.GetLogger(typeof(ModbusTcpSlave));
        private readonly Dictionary<string, ModbusMasterTcpConnection> _masters = new Dictionary<string, ModbusMasterTcpConnection>();
        private TcpListener _server;
        private XmlNodeList iPfilter = null;
        private List<string[]> _ipfilter = null; // [0] startIP 192.168.1.1 [1] endIP 255 // 192.168.1.1~255


        private ModbusTcpSlave(byte unitId, TcpListener tcpListener, XmlNodeList iPfilter)
            : base(unitId, new EmptyTransport())
        {
            if (tcpListener == null)
                throw new ArgumentNullException("tcpListener");

            this.iPfilter = iPfilter;
            _server = tcpListener;
        }
        private ModbusTcpSlave(byte unitId, TcpListener tcpListener, List<string[]> ipfilter)
            : base(unitId, new EmptyTransport())
        {
            if (tcpListener == null)
                throw new ArgumentNullException("tcpListener");

            this._ipfilter = ipfilter;
            _server = tcpListener;
        }

        /// <summary>
        /// Gets the Modbus TCP Masters connected to this Modbus TCP Slave.
        /// </summary>
        public ReadOnlyCollection<TcpClient> Masters
        {
            get
            {
                lock (_mastersLock)
                {
                    return new ReadOnlyCollection<TcpClient>(
                        SequenceUtility.ToList(_masters.Values, delegate(ModbusMasterTcpConnection connection) { return connection.TcpClient; }));
                }
            }
        }

        /// <summary>
        /// Gets the server.
        /// </summary>
        /// <value>The server.</value>
        /// <remarks>
        /// This property is not thread safe, it should only be consumed within a lock.
        /// </remarks>
        private TcpListener Server
        {
            get
            {
                //if (_server == null)
                //{
                //    throw new ObjectDisposedException("Server");
                //}
                return _server;
            }
        }

        /// <summary>
        /// Modbus TCP slave factory method.
        /// </summary>
        public static ModbusTcpSlave CreateTcp(byte unitId, TcpListener tcpListener, XmlNodeList iPfilter)
        {
            return new ModbusTcpSlave(unitId, tcpListener, iPfilter);
        }
        /// <summary>
        /// Modbus TCP slave factory method.
        /// </summary>
        public static ModbusTcpSlave CreateTcp(byte unitId, TcpListener tcpListener, List<string[]> ipfilter)
        {
            return new ModbusTcpSlave(unitId, tcpListener, ipfilter);
        }

        /// <summary>
        /// Start slave listening for requests.
        /// </summary>
        public override void Listen()
        {
            //###???_logger.Debug("Start Modbus Tcp Server.");
            if (LogUtility.IsEnable) LogUtility.Debug("Start Modbus Tcp Server.");

            lock (_serverLock)
            {
                try
                {
                    if (Server != null)
                    {
                        Server.Start();
#if true // method 1, BeginAccept, BeginRead, BeginWrite
                        // use Socket async API for compact framework compat
                        if (Server.Server != null) Server.Server.BeginAccept(AcceptCompleted, this);

#else // method 2, TcpListen thread
                        // create a background thread to Tcp Listen (for client tcp service)
                        Thread _listen_thread = new Thread(() => ModbusSlave_TcpListenThread());
                        _listen_thread.IsBackground = true;
                        _listen_thread.Start();

                        // create a background thread to handle request message from modbus master
                        Thread _request_thread = new Thread(() => ModbusSlave_RequestThread());
                        _request_thread.IsBackground = true;
                        _request_thread.Start();
#endif
                    }
                }
                catch (ObjectDisposedException)
                {
                    // this happens when the server stops
                }
            }
        }
        /// <summary>
        /// Stop TCP Master.
        /// </summary>
        public override void Stop()
        {
            // double-check locking
            if (Server != null)
            {
                lock (_serverLock)
                {
                    try
                    {
                        if (Server != null)
                        {
                            SocketDisconnect(Server.Server);
                            Server.Stop();
                        }
                    }
                    catch (Exception)// ex)
                    {
                        // this happens when the server stops
                    }
                }
            }
            _is_tcp_listening = false;
        }
        /// <summary>
        /// Releases unmanaged and - optionally - managed resources
        /// </summary>
        /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        /// <remarks>Dispose is thread-safe.</remarks>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Stop();
                _server = null;
            }
        }

        internal void RemoveMaster(string endPoint)
        {
            lock (_mastersLock)
            {
                if (!_masters.Remove(endPoint))
                    throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "EndPoint {0} cannot be removed, it does not exist.", endPoint));
            }

            //###???_logger.InfoFormat("Removed Master {0}", endPoint);
            if (LogUtility.IsEnable) LogUtility.Debug(string.Format("Removed Master {0}", endPoint), "RemoveMaster");
        }

        private int[] returnIPValueArray(string ip)
        {
            int[] re = new int[4];
            string[] data = ip.Split('.');
            for (int i = 0; i < data.Length; i++)
            {
                if (i > 3) { break; }
                if (!int.TryParse(data[i], out re[i])) { re[i] = 0; }
            }
            return re;
        }
        private bool checkIPfilter(Socket sock)
        {
            if (sock == null) return false;

            if (iPfilter != null && iPfilter.Count > 0) // XmlNodeList //ipfilter的ip區間為白名單
            {
                int[] remoteip = returnIPValueArray(sock.RemoteEndPoint.ToString().Split(':')[0]);
                foreach (XmlNode node in iPfilter)
                {
                    if (node.Attributes["startIP"] != null && node.Attributes["endIP"] != null)
                    {
                        int[] comps = returnIPValueArray(node.Attributes["startIP"].Value); // start ip
                        int compe = 0;
                        if (!int.TryParse(node.Attributes["endIP"].Value, out compe)) compe = 0; // end ip

                        if (comps.Length == 4 && remoteip.Length == 4 &&
                            comps[0] == remoteip[0] && comps[1] == remoteip[1] && comps[2] == remoteip[2])
                        {
                            if (remoteip[3] >= comps[3] && remoteip[3] <= compe)
                            {
                                return true; // ip ok
                            }
                        }
                    }
                }
            }
            else if (_ipfilter != null && _ipfilter.Count > 0) // List<string[]> //ipfilter的ip區間為白名單
            {
                int[] remoteip = returnIPValueArray(sock.RemoteEndPoint.ToString().Split(':')[0]); // remote ip
                for (int i = 0; i < _ipfilter.Count; i++)
                {
                    string[] ipf = _ipfilter[i];
                    if (ipf == null || ipf.Length < 2) continue;
                    if (String.IsNullOrEmpty(ipf[0]) == false && String.IsNullOrEmpty(ipf[1]) == false)
                    {
                        int[] comps = returnIPValueArray(ipf[0]); // start ip
                        int compe = 0;
                        if (!int.TryParse(ipf[1], out compe)) compe = 0; // end ip

                        if (comps.Length == 4 && remoteip.Length == 4 &&
                            comps[0] == remoteip[0] && comps[1] == remoteip[1] && comps[2] == remoteip[2])
                        {
                            if (remoteip[3] >= comps[3] && remoteip[3] <= compe)
                            {
                                return true; // ip ok
                            }
                        }
                    }
                }
            }
            else // no ip filter, always true
            {
                return true; // ip ok
            }
            return false; // reject ip
        }
        internal void AcceptCompleted(IAsyncResult ar)
        {
            ModbusTcpSlave slave = null;
            try
            {
                slave = (ModbusTcpSlave)ar.AsyncState;

                // use Socket async API for compact framework compat
                Socket socket = null;
                lock (_serverLock)
                {
                    if (Server != null && Server.Server != null)
                        socket = Server.Server.EndAccept(ar);
                }

                bool run = checkIPfilter(socket);
                if (run)
                {
                    TcpClient client = new TcpClient { Client = socket };

                    var masterConnection = new ModbusMasterTcpConnection(client, slave);
                    masterConnection.ModbusMasterTcpConnectionClosed += (sender, eventArgs) => RemoveMaster(eventArgs.EndPoint);

                    lock (_mastersLock)
                        _masters.Add(client.Client.RemoteEndPoint.ToString(), masterConnection);
                }
                else
                {
                    SocketDisconnect(socket);
                    socket = null;
                    //###???通知被close
                }
                //###???_logger.Debug("Accept completed.");
                if (LogUtility.IsEnable) LogUtility.Debug("Accept completed.", "AcceptCompleted");

                // Accept another client
                // use Socket async API for compact framework compat
                lock (_serverLock)
                {
                    if (Server != null && Server.Server != null)
                        Server.Server.BeginAccept(AcceptCompleted, slave);
                }
            }
            catch (ObjectDisposedException)
            {
                // this happens when the server stops
                lock (_serverLock)
                {
                    if (Server != null && Server.Server != null)
                        Server.Server.BeginAccept(AcceptCompleted, slave);
                }
            }
            catch (Exception ex)
            {
                if (IsWork() == true)
                {
                    this.TragetError(0x44002, "Modbus Slave AcceptCompleted error. " + ex.Message);
                    lock (_serverLock)
                    {
                        if (Server != null && Server.Server != null)
                            Server.Server.BeginAccept(AcceptCompleted, slave);
                    }
                }
                //lock (_serverLock)
                //    Server.Server.BeginAccept(AcceptCompleted, slave);

            }
        }

        // method 2, TcpListen thread
        private bool _is_tcp_listening = false;

        private object _message_Lock = new object();
        private List<ModbusTcpMessage> _request_message;
        struct ModbusTcpMessage
        {
            public Socket socket;  // Socket
            public byte[] request; // modbus master request message
            public ModbusTcpMessage(Socket sock, byte[] frame)
            {
                socket = sock;
                request = frame;
            }
        }
        // __Thread__ // for Server TcpListener
        private void ModbusSlave_TcpListenThread()
        {
            try
            {
                _is_tcp_listening = true;

                lock (_message_Lock)
                {
                    _request_message = new List<ModbusTcpMessage>();
                    _request_message.Clear();
                }

                while (Server != null && _is_tcp_listening == true)
                {
                    // check whether is any connection request
                    if (Server != null && Server.Pending() == false)
                    {
                        SpinWait.SpinUntil(() => false, 100);
                        continue;
                    }

                    // accept client socket
                    Socket client = null;
                    if (Server != null) client = Server.AcceptSocket();

                    // check ip filter
                    bool run = checkIPfilter(client);
                    if (run == false) // could not accept, close it.
                    {
                        SocketDisconnect(client);
                        client = null;
                    }
                    else
                    {
                        // create thread for tcp receive service
                        Thread _receive_thread = new Thread(new ParameterizedThreadStart(ModbusSlave_TcpReceive));
                        _receive_thread.IsBackground = true;
                        _receive_thread.Start(client);
                    }
                }
            }
            catch (Exception ex)
            {
                //MsgBoxLogger.Show(ex.Message, MsgBoxLogger.StatusType.LoggerOnly);
            }

            _is_tcp_listening = false;
            // stop listener
            try { Stop(); }
            catch (Exception ex)
            {
                //MsgBoxLogger.Show(ex.Message, MsgBoxLogger.StatusType.LoggerOnly);
            }
        }
        // __Thread__ // for Server to service client RX
        private void ModbusSlave_TcpReceive(object param)
        {
            Socket client = null;
            try
            {
                client = param as Socket;
                if (client == null) return;
                SetKeepAlive(client, 5000, 1000);
            }
            catch (Exception)// ex)
            {
                SocketDisconnect(client);
                client = null;
                return;
            }
            try
            {
                if (client == null) return;

                EndPoint remote_ep = client.RemoteEndPoint;
                List<byte> RecvBuffer = new List<byte>();

                while (Server != null && _is_tcp_listening == true)
                {
                    bool flag = SpinWait.SpinUntil(() => (Server == null || _is_tcp_listening == false ||
                                                          client == null || client.Connected == false ||
                                                          client.Poll(0, SelectMode.SelectRead)), -1);
                    if (Server == null || _is_tcp_listening == false) break;
                    if (client == null) break;
                    if (client.Connected == false) break;

                    int length = client.Available;
                    if (length == 0) break; // 當大於0時表示傳遞資料過來 等於0時表示"正常"斷線

                    byte[] buffer = new byte[length];
                    var received_length = client.ReceiveFrom(buffer, ref remote_ep); // hold in here
                    if (received_length == 0) // 當大於0時表示傳遞資料過來 等於0時表示"正常"斷線
                    {
                        break;
                    }
                    else
                    {
                        // add receive data
                        RecvBuffer.AddRange(buffer);
                        // decode receive buffer
                        while (Server != null && client != null)
                        {
                            byte[] request = OnReceive(ref RecvBuffer);
                            if (request == null) // null is waitting next received data
                            {
                                break;
                            }
                            else // not null is decode a frame ok
                            {
                                // add request message
                                lock (_message_Lock)
                                {
                                    ModbusTcpMessage item = new ModbusTcpMessage(client, request);
                                    _request_message.Add(item);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)// ex)
            {
            }
            SocketDisconnect(client);
            client = null;
        }
        // decode received bytes
        private byte[] OnReceive(ref List<byte> recvbytes)
        {
            try // parsing received data
            {
                if (recvbytes == null || recvbytes.Count < 6) return null; // least 6 chars

                ushort frameLength = (ushort)(recvbytes[4] << 8 | recvbytes[5]);

                if (recvbytes == null || recvbytes.Count < (6 + frameLength)) return null;

                byte[] data = new byte[6 + frameLength];

                for (int i = 0; i < data.Length; i++)
                {
                    if (i >= recvbytes.Count) return null;
                    data[i] = recvbytes[i];
                }
                recvbytes.RemoveRange(0, data.Length);
                return data;
            }
            catch (Exception)// ex)
            {
            }
            return null;
        }
        // check request message 
        private bool isRequestMessage()
        {
            lock (_message_Lock)
            {
                if (_request_message != null && _request_message.Count >= 1) return true;
                return false;
            }
        }
        // __Thread__  // handle request message
        private void ModbusSlave_RequestThread()
        {
            SpinWait.SpinUntil(() => _is_tcp_listening == true, 1000); // wait listening to true

            uint HoldCount = 0;
            while (Server != null && _is_tcp_listening == true)
            {
                try
                {
                    if (isRequestMessage() == false)
                    {
                        HoldCount = 0;
                        SpinWait.SpinUntil(() => (Server == null || _is_tcp_listening == false || isRequestMessage() == true), 100); // timeout
                    }
                    else
                    {
                        HoldCount++;
                        if (HoldCount >= 0x1000)
                        {
                            HoldCount = 0;
                            Thread.Sleep(2);
                        }
                    }
                    if (Server == null || _is_tcp_listening == false) break;
                    if (isRequestMessage() == false) continue;

                    ModbusTcpMessage item;
                    lock (_message_Lock)
                    {
                        item = _request_message[0];
                        _request_message.RemoveAt(0);
                    }
                    if (item.socket == null || item.request == null) continue;

                    // handle request message
                    TcpClient client = new TcpClient { Client = item.socket };
                    var masterConnection = new ModbusMasterTcpConnection(client, this, item.request);
                }
                catch (Exception)// ex)
                {
                }
            }
            lock (_message_Lock)
            {
                for (int i = 0; _request_message != null && i < _request_message.Count; i++)
                {
                    SocketDisconnect(_request_message[i].socket);
                }
                _request_message.Clear();
            }
        }
        /// <summary>
        /// Disconnect a Socket
        /// </summary>
        /// <param name="sock"></param>
        /// <returns></returns>
        private bool SocketDisconnect(Socket sock)
        {
            if (sock != null)
            {
                try { sock.Shutdown(SocketShutdown.Both); }
                catch { }
                try { sock.Disconnect(false); }
                catch { }
                try { sock.Close(); }
                catch { }
                try { sock.Dispose(); }
                catch { }
                return true;
            }
            return false;
        }
        /// <summary>
        /// this function is used to set keepalive option of socket
        /// </summary>
        /// <param name="sock"></param>
        /// works on which socket
        /// <param name="time"></param>
        /// if there is no any conmunication in this scoket, after "time", starting sending beacon to check whether the remote site is alive
        /// "time" is millisecond.
        /// <param name="interval"></param>
        ///  time interval between each two detection beacons
        /// <returns></returns>
        private bool SetKeepAlive(Socket sock, ulong time, ulong interval)
        {
            if (sock == null) return false;
            const int bytesperlong = 4; // 32 / 8
            const int bitsperbyte = 8;
            try
            {
                // resulting structure
                byte[] SIO_KEEPALIVE_VALS = new byte[3 * bytesperlong];

                // array to hold input values
                ulong[] input = new ulong[3];

                // put input arguments in input array
                if (time == 0 || interval == 0) // enable disable keep-alive
                    input[0] = (0UL); // off
                else
                    input[0] = (1UL); // on

                input[1] = (time); // time millis
                input[2] = (interval); // interval millis

                // pack input into byte struct
                for (int i = 0; i < input.Length; i++)
                {
                    SIO_KEEPALIVE_VALS[i * bytesperlong + 3] = (byte)(input[i] >> ((bytesperlong - 1) * bitsperbyte) & 0xff);
                    SIO_KEEPALIVE_VALS[i * bytesperlong + 2] = (byte)(input[i] >> ((bytesperlong - 2) * bitsperbyte) & 0xff);
                    SIO_KEEPALIVE_VALS[i * bytesperlong + 1] = (byte)(input[i] >> ((bytesperlong - 3) * bitsperbyte) & 0xff);
                    SIO_KEEPALIVE_VALS[i * bytesperlong + 0] = (byte)(input[i] >> ((bytesperlong - 4) * bitsperbyte) & 0xff);
                }
                // create bytestruct for result (bytes pending on server socket)
                byte[] result = BitConverter.GetBytes(0);
                // write SIO_VALS to Socket IOControl
                sock.IOControl(IOControlCode.KeepAliveValues, SIO_KEEPALIVE_VALS, result);
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

    }

}
