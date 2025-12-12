using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Threading;
using Modbus.IO;
using Modbus.Message;
using Modbus.Utility;



namespace Modbus.Device
{
    internal class ModbusMasterTcpConnection : ModbusDevice, IDisposable
    {

        //private readonly ILog _logger = LogManager.GetLogger(Assembly.GetCallingAssembly(),
            //String.Format(CultureInfo.InvariantCulture, "{0}.Instance{1}", typeof(ModbusMasterTcpConnection).FullName, Interlocked.Increment(ref _instanceCounter)));

        private readonly TcpClient _client;
        private readonly string _endPoint;
        private readonly Stream _stream;
        private readonly ModbusTcpSlave _slave;
        //////private static int _instanceCounter;
        private byte[] _mbapHeader = new byte[6];
        private byte[] _messageFrame;

        public ModbusMasterTcpConnection(TcpClient client, ModbusTcpSlave slave)
            : base(new ModbusIpTransport(new TcpClientAdapter(client)))
        {
            if (client == null)
                throw new ArgumentNullException("client");
            if (slave == null)
                throw new ArgumentException("slave");

            _client = client;
            _endPoint = client.Client.RemoteEndPoint.ToString();
            _stream = client.GetStream();
            _slave = slave;
            //###???_logger.DebugFormat("Creating new Master connection at IP:{0}", EndPoint);
            if (LogUtility.IsEnable) LogUtility.Debug(string.Format("Creating new Master connection at IP:{0}", EndPoint), "ModbusMasterTcpConnection");

            //###???_logger.Debug("Begin reading header.");
            if (LogUtility.IsEnable) LogUtility.Debug(string.Format("Begin reading header."), "ModbusMasterTcpConnection");

            Stream.BeginRead(_mbapHeader, 0, 6, ReadHeaderCompleted, null);
            // BeginRead(Header)->BeginRead(Frame)->ApplyRequest->BeginWrite(Response)->BeginRead(Header)...
        }
        public ModbusMasterTcpConnection(TcpClient client, ModbusTcpSlave slave, byte[] frame)
            : base(new ModbusIpTransport(new TcpClientAdapter(client)))
        {
            if (client == null)
                throw new ArgumentNullException("client");
            if (slave == null)
                throw new ArgumentException("slave");

            _client = client;
            _endPoint = client.Client.RemoteEndPoint.ToString();
            _stream = client.GetStream();
            _slave = slave;

            //###???_logger.DebugFormat("Creating new Master connection at IP:{0}", EndPoint);
            if (LogUtility.IsEnable) LogUtility.Debug(string.Format("Creating new Master connection at IP:{0}", EndPoint), "ModbusMasterTcpConnection");

            IModbusMessage request = ModbusMessageFactory.CreateModbusRequest(CollectionUtility.Slice(frame, 6, frame.Length - 6));
            request.TransactionId = (ushort)IPAddress.NetworkToHostOrder(BitConverter.ToInt16(frame, 0));

            if (LogUtility.IsEnable) LogUtility.Debug(string.Format("RX: {0}", StringUtility.Join(", ", frame)), "ModbusMasterTcpConnection");

            // perform action and build response
            IModbusMessage response = _slave.ApplyRequest(request);
            response.TransactionId = request.TransactionId;

            // write response
            byte[] responseFrame = Transport.BuildMessageFrame(response);

            //###???_logger.InfoFormat("TX: {0}", StringUtility.Join(", ", responseFrame));
            if (LogUtility.IsEnable) LogUtility.Debug(string.Format("TX: {0}", StringUtility.Join(", ", responseFrame)), "ModbusMasterTcpConnection");

            Stream.BeginWrite(responseFrame, 0, responseFrame.Length, WriteCompletedWithoutRead, null);
        }

        /// <summary>
        /// Occurs when a Modbus master TCP connection is closed.
        /// </summary>
        public event EventHandler<TcpConnectionEventArgs> ModbusMasterTcpConnectionClosed;

        public string EndPoint
        {
            get
            {
                return _endPoint;
            }
        }

        public Stream Stream
        {
            get
            {
                return _stream;
            }
        }

        public TcpClient TcpClient
        {
            get
            {
                return _client;
            }
        }

        internal void ReadHeaderCompleted(IAsyncResult ar)
        {
            //###???_logger.Debug("Read header completed.");
            if (LogUtility.IsEnable) LogUtility.Debug(string.Format("Read header completed."), "ReadHeaderCompleted");

            CatchExceptionAndRemoveMasterEndPoint(() =>
            {
                // this is the normal way a master closes its connection
                if (Stream.EndRead(ar) == 0)
                {
                    //###???_logger.Debug("0 bytes read, Master has closed Socket connection.");
                    if (LogUtility.IsEnable) LogUtility.Debug(string.Format("0 bytes read, Master has closed Socket connection."), "ReadHeaderCompleted");

                    EventHandler<TcpConnectionEventArgs> handler = ModbusMasterTcpConnectionClosed;
                    if (handler != null)
                        handler(this, new TcpConnectionEventArgs(EndPoint));

                    return;
                }

                //###???_logger.DebugFormat("MBAP header: {0}", StringUtility.Join(", ", _mbapHeader));
                if (LogUtility.IsEnable) LogUtility.Debug(string.Format("MBAP header: {0}", StringUtility.Join(", ", _mbapHeader)), "ReadHeaderCompleted");

                ushort frameLength = (ushort)IPAddress.HostToNetworkOrder(BitConverter.ToInt16(_mbapHeader, 4));

                //###???_logger.DebugFormat("{0} bytes in PDU.", frameLength);
                if (LogUtility.IsEnable) LogUtility.Debug(string.Format("{0} bytes in PDU.", frameLength), "ReadHeaderCompleted");

                _messageFrame = new byte[frameLength];

                Stream.BeginRead(_messageFrame, 0, frameLength, ReadFrameCompleted, null);
            }, EndPoint, 1);
        }

        internal void ReadFrameCompleted(IAsyncResult ar)
        {
            CatchExceptionAndRemoveMasterEndPoint(() =>
            {
                //###???_logger.DebugFormat("Read Frame completed {0} bytes", Stream.EndRead(ar));
                if (LogUtility.IsEnable) LogUtility.Debug(string.Format("Read Frame completed {0} bytes", Stream.EndRead(ar)), "ReadFrameCompleted");

                byte[] frame = CollectionUtility.Concat(_mbapHeader, _messageFrame);

                //###???_logger.InfoFormat("RX: {0}", StringUtility.Join(", ", frame));
                if (LogUtility.IsEnable) LogUtility.Debug(string.Format("RX: {0}", StringUtility.Join(", ", frame)), "ReadFrameCompleted");

                IModbusMessage request = ModbusMessageFactory.CreateModbusRequest(CollectionUtility.Slice(frame, 6, frame.Length - 6));
                request.TransactionId = (ushort)IPAddress.NetworkToHostOrder(BitConverter.ToInt16(frame, 0));

                // perform action and build response
                IModbusMessage response = _slave.ApplyRequest(request);
                response.TransactionId = request.TransactionId;

                // write response
                byte[] responseFrame = Transport.BuildMessageFrame(response);

                //###???_logger.InfoFormat("TX: {0}", StringUtility.Join(", ", responseFrame));
                if (LogUtility.IsEnable) LogUtility.Debug(string.Format("TX: {0}", StringUtility.Join(", ", responseFrame)), "ReadFrameCompleted");

                Stream.BeginWrite(responseFrame, 0, responseFrame.Length, WriteCompleted, null);
            }, EndPoint, 2);
        }

        internal void WriteCompleted(IAsyncResult ar)
        {
            //###???_logger.Debug("End write.");
            if (LogUtility.IsEnable) LogUtility.Debug(string.Format("End write."), "WriteCompleted");

            CatchExceptionAndRemoveMasterEndPoint(() =>
            {
                Stream.EndWrite(ar);
                //###???_logger.Debug("Begin reading another request.");
                if (LogUtility.IsEnable) LogUtility.Debug(string.Format("Begin reading another request."), "WriteCompleted");

                Stream.BeginRead(_mbapHeader, 0, 6, ReadHeaderCompleted, null);
            }, EndPoint, 3);
        }
        internal void WriteCompletedWithoutRead(IAsyncResult ar)
        {
            //###???_logger.Debug("End write.");
            if (LogUtility.IsEnable) LogUtility.Debug(string.Format("End write."), "WriteCompletedWithoutRead");

            CatchExceptionAndRemoveMasterEndPoint(() =>
            {
                Stream.EndWrite(ar);
            }, EndPoint, 4);
        }

        /// <summary>
        /// Catches all exceptions thrown when action is executed and removes the master end point.
        /// The exception is ignored when it simply signals a master closing its connection.
        /// </summary>
        internal void CatchExceptionAndRemoveMasterEndPoint(System.Action action, string endPoint, int index = 0)
        {
            if (action == null)
                throw new ArgumentNullException("action");
            if (endPoint == null)
                throw new ArgumentNullException("endPoint");
            if (String.IsNullOrEmpty(endPoint))
                throw new ArgumentException("Argument endPoint cannot be empty.");

            try
            {
                action.Invoke();
            }
            catch (IOException ioe)
            {
                //###???_logger.DebugFormat("IOException encountered in ReadHeaderCompleted - {0}", ioe.Message);
                if (LogUtility.IsEnable) LogUtility.Debug(string.Format("IOException encountered in ReadHeaderCompleted - {0}", ioe.Message), "CatchExceptionAndRemoveMasterEndPoint");

                EventHandler<TcpConnectionEventArgs> handler = ModbusMasterTcpConnectionClosed;
                if (handler != null)
                    handler(this, new TcpConnectionEventArgs(EndPoint));

                SocketException socketException = ioe.InnerException as SocketException;
                if (socketException != null && socketException.ErrorCode == Modbus.ConnectionResetByPeer)
                {
                    //###???_logger.Debug("Socket Exception ConnectionResetByPeer, Master closed connection.");
                    if (LogUtility.IsEnable) LogUtility.Debug(string.Format("Socket Exception ConnectionResetByPeer, Master closed connection."), "CatchExceptionAndRemoveMasterEndPoint");
                    return;
                }
                _slave.TragetError(0x44001, " BeginRead error " + index.ToString() + " " + ioe.ToString());
            }
            catch (Exception e)
            {
                _slave.TragetError(0x44001, " BeginRead error " + index.ToString() + " " + e.ToString());
                //###???_logger.Error("Unexpected exception encountered", e);
                if (LogUtility.IsEnable) LogUtility.Debug(string.Format("Unexpected exception encountered", e), "CatchExceptionAndRemoveMasterEndPoint");
            }
        }
    }
}
