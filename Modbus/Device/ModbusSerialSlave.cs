using System;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using Modbus.IO;
using Modbus.Message;
using Modbus.Utility;

namespace Modbus.Device
{
	/// <summary>
	/// Modbus serial slave device.
	/// </summary>
	public class ModbusSerialSlave : ModbusSlave
	{
		//private static readonly ILog _logger = LogManager.GetLogger(typeof(ModbusSerialSlave));
        private static SerialPort _sp = null;

		private ModbusSerialSlave(byte unitId, ModbusTransport transport)
			: base(unitId, transport)
		{
		}

		private ModbusSerialTransport SerialTransport
		{
			get
			{
				var transport = Transport as ModbusSerialTransport;
				if (transport == null)
					throw new ObjectDisposedException("SerialTransport");

				return transport;
			}
		}

		/// <summary>
		/// Modbus ASCII slave factory method.
		/// </summary>
		public static ModbusSerialSlave CreateAscii(byte unitId, SerialPort serialPort)
		{
			if (serialPort == null)
				throw new ArgumentNullException("serialPort");
			
			return CreateAscii(unitId, new SerialPortAdapter(serialPort));
		}

		/// <summary>
		/// Modbus ASCII slave factory method.
		/// </summary>
		public static ModbusSerialSlave CreateAscii(byte unitId, IStreamResource streamResource)
		{
			if (streamResource == null)
				throw new ArgumentNullException("streamResource");
			
			return new ModbusSerialSlave(unitId, new ModbusAsciiTransport(streamResource));
		}

		/// <summary>
		/// Modbus RTU slave factory method.
		/// </summary>
		public static ModbusSerialSlave CreateRtu(byte unitId, SerialPort serialPort)
		{
			if (serialPort == null)
				throw new ArgumentNullException("serialPort");

            _sp = serialPort;
			return CreateRtu(unitId, new SerialPortAdapter(serialPort));
		}

		/// <summary>
		/// Modbus RTU slave factory method.
		/// </summary>
		public static ModbusSerialSlave CreateRtu(byte unitId, IStreamResource streamResource)
		{
			if (streamResource == null)
				throw new ArgumentNullException("streamResource");

			return new ModbusSerialSlave(unitId, new ModbusRtuTransport(streamResource));
		}
        /// <summary>
        /// Modbus RTU slave Stop.
        /// </summary>
        public override void Stop()
        {
            //SerialTransport.
            //###???
            if (_sp != null)
            {
                _sp.Close();
                _sp.Dispose();
            }
            _sp = null;
        }
		/// <summary>
		/// Start slave listening for requests.
		/// </summary>
		public override void Listen()
		{
			while (true)
			{
				try
				{
					try
					{
						// read request and build message
						byte[] frame = SerialTransport.ReadRequest();
						IModbusMessage request = ModbusMessageFactory.CreateModbusRequest(frame);

						if (SerialTransport.CheckFrame && !SerialTransport.ChecksumsMatch(request, frame))
						{
                            string errorMessage = String.Format(CultureInfo.InvariantCulture, "Checksums failed to match {0} != {1}", StringUtility.Join(", ", request.MessageFrame), StringUtility.Join(", ", frame));
							//_logger.Error(errorMessage);
                            if (LogUtility.IsEnable) LogUtility.Debug(errorMessage, "Listen");
							throw new IOException(errorMessage);
						}

						// only service requests addressed to this particular slave
						if (request.SlaveAddress != UnitId)
						{
							//_logger.DebugFormat("NModbus Slave {0} ignoring request intended for NModbus Slave {1}", UnitId, request.SlaveAddress);
                            if (LogUtility.IsEnable) LogUtility.Debug(string.Format("NModbus Slave {0} ignoring request intended for NModbus Slave {1}", UnitId, request.SlaveAddress), "Listen");
							continue;
						}

						// perform action
						IModbusMessage response = ApplyRequest(request);

						// write response
						SerialTransport.Write(response);
					}
					catch (IOException ioe)
					{
						//_logger.ErrorFormat("IO Exception encountered while listening for requests - {0}", ioe.Message);
                        if (LogUtility.IsEnable) LogUtility.Debug(string.Format("IO Exception encountered while listening for requests - {0}", ioe.Message), "Listen");
						SerialTransport.DiscardInBuffer();
					}
					catch (TimeoutException te)
					{
						//_logger.ErrorFormat("Timeout Exception encountered while listening for requests - {0}", te.Message);
                        if (LogUtility.IsEnable) LogUtility.Debug(string.Format("Timeout Exception encountered while listening for requests - {0}", te.Message), "Listen");
						SerialTransport.DiscardInBuffer();
					}

					// TODO better exception handling here, missing FormatException, NotImplemented...
				}
				catch (InvalidOperationException)
				{
					// when the underlying transport is disposed
					break;
				}
			}
		}
	}
}
