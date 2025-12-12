using System;
using System.IO;
using System.Text;
using System.Threading;

namespace Modbus.Utility
{
    // 格式大致為
    // HH:mm:ss.fff: LOG_Message
    class LogUtility
    {
        // static locker
        private static object _LOG_Lock = new object();
        private const string LOG_FOLDER = @".\HMILOG";
        private const string LOG_SUBFOLDER = "Modbus"; // 需手動建立
        private const string LOG_PREFIX = "modbus_";
        private const string LOG_EXTENSION = ".log";

        // enable or disable
        private static bool LogEnable = false; // default disable
        // file size
        private static int LOG_FILE_SIZE = 102400000;// 100mb
        // log message
        private static StringBuilder LogMessage = new StringBuilder();

        private static bool _IsActive = false;
        private static Thread ActiveThread = null;

        static LogUtility()
        {
            if (LogEnable == false) return;
            if (ActiveThread != null) return;

            _IsActive = true;
            ActiveThread = new Thread(LogThread);
            ActiveThread.IsBackground = true;
            ActiveThread.Start();
        }
        ~LogUtility()
        {
            _IsActive = false;
        }

        private static void LogThread() // thread to write file from string builder
        {
            LogInit();
            while (_IsActive)
            {
                int s = (new Random()).Next(1000, 3000);
                SpinWait.SpinUntil(() => _IsActive == false, s);
                if (_IsActive == false) break;
                LogWrite();
            }
            LogFree();
            ActiveThread = null;
        }

        private static string LogGetFileName() // 取得要寫入的檔案名稱
        {
            // Append Date
            string logfile = null;
            try
            {
                // check subfolder
                string subfolder = System.IO.Path.Combine(LOG_FOLDER, LOG_SUBFOLDER);
                if (Directory.Exists(subfolder) == false) return String.Empty; // if subfolder is not exist, no need to write

                logfile = System.IO.Path.Combine(LOG_FOLDER, LOG_SUBFOLDER, LOG_PREFIX + DateTime.Now.ToString("yyyyMMdd") + LOG_EXTENSION);
                if (File.Exists(logfile))
                {
                    FileInfo fi = new FileInfo(logfile);
                    if (fi == null) return logfile;

                    if (fi.Length > LOG_FILE_SIZE)
                    {
                        try
                        {
                            string backupfile = System.IO.Path.Combine(LOG_FOLDER, LOG_SUBFOLDER, LOG_PREFIX + DateTime.Now.ToString("yyyyMMdd_HHmmss") + LOG_EXTENSION);
                            if (File.Exists(logfile) == true && String.IsNullOrEmpty(backupfile) == false && File.Exists(backupfile) == false)
                            { // file move
                                File.Move(logfile, backupfile);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return logfile;
        }
        // mesg      ref to XXXString
        // append    0 is overwrite, 1 is append
        private static void LogWriteFile(ref StringBuilder mesg, int append) // 0 is overwrite, 覆寫內容 // 1 is append, 附加內容
        {
            if (LogInit() == false) return;
            try
            {
                lock (_LOG_Lock)
                {
                    string filep = LogGetFileName();
                    if (String.IsNullOrEmpty(filep)) return;

                    StreamWriter sw = null;
                    try
                    {
                        if (append == 1) // 1 is append
                            sw = new StreamWriter(filep, true); // true is append
                        else
                            sw = new StreamWriter(filep, false); // false is overwrite

                        if (sw != null)
                        {
                            if (mesg != null) sw.Write(mesg.ToString());
                            sw.Flush();
                            sw.Close();
                        }
                    }
                    catch (Exception)// e)
                    {
                        //Debug.WriteLine("1LogWriteFile: " + e);
                    }
                    finally
                    {
                        if (sw != null) sw.Dispose();
                    }
                }
            }
            catch (Exception)// e)
            {
                //Debug.WriteLine("2LogWriteFile: " + e);
            }
        }
        private static void log_append(ref StringBuilder sb, params string[] mesg) // append message to string builder
        {
            lock (_LOG_Lock)
            {
                //if (LogEnable == false) return;

                if (sb == null) sb = new StringBuilder();
                if (sb != null)
                {
                    if (mesg == null || mesg[0] == null) // clear 
                    {
                        sb.Clear();
                    }
                    else // append string
                    {
                        foreach (var s in mesg)
                        {
                            if (s != null) sb.Append(s);
                        }
                    }
                }
            }
        }
        private static void log_write(ref StringBuilder sb, int append) // overwrite file or append file
        {
            lock (_LOG_Lock)
            {
                //if (LogEnable == false) return;

                if (sb == null) sb = new StringBuilder();
                if (sb.Length <= 0) return; // no need write

                LogWriteFile(ref sb, append);
                sb.Clear();
            }
        }

        private static void LogFree()
        {
            LogWrite();
            if (LogMessage != null)
            {
                LogMessage.Clear();
            }
        }
        private static bool LogInit()
        {
            //if (LogEnable == true)
            {
                try
                {
                    if (Directory.Exists(LOG_FOLDER) == false)
                    {
                        Directory.CreateDirectory(LOG_FOLDER);
                    }
                    if (Directory.Exists(LOG_FOLDER) == true)
                    {
                        string s = Path.Combine(LOG_FOLDER, LOG_SUBFOLDER);
                        if (Directory.Exists(s) == true) // don't create Modbus sub-folder (for debug)(手動建立)
                        {
                            return true;
                        }
                    }
                }
                catch { }
            }
            return false;
        }
        private static void LogWrite()
        {
            if (LogEnable == false) return;

            log_write(ref LogMessage, 1);
        }

        public static bool IsEnable
        {
            get { return LogEnable; }
            set { LogEnable = value; }
        }
        // log
        public static void Debug(string msg, string title = "Modbus")
        {
            if (LogEnable == false) return;

            if (msg == null) return;
            log_append(ref LogMessage, string.Format("{0}: [{1}] {2}{3}", DateTime.Now.ToString("HH:mm:ss.fff"), title, msg, Environment.NewLine));
        }
        public static void Debug2(string msg, string title = "Modbus")
        {
            if (msg == null) return;
            log_append(ref LogMessage, string.Format("{0}: [{1}] {2}{3}", DateTime.Now.ToString("HH:mm:ss.fff"), title, msg, Environment.NewLine));
            log_write(ref LogMessage, 1);
        }
    }
}

