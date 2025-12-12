using static StackExchange.Redis.Role;

namespace Base.Models
{
    /// <summary>
    /// get from appSettings.json Config section
    /// </summary>
    public class ConfigDto
    {
        //constructor
        public ConfigDto()
        {
            SystemName = "MIS System";
            Locale = "zh-TW";
            ServerId = "01";//客戶碼
            SlowSql = 1000;
            LogDebug = false;
            LogSql = false;
            RootEmail = "";
            TesterEmail = "";
            UploadFileMax = 5;
            //CacheSecond = 3600;
            SSL = false;
            Smtp = "";
            HtmlImageUrl = "";
            Redis = "";
            AutoDOC_LoopTime = 30000;
            OutPackStationName = "委外加工";
            IsOutPackStationStore = false;
            IsAutoDispatch = true;
            IsAutoDispatch_IsAutoUpdate_WO = true;
            DefaultStoreNO = "";
            DefaultFactoryName = "";
            DefaultLineName = "";
            APS_Simulation_ErrorData_Clear_Day = 0;
            Default_Simulation_AGE01 = false;
            Default_Simulation_AGE02 = false;
            Default_Simulation_AGE03 = false;
            Default_WorkingPaper_AGE01 = true;
            Default_WorkingPaper_AGE02 = 15;
            Default_WorkingPaper_AGE03 = "";
            Default_WorkingPaper_AGE03_StoreSpacesNO = "";
            Default_WorkingPaper_NOT_SafeQTY_Order = false;
            Default_WorkingPaper_NOT_SafeQTY_DOC1 = false;
            Default_SimulationDelay = 900;
            AdministratorEmail = "kurt@softnet.tw";
            SystemEmail = "kurt@softnet.tw";
            MailSmtpServer = "smtp.gmail.com";
            MailSmtpPort = 587;
            MailCredentialsAccount = "";
            MailCredentialsPWD = "";
            MailSubjectIsSame = false;
            ElectronicTagsURL = "127.0.0.1";
            MasterServiceIP = "127.0.0.1";
            LocalWebURL = "127.0.0.1";
            WesocketPort = 8089;
            WesocketURL = "ws://localhost:8089/echo";
            RUNMode = '2';
            RunTimeServerLoopTime = 60000;
            AdminKey03 = 10000;//分析效能筆數
            AdminKey14 = false;
            DefaultCalendarName = "";
            Default_EStore_ControlURL = "";
            Default_EStore_MachineToken = "";
            APS_CT_Custom_Rate = 100;
            SendMonitorMail00 = "";
            SendMonitorMail01 = "";
            SendMonitorMail02 = "";
            SendMonitorMail03 = "";
            SendMonitorMail04 = "";
            SendMonitorMail05 = "";
            SendMonitorMail06 = "";
            SendMonitorMail07 = "";
            SendMonitorMail08 = "";
            SendMonitorMail09 = "";
            SendMonitorMail10 = "";
            SendMonitorMail11 = "";
            SendMonitorMail12 = "";
            SendMonitorMail13 = "";
            SendMonitorMail14 = "";
            SendMonitorMail15 = "";
            SendMonitorMail16 = "";
            SendMonitorMail17 = "";
            SendMonitorMail18 = "";
            SendMonitorMail19 = "";
            SendMonitorMail20 = "";
            SendMonitorMail21 = "";
            SendMonitorMail22 = "";
            SendMonitorMail23 = "";

        }
        public string SendMonitorMail00 { get; set; }
        public string SendMonitorMail01 { get; set; }
        public string SendMonitorMail02 { get; set; }
        public string SendMonitorMail03 { get; set; }
        public string SendMonitorMail04 { get; set; }
        public string SendMonitorMail05 { get; set; }
        public string SendMonitorMail06 { get; set; }
        public string SendMonitorMail07 { get; set; }
        public string SendMonitorMail08 { get; set; }
        public string SendMonitorMail09 { get; set; }
        public string SendMonitorMail10 { get; set; }
        public string SendMonitorMail11 { get; set; }
        public string SendMonitorMail12 { get; set; }
        public string SendMonitorMail13 { get; set; }
        public string SendMonitorMail14 { get; set; }
        public string SendMonitorMail15 { get; set; }
        public string SendMonitorMail16 { get; set; }
        public string SendMonitorMail17 { get; set; }
        public string SendMonitorMail18 { get; set; }
        public string SendMonitorMail19 { get; set; }
        public string SendMonitorMail20 { get; set; }
        public string SendMonitorMail21 { get; set; }
        public string SendMonitorMail22 { get; set; }
        public string SendMonitorMail23 { get; set; }



        public int APS_CT_Custom_Rate { get; set; }
        public string Default_EStore_ControlURL { get; set; }
        public string Default_EStore_MachineToken { get; set; }
        public int APS_Simulation_ErrorData_Clear_Day { get; set; }
        public bool Default_WorkingPaper_NOT_SafeQTY_Order { get; set; }
        public bool MailSubjectIsSame { get; set; }
        public bool Default_WorkingPaper_NOT_SafeQTY_DOC1 { get; set; }
        public bool Default_Simulation_AGE01 { get; set; }
        public bool Default_Simulation_AGE02 { get; set; }
        public bool Default_Simulation_AGE03 { get; set; }
        public bool Default_WorkingPaper_AGE01 { get; set; }
        public int Default_WorkingPaper_AGE02 { get; set; }
        public string Default_WorkingPaper_AGE03 { get; set; }
        public string Default_WorkingPaper_AGE03_StoreSpacesNO { get; set; }
        public int Default_SimulationDelay { get; set; }
        public char RUNMode { get; set; }
        public string LocalWebURL { get; set; }
        public int WesocketPort { get; set; }
        public string WesocketURL { get; set; }
        public string DefaultCalendarName { get; set; }
        public string DefaultFactoryName { get; set; }
        public string DefaultLineName { get; set; }
        public string DefaultStoreNO { get; set; }
        public string OutPackStationName { get; set; }
        public bool IsOutPackStationStore { get; set; }
        public bool IsAutoDispatch { get; set; }
        public bool IsAutoDispatch_IsAutoUpdate_WO { get; set; }
        public string AdministratorEmail { get; set; }
        public string SystemEmail { get; set; }
        public string MailSmtpServer { get; set; }
        public int MailSmtpPort { get; set; }
        public string MailCredentialsAccount { get; set; }
        public string MailCredentialsPWD { get; set; }
        
        public string MasterServiceIP { get; set; }
        public string ElectronicTagsURL { get; set; }
        public int RunTimeServerLoopTime { get; set; }
        public int AdminKey03 { get; set; }
        public bool AdminKey14 { get; set; }
        public int AutoDOC_LoopTime { get; set; }
        //db connect string
        public string Db { get; set; }

        //system name
        public string SystemName { get; set; }

        //default locale code
        public string Locale { get; set; }

        //server Id for new key
        public string ServerId { get; set; }

        //log error for slow sql(mini secode)
        public int SlowSql { get; set; }

        //log debug
        public bool LogDebug { get; set; }

        //log sql
        public bool LogSql { get; set; }

        //root email address for send error
        public string RootEmail { get; set; }

        //tester email address
        public string TesterEmail { get; set; }

        //upload file max size(MB)
        public int UploadFileMax { get; set; }

        //cache time(second)
        //public int CacheSecond { get; set; }

        //SSL or not
        public bool SSL { get; set; }

        //smtp, format: 0(Host),1(Port),2(Ssl),3(Id),4(Pwd),5(FromEmail),6(FromName) 
        public string Smtp { get; set; }

        //email image path list: Id,Path.., ex: _TopImage, c:/xx/xx.png
        public string EmailImagePairs { get; set; }

        /// <summary>
        /// html image root url for sublime, ex: http://xxx.xx/image, auto add right slash
        /// </summary>
        public string HtmlImageUrl { get; set; }

        /// <summary>
        /// redis server for session, ex: "127.0.0.1:6379,ssl=true,password=xxx,defaultDatabase=x", 
        /// empty for memory cache
        /// </summary>
        public string Redis { get; set; }

    }
}
