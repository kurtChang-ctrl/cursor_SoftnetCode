using Base.Enums;
using System;
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
            APS_DOC3stock_Action_Advance_Day = 0;
            APS_DOC3stock_Action_Delay_Day = 0;
            DOCInvoice_Expiry_SendError = true;
            Material_Short_IsWarning=true;
            ElectronicTagsIsFailSendMail = true;
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
            APS_PlanIdleDays = 0;
            UseSQLVector = false;
            RAG_VectorStoreRetentionDays = 90;
            GeminiApiKey = "";
            GeminiModels = new[] { "gemini-2.5-flash" };
            OpenAIApiKey = "";
            AIReportDDay = DaysOfWeek.None;
            AIReportQuestion = "";
            SCM_IsEnabled = false;
            RAG_IsEnabled = false;
            SCM_WorkingDoWhileTime = 30;
            SCM_MaxRetryCount = 3;
            SCM_MailSubjectPrefix = "[SoftNetSCM]";
            SCM_EnableSignature = false;
            OrderNO_IsAutoClose = false;
            SCM_SignatureKey = "";
            SCM_DataExchange_Mail = "";
            SCM_LocalDecision_Mail = "";
            SCM_ImapServer = "imap.gmail.com";
            SCM_ImapPort = 993;
            WebPageDataDisCount = 20;
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
            SendMonitorMail24 = "";
            SendMonitorMail25 = "";
            SendMonitorMail26 = "";
            SendMonitorMail27 = "";
            SendMonitorMail28 = "";
            SendMonitorMail29 = "";
            SendMonitorMail30 = "";
            SendMonitorMail31 = "";
            SendMonitorMail32 = "";
            SendMonitorMail33 = "";

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
        public string SendMonitorMail24 { get; set; }
        public string SendMonitorMail25 { get; set; }
        public string SendMonitorMail26 { get; set; }
        public string SendMonitorMail27 { get; set; }
        public string SendMonitorMail28 { get; set; }
        public string SendMonitorMail29 { get; set; }
        public string SendMonitorMail30 { get; set; }
        public string SendMonitorMail31 { get; set; }
        public string SendMonitorMail32 { get; set; }
        public string SendMonitorMail33 { get; set; }


        public int APS_CT_Custom_Rate { get; set; }
        /// <summary>
        /// 監看計畫久未執行的天數門檻
        /// </summary>
        public int APS_PlanIdleDays { get; set; }
        /// <summary>是否使用 SQL Server 原生 vector 型別做向量搜尋（目前保守設為 false，之後可改為 true）</summary>
        public bool UseSQLVector { get; set; }
        /// <summary>RAG 向量庫保留天數，超過的資料移至歷史表</summary>
        public int RAG_VectorStoreRetentionDays { get; set; }
        public string Default_EStore_ControlURL { get; set; }
        public string Default_EStore_MachineToken { get; set; }
        public string GeminiApiKey { get; set; }
        /// <summary>Gemini 模型嘗試清單（依序）</summary>
        public string[] GeminiModels { get; set; }
        public string OpenAIApiKey { get; set; }
        public DaysOfWeek AIReportDDay { get; set; }
        public string AIReportQuestion { get; set; }
        /// <summary>是否啟用 SCM Mail 交換流程</summary>
        public bool SCM_IsEnabled { get; set; }
        public bool RAG_IsEnabled { get; set; }
        /// <summary>SCM 輪巡週期（分鐘）</summary>
        public int SCM_WorkingDoWhileTime { get; set; }
        /// <summary>Outbox 寄送最大重試次數</summary>
        public int SCM_MaxRetryCount { get; set; }
        /// <summary>SCM 郵件主旨前綴（例如 [SoftNetSCM]）</summary>
        public string SCM_MailSubjectPrefix { get; set; }
        /// <summary>是否啟用簽章驗證（預留）</summary>
        public bool SCM_EnableSignature { get; set; }
        public bool OrderNO_IsAutoClose { get; set; }
        /// <summary>SCM 簽章密鑰（建議放 user-secrets）</summary>
        public string SCM_SignatureKey { get; set; }
        /// <summary>供應鏈交換共用信箱</summary>
        public string SCM_DataExchange_Mail { get; set; }
        /// <summary>本地人工決策通知信箱</summary>
        public string SCM_LocalDecision_Mail { get; set; }
        /// <summary>收件 IMAP 主機</summary>
        public string SCM_ImapServer { get; set; }
        /// <summary>收件 IMAP Port</summary>
        public int SCM_ImapPort { get; set; }
        /// <summary>網頁表格每頁預設顯示筆數</summary>
        public int WebPageDataDisCount { get; set; }
        public int APS_Simulation_ErrorData_Clear_Day { get; set; }
        public int APS_DOC3stock_Action_Advance_Day { get; set; }
        public int APS_DOC3stock_Action_Delay_Day { get; set; }

        public bool DOCInvoice_Expiry_SendError { get; set; }
        public bool Material_Short_IsWarning { get; set; }
        public bool ElectronicTagsIsFailSendMail { get; set; }
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
