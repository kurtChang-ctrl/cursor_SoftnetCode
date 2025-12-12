using Base.Enums;
using Base.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using ZXing;
using ZXing.Common;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using System.Net.Http;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace Base.Services
{
    //global function
    #pragma warning disable CA2211 // 非常數欄位不應可見
    public static class _Fun
    {
        public static bool Is_Thread_For_Test = false;//for測試RUNTimeServer code用, 目前無作用
        public static bool Is_Thread_ForceClose = false;//網頁強制關閉後台Thread

        public static long _a01 = 0;
        public static long _a02 = 0;
        public static long _a03 = 0;
        public static long _a04 = 0;
        public static long _a05 = 0;
        public static long _a06 = 0;
        public static long _a07 = 0;
        public static long _a08 = 0;
        public static long _a09 = 0;
        public static long _a10 = 0;
        public static long _a11 = 0;
        public static long _a12 = 0;
        public static long _a13 = 0;
        public static long _a14 = 0;
        public static long _a15 = 0;
        public static long _a16 = 0;

        //監看Thead狀態  0=SfcTimerloopUpdateTagValue_Tick 1=SfcTimerloopautoRUN_DOC_Tick 2=SfcTimerloopautoRUN_Json_Tick 3=SfcTimerloopthread_Tick 4=DeviceConnectCheck_Tick
        public static bool[] Is_RUNTimeServer_Thread_State = new bool[] { false, false, false, false, false };

        //public static object Lock_Simulation_Flag = new object();
        public static bool Is_SNWebSocketService_OK = false;
        public static string RUNProcessSYSName = "";
        #region constant
        //hidden input for CRSF
        //public const string HideKey = "_hideKey";

        //system session name
        public const string Rows = "_rows";         //rows fid for CrudEdit
        public const string Childs = "_childs";     //childs fid for CrudEdit

        public const string BaseUser = "_BaseUser";         //base user info
        public const string ProgAuthStrs = "_ProgAuthStrs"; //program autu string list

        //c# datetime format, when js send to c#, will match to _fun.MmDtFmt
        public const string CsDtFmt = "yyyy/MM/dd HH:mm:ss";

        //carrier
        public const string TextCarrier = "\r\n";     //for string
        public const string HtmlCarrier = "<br>";     //for html

        //crud update/view for AuthType=Data only in xxxEdit.cs
        public const string UserFid = "_userId";
        public const string DeptFid = "_deptId";

        //session timeout, map to _BR.js
        public const string TimeOutFid = "TimeOut";

        //indicate error
        public const string PreError = "E:";
        public const string PreBrError = "B:";  //_BR code error
        //public const string PreSystemError = "S:";

        //default view cols for layout(row div, label=2, input=3)(horizontal) 
        public static List<int> DefHoriCols = new() { 2, 3 };

        //directory tail seperator
        public static char DirSep = Path.DirectorySeparatorChar;

        //class name for hide element in RWD phone
        public static string HideRwd = "xg-hide-rwd";
        #endregion

        #region variables which PG can change
        //max export rows count
        public static int MaxExportCount = 3000;

        //crud read for AuthType=Data only in xxxRead.cs
        public static string WhereUserFid = "u.Id='{0}'";
        public static string WhereDeptFid = "u.DeptId='{0}'";

        public static string SystemError = "System Error, Please Contact Administrator.";
        #endregion

        #region input parameters
        //is devironment or not
        public static bool IsDev;

        //private static ServiceContainer _DI  將app.ApplicationServices存下來;
        public static IServiceProvider DiBox;

        //database type
        public static DbTypeEnum DbType;

        //program auth type
        public static AuthTypeEnum AuthType;
        #endregion

        #region base varibles
        //ap physical path, has right slash
        public static string DirRoot = _Str.GetLeft(AppDomain.CurrentDomain.BaseDirectory, "bin" + DirSep);

        //temp folder
        public static string DirTemp = DirRoot + "_temp" + DirSep;
        #endregion

        #region Db variables
        //datetime format for read/write db
        public static string DbDtFmt;
        public static string DbDateFmt;

        //for read page rows
        public static string ReadPageSql;

        //for delete rows
        public static string DeleteRowsSql;
        #endregion

        #region others variables
        //from config file  存放在appsettings.json的設定
        public static ConfigDto Config;

        public static SmtpDto Smtp = default;
        #endregion

        #region 送電子標籤flag判斷
        public static bool Is_Tag_Connect = false;
        public static bool Has_Tag_httpClient = false;
        public static object Lock_Send_macID = new object();
        public static Dictionary<string, string> SendTAGDATA = new Dictionary<string, string>();
        public static void Tag_Write(DBADO db,string macID,string actionType, string json)
        {
            lock (Lock_Send_macID)
            {
                DateTime now = DateTime.Now;
                try
                {
                    if (SendTAGDATA.ContainsKey(macID))
                    { SendTAGDATA[macID] = json; }
                    else { SendTAGDATA.Add(macID, json); }
                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ActionType],[JSON]) VALUES ('{_Str.NewId('L')}','{macID}','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','{actionType}','{json}')");
                }
                catch (Exception ex)
                {
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"寫_sendTAGDATA標籤資料失敗 日期:{now.ToString("yyyy/MM/dd HH:mm:ss.fff")} Exception:{ex.Message} {ex.StackTrace}", true);
                }
            }
        }
        #endregion

        public static MailMessage mms = new MailMessage();
        private static object Lock_Mail_Send = new object();
        public static void Mail_Send(string[] MailTos, string[] Ccs, string MailSub, string MailBody, string[] filePaths, bool deleteFileAttachment)
        {

            if (_Fun.Config.AdministratorEmail.Trim() == "" || _Fun.Config.MailSmtpServer.Trim() == "" || MailTos == null || MailTos.Length <= 0) { return; }
            lock (Lock_Mail_Send)
            {
                try
                {
                    //建立MailMessage物件
                    //MailMessage mms = new MailMessage
                    //{
                    //    //指定一位寄信人MailAddress
                    //    From = new MailAddress(_Fun.Config.AdministratorEmail),
                    //    //信件主旨
                    //    Subject = MailSub
                    //};
                    //信件內容
                    mms.From = new MailAddress(_Fun.Config.AdministratorEmail);
                    if (_Fun.Config.MailSubjectIsSame)
                    {
                        string time = DateTime.Now.ToString("MM/dd HH:mm:ss");
                        MailSub = $"{time} {MailSub}";
                    }
                    mms.Subject = MailSub;
                    //MailBody = MailBody.Replace("\n", "<br>").Replace("\r", "");
                    mms.Body = MailBody;
                    //信件內容 是否採用Html格式
                    mms.IsBodyHtml = true;
                    mms.SubjectEncoding = Encoding.UTF8;
                    if (MailTos != null)//防呆
                    {
                        for (int i = 0; i < MailTos.Length; i++)
                        {
                            //加入信件的收信人(們)address
                            if (!string.IsNullOrEmpty(MailTos[i].Trim()))
                            {
                                mms.To.Add(new MailAddress(MailTos[i].Trim()));
                            }

                        }
                    }//End if (MailTos !=null)//防呆
                    if (mms.To.Count <= 0) { return; }
                    if (Ccs != null && Ccs.Length>0) //防呆
                    {
                        for (int i = 0; i < Ccs.Length; i++)
                        {
                            if (!string.IsNullOrEmpty(Ccs[i].Trim()))
                            {
                                //加入信件的副本(們)address
                                mms.CC.Add(new MailAddress(Ccs[i].Trim()));
                            }

                        }
                    }//End if (Ccs!=null) //防呆
                    if (filePaths != null)//防呆
                    {//有夾帶檔案
                        for (int i = 0; i < filePaths.Length; i++)
                        {
                            if (!string.IsNullOrEmpty(filePaths[i].Trim()))
                            {
                                Attachment file = new Attachment(filePaths[i].Trim());
                                //加入信件的夾帶檔案
                                mms.Attachments.Add(file);
                            }
                        }
                    }
                    string smtpServer = _Fun.Config.MailSmtpServer;
                    int smtpPort = _Fun.Config.MailSmtpPort;
                    string mailAccount = _Fun.Config.MailCredentialsAccount;
                    string mailPwd = _Fun.Config.MailCredentialsPWD;
                    using (SmtpClient client = new SmtpClient(smtpServer, smtpPort))//或公司、客戶的smtp_server
                    {
                        if (!string.IsNullOrEmpty(mailAccount) && !string.IsNullOrEmpty(mailPwd))
                        {
                            client.DeliveryMethod = SmtpDeliveryMethod.Network;
                            client.EnableSsl = true;
                            client.Credentials = new NetworkCredential(mailAccount, mailPwd);
                            client.EnableSsl = true;
                        }
                        //await client.SendMailAsync(mms);
                        client.Send(mms);//send
                    }//end using 
                     //釋放每個附件，才不會Lock住
                    if (mms.Attachments != null && mms.Attachments.Count > 0)
                    {
                        for (int i = 0; i < mms.Attachments.Count; i++)
                        {
                            mms.Attachments[i].Dispose();
                        }
                    }
                    //mms.Dispose();
                    //mms = null;

                    #region 要刪除附檔
                    if (deleteFileAttachment && filePaths != null && filePaths.Length > 0)
                    {

                        foreach (string filePath in filePaths)
                        {
                            //###???File.Delete(filePath.Trim());
                        }

                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    System.Threading.Tasks.Task task = _Log.ErrorAsync($"SNWebSocketService.cs Mail_Send Exception: {ex.Message} {ex.StackTrace}", false);
                }
            }
        }

        public static byte[] Craete2BarCode(string stationNO,int width,int height,string type)
        {
            BarcodeWriterPixelData writer = new BarcodeWriterPixelData()
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new EncodingOptions
                {
                    Height = height,
                    Width = width,
                    PureBarcode = false, // this should indicate that the text should be displayed, in theory. Makes no difference, though.
                    Margin = 10
                }
            };
            string url = "";// $"http://{Config.LocalWebURL}/Home/Index";
            if (type == "1")
            {
                url = $"http://{Config.LocalWebURL}/LabelStroe/Index/{stationNO}";
            }
            else if (type == "2")
            {
                url = $"http://{Config.LocalWebURL}/LabelStoreSpacesNO/StroeFUN/{stationNO}";
            }
            else
            {
                if (_Fun.Config.RUNMode == '1')
                {
                    url = $"http://{Config.LocalWebURL}/LabelProject/Index/{stationNO}";
                }
                else
                {
                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                    {
                        DataRow tmp_dt = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}' and Station_Type='1'");
                        if (tmp_dt != null)
                        {
                            url = $"http://{Config.LocalWebURL}/LabelWork/Index/{stationNO};0;;0";
                            /*
                            tmp_dt = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[LabelStateINFO] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}'");
                            if (tmp_dt != null)
                            {
                                if (tmp_dt["SimulationId"].ToString() != "")
                                {
                                    url = $"http://{Config.LocalWebURL}/LabelWork/Index/{stationNO};2;{tmp_dt["SimulationId"].ToString()};{tmp_dt["IndexSN"].ToString()}";
                                }
                                else
                                {
                                    if (tmp_dt["OrderNO"].ToString() == "")
                                    { url = $"http://{Config.LocalWebURL}/LabelWork/Index/{stationNO};0;;0"; }
                                    else if (tmp_dt["OrderNO"].ToString() != "" && tmp_dt["IndexSN"].ToString().Trim() != "0")
                                    {
                                        url = $"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{stationNO};1;{tmp_dt["OrderNO"].ToString()};{tmp_dt["IndexSN"].ToString()}";
                                    }
                                }
                            }
                            else
                            {
                                //"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{stationNO};0;;0\"
                                tmp_dt = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[Manufacture] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}'");
                                if (tmp_dt != null)
                                {
                                    if (tmp_dt["SimulationId"].ToString() != "")
                                    {
                                        url = $"http://{Config.LocalWebURL}/LabelWork/Index/{stationNO};2;{tmp_dt["SimulationId"].ToString()};{tmp_dt["IndexSN"].ToString()}";
                                        string _s = "";
                                    }
                                    else if (tmp_dt["OrderNO"].ToString() == "")
                                    { url = $"http://{Config.LocalWebURL}/LabelWork/Index/{stationNO};0;;0"; }
                                    else if (tmp_dt["OrderNO"].ToString() != "" && tmp_dt["IndexSN"].ToString().Trim() != "0")
                                    {
                                        url = $"http://{_Fun.Config.LocalWebURL}/LabelWork/Index/{stationNO};1;{tmp_dt["OrderNO"].ToString()};{tmp_dt["IndexSN"].ToString()}";
                                    }

                                }
                            }
                            */
                        }
                    }
                }

            }
            var pixelData = writer.Write(url);
            using (var bitmap = new Bitmap(pixelData.Width, pixelData.Height, System.Drawing.Imaging.PixelFormat.Format32bppRgb))
            {
                using (var ms = new System.IO.MemoryStream())
                {
                    var bitmapData = bitmap.LockBits(new Rectangle(0, 0, pixelData.Width, pixelData.Height), System.Drawing.Imaging.ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppRgb);
                    try
                    {
                        System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bitmapData.Scan0, pixelData.Pixels.Length);
                    }
                    finally
                    {
                        bitmap.UnlockBits(bitmapData);
                    }

                    bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                    return ms.ToArray();
                    //return File(ms.ToArray(), "image/jpeg");
                    //bitmap.Save(HttpContext.Response.Body, System.Drawing.Imaging.ImageFormat.Png);
                    //return bitmap;
                }
            }
        }

        public static DataRow GetAvgCTWTandTotalOutput(DBADO db, bool isTrace, string wo, string stationNO, string indexSN)
        {
            string sql = "";
            string tmp_indexSN = "";
            if (indexSN.Trim() != "") { tmp_indexSN = $" AND IndexSN={indexSN}"; }

            if (isTrace)
            {

                sql = $@"SELECT A.AvgCT,A.AvgWT,B.TotalOutput FROM
                        (
                            SELECT 
                                OrderNO,CAST(ROUND(AVG(CycleTime/1.0),2) AS NUMERIC(10,2)) AvgCT,
                                IIF(COUNT(1)-1>0,CAST(ROUND(SUM(WaitTime/1.0)/(COUNT(1)-1),2) AS NUMERIC(10,2)),0) AvgWT
                            FROM SoftNetLogDB.[dbo].SFC_StationDetail  
                            WHERE IsDel=0 AND StationNO='{stationNO}' AND OrderNO='{wo}' {tmp_indexSN} AND (OutputFlag = '1' OR  FailFlag = '1') 
	                        GROUP BY OrderNO
	                    ) A 
	                    LEFT JOIN
	                    (
	                        SELECT 
                                OrderNO,COUNT(1) TotalOutput 
                            FROM SoftNetLogDB.[dbo].SFC_StationDetail  
                            WHERE IsDel=0 AND StationNO='{stationNO}' {tmp_indexSN} AND OrderNO='{wo}' AND OutputFlag = '1' 
	                        GROUP BY OrderNO
	                    ) B ON A.OrderNO=B.OrderNO";
            }
            else
            {
                sql = $@"SELECT
                            IIF(SUM(A.CT)/SUM(A.InQty) IS NULL,0,CAST(ROUND((SUM(A.CT)/1.0)/SUM(A.InQty),2) AS NUMERIC(10,2))) 'AvgCT',
							IIF((SUM(A.InQty)-1)>0,CAST(ROUND((SUM(A.WT)/1.0)/(SUM(A.InQty)-1),2) AS NUMERIC(10,2)),0) 'AvgWT',
							IIF(SUM(A.OutQty) IS NULL,0,SUM(A.OutQty)) 'TotalOutput' 
                        FROM
                        (
                            SELECT SUM(B.CT) 'CT',SUM(B.WT) 'WT',SUM(B.InQty) 'InQty',SUM(B.OutQty) 'OutQty' FROM
                            (
                                SELECT 
                                    CycleTime * (SUM(ProductFinishedQty) + SUM(ProductFailedQty)) 'CT', 
			                        WaitTime  * (SUM(ProductFinishedQty) + SUM(ProductFailedQty)-1) 'WT',
			                        SUM(ProductFinishedQty) + SUM(ProductFailedQty) 'InQty', 
			                        SUM(ProductFinishedQty) 'OutQty'
                                FROM SoftNetLogDB.[dbo].SFC_StationDetail
                                WHERE IsDel = 0 and StationNO = '{stationNO}' {tmp_indexSN} and OrderNO = '{wo}' and (OutputFlag = '1' OR  FailFlag = '1')
                                GROUP BY CycleTime,WaitTime
                            ) B
						)
						A GROUP BY OutQty";
            }
            return db.DB_GetFirstDataByDataRow(sql);
        }

        /*
        //constructor
        static _Fun()
        {
        }
        */

        /// <summary>
        /// initial db environment for Ap with db function !!
        /// </summary>
        /// <param name="isDev">is devironment or not</param>
        /// <param name="diBox"></param>
        /// <param name="dbType"></param>
        /// <param name="authType"></param>
        /// <returns>error msg if any</returns>
        public static string Init(bool isDev, IServiceProvider diBox, DbTypeEnum dbType = DbTypeEnum.MSSql, 
            AuthTypeEnum authType = AuthTypeEnum.None)
        {
            //set instance variables
            IsDev = isDev;
            DiBox = diBox;
            DbType = dbType;
            AuthType = authType;

            Config.HtmlImageUrl = _Str.AddSlash(Config.HtmlImageUrl); //初始內容為 "/"

            #region set smtp
            var smtp = Config.Smtp;
            if (smtp != "")
            {
                try
                {
                    var cols = smtp.Split(',');
                    Smtp = new SmtpDto()
                    {
                        Host = cols[0],
                        Port = int.Parse(cols[1]),
                        Ssl = bool.Parse(cols[2]),
                        Id = cols[3],
                        Pwd = cols[4],
                        FromEmail = cols[5],
                        FromName = cols[6]
                    };
                }
                catch
                {
                    return "config Smtp is wrong(7 cols): Host,Port,Ssl,Id,Pwd,FromEmail,FromName";
                }
            }
            #endregion

            #region set DB variables
            //0:select, 1:order by, 2:start row(base 0), 3:rows count
            switch (dbType)
            {
                case DbTypeEnum.MSSql:
                    DbDtFmt = "yyyy-MM-dd HH:mm:ss";
                    DbDateFmt = "yyyy-MM-dd";

                    //for sql 2012, more easy
                    //note: offset/fetch not sql argument
                    ReadPageSql = @"
select {0} {1}
offset {2} rows fetch next {3} rows only
";
                    DeleteRowsSql = "delete {0} where {1} in ({2})";    
                    break;

                case DbTypeEnum.MySql:
                    #region
                    DbDtFmt = "YYYY-MM-DD HH:mm:SS";
                    DbDateFmt = "YYYY-MM-DD";

                    ReadPageSql = @"
select {0} {1}
limit {2},{3}
";
                    //TODO: set delete sql for MySql
                    //DeleteRowsSql = 
                    #endregion
                    break;

                case DbTypeEnum.Oracle:
                    #region
                    DbDtFmt = "YYYY/MM/DD HH24:MI:SS";
                    DbDateFmt = "YYYY/MM/DD";

                    //for Oracle 12c after(same as mssql)
                    ReadPageSql = @"
Select {0} {1}
Offset {2} Rows Fetch Next {3} Rows Only
";
                    //TODO: set delete sql for Oracle
                    //DeleteRowsSql = 
                    #endregion
                    break;
            }
            #endregion

            //case of ok
            return "";
        }

        //get current userId
        public static string Dir(string folder, bool tailSep = true)
        {
            return _Fun.DirRoot + folder + (tailSep ? DirSep : "");
        }

        //get current userId
        public static string UserId()
        {
            if (GetBaseUser() == null) { return ""; }
            return GetBaseUser().UserId;
        }

        public static string DeptId()
        {
            if (GetBaseUser() == null) { return ""; }
            return GetBaseUser().DeptId;
        }

        //check is AuthType=Data
        public static bool IsAuthTypeData()
        {
            return (AuthType == AuthTypeEnum.Data);
        }

        /*
        public static string GetBrError(string msg)
        {
            return IsError(msg)
                ? msg.Substring(PreError.Length)
                : "";
        }

        public static bool IsError(string msg)
        {
            var len = PreError.Length;
            return !_Str.IsEmpty(msg) &&
                msg.Length >= len &&
                msg.Substring(0, len) == PreError;
        }
        */

        /// <summary>
        /// get base user info for base component
        /// </summary>
        /// <returns>BaseUserInfoDto</returns>
        public static BaseUserDto GetBaseUser()
        {
            var service = (IBaseUserService)DiBox.GetService(typeof(IBaseUserService));
            return service.GetData();
        }

        /// <summary>
        /// check and open db
        /// </summary>
        /// <param name="db"></param>
        /// <param name="hasDb"></param>
        /// <param name="dbStr"></param>
        public static void CheckOpenDb(ref Db db, ref bool hasDb, string dbStr = "")
        {
            hasDb = (db != null);
            if (!hasDb)
            { db = new Db(dbStr); }
        }

        /// <summary>
        /// check and close db
        /// </summary>
        /// <param name="db"></param>
        /// <param name="hasDb"></param>
        public static async Task CheckCloseDb(Db db, bool hasDb)
        {
            if (!hasDb)
                await db.DisposeAsync();
        }

        public static void Except(string error = "")
        {
            throw new Exception(_Str.EmptyToValue(error, SystemError));
        }

        public static string GetHtmlImageUrl(string subPath)
        {
            return $"{Config.HtmlImageUrl}{subPath}";
        }

    } //class
    #pragma warning restore CA2211 // 非常數欄位不應可見
}
