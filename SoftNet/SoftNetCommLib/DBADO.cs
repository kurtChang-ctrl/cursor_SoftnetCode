using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;


namespace SoftNetCommLib
{


    public class DBADO : IDisposable
    {
        //###??? 將來要寫reTry connect
        string ms_connectType = "";
        string ms_connectString = "";
  
        public delegate void ERROR(string sql, string MEG);
        public event ERROR Error;

        public DBADO()
        {

        }
        public DBADO(string connectType, string connectString)
        {
            ms_connectType = connectType;
            ms_connectString = connectString;
        }
        public DBADO(string connectType, string connectString, ref string errMEG)
        {
            errMEG = "";
            ms_connectType = connectType;
            ms_connectString = connectString;
            try
            {
                switch (ms_connectType)
                {
                    case "1"://MSSQL
                        using (SqlConnection mssqlconn = new SqlConnection(ms_connectString)) { mssqlconn.Open(); }
                        break;
                }
            }
            catch (Exception ex)
            {
                
                    Mail_Send("kurt@softnet.tw", ex.Message, $"New資料庫物件時錯誤   來源Connect:{ms_connectString}");
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }


        private bool disposed = false;
        protected virtual void Dispose(bool disposing)
        {
            if (disposed) { return; }
            if (disposing)
            {
                //清理CLR託管資源
                /*
                if (mssqlconn != null)
                {
                    mssqlconn.Close();
                    mssqlconn.Dispose();
                }
                if (oledbconn != null)
                {
                    oledbconn.Close();
                    oledbconn.Dispose();
                }
                if (odbcconn != null)
                {
                    odbcconn.Close();
                    odbcconn.Dispose();
                }
                 */
            }
            //清理非託管資源,寫在下方,如果有的話
            disposed = true;

        }

        public string ConnectType
        {
            get { return ms_connectType; }
            set { ms_connectType = value; }
        }
        public string ConnectString
        {
            get { return ms_connectString; }
            set { ms_connectString = value; }
        }
        public int DB_GetQueryCount(string sql) //取得筆數 sql=SQL語法
        {
            if (sql == "") { return 0; }
            string err = "";
            DataTable drN = fn_GetDataTable(sql, ref err);
            if (drN != null)
            {
                return drN.Rows.Count;
            }
            return 0;
        }
        public int DB_GetQueryCount(string sql, ref string err) //取得筆數 sql=SQL語法
        {
            if (sql == "") { return 0; }
            DataTable drN = fn_GetDataTable(sql, ref err);
            if (drN != null)
            {
                return drN.Rows.Count;
            }
            return 0;
        }
        public JArray DB_Test(string sql, List<object> sqlArgs)//取得資料
        {
            if (sql == "") { return null; }
            string err = "1";
            DataTable drN = fn_GetDataTableByArgs(sql, sqlArgs, ref err);
            if (drN != null)
            {
                return Newtonsoft.Json.Linq.JArray.FromObject(drN);
            }
            return null;
        }
        public JObject DB_TestROW(string sql, List<object> sqlArgs)//取得第一筆資料
        {
            if (sql == "") { return null; }
            string err = "";
            DataTable drN = fn_GetDataTableByArgs(sql, sqlArgs, ref err);
            if (drN != null)
            {
                if (drN.Rows.Count > 0)
                {
                    return (JObject)Newtonsoft.Json.Linq.JArray.FromObject(drN)[0];
                }
            }
            return null;
        }
        public JArray DB_TestROWByArray(string sql, List<object> sqlArgs)//取得第一筆資料
        {
            if (sql == "") { return null; }
            string err = "";
            DataTable drN = fn_GetDataTableByArgs(sql, sqlArgs, ref err);
            if (drN != null)
            {
                if (drN.Rows.Count > 0)
                {
                    return (JArray)Newtonsoft.Json.Linq.JArray.FromObject(drN)[0];
                }
            }
            return null;
        }
        public int DB_TestQueryCount(string sql, List<object> sqlArgs) //取得筆數 sql=SQL語法
        {
            if (sql == "") { return 0; }
            string err = "";
            DataTable drN = fn_GetDataTable(sql, ref err);
            if (drN != null)
            {
                return drN.Rows.Count;
            }
            return 0;
        }

        public DataTable? DB_GetData(string sql)//取得資料
        {
            if (sql == "") { return null; }
            string err = "";
            DataTable drN = fn_GetDataTable(sql, ref err);
            if (drN != null)
            {
                return drN;
            }
            return null;
        }
        public DataTable? DB_GetData(string sql, ref string err)//取得資料
        {
            if (sql == "") { return null; }
            DataTable drN = fn_GetDataTable(sql, ref err);
            if (drN != null)
            {
                return drN;
            }
            return null;
        }
        public string[]? DB_GetFirstDataByStringArray(string sql)//取得第一筆資料
        {
            if (sql == "") { return null; }
            string err = "";
            DataTable drN = fn_GetDataTable(sql, ref err);
            if (drN != null)
            {
                if (drN.Rows.Count > 0)
                {
                    if (drN.Rows[0].ItemArray.Length > 0)
                    {
                        string[] re = new string[drN.Rows[0].ItemArray.Length];
                        int i = 0; int j = 0;
                        for (i = 0; i < drN.Rows[0].ItemArray.Length; i++)
                        {
                            re[i] = drN.Rows[0][i].ToString();
                        }
                        return re;
                    }
                }
            }
            return null;
        }
        public string[]? DB_GetFirstDataByStringArray(string sql, ref string err)//取得第一筆資料
        {
            if (sql == "") { return null; }
            DataTable drN = fn_GetDataTable(sql, ref err);
            if (drN != null)
            {
                if (drN.Rows.Count > 0)
                {
                    if (drN.Rows[0].ItemArray.Length > 0)
                    {
                        string[] re = new string[drN.Rows[0].ItemArray.Length];
                        int i = 0; int j = 0;
                        for (i = 0; i < drN.Rows[0].ItemArray.Length; i++)
                        {
                            re[i] = drN.Rows[0][i].ToString();
                        }
                        return re;
                    }
                }
                else
                    return new string[0];

            }
            return null;
        }
        public DataRow? DB_GetFirstDataByDataRow(string sql)//取得第一筆資料
        {
            if (sql == "") { return null; }
            string err = "";
            DataTable drN = fn_GetDataTable(sql, ref err);
            if (drN != null)
            {
                if (drN.Rows.Count > 0)
                {
                    return drN.Rows[0];
                }
            }
            return null;
        }
        public DataRow? DB_GetFirstDataByDataRow(string sql, ref string err)//取得第一筆資料
        {
            if (sql == "") { return null; }
            DataTable drN = fn_GetDataTable(sql, ref err);
            if (drN != null)
            {
                if (drN.Rows.Count > 0)
                {
                    return drN.Rows[0];
                }
            }
            return null;
        }

        public bool DB_SetData(string sql)//異動資料
        {
            if (sql == "") { return false; }
            if (sql.Trim() != "")
            {
                string err = "";
                return fn_SetDataTable(sql, ref err);
            }
            return false;
        }

        /// <summary>
        /// Parameterized Insert/Update/Delete with automatic parameter creation
        /// </summary>
        public bool DB_SetDataByParams(string sql, Dictionary<string, object> parameters)
        {
            if (string.IsNullOrEmpty(sql)) { return false; }
            try
            {
                using (SqlConnection conn = new SqlConnection(ms_connectString))
                {
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        if (parameters != null)
                        {
                            foreach (var kvp in parameters)
                            {
                                cmd.Parameters.AddWithValue("@" + kvp.Key, kvp.Value ?? DBNull.Value);
                            }
                        }
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                if (Error != null)
                {
                    Error(sql, $"DB_SetDataByParams fail: {ex.Message}");
                }
                return false;
            }
        }
        public bool DB_SetData(string sql, ref string err)//異動資料
        {
            if (sql == "") { return false; }
            if (sql.Trim() != "")
            {
                return fn_SetDataTable(sql, ref err);
            }
            return false;
        }
        public bool DB_SetMSSQLData(string dbname, DataTable dt)
        {
            if (dt != null)
            {
                try
                {
                    using (SqlDataAdapter dataAdapter = new SqlDataAdapter(string.Format("select * from {0}", dbname), ms_connectString))
                    {
                        //SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                        dataAdapter.Update(dt);
                    }
                    return true;
                }
                catch (SqlException ex)
                {

                     if (Error != null) { Error(string.Format("select * from {0}", dbname), string.Format("DB_SetMSSQLData fail:{0}", ex.Message)); } //012
                }
                catch (Exception ex)
                {
 
                    if (Error != null) { Error(string.Format("select * from {0}", dbname), string.Format("DB_SetMSSQLData fail :{0}", ex.Message)); } //012
                }
            }
            return false;
        }
        public bool DB_SetMSSQLData(string connectString, string dbname, DataTable dt)
        {
            if (dt != null)
            {
                try
                {
                    using (SqlDataAdapter dataAdapter = new SqlDataAdapter(string.Format("select * from {0}", dbname), connectString))
                    {
                        //SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                        dataAdapter.Update(dt);
                    }
                    return true;
                }
                catch (SqlException ex)
                {

                    if (Error != null) { Error(string.Format("select * from {0}", dbname), string.Format("DB_SetMSSQLData fail:{0}", ex.Message)); }
                } //012
                catch (Exception ex)
                {

                    if (Error != null) { Error(string.Format("select * from {0}", dbname), string.Format("DB_SetMSSQLData fail :{0}", ex.Message)); }
                } //012
            }
            return false;
        }
        public string[][] DB_GetDataToListArray(string sql)//資料轉換2維陣列
        {
            string err = "";
            StringBuilder sb = new StringBuilder();
            DataTable drN = fn_GetDataTable(sql, ref err);
            if (drN != null)
            {
                if (drN.Rows.Count > 0)
                {
                    if (drN.Rows[0].ItemArray.Length > 0)
                    {
                        string[,] re = new string[drN.Rows.Count, drN.Rows[0].ItemArray.Length];
                        int i = 0; int j = 0;
                        for (i = 0; i < drN.Rows.Count; i++)
                        {
                            for (j = 0; j < drN.Rows[0].ItemArray.Length; j++)
                            {
                                re[i, j] = drN.Rows[i][j].ToString();
                            }
                        }
                    }

                }
            }
            return null;
        }
        public string[][] DB_GetDataToListArray(string sql, ref string err)//資料轉換2維陣列
        {
            StringBuilder sb = new StringBuilder();
            DataTable drN = fn_GetDataTable(sql, ref err);
            if (drN != null)
            {
                if (drN.Rows.Count > 0)
                {
                    if (drN.Rows[0].ItemArray.Length > 0)
                    {
                        string[,] re = new string[drN.Rows.Count, drN.Rows[0].ItemArray.Length];
                        int i = 0; int j = 0;
                        for (i = 0; i < drN.Rows.Count; i++)
                        {
                            for (j = 0; j < drN.Rows[0].ItemArray.Length; j++)
                            {
                                re[i, j] = drN.Rows[i][j].ToString();
                            }
                        }
                    }

                }
            }
            return null;
        }
        public DataTable GetMSSQL_Schema(string type)
        {
            DataTable re = null;
            using (SqlConnection mssqlconn = new SqlConnection(ms_connectString))
            {
                try
                {
                    mssqlconn.Open();
                    re = mssqlconn.GetSchema(type);
                }
                catch (SqlException ex)
                {

                    if (Error != null) { Error(type, string.Format("GetMSSQL_Schema fail:{0}", ex.Message)); }
                } //013
                catch (Exception ex)
                {

                    if (Error != null) { Error(type, string.Format("GetMSSQL_Schema fail:{0}", ex.Message)); }
                } //013
            }
            return re;
        }

        private DataTable fn_GetDataTable(string getSql, ref string err ) //sql=SQL語法
        {
            err = "";
            DataTable re = null;
            //            if (getSql.Trim().ToLower().IndexOf("select") != 0) { return null; }
            DataSet ds = new DataSet();
            switch (ms_connectType)
            {
                case "1"://MSSQL
                    using (SqlConnection mssqlconn = new SqlConnection(ms_connectString))
                    {
                        using (SqlCommand comm = new SqlCommand())
                        {
                            using (SqlDataAdapter dbAdapter = new SqlDataAdapter())
                            {
                                try
                                {
                                    List<object> sqlArgs = new List<object>();
                                    if (err == "1")
                                    {
                                        sqlArgs.Add("ServerId");
                                        sqlArgs.Add("01");
                                    }
                                    if (sqlArgs != null && sqlArgs.Count > 1)
                                    {
                                        DateTime result = new DateTime();
                                        for (var i = 0; i < sqlArgs.Count; i += 2)
                                        {
                                            var arg = comm.CreateParameter();
                                            arg.ParameterName = "@" + sqlArgs[i].ToString();
                                            //can not use iif here, has different type
                                            if (sqlArgs[i + 1] == null)
                                            { arg.Value = DBNull.Value; }
                                            else if (sqlArgs[i + 1].GetType() == typeof(DateTime) || DateTime.TryParse(sqlArgs[(i + 1)].ToString(), out result))
                                            {
                                                arg.Value = result.ToString("MM/dd/yyyy HH:mm:ss");
                                            }
                                            else
                                            { arg.Value = sqlArgs[i + 1].ToString(); }
                                            comm.Parameters.Add(arg);
                                        }
                                    }
                                    mssqlconn.Open();
                                    comm.Connection = mssqlconn;
                                    comm.CommandText = getSql;
                                    dbAdapter.SelectCommand = comm;
                                    dbAdapter.Fill(ds, "1");
                                    re = ds.Tables["1"];
                                }
                                catch (SqlException ex)
                                {

                                    if (Error != null) { Error(getSql, string.Format("fn_GetDataTable fail:{0}", ex.Message)); } //014
                                    else
                                    {
                                        err = ex.Message;
                                        //0000                                         Mail_Send("kurt@softnet.tw", $"SQL={getSql}<p>ERR={err}</p>", $"資料庫法語錯誤<p> 來源Connect:{ms_connectString}");
                                    }
                                }
                                catch (Exception ex)
                                {

                                    if (Error != null) { Error(getSql, string.Format("fn_GetDataTable fail:{0}", ex.Message)); } //014
                                    else
                                    {
                                        err = ex.Message;
                                        //0000                                        Mail_Send("kurt@softnet.tw", $"SQL={getSql}<p>ERR={err}</p>", $"資料庫法語錯誤<p> 來源Connect:{ms_connectString}");
                                    }
                                }
                            }
                        }
                    }
                    break;
            }
            ds.Dispose();
            return re;
        }
        private DataTable fn_GetDataTableByArgs(string getSql, List<object> sqlArgs, ref string err) //sql=SQL語法
        {
            err = "";
            DataTable re = null;
            DataSet ds = new DataSet();
            switch (ms_connectType)
            {
                case "1"://MSSQL
                    using (SqlConnection mssqlconn = new SqlConnection(ms_connectString))
                    {
                        using (SqlCommand comm = new SqlCommand())
                        {
                            using (SqlDataAdapter dbAdapter = new SqlDataAdapter())
                            {
                                try
                                {
                                    if (sqlArgs != null && sqlArgs.Count > 1)
                                    {
                                        DateTime result = new DateTime();
                                        for (var i = 0; i < sqlArgs.Count; i += 2)
                                        {
                                            var arg = comm.CreateParameter();
                                            arg.ParameterName = "@" + sqlArgs[i].ToString();
                                            //can not use iif here, has different type
                                            if (sqlArgs[i + 1] == null)
                                            { arg.Value = DBNull.Value; }
                                            else if (sqlArgs[i + 1].GetType() == typeof(DateTime) || DateTime.TryParse(sqlArgs[(i + 1)].ToString(), out result))
                                            {
                                                arg.Value = result.ToString("MM/dd/yyyy HH:mm:ss");
                                            }
                                            else
                                            { arg.Value = sqlArgs[i + 1].ToString(); }
                                            comm.Parameters.Add(arg);
                                        }
                                    }
                                    mssqlconn.Open();
                                    comm.Connection = mssqlconn;
                                    comm.CommandText = getSql;
                                    dbAdapter.SelectCommand = comm;
                                    dbAdapter.Fill(ds, "1");
                                    re = ds.Tables["1"];
                                }
                                catch (SqlException ex)
                                {
                                    if (Error != null) 
                                    {
                                                          Error(getSql, string.Format("fn_GetDataTable fail:{0}", ex.Message)); 
                                    } //014
                                    else
                                    {
                                        err = ex.Message;
                                        //0000                                      Mail_Send("kurt@softnet.tw", $"SQL={getSql}<p>ERR={err}</p>", $"資料庫法語錯誤<p> 來源Connect:{ms_connectString}");
                                    }
                                }
                                catch (Exception ex)
                                {

                                    if (Error != null) 
                                    {
                                                          Error(getSql, string.Format("fn_GetDataTable fail:{0}", ex.Message));
                                    } //014
                                    else
                                    {
                                        err = ex.Message;
                                        //0000                                      Mail_Send("kurt@softnet.tw", $"SQL={getSql}<p>ERR={err}</p>", $"資料庫法語錯誤<p> 來源Connect:{ms_connectString}");
                                    }
                                }
                            }
                        }
                    }
                    break;
            }
            ds.Dispose();
            return re;
        }


        private bool fn_SetDataTable(string setSql, ref string err) //sql=SQL語法
        {
            err = "";
            bool re = false;
            DataSet ds = new DataSet();
            switch (ms_connectType)
            {
                case "1"://MSSQL
                    using (SqlConnection mssqlconn = new SqlConnection(ms_connectString))
                    {
                        using (SqlCommand comm = new SqlCommand())
                        {
                            try
                            {
                                mssqlconn.Open();
                                comm.Connection = mssqlconn;
                                if (setSql.Trim() != "")
                                {
                                    comm.CommandText = setSql;
                                    comm.ExecuteNonQuery();
                                    re = true;
                                }
                            }
                            catch (SqlException ex)
                            {
                                if (Error != null) 
                                {
                      //                                     Error(setSql, string.Format("fn_SetDataTable fail:{0} sql={1}", ex.Message, setSql)); 
                                }
                                else 
                                { 
                                    err = ex.Message;
                                    //0000                       Mail_Send("kurt@softnet.tw", $"SQL={setSql}<p>ERR={err}</p>", $"資料庫法語錯誤<p> 來源Connect:{ms_connectString}");
                                }
                            } //015
                            catch (Exception ex)
                            {
                                if (Error != null)
                                {
                      //                                    Error(setSql, string.Format("fn_SetDataTable fail:{0} sql={1}", ex.Message, setSql));
                                }
                                else 
                                { 
                                    err = ex.Message;
                                    //0000                     Mail_Send("kurt@softnet.tw", $"SQL={setSql}<p>ERR={err}</p>", $"資料庫法語錯誤<p> 來源Connect:{ms_connectString}");
                                }
                            } //015
                        }
                    }
                    break;
            }
            ds.Dispose();
            return re;
        }
        private void Mail_Send(string MailTos, string MailBody, string MailSub)
        {
            MailMessage mms = new MailMessage();
            object Lock_Mail_Send = new object();
            if (MailTos.Trim() == "") { return; }
            lock (Lock_Mail_Send)
            {
                try
                {

                    mms.From = new MailAddress("kurt@softnet.tw");
                    mms.Subject = MailSub;
                    //MailBody = MailBody.Replace("\n", "<br>").Replace("\r", "");
                    mms.Body = MailBody;
                    //信件內容 是否採用Html格式
                    mms.IsBodyHtml = true;
                    mms.SubjectEncoding = Encoding.UTF8;
                    mms.To.Add(new MailAddress(MailTos));
                    if (mms.To.Count <= 0) { return; }

                    string smtpServer = "smtp.gmail.com";
                    int smtpPort = 587;
                    string mailAccount = "kurt@softnet.tw";
                    string mailPwd = "ynnzircjohfcobov";
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

                }
                catch (Exception ex)
                {
                    string _s = "";
                }
            }
            mms.Dispose();
        }


        #region 2019/04/19 Neil add for sql parameter
        public enum SqlNonQueryStyle { WithRollBack, None }//在NonQuery語法是否做RollBack NeilBug
        public DataTable DB_Query(string connectionstring, string sqlexcutestring, Dictionary<string, object> sqlparameters = null)
        {
            DataTable dt = new DataTable();

            if (string.IsNullOrEmpty(sqlexcutestring) || string.IsNullOrEmpty(connectionstring))
                return dt;
            switch (ms_connectType)
            {
                case "1"://MSSQL
                    using (SqlConnection sqlConnection = new SqlConnection(connectionstring))
                    {
                        sqlConnection.Open();
                        try
                        {

                            SqlCommand sqlCommand = new SqlCommand(sqlexcutestring, sqlConnection);
                            if (sqlparameters != null)
                            {
                                foreach (string key in sqlparameters.Keys)
                                {
                                    sqlCommand.Parameters.Add(new SqlParameter(key, sqlparameters[key]));
                                }
                            }

                            using (SqlDataReader sqlDataReader = sqlCommand.ExecuteReader())
                            {
                                dt.Load(sqlDataReader);
                            }
                        }
                        catch (SqlException sqlex)
                        {
                            sqlConnection.Close();
                            throw new Exception("Database exception", sqlex); //017

                        }

                        catch (Exception ex)
                        {
                            sqlConnection.Close();
                            throw new Exception("Database exception", ex); //018
                        }
                        break;

                    }
            }
            return dt;
        }
        public DataTable DB_Query(string sqlexcutestring, Dictionary<string, object> sqlparameters = null)
        {
            DataTable dt = new DataTable();
            if (string.IsNullOrEmpty(sqlexcutestring))
                return dt;
            switch (ms_connectType)
            {
                case "1"://MSSQL
                    using (SqlConnection sqlConnection = new SqlConnection(ms_connectString))
                    {

                        try
                        {
                            sqlConnection.Open();
                            SqlCommand sqlCommand = new SqlCommand(sqlexcutestring, sqlConnection);
                            if (sqlparameters != null)
                            {
                                foreach (string key in sqlparameters.Keys)
                                {
                                    sqlCommand.Parameters.Add(new SqlParameter(key, sqlparameters[key]));
                                }
                            }

                            using (SqlDataReader sqlDataReader = sqlCommand.ExecuteReader())
                            {
                                dt.Load(sqlDataReader);
                            }
                        }
                        catch (SqlException sqlex)
                        {
                            sqlConnection.Close();
                            throw new Exception("Database exception", sqlex); //017

                        }

                        catch (Exception ex)
                        {
                            sqlConnection.Close();

                            throw new Exception("Database exception", ex); //018
                        }
                        break;

                    }
            }
            return dt;
        }
        public DataSet DB_Querys(string sqlexcutestring, Dictionary<string, object> sqlparameters = null)
        {
            DataSet dts = new DataSet();
            switch (ms_connectType)
            {
                case "1"://MSSQL
                    using (SqlConnection sqlConnection = new SqlConnection(ms_connectString))
                    {

                        try
                        {
                            sqlConnection.Open();
                            SqlCommand sqlCommand = new SqlCommand(sqlexcutestring, sqlConnection);
                            if (sqlparameters != null)
                            {
                                foreach (string key in sqlparameters.Keys)
                                {
                                    sqlCommand.Parameters.Add(new SqlParameter(key, sqlparameters[key]));
                                }
                            }

                            using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand))
                            {
                                sqlDataAdapter.Fill(dts);
                            }
                        }
                        catch (SqlException sqlex)
                        {
                            sqlConnection.Close();
                            throw new Exception("Database exception", sqlex); //017
                        }
                        catch (Exception ex)
                        {
                            sqlConnection.Close();
                            throw new Exception("Database exception", ex); //018
                        }
                    }
                    break;
            }

            return dts;
        }
        public void DB_NonQuery(SqlNonQueryStyle sqlNonQueryStyle, string sqlexcutestring, Dictionary<string, object> sqlparameters = null)
        {
            switch (ms_connectType)
            {
                case "1"://MSSQL
                    using (SqlConnection sqlConnection = new SqlConnection(ms_connectString))
                    {

                        switch (sqlNonQueryStyle)
                        {
                            case SqlNonQueryStyle.WithRollBack:
                                SqlTransaction sqlTrans = null;
                                try
                                {
                                    sqlConnection.Open();
                                    sqlTrans = sqlConnection.BeginTransaction();
                                    SqlCommand sqlCommand = new SqlCommand(sqlexcutestring, sqlConnection);
                                    if (sqlparameters != null)
                                    {
                                        foreach (string key in sqlparameters.Keys)
                                        {
                                            sqlCommand.Parameters.Add(new SqlParameter(key, sqlparameters[key]));
                                        }
                                    }

                                    sqlCommand.Transaction = sqlTrans;
                                    sqlCommand.ExecuteNonQuery();
                                    sqlTrans.Commit();
                                }
                                catch (SqlException sqlex)
                                {
                                    if (sqlTrans != null)
                                        sqlTrans.Rollback();
                                    sqlConnection.Close();
                                    throw new Exception("Database exception", sqlex); //017
                                }
                                catch (Exception ex)
                                {

                                    sqlConnection.Close();
                                    throw new Exception("Database exception", ex); //018
                                }
                                break;
                            case SqlNonQueryStyle.None:
                                try
                                {
                                    sqlConnection.Open();
                                    SqlCommand sqlCommand = new SqlCommand(sqlexcutestring, sqlConnection);
                                    if (sqlparameters != null)
                                    {
                                        foreach (string key in sqlparameters.Keys)
                                        {
                                            sqlCommand.Parameters.Add(new SqlParameter(key, sqlparameters[key]));
                                        }
                                    }
                                    sqlCommand.ExecuteNonQuery();
                                }
                                catch (SqlException sqlex)
                                {
                                    sqlConnection.Close();
                                    throw new Exception("Database exception", sqlex); //017
                                }
                                catch (Exception ex)
                                {
                                    sqlConnection.Close();
                                    throw new Exception("program exception", ex); //018
                                }
                                break;
                        }
                    }
                    break;
            }
        }
        /// <summary>
        /// 從DataTable批次匯入至Sql Table
        /// </summary>
        /// <param name="Data">匯入資料來源</param>
        /// <param name="TableName">匯入database table name</param>
        /// <param name="dataColumns">匯入欄位</param>
        public void DB_BulkInsert(DataTable Data, string TableName, DataColumn[] dataColumns = null)
        {
            using (SqlConnection sqlConnection = new SqlConnection(ms_connectString))
            {
                sqlConnection.Open();

                SqlTransaction sqlTrans = sqlConnection.BeginTransaction();
                try
                {

                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlConnection, SqlBulkCopyOptions.Default, sqlTrans))
                    {

                        sqlBulkCopy.DestinationTableName = TableName;
                        if (dataColumns != null)
                            foreach (DataColumn dc in dataColumns)
                                sqlBulkCopy.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);

                        sqlBulkCopy.WriteToServer(Data);
                        sqlTrans.Commit();
                    }
                }
                catch (SqlException sqlex)
                {
                    sqlTrans.Rollback();
                    sqlConnection.Close();

                    throw new Exception("Database exception", sqlex); //017

                }
                catch (Exception ex)
                {
                    sqlTrans.Rollback();
                    sqlConnection.Close();
                    throw new Exception("program exception", ex); //018
                }

            }
        }

        //從DataSet批次匯入至Sql Table
        public void DB_BulkInsertDs(DataSet dsData, string DBConnectString)
        {
            using (SqlConnection sqlConnection = new SqlConnection(DBConnectString))
            {
                sqlConnection.Open();
                SqlTransaction sqlTrans = sqlConnection.BeginTransaction();
                try
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlConnection, SqlBulkCopyOptions.Default, sqlTrans))
                    {
                        foreach (DataTable dt in dsData.Tables)
                        {
                            sqlBulkCopy.ColumnMappings.Clear();
                            sqlBulkCopy.DestinationTableName = dt.TableName;
                            foreach (DataColumn dc in dt.Columns)
                            {
                                sqlBulkCopy.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                            }
                            sqlBulkCopy.WriteToServer(dt);
                        }
                        sqlTrans.Commit();
                    }
                }
                catch (SqlException sqlex)
                {
                    sqlTrans.Rollback();
                    sqlConnection.Close();
                    throw new Exception("Database exception", sqlex); //017
                }
                catch (Exception ex)
                {
                    sqlTrans.Rollback();
                    sqlConnection.Close();
                    throw new Exception("program exception", ex); //018
                }
            }
        }
        #endregion
    }

}
