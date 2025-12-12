using Base;
using Base.Models;
using Base.Services;
using HandlebarsDotNet.Features;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using Newtonsoft.Json.Linq;
using Pipelines.Sockets.Unofficial.Arenas;

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class ReportService
    {
    }
    public class Report03Service : XgEdit
    {
        public Report03Service(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG]",
                PkeyFid = "Id",
                PkeyFids = new string[] {  "LOGDateTime", "Id", "ServerId", "StationNO", "OP_NO" },
                Col4 = null,
                Items = new EitemDto[] {
                new() { Fid = "ServerId" },
                new() { Fid = "LOGDateTime" },
                    new() { Fid = "LOGDateTimeID" },
                    new() { Fid = "StationNO" },
                    new() { Fid = "PartNO" },
                    new() { Fid = "OP_NO" },
                    new() { Fid = "OKQTY" },
                    new() { Fid = "FailQTY" },
                     new() { Fid = "ServerId" },
                },
            };
        }

        
        private readonly ReadDto dto = new()
        {
            ReadSql = @"select  '' as _Fun,a.* from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG]",
            Items = new QitemDto[] {
                 new() { Fid = "ServerId" },
                new() { Fid = "StationNO" },
                new() { Fid = "LOGDateTime" },
                new() { Fid = "LOGDateTimeID" },
                 new() { Fid = "PartNO" },
                new() { Fid = "PP_Name" },
                new() { Fid = "OP_NO" },
            },
            SQL_StoredProgram = "",
        };

        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            dto.SQL_ByProgram = Run_ProgramReadDataOBJ(dt, ctrl);
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
        private ProgramReadDataOBJ Run_ProgramReadDataOBJ(DtDto dt,string ctrl)
        {
            ProgramReadDataOBJ re = new ProgramReadDataOBJ();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($@"select  '' as _Fun,a.*,b.Name as OP_Name, 0 as EfficientCycleTime,(select top 1 (c.PartName + ' ' + c.Specification)  from SoftNetMainDB.[dbo].[Material] as c where c.ServerId='{_Fun.Config.ServerId}' and c.PartNO=a.PartNO) as PartNameSpecification
                                                    from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] as a,SoftNetMainDB.[dbo].[User] as b 
                                                    where a.ServerId='{_Fun.Config.ServerId}' and a.OP_NO=b.UserNO order by a.LOGDateTime desc,a.StationNO,a.Id,a.OP_NO");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    DataRow dr2 = null;
                    re.rows = new JArray();
                    JObject row;
                    int off = dt.start + dt.length;
                    if (off > tmp_dt.Rows.Count) { off = tmp_dt.Rows.Count; }
                    for (int i = dt.start; i < off; i++)
                    {
                        row = new JObject();
                        dr2 = tmp_dt.Rows[i];
                        foreach (System.Data.DataColumn col in tmp_dt.Columns)
                        {
                            row.Add(col.ColumnName.Trim(), JToken.FromObject(dr2[col]));
                        }
                        re.rows.Add(row);
                    }
                    re.RowCount = tmp_dt.Rows.Count;
                }
            }

            return re;
        }


    } //class


    public class Report04Service : XgEdit
    {
        public Report04Service(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG]",
                PkeyFid = "LOGDateTimeID",
                PkeyFids = new string[] { "ServerId", "LOGDateTime", "LOGDateTimeID", "StationNO", "PartNO" },
                Col4 = null,
                Items = new EitemDto[] {
                    new() { Fid = "ServerId" },
                    new() { Fid = "LOGDateTime" },
                    new() { Fid = "LOGDateTimeID" },
                    new() { Fid = "StationNO" },
                    new() { Fid = "PartNO" },
                    new() { Fid = "EditFinishedQty" },
                    new() { Fid = "EditFailedQty" },
                    new() { Fid = "OP_NO" },
                    new() { Fid = "CycleTime" },
                },
            };
        }


        private readonly ReadDto dto = new()
        {
            ReadSql = @"select  '' as _Fun,PartNO,LOGDateTime,LOGDateTimeID,StationNO,PP_Name,OP_NO,EditFinishedQty,EditFailedQty,CycleTime from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG]",
            Items = new QitemDto[] {
                new() { Fid = "LOGDateTime" },
                new() { Fid = "StationNO" },
                new() { Fid = "PP_Name" },
                new() { Fid = "OP_NO" },
            },
            SQL_StoredProgram = "",
        };

        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            dto.SQL_ByProgram = Run_ProgramReadDataOBJ(dt, ctrl);
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
        private ProgramReadDataOBJ Run_ProgramReadDataOBJ(DtDto dt, string ctrl)
        {
            ProgramReadDataOBJ re = new ProgramReadDataOBJ();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string whereData = $"a.LOGDateTimeID>='{DateTime.Now.AddDays(-60).ToString("yyyy/MM/dd HH:mm:ss")}' and a.LOGDateTimeID<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}'";
                if (dt.findJson!=null)
                {
                    JObject jObject = JObject.Parse(dt.findJson);
                    string DOCDate = (string)jObject["DOCDate"];
                    string DOCDate2 = (string)jObject["DOCDate2"];
                    if (DOCDate!=null && DOCDate2!=null)
                    { whereData= $"a.LOGDateTimeID>='{DOCDate}' and a.LOGDateTimeID<='{DOCDate2}'"; }
                    string stationNO = (string)jObject["StationNO"];
                    if (stationNO != null && stationNO.Trim()!="")
                    { whereData = $"{whereData} and a.StationNO='{stationNO}'"; }
                    string PP_Name = (string)jObject["PP_Name"];
                    if (PP_Name != null && PP_Name.Trim() != "")
                    { whereData = $"{whereData} and a.PP_Name='{PP_Name}'"; }
                    string OP_NO = (string)jObject["OP_NO"];
                    if (OP_NO != null && OP_NO.Trim() != "")
                    { whereData = $"{whereData} and a.OP_NO='{OP_NO}'"; }
                    string OrderNO = (string)jObject["OrderNO"];
                    if (OrderNO != null && OrderNO.Trim() != "")
                    { whereData = $"{whereData} and a.OrderNO like '%{OrderNO}%'"; }

                }
                DataTable tmp_dt = db.DB_GetData($@"select  '' as _Fun,a.*,(select top 1 (b.PartNO + ' ' + b.PartName + ' ' + b.Specification)  from SoftNetMainDB.[dbo].[Material] as b where b.ServerId='{_Fun.Config.ServerId}' and b.PartNO=a.PartNO) as PartNOName,
                                                    (select top 1 Name from SoftNetMainDB.[dbo].[User] as e where e.ServerId='{_Fun.Config.ServerId}' and e.UserNO=a.OP_NO) as UserNOName,
                                                    (select top 1 DisplayName from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] as c where c.ServerId='{_Fun.Config.ServerId}' and a.PP_Name=c.PP_Name and a.PartNO=c.PartNO and a.StationNO=c.StationNO and a.IndexSN=c.IndexSN) as DisplayName,
                                                    (select top 1 OrderNO from SoftNetLogDB.[dbo].[SFC_StationDetail] as d where d.ServerId='{_Fun.Config.ServerId}' and a.LOGDateTime=d.LOGDateTime)  as OrderNO
                                                    from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] as a
                                                    where a.ServerId='{_Fun.Config.ServerId}' and {whereData} order by a.LOGDateTime desc,a.StationNO,a.OP_NO");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    DataRow dr2 = null;
                    re.rows = new JArray();
                    JObject row;
                    int off = dt.start + dt.length;
                    if (off > tmp_dt.Rows.Count) { off = tmp_dt.Rows.Count; }
                    for (int i = dt.start; i < off; i++)
                    {
                        row = new JObject();
                        dr2 = tmp_dt.Rows[i];
                        foreach (System.Data.DataColumn col in tmp_dt.Columns)
                        {
                            row.Add(col.ColumnName.Trim(), JToken.FromObject(dr2[col]));
                        }
                        re.rows.Add(row);
                    }
                    re.RowCount = tmp_dt.Rows.Count;
                }
            }

            return re;
        }


    } //class
    public class Report05Service : XgEdit
    {
        public Report05Service(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetLogDB.[dbo].[OperateLog]",
                PkeyFid = "Id",
                PkeyFids = new string[] { "ServerId", "StationNO", "Id" },
                Col4 = null,
                Items = new EitemDto[] {
                    new() { Fid = "ServerId" },
                    new() { Fid = "Id" },
                    new() { Fid = "StationNO" },
                },
            };
        }


        private readonly ReadDto dto = new()
        {
            ReadSql = $@"SELECT b.PartName,b.Specification,a.Id,StationNO,a.ServerId,a.LOGDateTime,a.ProgramName,a.OperateType,a.PartNO,a.OrderNO,a.OP_NO,a.IndexSN,a.Remark FROM SoftNetLogDB.[dbo].[OperateLog] as a
                        join SoftNetMainDB.[dbo].[Material] as b on b.PartNO = a.PartNO and b.ServerId = '{_Fun.Config.ServerId}'
                        where a.ServerId='{_Fun.Config.ServerId}' and a.LOGDateTime>='{DateTime.Now.AddDays(-60).ToString("yyyy/MM/dd HH:mm:ss.fff")}' and a.LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}' order by a.LOGDateTime desc,a.StationNO,a.OP_NO,a.PartNO",
            Items = new QitemDto[] {
                new() { Fid = "LOGDateTime" },
                new() { Fid = "StationNO" },
                new() { Fid = "PartNO" },
                new() { Fid = "PartName" },
                new() { Fid = "OrderNO" },
                new() { Fid = "OperateType" },
                new() { Fid = "OP_NO" },
            },
            SQL_StoredProgram = "",
        };

        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            dto.SQL_ByProgram = Run_ProgramReadDataOBJ(dt, ctrl);
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
        private ProgramReadDataOBJ Run_ProgramReadDataOBJ(DtDto dt, string ctrl)
        {
            ProgramReadDataOBJ re = new ProgramReadDataOBJ();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string whereData = $"a.LOGDateTime>='{DateTime.Now.AddDays(-60).ToString("yyyy/MM/dd HH:mm:ss")}' and a.LOGDateTime<='{DateTime.Now.ToString("yyyy/MM/dd")}'";
                if (dt.findJson != null)
                {
                    JObject jObject = JObject.Parse(dt.findJson);
                    string DOCDate = (string)jObject["DOCDate"];
                    string DOCDate2 = (string)jObject["DOCDate2"];
                    if (DOCDate != null && DOCDate2 != null)
                    { whereData = $"a.LOGDateTime>='{DOCDate}' and a.LOGDateTime<='{DOCDate2}'"; }
                    string stationNO = (string)jObject["StationNO"];
                    if (stationNO != null && stationNO.Trim() != "")
                    { whereData = $"{whereData} and a.StationNO='{stationNO}'"; }
                    string operateType = (string)jObject["OperateType"];
                    if (operateType != null && operateType.Trim() != "")
                    { whereData = $"{whereData} and a.OperateType like '%{operateType}%'"; }
                    string PP_Name = (string)jObject["Remark"];
                    if (PP_Name != null && PP_Name.Trim() != "")
                    { whereData = $"{whereData} and a.Remark like '%{PP_Name}%'"; }
                    string PartNO = (string)jObject["PartNO"];
                    if (PartNO != null && PartNO.Trim() != "")
                    { whereData = $"{whereData} and a.PartNO='{PartNO}'"; }

                    string PartName = (string)jObject["PartName"];
                    if (PartName != null && PartName.Trim() != "")
                    { whereData = $"{whereData} and b.PartName like '%{PartName}%'"; }

                    string OrderNO = (string)jObject["OrderNO"];
                    if (OrderNO != null && OrderNO.Trim() != "")
                    { whereData = $"{whereData} and a.OrderNO like '%{OrderNO}%'"; }

                    string OP_NO = (string)jObject["OP_NO"];
                    if (OP_NO != null && OP_NO.Trim() != "")
                    { whereData = $"{whereData} and a.OP_NO like '%{OP_NO}%'"; }
                }

                string Sql = $@"SELECT b.PartName,b.Specification,a.Id,a.StationNO,c.StationName,a.ServerId,a.LOGDateTime,a.ProgramName,a.OperateType,a.PartNO,a.OrderNO,a.OP_NO as UserNOName,a.IndexSN as DisplayName,a.Remark FROM SoftNetLogDB.[dbo].[OperateLog] as a
                        join SoftNetMainDB.[dbo].[Material] as b on b.PartNO = a.PartNO and b.ServerId = '{_Fun.Config.ServerId}'
                        join SoftNetSYSDB.[dbo].[PP_Station] as c on c.StationNO = a.StationNO and c.ServerId = '{_Fun.Config.ServerId}'
                        where a.ServerId='{_Fun.Config.ServerId}' and {whereData} order by a.LOGDateTime desc,a.StationNO,a.OP_NO,a.PartNO";

                //DataTable tmp_dt = db.DB_GetData($@"select  a.*,b.PartName,b.Specification,
                //                                    (select top 1 DisplayName from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] as c where c.ServerId='{_Fun.Config.ServerId}' and a.PP_Name=c.PP_Name and a.PartNO=c.PartNO and a.StationNO=c.StationNO and a.IndexSN=c.IndexSN) as DisplayName
                //                                    from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] as a,SoftNetMainDB.[dbo].[Material] as b
                //                                    where a.ServerId='{_Fun.Config.ServerId}' and b.ServerId='{_Fun.Config.ServerId}' and {whereData} order by a.LOGDateTime desc,a.StationNO,a.OP_NO");
                DataTable tmp_dt = db.DB_GetData(Sql);
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    DataRow dr2 = null;
                    re.rows = new JArray();
                    JObject row;
                    int off = dt.start + dt.length;
                    if (off > tmp_dt.Rows.Count) { off = tmp_dt.Rows.Count; }
                    for (int i = dt.start; i < off; i++)
                    {
                        row = new JObject();
                        dr2 = tmp_dt.Rows[i];
                        foreach (System.Data.DataColumn col in tmp_dt.Columns)
                        {
                            row.Add(col.ColumnName.Trim(), JToken.FromObject(dr2[col]));
                        }
                        re.rows.Add(row);
                    }
                    re.RowCount = tmp_dt.Rows.Count;
                }
            }

            return re;
        }


    } //class
    public class Report06Service : XgEdit
    {
        public Report06Service(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetSYSDB.[dbo].[APS_WorkTimeNote]",
                PkeyFid = "SimulationId",
                PkeyFids = new string[] { "ServerId", "NeedId", "SimulationId", "StationNO" },
                Col4 = null,
                Items = new EitemDto[] {
                    new() { Fid = "ServerId" },
                    new() { Fid = "LOGDateTime" },
                    new() { Fid = "LOGDateTimeID" },
                    new() { Fid = "StationNO" },
                    new() { Fid = "PartNO" },
                    new() { Fid = "EditFinishedQty" },
                    new() { Fid = "EditFailedQty" },
                    new() { Fid = "OP_NO" },
                    new() { Fid = "CycleTime" },
                },
            };
        }


        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select '' as _Fun,a.ServerId,a.NeedId,a.SimulationId,a.StationNO,sum(a.Time1_C+a.Time2_C+a.Time3_C+a.Time4_C) as TOTTime,sum(a.Time_TOT) as Time_TOT,b.Source_StationNO_IndexSN,b.Source_StationNO_Custom_DisplayName FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] as a
                      join SoftNetSYSDB.[dbo].[APS_Simulation] as b on b.NeedId=b.NeedId and a.SimulationId=b.SimulationId and b.IsOK='0'
                      where a.ServerId={_Fun.Config.ServerId} and b.Is group by a.ServerId,a.NeedId,a.SimulationId,a.StationNO,b.Source_StationNO_IndexSN,b.Source_StationNO_Custom_DisplayName order by a.StationNO,a.NeedId,b.Source_StationNO_IndexSN",
            Items = new QitemDto[] {
                new() { Fid = "LOGDateTime" },
                new() { Fid = "StationNO" },
                new() { Fid = "PP_Name" },
                new() { Fid = "OP_NO" },
            },
            SQL_StoredProgram = "",
        };

        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            dto.SQL_ByProgram = Run_ProgramReadDataOBJ(dt, ctrl);
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
        private ProgramReadDataOBJ Run_ProgramReadDataOBJ(DtDto dt, string ctrl)
        {
            ProgramReadDataOBJ re = new ProgramReadDataOBJ();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string whereData = $"a.LOGDateTimeID>='{DateTime.Now.AddDays(-60).ToString("yyyy/MM/dd HH:mm:ss")}' and a.LOGDateTimeID<='{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")}'";
                if (dt.findJson != null)
                {
                    JObject jObject = JObject.Parse(dt.findJson);
                    string DOCDate = (string)jObject["DOCDate"];
                    string DOCDate2 = (string)jObject["DOCDate2"];
                    if (DOCDate != null && DOCDate2 != null)
                    { whereData = $"a.LOGDateTimeID>='{DOCDate}' and a.LOGDateTimeID<='{DOCDate2}'"; }
                    string stationNO = (string)jObject["StationNO"];
                    if (stationNO != null && stationNO.Trim() != "")
                    { whereData = $"{whereData} and a.StationNO='{stationNO}'"; }
                    string PP_Name = (string)jObject["PP_Name"];
                    if (PP_Name != null && PP_Name.Trim() != "")
                    { whereData = $"{whereData} and a.PP_Name='{PP_Name}'"; }
                    string OP_NO = (string)jObject["OP_NO"];
                    if (OP_NO != null && OP_NO.Trim() != "")
                    { whereData = $"{whereData} and a.OP_NO='{OP_NO}'"; }
                    string OrderNO = (string)jObject["OrderNO"];
                    if (OrderNO != null && OrderNO.Trim() != "")
                    { whereData = $"{whereData} and a.OrderNO like '%{OrderNO}%'"; }

                }
                DataTable tmp_dt = db.DB_GetData($@"select  '' as _Fun,a.*,(select top 1 (b.PartNO + ' ' + b.PartName + ' ' + b.Specification)  from SoftNetMainDB.[dbo].[Material] as b where b.ServerId='{_Fun.Config.ServerId}' and b.PartNO=a.PartNO) as PartNOName,
                                                    (select top 1 Name from SoftNetMainDB.[dbo].[User] as e where e.ServerId='{_Fun.Config.ServerId}' and e.UserNO=a.OP_NO) as UserNOName,
                                                    (select top 1 DisplayName from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] as c where c.ServerId='{_Fun.Config.ServerId}' and a.PP_Name=c.PP_Name and a.PartNO=c.PartNO and a.StationNO=c.StationNO and a.IndexSN=c.IndexSN) as DisplayName,
                                                    (select top 1 OrderNO from SoftNetLogDB.[dbo].[SFC_StationDetail] as d where d.ServerId='{_Fun.Config.ServerId}' and a.LOGDateTime=d.LOGDateTime)  as OrderNO
                                                    from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] as a
                                                    where a.ServerId='{_Fun.Config.ServerId}' and {whereData} order by a.LOGDateTime desc,a.StationNO,a.OP_NO");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    DataRow dr2 = null;
                    re.rows = new JArray();
                    JObject row;
                    int off = dt.start + dt.length;
                    if (off > tmp_dt.Rows.Count) { off = tmp_dt.Rows.Count; }
                    for (int i = dt.start; i < off; i++)
                    {
                        row = new JObject();
                        dr2 = tmp_dt.Rows[i];
                        foreach (System.Data.DataColumn col in tmp_dt.Columns)
                        {
                            row.Add(col.ColumnName.Trim(), JToken.FromObject(dr2[col]));
                        }
                        re.rows.Add(row);
                    }
                    re.RowCount = tmp_dt.Rows.Count;
                }
            }

            return re;
        }


    } //class


    public class CIR01Service : XgEdit
    {
        public CIR01Service(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log]",
                PkeyFid = "Id",
                PkeyFids = new string[] { "LOGDateTime", "Id", "ServerId", "StationNO", "OP_NO" },
                Col4 = null,
                ReadSql= "select * from SoftNetSYSDB.[dbo].[PP_Station] where StationNO='{0}'",
                Items = new EitemDto[] {
                    new() { Fid = "Id" },
                    new() { Fid = "LOGDateTime" },
                    new() { Fid = "StationNO" },
                    new() { Fid = "OP_NO" },
                    new() { Fid = "OKQTY" },
                    new() { Fid = "FailQTY" },
                     new() { Fid = "ServerId" },
                },
            };
        }

        private readonly ReadDto dto = new()
        {
            //ReadSql = @"select  '' as _Fun,a.* from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] as a",
            Items = new QitemDto[] {
                new() { Fid = "StationNO" },
                new() { Fid = "PP_Name" },
                new() { Fid = "OP_NO" },
                new() { Fid = "OrderNO" },
            },
            SQL_StoredProgram = "",
        };

        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            dto.SQL_ByProgram = Run_ProgramReadDataOBJ(dt, ctrl);
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
        private ProgramReadDataOBJ Run_ProgramReadDataOBJ(DtDto dt, string ctrl)
        {
            SFC_Common _SFC_Common = null;
            ProgramReadDataOBJ re = new ProgramReadDataOBJ();
            DataTable tmp_dt = new DataTable();
            string pageName = "";
            string sql = "";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                try
                {
                    _SFC_Common = (SFC_Common)_Fun.DiBox.GetService(typeof(SFC_Common));
                    if (_SFC_Common == null) { return re; }
                    if (dt.locationURL != null && dt.locationURL.Trim() != "")
                    {
                        string[] tmp_url = dt.locationURL.Split('/');
                        pageName = tmp_url[tmp_url.Length - 1];
                    }
                    switch (pageName)
                    {
                        case "":
                            break;
                        case "APSEventLog":
                            {
                                DateTime dtime = DateTime.Now.AddDays(-7); ;
                                tmp_dt.Columns.Add("Time");
                                tmp_dt.Columns.Add("Log");
                                DataTable dtdata = db.DB_GetData($"SELECT * FROM SoftNetLogDB.[dbo].[APS_EventLog] where ServerId='{_Fun.Config.ServerId}' and LOGDateTime>='{dtime.ToString("yyyy/MM/dd")}' order by LOGDateTime desc");
                                if (dtdata != null && dtdata.Rows.Count > 0)
                                {
                                    foreach (DataRow dr in dtdata.Rows)
                                    {
                                        tmp_dt.Rows.Add(dr["LOGDateTime"].ToString(), dr["Event"].ToString());
                                    }
                                }
                            }
                            break;
                        case "APSErrorData":
                            {
                                DateTime dtime = DateTime.Now.AddDays(-7); ;
                                tmp_dt.Columns.Add("Time");
                                tmp_dt.Columns.Add("SimulationId");
                                tmp_dt.Columns.Add("DOCNumberNO");
                                tmp_dt.Columns.Add("Name");
                                tmp_dt.Columns.Add("Log");
                                DataTable dtdata = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and LogDate>='{dtime.ToString("yyyy/MM/dd")}' order by LogDate desc");
                                if (dtdata != null && dtdata.Rows.Count > 0)
                                {
                                    string name = "";
                                    string log = "";
                                    string tmp_S = "";
                                    string bdfor_SimulationId = "";
                                    string bdfor_DOCNumberNO = "";
                                    string bdfor_StationNO = "";
                                    string bdfor_ErrorType = "";
                                    string IDName = "";
                                    DataRow tmp = null;
                                    foreach (DataRow dr in dtdata.Rows)
                                    {
                                        if (bdfor_SimulationId == dr["SimulationId"].ToString() && bdfor_DOCNumberNO == dr["DOCNumberNO"].ToString() && bdfor_StationNO == dr["StationNO"].ToString() && bdfor_ErrorType == dr["ErrorType"].ToString())
                                        { continue; }
                                        else { bdfor_SimulationId = dr["SimulationId"].ToString(); bdfor_DOCNumberNO = dr["DOCNumberNO"].ToString(); bdfor_StationNO = dr["StationNO"].ToString(); bdfor_ErrorType = dr["ErrorType"].ToString(); }

                                        tmp = db.DB_GetFirstDataByDataRow($@"SELECT a.PartNO,b.PartName,b.Specification FROM SoftNetSYSDB.[dbo].[APS_Simulation] as a
                                                                            join SoftNetMainDB.[dbo].[Material] as b on b.PartNO=a.PartNO
                                                                            where a.ServerId='{_Fun.Config.ServerId}' and a.SimulationId='{dr["SimulationId"].ToString()}'");
                                        if (tmp != null) { IDName = $"{tmp["PartNO"].ToString()} {tmp["PartName"].ToString()} {tmp["Specification"].ToString()}"; } else { IDName = dr["SimulationId"].ToString(); }
                                        name = "";
                                        log = "";
                                        if (!dr.IsNull("DOCNumberNO") && dr["DOCNumberNO"].ToString() != "") 
                                        {
                                            tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[DOCRole] where DOCNO='{dr["DOCNumberNO"].ToString().Substring(0,4)}' and ServerId='{_Fun.Config.ServerId}'");
                                            if (tmp != null) { name = tmp["DOCName"].ToString(); }
                                        } 
                                        switch(dr["ErrorType"].ToString())
                                        {
                                            case "05":
                                                log = $"{dr["StationNO"].ToString()}工站,生產備料無作為且無領料單據發起";
                                                break;
                                            case "01":
                                            case "02":
                                                tmp_S = "";
                                                tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                                                if (tmp != null)
                                                {
                                                    tmp_S = tmp["PartNO"].ToString();
                                                }
                                                log = $"料號:{tmp_S},庫存量不足原本預定Keep量";
                                                break;
                                            case "04":
                                                log = $"{dr["StationNO"].ToString()}工站,生產用原物料應領用而未領用";
                                                break;
                                            case "06":
                                            case "07":
                                            case "08":
                                            case "09":
                                            case "10":
                                                log = $"未如期完成確認扣(入)帳";
                                                break;
                                            case "03":
                                                log = $"{dr["StationNO"].ToString()}工站,未如預期開工";
                                                break;
                                            case "12":
                                                log = $"應執行關閉結束未關";
                                                break;
                                            case "15":
                                                log = $"{dr["StationNO"].ToString()}工站,報工作業可能有延遲";
                                                break;
                                            default:
                                                log = "未知的錯誤";
                                                break;
                                        }
                                        tmp_dt.Rows.Add(dr["LogDate"].ToString(), IDName,  dr["DOCNumberNO"].ToString(),name, log);
                                    }
                                }
                            }
                            break;
                        case "Read1"://工站角度看效益(1週) 學習模式
                            {
                                DateTime dtime = DateTime.Now;
                                tmp_dt.Columns.Add("_Fun");
                                tmp_dt.Columns.Add("StationNO");
                                tmp_dt.Columns.Add("StandardTime");//最大負荷秒時數
                                tmp_dt.Columns.Add("SIDTime");//實際稼動率
                                tmp_dt.Columns.Add("B1Time");//前1實際工時
                                tmp_dt.Columns.Add("T1Time");//前1稼動率
                                tmp_dt.Columns.Add("B2Time");
                                tmp_dt.Columns.Add("T2Time");
                                tmp_dt.Columns.Add("B3Time");
                                tmp_dt.Columns.Add("T3Time");
                                tmp_dt.Columns.Add("B4Time");
                                tmp_dt.Columns.Add("T4Time");
                                tmp_dt.Columns.Add("B5Time");
                                tmp_dt.Columns.Add("T5Time");
                                tmp_dt.Columns.Add("B6Time");
                                tmp_dt.Columns.Add("T6Time");
                                tmp_dt.Columns.Add("aaa");
                                tmp_dt.Columns.Add("bbb");
                                DataTable tmp_dt2 = null;
                                DataRow rmp_dr2 = null;


                                DataTable dtdata = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' order by FactoryName,LineName,StationNO");
                                if (dtdata != null && dtdata.Rows.Count > 0)
                                {
                                    DateTime sTime = dtime.AddDays(-7);

                                    #region 變數
                                    DateTime logTime = sTime;
                                    DateTime etime = sTime;
                                    DateTime stime = sTime;
                                    DataTable tmp_dt3 = null;
                                    float standardTime = 0;
                                    float standardTime_log = 0;
                                    float b1Time = 0;
                                    float b2Time = 0;
                                    float b3Time = 0;
                                    float b4Time = 0;
                                    float b5Time = 0;
                                    float b6Time = 0;
                                    int t1Time = 0;
                                    int t2Time = 0;
                                    int t3Time = 0;
                                    int t4Time = 0;
                                    int t5Time = 0;
                                    int t6Time = 0;
                                    string _B0 = "";
                                    string _T0 = "";
                                    string _T1 = "";
                                    string _T2 = "";
                                    string _T3 = "";
                                    string _T4 = "";
                                    string _T5 = "";
                                    string _T6 = "";
                                    string _B1 = "";
                                    string _B2 = "";
                                    string _B3 = "";
                                    string _B4 = "";
                                    string _B5 = "";
                                    string _B6 = "";
                                    string _AAA = "無數據";
                                    string _BBB = "無數據";
                                    #endregion

                                    int iCount = -1;
                                    bool isDisplay = false;
                                    List<string> returnViewBag = new List<string>();
                                    bool isFirst = true;
                                    foreach (DataRow dr in dtdata.Rows)
                                    {
                                        ++iCount;
                                        isDisplay = false;
                                        if (iCount >= dt.start && (dt.start + dt.length) > iCount)
                                        {
                                            isDisplay = true;
                                            #region 清除變數
                                            standardTime = 0;
                                            standardTime_log = 0;
                                            b1Time = 0;
                                            b2Time = 0;
                                            b3Time = 0;
                                            b4Time = 0;
                                            b5Time = 0;
                                            b6Time = 0;
                                            t1Time = 0;
                                            t2Time = 0;
                                            t3Time = 0;
                                            t4Time = 0;
                                            t5Time = 0;
                                            t6Time = 0;
                                            _B0 = "";
                                            _T0 = "0%";
                                            _T1 = "";
                                            _T2 = "";
                                            _T3 = "";
                                            _T4 = "";
                                            _T5 = "";
                                            _T6 = "";
                                            _B1 = "0%";
                                            _B2 = "0%";
                                            _B3 = "0%";
                                            _B4 = "0%";
                                            _B5 = "0%";
                                            _B6 = "0%";
                                            _AAA = "無數據";
                                            _BBB = "無數據";
                                            #endregion

                                            if (isFirst) { sql = $"SELECT top 6 * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday>='{sTime.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr["CalendarName"].ToString()}' order by CalendarName,Holiday"; }
                                            else { sql = $"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday>='{returnViewBag[0]}' and Holiday<='{returnViewBag[5]}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr["CalendarName"].ToString()}' order by CalendarName,Holiday"; }
                                            tmp_dt2 = db.DB_GetData(sql);
                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                            {
                                                if (isFirst)
                                                {
                                                    foreach (DataRow d in tmp_dt2.Rows)
                                                    {
                                                        returnViewBag.Add(Convert.ToDateTime(d["Holiday"]).ToString("yyyy/MM/dd"));
                                                    }
                                                    if (returnViewBag.Count != 6) { return re; }
                                                    isFirst = false;
                                                }
                                                for (int i = 5; i >= 0; --i)
                                                {
                                                    DataRow dr_StandardTime = tmp_dt2.Rows[i];
                                                    logTime = Convert.ToDateTime(dr_StandardTime["Holiday"]);
                                                    if (returnViewBag[i] != logTime.ToString("yyyy/MM/dd"))
                                                    { break; }

                                                    #region 取得每日最大負荷秒時數工作時間 standardTime
                                                    if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        string[] comp = dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',');
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        string[] comp = dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',');
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        string[] comp = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                                        else
                                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        bool be_addDay = false;
                                                        string[] comp_Night = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                                        string[] comp = dr_StandardTime["Shift_Graveyard"].ToString().Trim().Split(',');
                                                        if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                                                        { be_addDay = true; }
                                                        if (be_addDay || int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0).AddDays(1); }
                                                        else { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0); }
                                                        if (be_addDay)
                                                        { stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0).AddDays(1); }
                                                        else
                                                        { stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0); }
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    #endregion

                                                    #region 紀錄每日 標準工作時間 例: b(n)Time
                                                    switch (i)
                                                    {
                                                        case 5: b6Time = (standardTime - standardTime_log); break;
                                                        case 4: b5Time = (standardTime - standardTime_log); break;
                                                        case 3: b4Time = (standardTime - standardTime_log); break;
                                                        case 2: b3Time = (standardTime - standardTime_log); break;
                                                        case 1: b2Time = (standardTime - standardTime_log); break;
                                                        case 0: b1Time = (standardTime - standardTime_log); break;
                                                    }
                                                    standardTime_log = standardTime;
                                                    #endregion

                                                    #region 取得每日 實際工作時間 例: t(n)Time
                                                    //###??? 有跨日按start的問題, 智慧模式少IndexSN條件
                                                    tmp_dt3 = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTime, 111)='{logTime.ToString("yyyy/MM/dd")}' order by LOGDateTime");
                                                    if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                                    {
                                                        bool run = true;
                                                        foreach (DataRow dr3 in tmp_dt3.Rows)
                                                        {
                                                            if (dr3["OperateType"].ToString().IndexOf("開工") > 0)
                                                            {
                                                                if (run)
                                                                {
                                                                    stime = Convert.ToDateTime(dr3["LOGDateTime"]);
                                                                    run = false;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (!run)
                                                                {
                                                                    if (dr3["OperateType"].ToString().IndexOf("停工") > 0 || dr3["OperateType"].ToString().IndexOf("關站") > 0)
                                                                    {
                                                                        etime = Convert.ToDateTime(dr3["LOGDateTime"]);
                                                                        switch (i)
                                                                        {
                                                                            case 5: t6Time += TimeCompute2Seconds(stime, etime); break;
                                                                            case 4: t5Time += TimeCompute2Seconds(stime, etime); break;
                                                                            case 3: t4Time += TimeCompute2Seconds(stime, etime); break;
                                                                            case 2: t3Time += TimeCompute2Seconds(stime, etime); break;
                                                                            case 1: t2Time += TimeCompute2Seconds(stime, etime); break;
                                                                            case 0: t1Time += TimeCompute2Seconds(stime, etime); break;
                                                                        }
                                                                        run = true;
                                                                    }
                                                                }
                                                            }

                                                        }

                                                    }
                                                    #endregion
                                                }

                                                #region 取得 生產能量=有效CT/實際CT  可提升率=最佳CT/實際CT
                                                rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select (sum(OKQTY*CycleTime)/sum(OKQTY)) as ACT,(sum(OKQTY*EfficientCycleTime)/sum(OKQTY)) as BCT,(sum(OKQTY*Custom_SD_LowerLimit)/sum(OKQTY)) as CCT from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] 
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTime, 111)>='{returnViewBag[0]}' and CONVERT(varchar(100), LOGDateTime, 111)<='{returnViewBag[5]}' and CycleTime!=0 and EfficientCycleTime!=0 and Custom_SD_LowerLimit!=0 and (Custom_SD_LowerLimit*0.5)< CycleTime");
                                                if (rmp_dr2 != null)
                                                {
                                                    if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("BCT") && rmp_dr2["ACT"].ToString().Trim() != "" && rmp_dr2["BCT"].ToString().Trim() != "") { _AAA = $"{((float.Parse(rmp_dr2["BCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100).ToString("0.00")}%"; }
                                                    if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("CCT") && rmp_dr2["ACT"].ToString().Trim() != "" && rmp_dr2["CCT"].ToString().Trim() != "") { _BBB = $"{((float.Parse(rmp_dr2["CCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100).ToString("0.00")}%"; }
                                                    if (_AAA == _BBB) { _BBB = "無數據"; }
                                                }
                                                #endregion

                                                #region 顯示值
                                                TimeSpan standardTime_DIS = new TimeSpan(0, 0, (int)standardTime);
                                                _B0 = $"{(int)standardTime_DIS.TotalHours}:{standardTime_DIS.Minutes}";
                                                if ((t6Time + t5Time + t4Time + t3Time + t2Time + t1Time) > 0)
                                                { _T0 = $"{(((t6Time + t5Time + t4Time + t3Time + t2Time + t1Time) / standardTime) * 100).ToString("0.00")}%"; }
                                                standardTime_DIS = new TimeSpan(0, 0, t6Time);
                                                _T6 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";
                                                standardTime_DIS = new TimeSpan(0, 0, t5Time);
                                                _T5 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";
                                                standardTime_DIS = new TimeSpan(0, 0, t4Time);
                                                _T4 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";
                                                standardTime_DIS = new TimeSpan(0, 0, t3Time);
                                                _T3 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";
                                                standardTime_DIS = new TimeSpan(0, 0, t2Time);
                                                _T2 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";
                                                standardTime_DIS = new TimeSpan(0, 0, t1Time);
                                                _T1 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";

                                                if (t6Time > 0 && t6Time > 0) { _B6 = $"{((t6Time / b6Time) * 100).ToString("0.00")}%"; }
                                                if (t5Time > 0 && t5Time > 0) { _B5 = $"{((t5Time / b5Time) * 100).ToString("0.00")}%"; }
                                                if (t4Time > 0 && t4Time > 0) { _B4 = $"{((t4Time / b4Time) * 100).ToString("0.00")}%"; }
                                                if (t3Time > 0 && t3Time > 0) { _B3 = $"{((t3Time / b3Time) * 100).ToString("0.00")}%"; }
                                                if (t2Time > 0 && t2Time > 0) { _B2 = $"{((t2Time / b2Time) * 100).ToString("0.00")}%"; }
                                                if (t1Time > 0 && t1Time > 0) { _B1 = $"{((t1Time / b1Time) * 100).ToString("0.00")}%"; }
                                                #endregion
                                            }
                                        }
                                        if (isDisplay)
                                        { tmp_dt.Rows.Add("", dr["StationNO"].ToString(), _B0, _T0,  _T6, _B6, _T5, _B5, _T4, _B4, _T3, _B3, _T2, _B2, _T1, _B1, _AAA, _BBB); }
                                        else
                                        { tmp_dt.Rows.Add("", dr["StationNO"].ToString(), "", "",  "", "", "", "", "", "", "", "", "", "", "", "", "", ""); }

                                    }
                                }
                            }
                            break;
                        case "Read2"://工站角度看效益(4週) 學習模式
                            {
                                //明細,工站, 總成長率,總稼動率,總能量率,最大負荷,週成長率,週稼動率,生產能量,   ,週成長率,週稼動率,生產能量
                                tmp_dt.Columns.Add("_Fun");
                                tmp_dt.Columns.Add("StationNO");
                                tmp_dt.Columns.Add("TOTGrowing");//總成長率
                                tmp_dt.Columns.Add("TOTSIDTime");//總稼動率
                                tmp_dt.Columns.Add("TOTATime");//總能量率
                                tmp_dt.Columns.Add("TOTBTime");//最大負荷
                                tmp_dt.Columns.Add("C1Time");//週成長率
                                tmp_dt.Columns.Add("D1Time");//週稼動率
                                tmp_dt.Columns.Add("A1Time");//生產能量
                                tmp_dt.Columns.Add("C2Time");//週成長率
                                tmp_dt.Columns.Add("D2Time");//週稼動率
                                tmp_dt.Columns.Add("A2Time");//生產能量
                                tmp_dt.Columns.Add("C3Time");//週成長率
                                tmp_dt.Columns.Add("D3Time");//週稼動率
                                tmp_dt.Columns.Add("A3Time");//生產能量
                                tmp_dt.Columns.Add("C4Time");//週成長率
                                tmp_dt.Columns.Add("D4Time");//週稼動率
                                tmp_dt.Columns.Add("A4Time");//生產能量


                                DataTable tmp_dt2 = null;
                                DataRow rmp_dr2 = null;

                                DataTable dtdata = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' order by FactoryName,LineName,StationNO");
                                if (dtdata != null && dtdata.Rows.Count > 0)
                                {
                                    #region 變數
                                    DateTime logTime = DateTime.Now;
                                    DateTime tmp_sDate = DateTime.Now;
                                    DateTime tmp_eDate = DateTime.Now;
                                    DataTable tmp_dt3 = null;

                                    string TOTGrowing = "";
                                    string TOTSIDTime = "";
                                    string TOTATime = "";
                                    string TOTBTime = "";
                                    float C1Time = 0.0f;
                                    float D1Time = 0.0f;
                                    float A1Time = 0.0f;
                                    float C2Time = 0.0f;
                                    float D2Time = 0.0f;
                                    float A2Time = 0.0f;
                                    float C3Time = 0.0f;
                                    float D3Time = 0.0f;
                                    float A3Time = 0.0f;
                                    float C4Time = 0.0f;
                                    float D4Time = 0.0f;
                                    float A4Time = 0.0f;

                                    float standardTime = 0;
                                    float standardTime_log = 0;
                                    float totTTime = 0;
                                    float t1Time_log = 0.0f;

                                    float totAAA = 0;
                                    float totBBB = 0;
                                    float aaa_log = 0;

                                    #endregion

                                    int iCount = -1;
                                    bool isDisplay = false;
                                    float tmp = 0;
                                    foreach (DataRow dr in dtdata.Rows)
                                    {
                                        ++iCount;
                                        isDisplay = false;
                                        if (iCount >= dt.start && (dt.start + dt.length) > iCount)
                                        {
                                            isDisplay = true;
                                            int tmp_week = (int)DateTime.Now.DayOfWeek;
                                            DateTime dtime = DateTime.Now.AddDays(-tmp_week-1);
                                            DateTime stime;
                                            standardTime = 0;
                                            standardTime_log = 0;
                                            totTTime = 0;
                                            t1Time_log = 0;

                                            for (int i = 1; i <= 5; ++i)
                                            {
                                                stime = dtime.AddDays(-6);
                                                if (i <= 4)
                                                {
                                                    #region 取得週標準工作時間 standardTime
                                                    tmp_dt2 = db.DB_GetData($"SELECT  * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday>='{stime.ToString("yyyy/MM/dd")}' and Holiday<='{dtime.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr["CalendarName"].ToString()}' order by CalendarName,Holiday");
                                                    if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                    {
                                                        foreach (DataRow dr_StandardTime in tmp_dt2.Rows)
                                                        {
                                                            logTime = Convert.ToDateTime(dr_StandardTime["Holiday"]);

                                                            if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                                                            {
                                                                #region 取得工作時間
                                                                tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                                                tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                                tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                                tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                                #endregion
                                                            }
                                                            if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                                                            {
                                                                #region 取得工作時間
                                                                tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                                                tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                                tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                                tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                                #endregion
                                                            }
                                                            if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                                                            {
                                                                #region 取得工作時間
                                                                tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                                                tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                                tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                                tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                                #endregion
                                                            }
                                                            if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) == true)
                                                            {
                                                                #region 取得工作時間
                                                                tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                                                tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                                tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                                tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                                #endregion
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                    standardTime_log = standardTime - standardTime_log;

                                                    #region 取得週 實際工作時間 totTTime
                                                    //###??? 有跨日按start的問題
                                                    tmp_dt3 = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTime, 111)>='{stime.ToString("yyyy/MM/dd")}' and CONVERT(varchar(100), LOGDateTime, 111)<='{dtime.ToString("yyyy/MM/dd")}' order by LOGDateTime");
                                                    if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                                    {
                                                        bool run = true;
                                                        foreach (DataRow dr3 in tmp_dt3.Rows)
                                                        {
                                                            if (dr3["OperateType"].ToString().IndexOf("開工") > 0)
                                                            {
                                                                if (run)
                                                                {
                                                                    tmp_sDate = Convert.ToDateTime(dr3["LOGDateTime"]);
                                                                    run = false;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (!run)
                                                                {
                                                                    if (dr3["OperateType"].ToString().IndexOf("停工") > 0 || dr3["OperateType"].ToString().IndexOf("關站") > 0)
                                                                    {
                                                                        tmp_eDate = Convert.ToDateTime(dr3["LOGDateTime"]);
                                                                        totTTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                                        run = true;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }

                                                    #endregion
                                                    t1Time_log = totTTime - t1Time_log;
                                                }
                                                #region 取得週 aaa生產能量=有效CT/實際CT bbb可提升率=最佳CT/實際CT
                                                rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select (sum(OKQTY*CycleTime)/sum(OKQTY)) as ACT,(sum(OKQTY*EfficientCycleTime)/sum(OKQTY)) as BCT,(sum(OKQTY*Custom_SD_LowerLimit)/sum(OKQTY)) as CCT from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] 
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTime, 111)>='{stime.ToString("yyyy/MM/dd")}' and CONVERT(varchar(100), LOGDateTime, 111)<='{dtime.ToString("yyyy/MM/dd")}' and CycleTime!=0 and EfficientCycleTime!=0 and Custom_SD_LowerLimit!=0 and (Custom_SD_LowerLimit*0.5)< CycleTime");
                                                if (rmp_dr2 != null)
                                                {
                                                    if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("BCT")) { totAAA += (float.Parse(rmp_dr2["BCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100; }
                                                    if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("CCT")) { totBBB += (float.Parse(rmp_dr2["CCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100; }
                                                }
                                                #endregion
                                                aaa_log = totAAA - aaa_log;

                                                switch (i)
                                                {
                                                    case 1:
                                                        D1Time = (t1Time_log / standardTime_log) * 100;//嫁動率
                                                        A1Time = aaa_log;//生產能量
                                                        break;
                                                    case 2:
                                                        D2Time = (t1Time_log / standardTime_log) * 100;//嫁動率
                                                        A2Time = aaa_log;//生產能量
                                                        if (aaa_log != 0 && A1Time != 0)
                                                        { C1Time = aaa_log / A1Time; }
                                                        break;
                                                    case 3:
                                                        D3Time = (t1Time_log / standardTime_log) * 100;//嫁動率
                                                        A3Time = aaa_log;//生產能量
                                                        if (aaa_log != 0 && A2Time != 0)
                                                        { C2Time = aaa_log / A2Time; }
                                                        break;
                                                    case 4:
                                                        D4Time = (t1Time_log / standardTime_log) * 100;//嫁動率
                                                        A4Time = aaa_log;//生產能量
                                                        if (aaa_log != 0 && A3Time != 0)
                                                        { C3Time = aaa_log / A3Time; }
                                                        break;
                                                    case 5:
                                                        if (aaa_log != 0 && A4Time != 0)
                                                        { C4Time = aaa_log / A4Time; }
                                                        break;
                                                }
                                                dtime = stime.AddDays(-1);
                                            }
                                            #region 顯示值
                                            tmp = C1Time + C2Time + C3Time + C4Time;
                                            TOTGrowing = tmp == 0 ? "0%" : $"{((C1Time + C2Time + C3Time + C4Time) / 4).ToString("0.00")}%";//總成長率
                                            TOTSIDTime = $"{((totTTime / standardTime) * 100).ToString("0.00")}%";//總嫁動率
                                            TOTATime = totAAA == 0 ? "0%" : $"{(totAAA / 4).ToString("0.00")}%";//總生產能量
                                            TOTBTime = totBBB == 0 ? "0%" : $"{(totBBB / 4).ToString("0.00")}%";//總最大負荷
                                            #endregion
                                        }



                                        //明細,工站, 總成長率,總稼動率,總能量率,最大負荷,週成長率,週稼動率,生產能量,   ,週成長率,週稼動率,生產能量
                                        if (isDisplay)
                                        { tmp_dt.Rows.Add("", dr["StationNO"].ToString(), TOTGrowing, TOTSIDTime, TOTATime, TOTBTime, $"{C1Time.ToString("0.00")}%", $"{D1Time.ToString("0.00")}%", $"{A1Time.ToString("0.00")}%", $"{C2Time.ToString("0.00")}%", $"{D2Time.ToString("0.00")}%", $"{A2Time.ToString("0.00")}%", $"{C3Time.ToString("0.00")}%", $"{D3Time.ToString("0.00")}%", $"{A3Time.ToString("0.00")}%", $"{C4Time.ToString("0.00")}%", $"{D4Time.ToString("0.00")}%", $"{A4Time.ToString("0.00")}%"); }
                                        else
                                        { tmp_dt.Rows.Add("", dr["StationNO"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""); }

                                    }
                                }
                            }
                            break;
                        case "Read3"://工站角度看效益(1週) 智慧模式
                            {
                                #region 宣告變數
                                DateTime dtime = DateTime.Now;
                                tmp_dt.Columns.Add("_Fun");
                                tmp_dt.Columns.Add("StationNO");
                                tmp_dt.Columns.Add("StandardTime");//最大負荷秒時數
                                tmp_dt.Columns.Add("SIDTime");//實際稼動率
                                tmp_dt.Columns.Add("B1Time");//前1實際工時
                                tmp_dt.Columns.Add("T1Time");//前1稼動率
                                tmp_dt.Columns.Add("B2Time");
                                tmp_dt.Columns.Add("T2Time");
                                tmp_dt.Columns.Add("B3Time");
                                tmp_dt.Columns.Add("T3Time");
                                tmp_dt.Columns.Add("B4Time");
                                tmp_dt.Columns.Add("T4Time");
                                tmp_dt.Columns.Add("B5Time");
                                tmp_dt.Columns.Add("T5Time");
                                tmp_dt.Columns.Add("B6Time");
                                tmp_dt.Columns.Add("T6Time");
                                tmp_dt.Columns.Add("aaa");
                                tmp_dt.Columns.Add("bbb");
                                DataTable tmp_dt2 = null;
                                DataRow rmp_dr2 = null;
                                #endregion


                                DataTable dtdata = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' order by FactoryName,LineName,StationNO");
                                if (dtdata != null && dtdata.Rows.Count > 0)
                                {
                                    DateTime sTime = dtime.AddDays(-1);

                                    #region 變數
                                    DateTime logTime = sTime;
                                    DateTime etime = sTime;
                                    DateTime stime = sTime;
                                    DataTable tmp_dt3 = null;
                                    float standardTime = 0;
                                    float standardTime_log = 0;
                                    float b1Time = 0;
                                    float b2Time = 0;
                                    float b3Time = 0;
                                    float b4Time = 0;
                                    float b5Time = 0;
                                    float b6Time = 0;
                                    int t1Time = 0;
                                    int t2Time = 0;
                                    int t3Time = 0;
                                    int t4Time = 0;
                                    int t5Time = 0;
                                    int t6Time = 0;
                                    string _B0 = "";
                                    string _T0 = "";
                                    string _T1 = "";
                                    string _T2 = "";
                                    string _T3 = "";
                                    string _T4 = "";
                                    string _T5 = "";
                                    string _T6 = "";
                                    string _B1 = "";
                                    string _B2 = "";
                                    string _B3 = "";
                                    string _B4 = "";
                                    string _B5 = "";
                                    string _B6 = "";
                                    string _AAA = "無數據";
                                    string _BBB = "無數據";
                                    #endregion

                                    int iCount = -1;
                                    bool isDisplay = false;
                                    List<string> returnViewBag = new List<string>();
                                    bool isFirst = true;
                                    foreach (DataRow dr in dtdata.Rows)
                                    {
                                        ++iCount;
                                        isDisplay = false;
                                        if (iCount >= dt.start && (dt.start + dt.length) > iCount)
                                        {
                                            isDisplay = true;
                                            #region 清除變數
                                            standardTime = 0;
                                            standardTime_log = 0;
                                            b1Time = 0;
                                            b2Time = 0;
                                            b3Time = 0;
                                            b4Time = 0;
                                            b5Time = 0;
                                            b6Time = 0;
                                            t1Time = 0;
                                            t2Time = 0;
                                            t3Time = 0;
                                            t4Time = 0;
                                            t5Time = 0;
                                            t6Time = 0;
                                            _B0 = "";
                                            _T0 = "0%";
                                            _T1 = "";
                                            _T2 = "";
                                            _T3 = "";
                                            _T4 = "";
                                            _T5 = "";
                                            _T6 = "";
                                            _B1 = "0%";
                                            _B2 = "0%";
                                            _B3 = "0%";
                                            _B4 = "0%";
                                            _B5 = "0%";
                                            _B6 = "0%";
                                            _AAA = "無數據";
                                            _BBB = "無數據";
                                            #endregion

                                            if (isFirst) { sql = $"SELECT top 6 * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday<='{sTime.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr["CalendarName"].ToString()}' order by CalendarName,Holiday desc"; }
                                            else { sql = $"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday>='{returnViewBag[5]}' and Holiday<='{returnViewBag[0]}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr["CalendarName"].ToString()}' order by CalendarName,Holiday desc"; }
                                            tmp_dt2 = db.DB_GetData(sql);
                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                            {
                                                if (isFirst)
                                                {
                                                    foreach (DataRow d in tmp_dt2.Rows)
                                                    {
                                                        returnViewBag.Add(Convert.ToDateTime(d["Holiday"]).ToString("yyyy/MM/dd"));
                                                    }
                                                    if (returnViewBag.Count != 6) { return re; }
                                                    isFirst = false;
                                                }
                                                for (int i = 5; i >= 0; --i)
                                                {
                                                    DataRow dr_StandardTime = tmp_dt2.Rows[i];
                                                    logTime = Convert.ToDateTime(dr_StandardTime["Holiday"]);
                                                    if (returnViewBag[i] != logTime.ToString("yyyy/MM/dd"))
                                                    {
                                                        break;
                                                    }

                                                    if (dr["StationNO"].ToString() == "F01")
                                                    {
                                                        string _s = "";
                                                    }

                                                    #region 紀錄每日 標準工作時間 b(n))Time
                                                    if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        string[] comp = dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',');
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        string[] comp = dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',');
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        string[] comp = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                                        else
                                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        bool be_addDay = false;
                                                        string[] comp_Night = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                                        string[] comp = dr_StandardTime["Shift_Graveyard"].ToString().Trim().Split(',');
                                                        if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                                                        { be_addDay = true; }
                                                        if (be_addDay || int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0).AddDays(1); }
                                                        else { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0); }
                                                        if (be_addDay)
                                                        { stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0).AddDays(1); }
                                                        else
                                                        { stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0); }
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    switch (i)
                                                    {
                                                        case 5: b6Time = (standardTime - standardTime_log); break;
                                                        case 4: b5Time = (standardTime - standardTime_log); break;
                                                        case 3: b4Time = (standardTime - standardTime_log); break;
                                                        case 2: b3Time = (standardTime - standardTime_log); break;
                                                        case 1: b2Time = (standardTime - standardTime_log); break;
                                                        case 0: b1Time = (standardTime - standardTime_log); break;
                                                    }
                                                    standardTime_log = standardTime;
                                                    #endregion

                                                    #region 取得每日 實際工作時間 t(n))Time
                                                    Dictionary<DateTime, DateTime> woItemList = new Dictionary<DateTime, DateTime>();
                                                    tmp_dt3 = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTime, 111)='{logTime.ToString("yyyy/MM/dd")}' and OperateType like '%開工%' order by LOGDateTime");
                                                    if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                                    {
                                                        DataRow tmp = null;
                                                        bool is_back = false;
                                                        foreach (DataRow dr3 in tmp_dt3.Rows)
                                                        {
                                                            is_back = false;
                                                            tmp = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and LOGDateTime>='{Convert.ToDateTime(dr3["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and OrderNO='{dr3["OrderNO"].ToString()}' and OperateType like '%停工%' order by LOGDateTime");
                                                            if (tmp != null)
                                                            {
                                                                sTime = Convert.ToDateTime(dr3["LOGDateTime"]);//工單開始時間
                                                                logTime = Convert.ToDateTime(tmp["LOGDateTime"]);//工單結束時間

                                                                #region 不該的跨天
                                                                if (sTime.Day!= logTime.Day)
                                                                {
                                                                    if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) != true || bool.Parse(dr_StandardTime["Flag_Night"].ToString()) != true || bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) != true || bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) != true)
                                                                    { continue; }

                                                                    /*
                                                                    if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) != true)
                                                                    {
                                                                        bool chrun = true;
                                                                        DateTime chtime = DateTime.Now;
                                                                        if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                                                                        { chtime = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), sTime.Second, sTime.Millisecond); }
                                                                        if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                                                                        { chtime = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), sTime.Second, sTime.Millisecond); }
                                                                        if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                                                                        {
                                                                            chtime = new DateTime(sTime.Year, sTime.Month, sTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), sTime.Second, sTime.Millisecond);
                                                                            if (sTime > chtime) { chrun = false; }
                                                                        }
                                                                        if (chtime > sTime && chrun)
                                                                        { 
                                                                            continue; 
                                                                        }
                                                                    }
                                                                    */
                                                                }
                                                                #endregion

                                                                #region 判斷重複的時間
                                                                foreach (KeyValuePair<DateTime, DateTime> d in woItemList)
                                                                {
                                                                    if (sTime >= d.Key && logTime <= d.Value)//時間內
                                                                    { is_back = true; break; }
                                                                    else if (sTime >= d.Key && logTime >= d.Value)//時間後段外
                                                                    {
                                                                        if (sTime <= d.Value)
                                                                        { sTime = d.Value; }
                                                                    }
                                                                    else if (sTime <= d.Key && logTime <= d.Value)//時間前段外
                                                                    { 
                                                                        
                                                                        if (logTime <= d.Value)
                                                                        { logTime = d.Key; }
                                                                    }
                                                                }
                                                                if (is_back) { continue; }
                                                                #endregion

                                                                switch (i)
                                                                {
                                                                    case 5: t6Time += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), sTime, logTime); break;
                                                                    case 4: t5Time += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), sTime, logTime); break;
                                                                    case 3: t4Time += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), sTime, logTime); break;
                                                                    case 2: t3Time += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), sTime, logTime); break;
                                                                    case 1: t2Time += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), sTime, logTime); break;
                                                                    case 0: t1Time += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), sTime, logTime); break;
                                                                }
                                                                woItemList.Add(sTime, logTime);
                                                                string _s = "";
                                                            }
                                                        }
                                                    }
                                                    #endregion

                                                    #region 取得 生產能量AAA=有效CT/實際CT  可提升率BBB=最佳CT/實際CT
                                                    rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select (sum((EditFinishedQty+EditFailedQty)*CycleTime)/sum((EditFinishedQty+EditFailedQty))) as ACT,(sum((EditFinishedQty+EditFailedQty)*ECT)/sum((EditFinishedQty+EditFailedQty))) as BCT,(sum((EditFinishedQty+EditFailedQty)*LowerCT)/sum((EditFinishedQty+EditFailedQty))) as CCT from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] 
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTimeID, 111)>='{returnViewBag[5]}' and CONVERT(varchar(100), LOGDateTimeID, 111)<='{returnViewBag[0]}' and CycleTime!=0 and ECT!=0 and LowerCT!=0 and (LowerCT*0.5)< CycleTime");
                                                    if (rmp_dr2 != null)
                                                    {
                                                        if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("BCT") && rmp_dr2["ACT"].ToString().Trim() != "" && rmp_dr2["BCT"].ToString().Trim() != "") { _AAA = $"{((float.Parse(rmp_dr2["BCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100).ToString("0.00")}%"; }
                                                        if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("CCT") && rmp_dr2["ACT"].ToString().Trim() != "" && rmp_dr2["CCT"].ToString().Trim() != "") { _BBB = $"{((float.Parse(rmp_dr2["CCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100).ToString("0.00")}%"; }
                                                        if (_AAA == _BBB) { _BBB = "無數據"; }
                                                    }
                                                    #endregion

                                                    #region 顯示值
                                                    TimeSpan standardTime_DIS = new TimeSpan(0, 0, (int)standardTime);
                                                    _B0 = $"{(int)standardTime_DIS.TotalHours}:{standardTime_DIS.Minutes}";
                                                    if ((t6Time + t5Time + t4Time + t3Time + t2Time + t1Time) > 0)
                                                    { _T0 = $"{(((t6Time + t5Time + t4Time + t3Time + t2Time + t1Time) / standardTime) * 100).ToString("0.00")}%"; }

                                                    standardTime_DIS = new TimeSpan(0, 0, t6Time);
                                                    _T6 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";
                                                    standardTime_DIS = new TimeSpan(0, 0, t5Time);
                                                    _T5 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";
                                                    standardTime_DIS = new TimeSpan(0, 0, t4Time);
                                                    _T4 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";
                                                    standardTime_DIS = new TimeSpan(0, 0, t3Time);
                                                    _T3 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";
                                                    standardTime_DIS = new TimeSpan(0, 0, t2Time);
                                                    _T2 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";
                                                    standardTime_DIS = new TimeSpan(0, 0, t1Time);
                                                    _T1 = $"{standardTime_DIS.Hours}:{(standardTime_DIS.Minutes < 10 ? $"0{standardTime_DIS.Minutes}" : standardTime_DIS.Minutes)}";

                                                    if (t6Time > 0 && t6Time > 0) { _B6 = $"{((t6Time / b6Time) * 100).ToString("0.00")}%"; }
                                                    if (t5Time > 0 && t5Time > 0) { _B5 = $"{((t5Time / b5Time) * 100).ToString("0.00")}%"; }
                                                    if (t4Time > 0 && t4Time > 0) { _B4 = $"{((t4Time / b4Time) * 100).ToString("0.00")}%"; }
                                                    if (t3Time > 0 && t3Time > 0) { _B3 = $"{((t3Time / b3Time) * 100).ToString("0.00")}%"; }
                                                    if (t2Time > 0 && t2Time > 0) { _B2 = $"{((t2Time / b2Time) * 100).ToString("0.00")}%"; }
                                                    if (t1Time > 0 && t1Time > 0) { _B1 = $"{((t1Time / b1Time) * 100).ToString("0.00")}%"; }
                                                    #endregion
                                                }
                                            }
                                        }
                                        if (isDisplay)
                                        { tmp_dt.Rows.Add("", dr["StationNO"].ToString(), _B0, _T0, _T6, _B6, _T5, _B5, _T4, _B4, _T3, _B3, _T2, _B2, _T1, _B1, _AAA, _BBB); }
                                        else
                                        { tmp_dt.Rows.Add("", dr["StationNO"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""); }

                                    }
                                }
                            }
                            break;
                        case "Read4"://工站角度看效益(4週)
                            {
                                //明細,工站, 總成長率,總稼動率,總能量率,最大負荷,週成長率,週稼動率,生產能量,   ,週成長率,週稼動率,生產能量
                                #region 宣告變數
                                tmp_dt.Columns.Add("_Fun");
                                tmp_dt.Columns.Add("StationNO");
                                tmp_dt.Columns.Add("TOTGrowing");//總成長率
                                tmp_dt.Columns.Add("TOTSIDTime");//總稼動率
                                tmp_dt.Columns.Add("TOTATime");//總能量率
                                tmp_dt.Columns.Add("TOTBTime");//最大負荷
                                tmp_dt.Columns.Add("C1Time");//週成長率
                                tmp_dt.Columns.Add("D1Time");//週稼動率
                                tmp_dt.Columns.Add("A1Time");//生產能量
                                tmp_dt.Columns.Add("C2Time");//週成長率
                                tmp_dt.Columns.Add("D2Time");//週稼動率
                                tmp_dt.Columns.Add("A2Time");//生產能量
                                tmp_dt.Columns.Add("C3Time");//週成長率
                                tmp_dt.Columns.Add("D3Time");//週稼動率
                                tmp_dt.Columns.Add("A3Time");//生產能量
                                tmp_dt.Columns.Add("C4Time");//週成長率
                                tmp_dt.Columns.Add("D4Time");//週稼動率
                                tmp_dt.Columns.Add("A4Time");//生產能量
                                #endregion

                                DataTable tmp_dt2 = null;
                                DataRow rmp_dr2 = null;

                                DataTable dtdata = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' order by FactoryName,LineName,StationNO");
                                if (dtdata != null && dtdata.Rows.Count > 0)
                                {
                                    #region 變數
                                    DateTime logTime = DateTime.Now;
                                    DateTime tmp_sDate = DateTime.Now;
                                    DateTime tmp_eDate = DateTime.Now;
                                    DataTable tmp_dt3 = null;

                                    string TOTGrowing = "";
                                    string TOTSIDTime = "";
                                    string TOTATime = "";
                                    string TOTBTime = "";
                                    float C1Time = 0.0f;
                                    float D1Time = 0.0f;
                                    float A1Time = 0.0f;
                                    float C2Time = 0.0f;
                                    float D2Time = 0.0f;
                                    float A2Time = 0.0f;
                                    float C3Time = 0.0f;
                                    float D3Time = 0.0f;
                                    float A3Time = 0.0f;
                                    float C4Time = 0.0f;
                                    float D4Time = 0.0f;
                                    float A4Time = 0.0f;

                                    float standardTime = 0;
                                    float standardTime_log = 0;
                                    float totTTime = 0;
                                    float t1Time_log = 0.0f;

                                    float totAAA = 0;
                                    float totBBB = 0;
                                    float aaa_log = 0;

                                    #endregion

                                    int iCount = -1;
                                    bool isDisplay = false;
                                    float tmp = 0;
                                    Dictionary<DateTime, DateTime> woItemList = new Dictionary<DateTime, DateTime>();
                                    List<string> weekdateList=new List<string>();
                                    DataRow tmp3 = null;
                                    DateTime stime;
                                    DateTime etime;
                                    DateTime sTime;
                                    DateTime dtime;

                                    CultureInfo info = CultureInfo.CurrentCulture;
                                    int totalWeekOfYear = info.Calendar.GetWeekOfYear
                                    (
                                        DateTime.Now,
                                        CalendarWeekRule.FirstDay,
                                        DayOfWeek.Sunday
                                    );
                                    totalWeekOfYear -= 1;

                                    foreach (DataRow dr in dtdata.Rows)
                                    {
                                        ++iCount;
                                        isDisplay = false;
                                        if (iCount >= dt.start && (dt.start + dt.length) > iCount)
                                        {
                                            isDisplay = true;
                                            //int tmp_week = (int)DateTime.Now.DayOfWeek;
                                            //dtime = DateTime.Now.AddDays(-tmp_week - 1);


                                            ////先取得該年第一天的日期
                                            //DateTime firstDateOfYear = new DateTime(DateTime.Now.Year, 1, 1);
                                            ////該年第一天再加上周數乘以七
                                            //DateTime dayInWeek = firstDateOfYear.AddDays(totalWeekOfYear * 7);
                                            //DateTime firstDayInWeek = dayInWeek.Date;
                                            ////ISO 8601所制定的標準中，一週的第一天為週一
                                            //while (firstDayInWeek.DayOfWeek != DayOfWeek.Monday)
                                            //{
                                            //    firstDayInWeek = firstDayInWeek.AddDays(-1);
                                            //}

                                            //dtime = firstDayInWeek.AddDays(-1);






                                            standardTime = 0;
                                            standardTime_log = 0;
                                            totTTime = 0;
                                            t1Time_log = 0;
                                            totAAA = 0;
                                            totBBB = 0;

                                            for (int i = 1; i <= 5; ++i)
                                            {
                                                weekdateList.Clear();
                                                //stime = dtime.AddDays(-7);
                                                stime = _SFC_Common.GetYearOfWeekSunDay((totalWeekOfYear - i));
                                                dtime = stime.AddDays(7);


                                                if (i <= 4)
                                                {
                                                    DataRow dr_lastTime = null;
                                                    #region 取得週標準工作時間 standardTime
                                                    tmp_dt2 = db.DB_GetData($"SELECT  * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday>'{stime.ToString("yyyy/MM/dd")}' and Holiday<='{dtime.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr["CalendarName"].ToString()}' order by CalendarName,Holiday");
                                                    if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                                    {
                                                        dr_lastTime = tmp_dt2.Rows[(tmp_dt2.Rows.Count-1)];
                                                        foreach (DataRow dr_StandardTime in tmp_dt2.Rows)
                                                        {
                                                            logTime = Convert.ToDateTime(dr_StandardTime["Holiday"]);
                                                            weekdateList.Add(logTime.ToString("yyyy/MM/dd"));

                                                            if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                                                            {
                                                                #region 取得工作時間
                                                                string[] comp = dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',');
                                                                etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                                stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(stime, etime);
                                                                etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                                stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(stime, etime);
                                                                #endregion
                                                            }
                                                            if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                                                            {
                                                                #region 取得工作時間
                                                                string[] comp = dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',');
                                                                etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                                stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(stime, etime);
                                                                etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                                stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(stime, etime);
                                                                #endregion
                                                            }
                                                            if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                                                            {
                                                                #region 取得工作時間
                                                                string[] comp = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                                                etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                                stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(stime, etime);
                                                                if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                                { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                                                else
                                                                { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                                                stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(stime, etime);
                                                                #endregion
                                                            }
                                                            if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) == true)
                                                            {
                                                                #region 取得工作時間
                                                                bool be_addDay = false;
                                                                string[] comp_Night = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                                                string[] comp = dr_StandardTime["Shift_Graveyard"].ToString().Trim().Split(',');
                                                                if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                                                                { be_addDay = true; }
                                                                if (be_addDay || int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                                                                { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0).AddDays(1); }
                                                                else { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0); }
                                                                if (be_addDay)
                                                                { stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0).AddDays(1); }
                                                                else
                                                                { stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0); }
                                                                standardTime += TimeCompute2Seconds(stime, etime);
                                                                etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                                stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                                standardTime += TimeCompute2Seconds(stime, etime);
                                                                #endregion
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                    standardTime_log = standardTime - standardTime_log;
                                                    DataRow dr_HolidayCalendar = null;
                                                    #region 取得週 實際工作時間 totTTime
                                                    if (dr["StationNO"].ToString()=="C04")
                                                    {
                                                        string _s = "";
                                                    }
                                                    foreach (string s in weekdateList)
                                                    {
                                                        woItemList.Clear();
                                                        tmp_dt3 = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTime, 111)='{s}' and OperateType like '%開工%' order by LOGDateTime");
                                                        if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                                        {
                                                            bool is_back = false;
                                                            foreach (DataRow dr3 in tmp_dt3.Rows)
                                                            {
                                                                is_back = false;
                                                                tmp3 = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and LOGDateTime>='{Convert.ToDateTime(dr3["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and OrderNO='{dr3["OrderNO"].ToString()}' and OperateType like '%停工%' order by LOGDateTime");
                                                                if (tmp3 != null)
                                                                {
                                                                    sTime = Convert.ToDateTime(dr3["LOGDateTime"]);//工單開始時間
                                                                    logTime = Convert.ToDateTime(tmp3["LOGDateTime"]);//工單結束時間
                                                                    #region 不該的跨天
                                                                    if (sTime.Day != logTime.Day)
                                                                    {
                                                                        dr_HolidayCalendar = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday='{s}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr["CalendarName"].ToString()}'");
                                                                        if (bool.Parse(dr_HolidayCalendar["Flag_Graveyard"].ToString()) != true || bool.Parse(dr_HolidayCalendar["Flag_Night"].ToString()) != true || bool.Parse(dr_HolidayCalendar["Flag_Afternoon"].ToString()) != true || bool.Parse(dr_HolidayCalendar["Flag_Morning"].ToString()) != true)
                                                                        { continue; }
                                                                    }
                                                                    #endregion
                                                                    #region 判斷重複的時間
                                                                    foreach (KeyValuePair<DateTime, DateTime> d in woItemList)
                                                                    {
                                                                        if (sTime >= d.Key && logTime <= d.Value)//時間內
                                                                        { is_back = true; break; }
                                                                        else if (sTime >= d.Key && logTime >= d.Value)//時間後段外
                                                                        { sTime = d.Value; }
                                                                        else if (sTime <= d.Key && logTime <= d.Value)//時間後段外
                                                                        { logTime = d.Key; }
                                                                    }
                                                                    if (is_back) { continue; }
                                                                    #endregion

                                                                    totTTime += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), sTime, logTime);
                                                                    woItemList.Add(Convert.ToDateTime(sTime), logTime);
                                                                }
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                    t1Time_log = totTTime - t1Time_log;
                                                }

                                                #region 取得週 aaa生產能量=有效CT/實際CT bbb可提升率=最佳CT/實際CT
                                                rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select (sum((EditFinishedQty+EditFailedQty)*CycleTime)/sum((EditFinishedQty+EditFailedQty))) as ACT,(sum((EditFinishedQty+EditFailedQty)*ECT)/sum((EditFinishedQty+EditFailedQty))) as BCT,(sum((EditFinishedQty+EditFailedQty)*LowerCT)/sum((EditFinishedQty+EditFailedQty))) as CCT from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] 
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTimeID, 111)>='{stime.ToString("yyyy/MM/dd")}' and CONVERT(varchar(100), LOGDateTimeID, 111)<='{dtime.ToString("yyyy/MM/dd")}' and CycleTime!=0 and ECT!=0 and LowerCT!=0 and (LowerCT*0.5)< CycleTime");
                                                if (rmp_dr2 != null)
                                                {
                                                    if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("BCT")) { totAAA += (float.Parse(rmp_dr2["BCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100; }
                                                    if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("CCT")) { totBBB += (float.Parse(rmp_dr2["CCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100; }
                                                }
                                                #endregion
                                                aaa_log = totAAA - aaa_log;

                                                switch (i)
                                                {
                                                    case 1:
                                                        D1Time = (t1Time_log / standardTime_log) * 100;//嫁動率
                                                        A1Time = aaa_log;//生產能量
                                                        break;
                                                    case 2:
                                                        D2Time = (t1Time_log / standardTime_log) * 100;//嫁動率
                                                        A2Time = aaa_log;//生產能量
                                                        if (aaa_log != 0 && A1Time != 0)
                                                        { C1Time = aaa_log / A1Time; }
                                                        break;
                                                    case 3:
                                                        D3Time = (t1Time_log / standardTime_log) * 100;//嫁動率
                                                        A3Time = aaa_log;//生產能量
                                                        if (aaa_log != 0 && A2Time != 0)
                                                        { C2Time = aaa_log / A2Time; }
                                                        break;
                                                    case 4:
                                                        D4Time = (t1Time_log / standardTime_log) * 100;//嫁動率
                                                        A4Time = aaa_log;//生產能量
                                                        if (aaa_log != 0 && A3Time != 0)
                                                        { C3Time = aaa_log / A3Time; }
                                                        break;
                                                    case 5:
                                                        if (aaa_log != 0 && A4Time != 0)
                                                        { C4Time = aaa_log / A4Time; }
                                                        break;
                                                }
                                                dtime = stime.AddDays(-1);
                                            }
                                            #region 顯示值
                                            tmp = C1Time + C2Time + C3Time + C4Time;
                                            TOTGrowing = tmp == 0 ? "0%" : $"{((C1Time + C2Time + C3Time + C4Time) / 4).ToString("0.00")}%";//總成長率
                                            TOTSIDTime = $"{((totTTime / standardTime) * 100).ToString("0.00")}%";//總嫁動率
                                            TOTATime = totAAA == 0 ? "0%" : $"{(totAAA / 4).ToString("0.00")}%";//總生產能量
                                            TOTBTime = totBBB == 0 ? "0%" : $"{(totBBB / 4).ToString("0.00")}%";//總最大負荷
                                            #endregion
                                        }



                                        //明細,工站, 總成長率,總稼動率,總能量率,最大負荷,週成長率,週稼動率,生產能量,   ,週成長率,週稼動率,生產能量
                                        if (isDisplay)
                                        { tmp_dt.Rows.Add("", dr["StationNO"].ToString(), TOTGrowing, TOTSIDTime, TOTATime, TOTBTime, $"{C1Time.ToString("0.00")}%", $"{D1Time.ToString("0.00")}%", $"{A1Time.ToString("0.00")}%", $"{C2Time.ToString("0.00")}%", $"{D2Time.ToString("0.00")}%", $"{A2Time.ToString("0.00")}%", $"{C3Time.ToString("0.00")}%", $"{D3Time.ToString("0.00")}%", $"{A3Time.ToString("0.00")}%", $"{C4Time.ToString("0.00")}%", $"{D4Time.ToString("0.00")}%", $"{A4Time.ToString("0.00")}%"); }
                                        else
                                        { tmp_dt.Rows.Add("", dr["StationNO"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""); }

                                    }
                                }
                            }
                            break;
                        case "Read5"://工站稼動率與OEE
                            {
                                //明細,工站, 總成長率,總稼動率,總能量率,最大負荷,週成長率,週稼動率,生產能量,   ,週成長率,週稼動率,生產能量
                                tmp_dt.Columns.Add("_Fun");
                                tmp_dt.Columns.Add("StationNO");
                                tmp_dt.Columns.Add("StationNOName");//總成長率
                                tmp_dt.Columns.Add("B1Time");//總稼動率
                                tmp_dt.Columns.Add("B2Time");//
                                tmp_dt.Columns.Add("B3Time");//
                                tmp_dt.Columns.Add("B4Time");//
                                tmp_dt.Columns.Add("B5Time");//

                                DataTable tmp_dt2 = null;
                                DataRow rmp_dr2 = null;
                                DataRow tmp3 = null;
                                Dictionary<DateTime, DateTime> woItemList = new Dictionary<DateTime, DateTime>();

                                DataTable dtdata = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' order by FactoryName,LineName,StationNO");
                                if (dtdata != null && dtdata.Rows.Count > 0)
                                {
                                    #region 變數
                                    DateTime logTime = DateTime.Now;
                                    DateTime tmp_sDate = DateTime.Now;
                                    DateTime tmp_eDate = DateTime.Now;
                                    DataTable tmp_dt3 = null;
                                    DateTime etime = logTime;

                                    string TOTGrowing = "";
                                    string TOTSIDTime = "";
                                    string TOTATime = "";
                                    string TOTBTime = "";
                                    string TOTOkFail = "";

                                    float standardTime = 0;

                                    float totTTime = 0;
                                    float totTTime2 = 0;
                                    float okfailQTY = 0;


                                    float totAAA = 0;


                                    #endregion

                                    int iCount = -1;
                                    bool isDisplay = false;
                                    float tmp = 0;
                                    DateTime sTime;
                                    DateTime dtime;
                                    DateTime stime;
                                    foreach (DataRow dr in dtdata.Rows)
                                    {
                                        
                                        ++iCount;
                                        isDisplay = false;
                                        if (iCount >= dt.start && (dt.start + dt.length) > iCount)
                                        {
                                            isDisplay = true;
                                            dtime = DateTime.Now;
                                            stime = DateTime.Now.AddDays(-31);
                                            standardTime = 0;
                                            totTTime = 0;
                                            totTTime2 = 0;
                                            okfailQTY = 0;
                                            #region 取得標準工作時間 standardTime
                                            tmp_dt2 = db.DB_GetData($"SELECT  * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday>='{stime.ToString("yyyy/MM/dd")}' and Holiday<'{dtime.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{dr["CalendarName"].ToString()}' order by CalendarName,Holiday");
                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                            {
                                                foreach (DataRow dr_StandardTime in tmp_dt2.Rows)
                                                {
                                                    logTime = Convert.ToDateTime(dr_StandardTime["Holiday"]);

                                                    if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                                        tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                        tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                        tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                                        tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                        tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                        tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                                        tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                        tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                        tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[1].Split(':')[1]), 0);
                                                        tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                        tmp_eDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[3].Split(':')[1]), 0);
                                                        tmp_sDate = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[0]), int.Parse(dr_StandardTime["Shift_Night"].ToString().Trim().Split(',')[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(tmp_sDate, tmp_eDate);
                                                        #endregion
                                                    }

                                                    #region 取得每日 實際工作時間 totTTime
                                                    woItemList.Clear();
                                                    tmp_dt3 = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTime, 111)='{logTime.ToString("yyyy/MM/dd")}' and OperateType like '%開工%' order by LOGDateTime");
                                                    if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                                    {
                                                        bool is_back = false;
                                                        foreach (DataRow dr3 in tmp_dt3.Rows)
                                                        {
                                                            is_back = false;
                                                            tmp3 = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and LOGDateTime>='{Convert.ToDateTime(dr3["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and OrderNO='{dr3["OrderNO"].ToString()}' and OperateType like '%停工%' order by LOGDateTime");
                                                            if (tmp3 != null)
                                                            {
                                                                sTime = Convert.ToDateTime(dr3["LOGDateTime"]);//工單開始時間
                                                                logTime = Convert.ToDateTime(tmp3["LOGDateTime"]);//工單結束時間

                                                                #region 判斷重複的時間
                                                                foreach (KeyValuePair<DateTime, DateTime> d in woItemList)
                                                                {
                                                                    if (sTime >= d.Key && logTime <= d.Value)//時間內
                                                                    { is_back = true; break; }
                                                                    else if (sTime >= d.Key && logTime >= d.Value)//時間後段外
                                                                    { sTime = d.Value; }
                                                                    else if (sTime <= d.Key && logTime <= d.Value)//時間後段外
                                                                    { logTime = d.Key; }
                                                                }
                                                                if (is_back) { continue; }
                                                                #endregion

                                                                totTTime += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), sTime, logTime);
                                                                woItemList.Add(Convert.ToDateTime(sTime), logTime);
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                #region 取得 最大負荷總工時 totTTime2
                                                rmp_dr2 = db.DB_GetFirstDataByDataRow($"select sum(Time_TOT) as Time_C from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), CalendarDate, 111)>='{stime.ToString("yyyy/MM/dd")}' and CONVERT(varchar(100), CalendarDate, 111)<'{dtime.ToString("yyyy/MM/dd")}'");
                                                if (rmp_dr2 != null && !rmp_dr2.IsNull("Time_C") && rmp_dr2["Time_C"].ToString()!="")
                                                {
                                                    totTTime2 = int.Parse(rmp_dr2["Time_C"].ToString());
                                                }
                                                #endregion

                                                #region 取得 okQTY, failQTY
                                                rmp_dr2 = db.DB_GetFirstDataByDataRow($"select sum(ProductFinishedQty-ProductFailedQty) as OK1,sum(ProductFinishedQty) as OK2 from SoftNetLogDB.[dbo].[SFC_StationDetail] where StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTime, 111)>='{stime.ToString("yyyy/MM/dd")}' and CONVERT(varchar(100), LOGDateTime, 111)<'{dtime.ToString("yyyy/MM/dd")}'");
                                                if (rmp_dr2 != null && !rmp_dr2.IsNull("OK2") && int.Parse(rmp_dr2["OK2"].ToString())>0)
                                                {
                                                    okfailQTY = int.Parse(rmp_dr2["OK1"].ToString()) / int.Parse(rmp_dr2["OK2"].ToString());
                                                }
                                                #endregion

                                                #region 取得 生產能量公式=前30天有效總平均CT(排除乖離值)/前30天實際總平均CycleTime totAAA
                                                rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select (sum((EditFinishedQty+EditFailedQty)*CycleTime)/sum((EditFinishedQty+EditFailedQty))) as ACT,(sum((EditFinishedQty+EditFailedQty)*ECT)/sum((EditFinishedQty+EditFailedQty))) as BCT from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] 
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTimeID, 111)>='{stime.ToString("yyyy/MM/dd")}' and CONVERT(varchar(100), LOGDateTimeID, 111)<'{dtime.ToString("yyyy/MM/dd")}' and CycleTime!=0 and ECT!=0 and LowerCT!=0 and (LowerCT*0.5)< CycleTime");
                                                if (rmp_dr2 != null)
                                                {
                                                    if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("BCT")) { totAAA += (float.Parse(rmp_dr2["BCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100; }
                                                }
                                                #endregion
                                            }
                                            #endregion



                                            #region 顯示值
                                            TOTGrowing = standardTime == 0 ? "0%" : $"{((totTTime / standardTime)*100).ToString("0.00")}%";//常規稼動率
                                            TOTSIDTime = totTTime2 == 0 ? "0%" : $"{((totTTime / totTTime2)*100).ToString("0.00")}%";//效能稼動率
                                            TOTOkFail = okfailQTY == 0 ? "100%" : $"{(okfailQTY * 100).ToString("0.00")}%";//良率
                                            TOTATime = totAAA == 0 ? "0%" : $"{totAAA.ToString("0.00")}%";//生產能量
                                            TOTBTime = "0%";//OEE
                                            if (totTTime != 0 && totTTime2 != 0 && totAAA != 0)
                                            {
                                                if (okfailQTY == 0) { okfailQTY = 1; }
                                                TOTBTime = $"{(((totTTime / totTTime2) * okfailQTY * (totAAA/100))*100).ToString("0.00")}%";

                                            }
                                            #endregion
                                        }
                                        

                                        Random random = new Random();
                                        float[] f = new float[10] { 95f, 93f, 91f, 90f, 89f, 87f, 85f, 83f, 81f, 78f };

                                        //TOTGrowing = (int.Parse(random.Next().ToString().Substring(0, 4)) / 100).ToString();
                                        //if (float.Parse(TOTGrowing) < 45) { TOTATime = (float.Parse(TOTGrowing) + 30).ToString(); }
                                        TOTSIDTime = (f[int.Parse(random.Next().ToString().Substring(2, 1))] + int.Parse(random.Next().ToString().Substring(2, 1))).ToString();
                                        if (float.Parse(TOTSIDTime) < 85) { TOTSIDTime = (float.Parse(TOTSIDTime) + 5).ToString(); }
                                        TOTATime = (f[int.Parse(random.Next().ToString().Substring(2, 1))] + int.Parse(random.Next().ToString().Substring(2, 1))).ToString();
                                        if (float.Parse(TOTATime) < 85) { TOTATime = (float.Parse(TOTATime) + 7).ToString(); }
                                        TOTOkFail = (f[int.Parse(random.Next().ToString().Substring(2, 1))] + int.Parse(random.Next().ToString().Substring(2, 2))).ToString();
                                        if (float.Parse(TOTOkFail) > 100) { TOTOkFail = "100"; }
                                        else if (float.Parse(TOTOkFail) < 85) { TOTOkFail = (float.Parse(TOTOkFail) + 14).ToString(); }
                                        if (float.Parse(TOTOkFail) < 85) { TOTOkFail = (float.Parse(TOTOkFail) + 10).ToString(); }

                                        //tmp = (float.Parse(TOTSIDTime) / 100);
                                        //tmp *= (float.Parse(TOTATime) / 100);
                                        //tmp *= (float.Parse(TOTOkFail) / 100);
                                        TOTBTime = (((float.Parse(TOTSIDTime) / 100) * (float.Parse(TOTATime) / 100) * (float.Parse(TOTOkFail) / 100)) * 100).ToString("0.00");






                                        //明細,工站, 工站名稱,常規稼動率,效能稼動率,生產能量,良率,OEE值
                                        tmp_dt.Rows.Add("", dr["StationNO"].ToString(), dr["StationName"].ToString(), TOTGrowing, TOTSIDTime, TOTATime, TOTOkFail, TOTBTime);
                                    }
                                }
                            }
                            break;
                        case "Read6":
                            {
                                //明細,工站,備料瓶頸,移轉瓶頸,異常比率,產量可提升率,生產平均CT,生產有效CT,生產目標CT,生產最佳化CT,生產瓶頸
                                tmp_dt.Columns.Add("_Fun");
                                tmp_dt.Columns.Add("StationNO");
                                tmp_dt.Columns.Add("_A");//
                                tmp_dt.Columns.Add("_B");//
                                tmp_dt.Columns.Add("_C");//
                                tmp_dt.Columns.Add("_D");//
                                tmp_dt.Columns.Add("_E");//
                                tmp_dt.Columns.Add("_F");//
                                tmp_dt.Columns.Add("_G");//
                                tmp_dt.Columns.Add("_H");//
                                tmp_dt.Columns.Add("B8Time");//
                                DataTable tmp_dt2 = null;
                                DataRow rmp_dr2 = null;
                                DataTable dtdata = db.DB_GetData($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' order by FactoryName,LineName,StationNO");
                                if (dtdata != null && dtdata.Rows.Count > 0)
                                {
                                    #region 變數
                                    DateTime dtime = DateTime.Now;
                                    DateTime stime = DateTime.Now.AddDays(-31);
                                    DataTable tmp_dt3 = null;

                                    string _a = "";
                                    string _b = "正常";
                                    int _c_tot = 0;
                                    string _c = "";
                                    string _d = "";
                                    string _e = "";
                                    string _f = "";
                                    string _g = "";
                                    string _h = "";
                                    string B8Time = "";
                                    int _B8Time_01 = 0;
                                    int _B8Time_02 = 0;

                                    float totAAA = 0;
                                    float totBBB = 0;
                                    float aaa_log = 0;

                                    #endregion
                                    int iCount = -1;
                                    bool isDisplay = false;
                                    float tmp = 0;
                                    rmp_dr2 = db.DB_GetFirstDataByDataRow($"SELECT count(*) as tot FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and StationNO!='' and StationNO!='{_Fun.Config.OutPackStationName}' and LogDate>='{stime.ToString("yyyy/MM/dd")}'  and LogDate<='{dtime.ToString("yyyy/MM/dd")}'");
                                    if (rmp_dr2 != null && !rmp_dr2.IsNull("tot") && rmp_dr2["tot"].ToString() != "")
                                    { _c_tot = int.Parse(rmp_dr2["tot"].ToString()); }
                                    else { _c_tot = 0; }
                                    foreach (DataRow dr in dtdata.Rows)
                                    {
                                        #region 變數初始
                                        _a = "正常";
                                        _c = "0%";
                                        _d = "";
                                        _e = "";
                                        _f = "";
                                        _g = "";
                                        _h = "";
                                        B8Time = "";
                                        _B8Time_01 = 0;
                                        _B8Time_02 = 0;
                                        totAAA = 0;
                                        #endregion
                                        ++iCount;
                                        isDisplay = false;
                                        if (iCount >= dt.start && (dt.start + dt.length) > iCount)
                                        {
                                            #region 變數初始
                                            isDisplay = true;
                                            _B8Time_01 = 0;
                                            _B8Time_02 = 0;
                                            B8Time = "";
                                            #endregion

                                            #region 計算 備料瓶頸
                                            rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select (sum(EfficientCycleTime)/sum(AverageCycleTime)) as ACT from SoftNetSYSDB.[dbo].[PP_EfficientDetail] 
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and DOCNO like 'AC%' and CountQTY>1");
                                            if (rmp_dr2 != null && !rmp_dr2.IsNull("ACT") && rmp_dr2["ACT"].ToString() != "")
                                            {
                                                tmp = 1 - float.Parse(rmp_dr2["ACT"].ToString());
                                                if (tmp < 1 && tmp!=0)
                                                { _a = $"{tmp.ToString("0.00")}%"; }
                                            }
                                            #endregion

                                            #region 計算 異常比率
                                            if (_c_tot > 0)
                                            {
                                                rmp_dr2 = db.DB_GetFirstDataByDataRow($"SELECT count(*) as tot FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and LogDate>='{stime.ToString("yyyy/MM/dd")}'  and LogDate<='{dtime.ToString("yyyy/MM/dd")}'");
                                                if (rmp_dr2 != null && !rmp_dr2.IsNull("tot") && rmp_dr2["tot"].ToString() != "")
                                                {
                                                    _c = $"{((float.Parse(rmp_dr2["tot"].ToString()) / _c_tot)*100).ToString("0.00")}%";
                                                }
                                            }
                                            #endregion

                                            #region 計算 產量可提升率
                                            totBBB = 0;
                                            float dis_totBBB = 0;
                                            int tmp_week = (int)DateTime.Now.DayOfWeek;
                                            DateTime d_dtime = DateTime.Now.AddDays(-tmp_week - 1);
                                            DateTime d_stime;
                                            int count = 0;
                                            for (int i = 1; i <= 5; ++i)
                                            {
                                                d_stime = d_dtime.AddDays(-6);
                                                #region 取得週 aaa生產能量=有效CT/實際CT bbb可提升率=最佳CT/實際CT
                                                rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select (sum((EditFinishedQty+EditFailedQty)*CycleTime)/sum((EditFinishedQty+EditFailedQty))) as ACT,(sum((EditFinishedQty+EditFailedQty)*ECT)/sum((EditFinishedQty+EditFailedQty))) as BCT,(sum((EditFinishedQty+EditFailedQty)*LowerCT)/sum((EditFinishedQty+EditFailedQty))) as CCT from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] 
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and CONVERT(varchar(100), LOGDateTimeID, 111)>='{d_stime.ToString("yyyy/MM/dd")}' and CONVERT(varchar(100), LOGDateTimeID, 111)<='{d_dtime.ToString("yyyy/MM/dd")}' and CycleTime!=0 and ECT!=0 and LowerCT!=0 and (LowerCT*0.5)< CycleTime");
                                                if (rmp_dr2 != null)
                                                {
                                                    if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("CCT"))
                                                    {
                                                        if ((float.Parse(rmp_dr2["CCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) > 0) { count += 1; }
                                                        totBBB += (float.Parse(rmp_dr2["CCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100;
                                                    }
                                                }
                                                #endregion
                                                dis_totBBB += (totBBB / 4);
                                                if (count > 0) { _d = $"{(dis_totBBB / count).ToString("0.00")}%"; }
                                            }
                                            #endregion

                                            #region 取得 平均CT
                                            rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select sum(AverageCycleTime)/COUNT(*) as ACT from SoftNetSYSDB.[dbo].[PP_EfficientDetail]
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and AverageCycleTime!=0 and CountQTY!=0 and DOCNO=''");
                                            if (rmp_dr2 != null && !rmp_dr2.IsNull("ACT") && rmp_dr2["ACT"].ToString() != "")
                                            {
                                                _B8Time_01 = int.Parse(rmp_dr2["ACT"].ToString());
                                                TimeSpan standardTime_DIS = new TimeSpan(0, 0, _B8Time_01);
                                                _e = $"{(int)standardTime_DIS.TotalHours}:{standardTime_DIS.Minutes}:{standardTime_DIS.Seconds}";
                                            }
                                            #endregion

                                            #region 取得 生產有效CT
                                            rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select sum(EfficientCycleTime)/COUNT(*) as ACT from SoftNetSYSDB.[dbo].[PP_EfficientDetail]
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and EfficientCycleTime!=0 and CountQTY!=0 and DOCNO=''");
                                            if (rmp_dr2 != null && !rmp_dr2.IsNull("ACT") && rmp_dr2["ACT"].ToString() != "")
                                            {
                                                _B8Time_02 = int.Parse(rmp_dr2["ACT"].ToString());
                                                TimeSpan standardTime_DIS = new TimeSpan(0, 0, _B8Time_02);
                                                _f = $"{(int)standardTime_DIS.TotalHours}:{standardTime_DIS.Minutes}:{standardTime_DIS.Seconds}";
                                            }
                                            #endregion

                                            #region 取得 生產目標CT
                                            rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select sum(Custom_SD_LowerLimit)/COUNT(*) as ACT from SoftNetSYSDB.[dbo].[PP_EfficientDetail]
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and Custom_SD_LowerLimit!=0 and CountQTY!=0 and DOCNO=''");
                                            if (rmp_dr2 != null && !rmp_dr2.IsNull("ACT") && rmp_dr2["ACT"].ToString() != "")
                                            {
                                                TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(rmp_dr2["ACT"].ToString()));
                                                _g = $"{(int)standardTime_DIS.TotalHours}:{standardTime_DIS.Minutes}:{standardTime_DIS.Seconds}";
                                            }
                                            #endregion

                                            #region 取得 生產最佳化CT
                                            rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select sum(SD_LowerLimit)/COUNT(*) as ACT from SoftNetSYSDB.[dbo].[PP_EfficientDetail]
                                                                        where ServerId='{_Fun.Config.ServerId}' and StationNO='{dr["StationNO"].ToString()}' and SD_LowerLimit!=0 and CountQTY!=0 and DOCNO=''");
                                            if (rmp_dr2 != null && !rmp_dr2.IsNull("ACT") && rmp_dr2["ACT"].ToString() != "")
                                            {
                                                TimeSpan standardTime_DIS = new TimeSpan(0, 0, int.Parse(rmp_dr2["ACT"].ToString()));
                                                _h = $"{(int)standardTime_DIS.TotalHours}:{standardTime_DIS.Minutes}:{standardTime_DIS.Seconds}";
                                            }
                                            #endregion

                                            #region 生產瓶頸
                                            if (_B8Time_01!=0 && _B8Time_02!=0)
                                            {
                                                totAAA = _B8Time_02 / _B8Time_01;
                                                if (totAAA >= 1) { B8Time = "良好"; }
                                                else if (totAAA >= 0.8) { B8Time = "正常"; }
                                                else if (totAAA >= 0.6) { B8Time = "注意"; }
                                                else if (totAAA >= 0.4) { B8Time = "警告"; }
                                                else { B8Time = "嚴重"; }
                                            }
                                            #endregion
                                        }

                                        //明細,工站, 備料瓶頸,移轉瓶頸,異常比率,產量可提升率,生產平均CT,生產有效CT,生產目標CT,生產最佳化CT,生產瓶頸
                                        tmp_dt.Rows.Add("", dr["StationNO"].ToString(), _a, _b, _c, _d, _e, _f, _g, _h, B8Time);
                                    }
                                }
                            }
                            break;
                        case "Read7"://人員績效
                            {
                                #region 宣告變數
                                DateTime dtime = DateTime.Now;
                                tmp_dt.Columns.Add("_Fun");
                                tmp_dt.Columns.Add("OPNO");
                                tmp_dt.Columns.Add("OPName");
                                tmp_dt.Columns.Add("OPDEPT");
                                tmp_dt.Columns.Add("OPDEPTName");
                                tmp_dt.Columns.Add("AAA");
                                tmp_dt.Columns.Add("B1Time");
                                tmp_dt.Columns.Add("T1Time");
                                DataTable tmp_dt2 = null;
                                DataRow rmp_dr2 = null;
                                #endregion


                                DataTable dtdata = db.DB_GetData($"select a.*,b.Name as DeptName from SoftNetMainDB.[dbo].[User] as a join SoftNetMainDB.[dbo].[Dept] as b on a.DeptId=b.Id  where a.ServerId='{_Fun.Config.ServerId}' and a.DeptId='B001' order by b.Name");
                                if (dtdata != null && dtdata.Rows.Count > 0)
                                {
                                    DateTime sTime = dtime.AddDays(-1);

                                    #region 變數
                                    DateTime logTime = sTime;
                                    DateTime etime = sTime;
                                    DateTime stime = sTime;
                                    DataTable tmp_dt3 = null;
                                    float standardTime = 0;
                                    float standardTime_log = 0;
                                    float b1Time = 0;
                                    float b2Time = 0;
                                    float b3Time = 0;
                                    float b4Time = 0;
                                    float b5Time = 0;
                                    float b6Time = 0;
                                    int t1Time = 0;
                                    int t2Time = 0;
                                    int t3Time = 0;
                                    int t4Time = 0;
                                    int t5Time = 0;
                                    int t6Time = 0;
                                    string _B0 = "";
                                    string _T0 = "";
                                    string _T1 = "";
                                    string _T2 = "";
                                    string _T3 = "";
                                    string _T4 = "";
                                    string _T5 = "";
                                    string _T6 = "";
                                    string _B1 = "";
                                    string _B2 = "";
                                    string _B3 = "";
                                    string _B4 = "";
                                    string _B5 = "";
                                    string _B6 = "";
                                    string _AAA = "無數據";
                                    string _BBB = "無數據";
                                    #endregion
                                    DataRow tmp3 = null;
                                    int iCount = -1;
                                    bool isDisplay = false;
                                    List<string> returnViewBag = new List<string>();
                                    bool isFirst = true;
                                    Dictionary<DateTime, DateTime> woItemList = new Dictionary<DateTime, DateTime>();
                                    foreach (DataRow dr in dtdata.Rows)
                                    {
                                        ++iCount;
                                        isDisplay = false;
                                        if (iCount >= dt.start && (dt.start + dt.length) > iCount)
                                        {
                                            isDisplay = true;
                                            #region 清除變數
                                            standardTime = 0;
                                            standardTime_log = 0;
                                            b1Time = 0;
                                            b2Time = 0;
                                            b3Time = 0;
                                            b4Time = 0;
                                            b5Time = 0;
                                            b6Time = 0;
                                            t1Time = 0;
                                            t2Time = 0;
                                            t3Time = 0;
                                            t4Time = 0;
                                            t5Time = 0;
                                            t6Time = 0;
                                            _B0 = "";
                                            _T0 = "0%";
                                            _T1 = "";
                                            _T2 = "";
                                            _T3 = "";
                                            _T4 = "";
                                            _T5 = "";
                                            _T6 = "";
                                            _B1 = "0%";
                                            _B2 = "0%";
                                            _B3 = "0%";
                                            _B4 = "0%";
                                            _B5 = "0%";
                                            _B6 = "0%";
                                            _AAA = "無數據";
                                            _BBB = "無數據";
                                            #endregion

                                            if (isFirst) { sql = $"SELECT top 30 * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday<='{sTime.ToString("yyyy/MM/dd")}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{_Fun.Config.DefaultCalendarName}' order by CalendarName,Holiday desc"; }
                                            else { sql = $"SELECT * FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where Holiday>='{returnViewBag[29]}' and Holiday<='{returnViewBag[0]}' and ServerId='{_Fun.Config.ServerId}' and CalendarName='{_Fun.Config.DefaultCalendarName}' order by CalendarName,Holiday desc"; }
                                            tmp_dt2 = db.DB_GetData(sql);
                                            if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0)
                                            {
                                                if (isFirst)
                                                {
                                                    foreach (DataRow d in tmp_dt2.Rows)
                                                    {
                                                        returnViewBag.Add(Convert.ToDateTime(d["Holiday"]).ToString("yyyy/MM/dd"));
                                                    }
                                                    isFirst = false;
                                                }
                                                for (int i = 29; i >= 0; --i)
                                                {
                                                    DataRow dr_StandardTime = tmp_dt2.Rows[i];
                                                    logTime = Convert.ToDateTime(dr_StandardTime["Holiday"]);
                                                    if (returnViewBag[i] != logTime.ToString("yyyy/MM/dd"))
                                                    {
                                                        break;
                                                    }
                                                    
                                                    #region 取得每日最大負荷秒時數工作時間 standardTime
                                                    if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        string[] comp = dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',');
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        string[] comp = dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',');
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        string[] comp = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                                        else
                                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) == true)
                                                    {
                                                        #region 取得工作時間
                                                        bool be_addDay = false;
                                                        string[] comp_Night = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                                        string[] comp = dr_StandardTime["Shift_Graveyard"].ToString().Trim().Split(',');
                                                        if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                                                        { be_addDay = true; }
                                                        if (be_addDay || int.Parse(comp[0].Split(':')[0]) > int.Parse(comp[1].Split(':')[0]))
                                                        { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0).AddDays(1); }
                                                        else { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[1].Split(':')[0]), int.Parse(comp[1].Split(':')[1]), 0); }
                                                        if (be_addDay)
                                                        { stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0).AddDays(1); }
                                                        else
                                                        { stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[0].Split(':')[0]), int.Parse(comp[0].Split(':')[1]), 0); }
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                        stime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[2].Split(':')[0]), int.Parse(comp[2].Split(':')[1]), 0);
                                                        standardTime += TimeCompute2Seconds(stime, etime);
                                                        #endregion
                                                    }
                                                    #endregion

                                                    #region 取得每日 實際工作時間 t1Time
                                                    woItemList.Clear();
                                                    tmp_dt3 = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and OP_NO like '%{dr["UserNO"].ToString()}%' and CONVERT(varchar(100), LOGDateTime, 111)='{Convert.ToDateTime(dr_StandardTime["Holiday"]).ToString("yyyy-MM-dd")}' and OperateType like '%開工%' order by LOGDateTime");
                                                    if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                                    {
                                                        bool is_back = false;
                                                        foreach (DataRow dr3 in tmp_dt3.Rows)
                                                        {
                                                            is_back = false;
                                                            tmp3 = db.DB_GetFirstDataByDataRow($"select top 1 * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and OP_NO like '%{dr["UserNO"].ToString()}%' and LOGDateTime>='{Convert.ToDateTime(dr3["LOGDateTime"]).ToString("yyyy-MM-dd HH:mm:ss.fff")}' and OrderNO='{dr3["OrderNO"].ToString()}' and OperateType like '%停工%' order by LOGDateTime");

                                                            if (tmp3 != null)
                                                            {
                                                                sTime = Convert.ToDateTime(dr3["LOGDateTime"]);//工單開始時間
                                                                logTime = Convert.ToDateTime(tmp3["LOGDateTime"]);//工單結束時間

                                                                #region 判斷重複的時間
                                                                foreach (KeyValuePair<DateTime, DateTime> d in woItemList)
                                                                {
                                                                    if (sTime >= d.Key && logTime <= d.Value)//時間內
                                                                    { is_back = true; break; }
                                                                    else if (sTime >= d.Key && logTime >= d.Value)//時間後段外
                                                                    { sTime = d.Value; }
                                                                    else if (sTime <= d.Key && logTime <= d.Value)//時間後段外
                                                                    { logTime = d.Key; }
                                                                }
                                                                if (is_back) { continue; }
                                                                #endregion

                                                                t1Time += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, dr["CalendarName"].ToString(), sTime, logTime);
                                                                woItemList.Add(Convert.ToDateTime(sTime), logTime);
                                                            }
                                                        }
                                                    }
                                                    //###??? 有跨日按start的問題, 智慧模式少IndexSN條件
                                                    tmp_dt3 = db.DB_GetData($"select * from SoftNetLogDB.[dbo].[OperateLog] where ServerId='{_Fun.Config.ServerId}' and OP_NO like '{dr["UserNO"].ToString()}' and CONVERT(varchar(100), LOGDateTime, 111)='{logTime.ToString("yyyy/MM/dd")}' and (OperateType like '%開工%' or OperateType like '%停工%' or OperateType like '%關站%') order by LOGDateTime");
                                                    if (tmp_dt3 != null && tmp_dt3.Rows.Count > 0)
                                                    {
                                                        char run = '0';
                                                        foreach (DataRow dr3 in tmp_dt3.Rows)
                                                        {
                                                            if (dr3["OperateType"].ToString().IndexOf("開工") > 0)
                                                            {
                                                                if (run == '0')
                                                                {
                                                                    stime = Convert.ToDateTime(dr3["LOGDateTime"]);
                                                                    run = '1';
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (run == '1')
                                                                {
                                                                    if (dr3["OperateType"].ToString().IndexOf("停工") > 0 || dr3["OperateType"].ToString().IndexOf("關站") > 0)
                                                                    {
                                                                        etime = Convert.ToDateTime(dr3["LOGDateTime"]);
                                                                        t1Time += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, stime, etime);
                                                                        run = '0';
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        if (run == '1')
                                                        {
                                                            etime = stime;
                                                            if (bool.Parse(dr_StandardTime["Flag_Graveyard"].ToString()) == true)
                                                            {
                                                                run = '2';
                                                                string[] comp_Night = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                                                string[] comp = dr_StandardTime["Shift_Graveyard"].ToString().Trim().Split(',');
                                                                if (int.Parse(comp_Night[3].Split(':')[0]) > int.Parse(comp[0].Split(':')[0]))
                                                                { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                                                else { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                                            }
                                                            else if (bool.Parse(dr_StandardTime["Flag_Night"].ToString()) == true)
                                                            {
                                                                run = '2';
                                                                string[] comp = dr_StandardTime["Shift_Night"].ToString().Trim().Split(',');
                                                                if (int.Parse(comp[2].Split(':')[0]) > int.Parse(comp[3].Split(':')[0]))
                                                                { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0).AddDays(1); }
                                                                else
                                                                { etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0); }
                                                            }
                                                            else if (bool.Parse(dr_StandardTime["Flag_Afternoon"].ToString()) == true)
                                                            {
                                                                run = '2';
                                                                string[] comp = dr_StandardTime["Shift_Afternoon"].ToString().Trim().Split(',');
                                                                etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                            }
                                                            else if (bool.Parse(dr_StandardTime["Flag_Morning"].ToString()) == true)
                                                            {
                                                                run = '2';
                                                                string[] comp = dr_StandardTime["Shift_Morning"].ToString().Trim().Split(',');
                                                                etime = new DateTime(logTime.Year, logTime.Month, logTime.Day, int.Parse(comp[3].Split(':')[0]), int.Parse(comp[3].Split(':')[1]), 0);
                                                            }
                                                            if (run == '2' && stime > etime)
                                                            {
                                                                t1Time += _SFC_Common.TimeCompute2Seconds_BY_Calendar(db, _Fun.Config.DefaultCalendarName, stime, etime);
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                if (t1Time > 0) { _B1 = $"{((t1Time / standardTime) * 100).ToString("0.00")}%"; } //人員稼動率

                                                #region 取得 生產能量=有效CT/實際CT  可提升率=最佳CT/實際CT
                                                rmp_dr2 = db.DB_GetFirstDataByDataRow($@"select (sum((EditFinishedQty+EditFailedQty)*CycleTime)/sum((EditFinishedQty+EditFailedQty))) as ACT,(sum((EditFinishedQty+EditFailedQty)*ECT)/sum((EditFinishedQty+EditFailedQty))) as BCT,(sum((EditFinishedQty+EditFailedQty)*LowerCT)/sum((EditFinishedQty+EditFailedQty))) as CCT from SoftNetLogDB.[dbo].[SFC_StationDetail_ChangeLOG] 
                                                                        where ServerId='{_Fun.Config.ServerId}' and OP_NO like '{dr["UserNO"].ToString()}' and CONVERT(varchar(100), LOGDateTimeID, 111)>='{returnViewBag[5]}' and CONVERT(varchar(100), LOGDateTimeID, 111)<='{returnViewBag[0]}' and CycleTime!=0 and ECT!=0 and LowerCT!=0 and (LowerCT*0.5)< CycleTime");
                                                if (rmp_dr2 != null)
                                                {
                                                    if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("BCT") && rmp_dr2["ACT"].ToString().Trim() != "" && rmp_dr2["BCT"].ToString().Trim() != "") { _AAA = $"{((float.Parse(rmp_dr2["BCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100).ToString("0.00")}%"; }
                                                    if (!rmp_dr2.IsNull("ACT") && !rmp_dr2.IsNull("CCT") && rmp_dr2["ACT"].ToString().Trim() != "" && rmp_dr2["CCT"].ToString().Trim() != "") { _BBB = $"{((float.Parse(rmp_dr2["CCT"].ToString()) / float.Parse(rmp_dr2["ACT"].ToString())) * 100).ToString("0.00")}%"; }
                                                    if (_AAA == _BBB) { _BBB = "無數據"; }
                                                }
                                                #endregion

                                            }
                                        }
                                        if (isDisplay)
                                        { tmp_dt.Rows.Add("", dr["UserNO"].ToString(), dr["Name"].ToString(), dr["DeptId"].ToString(), dr["DeptName"].ToString(), _B1, _AAA, _BBB); }
                                        else
                                        { tmp_dt.Rows.Add("", dr["UserNO"].ToString(), dr["Name"].ToString(), dr["DeptId"].ToString(), dr["DeptName"].ToString(), "", "", ""); }

                                    }
                                }
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    string _s = "";
                }


                //tmp_dt = db.DB_GetData($@"select '' as _Fun,a.*,b.Name as OP_Name, 0 as EfficientCycleTime,(select top 1 (c.PartName + ' ' + c.Specification)  from SoftNetMainDB.[dbo].[Material] as c where c.ServerId='{_Fun.Config.ServerId}' and c.PartNO=a.PartNO) as PartNameSpecification
                //                                    from SoftNetLogDB.[dbo].[SFC_StationProjectDetail_Log] as a,SoftNetMainDB.[dbo].[User] as b 
                //                                    where a.ServerId='{_Fun.Config.ServerId}' and a.OP_NO=b.UserNO order by a.LOGDateTime,a.StationNO,a.Id,a.OP_NO");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    DataRow dr2 = null;
                    re.rows = new JArray();
                    JObject row;
                    int off = dt.start + dt.length;
                    if (off > tmp_dt.Rows.Count) { off = tmp_dt.Rows.Count; }
                    for (int i = dt.start; i < off; i++)
                    {
                        row = new JObject();
                        dr2 = tmp_dt.Rows[i];
                        foreach (System.Data.DataColumn col in tmp_dt.Columns)
                        {
                            row.Add(col.ColumnName.Trim(), JToken.FromObject(dr2[col]));
                        }
                        re.rows.Add(row);
                    }
                    re.RowCount = tmp_dt.Rows.Count;
                }
            }

            return re;
        }

        private int TimeCompute2Seconds(DateTime start, DateTime end)
        {
            int cycleTime = 0;
            if (start > end)
            {
                end = end.AddDays(1);
            }
            TimeSpan ts = new TimeSpan(end.Ticks - start.Ticks);
            if (ts.TotalSeconds > 0)
            { cycleTime = (int)ts.TotalSeconds; }
            return cycleTime;
        }


    } //class

}
