using Base;
using Base.Models;
using Base.Services;
using BaseWeb.Models;
using DocumentFormat.OpenXml.Bibliography;
using Newtonsoft.Json.Linq;

using System.Data;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class SelectService : XgEdit
    {
        public SelectService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "dbo.[DOC3stockII]",
                PkeyFid = "Id",
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "Id" },
                    new() { Fid = "DOCNumberNO", Required = true  },
                    new() { Fid = "PartNO" , Required = true },
                    new() { Fid = "Price" , Required = true },
                    new() { Fid = "QTY", Required = true  },
                },
            };
        }


        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select c.DOCName,a.*,b.DOCNO,b.FlowStatus from SoftNetMainDB.[dbo].[DOC3stockII] a,SoftNetMainDB.[dbo].[DOC3stock] as b,SoftNetMainDB.[dbo].[DOCRole] as c
                        where b.ServerId='{_Fun.Config.ServerId}' and c.ServerId='{_Fun.Config.ServerId}' and a.DOCNumberNO=b.DOCNumberNO and c.DOCNO=b.DOCNO order by a.IsOK,b.DOCNO,a.DOCNumberNO,a.OUT_StoreNO,a.IN_StoreNO",
            Items = new QitemDto[] {
                new() { Fid = "DOCNO" },
                new() { Fid = "IN_StoreNO" },
                new() { Fid = "OUT_StoreNO" },
            },
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
    }
    public class Select2Service : XgEdit
    {
        public Select2Service(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetMainDB.[dbo].[TotalStock]",
                PkeyFid = "Id",
                Col4 = null,
                ReadSql = "select * from SoftNetMainDB.[dbo].[TotalStock] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "Id" },
                    new() { Fid = "QTY" },
                    new() { Fid = "StoreNO", Required = true  },
                    new() { Fid = "StoreSpacesNO" , Required = true },
                },
            };
        }


        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select '' as _Fun,a.Id,a.StoreNO,a.StoreSpacesNO,a.PartNO,a.QTY,sum(b.KeepQTY+b.OverQTY) as KeepQTY,c.PartName,c.Specification,c.Unit,d.StoreName,IIF(a.QTY-sum(b.KeepQTY+b.OverQTY) IS NULL,a.QTY,a.QTY-sum(b.KeepQTY+b.OverQTY)) as TotalQTY from SoftNetMainDB.[dbo].[TotalStock] as a
                        left join SoftNetMainDB.[dbo].[TotalStockII] as b on a.Id=b.Id 
                        join SoftNetMainDB.[dbo].[Material] as c on a.PartNO=c.PartNO
                        join SoftNetMainDB.[dbo].[Store] as d on a.StoreNO=d.StoreNO
                        where a.ServerId='{_Fun.Config.ServerId}'  group by a.Id,a.StoreNO,a.StoreSpacesNO,a.PartNO,a.QTY,c.PartName,c.Specification,c.Unit,d.StoreName Order by a.StoreNO,a.PartNO",
            Items = new QitemDto[] {
                new() { Fid = "StoreNO" },
                new() { Fid = "PartNO" },
            },
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
    }
    public class Select3Service : XgEdit
    {
        public Select3Service(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetMainDB.[dbo].[TotalStock]",
                PkeyFid = "Id",
                Col4 = null,
                //ReadSql = "select * from SoftNetMainDB.[dbo].[TotalStock] where Id='{0}'",
                ReadSql = @"select a.ServerId,a.Id,a.StoreNO,a.StoreSpacesNO,a.PartNO,a.QTY,c.PartName,c.Specification,c.Unit,d.StoreName from SoftNetMainDB.[dbo].[TotalStock] as a
                        join SoftNetMainDB.[dbo].[Material] as c on a.PartNO=c.PartNO
                        join SoftNetMainDB.[dbo].[Store] as d on a.StoreNO=d.StoreNO
                        where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "Id" },
                    new() { Fid = "StoreNO" },
                    new() { Fid = "StoreName" },
                    new() { Fid = "PartNO" },
                    new() { Fid = "PartName" },
                    new() { Fid = "Specification" },
                    new() { Fid = "Unit" },
                    new() { Fid = "QTY" , Required = true},

                },
            };
        }


        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select '' as _Fun,a.ServerId,a.Id,a.StoreNO,a.StoreSpacesNO,a.PartNO,a.QTY,c.PartName,c.Specification,c.Unit,d.StoreName from SoftNetMainDB.[dbo].[TotalStock] as a
                        join SoftNetMainDB.[dbo].[Material] as c on a.PartNO=c.PartNO
                        join SoftNetMainDB.[dbo].[Store] as d on a.StoreNO=d.StoreNO
                        where a.ServerId='{_Fun.Config.ServerId}' and c.ServerId='{_Fun.Config.ServerId}' and d.ServerId='{_Fun.Config.ServerId}'  Order by a.PartNO,a.StoreNO",
            Items = new QitemDto[] {
                new() { Fid = "StoreNO" },
                new() { Fid = "StoreName" },
                new() { Fid = "PartNO" },
                new() { Fid = "PartName" },
                new() { Fid = "Specification" },
                new() { Fid = "Unit" },
            },
            Db_forFind_From_AS_Name = "a.",
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
    }
    public class Report01Service : XgEdit
    {
        public Report01Service(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "dbo.[TotalStock]",
                PkeyFid = "Id",
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "Id" },
                    new() { Fid = "StoreNO", Required = true  },
                    new() { Fid = "StoreSpacesNO" , Required = true },
                },
            };
        }


        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select a.* FROM SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] as a 
                        join SoftNetSYSDB.[dbo].[PP_WorkOrder] as b on a.OrderNO=b.OrderNO and (b.EndTime is null or b.EndTime='') and b.StartTime is not null
                        where a.ServerId='{_Fun.Config.ServerId}' and a.StartTime is not null order by a.OrderNO,a.StationNO,a.DisplaySN,a.StartTime",
            Items = new QitemDto[] {
                new() { Fid = "OrderNO" },
                new() { Fid = "StationNO" },
                new() { Fid = "PP_Name" },
            },
            SQL_StoredProgram="",
        };
        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            dto.SQL_ByProgram = Run_ProgramReadDataOBJ(dt, ctrl);
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
        private ProgramReadDataOBJ Run_ProgramReadDataOBJ(DtDto dt, string ctrl)
        {
            ProgramReadDataOBJ re = new ProgramReadDataOBJ();
            return re;
        }
    }
    public class Report02Service : XgEdit
    {
        public Report02Service(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "dbo.[TotalStock]",
                PkeyFid = "Id",
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "Id" },
                    new() { Fid = "StoreNO", Required = true  },
                    new() { Fid = "StoreSpacesNO" , Required = true },
                },
            };
        }

        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select a.* FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] where ServerId='{_Fun.Config.ServerId}'",
            Items = new QitemDto[] {
                new() { Fid = "StationNO" },
                new() { Fid = "PartNO" },
                new() { Fid = "Sub_PartNO" },
            },
            SQL_StoredProgram = "",
        };
        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            dto.SQL_ByProgram = Run_ProgramReadDataOBJ(dt, ctrl);
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
        private ProgramReadDataOBJ Run_ProgramReadDataOBJ(DtDto dt, string ctrl)
        {
            ProgramReadDataOBJ re = new ProgramReadDataOBJ();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable tmp_dt = db.DB_GetData($"SELECT * FROM SoftNetMainDB.[dbo].[Manufacture] where  ServerId='{_Fun.Config.ServerId}' and StationNO!='{_Fun.Config.OutPackStationName}' order by OrderNO,State");
                if (tmp_dt != null && tmp_dt.Rows.Count > 0)
                {
                    DataTable tmp_dt1 = new DataTable();
                    DataTable tmp_dt2 = null;
                    DataRow tmp_dr = null;
                    string partNO = "";
                    string sql = "";
                    foreach (DataRow dr in tmp_dt.Rows)
                    {
                        sql = "";
                        if (dr["SimulationId"].ToString().Trim() != "")
                        {
                            tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[APS_Simulation] where SimulationId='{dr["SimulationId"].ToString()}'");
                            if (tmp_dr != null)
                            {
                                partNO = tmp_dr["PartNO"].ToString();
                                sql = $@"SELECT StationNO,[Sub_PartNO] as PNO,SUM([CountQTY]) as TOTCount,SUM([AverageCycleTime]*[CountQTY])/SUM([CountQTY]) as agv,SUM([EfficientCycleTime]*[CountQTY])/SUM([CountQTY]) as eCT,
                                    SUM([Custom_SD_LowerLimit]*[CountQTY])/SUM([CountQTY]) as customCT,SUM([SD_LowerLimit]*[CountQTY])/SUM([CountQTY]) as lowerCT,
                                    SUM([SD_UpperLimit]*[CountQTY])/SUM([CountQTY]) as upperCT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] 
                                    WHERE ServerId='{_Fun.Config.ServerId}' and [StationNO] = '{dr["StationNO"].ToString()}' and IndexSN={dr["IndexSN"].ToString()} AND [Sub_PartNO]='{partNO}' and Next_StationNO is null GROUP BY StationNO,[Sub_PartNO]";
                                tmp_dt2 = db.DB_GetData(sql);
                                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0) { tmp_dt1.Merge(tmp_dt2); }
                            }
                        }
                        else
                        {
                            tmp_dr = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{_Fun.Config.ServerId}' and OrderNO='{dr["OrderNO"].ToString()}'");
                            if (tmp_dr != null)
                            {
                                partNO = tmp_dr["PartNO"].ToString();
                                sql = $@"SELECT StationNO,[PartNO] as PNO,SUM([CountQTY]) as TOTCount,SUM([AverageCycleTime]*[CountQTY])/SUM([CountQTY]) as agv,SUM([EfficientCycleTime]*[CountQTY])/SUM([CountQTY]) as eCT,
                                    SUM([Custom_SD_LowerLimit]*[CountQTY])/SUM([CountQTY]) as customCT,SUM([SD_LowerLimit]*[CountQTY])/SUM([CountQTY]) as lowerCT,
                                    SUM([SD_UpperLimit]*[CountQTY])/SUM([CountQTY]) as upperCT FROM SoftNetSYSDB.[dbo].[PP_EfficientDetail] 
                                    WHERE ServerId='{_Fun.Config.ServerId}' and [StationNO] = '{dr["StationNO"].ToString()}' and IndexSN={dr["IndexSN"].ToString()} AND [PartNO]='{partNO}' and Next_StationNO is null GROUP BY StationNO,[PartNO]";
                                tmp_dt2 = db.DB_GetData(sql);
                                if (tmp_dt2 != null && tmp_dt2.Rows.Count > 0) { tmp_dt1.Merge(tmp_dt2); }
                            }
                        }
                    }
                    if (tmp_dt1.Rows.Count > 0)
                    {
                        re.rows = new JArray();
                        JObject row;
                        DataRow dr2 = null;
                        int off = dt.start + dt.length;
                        if (off > tmp_dt1.Rows.Count) { off = tmp_dt1.Rows.Count; }
                        for (int i = dt.start;i< off;i++)
                        {
                            row = new JObject();
                            dr2 = tmp_dt1.Rows[i];
                            foreach (System.Data.DataColumn col in tmp_dt1.Columns)
                            {
                                row.Add(col.ColumnName.Trim(), JToken.FromObject(dr2[col]));
                            }
                            re.rows.Add(row);
                        }
                        re.RowCount = tmp_dt1.Rows.Count;
                    }
                }
            }
            return re;
        }
    }

}
