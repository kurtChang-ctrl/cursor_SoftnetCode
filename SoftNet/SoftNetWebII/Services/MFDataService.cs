using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;

using System;
using System.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class MFDataService : XgEdit
    {
        public MFDataService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "dbo.[MFData]",
                PkeyFid = "MFNO",
                PkeyFids = new string[] { "ServerId", "MFNO" },//若多主Key欄位,要定義否則定義null
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "MFNO", Required = true },
                    new() { Fid = "MFName"  },
                    new() { Fid = "SName",  },
                    new() { Fid = "UniFormNO" , },
                    new() { Fid = "TEL",  },
                    new() { Fid = "FAX" ,  },
                    new() { Fid = "ContactMan" ,  },
                    new() { Fid = "ContactTEL" ,  },
                    new() { Fid = "EMail" ,  },
                    new() { Fid = "Address" ,  },
                    new() { Fid = "PaymentMethod" ,  },
                    new() { Fid = "PaymentDate" ,  },
                    new() { Fid = "PaymentTerms" ,  },
                    new() { Fid = "StoreNO" ,  },
                    new() { Fid = "CTDataWeights" , },
                    new() { Fid = "Remark" , },
                    new() { Fid = "ServerId" , },
                },

            };
        }

        private readonly ReadDto dto = new()
        {
            //ReadSql = @"select MFNO,MFName,UniFormNO,TEL,FAX,ContactMan,ContactTEL,EMail,Address,PaymentMethod,PaymentDate,PaymentTerms,StoreNO,Remark,_Crud='' from dbo.[MFData] order by MFNO",
            ReadSql = $@"select * from dbo.[MFData] where ServerId='{_Fun.Config.ServerId}' order by MFNO",
            Items = new QitemDto[] {
                new() { Fid = "MFNO", Op = ItemOpEstr.Like },
                new() { Fid = "MFName", Op = ItemOpEstr.Like },
                new() { Fid = "StoreNO", Op = ItemOpEstr.Like },
            },
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
        public async Task<ResultDto> OtherCrateAsync(JObject json)
        {
            ResultDto re = await CreateAsync(json);
            //if (re.ErrorMsg == "")
            //{
            //    var rows = json["_rows"] as JArray;
            //    string s01 = rows[0]["StationNO"].ToString();
            //    string sql = $"INSERT INTO dbo.[Manufacture] (StationNO) VALUES ('{rows[0]["StationNO"].ToString()}')";
            //    using (DBADO db = new DBADO("1", _Fun.Config.Db))
            //    {
            //        db.DB_SetData(sql);
            //        db.DB_SetData($"INSERT INTO SoftNetSYSDB.[dbo].[APS_WorkTimeNote] (StationNO,CalendarDate) VALUES ('{rows[0]["StationNO"].ToString()}','{DateTime.Now.ToString("yyyy-MM-dd")}')");
            //    }
            //}
            return re;
        }
        public async Task<ResultDto> OtherDeleteAsync(string key)
        {
            ResultDto re = await DeleteAsync(key);
            //if (re.ErrorMsg == "")
            //{
            //    string sql = $"DELETE FROM dbo.[Manufacture] WHERE StationNO='{key}'";
            //    using (DBADO db = new DBADO("1", _Fun.Config.Db))
            //    {
            //        db.DB_SetData(sql);
            //        db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] WHERE StationNO='{key}'");
            //    }
            //}
            return re;
        }

    }
}

