using Base;
using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;

using System;
using System.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class StationService : XgEdit
    {
        public StationService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetSYSDB.[dbo].[PP_Station]",
                PkeyFid = "StationNO",
                PkeyFids = new string[] { "ServerId", "StationNO" },//若多主Key欄位,要定義否則定義null
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "StationNO", Required = true },
                    new() { Fid = "FactoryName", Required = true  },
                    new() { Fid = "LineName" , Required = true },
                    new() { Fid = "StationName", Required = true  },
                    new() { Fid = "ServerId" ,  },
                    new() { Fid = "RMSName" ,  },
                    new() { Fid = "Station_Type" , Required = true },
                    new() { Fid = "StationUI_type" , Required = true },
                    new() { Fid = "Remark" , },
                    new() { Fid = "CalendarName" , },
                },
            };
        }

        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select StationNO,FactoryName,LineName,StationName,RMSName,Station_Type,StationUI_type,Remark,CalendarName,Station_Type,_Crud='' from SoftNetSYSDB.[dbo].[PP_Station] where  ServerId='{_Fun.Config.ServerId}' order by FactoryName,LineName,StationNO",
            Items = new QitemDto[] {
                new() { Fid = "StationNO", Op = ItemOpEstr.Like },
                 new() { Fid = "FactoryName", Op = ItemOpEstr.Like },
                  new() { Fid = "LineName", Op = ItemOpEstr.Like },
                  new() { Fid = "RMSName", Op = ItemOpEstr.Like },
            },
        };
       
        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
        public async Task<ResultDto> OtherCrateAsync(JObject json)
        {
            ResultDto re= await CreateAsync(json);
            if (re.ErrorMsg == "")
            {
                var rows= json["_rows"] as JArray;
                string Config_MutiWO = "0";
                if (rows[0]["Station_Type"].ToString().Trim() == "8") { Config_MutiWO = "1"; }
                string sql = $"INSERT INTO SoftNetMainDB.[dbo].[Manufacture] (ServerId,StationNO,Config_MutiWO) VALUES ('{_Fun.Config.ServerId}','{rows[0]["StationNO"].ToString()}','{Config_MutiWO}')";
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    db.DB_SetData(sql);
                }
            }
            return re;
        }
        public async Task<ResultDto> OtherDeleteAsync(string key)
        {
            ResultDto re = await DeleteAsync(key);
            if (re.ErrorMsg == "")
            {
                string sql = $"DELETE FROM SoftNetMainDB.[dbo].[Manufacture] WHERE ServerId='{_Fun.Config.ServerId}' and StationNO='{key}'";
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    db.DB_SetData(sql);
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] WHERE StationNO='{key}'");
                }
            }
            return re;
        }

    }
}
