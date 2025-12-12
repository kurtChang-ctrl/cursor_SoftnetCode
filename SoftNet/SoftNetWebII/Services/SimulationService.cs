using Base;
using Base.Models;
using Base.Services;
using BaseApi.Services;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json.Linq;

using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class SimulationService : XgEdit
    {
        public SimulationService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetSYSDB.[dbo].APS_NeedData",
                PkeyFid = "Id",
                PkeyFids = new string[] { "Id", "ServerId" },//若多主Key欄位,要定義否則定義null
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "ServerId" },
                    new() { Fid = "Id" },
                    new() { Fid = "IsAdd_SafeQTY" },
                    new() { Fid = "NeedDate" , Required = true},
                    new() { Fid = "PartNO" , Required = true},
                    new() { Fid = "NeedQTY" , Required = true},
                    new() { Fid = "Apply_PP_Name" , Required = true},
                    new() { Fid = "CalendarName" , Required = true},
                    new() { Fid = "BOMId" },
                    new() { Fid = "CTNO" },
                    new() { Fid = "CTName" },
                    new() { Fid = "FactoryName" },
                    new() { Fid = "NeedSource" },
                    new() { Fid = "BufferTime" , Required = true},
                },
            };
        }


        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select '' as _F0,'' as _F1,a.NeedSource,a.CTNO,a.CTName,a.ServerId,a.Id,a.NeedType,a.NeedDate,a.PartNO,b.PartName,b.Specification,b.Class,a.BufferTime,a.NeedQTY,a.ChangeQTY,a.StockQTY,a.StateINFO,a.WIP7Day,a.WIP30Daya,a.CalendarName,a.State,a.NeedSimulationDate 
            from SoftNetSYSDB.[dbo].APS_NeedData as a, SoftNetMainDB.[dbo].Material as b 
            where a.ServerId='{_Fun.Config.ServerId}' and b.ServerId='{_Fun.Config.ServerId}' and a.NeedType!='5' and a.State!='9' and a.State!='6' and a.State!='7' and a.PartNO=b.PartNO order by a.NeedDate",
            Items = new QitemDto[] {
                new() { Fid = "ServerId" },
                new() { Fid = "NeedSource" },
                   new() { Fid = "CTNO" },
                      new() { Fid = "CTName" },
                 new() { Fid = "Id" },
                new() { Fid = "NeedType" },
                new() { Fid = "NeedDate" },
                new() { Fid = "PartNO" },
                new() { Fid = "PartName" },
                new() { Fid = "Specification" },
                new() { Fid = "Class" },
                new() { Fid = "BufferTime" },
                new() { Fid = "NeedQTY" },
                new() { Fid = "ChangeQTY" },
                new() { Fid = "StockQTY" },
                new() { Fid = "StateINFO" },
                new() { Fid = "WIP7Day" },
                new() { Fid = "WIP30Daya" },
                new() { Fid = "State" },
                new() { Fid = "CalendarName" },
                new() { Fid = "NeedSimulationDate" },
            },
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

        public async Task<ResultDto> OtherDeleteAsync(string key)
        {
            ResultDto re = await DeleteAsync(key);
            if (re.ErrorMsg == "")
            {
                string[] tmp = key.Split(';');
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation] where NeedId='{tmp[0]}'");
                    db.DB_SetData($"DELETE FROM SoftNetMainDB.[dbo].[TotalStockII] where NeedId='{tmp[0]}'");
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where NeedId='{tmp[0]}'");
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] where NeedId='{tmp[0]}'");
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_Simulation_ErrorData] where NeedId='{tmp[0]}'");
                    db.DB_SetData($"DELETE FROM SoftNetSYSDB.[dbo].[APS_WarningData] where NeedId='{tmp[0]}'");
                    db.DB_SetData($"DELETE FROM SoftNetLogDB.[dbo].[APS_Change_S_Date_Log] where Trigger_NeedId='{tmp[0]}'");
                }
            }
            return re;
        }

    } //class
}
