using Base.Models;
using Base.Services;
using BaseApi.Services;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class STViewService : XgEdit
    {
        public STViewService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetMainDB.[dbo].[Manufacture]",
                PkeyFid = "StationNO",
                Col4 = null,
                Items = new EitemDto[] {
                    new() { Fid = "State" },
                },
            };
        }

        //CumulativeTime=現在時間-Start 
        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select c.StationUI_type,c.FactoryName,b.OP_NO,b.State,b.StationNO,a.OrderNO,a.PP_Name,a.IndexSN,a.IsLastStation,a.StartTime,a.CumulativeTime,a.StatndCycleTime,a.AvarageWaitTime,a.AvarageCycleTime,a.TotalInput,a.TotalOutput,a.TotalFail,a.YieldRate,'' as _Fun 
            from SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] as a, SoftNetMainDB.[dbo].[Manufacture] as b, SoftNetSYSDB.[dbo].PP_Station as c 
            where b.Config_MutiWO='0' and b.ServerId='{_Fun.Config.ServerId}' and a.StationNO!='{_Fun.Config.OutPackStationName}' and a.StationNO=b.StationNO and b.StationNO=c.StationNO and (a.EndTime is null or a.EndTime='') and a.StartTime is not null
            order by c.FactoryName,a.StationNO,a.OrderNO",
            Items = new QitemDto[] {
                new() { Fid = "FactoryName" },
            },

        };

        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

    } //class
    public class STView2Service : XgEdit
    {
        public STView2Service(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetMainDB.[dbo].[Manufacture]",
                PkeyFid = "StationNO",
                Col4 = null,
                Items = new EitemDto[] {
                    new() { Fid = "State" },
                },
            };
        }

        //CumulativeTime=現在時間-Start 
        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select c.StationUI_type,c.FactoryName,b.OP_NO,b.State,b.StationNO,a.OrderNO,a.PP_Name,a.IndexSN,a.IsLastStation,a.StartTime,a.CumulativeTime,a.StatndCycleTime,a.AvarageWaitTime,a.AvarageCycleTime,a.TotalInput,a.TotalOutput,a.TotalFail,a.YieldRate,'' as _Fun 
            from SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] as a, SoftNetMainDB.[dbo].[Manufacture] as b, SoftNetSYSDB.[dbo].PP_Station as c 
            where b.Config_MutiWO='1' and b.ServerId='{_Fun.Config.ServerId}' and a.StationNO=b.StationNO and b.StationNO=c.StationNO and (a.EndTime is null or a.EndTime='') and a.StartTime is not null
            order by c.FactoryName,a.StationNO,a.OrderNO",
            Items = new QitemDto[] {
                new() { Fid = "FactoryName" },
            },

        };

        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

    } //class
    public class STView2WorkService : XgEdit
    {
        public STView2WorkService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetMainDB.[dbo].[Manufacture]",
                PkeyFid = "StationNO",
                Col4 = null,
                Items = new EitemDto[] {
                    new() { Fid = "State" },
                },
            };
        }

        //CumulativeTime=現在時間-Start 
        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select c.StationUI_type,c.FactoryName,b.OP_NO,b.State,b.StationNO,a.OrderNO,a.PP_Name,a.IndexSN,a.IsLastStation,a.StartTime,a.CumulativeTime,a.StatndCycleTime,a.AvarageWaitTime,a.AvarageCycleTime,a.TotalInput,a.TotalOutput,a.TotalFail,a.YieldRate,'' as _Fun 
            from SoftNetSYSDB.[dbo].[PP_WorkOrder_Settlement] as a, SoftNetMainDB.[dbo].[Manufacture] as b, SoftNetSYSDB.[dbo].PP_Station as c 
            where b.Config_MutiWO='1' and b.ServerId='{_Fun.Config.ServerId}' and a.StationNO=b.StationNO and b.StationNO=c.StationNO and (a.EndTime is null or a.EndTime='') and a.StartTime is not null
            order by c.FactoryName,a.StationNO,a.OrderNO",
            Items = new QitemDto[] {
                new() { Fid = "FactoryName" },
            },

        };

        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

    } //class
}

