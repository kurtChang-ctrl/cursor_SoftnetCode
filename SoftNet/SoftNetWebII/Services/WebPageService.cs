using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;

using System.Data;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class WebPageService
    {
    }
    public class ProductProcessService : XgEdit
    {
        public ProductProcessService(string ctrl) : base(ctrl) { }

        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select ServerId,PP_Name,FactoryName,LineName,CalendarName,UpdateTime,IsAutoCloseWO,Remark,'' as _F1,'' as _Fun from SoftNetSYSDB.[dbo].[PP_ProductProcess]
                        where ServerId='{_Fun.Config.ServerId}' order by FactoryName,LineName",
            Items = new QitemDto[] {
                new() { Fid = "FactoryName" },
            },
        };

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetSYSDB.[dbo].[PP_ProductProcess]",
                PkeyFid = "PP_Name",
                PkeyFids = new string[] { "ServerId", "PP_Name" },
                Col4 = null,
                Items = new EitemDto[] {
                    new() { Fid = "ServerId" },
                    new() { Fid = "PP_Name" },
                    new() { Fid = "FactoryName" },
                    new() { Fid = "LineName" },
                    new() { Fid = "CalendarName" },
                    new() { Fid = "UpdateTime" },
                    new() { Fid = "Remark" },
                    new() { Fid = "IsAutoCloseWO" },
                },
                Childs = new EditDto[]
                {
                    new()
                    {
                        Table = "SoftNetSYSDB.[dbo].[PP_ProductProcess_Item]",
                        PkeyFid = "Id",
                        FkeyFid = "PP_Name",
                        OrderBy = "sn",
                        Col4 = null,
                        IsKeyValueNUM=true,
                        ReadSql="select Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,Sub_PP_Name,Station_Custom_IndexSN,DisplayName,OutPackType,MFNO from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where ServerId='{0}' and PP_Name='{1}'",
                        Items = new EitemDto[] {
                            new() { Fid = "Id" },
                            new() { Fid = "ServerId" },
                            new() { Fid = "FactoryName" },
                            new() { Fid = "LineName" },
                            new() { Fid = "PP_Name" },
                            new() { Fid = "DisplaySN" },
                            new() { Fid = "IndexSN" },
                            new() { Fid = "IndexSN_Merge" },
                            new() { Fid = "StationNO" },
                            new() { Fid = "Sub_PP_Name" },
                            new() { Fid = "Station_Custom_IndexSN" },
                            new() { Fid = "DisplayName" },
                            new() { Fid = "OutPackType" },
                            new() { Fid = "MFNO" },
                        },
                    },
                },
            };
        }


        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

    }
}
