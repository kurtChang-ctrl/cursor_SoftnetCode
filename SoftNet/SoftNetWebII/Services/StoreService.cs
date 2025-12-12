using Base.Models;
using Base.Services;
using BaseApi.Services;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class StoreService : XgEdit
    {
        public StoreService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetMainDB.dbo.[Store]",
                PkeyFid = "StoreNO",
                PkeyFids = new string[] { "ServerId", "StoreNO" },
                Col4 = null,
                Items = new EitemDto[] {
                    new() { Fid = "ServerId" },
                    new() { Fid = "StoreNO" },
                    new() { Fid = "StoreName" },
                    new() { Fid = "FactoryName" },
                    new() { Fid = "Class" },
                    new() { Fid = "Simulation_FirstNO" },
                    new() { Fid = "PostNO" },
                    new() { Fid = "Address" },
                                        new() { Fid = "Config_macID" },
                    new() { Fid = "Default_IN_OUT" },
                },
                Childs = new EditDto[]
                {
                    new()
                    {
                        IsKeyValueNUM = true,
                        ReadSql = "SELECT A.Id,A.ServerId,A.StoreNO,A.StoreSpacesNO,A.StoreSpacesName,A.Remark,A.Config_macID FROM SoftNetMainDB.[dbo].[StoreII] as A where A.ServerId='{0}' and A.StoreNO='{1}' order by A.StoreNO,A.StoreSpacesNO",
                        Table = "SoftNetMainDB.dbo.[StoreII]",
                        PkeyFid = "Id",
                        FkeyFid = "StoreNO",
                        FkeyFids = new string[] { "ServerId", "StoreNO" },
                        OrderBy = "StoreNO",
                        Col4 = null,
                        Items = new EitemDto[] {
                            new() { Fid = "Id" },
                            new() { Fid = "StoreNO" },
                             new() { Fid = "ServerId" },
                            new() { Fid = "StoreSpacesNO" },
                            new() { Fid = "StoreSpacesName" },
                            new() { Fid = "Config_macID" },
                            new() { Fid = "Remark" },
                        },
                    },
                },
            };
        }


        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select  ServerId,StoreNO,StoreName,FactoryName,Class,Simulation_FirstNO,PostNO,Config_macID,Default_IN_OUT,Address,'' as _Fun from SoftNetMainDB.[dbo].[Store] where ServerId='{_Fun.Config.ServerId}' order by StoreNO",
            Items = new QitemDto[] {
                new() { Fid = "StoreNO" },
                new() { Fid = "FactoryName" },
                new() { Fid = "Class" },
            },
        };

        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }


    } //class
}
