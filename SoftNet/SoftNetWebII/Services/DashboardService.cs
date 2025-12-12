using Base.Models;
using Base.Services;
using BaseApi.Services;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class DashboardService : XgEdit
    {
        public DashboardService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "dbo.[Factory]",
                PkeyFid = "Id",
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "Id" },
                    new() { Fid = "FactoryName", Required = true  },
                    new() { Fid = "Manager" , Required = true },
                    new() { Fid = "Address" , Required = true },
                    new() { Fid = "Telephone", Required = true  },
                },
            };
        }


        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select '' as _F1,b.FactoryName,b.LineName,b.StationNO,b.StationName,a.State,a.OrderNO,a.OP_NO,a.Master_PP_Name,a.PP_Name,a.IndexSN,a.Station_Custom_IndexSN from SoftNetMainDB.[dbo].[Manufacture] a,SoftNetSYSDB.[dbo].[PP_Station] as b where a.ServerId='{_Fun.Config.ServerId}' and a.StationNO=b.StationNO order by b.FactoryName,b.LineName,b.StationNO",
            Items = new QitemDto[] {
                new() { Fid = "StoreNO" },
                new() { Fid = "FactoryName" },
                new() { Fid = "Class" },
            },
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }


    } //class
}
