using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;
using SoftNetWebII.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class APSService : XgEdit
    {

        public APSService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetSYSDB.[dbo].[PP_Station]",
                PkeyFid = "StationNO",
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "StationNO", Required = true },
                    new() { Fid = "FactoryName", Required = true  },
                    new() { Fid = "LineName" , Required = true },
                    new() { Fid = "StationName", Required = true  },
                    new() { Fid = "RMSName" ,  },
                    new() { Fid = "Station_Type" , Required = true },
                    new() { Fid = "StationUI_type" , Required = true },
                    new() { Fid = "Remark" , },
                },
            };
        }

        private readonly ReadDto dto = new()
        {
            //###??? SQL不正確 可參考RUNTimeServer的自動開立code 改用 SQL_ByProgram寫
            ReadSql = $@"select a.PartNO,a.CalendarDate,a.NeedId,a.SimulationId,a.NeedQTY,b.PartName,b.Specification,_Fun=''
                        from SoftNetSYSDB.[dbo].[APS_PartNOTimeNote] as a 
                        join SoftNetSYSDB.[dbo].[APS_Simulation] as a2 on a.SimulationId=a2.SimulationId and a2.NeedQTY>a2.Math_TotalStock_HasUseQTY and a2.DOCNumberNO=''
                        Left join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO 
                        where b.ServerId='{_Fun.Config.ServerId}' and a.DOCNumberNO='' and (a.Class='1' or a.Class='2' or a.Class='3') order by a.PartNO,a.NeedId,a.CalendarDate",
            Items = new QitemDto[] {
                new() { Fid = "PartNO", Op = ItemOpEstr.Like },


            },
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }



    }
}
