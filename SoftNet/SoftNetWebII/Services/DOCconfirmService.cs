using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;

using System;
using System.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{

    public class DOCconfirmService : XgEdit
    {
        public DOCconfirmService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetMainDB.[dbo].[DOC1BuyII]",
                PkeyFid = "Id",
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "Id", Required = true },
                    new() { Fid = "DOCNumberNO", Required = true  },
                    new() { Fid = "PartNO" , Required = true },
                    new() { Fid = "Price",   },
                    new() { Fid = "Unit" ,  },
                    new() { Fid = "QTY" , },
                    new() { Fid = "Remark" ,  },
                    new() { Fid = "ArrivalDate" , },
                    new() { Fid = "SimulationId" , },
                },
            };
        }


        private readonly ReadDto dto = new()
        {

            ReadSql = @"select Id,DOCNumberNO,PartNO,Price,Unit,QTY,Remark,ArrivalDate='',SimulationId,_Crud='' from SoftNetMainDB.[dbo].[DOC3stockII] where IsOK='0' order by DOCNumberNO,PartNO,ArrivalDate,Id",

            Items = new QitemDto[] {
                new() { Fid = "DOCNumberNO", Op = ItemOpEstr.Like },
                new() { Fid = "Price", Op = ItemOpEstr.Like },

            },
            SQL_StoredProgram = "EXEC SoftNetMainDB.[dbo].[spGet_DOC1Buy_IsOK]",
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

    }

    public class DOC4confirmService : XgEdit
    {
        public DOC4confirmService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetMainDB.[dbo].[DOC4ProductionII]",
                PkeyFid = "Id",
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "Id", Required = true },
                    new() { Fid = "DOCNumberNO", Required = true  },
                    new() { Fid = "PartNO" , Required = true },
                    new() { Fid = "Price",   },
                    new() { Fid = "Unit" ,  },
                    new() { Fid = "QTY" , },
                    new() { Fid = "Remark" ,  },
                    new() { Fid = "ArrivalDate" , },
                    new() { Fid = "SimulationId" , },
                },
            };
        }


        private readonly ReadDto dto = new()
        {

            ReadSql = @"select Id,DOCNumberNO,PartNO,Price,Unit,QTY,Remark,ArrivalDate='',SimulationId,_Crud='' from SoftNetMainDB.[dbo].[DOC4ProductionII] where IsOK='0' order by DOCNumberNO,PartNO,ArrivalDate,Id",

            Items = new QitemDto[] {
                new() { Fid = "DOCNumberNO", Op = ItemOpEstr.Like },
                new() { Fid = "Price", Op = ItemOpEstr.Like },

            },
            SQL_StoredProgram = "EXEC SoftNetMainDB.[dbo].[spGet_DOC4_IsOK]",
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

    }

}
