using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class DOCService : XgEdit
    {
        public DOCService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "dbo.[DOCRole]",
                PkeyFid = "Id",
                Col4 = null,
                //ReadSql = "select * from SoftNetMainDB.[dbo].[Factory] where Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "Id" },
                    new() { Fid = "DOCType", Required = true  },
                    new() { Fid = "DOCNO" , Required = true },
                    new() { Fid = "DOCName" , Required = true },
                    new() { Fid = "Remark"  },
                },
            };
        }

        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select Id,DOCType,DOCNO,DOCName,Remark,_Crud='' from SoftNetMainDB.[dbo].[DOCRole] where ServerId='{_Fun.Config.ServerId}' order by DOCNO",
            Items = new QitemDto[] {
                new() { Fid = "DOCType", Op = ItemOpEstr.Like },
                 new() { Fid = "DOCNO", Op = ItemOpEstr.Like },
                  new() { Fid = "DOCName", Op = ItemOpEstr.Like },
            },
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
    } 
}