using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class DOCSSignService : XgEdit
    {
        public DOCSSignService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "dbo.[XpFlowSign]",
                PkeyFid = "Id",
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
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
            ReadSql = $@"
                select b.DOCName,a.* from SoftNetMainDB.dbo.[XpFlowSign] as a
                left join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.SourceId,1,4)
                where a.SignerId='{_Fun.UserId()}' order by a.SourceId",
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }
    }
}