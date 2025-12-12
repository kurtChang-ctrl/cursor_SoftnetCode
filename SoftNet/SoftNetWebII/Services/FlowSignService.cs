using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class FlowSignService : XgEdit
    {

        public FlowSignService(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            var locale = _Xp.GetLocale0();
            return new EditDto
            {
                ReadSql = $@"
select
    SignId=s.Id,
    LeaveId=l.Id,
    l.StartTime, l.EndTime, l.Hours, l.Created, l.FileName,
    LeaveName=c.Name_{locale},
    UserName=u.Name,
    AgentName=u2.Name
from dbo.XpFlowSign s
join dbo.Leave l on s.SourceId=l.Id and s.FlowLevel=l.FlowLevel and s.SignStatus='0'
join dbo.[User] u on l.UserId=u.Id
join dbo.[User] u2 on l.AgentId=u2.Id
join dbo.XpCode c on c.Type='LeaveType' and l.LeaveType=c.Value
where s.Id='{{0}}'
",
            };
        }

        private readonly ReadDto dto = new()
        {
            ReadSql = $@"
                select b.DOCName,a.* from SoftNetMainDB.[dbo].[XpFlowSign] as a
                left join SoftNetMainDB.[dbo].[DOCRole] as b on b.DOCNO=SUBSTRING(a.SourceId,1,4) and b.ServerId='{_Fun.Config.ServerId}'
                where a.SignerId='{_Fun.UserId()}' and SignStatus='0' order by a.SourceId",
        };



        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

    }
}
