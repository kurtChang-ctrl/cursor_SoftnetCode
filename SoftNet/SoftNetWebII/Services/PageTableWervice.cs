using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class PageTableWervice : XgEdit
    {
        public PageTableWervice(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetSYSDB.[dbo].[APS_WorkingPaper]",
                PkeyFid = "Id",
                Col4 = null,
                ReadSql = "select a.*,b.PartName,b.Specification from SoftNetSYSDB.[dbo].[APS_WorkingPaper] as a join SoftNetMainDB.[dbo].[Material] as b on b.PartNO=a.PartNO where a.Id='{0}'",
                Items = new EitemDto[]
                {
                    new() { Fid = "Id" },
                    new() { Fid = "PartNO",   },
                     new() { Fid = "NeedQTY", Required = true  },
                      new() { Fid = "Price"  },
                       new() { Fid = "StartTime"  },
                        new() { Fid = "ArrivalDate"  },
                        new() { Fid = "APS_StationNO"  },
                    new() { Fid = "MFNO" , Required = true },
                    new() { Fid = "Remark" },
                },
            };
        }

        private readonly ReadDto dto = new()
        {
            //ReadSql = $"select _Crud='',Id,WorkType,Class,PartNO,NeedQTY,Price,MFNO,IN_StoreNO,IN_StoreSpacesNO,APS_StationNO,StartTime,ArrivalDate,Remark from SoftNetSYSDB.[dbo].[APS_WorkingPaper] where ServerId='{_Fun.Config.ServerId}' order by WorkType",
            ReadSql = $@"select _Fun='',a.Id,a.WorkType,a.Class,a.PartNO,a.NeedQTY,a.Price,a.MFNO,a.IN_StoreNO,a.IN_StoreSpacesNO,a.APS_StationNO,a.StartTime,a.ArrivalDate,a.Remark,a.DOCNumberNO,(b.PartName+ ' ' +b.Specification) as PartNameSpecification
                        ,(select (c.MFNO+ ' ' +c.SName) from SoftNetMainDB.[dbo].[MFData] as c where c.ServerId='{_Fun.Config.ServerId}' and a.MFNO!='' and a.MFNO=c.MFNO group by c.MFNO,c.SName) as SName 
                        from SoftNetSYSDB.[dbo].[APS_WorkingPaper] as a,SoftNetMainDB.[dbo].[Material] as b
                        where a.IsOK='0' and a.ServerId='{_Fun.Config.ServerId}' and b.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO order by a.WorkType,a.SimulationId",
            Items = new QitemDto[] {
                new() { Fid = "PartNO", Op = ItemOpEstr.Like },
            },
            SpecifySQLFrom = $"Select Count(*) as _count from SoftNetSYSDB.[dbo].[APS_WorkingPaper] as a,SoftNetMainDB.[dbo].[Material] as b where a.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO",
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }



    }
}
