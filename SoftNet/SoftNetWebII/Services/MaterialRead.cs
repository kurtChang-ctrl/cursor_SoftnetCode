using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class MaterialRead
    {
        private readonly ReadDto dto = new()
        {
            ReadSql = $@"select ServerId,PartNO,PartName,Specification,Model,Class,PartType,SafeQTY,SafeWeightQty,Unit,StoreSTime,BuySTime,Buy_Flag,OutSTime,Out_Flag,APS_Default_MFNO,APS_Default_StoreNO,APS_Default_StoreSpacesNO,IS_Store_Test,Use_ONLineTime,Use_LimitDay,Use_Count,Remark,'' as _Crud 
                        from dbo.[Material] where ServerId='{_Fun.Config.ServerId}' order by PartNO",
            Items = new QitemDto[] 
            {
                new() { Fid = "PartNO", Op = ItemOpEstr.Like },
                 new() { Fid = "PartName", Op = ItemOpEstr.Like },
                  new() { Fid = "Model", Op = ItemOpEstr.Like },
                   new() { Fid = "Class" },
            },
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            //3.call CrudRead.GetPage()
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

    } 
}