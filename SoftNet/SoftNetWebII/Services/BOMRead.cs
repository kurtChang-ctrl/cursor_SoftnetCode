using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class BOMRead
    {
        private readonly ReadDto dto = new()
        {
            //ReadSql = $@"select a.Id,a.ServerId,a.PartNO,[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],[OutPackType],'' as _Fun,b.PartName,b.Specification 
            //            from SoftNetMainDB.[dbo].[BOM] as a,SoftNetMainDB.[dbo].[Material] as b where a.ServerId='{_Fun.Config.ServerId}' and a.PartNO=b.PartNO order by a.Apply_PP_Name,a.IndexSN",
            ReadSql = $@"select a.Id,a.ServerId,a.PartNO,Main_Item,EffectiveDate,ExpiryDate,Version,Apply_PP_Name,Apply_StationNO,IsEnd,IndexSN,Station_Custom_IndexSN,StationNO_Custom_DisplayName,OutPackType,'' as _Fun,b.PartName,b.Specification from SoftNetMainDB.[dbo].[BOM] as a
                        join SoftNetMainDB.[dbo].[Material] as b on a.PartNO=b.PartNO
                        where a.ServerId='{_Fun.Config.ServerId}' order by a.Apply_PP_Name,a.IndexSN",
            TableAs = "a",
            Items = new QitemDto[] {
                new() { Fid = "PartNO", Op = ItemOpEstr.Like },
                new() { Fid = "Apply_PP_Name" },
                 new() { Fid = "EffectiveDate" },
                  new() { Fid = "ExpiryDate" },
            },
        };

        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

        /// <summary>
        /// export excel
        /// </summary>
        /// <param name="ctrl">controller name for authorize</param>
        /// <param name="find"></param>
        /// <returns></returns>

    } //class
}