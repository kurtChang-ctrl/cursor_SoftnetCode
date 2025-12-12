using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class FactoryRead
    {
        private readonly ReadDto dto = new()
        {
            ReadSql = @"select _Crud='',Id,FactoryName,Manager,Address,Telephone from SoftNetMainDB.[dbo].[Factory] order by Id",
            Items = new QitemDto[] {
                new() { Fid = "FactoryName", Op = ItemOpEstr.Like },
            },
        };

        public async Task<JObject> GetPageAsync(DtDto dt, string ctrl)
        {
            //3.call CrudRead.GetPage()
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

    } 
}