using Base.Models;
using Base.Services;
using BaseApi.Services;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class BOMEdit : XgEdit
    {
        public BOMEdit(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "SoftNetMainDB.[dbo].[BOM]",
                PkeyFid = "Id",
                Col4 = null,
                EmptyToNulls = new string[] { "EffectiveDate", "ExpiryDate" },
                Items = new EitemDto[] {
                    new() { Fid = "Id" },
                    new() { Fid = "ServerId" },
                    new() { Fid = "PartNO" },
                    new() { Fid = "Main_Item" },
                    new() { Fid = "Apply_PP_Name" },
                    new() { Fid = "Apply_StationNO" },
                    new() { Fid = "IndexSN" },
                    new() { Fid = "EffectiveDate" },
                    new() { Fid = "ExpiryDate" },
                    new() { Fid = "Version" },
                    new() { Fid = "OutPackType" },
                    new() { Fid = "IsEnd" },
                },
                Childs = new EditDto[]
                {
                    new()
                    {
                        Table = "SoftNetMainDB.[dbo].[BOMII]",
                        PkeyFid = "Id",
                        FkeyFid = "BOMId",
                        OrderBy = "sn",
                        Col4 = null,
                        IsKeyValueNUM=true,
                        Items = new EitemDto[] {
                            new() { Fid = "Id" },
                            new() { Fid = "BOMId" },
                             new() { Fid = "ServerId" },
                            new() { Fid = "sn" },
                            new() { Fid = "PartNO" },
                            new() { Fid = "BOMQTY" },
                            new() { Fid = "Class" },
                            new() { Fid = "AttritionRate" },
                        },
                    },
                },
            };
        }

    } //class
}
