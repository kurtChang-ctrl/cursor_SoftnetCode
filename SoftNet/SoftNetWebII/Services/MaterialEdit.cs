using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;

namespace SoftNetWebII.Services
{
    public class MaterialEdit : XgEdit
    {
        public MaterialEdit(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "dbo.[Material]",
                PkeyFid = "PartNO",
                PkeyFids = new string[] { "ServerId", "PartNO" },//若多主Key欄位,要定義否則定義null
                Col4 = null,
                //ReadSql = "select * from dbo.[Material] where PartNO='{0}'",
                Items = new EitemDto[] 
				{
					new() { Fid = "PartNO" , Required = true },
					new() { Fid = "PartName", Required = true  },
					new() { Fid = "Specification" },
					new() { Fid = "Model" },
                    new() { Fid = "ServerId" },
                    new() { Fid = "Class", Required = true  },
                    new() { Fid = "PartType", Required = true  },
                    new() { Fid = "Unit", Required = true  },
                    new() { Fid = "SafeQTY"  },
                    new() { Fid = "StoreSTime" },
                    new() { Fid = "APS_Default_MFNO" },
                    new() { Fid = "APS_Default_StoreNO" },
                    new() { Fid = "APS_Default_StoreSpacesNO" },
                    new() { Fid = "IS_Store_Test" },
                    new() { Fid = "Use_ONLineTime" },
                    new() { Fid = "Use_LimitDay" },
                    new() { Fid = "Use_Count" },

                },
            };
        }

    } //class
}
