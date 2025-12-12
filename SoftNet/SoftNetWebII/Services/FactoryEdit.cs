using Base.Enums;
using Base.Models;
using Base.Services;
using Newtonsoft.Json.Linq;

namespace SoftNetWebII.Services
{
    public class FactoryEdit : XgEdit
    {
        public FactoryEdit(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "dbo.[Factory]",
                PkeyFid = "Id",
                Col4 = null,
                //ReadSql = "select * from dbo.[Factory] where Id='{0}'",
                Items = new EitemDto[] 
				{
					new() { Fid = "Id" },
					new() { Fid = "FactoryName", Required = true  },
					new() { Fid = "Manager" },
					new() { Fid = "Address" , Required = true },
                    new() { Fid = "Telephone" },
                },
            };
        }

    } //class
}
