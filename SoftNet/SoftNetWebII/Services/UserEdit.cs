using Base.Models;
using Base.Services;

namespace SoftNetWebII.Services
{
    public class UserEdit : XgEdit
    {
        public UserEdit(string ctrl) : base(ctrl) { }

        override public EditDto GetDto()
        {
            return new EditDto
            {
				Table = "dbo.[User]",
                PkeyFid = "Id",
                Col4 = null,
                Items = new EitemDto[] 
				{
					new() { Fid = "Id" },
					new() { Fid = "Account" },
                    new() { Fid = "UserNO" },
                    new() { Fid = "Name" },
					new() { Fid = "Pwd" },
					new() { Fid = "DeptId" },
					new() { Fid = "Status" },
                    new() { Fid = "ServerId" },
                },
                Childs = new EditDto[]
                {
                    new EditDto
                    {
                        Table = "dbo.[XpUserRole]",
                        PkeyFid = "Id",
                        FkeyFid = "UserId",
                        Col4 = null,
                        Items = new EitemDto[] 
						{
							new() { Fid = "Id" },
							new() { Fid = "UserId" },
							new() { Fid = "RoleId", Required = true },
                        },
                    },
                },
            };
        }

    } //class
}
