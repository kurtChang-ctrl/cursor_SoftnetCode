using Base.Models;
using Base.Services;
using BaseApi.Services;
using BaseWeb.Extensions;

namespace SoftNetWebII.Services
{
    public class MyBaseUserService : IBaseUserService
    {

        public MyBaseUserService()
        {
            string _s = "";
        }
        //get base user info
        public BaseUserDto GetData()
        {
            return _Http.GetSession().Get<BaseUserDto>(_Fun.BaseUser);   //extension method
        }
    }
}
