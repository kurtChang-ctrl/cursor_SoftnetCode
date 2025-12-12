using Base;
using Base.Models;
using Base.Services;
using BaseApi.Services;
using BaseWeb.Attributes;
using BaseWeb.Extensions;
using Microsoft.AspNetCore.Diagnostics;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Models;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class HomeController : Controller
    {
        [XgLogin]
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Login(string url = "")
        {
            return View(new LoginVo() { FromUrl = url });
        }
        //###??? 將來利用 ActionFilterAttribute 執行呼叫Controller前 檢查 WebScket的狀態
        //[XgProgAuth] 參考HrAdm的LeaveController

        public string CheckData(string key)
        {
            BaseUserDto data = _Fun.GetBaseUser();
            if (data == null) 
            {
                return "NG";
            }
            return "OK";
        }

        [HttpPost]
        public async Task<ActionResult> Login(LoginVo vo)
        {
            if (vo == null) { vo = new LoginVo(); }
            #region 1.check input account & password
            //reset UI msg
            vo.AccountMsg = "";
            vo.PwdMsg = "";

            if (_Str.IsEmpty(vo.Account))
            {
                vo.AccountMsg = "field is required.";
                goto lab_exit;
            }
            if (_Str.IsEmpty(vo.Pwd))
            {
                vo.PwdMsg = "field is required.";
                goto lab_exit;
            }
            #endregion

            #region 2.check DB password & get user info
            var sql = @"
            select u.Id as UserId, u.Name as UserName, u.Pwd,u.UserNO,
                u.DeptId, d.Name as DeptName
            from SoftNetMainDB.[dbo].[User] u
            join SoftNetMainDB.[dbo].[Dept] d on u.DeptId=d.Id
            where u.Account=@Account
            ";
            DBADO db = new DBADO("1", _Fun.Config.Db);
            var row = db.DB_TestROW(sql, new List<object>() { "Account", vo.Account });
            db.Dispose();
            //var row = await _Db.GetJsonAsync(sql, new List<object>() { "Account", vo.Account });
            //TODO: encode if need
            //if (row == null || row["Pwd"].ToString() != _Str.Md5(vo.Pwd))
            if (row == null || row["Pwd"].ToString() != vo.Pwd)
            {
                vo.AccountMsg = "input wrong.";
                goto lab_exit;
            }
            #endregion

            #region 3.set base user info
            var userId = row["UserId"].ToString().Trim();
            var userInfo = new BaseUserDto()
            {
                UserId = userId,
                UserName = row["UserName"].ToString(),
                UserNO = row["UserNO"].ToString(),
                DeptId = row["DeptId"].ToString(),
                DeptName = row["DeptName"].ToString(),
                Locale = _Fun.Config.Locale,
                ProgAuthStrs = await _XgProg.GetAuthStrsAsync(userId),
                IsLogin = true,
            };
            #endregion

            //4.set session of base user info
            _Http.GetSession().Set(_Fun.BaseUser, userInfo);   //extension method

            //5.redirect if need
            var url = _Str.IsEmpty(vo.FromUrl) ? "/Home/Index" : vo.FromUrl;
            return Redirect(url);

        lab_exit:
            return View(vo);
        }

        public ActionResult Logout()
        {
            _Http.GetSession().Clear();
            return Redirect("/Home/Index");
        }

        public ActionResult Privacy()
        {
            return View();
        }

        public ActionResult Error()
        {
            var error = HttpContext.Features.Get<IExceptionHandlerFeature>();
            return View("Error", (error == null) ? _Fun.SystemError : error.Error.Message);
        }
    }
}
