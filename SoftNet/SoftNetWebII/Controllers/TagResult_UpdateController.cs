using Base;
using Base.Services;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Models;
using System;

namespace SoftNetWebII.Controllers
{
    public class TagResult_UpdateController : Controller
    {
        public IActionResult Index([FromBody] API_UpdateTagResult input)
        {
            DateTime now = DateTime.Now;
            if (!input.result)
            {
                using (DBADO db = new DBADO("1", _Fun.Config.Db))
                {
                    db.DB_SetData($"INSERT INTO SoftNetLogDB.[dbo].[LabelStateLog] ([Id],[macID],[LOGDateTime],[ReceiveType],[INFO]) VALUES ('{_Str.NewId('L')}','{input.mac}','{now.ToString("yyyy/MM/dd HH:mm:ss.fff")}','接收發送失敗','{input.message}')");
                }
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"TagResult_UpdateController.cs 標籤傳送失敗: macID={input.mac} message={input.message}", true);
            }
            return StatusCode(200);
        }
    }
}
