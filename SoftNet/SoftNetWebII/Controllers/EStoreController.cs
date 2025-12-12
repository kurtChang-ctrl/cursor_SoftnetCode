using Base;
using Base.Services;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Models;
using SoftNetWebII.Services;

namespace SoftNetWebII.Controllers
{
    public class EStoreController : Controller
    {
        private SNWebSocketService _WebSocket = null;
        private SFC_Common _SFC_Common = null;
        public EStoreController(SNWebSocketService websocket, SFC_Common sfc_Common)
        {
            if (_WebSocket == null)
            {
                _WebSocket = websocket;
            }
            if (_SFC_Common == null)
            {
                _SFC_Common = sfc_Common;
            }
        }
        public ActionResult CallBackINFO([FromBody] API_EStore_CallBack input)
        {
            if (input == null)
            {
                System.Threading.Tasks.Task task = _Log.ErrorAsync($"EStoreController.cs 接收關門失敗 Data=NULL", true);
                return StatusCode(404);
            }
            else
            {
                if (input.isFinished)
                {
                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                    {
                        string _s = "";

                        int count = input.result.Count;
                        foreach (API_EStore_Close item in input.result)
                        {
                            _s = $"{_s}/r/nX={item.x.ToString()},Y={item.y.ToString()},cabinetNo={item.cabinetNo},weight={item.cabinetNo.ToString()}";
                        }
                    }
                }
            }
            return StatusCode(200);
        }

    }
}
