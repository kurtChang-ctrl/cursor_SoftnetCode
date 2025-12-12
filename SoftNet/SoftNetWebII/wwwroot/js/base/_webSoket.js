
var ws;

var _webSocket = {

    Url: 'ws://127.0.0.1:8089/echo',//###???暫時寫死   'ws://localhost:8080/echo',
    init: function (url) {
        Url = url;
        this.connectWesocket();
    },
    isWork: function () {
        if (ws.readyState == 1) { return true; }
        else { return false; }
    },


    /**
     * initial forms(recursive)
     * param edit {object} EditOne/EditMany object
     */
    connectWesocket: function () {
        if ("WebSocket" in window) {
            //alert("您的浏览器支持 WebSocket!");
            ws = new WebSocket(Url);

            ws.onopen = function () {
                // Web Socket 已连接上，使用 send() 方法发送数据
                //ws.send("发送数据");
                //ws.send(JSON.stringify({x:254,y:100}));

                //###???將來改送RMSProtocol object
                //var RMSProtocol = new Object();
                //RMSProtocol.DataType = 0;
                //RMSProtocol.Data = 'Mustang';
                //RMSProtocol.PoolID = 0;
                //RMSProtocol.IsReturnID = false;
                //RMSProtocol.TransferID = 0;

                //ws.send(JSON.stringify({RMSProtocol}));
                ws.send(JSON.stringify({ DataType: 1, Data: "WebSocket_Login,SoftNet_I,aa", PoolID: 0, IsReturnID: "false", TransferID: 0 }));

            };

            ws.onmessage = function (evt) {
                var received_msg = evt.data;
                var cmd = received_msg.split(',')
                //debugger;
                switch (cmd[0]) {
                    case "StationStatusChange":

                        if (_me.hasStaionStatus) {
                            _crud.pageReload();
                        }
                        break;
                }
            };

            ws.onclose = function () {
                // 关闭 websocket
                alert("连接已关闭...");
            };
        }

        else {
            // 浏览器不支持 WebSocket
            alert("您的浏览器不支持 WebSocket!");
        }
    },
    Send: function (data) {
        if (this.isWork) {
            ws.send(JSON.stringify(data));
        }
        if (ws.readyState == 1) { return true; }
        else { return false; }
    },
};//class