using Base;
using Base.Services;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    /// <summary>
    /// 需求暫停 (APS_NeedData.State='P') 時，於現場作業成功後背景解除暫停並重算排程。
    /// 觸發條件：僅 NeedData.State='P'（不再要求 PP_WorkOrder.PauseTime）。
    /// 失敗：ErrorAsync Mail，並將 State 還原為 P、還原 PauseTime（若先前有值）。
    /// </summary>
    public static class ApsNeedPauseResumeQueue
    {
        private static string SqlSafe(string s) => (s ?? "").Replace("'", "''");

        /// <summary>由 Manufacture 單筆工站取得 NeedId 後，若 State=P 則排程背景處理。</summary>
        public static void TryEnqueueFromManufactureStation(DBADO dbSync, string stationNo, string triggerAction)
        {
            if (string.IsNullOrWhiteSpace(stationNo)) return;
            var dr = dbSync.DB_GetFirstDataByDataRow(
                $"SELECT OrderNO,SimulationId FROM SoftNetMainDB.[dbo].[Manufacture] where ServerId='{SqlSafe(_Fun.Config.ServerId)}' and StationNO='{SqlSafe(stationNo)}'");
            if (dr == null) return;
            string orderNo = dr["OrderNO"]?.ToString().Trim() ?? "";
            string simulationId = dr["SimulationId"]?.ToString().Trim() ?? "";
            TryEnqueueFromOrderAndSimulation(dbSync, orderNo, simulationId, stationNo, triggerAction);
        }

        /// <summary>由 ManufactureII 列（多工）取得 NeedId；dedupe 避免同一 Need 重複排程。</summary>
        public static void TryEnqueueFromManufactureII(DBADO dbSync, HashSet<string> dedupe, DataRow drMII, string triggerAction)
        {
            if (drMII == null || dedupe == null) return;
            string orderNo = drMII["OrderNO"]?.ToString().Trim() ?? "";
            string simulationId = drMII["SimulationId"]?.ToString().Trim() ?? "";
            string stationNo = drMII["StationNO"]?.ToString().Trim() ?? "";
            TryEnqueueFromOrderAndSimulation(dbSync, orderNo, simulationId, stationNo, triggerAction, dedupe);
        }

        /// <summary>
        /// 【入隊入口】現場作業成功後呼叫：檢查該工單對應需求是否為暫停(State=P)，若是則丟到背景執行解除暫停與排程延後重算。
        /// 本方法本身不做重算，只做前置過濾後 Task.Run 呼叫 ExecuteBackground。
        /// </summary>
        /// <param name="dbSync">呼叫端當前 DB 連線（僅用於同步查詢，背景作業會另開連線）</param>
        /// <param name="orderNo">工單編號（PP_WorkOrder.OrderNO）</param>
        /// <param name="simulationId">目前作業的子排程編號（供 RefreshRunSetSimulation 當觸發點）</param>
        /// <param name="stationNo">觸發工站（僅供 Log）</param>
        /// <param name="triggerAction">觸發來源說明，例如 LabelWork_報工、STView2Work_智慧開工（僅供 Log）</param>
        /// <param name="dedupe">可選。同一請求內避免同一 NeedId 重複入隊（多工站一次勾選多筆時使用）</param>
        public static void TryEnqueueFromOrderAndSimulation(DBADO dbSync, string orderNo, string simulationId, string stationNo, string triggerAction, HashSet<string> dedupe = null)
        {
            // 缺工單或 SimulationId 無法定位需求／重算觸發點 → 直接略過
            if (string.IsNullOrWhiteSpace(orderNo) || string.IsNullOrWhiteSpace(simulationId)) return;

            // 工單 → NeedId（APS_NeedData.Id）
            var wo = dbSync.DB_GetFirstDataByDataRow(
                $"SELECT NeedId FROM SoftNetSYSDB.[dbo].[PP_WorkOrder] where ServerId='{SqlSafe(_Fun.Config.ServerId)}' and OrderNO='{SqlSafe(orderNo)}'");
            if (wo == null || wo.IsNull("NeedId") || string.IsNullOrWhiteSpace(wo["NeedId"]?.ToString())) return;
            string needId = wo["NeedId"].ToString();

            // 同一批次已排過此 NeedId 則略過（HashSet.Add 回傳 false 表示已存在）
            if (dedupe != null && !dedupe.Add(needId)) return;

            // 僅處理「暫停中」的需求；State≠P（例如已是 6 正式生產）不動作
            var nd = dbSync.DB_GetFirstDataByDataRow(
                $"SELECT State FROM SoftNetSYSDB.[dbo].[APS_NeedData] where ServerId='{SqlSafe(_Fun.Config.ServerId)}' and Id='{SqlSafe(needId)}'");
            if (nd == null || nd["State"]?.ToString() != "P") return;

            // 背景執行：不阻塞現場作業回應；連線字串與參數先複製，避免呼叫端 db 關閉後背景仍用到
            string serverId = _Fun.Config.ServerId;
            string connStr = _Fun.Config.Db;
            _ = Task.Run(() => ExecuteBackground(connStr, serverId, needId, orderNo, stationNo ?? "", simulationId, triggerAction ?? ""));
        }

        /// <summary>
        /// 【背景執行】解除需求暫停並做局部排程重算（RefreshRunSetSimulation）。
        /// 流程：
        ///   1) 交易內取得 applock，避免同一 NeedId 並發解除暫停
        ///   2) 再次確認 State 仍為 P（雙重檢查，避免入隊後狀態已變）
        ///   3) 記住工單 PauseTime（失敗時補償還原用）
        ///   4) APS_NeedData：P → 6（恢復正式生產）；PP_WorkOrder.PauseTime 清空
        ///   5) 交易 Commit 後，呼叫 RefreshRunSetSimulation 依 simulationId 做局部延後重算
        ///   6) 更新 NeedSimulationDate 為尚未結案工序的最大 SimulationDate
        /// 失敗：CompensateRevertToPause 將 State 還原為 P，並還原 PauseTime。
        /// 注意：此路徑不會 DELETE 整個 NeedId 的 APS_Simulation，僅局部調整時間。
        /// </summary>
        private static void ExecuteBackground(string connStr, string serverId, string needId, string orderNo, string stationNo, string simulationId, string triggerAction)
        {
            SFC_Common sfc = null;
            DateTime? pauseBefore = null; // 解除前的 PauseTime；例外時用來補償還原
            try
            {
                sfc = new SFC_Common("1", _Fun.Config.Db);

                // ---------- 階段 A：交易內解除暫停（State P→6）----------
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    using (var tran = conn.BeginTransaction())
                    {
                        // 以 NeedId 為資源名稱加排他鎖，避免多工站同時觸發造成重複解除／重算
                        string lockName = $"APS_NR_{serverId}_{needId}";
                        if (lockName.Length > 255) lockName = lockName.Substring(0, 255);

                        int lockRet;
                        using (var cmdLock = new SqlCommand("DECLARE @r int; EXEC @r = sp_getapplock @Resource=@res, @LockMode=N'Exclusive', @LockOwner=N'Transaction', @LockTimeout=60000; SELECT @r;", conn, tran))
                        {
                            cmdLock.Parameters.Add("@res", SqlDbType.NVarChar, 255).Value = lockName;
                            lockRet = Convert.ToInt32(cmdLock.ExecuteScalar());
                        }
                        // sp_getapplock：>=0 成功；<0 失敗（逾時／取消等）
                        if (lockRet < 0)
                        {
                            tran.Rollback();
                            _ = _Log.ErrorAsync($"ApsNeedPauseResumeQueue 無法取得 applock NeedId={needId} OrderNO={orderNo} Station={stationNo} Trigger={triggerAction} lockRet={lockRet}", true);
                            return;
                        }

                        // 雙重檢查：入隊時是 P，執行時可能已被其他程序改掉
                        using (var cmdState = new SqlCommand(
                            "SELECT State FROM SoftNetSYSDB.dbo.APS_NeedData WHERE ServerId=@sid AND Id=@nid", conn, tran))
                        {
                            cmdState.Parameters.AddWithValue("@sid", serverId);
                            cmdState.Parameters.AddWithValue("@nid", needId);
                            var st = cmdState.ExecuteScalar()?.ToString();
                            if (st != "P")
                            {
                                tran.Rollback();
                                return;
                            }
                        }

                        // 備份 PauseTime，供 catch 補償還原
                        using (var cmdPt = new SqlCommand(
                            "SELECT PauseTime FROM SoftNetSYSDB.dbo.PP_WorkOrder WHERE ServerId=@sid AND OrderNO=@ord", conn, tran))
                        {
                            cmdPt.Parameters.AddWithValue("@sid", serverId);
                            cmdPt.Parameters.AddWithValue("@ord", orderNo);
                            var o = cmdPt.ExecuteScalar();
                            if (o != null && o != DBNull.Value) pauseBefore = Convert.ToDateTime(o);
                        }

                        // 條件更新：僅當仍為 P 才改成 6；影響列數 0 表示已被其他程序搶先處理
                        using (var cmdUpNd = new SqlCommand(
                            "UPDATE SoftNetSYSDB.dbo.APS_NeedData SET State=N'6', StateINFO=NULL, UpdateTime=GETDATE() WHERE ServerId=@sid AND Id=@nid AND State=N'P'", conn, tran))
                        {
                            cmdUpNd.Parameters.AddWithValue("@sid", serverId);
                            cmdUpNd.Parameters.AddWithValue("@nid", needId);
                            if (cmdUpNd.ExecuteNonQuery() == 0)
                            {
                                tran.Rollback();
                                return;
                            }
                        }

                        // 清除工單暫停時間標記
                        using (var cmdUpWo = new SqlCommand(
                            "UPDATE SoftNetSYSDB.dbo.PP_WorkOrder SET PauseTime=NULL WHERE ServerId=@sid AND OrderNO=@ord", conn, tran))
                        {
                            cmdUpWo.Parameters.AddWithValue("@sid", serverId);
                            cmdUpWo.Parameters.AddWithValue("@ord", orderNo);
                            cmdUpWo.ExecuteNonQuery();
                        }

                        tran.Commit(); // Commit 後 applock 隨交易釋放
                    }
                }

                // ---------- 階段 B：局部排程重算（交易外，避免長時間鎖住）----------
                using (var db = new DBADO("1", _Fun.Config.Db))
                {
                    var webList = new List<string>();
                    var swRefresh = System.Diagnostics.Stopwatch.StartNew();
                    // data[1]=simulationId：以該子排程為觸發點，延後後續工序時間（非整批 DELETE 重建）
                    bool refreshOk = sfc.RefreshRunSetSimulation(db, new[] { "", simulationId }, ref webList);
                    swRefresh.Stop();
                    if (!refreshOk)
                    {
                        // 重算失敗：State 已是 6，僅記 Log，不自動還原（避免現場已恢復生產又被改回暫停）
                        _ = _Log.ErrorAsync($"ApsNeedPauseResumeQueue RefreshRunSetSimulation 回傳 false (暫停狀態未恢復,僅記Log) NeedId={needId} SimulationId={simulationId} OrderNO={orderNo} Station={stationNo} Trigger={triggerAction} Elapsed={swRefresh.ElapsedMilliseconds}ms", false);
                    }
                    else if (swRefresh.ElapsedMilliseconds > 10000)
                    {
                        _ = _Log.ErrorAsync($"ApsNeedPauseResumeQueue RefreshRunSetSimulation 耗時過長 NeedId={needId} SimulationId={simulationId} Elapsed={swRefresh.ElapsedMilliseconds}ms Trigger={triggerAction}", false);
                    }

                    // 同步需求預計完成日：取尚未結案(IsOK≠1)工序的最大 SimulationDate
                    string mxSql = $@"SELECT MAX(SimulationDate) AS mx FROM SoftNetSYSDB.dbo.APS_Simulation 
                        WHERE NeedId='{SqlSafe(needId)}' AND IsOK <> N'1'";
                    var mxRow = db.DB_GetFirstDataByDataRow(mxSql);
                    if (mxRow != null && !mxRow.IsNull("mx"))
                    {
                        string mx = Convert.ToDateTime(mxRow["mx"]).ToString("yyyy-MM-dd HH:mm:ss");
                        db.DB_SetData($"UPDATE SoftNetSYSDB.dbo.APS_NeedData SET NeedSimulationDate='{mx}', UpdateTime=GETDATE() WHERE ServerId='{SqlSafe(serverId)}' AND Id='{SqlSafe(needId)}'");
                    }
                }
            }
            catch (Exception ex)
            {
                // 例外時補償：State 還原為 P，並還原 PauseTime（若解除前有值）
                CompensateRevertToPause(connStr, serverId, needId, orderNo, pauseBefore);
                _ = _Log.ErrorAsync(
                    $"ApsNeedPauseResumeQueue 失敗 NeedId={needId} OrderNO={orderNo} Station={stationNo} SimulationId={simulationId} Trigger={triggerAction} {ex.Message} {ex.StackTrace}", true);
            }
            finally
            {
                sfc?.Dispose();
            }
        }

        private static void CompensateRevertToPause(string connStr, string serverId, string needId, string orderNo, DateTime? pauseBefore)
        {
            try
            {
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(
                        "UPDATE SoftNetSYSDB.dbo.APS_NeedData SET State=N'P', UpdateTime=GETDATE() WHERE ServerId=@sid AND Id=@nid", conn))
                    {
                        cmd.Parameters.AddWithValue("@sid", serverId);
                        cmd.Parameters.AddWithValue("@nid", needId);
                        cmd.ExecuteNonQuery();
                    }
                    if (pauseBefore.HasValue)
                    {
                        using (var cmd2 = new SqlCommand(
                            "UPDATE SoftNetSYSDB.dbo.PP_WorkOrder SET PauseTime=@pt WHERE ServerId=@sid AND OrderNO=@ord", conn))
                        {
                            cmd2.Parameters.AddWithValue("@sid", serverId);
                            cmd2.Parameters.AddWithValue("@ord", orderNo);
                            cmd2.Parameters.AddWithValue("@pt", pauseBefore.Value);
                            cmd2.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex2)
            {
                _ = _Log.ErrorAsync($"ApsNeedPauseResumeQueue 補償還原 P 失敗 NeedId={needId} OrderNO={orderNo} {ex2.Message} {ex2.StackTrace}", true);
            }
        }
    }
}
