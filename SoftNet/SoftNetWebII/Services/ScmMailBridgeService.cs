using Base;
using Base.Services;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    /// <summary>
    /// SCM Mail Bridge 核心服務：
    /// 1) 處理 Outbox 寄件與重送
    /// 2) 從 Mail 收件後落地 Inbox
    /// 3) 依 MessageType 轉入對應業務表
    /// </summary>
    public class ScmMailBridgeService
    {
        // ParseStatus：Inbox 解析結果
        private const string ParseOk = "0";
        private const string ParseInvalidSubject = "1";
        private const string ParseInvalidJson = "2";
        private const string ParseNotTarget = "3";

        // ProcessStatus：Inbox 業務處理結果
        private const string ProcessPending = "0";
        private const string ProcessDone = "1";
        private const string ProcessFailed = "2";
        private const string ProcessDuplicated = "3";

        private readonly object _lock = new object();

        /// <summary>
        /// SCM 一次輪巡入口（由 RUNTimeServer 週期呼叫）
        /// </summary>
        public void RunCycle(DBADO db, CancellationToken token)
        {
            if (!IsEnabled())
            {
                return;
            }

            lock (_lock)
            {
                SendPendingMessages(db);
                FetchInboxMessages(db, token);
                ProcessInboxMessages(db);
            }
        }

        public void EnqueueCreateNeed(DBADO db, string docType, string docNumberNo, string fromServerId, string toServerId,
            string needIdA, string simulationIdA, string partNo, int qty, string arrivalDate, string remark)
        {
            if (!IsEnabled() || string.IsNullOrWhiteSpace(toServerId))
            {
                return;
            }

            string docTypeNorm = (docType ?? "").Trim().ToUpperInvariant();
            if (docTypeNorm != "DOC1" && docTypeNorm != "DOC4")
            {
                return;
            }

            // 基礎去重：同來源/目的/單據/Need 對應已存在時不重複入 Outbox（COUNT 結果須用 DB_GetQueryCount 正確讀取彙總值）
            string dedupeSql = $@"
SELECT COUNT(1)
FROM SoftNetLogDB.[dbo].[SCM_MailOutbox]
WHERE MessageType='CreateNeed'
  AND FromServerId=N'{EscapeSql(fromServerId)}'
  AND ToServerId=N'{EscapeSql(toServerId)}'
  AND Subject LIKE N'%Type:CreateNeed%'
  AND BodyJson LIKE N'%""docNumberNO"":""{EscapeSql(docNumberNo)}""%'
  AND BodyJson LIKE N'%""needIdA"":""{EscapeSql(needIdA)}""%'
  AND BodyJson LIKE N'%""simulationIdA"":""{EscapeSql(simulationIdA)}""%'";
            if (db.DB_GetQueryCount(dedupeSql) > 0)
            {
                return;
            }

            string messageId = $"{fromServerId}{Guid.NewGuid():N}".Substring(0, 32);
            JObject body = new JObject
            {
                ["messageId"] = messageId,
                ["messageType"] = "CreateNeed",
                ["fromServerId"] = fromServerId,
                ["toServerId"] = toServerId,
                ["eventTime"] = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz"),
                ["docType"] = docTypeNorm,
                ["docNumberNO"] = docNumberNo,
                ["needMapping"] = new JObject
                {
                    ["needIdA"] = needIdA,
                    ["simulationIdA"] = simulationIdA,
                    ["needIdB"] = ""
                },
                ["payload"] = new JObject
                {
                    ["partNO"] = partNo,
                    ["qty"] = qty,
                    ["arrivalDate"] = arrivalDate,
                    ["remark"] = remark
                },
                ["version"] = 1,
                ["signature"] = ""
            };

            string subject = BuildSubject(fromServerId, toServerId, "CreateNeed", messageId, 1);
            db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[SCM_MailOutbox]
(MessageId,FromServerId,ToServerId,MessageType,Subject,BodyJson,[Status],RetryCount,NextRetryTime,CreatedTime)
VALUES
(N'{EscapeSql(messageId)}',N'{EscapeSql(fromServerId)}',N'{EscapeSql(toServerId)}',N'CreateNeed',
N'{EscapeSql(subject)}',N'{EscapeSql(body.ToString(Formatting.None))}','0',0,NULL,'{DateTime.Now:yyyy-MM-dd HH:mm:ss}')");
        }

        private bool IsEnabled()
        {
            return _Fun.Config.SCM_IsEnabled && !string.IsNullOrWhiteSpace(_Fun.Config.SCM_DataExchange_Mail);
        }

        private void SendPendingMessages(DBADO db)
        {
            int maxRetry = Math.Max(1, _Fun.Config.SCM_MaxRetryCount);
            string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string sql = $@"
SELECT TOP 30 * 
FROM SoftNetLogDB.[dbo].[SCM_MailOutbox]
WHERE [Status] IN ('0','2')
  AND (NextRetryTime IS NULL OR NextRetryTime<='{now}')
ORDER BY CreatedTime";

            DataTable dt = db.DB_GetData(sql);
            if (dt == null || dt.Rows.Count == 0)
            {
                return;
            }

            // 將待寄與可重送訊息送至共用交換信箱
            foreach (DataRow row in dt.Rows)
            {
                string id = row["Id"].ToString();
                string subject = row["Subject"].ToString();
                string body = row["BodyJson"].ToString();
                string messageId = row["MessageId"].ToString();
                int retryCount = int.TryParse(row["RetryCount"].ToString(), out int rc) ? rc : 0;

                try
                {
                    _Fun.Mail_Send(new[] { _Fun.Config.SCM_DataExchange_Mail }, Array.Empty<string>(), subject, body, null, false);
                    db.DB_SetData($@"UPDATE SoftNetLogDB.[dbo].[SCM_MailOutbox]
SET [Status]='1',SentTime='{DateTime.Now:yyyy-MM-dd HH:mm:ss}',ErrorMessage=NULL
WHERE Id={id}");
                }
                catch (Exception ex)
                {
                    int nextRetryCount = retryCount + 1;
                    // 超過次數後改為 Cancelled(3)，避免無限重送
                    string nextStatus = nextRetryCount > maxRetry ? "3" : "0";
                    string nextRetryTime = DateTime.Now.AddMinutes(Math.Max(1, _Fun.Config.SCM_WorkingDoWhileTime)).ToString("yyyy-MM-dd HH:mm:ss");
                    db.DB_SetData($@"UPDATE SoftNetLogDB.[dbo].[SCM_MailOutbox]
SET [Status]='{nextStatus}',
    RetryCount={nextRetryCount},
    NextRetryTime='{nextRetryTime}',
    ErrorMessage=N'{EscapeSql($"MessageId={messageId}; {ex.Message}")}'
WHERE Id={id}");
                }
            }
        }

        private void FetchInboxMessages(DBADO db, CancellationToken token)
        {
            string server = string.IsNullOrWhiteSpace(_Fun.Config.SCM_ImapServer) ? "imap.gmail.com" : _Fun.Config.SCM_ImapServer.Trim();
            int port = _Fun.Config.SCM_ImapPort <= 0 ? 993 : _Fun.Config.SCM_ImapPort;
            if (string.IsNullOrWhiteSpace(_Fun.Config.MailCredentialsAccount) || string.IsNullOrWhiteSpace(_Fun.Config.MailCredentialsPWD))
            {
                return;
            }

            try
            {
                using var client = new ImapClient();
                client.Connect(server, port, true, token);
                client.Authenticate(_Fun.Config.MailCredentialsAccount, _Fun.Config.MailCredentialsPWD, token);
                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite, token);

                // 以「上次最後收件時間」做增量收信，避免大量未讀時因固定上限造成延遲/遺漏
                DateTime? lastReceiveTime = GetLastInboxReceiveTime(db);
                IList<UniqueId> ids;
                if (lastReceiveTime.HasValue)
                {
                    DateTime fetchFrom = lastReceiveTime.Value.AddMinutes(-5);
                    ids = inbox.Search(SearchQuery.NotSeen.And(SearchQuery.DeliveredAfter(fetchFrom)), token);

                    // 郵件伺服器時間可能有誤差，若增量結果為空則回補掃描全部未讀，避免漏信
                    if (ids == null || ids.Count == 0)
                    {
                        ids = inbox.Search(SearchQuery.NotSeen, token);
                    }
                }
                else
                {
                    ids = inbox.Search(SearchQuery.NotSeen, token);
                }

                if (ids == null || ids.Count == 0)
                {
                    client.Disconnect(true, token);
                    return;
                }

                // 已掃過的 IMAP UID 不再進入處理（略過的非 SCM、非本機等皆會推進游標，無須標已讀）
                long lastImapUid = GetLastProcessedImapUid(db);
                List<UniqueId> uidList = ids
                    .Where(u => (long)u.Id > lastImapUid)
                    .OrderBy(u => (long)u.Id)
                    .ToList();
                if (uidList.Count == 0)
                {
                    client.Disconnect(true, token);
                    return;
                }

                foreach (UniqueId uid in uidList)
                {
                    try
                    {
                    var message = inbox.GetMessage(uid, token);
                    string subject = message.Subject ?? "";
                    // 非 SCM 主旨前綴者略過（游標仍推進，下輪不會再進本 foreach）
                    string subjectPrefix = string.IsNullOrWhiteSpace(_Fun.Config.SCM_MailSubjectPrefix)
                        ? "[SoftNetSCM]"
                        : _Fun.Config.SCM_MailSubjectPrefix.Trim();
                    if (string.IsNullOrWhiteSpace(subject) || !subject.StartsWith(subjectPrefix, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    SubjectParseResult parse = ParseSubject(subject);
                    string bodyJson = message.TextBody;
                    if (string.IsNullOrWhiteSpace(bodyJson))
                    {
                        bodyJson = message.HtmlBody ?? "";
                    }


                    string parseStatus = ParseOk;
                    string processStatus = ProcessPending;
                    string error = "";
                    string messageId = parse.MessageId;
                    string fromServerId = parse.FromServerId;
                    string toServerId = parse.ToServerId;
                    string messageType = parse.MessageType;
                    DateTime receiveTime = message.Date != default ? message.Date.LocalDateTime : DateTime.Now;

                    // 先驗主旨，再驗目標 Server，再驗 Body JSON
                    if (!parse.IsValid)
                    {
                        parseStatus = ParseInvalidSubject;
                        processStatus = ProcessFailed;
                        error = parse.ErrorMessage;
                    }
                    else if (!string.Equals(toServerId, _Fun.Config.ServerId, StringComparison.OrdinalIgnoreCase))
                    {
                        // 非本機目標直接略過，不落地 Inbox
                        continue;
                    }
                    else
                    {
                        if (!TryNormalizeBodyJson(bodyJson, ref messageId, out string normalizedJson, out string bodyErr))
                        {
                            parseStatus = ParseInvalidJson;
                            processStatus = ProcessFailed;
                            error = bodyErr;
                        }
                        else
                        {
                            bodyJson = normalizedJson;
                        }
                    }

                    // 主旨合法且目標為本機的郵件，才會落地 Inbox
                    string insertSql = $@"
INSERT INTO SoftNetLogDB.[dbo].[SCM_MailInbox]
(
    MessageId,FromServerId,ToServerId,MessageType,Subject,BodyJson,ReceiveTime,ParseStatus,ProcessStatus,ErrorMessage
)
VALUES
(
N'{EscapeSql(messageId)}',N'{EscapeSql(fromServerId)}',N'{EscapeSql(toServerId)}',N'{EscapeSql(messageType)}',
N'{EscapeSql(subject)}',N'{EscapeSql(bodyJson)}','{receiveTime:yyyy-MM-dd HH:mm:ss}',
'{parseStatus}','{processStatus}',{(string.IsNullOrWhiteSpace(error) ? "NULL" : $"N'{EscapeSql(error)}'")}
)";
                    db.DB_SetData(insertSql);
                    inbox.AddFlags(uid, MessageFlags.Seen, true, token);
                    }
                    finally
                    {
                        // 每封處理過的郵件皆更新 UID 游標，增量條件與 Inbox 無關，略過的信不會卡住「上次時間」
                        UpsertPollImapUidMax(db, (long)uid.Id);
                    }
                }

                client.Disconnect(true, token);
            }
            catch (Exception ex)
            {
                _ = _Log.ErrorAsync($"ScmMailBridgeService.FetchInboxMessages Exception: {ex.Message}", false);
            }
        }

        /// <summary>
        /// 本機已掃過的最大 IMAP UniqueId；用於下輪排除已略過之郵件（非 SCM 主旨等），避免重複進入處理迴圈。
        /// 若尚未執行建表腳本（無 SCM_MailPollState），回傳 0 不拋錯。
        /// </summary>
        private long GetLastProcessedImapUid(DBADO db)
        {
            string sid = SqlServerId2(_Fun.Config.ServerId);
            DataRow dr = db.DB_GetFirstDataByDataRow($@"
IF OBJECT_ID(N'SoftNetLogDB.dbo.SCM_MailPollState', N'U') IS NULL
    SELECT CAST(0 AS BIGINT) AS LastImapUniqueId
ELSE
    SELECT CAST(ISNULL((
        SELECT TOP 1 LastImapUniqueId FROM SoftNetLogDB.[dbo].[SCM_MailPollState] WHERE ServerId=N'{EscapeSql(sid)}'
    ), 0) AS BIGINT) AS LastImapUniqueId");
            if (dr == null || dr["LastImapUniqueId"] == DBNull.Value)
            {
                return 0;
            }
            return long.TryParse(dr["LastImapUniqueId"].ToString(), out long v) ? v : 0;
        }

        private void UpsertPollImapUidMax(DBADO db, long uidVal)
        {
            string sid = SqlServerId2(_Fun.Config.ServerId);
            db.DB_SetData($@"
IF OBJECT_ID(N'SoftNetLogDB.dbo.SCM_MailPollState', N'U') IS NOT NULL
BEGIN
    IF EXISTS (SELECT 1 FROM SoftNetLogDB.[dbo].[SCM_MailPollState] WHERE ServerId=N'{EscapeSql(sid)}')
        UPDATE SoftNetLogDB.[dbo].[SCM_MailPollState]
        SET LastImapUniqueId = CASE WHEN {uidVal} > LastImapUniqueId THEN {uidVal} ELSE LastImapUniqueId END,
            UpdatedTime = GETDATE()
        WHERE ServerId=N'{EscapeSql(sid)}'
    ELSE
        INSERT INTO SoftNetLogDB.[dbo].[SCM_MailPollState] (ServerId, LastImapUniqueId, UpdatedTime)
        VALUES (N'{EscapeSql(sid)}', {uidVal}, GETDATE())
END");
        }

        private static string SqlServerId2(string serverId)
        {
            string s = (serverId ?? "").Trim();
            if (s.Length >= 2)
            {
                return s.Substring(0, 2);
            }
            return s.PadRight(2);
        }

        /// <summary>
        /// 取得 Inbox 已落地的最後收件時間（做為下一輪 IMAP DeliveredAfter 增量起點）
        /// </summary>
        private static DateTime? GetLastInboxReceiveTime(DBADO db)
        {
            DataRow dr = db.DB_GetFirstDataByDataRow("SELECT MAX(ReceiveTime) AS LastReceiveTime FROM SoftNetLogDB.[dbo].[SCM_MailInbox]");
            if (dr == null || dr["LastReceiveTime"] == DBNull.Value)
            {
                return null;
            }

            if (DateTime.TryParse(dr["LastReceiveTime"].ToString(), out DateTime dt))
            {
                return dt;
            }
            return null;
        }

        private void ProcessInboxMessages(DBADO db)
        {
            DataTable dt = db.DB_GetData($@"
SELECT TOP 50 *
FROM SoftNetLogDB.[dbo].[SCM_MailInbox]
WHERE ParseStatus='{ParseOk}' AND ProcessStatus='{ProcessPending}'
ORDER BY ReceiveTime");
            if (dt == null || dt.Rows.Count == 0)
            {
                return;
            }

            foreach (DataRow row in dt.Rows)
            {
                string id = row["Id"].ToString();
                string messageId = row["MessageId"].ToString();
                string fromServerId = row["FromServerId"].ToString().Trim();
                string toServerId = row["ToServerId"].ToString().Trim();
                string type = row["MessageType"].ToString().Trim();
                string body = row["BodyJson"].ToString();

                try
                {
                    // 技術冪等：同 MessageId 已成功處理過，後續標為 Duplicated
                    if (db.DB_GetQueryCount($@"SELECT COUNT(1)
FROM SoftNetLogDB.[dbo].[SCM_MailInbox]
WHERE MessageId=N'{EscapeSql(messageId)}' AND ProcessStatus='{ProcessDone}' AND Id<>{id}") > 0)
                    {
                        db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SCM_MailInbox] SET ProcessStatus='{ProcessDuplicated}' WHERE Id={id}");
                        continue;
                    }

                    JObject json = JObject.Parse(body);
                    HandleByMessageType(db, type, json, fromServerId, toServerId);
                    db.DB_SetData($"UPDATE SoftNetLogDB.[dbo].[SCM_MailInbox] SET ProcessStatus='{ProcessDone}',ErrorMessage=NULL WHERE Id={id}");
                }
                catch (Exception ex)
                {
                    db.DB_SetData($@"UPDATE SoftNetLogDB.[dbo].[SCM_MailInbox]
SET ProcessStatus='{ProcessFailed}',ErrorMessage=N'{EscapeSql(ex.Message)}'
WHERE Id={id}");
                }
            }
        }

        private void HandleByMessageType(DBADO db, string messageType, JObject json, string fromServerId, string toServerId)
        {
            // 目前採最小可運作流程：
            // - NeedCreated 寫入關聯
            // - PromiseDate/ProdStatus/Exception 寫入快照
            // - 其餘關鍵訊息先做回覆佇列，便於流程驗證
            switch (messageType)
            {
                case "NeedCreated":
                    UpsertNeedLink(db, json, fromServerId, toServerId, "1");
                    break;
                case "PromiseDate":
                case "ProdStatus":
                case "Exception":
                    UpsertStatusSnapshot(db, json);
                    break;
                case "CreateNeed":
                    EnqueueReply(db, "PromiseDate", toServerId, fromServerId, json, null);
                    break;
                case "OKNeed":
                    EnqueueReply(db, "NeedCreated", toServerId, fromServerId, json, j => EnsureNeedIdB(j));
                    break;
                case "NGNeed":
                    EnqueueReply(db, "PromiseDate", toServerId, fromServerId, json, null);
                    break;
                case "CancelNeed":
                    EnqueueReply(db, "CloseNeed", toServerId, fromServerId, json, null);
                    break;
                case "AskStatus":
                    EnqueueReply(db, "ProdStatus", toServerId, fromServerId, json, null);
                    break;
                default:
                    EnqueueReply(db, "AckMsg", toServerId, fromServerId, json, null);
                    break;
            }
        }

        private void UpsertNeedLink(DBADO db, JObject json, string fromServerId, string toServerId, string linkStatus)
        {
            var mapping = json["needMapping"] as JObject ?? new JObject();
            string needIdA = mapping["needIdA"]?.ToString() ?? "";
            string simulationIdA = mapping["simulationIdA"]?.ToString() ?? "";
            string needIdB = mapping["needIdB"]?.ToString() ?? "";
            string docType = json["docType"]?.ToString() ?? "";
            string docNumberNo = json["docNumberNO"]?.ToString() ?? "";
            // 以 NeedId_A + NeedId_B 組成固定長度關聯鍵
            string linkId = BuildLinkId(needIdA, needIdB);

            if (db.DB_GetQueryCount($"SELECT COUNT(1) FROM SoftNetLogDB.[dbo].[SCM_NeedLink] WHERE LinkId=N'{EscapeSql(linkId)}'") > 0)
            {
                db.DB_SetData($@"UPDATE SoftNetLogDB.[dbo].[SCM_NeedLink]
SET NeedId_B=N'{EscapeSql(needIdB)}',LinkStatus='{EscapeSql(linkStatus)}',UpdatedTime='{DateTime.Now:yyyy-MM-dd HH:mm:ss}'
WHERE LinkId=N'{EscapeSql(linkId)}'");
            }
            else
            {
                db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[SCM_NeedLink]
(LinkId,FromServerId,ToServerId,NeedId_A,SimulationId_A,NeedId_B,DocType,DocNumberNO,LinkStatus,CreatedTime,UpdatedTime)
VALUES
(N'{EscapeSql(linkId)}',N'{EscapeSql(fromServerId)}',N'{EscapeSql(toServerId)}',N'{EscapeSql(needIdA)}',N'{EscapeSql(simulationIdA)}',
N'{EscapeSql(needIdB)}',N'{EscapeSql(docType)}',N'{EscapeSql(docNumberNo)}','{EscapeSql(linkStatus)}',
'{DateTime.Now:yyyy-MM-dd HH:mm:ss}','{DateTime.Now:yyyy-MM-dd HH:mm:ss}')");
            }
        }

        private void UpsertStatusSnapshot(DBADO db, JObject json)
        {
            var mapping = json["needMapping"] as JObject ?? new JObject();
            string needIdA = mapping["needIdA"]?.ToString() ?? "";
            string needIdB = mapping["needIdB"]?.ToString() ?? "";
            string linkId = BuildLinkId(needIdA, needIdB);
            string promiseDateSql = "NULL";
            if (DateTime.TryParse(json["payload"]?["promiseDate"]?.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime pDate))
            {
                promiseDateSql = $"'{pDate:yyyy-MM-dd HH:mm:ss}'";
            }

            string planState = json["payload"]?["state"]?.ToString() ?? json["payload"]?["planState"]?.ToString() ?? "";
            string reportProgress = json["payload"]?["reportProgress"]?.ToString() ?? "";
            string exceptionCode = json["payload"]?["exceptionCode"]?.ToString() ?? "";
            string exceptionRemark = json["payload"]?["exceptionRemark"]?.ToString() ?? "";

            // 快照採 append-only，保留歷史軌跡供追溯
            db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[SCM_StatusSnapshot]
(LinkId,NeedId_B,PromiseDate,PlanState,ReportProgress,ExceptionCode,ExceptionRemark,SnapshotTime)
VALUES
(N'{EscapeSql(linkId)}',N'{EscapeSql(needIdB)}',{promiseDateSql},N'{EscapeSql(planState)}',N'{EscapeSql(reportProgress)}',
{(string.IsNullOrWhiteSpace(exceptionCode) ? "NULL" : $"N'{EscapeSql(exceptionCode)}'")},
{(string.IsNullOrWhiteSpace(exceptionRemark) ? "NULL" : $"N'{EscapeSql(exceptionRemark)}'")},
'{DateTime.Now:yyyy-MM-dd HH:mm:ss}')");
        }

        private void EnqueueReply(DBADO db, string replyType, string fromServerId, string toServerId, JObject source, Action<JObject> beforeSave)
        {
            JObject body = (JObject)source.DeepClone();
            string messageId = $"{fromServerId}{Guid.NewGuid():N}".Substring(0, 32);
            body["messageId"] = messageId;
            body["messageType"] = replyType;
            body["fromServerId"] = fromServerId;
            body["toServerId"] = toServerId;
            body["eventTime"] = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz");
            if (beforeSave != null)
            {
                beforeSave(body);
            }

            string subject = BuildSubject(fromServerId, toServerId, replyType, messageId, 1);
            if (db.DB_GetQueryCount($"SELECT COUNT(1) FROM SoftNetLogDB.[dbo].[SCM_MailOutbox] WHERE MessageId=N'{EscapeSql(messageId)}'") > 0)
            {
                return;
            }

            db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[SCM_MailOutbox]
(MessageId,FromServerId,ToServerId,MessageType,Subject,BodyJson,[Status],RetryCount,NextRetryTime,CreatedTime)
VALUES
(N'{EscapeSql(messageId)}',N'{EscapeSql(fromServerId)}',N'{EscapeSql(toServerId)}',N'{EscapeSql(replyType)}',
N'{EscapeSql(subject)}',N'{EscapeSql(body.ToString(Formatting.None))}','0',0,NULL,'{DateTime.Now:yyyy-MM-dd HH:mm:ss}')");
        }

        private static bool TryNormalizeBodyJson(string bodyJson, ref string messageId, out string normalizedJson, out string err)
        {
            normalizedJson = bodyJson ?? "";
            err = "";
            try
            {
                JObject json = JObject.Parse(normalizedJson);
                // 依 PRD 最低必要欄位檢查
                string[] required = { "messageType", "fromServerId", "toServerId", "eventTime", "docType", "docNumberNO", "needMapping", "payload" };
                foreach (string key in required)
                {
                    if (json[key] == null)
                    {
                        err = $"Missing field: {key}";
                        return false;
                    }
                }

                if (json["messageId"] != null)
                {
                    messageId = json["messageId"]!.ToString();
                }
                normalizedJson = json.ToString(Formatting.None);
                return true;
            }
            catch (Exception ex)
            {
                err = ex.Message;
                return false;
            }
        }

        private SubjectParseResult ParseSubject(string subject)
        {
            string prefix = string.IsNullOrWhiteSpace(_Fun.Config.SCM_MailSubjectPrefix) ? "[SoftNetSCM]" : _Fun.Config.SCM_MailSubjectPrefix;
            if (string.IsNullOrWhiteSpace(subject) || !subject.StartsWith(prefix, StringComparison.Ordinal))
            {
                return SubjectParseResult.Fail("Subject prefix is invalid.");
            }

            Dictionary<string, string> map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string part in subject.Split('|'))
            {
                int idx = part.IndexOf(':');
                if (idx <= 0)
                {
                    continue;
                }
                map[part.Substring(0, idx)] = part[(idx + 1)..];
            }

            if (!map.TryGetValue("From", out string from) ||
                !map.TryGetValue("To", out string to) ||
                !map.TryGetValue("Type", out string type) ||
                !map.TryGetValue("Msg", out string msg))
            {
                return SubjectParseResult.Fail("Subject keys are incomplete.");
            }

            return SubjectParseResult.Ok(msg.Trim(), from.Trim(), to.Trim(), type.Trim());
        }

        private static string BuildSubject(string fromServerId, string toServerId, string type, string messageId, int version)
        {
            string prefix = string.IsNullOrWhiteSpace(_Fun.Config.SCM_MailSubjectPrefix) ? "[SoftNetSCM]" : _Fun.Config.SCM_MailSubjectPrefix.Trim();
            return $"{prefix}|From:{fromServerId}|To:{toServerId}|Type:{type}|Msg:{messageId}|V:{version}";
        }

        private static string EscapeSql(string value)
        {
            return (value ?? "").Replace("'", "''");
        }

        private static string BuildLinkId(string needIdA, string needIdB)
        {
            string a = (needIdA ?? "").PadRight(6, '0');
            string b = (needIdB ?? "").PadRight(6, '0');
            return $"{a[..6]}{b[..6]}";
        }

        private static void EnsureNeedIdB(JObject body)
        {
            JObject mapping = body["needMapping"] as JObject ?? new JObject();
            if (string.IsNullOrWhiteSpace(mapping["needIdB"]?.ToString()))
            {
                mapping["needIdB"] = _Str.NewId('Y');
            }
            body["needMapping"] = mapping;
        }

        private sealed class SubjectParseResult
        {
            public bool IsValid { get; init; } = false;
            public string MessageId { get; init; } = "";
            public string FromServerId { get; init; } = "";
            public string ToServerId { get; init; } = "";
            public string MessageType { get; init; } = "";
            public string ErrorMessage { get; init; } = "";

            public static SubjectParseResult Ok(string messageId, string fromServerId, string toServerId, string messageType)
                => new SubjectParseResult
                {
                    IsValid = true,
                    MessageId = messageId,
                    FromServerId = fromServerId,
                    ToServerId = toServerId,
                    MessageType = messageType
                };

            public static SubjectParseResult Fail(string err)
                => new SubjectParseResult { IsValid = false, ErrorMessage = err };
        }
    }
}
