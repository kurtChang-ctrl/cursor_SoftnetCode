USE [SoftNetLogDB];
GO

/* =========================================================
   SCM Mail 交換最小驗證 SQL
   用途：
   1) 確認 CreateNeed 是否成功進 Outbox
   2) 確認收件是否正確落地 Inbox
   3) 確認 NeedLink / StatusSnapshot 是否有資料
   ========================================================= */

DECLARE @TopN INT = 50;

PRINT N'1) Outbox 最新狀態';
SELECT TOP (@TopN)
    Id,
    MessageId,
    FromServerId,
    ToServerId,
    MessageType,
    [Status],         -- 0=Pending,1=Sent,2=Failed,3=Cancelled
    RetryCount,
    NextRetryTime,
    CreatedTime,
    SentTime,
    ErrorMessage
FROM dbo.SCM_MailOutbox
ORDER BY Id DESC;
GO

PRINT N'2) Inbox 最新狀態';
SELECT TOP (@TopN)
    Id,
    MessageId,
    FromServerId,
    ToServerId,
    MessageType,
    ParseStatus,      -- 0=OK,1=InvalidSubject,2=InvalidJson,3=NotTarget
    ProcessStatus,    -- 0=Pending,1=Done,2=Failed,3=Duplicated
    ReceiveTime,
    ErrorMessage
FROM dbo.SCM_MailInbox
ORDER BY Id DESC;
GO

PRINT N'3) NeedLink 最新關聯';
SELECT TOP (@TopN)
    LinkId,
    FromServerId,
    ToServerId,
    NeedId_A,
    SimulationId_A,
    NeedId_B,
    DocType,
    DocNumberNO,
    LinkStatus,       -- 0=Pending,1=Linked,2=Closed,3=Rejected
    CreatedTime,
    UpdatedTime
FROM dbo.SCM_NeedLink
ORDER BY UpdatedTime DESC, CreatedTime DESC;
GO

PRINT N'4) StatusSnapshot 最新快照';
SELECT TOP (@TopN)
    SnapshotId,
    LinkId,
    NeedId_B,
    PromiseDate,
    PlanState,
    ReportProgress,
    ExceptionCode,
    ExceptionRemark,
    SnapshotTime
FROM dbo.SCM_StatusSnapshot
ORDER BY SnapshotId DESC;
GO

PRINT N'5) 異常檢查：Outbox 失敗或取消';
SELECT TOP (@TopN)
    Id,
    MessageId,
    MessageType,
    [Status],
    RetryCount,
    ErrorMessage,
    CreatedTime,
    SentTime
FROM dbo.SCM_MailOutbox
WHERE [Status] IN ('2', '3')
ORDER BY Id DESC;
GO

PRINT N'6) 異常檢查：Inbox 解析/處理失敗';
SELECT TOP (@TopN)
    Id,
    MessageId,
    MessageType,
    ParseStatus,
    ProcessStatus,
    ErrorMessage,
    ReceiveTime
FROM dbo.SCM_MailInbox
WHERE ParseStatus IN ('1', '2')
   OR ProcessStatus IN ('2')
ORDER BY Id DESC;
GO

PRINT N'7) 異常檢查：重複訊息';
SELECT TOP (@TopN)
    MessageId,
    COUNT(1) AS Cnt
FROM dbo.SCM_MailInbox
GROUP BY MessageId
HAVING COUNT(1) > 1
ORDER BY Cnt DESC, MessageId;
GO

PRINT N'8) 快速 KPI：各狀態筆數';
SELECT
    (SELECT COUNT(1) FROM dbo.SCM_MailOutbox WHERE [Status]='0') AS Outbox_Pending,
    (SELECT COUNT(1) FROM dbo.SCM_MailOutbox WHERE [Status]='1') AS Outbox_Sent,
    (SELECT COUNT(1) FROM dbo.SCM_MailOutbox WHERE [Status]='2') AS Outbox_Failed,
    (SELECT COUNT(1) FROM dbo.SCM_MailOutbox WHERE [Status]='3') AS Outbox_Cancelled,
    (SELECT COUNT(1) FROM dbo.SCM_MailInbox WHERE ParseStatus='0' AND ProcessStatus='1') AS Inbox_Done,
    (SELECT COUNT(1) FROM dbo.SCM_MailInbox WHERE ParseStatus='3') AS Inbox_NotTarget,
    (SELECT COUNT(1) FROM dbo.SCM_MailInbox WHERE ProcessStatus='3') AS Inbox_Duplicated;
GO

PRINT N'9) 針對指定單據查關聯（請自行替換 DOCNumberNO）';
DECLARE @DocNumberNO NVARCHAR(20) = N'';
SELECT
    nl.LinkId,
    nl.DocNumberNO,
    nl.NeedId_A,
    nl.SimulationId_A,
    nl.NeedId_B,
    nl.LinkStatus,
    ss.PromiseDate,
    ss.PlanState,
    ss.ReportProgress,
    ss.ExceptionCode,
    ss.SnapshotTime
FROM dbo.SCM_NeedLink nl
LEFT JOIN dbo.SCM_StatusSnapshot ss ON ss.LinkId = nl.LinkId
WHERE (@DocNumberNO = N'' OR nl.DocNumberNO = @DocNumberNO)
ORDER BY nl.UpdatedTime DESC, ss.SnapshotTime DESC;
GO

