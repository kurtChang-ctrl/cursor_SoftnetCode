USE [SoftNetLogDB];
GO

SET ANSI_NULLS ON;
GO
SET QUOTED_IDENTIFIER ON;
GO

/* =========================================================
   SCM_MailOutbox
   用途：寄件佇列表（待寄、重送、寄送結果）
   ========================================================= */
IF OBJECT_ID(N'dbo.SCM_MailOutbox', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SCM_MailOutbox
    (
        Id              BIGINT IDENTITY(1,1) NOT NULL,
        MessageId       NVARCHAR(50)         NOT NULL,
        FromServerId    NCHAR(2)             NOT NULL,
        ToServerId      NCHAR(2)             NOT NULL,
        MessageType     NVARCHAR(20)         NOT NULL,
        Subject         NVARCHAR(500)        NOT NULL,
        BodyJson        NVARCHAR(MAX)        NOT NULL,
        [Status]        CHAR(1)              NOT NULL CONSTRAINT DF_SCM_MailOutbox_Status DEFAULT ('0'),
        RetryCount      INT                  NOT NULL CONSTRAINT DF_SCM_MailOutbox_RetryCount DEFAULT (0),
        NextRetryTime   DATETIME             NULL,
        CreatedTime     DATETIME             NOT NULL CONSTRAINT DF_SCM_MailOutbox_CreatedTime DEFAULT (GETDATE()),
        SentTime        DATETIME             NULL,
        ErrorMessage    NVARCHAR(1000)       NULL,
        CONSTRAINT PK_SCM_MailOutbox PRIMARY KEY CLUSTERED (Id),
        CONSTRAINT UX_SCM_MailOutbox_MessageId UNIQUE NONCLUSTERED (MessageId)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_SCM_MailOutbox_Status_NextRetryTime' AND object_id = OBJECT_ID(N'dbo.SCM_MailOutbox'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_SCM_MailOutbox_Status_NextRetryTime
        ON dbo.SCM_MailOutbox([Status], NextRetryTime);
END
GO

/* =========================================================
   SCM_MailInbox
   用途：收件暫存表（收信、解析、處理）
   ========================================================= */
IF OBJECT_ID(N'dbo.SCM_MailInbox', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SCM_MailInbox
    (
        Id              BIGINT IDENTITY(1,1) NOT NULL,
        MessageId       NVARCHAR(50)         NOT NULL,
        FromServerId    NCHAR(2)             NOT NULL,
        ToServerId      NCHAR(2)             NOT NULL,
        MessageType     NVARCHAR(20)         NOT NULL,
        Subject         NVARCHAR(500)        NOT NULL,
        BodyJson        NVARCHAR(MAX)        NOT NULL,
        ReceiveTime     DATETIME             NOT NULL CONSTRAINT DF_SCM_MailInbox_ReceiveTime DEFAULT (GETDATE()),
        ParseStatus     CHAR(1)              NOT NULL CONSTRAINT DF_SCM_MailInbox_ParseStatus DEFAULT ('0'),
        ProcessStatus   CHAR(1)              NOT NULL CONSTRAINT DF_SCM_MailInbox_ProcessStatus DEFAULT ('0'),
        ErrorMessage    NVARCHAR(1000)       NULL,
        CONSTRAINT PK_SCM_MailInbox PRIMARY KEY CLUSTERED (Id)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_SCM_MailInbox_MessageId' AND object_id = OBJECT_ID(N'dbo.SCM_MailInbox'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_SCM_MailInbox_MessageId
        ON dbo.SCM_MailInbox(MessageId);
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_SCM_MailInbox_ToServerId_ReceiveTime' AND object_id = OBJECT_ID(N'dbo.SCM_MailInbox'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_SCM_MailInbox_ToServerId_ReceiveTime
        ON dbo.SCM_MailInbox(ToServerId, ReceiveTime DESC);
END
GO

/* =========================================================
   SCM_MailPollState
   用途：收件輪詢游標（IMAP UniqueId），已掃過的郵件（含略過的非 SCM 主旨）
         下輪以 UID 過濾不再進入處理迴圈，且不依賴標已讀。
   ========================================================= */
IF OBJECT_ID(N'dbo.SCM_MailPollState', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SCM_MailPollState
    (
        ServerId           NCHAR(2)             NOT NULL,
        LastImapUniqueId   BIGINT               NOT NULL CONSTRAINT DF_SCM_MailPollState_LastImapUniqueId DEFAULT (0),
        UpdatedTime        DATETIME             NOT NULL CONSTRAINT DF_SCM_MailPollState_UpdatedTime DEFAULT (GETDATE()),
        CONSTRAINT PK_SCM_MailPollState PRIMARY KEY CLUSTERED (ServerId)
    );
END
GO

/* =========================================================
   SCM_NeedLink
   用途：A/B 需求關聯主表
   ========================================================= */
IF OBJECT_ID(N'dbo.SCM_NeedLink', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SCM_NeedLink
    (
        LinkId          NCHAR(12)            NOT NULL,
        FromServerId    NCHAR(2)             NOT NULL,
        ToServerId      NCHAR(2)             NOT NULL,
        NeedId_A        NCHAR(12)            NOT NULL,
        SimulationId_A  NCHAR(12)            NULL,
        NeedId_B        NCHAR(12)            NOT NULL,
        DocType         NCHAR(4)             NOT NULL,
        DocNumberNO     NVARCHAR(20)         NOT NULL,
        LinkStatus      CHAR(1)              NOT NULL CONSTRAINT DF_SCM_NeedLink_LinkStatus DEFAULT ('0'),
        CreatedTime     DATETIME             NOT NULL CONSTRAINT DF_SCM_NeedLink_CreatedTime DEFAULT (GETDATE()),
        UpdatedTime     DATETIME             NOT NULL CONSTRAINT DF_SCM_NeedLink_UpdatedTime DEFAULT (GETDATE()),
        CONSTRAINT PK_SCM_NeedLink PRIMARY KEY CLUSTERED (LinkId)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_SCM_NeedLink_NeedIdA_SimulationIdA' AND object_id = OBJECT_ID(N'dbo.SCM_NeedLink'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_SCM_NeedLink_NeedIdA_SimulationIdA
        ON dbo.SCM_NeedLink(NeedId_A, SimulationId_A);
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_SCM_NeedLink_NeedIdB' AND object_id = OBJECT_ID(N'dbo.SCM_NeedLink'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_SCM_NeedLink_NeedIdB
        ON dbo.SCM_NeedLink(NeedId_B);
END
GO

/* =========================================================
   SCM_StatusSnapshot
   用途：供應鏈進度快照表
   ========================================================= */
IF OBJECT_ID(N'dbo.SCM_StatusSnapshot', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SCM_StatusSnapshot
    (
        SnapshotId      BIGINT IDENTITY(1,1) NOT NULL,
        LinkId          NCHAR(12)            NOT NULL,
        NeedId_B        NCHAR(12)            NOT NULL,
        PromiseDate     DATETIME             NULL,
        PlanState       NVARCHAR(50)         NOT NULL CONSTRAINT DF_SCM_StatusSnapshot_PlanState DEFAULT (N''),
        ReportProgress  NVARCHAR(200)        NOT NULL CONSTRAINT DF_SCM_StatusSnapshot_ReportProgress DEFAULT (N''),
        ExceptionCode   NVARCHAR(30)         NULL,
        ExceptionRemark NVARCHAR(500)        NULL,
        SnapshotTime    DATETIME             NOT NULL CONSTRAINT DF_SCM_StatusSnapshot_SnapshotTime DEFAULT (GETDATE()),
        CONSTRAINT PK_SCM_StatusSnapshot PRIMARY KEY CLUSTERED (SnapshotId),
        CONSTRAINT FK_SCM_StatusSnapshot_SCM_NeedLink FOREIGN KEY (LinkId) REFERENCES dbo.SCM_NeedLink(LinkId)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_SCM_StatusSnapshot_LinkId_SnapshotTime' AND object_id = OBJECT_ID(N'dbo.SCM_StatusSnapshot'))
BEGIN
    CREATE NONCLUSTERED INDEX IX_SCM_StatusSnapshot_LinkId_SnapshotTime
        ON dbo.SCM_StatusSnapshot(LinkId, SnapshotTime DESC);
END
GO

