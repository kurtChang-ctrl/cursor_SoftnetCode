-- ============================================================
-- RAG_VectorStore 結構調整：Payload 分表 + 90 天保留 + 歷史表
-- 執行前請備份 SoftNetLogDB。請在 SoftNetLogDB 上執行。
-- ============================================================

USE SoftNetLogDB;
GO

-- ---------------------------------------------------------------------------
-- 1. 建立 RAG_VectorStorePayload（存放 Text, Metadata, Embedding）
-- ---------------------------------------------------------------------------
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RAG_VectorStorePayload')
BEGIN
    CREATE TABLE [dbo].[RAG_VectorStorePayload] (
        [Id]           NVARCHAR(50) NOT NULL,
        [Text]         NVARCHAR(MAX) NULL,
        [Metadata]     NVARCHAR(MAX) NULL,
        [Embedding]    NVARCHAR(MAX) NULL,
        CONSTRAINT [PK_RAG_VectorStorePayload] PRIMARY KEY CLUSTERED ([Id] ASC),
        CONSTRAINT [FK_RAG_VectorStorePayload_VectorStore] FOREIGN KEY ([Id])
            REFERENCES [dbo].[RAG_VectorStore]([Id]) ON DELETE CASCADE
    );
    PRINT 'Created table RAG_VectorStorePayload.';
END
ELSE
    PRINT 'Table RAG_VectorStorePayload already exists.';
GO

-- ---------------------------------------------------------------------------
-- 2. 建立歷史表（與主表+Payload 結構對應，供 90 天前資料歸檔）
-- ---------------------------------------------------------------------------
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RAG_VectorStore_History')
BEGIN
    CREATE TABLE [dbo].[RAG_VectorStore_History] (
        [Id]        NVARCHAR(50) NOT NULL,
        [ServerId]  NCHAR(10) NULL,
        [DataType]  NVARCHAR(50) NULL,
        [DataId]    NVARCHAR(50) NULL,
        [CreateTime] DATETIME NULL,
        [UpdateTime] DATETIME NULL,
        CONSTRAINT [PK_RAG_VectorStore_History] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
    PRINT 'Created table RAG_VectorStore_History.';
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RAG_VectorStorePayload_History')
BEGIN
    CREATE TABLE [dbo].[RAG_VectorStorePayload_History] (
        [Id]        NVARCHAR(50) NOT NULL,
        [Text]      NVARCHAR(MAX) NULL,
        [Metadata]  NVARCHAR(MAX) NULL,
        [Embedding] NVARCHAR(MAX) NULL,
        CONSTRAINT [PK_RAG_VectorStorePayload_History] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
    PRINT 'Created table RAG_VectorStorePayload_History.';
END
GO

-- ---------------------------------------------------------------------------
-- 3. 若主表仍有 Text/Metadata/Embedding，先複製到 Payload（僅補尚未存在的 Id）
-- ---------------------------------------------------------------------------
IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.RAG_VectorStore') AND name = 'Text')
BEGIN
    INSERT INTO [dbo].[RAG_VectorStorePayload] ([Id], [Text], [Metadata], [Embedding])
    SELECT v.[Id], v.[Text], v.[Metadata], v.[Embedding]
    FROM [dbo].[RAG_VectorStore] v
    WHERE NOT EXISTS (SELECT 1 FROM [dbo].[RAG_VectorStorePayload] p WHERE p.[Id] = v.[Id]);
    PRINT 'Copied existing RAG_VectorStore data into RAG_VectorStorePayload.';
END
GO

-- ---------------------------------------------------------------------------
-- 4. 從 RAG_VectorStore 移除 Text, Metadata, Embedding 欄位（若存在）
--    保留 Id, ServerId, DataType, DataId, CreateTime, UpdateTime；若有 EmbeddingVector 則保留供未來使用
-- ---------------------------------------------------------------------------
-- 4-0. 先移除所有依賴 Text/Metadata/Embedding 的索引（避免 DROP COLUMN 失敗）
IF EXISTS (SELECT 1 FROM sys.tables WHERE object_id = OBJECT_ID(N'dbo.RAG_VectorStore'))
BEGIN
    DECLARE @idxName SYSNAME;
    DECLARE idx_cursor CURSOR FAST_FORWARD FOR
    SELECT DISTINCT i.name
    FROM sys.indexes i
    INNER JOIN sys.index_columns ic ON ic.object_id = i.object_id AND ic.index_id = i.index_id
    INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
    WHERE i.object_id = OBJECT_ID(N'dbo.RAG_VectorStore')
      AND i.is_primary_key = 0
      AND i.is_unique_constraint = 0
      AND c.name IN (N'Text', N'Metadata', N'Embedding');

    OPEN idx_cursor;
    FETCH NEXT FROM idx_cursor INTO @idxName;
    WHILE @@FETCH_STATUS = 0
    BEGIN
        BEGIN TRY
            DECLARE @dropSql NVARCHAR(MAX) = N'DROP INDEX ' + QUOTENAME(@idxName) + N' ON [dbo].[RAG_VectorStore];';
            EXEC sp_executesql @dropSql;
            PRINT N'Dropped dependent index on RAG_VectorStore: ' + @idxName;
        END TRY
        BEGIN CATCH
            PRINT N'WARN: Failed to drop index ' + ISNULL(@idxName, N'(null)') + N' : ' + ERROR_MESSAGE();
        END CATCH
        FETCH NEXT FROM idx_cursor INTO @idxName;
    END
    CLOSE idx_cursor;
    DEALLOCATE idx_cursor;
END
GO

IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.RAG_VectorStore') AND name = 'Text')
BEGIN
    ALTER TABLE [dbo].[RAG_VectorStore] DROP COLUMN [Text];
    PRINT 'Dropped column RAG_VectorStore.Text';
END
IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.RAG_VectorStore') AND name = 'Metadata')
BEGIN
    ALTER TABLE [dbo].[RAG_VectorStore] DROP COLUMN [Metadata];
    PRINT 'Dropped column RAG_VectorStore.Metadata';
END
IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.RAG_VectorStore') AND name = 'Embedding')
BEGIN
    ALTER TABLE [dbo].[RAG_VectorStore] DROP COLUMN [Embedding];
    PRINT 'Dropped column RAG_VectorStore.Embedding';
END
GO

-- ---------------------------------------------------------------------------
-- 5. 索引：查詢與 90 天過濾用
-- ---------------------------------------------------------------------------
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.RAG_VectorStore') AND name = 'IX_RAG_VectorStore_ServerId_DataType_CreateTime')
BEGIN
    CREATE NONCLUSTERED INDEX [IX_RAG_VectorStore_ServerId_DataType_CreateTime]
        ON [dbo].[RAG_VectorStore]([ServerId], [DataType], [CreateTime] DESC);
    PRINT 'Created index IX_RAG_VectorStore_ServerId_DataType_CreateTime.';
END
GO

-- ---------------------------------------------------------------------------
-- 6. 預存程序：將超過保留天數的資料搬至歷史表並自主表刪除（預設 90 天）
-- ---------------------------------------------------------------------------
CREATE OR ALTER PROCEDURE [dbo].[sp_RAG_VectorStore_PurgeOldData]
    @RetentionDays INT = 90,
    @ServerId      NCHAR(10) = NULL  -- NULL = 全部 ServerId
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @Cutoff DATE = DATEADD(DAY, -@RetentionDays, GETDATE());

    -- 先搬 Payload 到歷史（依主表將被搬移的 Id）
    INSERT INTO [dbo].[RAG_VectorStorePayload_History] ([Id], [Text], [Metadata], [Embedding])
    SELECT p.[Id], p.[Text], p.[Metadata], p.[Embedding]
    FROM [dbo].[RAG_VectorStorePayload] p
    INNER JOIN [dbo].[RAG_VectorStore] v ON v.[Id] = p.[Id]
    WHERE v.[CreateTime] < @Cutoff
      AND (@ServerId IS NULL OR v.[ServerId] = @ServerId);

    -- 再搬主表到歷史
    INSERT INTO [dbo].[RAG_VectorStore_History] ([Id], [ServerId], [DataType], [DataId], [CreateTime], [UpdateTime])
    SELECT [Id], [ServerId], [DataType], [DataId], [CreateTime], [UpdateTime]
    FROM [dbo].[RAG_VectorStore]
    WHERE [CreateTime] < @Cutoff
      AND (@ServerId IS NULL OR [ServerId] = @ServerId);

    -- 刪除主表（CASCADE 會一併刪除 Payload）
    DELETE FROM [dbo].[RAG_VectorStore]
    WHERE [CreateTime] < @Cutoff
      AND (@ServerId IS NULL OR [ServerId] = @ServerId);

    SELECT @@ROWCOUNT AS RowsMovedToHistory;
END
GO

PRINT 'RAG_VectorStore migration (Payload + History + Purge SP) completed.';
GO
