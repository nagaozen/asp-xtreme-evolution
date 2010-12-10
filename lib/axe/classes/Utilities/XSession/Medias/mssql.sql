/* 
 * File: mssql.sql
 * 
 * AXE(ASP Xtreme Evolution) XSession_Media_MSSQL structure definition.
 * 
 * License:
 * 
 * The MIT License
 * Copyright (C) 2010 Fabio Zendhi Nagao
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */



/* Table: axe_xsession
 * 
 * Holds all the XSession data.
 * 
 * Primary key:
 * 
 *  (int) - PK_axe_xsession_key
 * 
 * Indexes:
 * 
 *  (clustered) - PK_axe_xsession_key
 * 
 */
CREATE TABLE [$owner].[axe_xsession] (
    [key] uniqueidentifier NOT NULL CONSTRAINT PK_axe_xsession_key PRIMARY KEY CLUSTERED,
    [content] ntext COLLATE SQL_Latin1_General_CP1_CI_AI NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO



/* Stored Procedure: SP_axe_xsession_create
 * 
 * Creates a new XSession record.
 * 
 * Parameters:
 * 
 *  (uniqueidentifier) - XSession id
 * 
 */
CREATE PROC [$owner].[SP_axe_xsession_create]
    @uid uniqueidentifier = NULL,
    @content ntext = NULL
AS
    IF @uid is NULL BEGIN
        PRINT 'This stored procedure syntax is: EXEC [$owner].[SP_axe_xsession_create] (uniqueidentifier)@uid, (ntext)@content'
        RETURN
    END
    
    INSERT INTO [$owner].[axe_xsession]([key], content)
    VALUES (@uid, @content)
GO

/* Stored Procedure: SP_axe_xsession_retrieve
 * 
 * Retrieves a XSession record.
 * 
 * Parameters:
 * 
 *  (uniqueidentifier) - XSession id
 * 
 */
CREATE PROC [$owner].[SP_axe_xsession_retrieve]
    @uid uniqueidentifier = NULL
AS
    IF @uid is NULL BEGIN
        PRINT 'This stored procedure syntax is: EXEC [$owner].[SP_axe_xsession_retrieve] (uniqueidentifier)@uid'
        RETURN
    END
    
    SELECT content
    FROM [$owner].[axe_xsession]
    WHERE [key] = @uid
GO

/* Stored Procedure: SP_axe_xsession_update
 * 
 * Updates a XSession record.
 * 
 * Parameters:
 * 
 *  (uniqueidentifier) - XSession id 
 *  (ntext) - XSession content
 * 
 */
CREATE PROC [$owner].[SP_axe_xsession_update]
    @uid uniqueidentifier = NULL,
    @content ntext = NULL
AS
    IF @uid is NULL BEGIN
        PRINT 'This stored procedure syntax is: EXEC [$owner].[SP_axe_xsession_update] (uniqueidentifier)@uid, (ntext)@content'
        RETURN
    END
    
    UPDATE [$owner].[axe_xsession]
    SET content = @content
    WHERE [key] = @uid
GO

/* Stored Procedure: SP_axe_xsession_delete
 * 
 * Deletes a XSession record.
 * 
 * Parameters:
 * 
 *  (uniqueidentifier) - XSession id 
 * 
 */
CREATE PROC [$owner].[SP_axe_xsession_delete]
    @uid uniqueidentifier = NULL
AS
    IF @uid is NULL BEGIN
        PRINT 'This stored procedure syntax is: EXEC [$owner].[SP_axe_xsession_delete] (uniqueidentifier)@uid'
        RETURN
    END
    
    DELETE FROM [$owner].[axe_xsession]
    WHERE [key] = @uid
GO



/* Stored Procedure: SP_axe_xsession_load
 * 
 * description
 * 
 * Parameters:
 * 
 *  (uniqueidentifier) - XSession id 
 * 
 */
CREATE PROC [$owner].[SP_axe_xsession_load]
    @uid uniqueidentifier = NULL
AS
    IF @uid is NULL BEGIN
        PRINT 'This stored procedure syntax is: EXEC [$owner].[SP_axe_xsession_load] (uniqueidentifier)@uid'
        RETURN
    END
    
    EXEC [$owner].[SP_axe_xsession_retrieve] @uid
GO

/* Stored Procedure: SP_axe_xsession_save
 * 
 * description
 * 
 * Parameters:
 * 
 *  (uniqueidentifier) - XSession id 
 * 
 */
CREATE PROC [$owner].[SP_axe_xsession_save]
    @uid uniqueidentifier = NULL,
    @content ntext = NULL
AS
    IF @uid is NULL BEGIN
        PRINT 'This stored procedure syntax is: EXEC [$owner].[SP_axe_xsession_save] (uniqueidentifier)@uid, (ntext)@content'
        RETURN
    END
    
    EXEC [$owner].[SP_axe_xsession_retrieve] @uid
    IF @@ROWCOUNT > 0
        BEGIN
            EXEC [$owner].[SP_axe_xsession_update] @uid, @content
        END
    ELSE
        BEGIN
            EXEC [$owner].[SP_axe_xsession_create] @uid, @content
        END
GO
