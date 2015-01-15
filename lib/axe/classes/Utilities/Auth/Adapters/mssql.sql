/* 
 * File: mssql.sql
 * 
 * AXE(ASP Xtreme Evolution) Auth_Adapter_MSSQL structure definition.
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



/* Table: axe_auth
 * 
 * Prototyped table used to show a basic MSSQL structure to use with AXE Auth class.
 * 
 * Primary key:
 * 
 *  (int) - PK_axe_auth_key
 * 
 * Indexes:
 * 
 *  (clustered) - PK_axe_auth_key
 *  (nonclustered) - UN_axe_auth_usr
 * 
 * Defaults:
 * 
 *  (int) - Defaults to auto number.
 * 
 */
CREATE TABLE [$owner].[axe_auth] (
    [key] int IDENTITY(1, 1) NOT NULL CONSTRAINT PK_axe_auth_key PRIMARY KEY CLUSTERED,
    [usr] nvarchar(128) COLLATE SQL_Latin1_General_CP1_CI_AI NOT NULL CONSTRAINT UN_axe_auth_usr UNIQUE NONCLUSTERED,
    [pwd] nvarchar(32) COLLATE SQL_Latin1_General_CP1_CI_AI NOT NULL,
    [active] bit NOT NULL CONSTRAINT DF_axe_auth_active DEFAULT 1
) ON [PRIMARY]
GO



/* Stored Procedure: SP_axe_auth_authenticate
 * 
 * This procedure tries to find the right user to test the given credential.
 * 
 * Parameters:
 * 
 *  (nvarchar) - User id
 * 
 */
CREATE PROC [$owner].[SP_axe_auth_authenticate]
    @usr nvarchar(128) = NULL
AS
    IF @usr is NULL BEGIN
        PRINT 'This stored procedure syntax is: EXEC [$owner].[SP_axe_auth_authenticate] (nvarchar)@usr'
        RETURN
    END
    
    SELECT pwd
    FROM [$owner].[axe_auth]
    WHERE active = 1 AND usr = @usr
GO
