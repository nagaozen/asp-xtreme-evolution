<%

' File: csv.asp
' 
' AXE(ASP Xtreme Evolution) implementation of CSV parser.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2007-2012 Fabio Zendhi Nagao
' 
' ASP Xtreme Evolution is free software: you can redistribute it and/or modify
' it under the terms of the GNU Lesser General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
' 
' ASP Xtreme Evolution is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU Lesser General Public License for more details.
' 
' You should have received a copy of the GNU Lesser General Public License
' along with ASP Xtreme Evolution. If not, see <http://www.gnu.org/licenses/>.

' Class: CSV
' 
' Comma-Separated Values (CSV) is an ancient data interchange format. Its widely
' used in the wild, therefore a must class for any serious framework. This
' implementation leverages Microsoft Jet OLEDB, so it requires a  writable
' folder to create schema files in order to help the CSV parsing. The best side
' of using MS Text Driver is that it's compatible with RFC4180, Excel and other
' Windows made CSV files.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Dec 2011
' 
' References:
' 
'     - Much ADO About Text Files <http://msdn.microsoft.com/en-us/library/ms974559.aspx> @ MSDN Library
'     - Common Format and MIME Type for Comma-Separated Values (CSV) Files <http://www.ietf.org/rfc/rfc4180.txt> @ IETF
' 
class CSV

    ' Property: classType
    ' 
    ' Class type.
    ' 
    ' Contains:
    ' 
    '   (string) - type
    ' 
    public classType

    ' Property: classVersion
    ' 
    ' Class version.
    ' 
    ' Contains:
    ' 
    '   (float) - version
    ' 
    public classVersion

    ' Property: writablesPath
    ' 
    ' Application writable physical path.
    ' 
    ' Contains:
    ' 
    '   (string) - application writable physical path.
    ' 
    public writablesPath

    ' Property: codepage
    ' 
    ' Microsoft character encoding numberic identification (codepage)
    ' 
    ' Contains:
    ' 
    '   (integer) - numeric character encoding identification. Defaults to 65001.
    ' 
    public codepage

    ' Property: charset
    ' 
    ' Microsoft character encoding human friendly identification (charset)
    ' 
    ' Contains:
    ' 
    '   (string) - human friendly character encoding identification. Defaults to UTF-8.
    ' 
    public charset

    ' Property: separator
    ' 
    ' CSV separator string. By the RFC4810 it should be ',' but some programs use <TAB> or ';'.
    ' 
    ' Contains:
    ' 
    '   (string) - CSV separator. Defaults to ','.
    ' 
    public separator

    ' Function: fromString
    ' 
    ' Reads a CSV string and convert it into an ADODB.Recordset.
    ' 
    ' Parameters:
    ' 
    '     (string)  - CSV string
    '     (boolean) - if the CSV string has headers or not
    ' 
    ' Returns:
    ' 
    '     (ADODB.Recordset) - In memory CSV model
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' set Parser = new CSV
    ' Parser.writablesPath = Server.mapPath("/app/writables/")
    ' Parser.charset   = "UTF-8"
    ' Parser.codepage  = 65001
    ' Parser.separator = ";"
    ' 
    ' set Data = Parser.fromString( join(array( _
    '     "id;firstname", _
    '     "1;Foo", _
    '     "2;Bar" _
    ' ), vbCrLf), true )
    ' Response.write( dump(Data) )
    ' Data.close()
    ' set Data = nothing
    ' 
    ' set Parser = nothing
    ' 
    ' (end code)
    ' 
    public function fromString(str, hasHeaders) : call [_ε]
        if( not isEmpty([_DataSourceFolder]) ) then
            call [_deleteFile]([_DataSourceFolder] & "\Schema.ini")
            call [_deleteFile]([_DataSourceFolder] & "\DataSource.csv")
            call [_deleteFolder]([_DataSourceFolder])
        end if
        
        [_DataSourceFolder] = writablesPath & "\" & guid()
        
        call [_createFolder]([_DataSourceFolder])
        call [_createFile]([_DataSourceFolder] & "\DataSource.csv" , str)
        call [_createSchema](hasHeaders)
        
        [_Conn].open( strsubstitute( _
            "Provider=Microsoft.Jet.OLEDB.4.0; Data Source={0};Extended Properties=""text;CharacterSet={1};"";", _
            array([_DataSourceFolder], codepage) _
        ) )
        
        set fromString = [_Conn].execute("SELECT * FROM [DataSource.csv]")
    end function

    ' Function: fromFile
    ' 
    ' Reads a CSV file and convert it into an ADODB.Recordset.
    ' 
    ' Parameters:
    ' 
    '     (string)  - file system path to the CSV file
    '     (boolean) - if the CSV file use headers or not
    ' 
    ' Returns:
    ' 
    '     (ADODB.Recordset) - In memory CSV model
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' set Parser = new CSV
    ' Parser.writablesPath = Server.mapPath("/app/writables/")
    ' Parser.charset   = "Windows-1252"
    ' Parser.codepage  = 1252
    ' Parser.separator = ","
    ' 
    ' set Data = Parser.fromFile( Server.mapPath("file.csv"), true )
    ' while(not Data.eof)
    '     Response.write( Data("firstname").value )
    '     Data.moveNext()
    ' wend
    ' Data.close()
    ' set Data = nothing
    ' 
    ' set Parser = nothing
    ' 
    ' (end code)
    ' 
    public function fromFile(path, hasHeaders)
        dim csv : csv = [_loadTextFile](path)
        set fromFile = fromString(csv, hasHeaders)
    end function

    ' Function: toString
    ' 
    ' Returns a RFC4810 compatible CSV with headers string.
    ' 
    ' Parameters:
    ' 
    '     (ADODB.Recordset) - CSV data source
    ' 
    ' Returns:
    ' 
    '     (string) - CSV
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Conn, Rs, Parser, csv
    ' set Conn = Server.createObject("ADODB.Connection")
    ' Conn.open(connectionstring)
    ' 
    ' set Rs = Conn.execute(...)
    ' if(not Rs.eof) then
    '     set Parser = new CSV
    '     Parser.writablesPath = Server.mapPath("/app/writables/")
    '     csv = Parser.toString(Rs)
    '     set Parser = nothing
    ' end if
    ' set Rs = nothing
    ' 
    ' Conn.close()
    ' set Conn = nothing
    ' 
    ' Response.write(csv)
    ' 
    ' (end code)
    ' 
    public function toString(Rs)
        dim line, csv, i
        
        set csv = new StringBuilder
        set line = new StringBuilder
        
        if(not Rs.eof) then
            ' build header
            for i = 0 to ( Rs.Fields.count - 1 )
                line.append separator
                line.append """"
                line.append replace(Rs(i).name, """", """""")
                line.append """"
            next
            line.append vbCrLf
            csv.append mid(line.toString(), len(separator) + 1)
            
            ' build body
            while(not Rs.eof)
                line.reset()
                
                for i = 0 to ( Rs.Fields.count - 1 )
                    line.append separator
                    line.append """"
                    line.append replace(Rs(i).value, """", """""")
                    line.append """"
                next
                line.append vbCrLf
                csv.append mid(line.toString(), len(separator) + 1)
                
                Rs.moveNext()
            wend
        end if
        
        toString = csv.toString()
        
        set line = nothing
        set csv = nothing
    end function

    ' --[Private Section]-------------------------------------------------------

    ' Property: SCHEMA_INI_TEMPLATE
    ' 
    ' {private} Schema.ini template
    ' 
    ' Contains:
    ' 
    '     (string) - Schema.ini template
    ' 
    private SCHEMA_INI_TEMPLATE

    ' Property: [_Conn]
    ' 
    ' {private} Because this parser leverages an ADODB.Recordset that requires
    ' an open ADODB.Connection object, a shared class scoped connection is required
    ' to let users navigate in the CSV.
    ' 
    ' Contains:
    ' 
    '     (ADODB.Connection) - Parser connection
    ' 
    private [_Conn]

    ' Property: [_DataSourceFolder]
    ' 
    ' {private} Full physical path to the folder containing the DataSource.csv and the Schema.ini files.
    ' 
    ' Contains:
    ' 
    '     (string) - Data source physical path
    ' 
    private [_DataSourceFolder]

    ' Subroutine: [_ε]
    ' 
    ' {private} Checks for a writable assignment.
    ' 
    private sub [_ε]
        if( isEmpty(writablesPath) ) then
            Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Missing a writable folder."
        end if
    end sub

    ' Subroutine: [_createSchema]
    ' 
    ' {private} Creates an ANSI encoded Schema.ini file in the same folder as the DataSource.csv file.
    ' 
    ' Parameters:
    ' 
    '   (boolean) - if the CSV file contains header information in the first line or not.
    ' 
    private sub [_createSchema](hasHeaders)
        dim Fso : set Fso = Server.createObject("Scripting.FileSystemObject")
        dim File : set File = Fso.createTextFile([_DataSourceFolder] & "\Schema.ini", true, false)
        File.write( strsubstitute( SCHEMA_INI_TEMPLATE, array(separator, hasHeaders) ) )
        File.close()
        set File = nothing
        set Fso = nothing
    end sub

    ' Function: [_fileExists]
    ' 
    ' {private} Checks if the file exists in the file system.
    ' 
    ' Parameters:
    ' 
    '   (string) - Full path plus file name with extension
    ' 
    ' Returns:
    ' 
    '   true  - if it's there
    '   false - otherwise
    ' 
    private function [_fileExists](sFilePath)
        dim Fso : set Fso = Server.createObject("Scripting.FileSystemObject")
        [_fileExists] = Fso.fileExists(sFilePath)
        set Fso = nothing
    end function

    ' Function: [_loadTextFile]
    ' 
    ' {private} If the file exists, read it all and return the content.
    ' 
    ' Parameters:
    ' 
    '   (string) - Full path plus file name with extension
    ' 
    ' Returns:
    ' 
    '   (string) - The file content
    ' 
    private function [_loadTextFile](sFilePath)
        if([_fileExists](sFilePath)) then
            dim Stream : set Stream = Server.createObject("ADODB.Stream")
            with Stream
                .type = adTypeText
                .mode = adModeReadWrite
                .charset = charset
                .open()
                
                .loadFromFile(sFilePath)
                .position = 0
                [_loadTextFile] = .readText()
                
                .close()
            end with
            set Stream = nothing
        else
            Err.raise 53, "Evolved AXE runtime error"
        end if
    end function

    ' Subroutine: [_createFile]
    ' 
    ' {private} Create/Overwrite a file in the file system.
    ' 
    ' Parameters:
    ' 
    '   (string) - Full path plus file name with extension
    '   (string) - File content
    ' 
    private sub [_createFile](sFilePath, sContent)
        select case charset
            case "UTF-8":
                call [_createFileWithoutBOM](sFilePath, sContent, 3)
            
            case "UTF-16":
                call [_createFileWithoutBOM](sFilePath, sContent, 2)
            
            case "UTF-32":
                call [_createFileWithoutBOM](sFilePath, sContent, 4)
            
            case else:
                call [_createFileStd](sFilePath, sContent)
        end select
    end sub

    ' Subroutine: [_createFileStd]
    ' 
    ' {private} Create/Overwrite a file in the file system without any special treatment.
    ' 
    ' Parameters:
    ' 
    '   (string) - Full path plus file name with extension
    '   (string) - File content
    ' 
    private sub [_createFileStd](sFilePath, sContent)
        dim Stream : set Stream = Server.createObject("ADODB.Stream")
        with Stream
            .type = adTypeText
            .mode = adModeReadWrite
            .charset = charset
            .open()
            
            .writeText(sContent)
            .setEOS()
            
            .position = 0
            .saveToFile sFilePath, adSaveCreateOverwrite
            
            .close()
        end with
        set Stream = nothing
    end sub

    ' Subroutine: [_createFileWithoutBOM]
    ' 
    ' {private} Create/Overwrite a file in the file system removing the Byte Order Mark. CSV files doesn't support them.
    ' 
    ' Parameters:
    ' 
    '   (string)  - Full path plus file name with extension
    '   (string)  - File content
    '   (integer) - How many bytes to jump in order to skip the BOM
    ' 
    private sub [_createFileWithoutBOM](sFilePath, sContent, p0)
        dim Stream, bin
        
        set Stream = Server.createObject("ADODB.Stream")
        with Stream
            .type = adTypeText
            .mode = adModeReadWrite
            .charset = charset
            .open()
            
            .writeText(sContent)
            .setEOS()
            
            .position = 0
            .type = adTypeBinary
            .position = p0' skips the Byte Order Mark (BOM)
            bin = .read()
            
            .close()
        end with
        set Stream = nothing
        
        set Stream = Server.createObject("ADODB.Stream")
        with Stream
            .type = adTypeBinary
            .open()
            
            .write(bin)
            .setEOS()
            
            .position = 0
            .saveToFile sFilePath, adSaveCreateOverwrite
            
            .close()
        end with
        set Stream = nothing
    end sub

    ' Subroutine: [_deleteFile]
    ' 
    ' {private} Delete a file in the file system.
    ' 
    ' Parameters:
    ' 
    '   (string) - Full path plus file name with extension
    ' 
    private sub [_deleteFile](sFilePath)
        dim Fso : set Fso = Server.createObject("Scripting.FileSystemObject")
        if(Fso.fileExists(sFilePath)) then
            call Fso.deleteFile(sFilePath, true)
        end if
        set Fso = nothing
    end sub

    ' Subroutine: [_createFolder]
    ' 
    ' {private} Create a folder in the file system.
    ' 
    ' Parameters:
    ' 
    '   (string) - Folder path without the last "\"
    ' 
    private sub [_createFolder](sFolderPath)
        dim Fso : set Fso = Server.createObject("Scripting.FileSystemObject")
        if( not Fso.folderExists(sFolderPath) ) then
            Fso.createFolder(sFolderPath)
        end if
        set Fso = nothing
    end sub

    ' Subroutine: [_deleteFolder]
    ' 
    ' {private} Delete a folder in the file system.
    ' 
    ' Parameters:
    ' 
    '   (string) - Folder path without the last "\"
    ' 
    private sub [_deleteFolder](sFolderPath)
        dim Fso : set Fso = Server.createObject("Scripting.FileSystemObject")
        if(Fso.folderExists(sFolderPath)) then
            call Fso.deleteFolder(sFolderPath, true)
        end if
        set Fso = nothing
    end sub

    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0"
        
        codepage   = 65001
        charset    = "UTF-8"
        separator  = ","
        
        SCHEMA_INI_TEMPLATE = join(array( _
            "[DataSource.csv]", _
            "Format = Delimited({0})", _
            "ColNameHeader = {1}" _
        ), vbNewline)
        
        set [_Conn] = Server.createObject("ADODB.Connection")
    end sub

    private sub Class_terminate()
        if( [_Conn].state = adStateOpen ) then [_Conn].close()
        set [_Conn] = nothing
        
        call [_deleteFile]([_DataSourceFolder] & "\Schema.ini")
        call [_deleteFile]([_DataSourceFolder] & "\DataSource.csv")
        call [_deleteFolder]([_DataSourceFolder])
    end sub

end class

%>
