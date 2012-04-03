<%

' File: upload.asp
' 
' AXE(ASP Xtreme Evolution) upload utility.
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



' Class: Upload
' 
' This class implements an easy interface to upload files to a webserver. It's
' built using standard only functionallity in ASP. It will automaticly start
' parsing the binary stream sent by the browser which might take a while if the
' uploaded files are large.
' 
' Notes:
' 
'   Large parts of this code is based on code by Gez Lemon of Juicy Studio
'   <http://www.juicystudio.com/>. It has been and partially rewritten and
'   enhanced by CrazyBeaver Software along with it's new class interface.
' 
' About:
' 
'   - Written by Karl-Johan Sjögren <http://www.crazybeavers.se/> @ June 2006
'   - Modified and normalized by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ January 2008
' 
class Upload

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

    ' Property: Form
    ' 
    ' Stores Request.Form elements.
    ' 
    ' Contains:
    ' 
    '   (scripting.dictionary) - With the elements
    ' 
    public Form

    ' Property: Files
    ' 
    ' Stores File elements.
    ' 
    ' Returns:
    ' 
    '   (scripting.dictionary) - With the Upload_File objects
    ' 
    public Files

    ' Property: errorText
    ' 
    ' If an error occurs it's message will be stored here.
    ' 
    ' Returns:
    ' 
    '   (string) - Error description
    ' 
    public errorText

    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "1.1.0"
        
        dim contentType, sData, iEndPos, sBoundary, aData
        contentType = Request.serverVariables("HTTP_CONTENT_TYPE")
        if( inStr(contentType, "multipart/form-data") > 0 ) then
            sData = parseRequest()
            iEndPos = inStrRev(contentType, "=")
            sBoundary = trim( right(contentType, len(contentType) - iEndPos ) )
            aData = split(sData, sBoundary)
            set Form = Server.createObject("Scripting.Dictionary")
            set Files = Server.createObject("Scripting.Dictionary")
            call parseFormData(aData)
        else
            errorText = "Encoding: multipart/form-data is required."
        end if
    end sub

    private sub Class_terminate()
        set Files = nothing
        set Form = nothing
    end sub

    private function parseRequest()
        const adLongVarChar = 201
        dim Rs, iBytesRead, iBlockSize, iTotalBytes
        
        iTotalBytes = Request.totalBytes
        iBlockSize = 16384
        iBytesRead = 0
        
        set Rs = Server.createObject("ADODB.Recordset")
        Rs.Fields.append "BinaryField", adLongVarChar, iTotalBytes
        Rs.open
        Rs.addNew
        do while(iBytesRead < iTotalBytes)
            if(iBytesRead + iBlockSize < iTotalBytes) then
                Rs("BinaryField").appendChunk Request.binaryRead(iBlockSize)
                iBytesRead = iBytesRead + iBlockSize
            else
                Rs("BinaryField").appendChunk Request.binaryRead(iTotalBytes - iBytesRead)
                iBytesRead = iTotalBytes
            end if
        loop
        
        Rs.update
        parseRequest = Rs.Fields("BinaryField").value
        Rs.close
        set Rs = nothing
    end function

    private sub parseFormData(byRef aData)
        dim iCounter, iEndMarker
        dim sFieldInfo, sFieldValue
        dim File
        for iCounter = 0 to UBound(aData)
            iEndMarker = instr(aData(iCounter), vbNewLine & vbNewLine)
            if iEndMarker > 0 then
                sFieldInfo = mid(aData(iCounter), 3, iEndMarker - 3)
                sFieldValue = mid(aData(iCounter), iEndMarker + 4, len(aData(iCounter)) - iEndMarker - 7)
                if(instr(sFieldInfo, "filename=") > 0) then
                    if(len(getFileName(sFieldInfo)) > 0) then
                        set File = new Upload_File
                        File.Name = getFileName(sFieldInfo)
                        File.FormName = getFieldName(sFieldInfo)
                        File.Data = sFieldValue
                        File.ContentType = getContentType(sFieldInfo)
                        call Files.add(getFieldName(sFieldInfo), File)
                        set File = nothing
                    end if
                else
                    call Form.Add(getFieldName(sFieldInfo), sFieldValue)
                end if
            end if
        next
    end sub

    private function getFileName(byVal sData)
        dim sBrowser, iStartPos, iEndPos, sQuote
        if(sData = "") then
            getFieldName = ""
            exit function
        end if
        sQuote = Chr(34)
        iStartPos = instr(sData, "filename=")
        iEndPos = instr(iStartPos + 10, sData, sQuote & ";")
        if iEndPos = 0 then
            iEndPos = instr(iStartPos + 10, sData, sQuote)
        end if
        sData = mid(sData, iStartPos + 10, iEndPos - (iStartPos + 10))
        sBrowser = UCase(Request.ServerVariables("HTTP_USER_AGENT"))
        if(instr(sBrowser, "WIN") > 0) then
            iStartPos = instrrev(sData, "\")
            sData = mid(sData, iStartPos + 1)
        else
            iStartPos = instrrev(sData, "/")
            sData = mid(sData, iStartPos + 1)
        end if
        getFileName = sData
    end function

    private function getFieldName(sData)
        dim iStartPos, iEndPos, sQuote
        if(sData = "") then
            getFieldName = ""
            exit function
        end if
        sQuote = chr(34)
        iStartPos = instr(sData, "name=")
        iEndPos = instr(iStartPos + 6, sData, sQuote & ";")
        if iEndPos = 0 then
            iEndPos = instr(iStartPos + 6, sData, sQuote)
        end if
        getFieldName = mid(sData, iStartPos + 6, iEndPos - (iStartPos + 6))
    end function

    private function getContentType(byVal sData)
        dim iStartPos, iEndPos
        if(sData = "") then
            getContentType = ""
            exit function
        end if
        iStartPos = instr(sData, "Content-Type: ")
        iEndPos = len(sData)
        getContentType = mid(sData, iStartPos + 14, iEndPos)
    end function

end class



' Class: Upload_File
' 
' Each entry of Upload.Files contains an object of this class. This is the class
' that user should use to retrieve information from the uploaded file.
'
' About:
' 
'   - Written by Karl-Johan Sjögren <http://www.crazybeavers.se/> @ June 2006
'   - Modified and normalized by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ January 2008
' 
class Upload_File

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

    ' Property: name
    ' 
    ' File name.
    ' 
    ' Returns:
    ' 
    '   (string) - Name
    ' 
    public name

    ' Property: data
    ' 
    ' File.
    ' 
    ' Returns:
    ' 
    '   (binary) - File
    ' 
    public data

    ' Property: contentType
    ' 
    ' File mimetype.
    ' 
    ' Returns:
    ' 
    '   (string) - MIMEType
    ' 
    public contentType

    ' Property: formName
    ' 
    ' Name of the input field in the form which the image came from.
    ' 
    ' Returns:
    ' 
    '   (string) - Name
    ' 
    public formName

    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "1.1.0"
    end sub

    ' Function: size
    ' 
    ' Compute the file size.
    ' 
    ' Returns:
    ' 
    '   (int) - Size
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim oUpload, oFile
    ' set oUpload = new Upload
    ' for each oFile in oUpload.Files
    '     set oFile = oUpload.Files(oFile)
    '     Response.write(oFile.name & " has " & oFile.size & " bytes<br />" & vbNewLine)
    '     set oFile = nothing
    ' next
    ' set oUpload = nothing
    ' 
    ' (end)
    ' 
    public function size()
        size = len(data)
    end function

    ' Function: saveToFile
    ' 
    ' Saves the binary in the hard drive.
    ' 
    ' Parameters:
    ' 
    '   (string) - Physical path
    ' 
    ' Returns:
    ' 
    '   true  - if saveToFile is successful
    '   false - otherwise
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim oUpload, oFile
    ' set oUpload = new Upload
    ' for each oFile in oUpload.Files
    '     set oFile = oUpload.Files(oFile)
    '     if(oFile.saveToFile(Server.mappath("saved/" & oFile.name))) then
    '         Response.write(oFile.name & " has been saved...<br />" & vbNewLine)
    '     end if
    '     set oFile = nothing
    ' next
    ' set oUpload = nothing
    ' 
    ' (end)
    ' 
    public function saveToFile(sPath)
        dim Fso, TextFile
        set Fso = Server.createObject("Scripting.FileSystemObject")
        
        on error resume next
        set TextFile = Fso.createTextFile(sPath, true)
        TextFile.write(data)
        TextFile.close
        if(Err <> 0) then
            saveToFile = false
        else
            saveToFile = true
        end if
        on error goto 0
        
        set TextFile = nothing
        set Fso = nothing
    end function

end class

%>
