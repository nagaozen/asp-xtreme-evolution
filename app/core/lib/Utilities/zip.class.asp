<%

'+-----------------------------------------------------------------------------+
'|This file is part of ASP Xtreme Evolution.                                   |
'|Copyright (C) 2007, 2009 Fabio Zendhi Nagao                                  |
'|                                                                             |
'|ASP Xtreme Evolution is free software: you can redistribute it and/or modify |
'|it under the terms of the GNU Lesser General Public License as published by  |
'|the Free Software Foundation, either version 3 of the License, or            |
'|(at your option) any later version.                                          |
'|                                                                             |
'|ASP Xtreme Evolution is distributed in the hope that it will be useful,      |
'|but WITHOUT ANY WARRANTY; without even the implied warranty of               |
'|MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                |
'|GNU Lesser General Public License for more details.                          |
'|                                                                             |
'|You should have received a copy of the GNU Lesser General Public License     |
'|along with ASP Xtreme Evolution.  If not, see <http://www.gnu.org/licenses/>.|
'+-----------------------------------------------------------------------------+

' Class: Zip
' 
' This class intends to help people use the CBZip in an easy way. CBZip brings
' the power of zip archives to asp.
' 
' Requirements:
' 
'   - CrazyBeavers Zip exe (CBZip.exe)
' 
' About:
' 
'   - Written by Karl-Johan Sjögren <http://www.crazybeavers.se/> @ February 2006
'   - Modified and normalized by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ January 2008
' 
class Zip

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

    ' Property: cbZipUri
    ' 
    ' This should be the URI to CBZip.exe including http:// etc.
    ' 
    ' Contains:
    ' 
    '   (string) - URI
    ' 
    public cbZipUri

    ' Property: zipName
    ' 
    ' The zip filename with extension.
    ' 
    ' Contains:
    ' 
    '   (string) - Filename
    ' 
    public zipName

    ' Property: Files
    ' 
    ' This is the container of all Zip_File objects.
    ' 
    ' Contains:
    ' 
    '   (scripting.dictionary) - Collection with the Zip_File's
    ' 
    public Files

    private Xml

    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "1.1.0"
        
        set Files = Server.createObject("Scripting.Dictionary")
    end sub

    private sub Class_terminate()
        set Files = nothing
        set Xml = nothing
    end sub

    private sub createFolder(sFolderPath)
        dim Fso
        set Fso = Server.createObject("Scripting.FileSystemObject")
        if(not Fso.folderExists(sFolderPath)) then
            Fso.createFolder(sFolderPath)
        end if
        set Fso = nothing
    end sub

    ' Subroutine: go
    ' 
    ' Parses the zip file and create the Zip_File objects.
    ' 
    public sub go()
        dim Xhr, Coll, Node, sCmd, i, File
        sCmd = cbZipUri & "?filename=" & zipName
        set Xhr = Server.createObject("MSXML2.ServerXMLHTTP.6.0")
        Xhr.open "GET", sCmd, false
        Xhr.send
        set Xml = Xhr.responseXML
        set Xhr = nothing
        
        if(Xml.parseError <> 0) then
            Xml.loadXml("<root><errorcode>" & Xml.parseError & "</errorcode><errortext>" & Xml.parseError.reason & "</errortext></root>")
        end if
        
        if(Xml.documentElement is nothing) then
            Xml.loadXml("<root><errorcode>-1</errorcode><errortext>Invalid XML returned. Check your parameters. (" & sCmd & ")</errortext></root>")
        end if
        
        set Coll = Xml.documentElement.selectNodes("zipinfo/files/file")
        for i = 0 to Coll.length - 1
            set File = new Zip_File
            File.name = Coll(i).selectSingleNode("filename").text
            File.compMethod = Coll(i).selectSingleNode("compmethod").text
            File.compSize = Coll(i).selectSingleNode("compsize").text
            File.uncompSize = Coll(i).selectSingleNode("uncompsize").text
            File.compRatio = Coll(i).selectSingleNode("compratio").text
            File.crc32 = Coll(i).selectSingleNode("crc32").text
            if(File.uncompSize = "") then
                File.isFolder = true
            else
                File.isFolder = false
            end if
            File.uri = cbZipUri & "/getfile?filename=" & zipName & "&target=" & File.name
            call Files.add(File.name, File)
            set File = nothing
        next
        set Coll = nothing
    end sub

    ' Subroutine: extract
    ' 
    ' Extract the files in the standard way.
    ' 
    ' Parameters:
    ' 
    '   (string) - Virtual Path where you intend to save your extracted files.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim oZip
    ' set oZip = new Zip
    ' oZip.cbZipUri = "http://localhost/cgi-bin/CBZIP.exe"
    ' oZip.zipName = Server.mapPath("docs.zip")
    ' call oZip.go()
    ' oZip.extract("files/") 'Remember that "files/" must exist before calling the extract.
    ' set oZip = nothing
    ' 
    ' (end)
    ' 
    public sub extract(sTargetPath)
        dim File
        for each File in Files
            set File = Files(File)
            if(not File.isFolder) then
                File.saveToFile(Server.mapPath(sTargetPath & File.name))
            else
                createFolder(Server.mapPath(sTargetPath & File.name))
            end if
            if(not Response.isClientConnected) then
                exit for
            end if
        next
        set File = nothing
    end sub

    ' Subroutine: extractWithoutFolders
    ' 
    ' Extract the files without caring about their relative path inside the zip.
    ' 
    ' Parameters:
    ' 
    '   (string) - Virtual Path where you intend to save your extracted files.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim oZip
    ' set oZip = new Zip
    ' oZip.cbZipUri = "http://localhost/cgi-bin/CBZIP.exe"
    ' oZip.zipName = Server.mapPath("docs.zip")
    ' call oZip.go()
    ' oZip.extractWithoutFolders("files/") 'Remember that "files/" must exist before calling the extract.
    ' set oZip = nothing
    ' 
    ' (end)
    ' 
    public sub extractWithoutFolders(sTargetPath)
        dim File
        for each File in Files
            set File = Files(File)
            if(not File.isFolder) then
                File.saveToFile(Server.mapPath(sTargetPath & File.fileName))
            end if
            if(not Response.isClientConnected) then
                exit for
            end if
        next
        set File = nothing
    end sub

    ' Function: inspect
    ' 
    ' Retrive a view from inside the zip.
    ' 
    ' Returns:
    ' 
    '   (string) - HTML table with the zip contents
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim oZip
    ' set oZip = new Zip
    ' oZip.cbZipUri = "http://localhost/cgi-bin/CBZIP.exe"
    ' oZip.zipName = Server.mapPath("docs.zip")
    ' call oZip.go()
    ' Response.write(oZip.inspect())
    ' set oZip = nothing
    ' 
    ' (end)
    ' 
    public function inspect()
        dim sXslt
        sXslt = join(array( _
        "<xsl:transform xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" version=""1.0"">", _
        "<xsl:output indent=""yes"" omit-xml-declaration=""yes"" />", _
        "<xsl:variable name=""headerFilename"">Filename</xsl:variable>", _
        "<xsl:variable name=""headerCompressedSize"">Compressed size</xsl:variable>", _
        "<xsl:variable name=""headerUncompressedSize"">Uncompressed size</xsl:variable>", _
        "<xsl:variable name=""headerCRC32"">CRC32</xsl:variable>", _
        "<xsl:variable name=""showCRC32"">1</xsl:variable>", _
        "<xsl:variable name=""enableDownload"">1</xsl:variable>", _
        "<xsl:variable name=""cbZipUri"">"& cbZipUri &"</xsl:variable>", _
        "<xsl:template match=""/"">", _
        "<table>", _
        "  <tr>", _
        "    <th><xsl:value-of select=""$headerFilename"" /></th>", _
        "    <th><xsl:value-of select=""$headerCompressedSize"" /></th>", _
        "    <th><xsl:value-of select=""$headerUncompressedSize"" /></th>", _
        "    <xsl:if test=""$showCRC32 = 1"">", _
        "      <th><xsl:value-of select=""$headerCRC32"" /></th>", _
        "    </xsl:if>", _
        "  </tr>", _
        "  <xsl:for-each select=""root/zipinfo/files/file"">", _
        "  <xsl:sort select=""filename"" order=""ascending"" />", _
        "  <tr>", _
        "    <xsl:if test=""(position() mod 2) = 1"">", _
        "      <xsl:attribute name=""class"">even</xsl:attribute>", _
        "    </xsl:if>", _
        "    <td>", _
        "      <xsl:choose>", _
        "        <xsl:when test=""$enableDownload = 1"">", _
        "          <a title=""get it"">", _
        "            <xsl:attribute name=""href""><xsl:value-of select=""$cbZipUri"" />/getfile?filename=<xsl:value-of select=""/root/zipinfo/filename"" />&amp;target=<xsl:value-of select=""filename"" /></xsl:attribute>", _
        "            <xsl:value-of select=""filename"" />", _
        "          </a>", _
        "        </xsl:when>", _
        "        <xsl:otherwise>", _
        "          <xsl:value-of select=""filename"" />", _
        "        </xsl:otherwise>", _
        "      </xsl:choose>", _
        "    </td>", _
        "    <xsl:choose>", _
        "      <xsl:when test=""string-length(compsize)=0"">", _
        "        <td align=""right""></td>", _
        "        <td align=""right""></td>", _
        "      </xsl:when>", _
        "      <xsl:otherwise>", _
        "        <td align=""right""><xsl:value-of select=""compsize"" /> bytes</td>", _
        "        <td align=""right""><xsl:value-of select=""uncompsize"" /> bytes</td>", _
        "      </xsl:otherwise>", _
        "    </xsl:choose>", _
        "    <xsl:if test=""$showCRC32 = 1"">", _
        "      <td align=""right""><xsl:value-of select=""crc32"" /></td>", _
        "    </xsl:if>", _
        "  </tr>", _
        "  </xsl:for-each>", _
        "</table>", _
        "</xsl:template>", _
        "</xsl:transform>"), vbNewLine)
        set Xslt = Server.createObject("MSXML2.DOMDocument.6.0")
        Xslt.loadXml(sXslt)
        inspect = Xml.transformNode(Xslt)
        set Xslt = nothing
    end function

end class

' Class: Zip_File
' 
' Each entry of Zip.Files contains an object of this class. This is the class
' that the user should use to retrieve information from a file in the zip.
' 
' About:
' 
'   - Written by Karl-Johan Sjögren <http://www.crazybeavers.se/> @ February 2006
'   - Modified and normalized by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ January 2008
' 
class Zip_File

    ' Property: name
    ' 
    ' Relative path plus the filename with it extension.
    ' 
    ' Contains:
    ' 
    '   (string) - File full path.
    ' 
    public name

    ' Property: compMethod
    ' 
    ' Compression method.
    ' 
    ' Contains:
    ' 
    '   (string) - Stored, Deflated, etc
    ' 
    public compMethod

    ' Property: compSize
    ' 
    ' The size of the file when it's compressed.
    ' 
    ' Contains:
    ' 
    '  (int) - Compressed file size
    ' 
    public compSize

    ' Property: uncompSize
    ' 
    ' The size of the file.
    ' 
    ' Contains:
    ' 
    '   (int) - File size
    ' 
    public uncompSize

    ' Property: compRatio
    ' 
    ' Compression effectiveness.
    ' 
    ' Contains:
    ' 
    '   (int) - Compression ratio
    ' 
    public compRatio

    ' Property: crc32
    ' 
    ' The CRC32 code which is a 32 bits length hash.
    ' 
    ' Contains:
    ' 
    '   (hex) - CRC32
    ' 
    public crc32

    ' Property: isFolder
    ' 
    ' A flag indicating if the entry is a folder or not.
    ' 
    ' Contains:
    ' 
    '   true - if it's a folder
    '   false - otherwise
    ' 
    public isFolder

    ' Property: uri
    ' 
    ' Uri for downloading the file.
    ' 
    ' Contains:
    ' 
    '   (string) - Href to download the file
    ' 
    public uri

    ' Function: fileName
    ' 
    ' Retrieve the file name.
    ' 
    ' Returns:
    ' 
    '   (string) - File name
    ' 
    public function fileName()
        dim aFileName : aFileName = split(name, "/")
        fileName = aFileName(uBound(aFileName))
    end function

    ' Function: fileExt
    ' 
    ' Retrieve the file extension.
    ' 
    ' Returns:
    ' 
    '   (string) - File extension
    ' 
    public function fileExt()
        dim aFileName : aFileName = split(name, ".")
        fileExt = aFileName(uBound(aFileName))
    end function

    ' Function: saveToFile
    ' 
    ' Save the file to the hard drive.
    ' 
    ' Parameters:
    ' 
    '   (string) - Full path
    ' 
    ' Returns:
    ' 
    '   true - if it's successfully saved
    '   false - otherwise
    ' 
    public function saveToFile(sPath)
        dim Xhr, Stream, data
        saveToFile = true
        on error resume next
        
        set Xhr = Server.createObject("MSXML2.ServerXMLHTTP.6.0")
        Xhr.open "GET", uri, false
        Xhr.send
        data = Xhr.responseBody
        set Xhr = nothing
        
        set Stream = Server.createObject("ADODB.Stream")
        Stream.mode = 3 'adModeReadWrite
        Stream.type = 1 'adTypeBinary
        Stream.open
        Stream.write(data)
        call Stream.saveToFile(sPath, 2) 'adSaveCreateOverwrite
        Stream.close
        set Stream = nothing
        
        if(Err <> 0) then
            saveToFile = false
        end if
        on error goto 0
    end function

end class

%>
