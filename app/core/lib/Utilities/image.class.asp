<%

'+-----------------------------------------------------------------------------+
'|This file is part of ASP Xtreme Evolution.                                   |
'|Copyright (C) 2007, 2008 Fabio Zendhi Nagao                                  |
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

' Class: Image
' 
' This class intends to help people use the Imager in an easy way. Imager brings
' some very useful image manipulations like resizing, stretching, rotating to
' asp. Then the result can be used directly from the memory or saved to the hard
' drive. You can also use it to retrieve image properties like width, height.
' 
' Requirements:
' 
'   - CrazyBeavers Imager Resizer dll (Imager.dll)
' 
' About:
' 
'   - Written by Karl-Johan Sjögren <http://www.crazybeavers.se/> @ April 2006
'   - Modified and normalized by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ January 2008
' 
class Image

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

    ' Property: imagerUri
    ' 
    ' This should be the URI to Imager.dll including http:// etc.
    ' 
    ' Contains:
    ' 
    '   (string) - URI
    ' 
    public imagerUri

    ' Property: image
    ' 
    ' Full path to the image.
    ' 
    ' Contains:
    ' 
    '   (string) - Physical path
    ' 
    public image

    ' Property: width
    ' 
    ' Target width.
    ' 
    ' Contains:
    ' 
    '   (int) - Width
    ' 
    public width

    ' Property: height
    ' 
    ' Target height.
    ' 
    ' Contains:
    ' 
    '   (int) - Height
    ' 
    public height

    ' Property: compression
    ' 
    ' Target compression.
    ' 
    ' Contains:
    ' 
    '   (int) - Compression
    ' 
    public compression

    ' Property: output
    ' 
    ' Output format.
    ' 
    ' Contains:
    ' 
    '   (string) - .jpg .gif .png and .tga. Note that it won't work without the dot at start.
    ' 
    public output

    ' Property: originalWidth
    ' 
    ' Original width. This can be retrieved with ProcessBinary=false.
    ' 
    ' Contains:
    ' 
    '   (int) - Width
    ' 
    public originalWidth

    ' Property: originalHeight
    ' 
    ' Original height. This can be retrieved with ProcessBinary=false.
    ' 
    ' Contains:
    ' 
    '   (int) - Height
    ' 
    public originalHeight

    ' Property: autoRotate
    ' 
    ' Use the autorotate function?
    ' 
    ' Contains:
    ' 
    '   (bool) - Defaults to true
    ' 
    public autoRotate

    ' Property: whitespace
    ' 
    ' Use the whitespace function?
    ' 
    ' Contains:
    ' 
    '   (bool) - Defaults to false
    ' 
    public whitespace

    ' Property: bgColor
    ' 
    ' 6 letter HEX color code used for whitespace/rotation functions.
    ' 
    ' Contains:
    ' 
    '   (string) - Defaults to an empty string
    ' 
    public bgColor

    ' Property: rotation
    ' 
    ' The amount of degrees to rotate the image.
    ' 
    ' Contains:
    ' 
    '   (int) - Defaults to 0
    ' 
    public rotation

    ' Property: processExif
    ' 
    ' Extract EXIF data?
    ' 
    ' Contains:
    ' 
    '   (bool) - Defaults to true
    ' 
    public processExif

    ' Property: processBinary
    ' 
    ' Return binary data?
    ' 
    ' Contains:
    ' 
    '   (bool) - Defaults to true
    ' 
    public processBinary

    ' Property: forceNewCache
    ' 
    ' Force the creation of a new file in the cache?
    ' 
    ' Contains:
    ' 
    '   (bool) - Defaults to false
    ' 
    public forceNewCache

    ' Property: useQueryString
    ' 
    ' Append the current querystring to the dll?
    ' 
    ' Contains:
    ' 
    '   (bool) - Defaults to true
    ' 
    public useQueryString

    ' Property: errorCode
    ' 
    ' Error code from the query.
    ' 
    ' Contains:
    ' 
    '   (int) - Default is 0 which means no error
    ' 
    public errorCode

    ' Property: errorText
    ' 
    ' Error description.
    ' 
    ' Returns:
    ' 
    '   (string) - Error description
    ' 
    public errorText

    ' Property: Xml
    ' 
    ' XML Dom object which holds the returned xml.
    ' 
    ' Returns:
    ' 
    '   (xml object) - XML
    ' 
    public Xml

    public sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "1.1.0"
        
        imagerUri = ""
        image = ""
        width = -1
        height = -1
        compression = -1
        output = ""
        originalWidth = -1
        originalHeight = -1
        autoRotate = true
        whitespace = false
        bgColor = ""
        rotation = 0
        processExif = true
        processBinary = true
        forceNewCache = false
        useQueryString = false
        errorCode = 0
        errorText = ""
    end sub

    public sub Class_terminate()
        set Xml = nothing
    end sub

    private function bool2str(value)
        if(value) then
            bool2str = "true"
        else
            bool2str = "false"
        end if
    end function

    ' Subroutine: go
    ' 
    ' Executes the request to the application to retrieve the XML data.
    ' 
    public sub go()
        dim sCmd, Xhr, Node
        
        if(useQueryString) then
            sCmd = imagerUri & "/xml?" & Request.serverVariables("QUERY_STRING")
        else
            sCmd = imagerUri & "/xml?Image=" & image & "&Width=" & width & "&Height=" & height & "&Autorotate=" & bool2str(autoRotate) & "&Whitespace=" & bool2str(whitespace) & "&BGColor=" & bgColor & "Rotation=" & rotation & "&Compression=" & compression & "&Output=" & output & "&ProcessExif=" & bool2str(processExif) & "&ProcessBinary=" & bool2str(processBinary) & "&ForceNewCache=" & bool2str(forceNewCache)
        end if
        
        set Xhr = Server.createObject("MSXML2.ServerXMLHTTP.6.0")
        Xhr.open "GET", sCmd, false
        Xhr.send()
        set Xml = Xhr.responseXml
        set Xhr = nothing
        
        if(Xml.parseError <> 0) then
            Xml.loadXml("<root><errorcode>" & Xml.parseError & "</errorcode><errortext>" & Xml.parseError.reason & "</errortext></root>")
        end if
        
        if(Xml.documentElement is nothing) then
            Xml.loadXML("<root><errorcode>-1</errorcode><errortext>Invalid XML returned. Check your parameters. (" & sCmd & ")</errortext></root>")
        end if
        
        set Node = Xml.selectSingleNode("/root/errorcode")
        if(not Node is nothing) then
            errorCode = Node.text
        end if
        set Node = Xml.selectSingleNode("/root/errortext")
        if(not Node is nothing) then
            errorText = Node.text
        end if
        
        set Node = Xml.selectSingleNode("/root/imageinfo/originalwidth")
        if(not Node is nothing) then
            OriginalWidth = Node.text
        end if
        set Node = Xml.selectSingleNode("/root/imageinfo/originalheight")
        if(not Node is nothing) then
            OriginalHeight = Node.text
        end if
        
        set Node = nothing
    end sub

    ' Function: saveToFile
    ' 
    ' Saves the image to a path specified by path.
    ' 
    ' Parameters:
    ' 
    '   (string)  - Path to write the file
    '   (boolean) - Overwrite true|false
    ' 
    ' Returns:
    ' 
    '   true - if the save was successful
    '   false - otherwise
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim oImage
    ' set oImage = new Image
    ' oImage.imagerUri = "http://localhost/cgi-bin/Imager.dll"
    ' oImage.image = Server.mapPath("image.png")
    ' oImage.width = 88
    ' oImage.height = 88
    ' oImage.compression = 88
    ' oImage.output = ".jpg"
    ' oImage.whitespace = true
    ' oImage.bgColor = "FFFFFF"
    ' oImage.rotation = 15
    ' call oImage.go()
    ' if(oImage.ErrorCode <> 0) then
    '     Response.write("<strong>Error:</strong> " & oImage.errorText)
    '     Response.end
    ' end if
    ' if(not oImage.saveToFile(Server.mapPath("saved/image.jpg"), true)) then
    '     Response.write(oImage.errorText)
    '     Response.end
    ' else
    '     Response.write("Saved")
    ' end if
    ' set oImage = nothing
    ' 
    ' (end)
    ' 
    public function saveToFile(path, overwrite)
        dim Stream, Node
        
        set Node = Xml.selectSingleNode("/root/imageinfo/imagedata")
        if(not Node is nothing) then
            set Stream = Server.createObject("ADODB.Stream")
            Stream.type = 1 'adTypeBinary
            Stream.mode = 3 'adModeReadWrite
            Stream.open
            Stream.write Node.nodeTypedValue
            Stream.position = 0
            if(overwrite) then
                overwrite = 2 'adSaveCreateOverWrite
            else
                overwrite = 1 'adSaveCreateNotExist
            end if
            
            on error resume next
            call Stream.saveToFile(path, overwrite)
            if(Err <> 0) then
                saveToFile = false
                errorCode = -2
                errorText = "Image.saveToFile failed: #" & Err.number & " - " & Err.description
            else
                saveToFile = true
            end if
            on error goto 0
            Stream.close
            set Stream = nothing
        else
            saveToFile = false
        end if
        set Node = nothing
    end function

    ' Function: getExif
    ' 
    ' Extract EXIF(EXchangeable Image Format) data from the Xml.
    ' 
    ' Returns:
    ' 
    '   (scripting.dictionary) - With the extracted EXIF data
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim oImage, Exif, sItem
    ' set oImage = new Image
    ' oImage.imagerUri = "http://localhost/cgi-bin/Imager.dll"
    ' oImage.image = Server.mapPath("image.png")
    ' call oImage.go()
    ' if(oImage.errorCode <> 0) then
    '     Response.write("<strong>Error:</strong> " & oImage.errorText)
    '     Response.end
    ' end if
    ' set Exif = oImage.getExif()
    ' for each sItem in Exif
    '     Response.write(sItem & " : " & Exif.item(sItem) & "<br />")
    ' next
    ' set Exif = nothing
    ' set oImage = nothing
    ' 
    ' (end)
    ' 
    public function getExif()
        dim Node, Collection
        set getExif = Server.createObject("Scripting.Dictionary")
        set Collection = Xml.selectSingleNode("/root/exifdata")
        if(not Collection is nothing) then
            for each Node in Collection.childNodes
                getExif.add Node.nodeName, Node.text
            next
        end if
        set Node = nothing
        set Collection = nothing
    end function

    ' Function: getBinary
    ' 
    ' Extract binary data from the Xml.
    ' 
    ' Returns:
    ' 
    '   (string) - With the image binary data
    ' 
    public function getBinary()
        dim Node
        set Node = Xml.selectSingleNode("/root/imageinfo/imagedata")
        if(not Node is nothing) then
            getBinary = Node.nodeTypedValue
        else
            getBinary = ""
        end if
        set Node = nothing
    end function

    ' Function: getMime
    ' 
    ' Extract MimeType from the Xml.
    ' 
    ' Returns:
    ' 
    '   (string) - Mimetype of the image
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim oImage
    ' set oImage = new Image
    ' oImage.imagerUri = "http://localhost/cgi-bin/Imager.dll"
    ' oImage.image = Server.mapPath("image.png")
    ' call oImage.go()
    ' if(oImage.errorCode <> 0) then
    '     Response.write("<strong>Error:</strong> " & oImage.errorText)
    '     Response.end
    ' end if
    ' Response.write(oImage.getMime())
    ' set oImage = nothing
    ' 
    ' (end)
    ' 
    public function getMime()
        dim Node
        set Node = Xml.selectSingleNode("/root/imageinfo/mime")
        if(not Node is nothing) then
            getMime = Node.text
        else
            getMime = ""
        end if
        set Node = nothing
    end function

    ' Function: getFilename
    ' 
    ' Extract filename from the imageinfo in Xml.
    ' 
    ' Returns:
    ' 
    '   (string) - Filename of the image
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim oImage
    ' set oImage = new Image
    ' oImage.imagerUri = "http://localhost/cgi-bin/Imager.dll"
    ' oImage.image = Server.mapPath("image.png")
    ' call oImage.go()
    ' if(oImage.errorCode <> 0) then
    '     Response.write("<strong>Error:</strong> " & oImage.errorText)
    '     Response.end
    ' end if
    ' Response.write(oImage.getFilename())
    ' set oImage = nothing
    ' 
    ' (end)
    ' 
    public function getFilename()
        dim Node
        set Node = Xml.selectSingleNode("/root/imageinfo/filename")
        if(not Node is nothing) then
            getFilename = Node.text
        else
            getFilename = ""
        end if
        set Node = nothing
    end function

end class

%>
