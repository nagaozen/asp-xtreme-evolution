<%

' File: kernel.asp
' 
' AXE(ASP Xtreme Evolution) framework kernel.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2007-2009 Fabio Zendhi Nagao
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



' Class: Kernel
' 
' This Class will be auto-loaded as the singleton "Core" for the entire
' application and it's objective is to provide support for the main features
' of this framework: MVC, URL-Rewriting, Cache System and XSLT.
' 
' Notes:
' 
'   - Some of the MVC approach is inspired by MvCasp <http://robrohan.com/2007/03/04/mvcasp-in-public-domain/> Copyright (c) 2007 Rob Rohan.
'   - Inpect mode is inspired in the debug mode from Symphony CMS <http://symphony21.com/>.
' 
' About:
' 
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2007
' 
class Kernel

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
    
    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "1.1.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: initialize
    ' 
    ' Initialize the application configuration, setup the (string)controller,
    ' (string)action and the (string[])argv sessions which will be used to run
    ' the MVC architecture.
    ' 
    ' Returns:
    ' 
    '   (Kernel) - Self instance to enable chaining
    ' 
    public function initialize()
        dim sController, sAction, sArgv
        sController = cStr(Request.QueryString("controller"))
        sAction = cStr(Request.QueryString("action"))
        sArgv = cStr(Request.QueryString("argv"))
        
        if( ( not Application("isConfigured") ) or ( strComp(lcase(Request.QueryString("reconfigure")), "true") = 0 ) or ( inStr(lcase(Request.QueryString("argv")), "reconfigure=true") > 0 ) ) then
            Server.execute("/lib/axe/application.configure.asp")
        end if
        
        if(sController <> "") then
            Session("controller") = sController
        end if
        
        if(sAction <> "") then
            Session("action") = sAction
        end if
        
        if(sArgv <> "") then
            Session("argv") = split(sArgv, "/")
        end if
        
        set initialize = Me
    end function

    ' Function: process
    '
    ' Process the controller if it exists.
    ' 
    ' Returns:
    ' 
    '   (Kernel) - Self instance to enable chaining
    ' 
    public function process()
        ' Transfer to cached version when in production environment
        if(strComp(Application("environment"), "production") = 0) then
            dim iCache, sCachePath
            iCache = cacheIndex(Session("controller"), Session("action"))
            if(iCache >= 0) then
                sCachePath = strsubstitute("/app/cache/__{0}_{1}_{2}__.html", array(Session("controller"), Session("action"), join(Session("argv"), "_")))
                if(fileExists(Server.mapPath(sCachePath))) then
                    if(dateDiff("h", Application("Cache.items")(iCache, 2), now()) <= Application("Cache.lifetime")) then
                        Session.abandon()
                        Server.transfer(sCachePath)
                    end if
                end if
            end if
        end if
        
        dim vpController
        vpController = "/app/controllers/" & Session("controller") & ".asp"
        if(fileExists(Server.mapPath(vpController))) then
            Server.execute(vpController)
        else
            Err.raise 53, "Evolved AXE runtime error", strsubstitute( _
                "Controller '{0}' source file '{1}' not found.", _
                array(Session("controller"), vpController) _
            )
        end if
        
        set process = Me
    end function

    ' Function: computeView
    ' 
    ' Compute the requested view with the current data and return it's source to
    ' be used inside the current process.
    ' 
    public function computeView()
        dim Xhr : set Xhr = Server.createObject("MSXML2.ServerXMLHTTP.6.0")
        Xhr.open "POST", strsubstitute("{0}/app/views/{1}.asp", array(Application("uri"), Session("view"))), false
        Xhr.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        Xhr.send(loadShuttle(Session("this")))
        computeView = Xhr.responseText
        set Xhr = nothing
    end function

    ' Function: dispatch
    ' 
    ' If no error exists, send the Output.value string to user. Otherwise,
    ' display the errors.
    ' 
    ' Returns:
    ' 
    '   (Kernel) - Self instance to enable chaining
    ' 
    public function dispatch()
        dim iCache
        
        ' Handle cache
        iCache = cacheIndex(Session("controller"), Session("action"))
        if(iCache >= 0) then
            createFile Server.mapPath( _
                sanitize( _
                    strsubstitute("/app/cache/__{0}_{1}_{2}__.html", array(Session("controller"), Session("action"), join(Session("argv"), "_"))), _
                    array("\",":","*","?", """","<",">","|"), _
                    array("" ,"" ,"" ,"" , ""  ,"" ,"" ,"") _
                ) _
            ), Session("this").item("Output.value")
            Application("Cache.items")(iCache, 2) = Now()
        end if
        
        ' Display
        if( ( strComp(Application("environment"), "development") = 0 ) and ( ( ( strComp(lcase(Request.QueryString("inspect")), "true") = 0 ) ) or ( inStr(lcase(Request.QueryString("argv")), "inspect=true") > 0 ) ) ) then
            Server.transfer("/app/views/inspect.asp")
        else
            Response.write(Session("this").item("Output.value"))
        end if
        
        set dispatch = Me
    end function

    ' Function: cacheIndex
    ' 
    ' Informs the request page cache index
    ' 
    ' Parameters:
    ' 
    '   (string) - Controller name
    '   (string) - Action name
    ' 
    ' Returns:
    ' 
    '   (integer) - (-1) if it's not in the cache array, array index otherwise
    ' 
    ' See also:
    ' 
    '   <config.xml>
    ' 
    public function cacheIndex(sController, sAction)
        cacheIndex = -1
        dim i : for i = 0 to uBound(Application("Cache.items")) - 1
            if(strComp(sController, Application("Cache.items")(i, 0)) = 0) then
                if(strComp(sAction, Application("Cache.items")(i, 1)) = 0) then
                    cacheIndex = i
                    exit function
                end if
            end if
        next
    end function

    ' Function: loadShuttle
    ' 
    ' A shuttle is needed to carry the user Session("this") information to the
    ' server Session("this") in order to the view be a document and not a
    ' variable. This function loads everything in Session("this") into a POST
    ' string.
    ' 
    ' Parameters:
    ' 
    '   (scripting dictionary) - The dictionary to be encoded
    ' 
    ' Returns:
    ' 
    '   (string) - Private representation of the dictionary
    ' 
    public function loadShuttle(Sd)
        if( Sd.count = 0 ) then loadShuttle = "" : exit function
        
        dim Stream : set Stream = Server.createObject("ADODB.Stream")
        Stream.type = adTypeText
        Stream.mode = adModeReadWrite
        Stream.open()
        
        dim aKeys : aKeys = Sd.keys
        dim i : for i = 0 to Sd.count - 1
            Stream.writeText(strsubstitute("&{0}={1}", array(aKeys(i), Server.urlEncode(Sd.item(aKeys(i))))))
        next
        
        Stream.position = 0
        loadShuttle = mid(Stream.readText(), 2)
        
        Stream.close()
        set Stream = nothing
    end function

    ' Subroutine: unloadShuttle
    ' 
    ' Unload a POST string into the Session("this") scripting dictionary.
    ' 
    public sub unloadShuttle()
        dim key : for each key in Request.Form
            Session("this").add key, Request.Form(key)
        next
    end sub

    ' Function: fileExists
    ' 
    ' Checks if the file exists in the file system.
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
    public function fileExists(sFilePath)
        dim Fso : set Fso = Server.createObject("Scripting.FileSystemObject")
        fileExists = Fso.fileExists(sFilePath)
        set Fso = nothing
    end function

    ' Function: loadTextFile
    ' 
    ' If the file exists, read it all and return the content.
    ' 
    ' Parameters:
    ' 
    '   (string) - Full path plus file name with extension
    ' 
    ' Returns:
    ' 
    '   (string) - The file content
    ' 
    public function loadTextFile(sFilePath)
        if(fileExists(sFilePath)) then
            dim Stream : set Stream = Server.createObject("ADODB.Stream")
            with Stream
                .type = adTypeText
                .mode = adModeReadWrite
                .charset = "UTF-8"
                .open()
                
                .loadFromFile(sFilePath)
                .position = 0
                loadTextFile = .readText()
                
                .close()
            end with
            set Stream = nothing
        else
            Err.raise 53, "Evolved AXE runtime error"
        end if
    end function

    ' Subroutine: createFile
    ' 
    ' Create || Overwrite a file in the file system.
    ' 
    ' Parameters:
    ' 
    '   (string) - Full path plus file name with extension
    '   (string) - File content
    ' 
    public sub createFile(sFilePath, sContent)
        dim Stream : set Stream = Server.createObject("ADODB.Stream")
        with Stream
            .type = adTypeText
            .mode = adModeReadWrite
            .charset = "UTF-8"
            .open()
            
            .writeText(sContent & vbNewLine)
            .writeText("<!--// CACHED FILE => Execution time is less than 1µs (Blazing fast performance) //-->")
            .setEOS()
            .position = 0
            .saveToFile sFilePath, adSaveCreateOverwrite
            
            .close()
        end with
        set Stream = nothing
    end sub

    ' Function: createLink
    ' 
    ' Creates a friendly relative link to be used as an URL-Rewrite hyperlink
    ' reference.
    ' 
    ' Parameters:
    ' 
    '   (string) - Controller name
    '   (string) - Action name
    ' 
    ' Returns:
    ' 
    '   (string) - Reference to be used in action or href
    ' 
    public function createLink(sController, sAction)
        createLink = strsubstitute("/{0}/{1}/", array(sController, sAction))
    end function

    ' Function: str2xml
    ' 
    ' Creates a XML Document from a string with it's source code.
    ' 
    ' Parameters:
    ' 
    '   (string) - The XML source code
    ' 
    ' Returns:
    ' 
    '   (xml object) - The XML
    ' 
    public function str2xml(sXml)
        set str2xml = Server.createObject("MSXML2.DOMDocument.6.0")
        str2xml.loadXML(sXml)
        if str2xml.parseError.errorCode <> 0 then
            dim ErrXml : set ErrXml = str2xml.parseError
            Err.raise 51, "Evolved AXE runtime error", strsubstitute("Internal error while parsing XML data. {0}", array(ErrXml.reason))
            set ErrXml = nothing
            Server.execute("/app/views/error.asp")
        end if
    end function

    ' Function: getXmlNodeValues
    ' 
    ' Creates a string array with the text of all nodes that match with the
    ' XPath.
    ' 
    ' Parameters:
    ' 
    '   (xml object) - The XML
    '   (string)     - The XPath
    ' 
    ' Returns:
    ' 
    '   (string[]) - An array with the XPath Node.text values
    ' 
    public function getXmlNodeValues(Xml, sXPath)
        dim Nodelist : set Nodelist = Xml.selectNodes(sXPath)
        dim sa : redim sa(Nodelist.length)
        dim i : i = 0
        dim Node : for each Node In Nodelist
            sa(i) = Node.text
            i = i + 1
        next
        set Node = nothing
        set Nodelist = nothing
        getXmlNodeValues = sa
    end function

    ' Function: strictTransform
    ' 
    ' Applies the xslt transformation AS IT IS, in other words it relies only in
    ' the built-in parser to do exactly what your transformation requested.
    ' 
    ' Parameters:
    ' 
    '   (xml object) - The dataset
    '   (xml object) - The transformation
    ' 
    ' Returns:
    ' 
    '   (string) - The evaluated transformation
    ' 
    ' See also:
    ' 
    '   <indentedTransform>
    ' 
    public function strictTransform(Xml, Xslt)
        strictTransform = Xml.transformNode(Xslt)
    end function

    ' Function: indentedTransform
    ' 
    ' Since MSXML are very conservative about how much whitespace they put into
    ' serialized XSLT results when xsl:output indent="yes", we are properly
    ' indenting XML, without relying on the processor's built-in indenting
    ' capability. This is useful if you are trying to print nicely indented
    ' outputs.
    ' 
    ' Parameters:
    ' 
    '   (xml object)Xml  - The dataset
    '   (xml object)Xslt - The transformation
    '   (string)sOutput  - The xsl:output directive
    '   (string)sIndent  - The string which will be used as indentation
    ' 
    ' Returns:
    ' 
    '   (string) - The evaluated nicely indented transformation
    ' 
    ' See also:
    ' 
    '   <strictTransform>
    ' 
    ' Notes:
    ' 
    ' This XSLT is strongly based on Mike J. Brown mike@skew.org ReIndent
    ' work: <http://skew.org/xml/stylesheets/reindent/reindent.xsl>
    ' 
    public function indentedTransform(Xml, Xslt, sOutput, sIndent)
        dim sPoorlyIndented : sPoorlyIndented = strictTransform(Xml, Xslt)
        dim sReIndent : sReIndent = join(array(_
            "<?xml version=""1.0"" encoding=""UTF-8""?>", _
            "<xsl:transform xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" version=""1.0"">", _
            "  " & sOutput, _
            "  <xsl:param name=""delete_comments"" select=""false()""/>", _
            "  <xsl:param name=""indent_string"" select=""'" & sIndent & "'""/>", _
            "  <xsl:strip-space elements=""*""/>", _
            "  <xsl:preserve-space elements=""xsl:text""/>", _
            "  <xsl:template name=""whitespace-before"">", _
            "    <xsl:if test=""ancestor::*"">", _
            "      <xsl:text>&#10;</xsl:text>", _
            "    </xsl:if>", _
            "    <xsl:for-each select=""ancestor::*"">", _
            "      <xsl:value-of select=""$indent_string""/>", _
            "    </xsl:for-each>", _
            "  </xsl:template>", _
            "  <xsl:template name=""whitespace-after"">", _
            "    <xsl:if test=""not(following-sibling::node())"">", _
            "      <xsl:text>&#10;</xsl:text>", _
            "      <xsl:for-each select=""../ancestor::*"">", _
            "        <xsl:value-of select=""$indent_string""/>", _
            "      </xsl:for-each>", _
            "    </xsl:if>", _
            "  </xsl:template>", _
            "  <xsl:template match=""*"">", _
            "    <xsl:call-template name=""whitespace-before""/>", _
            "    <xsl:copy>", _
            "      <xsl:apply-templates select=""@*|node()""/>", _
            "    </xsl:copy>", _
            "    <xsl:call-template name=""whitespace-after""/>", _
            "  </xsl:template>", _
            "  <xsl:template match=""@*"">", _
            "    <xsl:copy/>", _
            "  </xsl:template>", _
            "  <xsl:template match=""processing-instruction()"">", _
            "    <xsl:call-template name=""whitespace-before""/>", _
            "    <xsl:copy/>", _
            "    <xsl:call-template name=""whitespace-after""/>", _
            "  </xsl:template>", _
            "  <xsl:template match=""comment()"">", _
            "    <xsl:if test=""not($delete_comments)"">", _
            "      <xsl:call-template name=""whitespace-before""/>", _
            "      <xsl:copy/>", _
            "      <xsl:call-template name=""whitespace-after""/>", _
            "    </xsl:if>", _
            "  </xsl:template>", _
            "</xsl:transform>" _
        ))
        
        set Xml = str2xml(sPoorlyIndented)
        set Xslt = str2xml(sReIndent)
        indentedTransform = strictTransform(Xml, Xslt)
        set Xslt = nothing
        set Xml = nothing
    end function

    ' Function: printerFriendlyCode
    ' 
    ' replaces { chr(60) , chr(62) } with their xhtml respectives.
    ' 
    ' Parameters:
    ' 
    '   (string) - The code to be encoded
    ' 
    ' Returns:
    ' 
    '   (string) - Encoded text
    ' 
    public function printerFriendlyCode(sCode)
        sCode = replace(sCode, chr(60), "&lt;", 1, -1, 1)
        sCode = replace(sCode, chr(62), "&gt;", 1, -1, 1)
        printerFriendlyCode = sCode
    end function

    ' Subroutine: exception
    ' 
    ' Triggers a generic internal error exception with the given message.
    ' 
    public sub exception(message)
        Err.raise 51, "Evolved AXE runtime error", strsubstitute("Internal error. {0}", array(message))
    end sub

end class

%>
