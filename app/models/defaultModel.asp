<!--#include virtual="/lib/axe/classes/Parsers/markdown.asp"-->
<%

' File: defaultModel.asp
' 
' ASP Xtreme Evolution after install defaultModel. *You should replace it with your own defaultModel.*
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
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2007
' 
class DefaultModel
    
    public function introduce()
        dim Parser : set Parser = new Markdown
        
        dim sXml : sXml = join(array( _
            "<?xml version=""1.0"" encoding=""UTF-8""?>", _
            "<messages>", _
            "  <message>", _
            "    <img src='/lib/axe/assets/img/logomark-icons.jpg' alt='ASP Xtreme Evolution' title='ASP Xtreme Evolution' />", _
            "    <p>ASP Xtreme Evolution goal is to be a versatile MVC URL-Friendly base for Classic ASP applications with some additional features that are not ASP native. It should implement things that are common to most applications removing the pain of starting a new software and helping you to structure it so that you get things right from the beginning. Our key concepts are choice and freedom over limiting conventions, polyglotism, sustained quality, extensibility which we try to implement in a clean, maintainable and extensible way.</p>", _
            "    <h2>Engine information</h2>", _
            "    <p>" & axeInfo() & "</p>", _
            "    <p>LCID: " & Session.lcid & " - Codepage: " & Session.codepage & " - Charset: " & Response.charset & "</p>", _
            "    <h2>Important Notes</h2>", _
            "    <ul>", _
            "      <li>This page doesn't mean that you installed the entire Framework right, but just a part of it. Don't be lazy and follow the INSTALL.md to the end (a copy of it is also available below).</li>", _
            "      <li>If you are receiving an error: '80020009' without any additional information, increase the metabase 'AspMaxRequestEntityAllowed' parameter size (Default is 204800 bytes).</li>", _
            "      <li>Set the application custom error 500;100 to '/app/views/error.asp' to integrate the framework errors with the IIS ASP errors.</li>", _
            "      <li>There are some very useful templates which all project should use. They are: model.asp, view.asp and controller.asp on /lib/axe/templates folder.</li>", _
            "      <li><strong>Don't forget, this still VBScript ASP! Be confident in your ASP skills and you will be fine.</strong> Also, the <a href='/lib/axe/docs/'><strong>AXE DOCUMENTATION</strong></a> and <a href='http://msdn.microsoft.com/en-us/library/d1et7k7c(VS.85).aspx'><strong>MICROSOFT SCRIPTING TECHNOLOGIES</strong></a> are always ready to help you.</li>", _
            "    </ul>", _
            Parser.makeHtml( Core.loadTextFile( Server.mapPath("/lib/axe/docs/INSTALL.md") ) ), _
            "    <h2>Getting started</h2>", _
            "    <h3>Reconfiguring the Application</h3>", _
            "    <p>append: <a href='?reconfigure=true'>?reconfigure=true</a> in the address bar</p>", _
            "    <h3>Inspecting</h3>", _
            "    <p>append: <a href='?inspect=true'>?inspect=true</a> in the address bar</p>", _
            "    <h3>Checking how it works</h3>", _
            "    <p>append: <a href='/default/another'>/default/another</a> in the address bar for a valid <code>/Controller/action</code> example</p>", _
            "    <p>append: <a href='/controller-alias/another-alias'>/controller-alias/another-alias</a> in the address bar for a valid full routing example</p>", _
            "    <p>append: <a href='/additional-controller-alias/another'>/additional-controller-alias/another</a> in the address bar for a valid simple routing example</p>", _
            "    <p>append: <a href='/foo/bar'>/foo/bar</a> in the address bar for an internal error example</p>", _
            Parser.makeHtml( Core.loadTextFile( Server.mapPath("/lib/axe/docs/CHANGES.md") ) ), _
            "  </message>", _
            "</messages>" _
        ), vbNewLine)
        
        set Parser = nothing
        
        set introduce = Core.str2xml(sXml)
    end function
    
    public function read(kind, name)
        read = Core.printerFriendlyCode( Core.loadTextFile( Server.mapPath( strsubstitute("/app/{0}s/{1}.asp", array(kind, name)) ) ) )
    end function
    
end class

%>
