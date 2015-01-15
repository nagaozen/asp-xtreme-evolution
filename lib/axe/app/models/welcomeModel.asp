<!--#include virtual="/lib/axe/classes/Parsers/markdown.asp"-->
<%

' File: welcomeModel.asp
' 
' ASP Xtreme Evolution after install welcomeModel.
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
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2007
' 
class WelcomeModel
    
    public function introduce()
        dim Parser : set Parser = new Markdown
        
        dim sXml : sXml = join(array( _
            "<?xml version='1.0' encoding='UTF-8'?>", _
            "<messages>", _
            "  <message>", _
            "    <img src='/lib/axe/assets/img/logomark-icons.jpg' alt='ASP Xtreme Evolution' title='ASP Xtreme Evolution' />", _
            "    <p>ASP Xtreme Evolution goal is to be a versatile MVC URL-Friendly base for Classic ASP applications with some additional features that are not ASP native. It should implement things that are common to most applications removing the pain of starting a new software and helping you to structure it so that you get things right from the beginning. Our key concepts are choice and freedom over limiting conventions, polyglotism, sustained quality, extensibility which we try to implement in a clean, maintainable and extensible way.</p>", _
            "    <h2>Engine information</h2>", _
            "    <p>" & axeInfo() & "</p>", _
            "    <p>LCID: " & Session.lcid & " - Codepage: " & Session.codepage & " - Charset: " & Response.charset & "</p>", _
            Parser.makeHtml( Core.loadTextFile( Server.mapPath("/lib/axe/docs/NOTES.md") ) ), _
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
    
    public function read(path)
        read = Core.printerFriendlyCode( Core.loadTextFile( Server.mapPath(path) ) )
    end function
    
end class

%>
