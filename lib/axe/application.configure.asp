<!--#include virtual="/lib/axe/base.asp"-->
<%

' File: application.configure.asp
' 
' It's here where the config file is parsed. This is automatically executed in
' the first time the application runs, but you can send "?reconfigure=true" in
' the queryString to request a new execution.
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
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2007
' 
dim Xml : set Xml = Server.createObject("MSXML2.DOMDocument.6.0")
Xml.async = false
Xml.load(Server.mapPath("/app/config.xml"))

loadCommon Xml
loadCache Xml
loadRoutes Xml

Application("isConfigured") = true

set Xml = nothing



' Subroutine: loadCommon
' 
' Loads common application configurations.
' 
' Parameters:
' 
'   (xml object)Xml - config.xml XML object
' 
' See also:
' 
'   <config.xml>
' 
sub loadCommon(Xml)
    dim Nodelist, Node
    set Nodelist = Xml.selectNodes("/configurations/common/child::*")
    for each Node in Nodelist
        Application(Node.nodeName) = Node.text
    next
    set Node = nothing
    set Nodelist = nothing
end sub

' Subroutine: loadCache
' 
' Configures cache life time and items to be cached.
' 
' Parameters:
' 
'   (xml object)Xml - config.xml XML object
' 
' See also:
' 
'   <config.xml>
' 
sub loadCache(Xml)
    dim Nodelist, Node, i : i = 0
    
    set Nodelist = Xml.selectNodes("/configurations/cache/lifetime")
    for each Node in Nodelist
        Application("Cache.lifetime") = Node.text
    next
    
    set Nodelist = Xml.selectNodes("/configurations/cache/item")
    dim saCache : redim saCache(Nodelist.length - 1, 2)
    for each Node in Nodelist
        saCache(i, 0) = Node.firstChild.text
        saCache(i, 1) = Node.lastChild.text
        saCache(i, 2) = dateAdd("yyyy", -1, now())
        i = i + 1
    next
    
    Application("Cache.items") = saCache
    
    set Node = nothing
    set Nodelist = nothing
end sub

' Subroutine: loadRoutes
' 
' Configure the application routes.
' 
' Parameters:
' 
'   (xml object)Xml - config.xml XML object
' 
' See also:
' 
'   <config.xml>
' 
public sub loadRoutes(Xml)
    dim Nodelist, Node, Actions, Action, i, j
    dim aRoutes, aActions
    
    set Nodelist = Xml.selectNodes("/configurations/routes/controller")
    redim aRoutes(Nodelist.length - 1)
    i = 0
    for each Node in Nodelist
        set Actions = Node.selectNodes("action")
        redim aActions(Actions.length - 1)
        j = 0
        for each Action in Actions
            aActions(j) = array( Action.getAttribute("in"), Action.getAttribute("out") )
            j = j + 1
        next
        set Actions = nothing
        aRoutes(i) = array(Node.getAttribute("in"), Node.getAttribute("out"), aActions)
        i = i + 1
    next
    
    Application("Routes") = aRoutes
    
    set Nodelist = nothing
end sub

%>
