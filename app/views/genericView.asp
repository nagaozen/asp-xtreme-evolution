<!--#include virtual="/app/core/base.asp"-->
<!--#include virtual="/app/core/lib/kernel.class.asp"-->
<!--#include virtual="/app/core/singletons.initialize.asp"-->
<!--#include virtual="/app/core/shuttle.unload.initialize.asp"-->
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

' genericView.asp
' 
' Use this view to dispatch the content assigned to "View.content". Useful view 
' for feeding XML and JSON directly from their models.
' 
' Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ Feb 2009
' 

Response.write( Session("this").item("View.content") )

%>
<!--#include virtual="/app/core/shuttle.unload.finalize.asp"-->
<!--#include virtual="/app/core/singletons.finalize.asp"-->
