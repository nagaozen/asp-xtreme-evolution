<!--#include virtual="/lib/axe/base.asp"-->
<!--#include virtual="/lib/axe/classes/kernel.class.asp"-->
<!--#include virtual="/lib/axe/singletons.initialize.asp"-->
<!--#include virtual="/lib/axe/shuttle.unload.initialize.asp"-->
<%

' File: genericView.asp
' 
' Use this view to dispatch the content assigned to "View.content". Useful view 
' for feeding XML and JSON directly from their models.
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
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2009
' 

Response.write( Session("this").item("View.content") )

%>
<!--#include virtual="/lib/axe/shuttle.unload.finalize.asp"-->
<!--#include virtual="/lib/axe/singletons.finalize.asp"-->
