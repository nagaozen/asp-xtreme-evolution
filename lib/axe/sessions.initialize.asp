<%

' File: sessions.initialize.asp
' 
' This snippet is one of the key features of this MVC approach. Since ASP
' does not have runtime includes (<!--#include-->'s are first included and then
' compiled with the entire ASP page), Request Object is read-only and
' Server.Execute runs the requested document in an isolated process we have no
' other built-in way than using the so called evil Session Scope to mimic a
' Request Scope where the framework can share data between controllers, models
' and views. This won't become a problem since we are destroying them after each
' end of request. For Session tracking, we suggest using Cookies and Database
' which will also increase your application scalability making it web farms
' ready.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2007-2011 Fabio Zendhi Nagao
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
dim saArgv(0)
Session("controller") = "defaultController"
Session("action") = "defaultAction"
Session("view") = "defaultView"
Session("argv") = saArgv
set Session("this") = Server.createObject("Scripting.Dictionary")

%>
