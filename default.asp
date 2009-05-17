<%

'+-----------------------------------------------------------------------------+
'|This file is part of ASP Xtreme Evolution.                                   |
'|Copyright © 2007, 2008 Fabio Zendhi Nagao                                    |
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
'|along with ASP Xtreme Evolution. If not, see <http://www.gnu.org/licenses/>. |
'+-----------------------------------------------------------------------------+

' File: default.asp
' 
' The main document, this is a gateway. All requests are handled here and only
' here. Used with IIRF, it provides a nice MVC architecture with URL-Rewriting
' logic.
' 
' About:
' 
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2007

%>
<!--#include virtual="/app/core/base.asp"-->
<!--#include virtual="/app/core/lib/kernel.class.asp"-->
<!--#include virtual="/app/core/singletons.initialize.asp"-->
<!--#include virtual="/app/core/sessions.initialize.asp"-->
<%

' Set the controller and action sessions with the values received from the URI
Core.initialize

' Run the controller
Core.process

' Run the view
Core.dispatch

' Print a comment with the execution time for optimizing purposes
if( Response.contentType = "text/html" ) then
    Response.write "<!--// This page required " & timer - Session("Request.time") & " seconds of execution time //-->"
end if

%>
<!--#include virtual="/app/core/sessions.finalize.asp"-->
<!--#include virtual="/app/core/singletons.finalize.asp"-->