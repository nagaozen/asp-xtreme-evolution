<%@

language = "VBScript"
codepage = 65001
lcid     = 1033

%><%

' File: default.asp
' 
' The main document, this is a gateway. All requests are handled here and only
' here. Used with IIRF, it provides a nice MVC architecture with URL-Rewriting
' logic.
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
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2007

%>
<!--#include virtual="/lib/axe/singletons.initialize.asp"-->
<!--#include virtual="/lib/axe/sessions.initialize.asp"-->
<%

' Set the controller and action sessions with the values received from the URI 
' (initialize), run the controller (process) and run the view (dispatch)
Core.initialize().process().dispatch()

' Print a comment with the execution time for optimizing purposes
if( lcase(Response.contentType) = "text/html" ) then
    Response.write "<!--// This page required " & timer - Session("Request.time") & " seconds of execution time //-->"
end if

%>
<!--#include virtual="/lib/axe/sessions.finalize.asp"-->
<!--#include virtual="/lib/axe/singletons.finalize.asp"-->