<%

' File: shuttle.unload.initialize.asp
' 
' The shuttle is another of this framework key solutions. It's loaded with some
' relevant user request scope data to be used in the server request scope. Use
' this snippet to initialize unloading the shuttle in all your views.
' 
' See also:
' 
'   <Framework.loadShuttle>, <Framework.computeView>
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
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ January 2007
' 
set Session("this") = Server.createObject("Scripting.Dictionary")
call Core.unloadShuttle()

%>
