<!--#include virtual="/app/core/base.asp"-->
<!--#include virtual="/app/core/lib/kernel.class.asp"-->
<!--#include virtual="/app/core/singletons.initialize.asp"-->
<%

'+-----------------------------------------------------------------------------+
'|This file is part of ASP Xtreme Evolution.                                   |
'|Copyright (C) 2007, 2008 Fabio Zendhi Nagao                                  |
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

' ${1:Template}Controller.asp
' 
' ${2:ASP Xtreme Evolution Template Controller.}
' 
' Written by ${3:author} <http://${4:url}> @ {5:month} {6:year}
' 
class ${1:Template}Controller
    
    public sub defaultAction()
        $0
    end sub
    
end class



dim Controller : set Controller = new ${1:Template}Controller
select case Session("action")
    
    case "defaultAction"
        call Controller.defaultAction()
    
    case else
        Core.addError("Action '" & Session("action") & "' is not available at controller '" & Session("controller") & "'")
    
end select
set Controller = nothing

%>
<!--#include virtual="/app/core/singletons.finalize.asp"-->
