<%

' File: mvc.bootstrapper.asp
' 
' The bootstrapper is responsible for the initialization of an application built
' using the ASP Xtreme Evolution MVC architecture with /Controller/action.
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
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ Oct 2011
' 
function [_ßootstrap_naming_filter](value)
    dim Re : set Re = new RegExp
    Re.pattern = "^\S+$"
    Re.ignoreCase = true
    Re.global = false
    
    if( Re.test(value) ) then
        [_ßootstrap_naming_filter] = value
    else
        Err.raise 5, "Evolved AXE runtime error", strsubstitute( _
            "Invalid procedure call or argument. Controller or action names doesn't apply for this bootstrapper naming policy.", _
            array(Session("controller"), vpController) _
        )
    end if
    
    set Re = nothing
end function

executeGlobal strsubstitute( _
    join(array( _
        "dim [ßootstrap_Controller] : set [ßootstrap_Controller] = new [{1}Controller]", _
        "call [ßootstrap_Controller].[{0}]()", _
        "set [ßootstrap_Controller] = nothing" _
    ), vbNewline), _
    array( _
        [_ßootstrap_naming_filter]( Session("action") ), _
        [_ßootstrap_naming_filter]( Session("controller") ) _
    ) _
)

%>
