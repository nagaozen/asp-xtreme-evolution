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

' ${1:Template}Model.asp
' 
' ${2:ASP Xtreme Evolution Template Model.}
' 
' Written by ${3:author} <http://${4:url}> @ {5:month} {6:year}
' 
class ${1:Template}Model
    
    $0
    
    ' Function: toXML
    ' 
    ' Returns a XML representation of this model.
    ' 
    ' Parameters:
    ' 
    '     (string) - root node of the recursion
    ' 
    ' Returns:
    ' 
    '     (string) - XML representation
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Model : set Model = new ${1:Template}Model
    ' Response.write(Model.toXML())
    ' set Model = nothing
    ' 
    ' (end code)
    ' 
    public function toXML()
        
    end function
    
    ' Function: toJSON
    ' 
    ' Returns a JSON representation of this model.
    ' 
    ' Parameters:
    ' 
    '     (string) - root node of the recursion
    ' 
    ' Returns:
    ' 
    '     (string) - JSON representation
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Model : set Model = new ${1:Template}Model
    ' Response.write(Model.toJSON())
    ' set Model = nothing
    ' 
    ' (end code)
    ' 
    public function toJSON()
        
    end function
    
    ' Function: toHTML
    ' 
    ' Returns a HTML representation of this model.
    ' 
    ' Parameters:
    ' 
    '     (string) - root node of the recursion
    ' 
    ' Returns:
    ' 
    '     (string) - HTML representation
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Model : set Model = new ${1:Template}Model
    ' Response.write(Model.toHTML())
    ' set Model = nothing
    ' 
    ' (end code)
    ' 
    public function toHTML()
        
    end function
    
end class

%>
