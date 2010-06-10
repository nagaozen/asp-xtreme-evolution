<%

' File: stringbuilder.asp
' 
' AXE(ASP Xtreme Evolution) string builder utility.
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



' Class: StringBuilder
' 
' While using Response.write is far away faster than concatenating strings,
' sometimes you need to store the entire output before printing it. That's when
' a special designed string builder class is required.
' 
' About:
' 
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ January 2009
' 
class StringBuilder

    ' Property: classType
    ' 
    ' Class type.
    ' 
    ' Contains:
    ' 
    '   (string) - type
    ' 
    public classType

    ' Property: classVersion
    ' 
    ' Class version.
    ' 
    ' Contains:
    ' 
    '   (float) - version
    ' 
    public classVersion

    private Stream

    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "2.0.0"
        
        set Stream = Server.createObject("ADODB.Stream")
        Stream.type = adTypeText
        Stream.mode = adModeReadWrite
        Stream.open()
    end sub

    private sub Class_terminate()
        Stream.close()
        set Stream = nothing
    end sub

    ' Function: reset
    ' 
    ' Clear the string builder data.
    ' 
    public function reset()
        Stream.position = 0
        Stream.setEOS()
    end function

    ' Subroutine: append
    ' 
    ' Add the incoming data to the buffer.
    ' 
    ' Parameters:
    ' 
    '   (string) - String fragments.
    ' 
    public sub append(data)
        Stream.writeText(data)
    end sub

    ' Function: toString
    ' 
    ' Reads the entire buffer and return it.
    ' 
    ' Returns:
    ' 
    '   (string) - concatenated string
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Builder : set Builder = new StringBuilder
    ' Builder.append "First line" & vbNewLine
    ' Builder.append "Second line" & vbNewLine
    ' Response.write( Builder.toString() )
    ' set Builder = nothing
    ' 
    ' (end)
    ' 
    public function toString()
        Stream.position = 0
        toString = Stream.readText()
    end function

end class

%>
