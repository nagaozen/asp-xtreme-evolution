<%

' File: stringbuilder.asp
' 
' AXE(ASP Xtreme Evolution) string builder utility.
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



' Class: StringBuilder
' 
' While using Response.write is far away faster than concatenating strings,
' sometimes you need to store the entire output before printing it. That's when
' a special designed string builder class is required.
' 
' About:
' 
'   - Written by Markus Paeschke <http://www.mpaeschke.de> @ November 2012
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

    private stream, counter

    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "3.0.0"

        Set stream = CreateObject("Scripting.Dictionary")
    end sub

    private sub Class_terminate()
        Set stream = Nothing
    end sub

    ' Function: reset
    ' 
    ' Clear the string builder data.
    ' 
    public function reset()
        Call stream.RemoveAll()
        counter = 0
    end function

    ' Subroutine: append
    ' 
    ' Add the incoming data to the buffer if the data is not empty.
    ' 
    ' Parameters:
    ' 
    '   (string) - String fragments.
    ' 
    public sub append(data)
        If Not isNull(data) And data <> "" Then
            Call stream.Add(counter, cStr(data))
            counter = counter +1
        End If
    end sub

    ' Function: toString
    ' 
    ' Join all items in the dictionary to a single string.
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
        If counter.Count > 0 Then
            toString = Join(stream.Items, "")
        End If
    end function

end class

%>
