<%

' File: text.asp
' 
' AXE(ASP Xtreme Evolution) text file adapter.
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



' Class: Logger_Adapter_Text
' 
' This class enables Logger to write messages into text files.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao (nagaozen) <http://zend.lojcomm.com.br> @ Sep 2010
' 
class Logger_Adapter_Text' implements Logger_Interface
    
    ' --[ Interface ]-----------------------------------------------------------
    public Interface
    
    ' --[ Adapter definition ]--------------------------------------------------
    
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
    
    ' Property: path
    ' 
    ' File System physical path
    ' 
    ' Contains:
    ' 
    '     (string) - file system physical path
    ' 
    public path
    
    ' Property: encoding
    ' 
    ' Text encoding
    ' 
    ' Contains:
    ' 
    '     (Stream.charset) - text encoding. Defaults to UTF-8
    ' 
    public encoding
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        set Interface = new Logger_Interface
        set Interface.Implementation = Me
        if(not Interface.check) then
            Err.raise 17, "Evolved AXE runtime error", strsubstitute( _
                "Can't perform requested operation. '{0}' is a bad interface implementation of '{1}'", _
                array(classType, typename(Interface)) _
            )
        end if
        
        encoding = "UTF-8"
    end sub
    
    private sub Class_terminate()
        set Interface.Implementation = nothing
        set Interface = nothing
    end sub
    
    ' Subroutine: addFilter
    ' 
    ' Add a filter that will be applied before writing the message in this adapter.
    ' 
    ' Parameters:
    ' 
    '     (mixed[]) - array defining the filter
    '
    ' Values:
    ' 
    '     ([RegExp]) - Skips the log operation if a test for the regular expression returns false
    '     ([Number, Operator]) - Skips the log operation based on the priority defined by Number and Operator.
    '     ([String, ...]) - Skips the log operation for the types listed in this array
    ' 
    public sub addFilter(mixed)
        call Interface.addFilter(mixed)
    end sub
    
    ' Subroutine: write
    ' 
    ' Adapter writing routine
    ' 
    ' Parameters:
    ' 
    '     (string)   - message with/without placeholders
    '     (string)   - log type
    ' 
    public sub write(message, tp)
        if(varType(path) = vbEmpty) then
            Err.raise 5, "Evolved AXE runtime error", strsubstitute( _
                "Invalid procedure call or argument. '{0}' property 'path' isn't defined", _
                array(classType) _
            )
        end if
        
        if(not Interface.[isAcceptable](message, tp)) then exit sub
        
        dim Fso : set Fso = Server.createObject("Scripting.FileSystemObject")
        dim Stream : set Stream = Server.createObject("ADODB.Stream")
        Stream.charset = encoding
        Stream.type = adTypeText
        Stream.mode = adModeReadWrite
        Stream.open()
        
        if(Fso.fileExists(path)) then
            Stream.loadFromFile(path)
            Stream.position = 0
            Stream.readText()
        end if
        
        call Stream.writeText(strsubstitute( _
            "({0}){3}{1}{3}{2}", _
            array(tp,message, now, vbTab) _
        ), adWriteLine)
        call Stream.saveToFile(path, adSaveCreateOverWrite)
        
        Stream.close()
        set Stream = nothing
        
        set Fso = nothing
    end sub
    
end class

%>
