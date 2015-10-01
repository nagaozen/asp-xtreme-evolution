<%

' File: firephp.asp
' 
' AXE(ASP Xtreme Evolution) firephp adapter.
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



' Class: Logger_Adapter_FirePHP
' 
' This class leverages FirePHP protocol to send logs directly to Firebug.
' 
' Dependencies:
' 
'   - StringBuilder class (/lib/axe/classes/Utilities/stringbuilder.asp)
' 
' Requirements:
' 
'   - Firebug Firefox Extension which you can download from <https://addons.mozilla.org/en-US/firefox/addon/1843>
'   - FirePHP Firefox Extension which you can download from <https://addons.mozilla.org/en-US/firefox/addon/6149>
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao (nagaozen) <http://zend.lojcomm.com.br> @ Sep 2010
' 
class Logger_Adapter_FirePHP' implements Logger_Interface
    
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
    '   (string) - version
    ' 
    public classVersion
    
    ' Property: [_ι]
    ' 
    ' {private} Current log index
    ' 
    ' Contains:
    ' 
    '     (type) - short_description
    ' 
    private [_ι]
    
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
        classVersion = "1.0.0.1"
        
        set Interface = new Logger_Interface
        set Interface.Implementation = Me
        if(not Interface.check) then
            Err.raise 17, "Evolved AXE runtime error", strsubstitute( _
                "Can't perform requested operation. '{0}' is a bad interface implementation of '{1}'", _
                array(classType, typename(Interface)) _
            )
        end if
        
        [_ι] = 1
        encoding = "UTF-8"
        
        Response.addHeader "X-Wf-Protocol-1",    "http://meta.wildfirehq.org/Protocol/JsonStream/0.2"
        Response.addHeader "X-Wf-1-Plugin-1",    "http://meta.firephp.org/Wildfire/Plugin/FirePHP/Library-FirePHPCore/0.3"
        Response.addHeader "X-Wf-1-Structure-1", "http://meta.firephp.org/Wildfire/Structure/FirePHP/FirebugConsole/0.1"
    end sub
    
    private sub Class_terminate()
        set Interface.Implementation = nothing
        set Interface = nothing
    end sub
    
    ' Function: [_μ]
    ' 
    ' {private} Maps Logger types to FirePHP types.
    ' 
    ' Parameters:
    ' 
    '     (string) - Logger type
    ' 
    ' Returns:
    ' 
    '     (string) - FirePHP type
    ' 
    private function [_μ](tp)
        if( _
            tp = "EMERG" or _
            tp = "ALERT" or _
            tp = "CRIT" or _
            tp = "ERROR" _
        ) then
            [_μ] = "ERROR"
        elseif( _
            tp = "WARN" _
        ) then
            [_μ] = "WARN"
        elseif( _
            tp = "NOTICE" or _
            tp = "INFO" _
        ) then
            [_μ] = "INFO"
        elseif( _
            tp = "DEBUG" _
        ) then
            [_μ] = "LOG"
        else
            [_μ] = tp
        end if
    end function
    
    ' Function: [_τ]
    ' 
    ' {private} Handle JSON string backslash escapes and maps Unicode characters to JSON \uxxxx.
    ' 
    ' Parameters:
    ' 
    '     (char) - Unicode character
    ' 
    ' Returns:
    ' 
    '     (string) - JSON equivalent
    ' 
    private function [_τ](u)
        select case u
            case """":
                [_τ] = "\"""
            
            case "\":
                [_τ] = "\\"
            
            case "/":
                [_τ] = "\/"
            
            case vbFormFeed:
                [_τ] = "\f"
            
            case vbLf:
                [_τ] = "\n"
            
            case vbCr:
                [_τ] = "\r"
            
            case vbNewline:
                [_τ] = "\r\n"
            
            case vbTab:
                [_τ] = "\t"
            
            case else:
                if( ascw(u) > 255 ) then
                    [_τ] = "\u" & right("0000" & dec2hex( ascw( u ) ), 4)
                else
                    [_τ] = u
                end if
        end select
    end function
    
    ' Function: [_escaped]
    ' 
    ' {private} Converts an UTF-8 string into a JSON unicode escaped string.
    ' 
    ' Parameters:
    ' 
    '     (string) - content
    ' 
    ' Returns:
    ' 
    '     (string) - JSON equivalent
    ' 
    private function [_escaped](u)
        dim i, Sb
        set Sb = new StringBuilder
        for i = 1 to len(u)
            Sb.append( [_τ]( mid( u, i, 1 ) ) )
        next
        [_escaped] = Sb.toString()
        set Sb = nothing
    end function
    
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
    public sub write(byval message, tp)
        if(not Interface.[isAcceptable](message, tp)) then exit sub
        
        message = strsubstitute("[{""Type"":""{0}""}, ""{1}""]", array([_μ](tp), [_escaped]( message )))
        
        dim size, Stream : set Stream = Server.createObject("ADODB.Stream")
        Stream.charset = encoding
        Stream.type = adTypeText
        Stream.mode = adModeReadWrite
        Stream.open()
        
        Stream.writeText(message)
        
        Stream.position = 0
        size = Stream.size - 3
        message = Stream.readText()
        
        Stream.close()
        set Stream = nothing
        Response.addHeader strsubstitute("X-Wf-1-1-1-{0}", array([_ι])), strsubstitute("{0}|{1}|", array(size, message))
        
        [_ι] = [_ι] + 1
    end sub
    
end class

%>
