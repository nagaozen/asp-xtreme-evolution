<%

' File: logger.asp
' 
' AXE(ASP Xtreme Evolution) log utility.
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



' Class: Logger
' 
' Logger is a component for general purpose logging. It supports multiple log 
' backends and filtering messages from being logged.
' 
' Dependencies:
' 
'   - List class (/lib/axe/classes/Utilities/list.asp)
'   - Template class (/lib/axe/classes/Utilities/template.asp)
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Sep 2010
' 
class Logger
    
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
    
    ' Property: Priorities
    ' 
    ' Scripting.Dictionary holding available log priorities. The built-in 
    ' priorities are the same available at BSD syslog <http://tools.ietf.org/html/rfc3164>.
    ' 
    ' Contains:
    ' 
    '     (Scripting.Dictionary) - Built-in priorities
    ' 
    ' Built-in values:
    ' 
    '     (EMERG)  - Emergency: system is unusable. Priority 0
    '     (ALERT)  - Alert: action must be taken immediately. Priority 1
    '     (CRIT)   - Critical: critical conditions. Priority 2
    '     (ERROR)  - Error: error conditions. Priority 3
    '     (WARN)   - Warning: warning conditions. Priority 4
    '     (NOTICE) - Notice: normal but significant condition. Priority 5
    '     (INFO)   - Informational: informational messages. Priority 6
    '     (DEBUG)  - Debug: debug messages. Priority 7
    ' 
    public Priorities
    
    ' Property: [_Adapters]
    ' 
    ' {private} In order to really write the messages somewhere, Log requires 
    ' adapters  implementing Logger_Interface specifying how and where to write 
    ' the data coming through this class.
    ' 
    ' Contains:
    ' 
    '     (Logger_Interface()) - media adapters
    ' 
    private [_Adapters]
    
    ' Property: [_Filters]
    ' 
    ' {private} Filters blocks a message from being written to the log.
    ' 
    ' Contains:
    ' 
    '     (Scripting.Dictionary) - filters
    ' 
    private [_Filters]
    
    ' Property: [_τ]
    ' 
    ' {private} Template instance to handle the string formating.
    ' 
    ' Contains:
    ' 
    '     (Template) - Object to handle the string formating.
    ' 
    private [_τ]
    
    ' Subroutine: [_ε]
    ' 
    ' {private} Checks for an adapter assignment.
    ' 
    private sub [_ε]
        if([_Adapters].count = 0) then
            Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Missing an Email_Interface adapter."
        end if
    end sub
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        set Priorities = Server.createObject("Scripting.Dictionary")
        Priorities.add "EMERG",  0
        Priorities.add "ALERT",  1
        Priorities.add "CRIT",   2
        Priorities.add "ERROR",  3
        Priorities.add "WARN",   4
        Priorities.add "NOTICE", 5
        Priorities.add "INFO",   6
        Priorities.add "DEBUG",  7
        
        set [_Adapters] = new List
        set [_Filters] = Server.createObject("Scripting.Dictionary")
        set [_τ] = new Template
    end sub
    
    private sub Class_terminate()
        set [_τ] = nothing
        set [_Filters] = nothing
        set [_Adapters] = nothing
        
        set Priorities = nothing
    end sub
    
    ' Function: [_compare]
    ' 
    ' {private} Compares two numbers by an order operator.
    ' 
    ' Parameters:
    ' 
    '     (number) - left side of comparison
    '     (number) - right side of comparison
    '     (string) - operator
    ' 
    ' Returns:
    ' 
    '     (boolean) - true, if the comparison is true; false, otherwise
    ' 
    private function [_compare](a, b, op)
        if( op = ">" or op = "gt" ) then
            [_compare] = ( a > b )
        elseif( op = ">=" or op = "ge" ) then
            [_compare] = ( a >= b )
        elseif( op = "=" or op = "==" or op = "eq" ) then
            [_compare] = ( a = b )
        elseif( op = "!=" or op = "<>" or op = "ne" ) then
            [_compare] = ( a <> b )
        elseif( op = "<=" or op = "le" ) then
            [_compare] = ( a <= b )
        elseif( op = "<" or op = "lt" ) then
            [_compare] = ( a < b )
        else
            Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Invalid operator."
        end if
    end function
    
    ' Function: [_isAcceptable]
    ' 
    ' {private} Check with Filters if a message can be logged.
    ' 
    ' Parameters:
    ' 
    '     (string) - message
    '     (string) - type
    ' 
    ' Returns:
    ' 
    '     (boolean) - true, if acceptable; false, otherwise
    ' 
    private function [_isAcceptable](message, tp)
        [_isAcceptable] = true
        dim filter, kind
        for each filter in [_Filters].items()
            if(not isArray(filter)) then
                Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Invalid filter."
            end if
            
            kind = lcase(typename(filter(0)))
            if( instr(kind, "regexp") > 0 ) then
                [_isAcceptable] = filter(0).test(message)
            elseif( kind = "integer" ) then
                [_isAcceptable] = [_compare](Priorities(tp), filter(0) , filter(1))
            elseif( kind = "string" ) then
                [_isAcceptable] = false
                dim i : for i = 0 to ubound(filter)
                    if( tp = filter(i) ) then [_isAcceptable] = true
                next
            end if
        next
    end function
    
    ' Subroutine: addAdapter
    ' 
    ' Pushes a new Logger_Interface adapter to Adapters list
    ' 
    ' Parameters:
    ' 
    '     (Logger_Interface) - Adapter
    ' 
    public sub addAdapter(Adapter)
        dim Node : set Node = new List_Node
        set Node.data = Adapter
        call [_Adapters].push(Node)
        set Node = nothing
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
    '     ([RegExp]) - Tests the message against a regular expression and accept it if true.
    '     ([Number, Operator]) - Accepts only the priorities in the range created by the number and operator.
    '     ([String, ...]) - Accepts only the types listed in the array
    ' 
    public sub addFilter(mixed)
        [_Filters].add [_Filters].count, mixed
    end sub
    
    ' Subroutine: write
    ' 
    ' Main writing routine
    ' 
    ' Parameters:
    ' 
    '     (string)   - message with/without placeholders
    '     (string[]) - replacements
    '     (string)   - log type
    ' 
    public sub write(message, arguments, tp) : call [_ε]
        set Node = [_Adapters].Head
        dim v : v = [_τ](message).[](arguments)
        if([_isAcceptable](v, tp)) then
            while(Node.hasNext)
                set Node = Node.pNext
                call Node.data.write(v, tp)
            wend
        end if
        set Node = nothing
    end sub
    
    ' Subroutine: emerg
    ' 
    ' Logs an emergency
    ' 
    ' Parameters:
    ' 
    '     (string)   - message with/without placeholders
    '     (string[]) - replacements
    ' 
    public sub emerg(message, arguments)
        call write(message, arguments, "EMERG")
    end sub
    
    ' Subroutine: alert
    ' 
    ' Logs an alert
    ' 
    ' Parameters:
    ' 
    '     (string)   - message with/without placeholders
    '     (string[]) - replacements
    ' 
    public sub alert(message, arguments)
        call write(message, arguments, "ALERT")
    end sub
    
    ' Subroutine: critical
    ' 
    ' Logs a critical
    ' 
    ' Parameters:
    ' 
    '     (string)   - message with/without placeholders
    '     (string[]) - replacements
    ' 
    public sub critical(message, arguments)
        call write(message, arguments, "CRIT")
    end sub
    
    ' Subroutine: error
    ' 
    ' Logs an error
    ' 
    ' Parameters:
    ' 
    '     (string)   - message with/without placeholders
    '     (string[]) - replacements
    ' 
    public sub error(message, arguments)
        call write(message, arguments, "ERROR")
    end sub
    
    ' Subroutine: warn
    ' 
    ' Logs a warning
    ' 
    ' Parameters:
    ' 
    '     (string)   - message with/without placeholders
    '     (string[]) - replacements
    ' 
    public sub warn(message, arguments)
        call write(message, arguments, "WARN")
    end sub
    
    ' Subroutine: notice
    ' 
    ' Logs a notice
    ' 
    ' Parameters:
    ' 
    '     (string)   - message with/without placeholders
    '     (string[]) - replacements
    ' 
    public sub notice(message, arguments)
        call write(message, arguments, "NOTICE")
    end sub
    
    ' Subroutine: info
    ' 
    ' Logs a info
    ' 
    ' Parameters:
    ' 
    '     (string)   - message with/without placeholders
    '     (string[]) - replacements
    ' 
    public sub info(message, arguments)
        call write(message, arguments, "INFO")
    end sub
    
    ' Subroutine: debug
    ' 
    ' Logs a debug
    ' 
    ' Parameters:
    ' 
    '     (string)   - message with/without placeholders
    '     (string[]) - replacements
    ' 
    public sub debug(message, arguments)
        call write(message, arguments, "DEBUG")
    end sub
    
end class

%>
