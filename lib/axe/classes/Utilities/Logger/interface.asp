<%

' File: interface.asp
' 
' AXE(ASP Xtreme Evolution) Logger interface definition.
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



' Class: Logger_Interface
' 
' Defines the common specifications required to implement a working adapter of 
' Logger class.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Sep 2010
' 
class Logger_Interface' extends Interface
    
    ' --[ Inheritance ]---------------------------------------------------------
    public Parent
    
    public property set Implementation(I)
        set Parent.Implementation = I
    end property
    
    public property get Implementation
        set Implementation = Parent.Implementation
    end property
    
    public property get requireds
        requireds = Parent.requireds
    end property
    
    public function check()
        check = Parent.check()
    end function
    
    ' --[ Interface definition ]------------------------------------------------
    
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
    
    ' Property: [_Filters]
    ' 
    ' {private} Filters blocks a message from being written to the log.
    ' 
    ' Contains:
    ' 
    '     (Scripting.Dictionary) - filters
    ' 
    private [_Filters]
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        set Parent = new Interface
        Parent.requireds = array("addFilter", "write")
        
        set Priorities = Server.createObject("Scripting.Dictionary")
        Priorities.add "EMERG",  0
        Priorities.add "ALERT",  1
        Priorities.add "CRIT",   2
        Priorities.add "ERROR",  3
        Priorities.add "WARN",   4
        Priorities.add "NOTICE", 5
        Priorities.add "INFO",   6
        Priorities.add "DEBUG",  7
        
        set [_Filters] = Server.createObject("Scripting.Dictionary")
    end sub
    
    private sub Class_terminate()
        set [_Filters] = nothing
        set Priorities = nothing
        
        set Parent = nothing
    end sub
    
    ' Subroutine: [_ρ]
    ' 
    ' {private} Accepts only protected calls
    ' 
    private sub [_ρ]()
        if( isEmpty(Parent.Implementation) ) then Err.raise 70, "Evolved AXE runtime error", "Permission denied. Method is protected."
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
    
    ' Function: [isAcceptable]
    ' 
    ' {protected} Check with Filters if a message can be logged.
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
    public function [isAcceptable](message, tp) : call [_ρ]
        [isAcceptable] = true
        dim filter, kind
        for each filter in [_Filters].items()
            if(not isArray(filter)) then
                Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Invalid filter."
            end if
            
            kind = lcase(typename(filter(0)))
            if( instr(kind, "regexp") > 0 ) then
                [isAcceptable] = filter(0).test(message)
            elseif( kind = "integer" ) then
                [isAcceptable] = [_compare](Priorities(tp), filter(0) , filter(1))
            elseif( kind = "string" ) then
                [isAcceptable] = false
                dim i : for i = 0 to ubound(filter)
                    if( tp = filter(i) ) then [isAcceptable] = true
                next
            end if
        next
    end function
    
    ' Subroutine: [addFilter]
    ' 
    ' {protected} Add a filter that will be applied before writing the message in this adapter.
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
    public sub [addFilter](mixed) : call [_ρ]
        [_Filters].add [_Filters].count, mixed
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
    end sub
    
end class

%>
