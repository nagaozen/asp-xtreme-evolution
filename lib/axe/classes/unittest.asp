<script language="VBScript" runat="server">

' File: unittest.asp
' 
' AXE(ASP Xtreme Evolution) Unit Test factory.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2007-2011 Fabio Zendhi Nagao
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



' Class: UnitTest
' 
' Unit tests are pretty popular these days. It's a method of testing that
' verifies the individual units of source code are working properly. A unit is
' the smallest tastable part of an application. In our case, these units are
' functions or methods.
' 
' Dependencies:
' 
'   - Class StringBuilder
' 
' About:
' 
'   - Written by Fabio Zendhi Nagao  @ October 2008
' 
class UnitTest
    
    ' Property: classType
    ' 
    ' Class type.
    ' 
    ' Contains:
    ' 
    '     (string) - type
    ' 
    public classType
    
    ' Property: classVersion
    ' 
    ' Class version.
    ' 
    ' Contains:
    ' 
    '     (float) - version
    ' 
    public classVersion
    
    ' Property: container
    ' 
    ' Holds the unit test container. The class name if you are testing methods
    ' or "" if you are testing functions.
    ' 
    ' Contains:
    ' 
    '     (string) - container
    ' 
    public container
    
    ' Property: Properties
    ' 
    ' This object stores the initialization properties to be called in the test.
    ' 
    ' Contains:
    ' 
    '     (scripting dictionary) - properties
    ' 
    public Properties
    
    ' Function: build
    ' 
    ' Generate the unit test code.
    ' 
    ' Parameters:
    ' 
    '     (string)    - function name
    '     (variant[]) - function argument values
    ' 
    ' Returns:
    ' 
    '     (string) - generated code
    ' 
    public function build(fn, argv)
        dim Builder, item, initialize, _
            arguments, i, _
            vbscript
        
        set Builder = new StringBuilder
        for each item in Properties
            Builder.append("Obj." & item & " = " & [_formatArgument](Properties.item(item)) & vbNewLine)
        next
        initialize = Builder.toString()
        set Builder = nothing
        
        if( isnull(argv) ) then
            arguments = ""
        else
            for i = 0 to ubound(argv)
                argv(i) = [_formatArgument](argv(i))
            next
            arguments = join(argv, ", ")
        end if
        
        if( container = "" ) then
            vbscript = "[_result] = " & fn & "(" & arguments & ")"
        else
            vbscript = join(array( _
                "dim Obj : set Obj = new " & container, _
                "", _
                initialize, _
                "[_result] = Obj." & fn & "(" & arguments & ")", _
                "", _
                "set Obj = nothing" _
            ), vbNewLine)
        end if
        
        build = vbscript
    end function
    
    ' Function: assert
    ' 
    ' Tests if the coded function or method returns an equal value to the
    ' expected value.
    ' 
    ' Parameters:
    ' 
    '     (string)  - function or method
    '     (array)   - arguments
    '     (variant) - expected output
    ' 
    ' Returns:
    ' 
    '     (boolean) - true if function or method returns expected, false otherwise
    ' 
    public function assert(fn, argv, expected)
        on error resume next
        
        [_result] = empty
        execute(build(fn, argv))
        
        if( [_result] = expected ) then
            assert = true
        else
            assert = false
        end if
        
        on error goto 0
    end function
    
    ' Function: test
    ' 
    ' Returns a friendly message of the test.
    ' 
    ' Parameters:
    ' 
    '     (string)  - function or method
    '     (array)   - arguments
    '     (variant) - expected output
    ' 
    ' Returns:
    ' 
    '     (string) - message
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' ' --[ Definitions ]-------------------------------------------------------------
    ' 
    ' ' NOTE: Usually you will be including the definition from somewhere else.
    ' 
    ' function helloWorld()
    '     helloWorld = "Hello World"
    ' end function
    ' 
    ' function moron(arg)
    '     moron = arg
    ' end function
    ' 
    ' class BasicMath
    '     
    '     public function add(a, b)
    '         add = a + b
    '     end function
    '     
    '     public function subtract(a, b)
    '         subtract = a - b
    '     end function
    '     
    '     public function multiply(a, b)
    '         multiply = a * b
    '     end function
    '     
    '     public function divide(p, q)
    '         divide = p / q
    '     end function
    '     
    ' end class
    ' 
    ' ' --[ Testing ]-----------------------------------------------------------------
    ' 
    ' dim Tester : set Tester = new UnitTest
    ' 
    ' Response.write "Testing functions" & vbNewLine
    ' Response.write "-----------------" & vbNewLine
    ' Response.write Tester.test("helloWorld", null, "Hello World") & vbNewLine
    ' Response.write Tester.test("moron", array("someValue"), "anotherValue") & vbNewLine
    ' 
    ' Response.write vbNewLine
    ' 
    ' Tester.container = "BasicMath"
    ' 
    ' Response.write "Testing BasicMath.add" & vbNewLine
    ' Response.write "---------------------" & vbNewLine
    ' Response.write Tester.test("add", array(2,2), 4) & vbNewLine
    ' Response.write Tester.test("add", array(2,0), 5) & vbNewLine
    ' Response.write Tester.test("add", array(2008,1), 2009) & vbNewLine
    ' Response.write Tester.test("add", array(1,2008), 2009) & vbNewLine
    ' 
    ' set Tester = nothing
    ' 
    ' (end code)
    ' 
    public function test(fn, argv, expected)
        dim arguments
        if( isnull(argv) ) then
            arguments = ""
        else
            dim a : a = argv
            dim i : for i = 0 to ubound(a)
                a(i) = [_formatArgument](a(i))
            next
            arguments = join(a, ", ")
        end if
        
        if(assert(fn, argv, expected)) then
            test = "Success: " & container & iif((container = ""), "", ".") & fn & "(" & arguments & ") == (" & lcase(typename(expected)) & ")" & expected
        else
            test = "Failure: " & container & iif((container = ""), "", ".") & fn & "(" & arguments & ") != (" & lcase(typename(expected)) & ")" & expected & ". Method returns: (" & lcase(typename([_result])) & ")" & [_result]
        end if
    end function
    
    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "1.0.0.0"
        
        container = ""
        set Properties = Server.createObject("Scripting.Dictionary")
    end sub
    
    private sub Class_terminate()
        set Properties = nothing
    end sub
    
    private [_result]
    
    ' Function: [_formatArgument]
    ' 
    ' {private} Format the input argument to be reused in a code generator.
    ' 
    ' Parameters:
    ' 
    '     (variant) - argument value
    ' 
    ' Returns:
    ' 
    '     (string) - argument to be used in the dynamic command.
    ' 
    private function [_formatArgument](arg)
        select case lcase(typename(arg))
            case "string"
                [_formatArgument] = """" & arg & """"
            case else
                [_formatArgument] = arg
        end select
    end function
    
end class

</script>
