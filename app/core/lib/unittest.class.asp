<%

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
    
    private result
    
    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "1.0.0"
        
        container = ""
        set Properties = Server.createObject("Scripting.Dictionary")
    end sub
    
    private sub Class_terminate()
        set Properties = nothing
    end sub
    
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
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Tester : set Tester = new UnitTest
    ' 
    ' Response.write Tester.assert("helloWorld", null, "Hello World") & vbNewLine ' prints true
    ' Response.write Tester.assert("helloWorld", null, "Hi World") & vbNewLine ' prints false
    ' 
    ' Tester.container = "Math"
    ' 
    ' Response.write "Testing Math.add" & vbNewLine
    ' Response.write "----------------" & vbNewLine
    ' Response.write Tester.assert("add", array(2,2), 4) & vbNewLine ' prints true
    ' Response.write Tester.assert("add", array(2,0), 5) & vbNewLine ' prints false
    ' 
    ' set Tester = nothing
    ' 
    ' Response.write vbNewLine
    ' 
    ' (end code)
    ' 
    public function assert(fn, argv, expected)
        on error resume next
        
        dim Builder : set Builder = new StringBuilder
        dim item
        for each item in Properties
            Builder.append("Obj." & item & " = " & Properties.item(item) & vbNewLine)
        next
        dim initialize : initialize = Builder.toString()
        set Builder = nothing
        
        dim arguments
        if( isnull(argv) ) then
            arguments = ""
        else
            dim i
            for i = 0 to ubound(argv)
                select case lcase(typename(argv(i)))
                    case "string"
                        argv(i) = """" & argv(i) & """"
                end select
            next
            arguments = join(argv, ", ")
        end if
        
        dim vbscript
        if( container = "" ) then
            vbscript = "result = " & fn & "(" & arguments & ")"
        else
            vbscript = join(array( _
                "dim Obj : set Obj = new " & container, _
                initialize, _
                "result = Obj." & fn & "(" & arguments & ")", _
                "set Obj = nothing" _
            ), vbNewLine)
        end if
        
        execute(vbscript)
        
        if( result = expected ) then
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
    ' dim Tester : set Tester = new UnitTest
    ' 
    ' Response.write "Testing functions" & vbNewLine
    ' Response.write "-----------------" & vbNewLine
    ' Response.write Tester.test("helloWorld", null, "Hello World") & vbNewLine
    ' Response.write Tester.test("dummy", array("someValue"), "anotherValue") & vbNewLine
    ' 
    ' Response.write vbNewLine
    ' 
    ' Tester.container = "Math"
    ' 
    ' Response.write "Testing Math.add" & vbNewLine
    ' Response.write "----------------" & vbNewLine
    ' Response.write Tester.test("add", array(2,2), 4) & vbNewLine
    ' Response.write Tester.test("add", array(2,0), 5) & vbNewLine
    ' Response.write Tester.test("add", array(2008,1), 2009) & vbNewLine
    ' Response.write Tester.test("add", array(1,2008), 2009) & vbNewLine
    ' 
    ' set Tester = nothing
    ' 
    ' Response.write vbNewLine
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
                select case lcase(typename(a(i)))
                    case "string"
                        a(i) = """" & a(i) & """"
                end select
            next
            arguments = join(a, ", ")
        end if
        
        if(assert(fn, argv, expected)) then
            test = "Success: " & container & iif((container = ""), "", ".") & fn & "(" & arguments & ") == (" & lcase(typename(expected)) & ")" & expected
        else
            test = "Failure: " & container & iif((container = ""), "", ".") & fn & "(" & arguments & ") != (" & lcase(typename(expected)) & ")" & expected & ". Method returns: (" & lcase(typename(result)) & ")" & result
        end if
    end function
    
end class

%>
