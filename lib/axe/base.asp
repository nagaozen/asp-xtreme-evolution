<%

' File: base.asp
' 
' All asp pages should include this file.
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
' 
' About:
' 
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2007
' 
option explicit 'Forces the developer to declare all variables that will be used
Response.buffer = true 'Buffer the whole page before sending it to avoid a lot of overhead
Response.charset = Session("Response.charset") 'Sets charset to the value configured in the global.asa

' Function: axeInfo
' 
' Prints information about the ASP Xtreme Evolution version and Server scripting engine.
' 
' Returns:
' 
'     (string) - Framework information
' 
function axeInfo()
    axeInfo = Application("name") & " v" & Application("version") & " running over " & ScriptEngine & " v" & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion
end function

' Function: iif
' 
' Mimics the C ternary operator ?:
' 
' Parameters:
' 
'     (expression) - The expression that is to be evaluated
'     (variant)    - Defines what the iif function returns if the evaluation of expression returns true
'     (variant)    - Defines what the iif function returns if the evaluation of expression returns false
' 
' Returns:
' 
'     (variant) - truepart or falsepart
' 
' Example:
' 
' (start code)
' 
' Response.write iif((5 < 10), "Yes, it is", "No, it isn't") ' prints "Yes, it is"
' Response.write iif((2 + 2 = 5), "Correct", "Wrong") ' prints "Wrong"
' 
' (end code)
' 
function iif(expr, truepart, falsepart)
    if(expr) then
        iif = truepart
    else
        iif = falsepart
    end if
end function

' Function: strsubstitute
' 
' Mimics the string placeholders.
' 
' Parameters:
' 
'     (string)   - The string template with placeholders
'     (string[]) - Replacement / Replacements array
' 
' Returns:
' 
'     (string) - parsed string
' 
' Example:
' 
' (start code)
' 
' Response.write( strsubstitute("{0} {1}", array("Hello", "World")) )
' 
' (end code)
' 
function strsubstitute(template, replacements)
    if( not isArray(replacements) ) then replacements = array(replacements)
    
    dim Re : set Re = new RegExp
    Re.global = true
    Re.ignoreCase = true
    Re.pattern = "{\d+}"
    
    dim Matches, Match, index : set Matches = Re.execute(template)
    for each Match in Matches
        index = replace(replace(Match, "{", ""), "}", "")
        template = replace(template, Match.value, replacements(clng(index)))
    next
    set Matches = nothing
    
    set Re = nothing
    
    strsubstitute = template
end function

' Function: sanitize
' 
' A generic sanitize function.
' 
' Parameters:
' 
'     (string)   - The value to be sanitized
'     (string[]) - Elements to be sanitized array
'     (string[]) - Elements replacements array
' 
' Returns:
' 
'     (string) - Sanitized string
' 
' Example:
' 
' (start code)
' 
' dim aWithAccent : aWithAccent = "àáâãäå"
' Response.write( sanitize( aWithAccent, array("à", "á", "â", "ã", "ä", "å"), array("a", "a", "a", "a", "a", "a") ) )
' 
' (end code)
' 
public function sanitize(value, placeholders, replacements)
    if( ( vartype( placeholders ) < 8192 ) or ( vartype( replacements ) < 8192 ) ) then
        sanitize = "2nd and 3rd arguments should be arrays"
        exit function
    end if
    if(ubound(placeholders) <> ubound(replacements)) then
        sanitize = "2nd and 3rd arguments doesn't have the same length"
        exit function
    end if
    
    dim i
    sanitize = value
    for i = 0 to ubound(placeholders)
        sanitize = replace(sanitize, placeholders(i), replacements(i))
    next
end function

' Function: isOdd
' 
' Checks if a natural is odd or not.
' 
' Parameters:
' 
'     (int) - Integer
' 
' Returns:
' 
'     (boolean) - true, if yes; false, otherwise
' 
' Example:
' 
' (start code)
' 
' Response.write isOdd(1) ' prints true
' Response.write isOdd(2) ' prints false
' 
' (end code)
' 
function isOdd(n)
    isOdd = cbool(n mod 2)
end function

' Function: isEven
' 
' Checks if a natural is even or not.
' 
' Parameters:
' 
'     (int) - Integer
' 
' Returns:
' 
'     (boolean) - true, if yes; false, otherwise
' 
' Example:
' 
' (start code)
' 
' Response.write isEven(1) ' prints false
' Response.write isEven(2) ' prints true
' 
' (end code)
' 
function isEven(n)
    isEven = (not isOdd(n))
end function

' Function: max
' 
' Returns the number with the highest value of an array.
' 
' Parameters:
' 
'     (float[]) - Decimals
' 
' Returns:
' 
'     (float) - Maximum
' 
' Example:
' 
' (start code)
' 
' Response.write max(array(0,1,2,3,4)) ' prints 4
' Response.write max(array(5,6,7,8,9)) ' prints 9
' 
' (end code)
' 
function max(a)
    dim i : i = 1
    max = a(0)
    while( i <= ubound(a) )
        if( cdbl(max) < cdbl(a(i)) ) then
            max = a(i)
        end if
        i = i + 1
    wend
end function

' Function: min
' 
' Returns the number with the lowest value of an array.
' 
' Parameters:
' 
'     (float[]) - Decimals
' 
' Returns:
' 
'     (float) - Minimum
' 
' Example:
' 
' (start code)
' 
' Response.write min(array(0,1,2,3,4)) ' prints 0
' Response.write min(array(5,6,7,8,9)) ' prints 5
' 
' (end code)
' 
function min(a)
    dim i : i = 1
    min = a(0)
    while( i <= ubound(a) )
        if( cdbl(min) > cdbl(a(i)) ) then
            min = a(i)
        end if
        i = i + 1
    wend
end function

' Function: floor
' 
' Returns the value of a number rounded downwards to the nearest integer.
' 
' Parameters:
' 
'     (float) - Decimal
' 
' Returns:
' 
'     (int) - Nearest integer equal-or-less than Decimal
' 
' Example:
' 
' (start code)
' 
' Response.write floor(2.71828182) ' prints 2
' Response.write floor(3.14159265) ' prints 3
' 
' (end code)
' 
function floor(n)
    dim r : r = round(n)
    if( cdbl(r) > cdbl(n) ) then
        r = r - 1
    end if
    floor = cint(r)
end function

' Function: ceiling
' 
' Returns the value of a number rounded upwards to the nearest integer.
' 
' Parameters:
' 
'     (float) - Decimal
' 
' Returns:
' 
'     (int) - Nearest integer equal-or-greater than Decimal
' 
' Example:
' 
' (start code)
' 
' Response.write ceiling(2.71828182) ' prints 3
' Response.write ceiling(3.14159265) ' prints 4
' 
' (end code)
' 
function ceiling(n)
    dim f : f = floor(n)
    if( cdbl(f) = cdbl(n) ) then
        ceiling = f
    else
        ceiling = f + 1
    end if
end function

' Function: hex2dec
' 
' Converts a hexadecimal number into a decimal one.
' 
' Parameters:
' 
'     (string) - Hexadecimal
' 
' Returns:
' 
'     (int) - Decimal
' 
' Example:
' 
' (start code)
' 
' Response.write hex2dec("FF") ' prints 255
' Response.write hex2dec("00") ' prints 0
' 
' (end code)
' 
function hex2dec(value)
    hex2dec = 0
    dim i, u, d
    i = 0 : u = len(value)
    while( u >= 1 )
        d = instr("0123456789ABCDEF", mid(value, u, 1))
        hex2dec = hex2dec + ( (d - 1) * (16 ^ i) )
        i = i + 1 : u = u - 1
    wend
end function

' Function: dec2hex
' 
' Converts a decimal number into a hexadecimal one.
' 
' Parameters:
' 
'     (int) - Decimal
' 
' Returns:
' 
'     (string) - Hexadecimal
' 
' Example:
' 
' (start code)
' 
' Response.write dec2hex(255) ' prints "FF"
' Response.write dec2hex(0)   ' prints 0
' 
' (end code)
' 
function dec2hex(value)
    dec2hex = hex(value)
end function

' Function: guid
' 
' Generates a Global Unique Identifier.
' 
' Returns:
' 
'     (uid) - global unique identifier string
' 
' Example:
' 
' (start code)
' 
' Response.write( guid() )
' 
' (end code)
' 
function guid()
    dim TypeLib : set TypeLib = Server.CreateObject("Scriptlet.TypeLib")
    guid = mid(cstr(TypeLib.Guid), 2, 36)
    set TypeLib = nothing
end function

' Function: dump
' 
' This function displays structured information about one or more expressions that includes its type and value. Arrays and objects are explored recursively with values indented to show structure.
' 
' Parameters:
' 
'     (mixed) - variable
' 
' Returns:
' 
'     (string) - information
' 
' Example:
' 
' (start code)
' 
' str = "nagaozen"
' num = 28
' flo = 3.1415
' bol = true
' dat = now
' arr = array(str, num, bol)
' mat = array(arr, arr)
' 
' 
' 
' set Re = new RegExp
' Re.pattern = "[A-Za-z0-9]"
' Re.global = true
' Re.ignoreCase = true
' Re.multiline = true
' 
' 
' 
' set Dic = Server.createObject("Scripting.Dictionary")
' Dic.add "firstname", "Fabio"
' Dic.add "lastname", "Nagao"
' Dic.add "age", 28
' Dic.add "male", true
' 
' 
' 
' set Rs = Server.createObject("ADODB.Recordset")
' Rs.Fields.append "name", 200, 64
' Rs.Fields.append "age", 3
' Rs.open()
' 
' Rs.addNew()
' Rs("name") = "Fabio Zendhi Nagao"
' Rs("age") = 28
' Rs.update()
' 
' Rs.addNew()
' Rs("name") = "Manoela Suzane de Alencar Rodrigues"
' Rs("age") = 27
' Rs.update()
' 
' Rs.moveFirst()
' 
' 
' 
' set St = Server.createObject("ADODB.Stream")
' St.type = 2' adTypeText
' St.mode = 3' adModeReadWrite
' St.open()
' St.position = 0
' St.writeText("ADODB.Stream text content")
' 
' 
' 
' Response.write( dump( str ) & vbNewline )
' Response.write( dump( num ) & vbNewline )
' Response.write( dump( flo ) & vbNewline )
' Response.write( dump( bol ) & vbNewline )
' Response.write( dump( dat ) & vbNewline )
' Response.write( dump( arr ) & vbNewline )
' Response.write( dump( mat ) & vbNewline )
' Response.write( dump( Re ) & vbNewline )
' Response.write( dump( Dic ) & vbNewline )
' Response.write( dump( Rs ) & vbNewline )
' Response.write( dump( St ) & vbNewline )
' 
' (end code)
' 
function dump(mixed)
    select case lcase(typename(mixed))
        case "string":
            dump = "(" & typename(mixed) & ") """ & mixed & """"
        case "iregexp2":
            dump = "(RegExp) /" & mixed.pattern & "/" & iif(mixed.global, "g", "") & iif(mixed.ignoreCase, "i", "") & iif(mixed.multiline, "m", "")
        case "variant()":
            dump = dump_hdlArray(mixed)
        case "dictionary":
            dump = dump_hdlDictionary(mixed)
        case "recordset":
            dump = dump_hdlRecordset(mixed)
        case "stream":
            dump = dump_hdlStream(mixed)
        case else:
            on error resume next
            dump = "(" & typename(mixed) & ") " & mixed
            if(Err.number <> 0) then
                dump = "(" & typename(mixed) & ")"
                Err.clear()
            end if
            on error goto 0
    end select
end function

function dump_hdlArray(arr)
    dim entry : dump_hdlArray = ""
    for each entry in arr
        dump_hdlArray = dump_hdlArray & "," & dump(entry)
    next
    dump_hdlArray = "(Array) [" & mid(dump_hdlArray, 2) & "]"
end function

function dump_hdlDictionary(Dic)
    dim key : dump_hdlDictionary = ""
    for each key in Dic.keys()
        dump_hdlDictionary = dump_hdlDictionary & ",""" & key & """: " & dump(Dic(key))
    next
    dump_hdlDictionary = "(Dictionary) {" & mid(dump_hdlDictionary, 2) & "}"
end function

function dump_hdlRecordset(Rs)
    dim i, row : dump_hdlRecordset = ""
    while( not Rs.eof )
        row = ""
        for i = 0 to Rs.Fields.count - 1
            row = row & ",""" & Rs(i).name & """: " & dump(Rs(i).value)
        next
        dump_hdlRecordset = dump_hdlRecordset & ",(Record) {" & mid(row, 2) & "}"
        Rs.moveNext()
    wend
    dump_hdlRecordset = "(Recorset) [" & mid(dump_hdlRecordset, 2) & "]"
end function

function dump_hdlStream(St)
    if(St.type = 1) then
        dump_hdlStream = "(Stream) adTypeBinary"
    else
        St.position = 0
        dump_hdlStream = "(Stream) """ & St.readText() & """"
    end if
end function

' Function: lambda
' 
' Evaluates a javascript function and returns it's reference.
' 
' Parameters:
' 
'     (string) - Javascript function definition
' 
' Returns:
' 
'     (function) - Pointer to the evaluated function
' 
' Example:
' 
' (start code)
' 
' dim greet : set greet = lambda("function(name){return 'Hello' + name;}")
' Response.write( greet("nagaozen") )' prints Hello nagaozen
' 
' (end code)
' 
%>
<script language="javascript" runat="server">
function lambda(f) {
    if(/^function\s*\([ a-z0-9.$_,]*\)\s*{[\S\s]*}$/gim.test(f)) {
        eval("f = " + f.replace(/(\r|\n)/g, ''));
        return f;
    } else {
        return function() {};
    }
}
</script>
