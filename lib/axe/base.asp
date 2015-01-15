<%

' File: base.asp
' 
' All asp pages should include this file.
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
function iif(byVal expr, byRef truepart, byRef falsepart)
    if(expr) then
        iif = truepart
    else
        iif = falsepart
    end if
end function

' Function: getDate
' 
' A location independent date generation function.
' 
' Parameters:
' 
'     (integer) - year
'     (integer) - month
'     (integer) - day
'     (integer[]) - optional [ hour, minute, second ]
' 
' Returns:
' 
'     (Date) - date
' 
' Example:
' 
' (start code)
' 
' dim shortDate, generalDate
' shortDate   = getDate(1982, 08, 31, null)
' generalDate = getDate(1982, 08, 31, array(03, 15, 00))
' 
' (end code)
' 
public function getDate(byVal yyyy, byVal mm, byVal dd, byRef time)
    dim dt, tm, h, n, s

    dt = dateSerial(yyyy, mm, dd)
    if( not isNull(time) ) then
        if( isArray(time) ) then
            h = time(0)
            n = time(1)
            s = time(2)
            tm = timeSerial(h, n, s)
        end if
    end if

    if( isEmpty(tm) ) then
        getDate = dt
    else
        getDate = cdate( dt & " " & tm )
    end if
end function

' Function: formatDate
' 
' Outputs the date into a specific format.
' 
' Parameters:
' 
'     (Date)   - date
'     (string) - format
' 
' Returns:
' 
'     (string) - formated date
' 
' Example:
' 
' (start code)
' 
' dim generalDate
' generalDate = getDate(1982, 08, 31, array(03, 15, 00))
' Response.write formatDate(generalDate, "dd/mm/yyyy") ' prints "31/08/1982"
' Response.write formatDate(generalDate, "mm/dd/yyyy h:n:s") ' prints "08/31/1982 03:15:00"
' 
' (end code)
' 
function formatDate(byVal dt, byVal fmt)
    formatDate = fmt

    formatDate = replace(formatDate, "yyyy", year(dt))
    formatDate = replace(formatDate, "mm", right("00" & month(dt), 2))
    formatDate = replace(formatDate, "dd", right("00" & day(dt), 2))

    formatDate = replace(formatDate, "h", right("00" & hour(dt), 2))
    formatDate = replace(formatDate, "n", right("00" & minute(dt), 2))
    formatDate = replace(formatDate, "s", right("00" & second(dt), 2))
end function

' Function: moment
' 
' Returns a string representation of the current time with miliseconds.
' 
' Returns:
' 
'     (string) - moment with miliseconds
' 
' Example:
' 
' (start code)
' 
' Response.write( moment() )
' 
' (end code)
' 
public function moment()
    dim hh, mm, ss, ms _
      , tm, x

    tm = timer
    x  = int(tm)

    ms = int( ( tm - x ) * 1000 )
    ss = x mod 60

    x  = int( x / 60 )

    mm = x mod 60
    hh = int( x / 60 )

    moment = strsubstitute( "{0}:{1}:{2}.{3}", array( hh, mm, ss, ms ) )
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
function strsubstitute(byVal template, byVal replacements)
    if( not isArray(replacements) ) then replacements = array(replacements)
    
    dim str, i
    str = template
    for i = 0 to ubound(replacements)
        if( not isNull( replacements(i) ) ) then
            str = replace(str, "{" & i & "}", replacements(i))
        end if
    next
    strsubstitute = str
end function

' Function: strfilter
' 
' Filters a string by a list of valid characters.
' 
' Parameters:
' 
'     (string) - unfiltered value
'     (string) - valid characters
' 
' Returns:
' 
'     (string) - filtered value
' 
' Example:
' 
' (start code)
' 
' dim phone : phone = "555 1234-5678"
' Response.write( strfilter(phone, "0123456789") )' prints 55512345678
' 
' (end code)
' 
public function strfilter(byVal value, byVal valids)
    dim i, c
    with ( Server.createObject("ADODB.Stream") )
        .type = adTypeText
        .mode = adModeReadWrite
        .open()

        for i = 1 to len(value)
            c = mid(value, i, 1)
            if( instr( 1, valids, c, 1 ) ) then
                .writeText(c)
            end if
        next

        .position = 0
        strfilter = .readText()
    end with
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
public function sanitize(byVal value, byRef placeholders, byRef replacements)
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
function isOdd(byVal n)
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
function isEven(byVal n)
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
function max(byRef a)
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
function min(byRef a)
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
function floor(byVal n)
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
function ceiling(byVal n)
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
function hex2dec(byVal value)
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
function dec2hex(byVal value)
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
' This function displays structured information about one or more expressions 
' that includes its type and value. Arrays and objects are explored recursively 
' with values indented to show structure.
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
function dump(byRef mixed)
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

function dump_hdlArray(byRef arr)
    dim entry : dump_hdlArray = ""
    for each entry in arr
        dump_hdlArray = dump_hdlArray & "," & dump(entry)
    next
    dump_hdlArray = "(Array) [" & mid(dump_hdlArray, 2) & "]"
end function

function dump_hdlDictionary(byRef Dic)
    dim key : dump_hdlDictionary = ""
    for each key in Dic.keys()
        dump_hdlDictionary = dump_hdlDictionary & ",""" & key & """: " & dump(Dic(key))
    next
    dump_hdlDictionary = "(Dictionary) {" & mid(dump_hdlDictionary, 2) & "}"
end function

function dump_hdlRecordset(byRef Rs)
    dim i, row : dump_hdlRecordset = ""
    while( not Rs.eof )
        row = ""
        for i = 0 to Rs.Fields.count - 1
ON ERROR RESUME NEXT : Err.clear()
            row = row & ",""" & Rs(i).name & """: " & dump(Rs(i).value)
            if( Err.number <> 0 ) then' probably a decimal type, and because VBScript can't handle'em, convert to currency
                row = row & ",""" & Rs(i).name & """: " & dump( ccur( Rs(i).value ) )
            end if
ON ERROR GOTO 0
        next
        dump_hdlRecordset = dump_hdlRecordset & ",(Record) {" & mid(row, 2) & "}"
        Rs.moveNext()
    wend
    dump_hdlRecordset = "(Recorset) [" & mid(dump_hdlRecordset, 2) & "]"
end function

function dump_hdlStream(byRef St)
    if(St.type = 1) then
        dump_hdlStream = "(Stream) adTypeBinary"
    else
        St.position = 0
        dump_hdlStream = "(Stream) """ & St.readText() & """"
    end if
end function

%>
<script language="javascript" runat="server">

/*
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
*/
function lambda(f) {
    if(/^function\s*\([ a-z0-9.$_,]*\)\s*{[\S\s]*}$/gim.test(f)) {
        eval("f = " + f.replace(/(\r|\n)/g, ''));
        return f;
    } else {
        return function() {};
    }
}

/*
' Singleton: XCookies
' 
' A simple and better little Cookies framework which implements RFC6265 enabling
' Classic ASP to work with modern Cookie options as max-age and httponly.
' 
' Example:
' 
' (start code)
' 
' ' @ setter.asp
' XCookies.setItem "Classic ASP Framework", "ASP Xtreme Evolution"'[, (variant)end, (string)domain, (string)path, (boolean)secure, (boolean)httpOnly]
' 
' ' @ getter.asp
' Response.write Request.Cookies("Classic ASP Framework")
'
' ' @ remover.asp
' XCookies.removeItem "Classic ASP Framework"
' 
' (end code)
' 
' Known bugs:
' 
'     (WONTFIX) - If one `set` a Cookie and tries to use it in the same Response Stream, it won't be available. A: This happens because the Cookie hasn't been delivered to the client and returnet yet. Therefore it's not in the current available Cookies Collection. So, better use the value assigned to the Cookie instead.
' 
' (start code)
'
' ' NOTE: WITH EMPTY COOKIES
' XCookies.setItem "Classic ASP Framework", "ASP Xtreme Evolution"
' Response.write Request.Cookies("Classic ASP Framework")' doesnt print anything
' 
' (end code)
' 
' Notes:
' 
'     - For never-expire-cookies we used the arbitrarily distant date Fri, 31 Dec 9999 23:59:59 GMT. If for any reason you are afraid of such a date, use the conventional date of end of the world <http://en.wikipedia.org/wiki/Year_2038_problem> `Tue, 19 Jan 2038 03:14:07 GMT` – which is the maximum number of seconds elapsed since since 1 January 1970 00:00:00 UTC expressible by a signed 32-bit integer (i.e.: 01111111111111111111111111111111, which is new Date(0x7fffffff * 1e3)).
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Aug 2013
' 
' More:
' 
'     - IETF HTTP State Management Mechanism - RFC 6265 <http://tools.ietf.org/html/rfc6265>
'     - MSDN Some Bad News and Some Good News <http://msdn.microsoft.com/en-us/library/ms972826>
'     - MSDN Response.Cookies Collection <http://msdn.microsoft.com/en-us/library/ms524757(v=vs.90).aspx>
'     - MSDN Response.AddHeader Method <http://msdn.microsoft.com/en-us/library/ms524327(v=vs.90).aspx>
'     - MDN A little Cookie framework <https://developer.mozilla.org/en-US/docs/Web/API/document.cookie>
' 
*/
var XCookies = {

    getItem: function(sKey) {
        return Request.Cookies(sKey);
    },

    setItem: function(sKey, vValue, vEnd, sDomain, sPath, bSecure, bHttpOnly) {
        if(!sKey || /^(?:expires|max\-age|path|domain|secure|httpOnly)$/i.test(sKey))
                return false;

        var cookie      = ""
          , cookiePair  = ""
          , cookieName  = encodeURIComponent(sKey)
          , cookieValue = null
          , vValueKey   = ""
          , sExpires    = ""
        ;

        function _isObject(v) {
            if( ( v instanceof Object ) && ( !v.hasOwnProperty("length") ) )
                return true;
            return false;
        }

        if(vEnd) {
            switch (vEnd.constructor) {
                case Number:
                    sExpires = ( vEnd === Infinity ? "; expires=Fri, 31 Dec 9999 23:59:59 GMT" : "; max-age=" + vEnd );
                    break;
                case String:
                    sExpires = "; expires=" + vEnd;
                    break;
                case Date:
                    sExpires = "; expires=" + vEnd.toGMTString();
                    break;
            }
        }

        if( _isObject(vValue) ) {
            cookieValue = [];
            for(vValueKey in vValue) {
                if(vValue.hasOwnProperty(vValueKey)) {
                    cookieValue.push( encodeURIComponent(vValueKey) + "=" + encodeURIComponent( vValue[vValueKey] ) );
                }
            }
            cookieValue = cookieValue.join("&");
        } else {
            cookieValue = encodeURIComponent(vValue);
        }

        cookiePair = [ cookieName , cookieValue ].join("=");

        cookie = cookiePair
               + sExpires
               + (sDomain ? "; domain=" + sDomain : "")
               + (sPath ? "; path=" + sPath : "")
               + (bSecure ? "; secure" : "")
               + (bHttpOnly ? "; httpOnly" : "")
        ;

        Response.AddHeader("Set-Cookie", cookie);

        return true;
    },

    removeItem: function(sKey, sPath) {
        if(!sKey) return false;

        var cookie = encodeURIComponent(sKey)
                   + "=; expires=Thu, 01 Jan 1970 00:00:00 GMT"
                   + ( sPath ? "; path=" + sPath : "" )
        ;

        Response.AddHeader("Set-Cookie", cookie);

        return true;
    }

};

</script>
