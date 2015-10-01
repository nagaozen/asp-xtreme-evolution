<%

' File: xstring.asp
' 
' AXE(ASP Xtreme Evolution) eXtended string utility.
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



' Class: XString
' 
' A eXtended string class for Classic ASP.
' 
' About:
' 
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ May 2008
' 
class XString
    
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
    
    ' Property: value
    ' 
    ' This property holds the current XString value.
    ' 
    ' Contains:
    ' 
    '   (string) - value
    ' 
    public value
    
    ' Property: separator
    ' 
    ' This property holds the standard XString separator.
    ' 
    ' Contains:
    ' 
    '   (char) - separator
    ' 
    public separator
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.1.0.1"
        
        value = ""
        separator = " "
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: setValue
    ' 
    ' This is the class default method. It sets the instance value property.
    ' 
    ' Parameters:
    ' 
    '     (string) - xstring value
    ' 
    ' Returns:
    ' 
    '     (XString) - a self object reference
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString
    ' S("{0} {1}").value' prints "{0} {1}"
    ' S("{0} {1}").substitute(array("Hello", "World"))' prints "Hello World"
    ' S("{0} {1}").toLowerCase(array("Hello", "World"))' prints "hello world"
    ' S("{0} {1}").<any_other_method_of_this_class>(arguments)' just works
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    public default function setValue(s)
        value = s
        set setValue = Me
    end function
    
    ' Function: substitute
    ' 
    ' Use this method to retrieve an evaluated version of the value with the
    ' given replacements.
    ' 
    ' Parameters:
    ' 
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (string) - Evaluated string
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.substitute(array("Hello", "World"))
    ' Response.write S.substitute(null)
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    public function substitute(byVal saArgs)
        if( not isArray(saArgs) ) then saArgs = array(saArgs)
        
        dim sXString, i
        sXString = value
        for i = 0 to ubound(saArgs)
            if( not isNull( saArgs(i) ) ) then
                sXString = replace(sXString, "{" & i & "}", saArgs(i))
            end if
        next
        substitute = sXString
    end function
    
    ' Function: []
    ' 
    ' Short hand name for substitute.
    ' 
    ' Parameters:
    ' 
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (string) - Evaluated string
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString
    ' Response.write S("Do you know {0} {1}?").[](array("foo", "bar")) ' prints Do you know foo bar?
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    public function [](saArgs)
        [] = substitute(saArgs)
    end function
    
    ' Function: clean
    ' 
    ' Removes all extraneous whitespace from a string and trims it.
    ' 
    ' Parameters:
    ' 
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (string) - A clean evaluated string.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}, I     like     ASP."
    ' Response.write S.clean(array("Hello", "World"))
    ' Response.write S.clean(null)
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    public function clean(saArgs)
        dim sXString : sXString = substitute(saArgs)
        dim Re : set Re = new RegExp
        Re.pattern = sPattern
        Re.ignoreCase = bIgnoreCase
        Re.global = bGlogal
        clean = trim(Re.replace(sXString, " "))
        set Re = nothing
    end function
    
    ' Function: length
    ' 
    ' Compute the final length of an evaluated xstring with the given
    ' replacements.
    ' 
    ' Parameters:
    ' 
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (int) - Length
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.length(array("Hello", "World"))
    ' Response.write S.length(null)
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    public function length(saArgs)
        length = len(substitute(saArgs))
    end function
    
    ' Function: toLowerCase
    ' 
    ' Compute the xstring and retrieve a lowercased version of it's value.
    ' 
    ' Parameters:
    ' 
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (string) - Lowercased evaluated string
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.toLowerCase(array("Hello", "World"))
    ' Response.write S.toLowerCase(null)
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    ' See also:
    ' 
    '   <toUpperCase>, <capitalize>, <camelize>, <hyphenate>
    ' 
    public function toLowerCase(saArgs)
        toLowerCase = lcase(substitute(saArgs))
    end function
    
    ' Function: toUpperCase
    ' 
    ' Compute the xstring and retrieve an uppercased version of it's value.
    ' 
    ' Parameters:
    ' 
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (string) - Uppercased evaluated string
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.toUpperCase(array("Hello", "World"))
    ' Response.write S.toUpperCase(null)
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    ' See also:
    ' 
    '   <toLowerCase>, <capitalize>, <camelize>, <hyphenate>
    ' 
    public function toUpperCase(saArgs)
        toUpperCase = ucase(substitute(saArgs))
    end function
    
    ' Function: capitalize
    ' 
    ' Compute the xstring and retrieve a capitalized version of it's value.
    ' 
    ' Parameters:
    ' 
    '   (char)     - Separator (OPTIONAL)
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (string) - Capitalized evaluated string
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.capitalize(, array("hello", "world"))
    ' Response.write S.capitalize(" ", null)
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    ' See also:
    ' 
    '   <toLowerCase>, <toUpperCase>, <camelize>, <hyphenate>
    ' 
    public function capitalize(cSeparator, saArgs)
        dim saSplit, i
        if( varType(cSeparator) = 10 ) then cSeparator = separator end if
        saSplit = split(toLowercase(saArgs), cSeparator)
        for i = 0 to ubound(saSplit)
            saSplit(i) = ucase(left(saSplit(i), 1)) & mid(saSplit(i), 2)
        next
        capitalize = join(saSplit, cSeparator)
    end function
    
    ' Function: camelize
    ' 
    ' Compute the xstring and retrieve a camelized version of it's value.
    ' 
    ' Parameters:
    ' 
    '   (char)     - Separator (OPTIONAL)
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (string) - Camelized evaluated string
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.camelize(, array("Hello", "World"))
    ' Response.write S.camelize(" ", null)
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    ' See also:
    ' 
    '   <toLowerCase>, <toUpperCase>, <capitalize>, <hyphenate>
    ' 
    public function camelize(cSeparator, saArgs)
        dim saSplit, i
        if( varType(cSeparator) = 10 ) then cSeparator = separator end if
        saSplit = split(toLowercase(saArgs), cSeparator)
        if( ubound(saSplit) >= 1 ) then
            for i = 1 to ubound(saSplit)
                saSplit(i) = ucase(left(saSplit(i), 1)) & mid(saSplit(i), 2)
            next
        end if
        camelize = join(saSplit, "")
    end function
    
    ' Function: propercase
    ' 
    ' Compute the xstring and retrieve a propercased (PascalCase) version of it's value.
    ' 
    ' Parameters:
    ' 
    '   (char)     - Separator (OPTIONAL)
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (string) - ProperCased evaluated string
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString
    ' Response.write( S("hello world").propercase(null, null) )
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    public function propercase(cSeparator, saArgs)
        propercase = capitalize(cSeparator, saArgs)
        propercase = replace(propercase, cSeparator, "")
    end function
    
    ' Function: hyphenate
    ' 
    ' Compute the xstring and retrieve a hyphenated version of it's value.
    ' 
    ' Parameters:
    ' 
    '   (char)     - Separator (OPTIONAL)
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (string) - Hyphenated evaluated string
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.hyphenate(, array("Hello", "World"))
    ' Response.write S.hyphenate(" ", null)
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    ' See also:
    ' 
    '   <toLowerCase>, <toUpperCase>, <capitalize>, <camelize>
    ' 
    public function hyphenate(cSeparator, saArgs)
        dim saSplit, i
        if( varType(cSeparator) = 10 ) then cSeparator = separator end if
        saSplit = split(toLowercase(saArgs), cSeparator)
        hyphenate = join(saSplit, "-")
    end function
    
    ' Function: charAt
    ' 
    ' Look up in the evaluated xstring for a character at the specified position.
    ' 
    ' Parameters:
    ' 
    '   (int)      - Index
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '     (char) - Character
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.contains(0, array("Hello", "World"))
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    public function charAt(i, saArgs)
        charAt = mid(substitute(saArgs), ( i + 1 ), 1)
    end function
    
    ' Function: substr
    ' 
    ' Extracts a specified number of characters in a evaluated xstring, from a
    ' start index.
    ' 
    ' Parameters:
    ' 
    '   (int)      - Where to start the extraction
    '   (int)      - How many characters to extract (OPTIONAL)
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (string) - Evaluated string fragment
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.substr(0, 5, array("Hello", "World"))
    ' Response.write S.substr(6,, array("Hello", "World"))
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    public function substr(iStart, iLength, saArgs)
        dim sEvaluated : sEvaluated = substitute(saArgs)
        dim iMaxLength : iMaxLength = len(sEvaluated) - iStart
        if( varType(iLength) = 10 ) then
            iLength = iMaxLength
        elseif( iLength > iMaxLength ) then
            iLength = iMaxLength
        end if
        substr = mid(sEvaluated, ( iStart + 1 ), iLength)
    end function
    
    ' Function: substring
    ' 
    ' Extracts the characters in a evaluated xstring between two specified
    ' indices.
    ' 
    ' Parameters:
    ' 
    '   (int)      - Where to start the extraction
    '   (int)      - Where to stop the extraction (OPTIONAL)
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '     (string) - Evaluated string fragment
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.substr(0, 2, array("Hello", "World"))
    ' Response.write S.substr(6,, array("Hello", "World"))
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    public function substring(iStart, iStop, saArgs)
        dim sEvaluated : sEvaluated = substitute(saArgs)
        dim iEvaluatedLength : iEvaluatedLength = len(sEvaluated)
        dim iMaxLength : iMaxLength = iEvaluatedLength - iStart
        if( varType(iStop) = 10 ) then
            iStop = iMaxLength
        elseif( iStop > iEvaluatedLength ) then
            iStop = iMaxLength
        else
            iStop = iStop - iStart
        end if
        substring = mid(sEvaluated, ( iStart + 1 ), iStop)
    end function
    
    ' Function: contains
    ' 
    ' Compute the xstring and look for the position of the first occurence of a
    ' specified fragment in the evaluated string.
    ' 
    ' Parameters:
    ' 
    '   (string)   - Fragment to look for
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   (int) - The position of the first occurence of the fragment within value
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.contains("Hello", array("Hello", "World"))
    ' Response.write S.contains("World", array("Hello", "World"))
    ' Response.write S.contains("Lorem ipsum", array("Hello", "World"))
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    public function contains(sFragment, saArgs)
        contains = instr(substitute(saArgs), sFragment) - 1
    end function
    
    ' Function: test
    ' 
    ' Compute the xstring and search for a match between the value and a
    ' relular expression.
    ' 
    ' Parameters:
    ' 
    '   (string)  - VBScript regular expression
    '   (boolean) - true|false indicating to enable a CI search or not
    '   (boolean) - true|false indicating to match or not all occurrences of the pattern
    '   (string[]) - Replacements
    ' 
    ' Returns:
    ' 
    '   true, if a match for the regular expression is found in the value
    '   false, otherwise
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim S : set S = new XString : S.value = "{0} {1}"
    ' Response.write S.test("Hello", false, true, array("Hello", "World"))
    ' Response.write S.test("hello", true, true, array("Hello", "World"))
    ' set S = nothing
    ' 
    ' (end code)
    ' 
    public function test(sPattern, bIgnoreCase, bGlobal, saArgs)
        dim sXString : sXString = substitute(saArgs)
        dim Re : set Re = new RegExp
        Re.pattern = sPattern
        Re.ignoreCase = bIgnoreCase
        Re.global = bGlogal
        dim M : set M = Re.Execute(sXString)
        if M.count > 0 then
            test = true
        else
            test = false
        end if
        set M = nothing
        set Re = nothing
    end function
    
end class

%>
