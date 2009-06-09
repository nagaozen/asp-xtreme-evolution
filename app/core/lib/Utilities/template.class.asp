<%

'+-----------------------------------------------------------------------------+
'|This file is part of ASP Xtreme Evolution.                                   |
'|Copyright (C) 2007, 2008 Fabio Zendhi Nagao                                  |
'|                                                                             |
'|ASP Xtreme Evolution is free software: you can redistribute it and/or modify |
'|it under the terms of the GNU Lesser General Public License as published by  |
'|the Free Software Foundation, either version 3 of the License, or            |
'|(at your option) any later version.                                          |
'|                                                                             |
'|ASP Xtreme Evolution is distributed in the hope that it will be useful,      |
'|but WITHOUT ANY WARRANTY; without even the implied warranty of               |
'|MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                |
'|GNU Lesser General Public License for more details.                          |
'|                                                                             |
'|You should have received a copy of the GNU Lesser General Public License     |
'|along with ASP Xtreme Evolution.  If not, see <http://www.gnu.org/licenses/>.|
'+-----------------------------------------------------------------------------+

' Class: Template
' 
' A string Template class for Classic ASP.
' 
' About:
' 
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ May 2008
' 
class Template
    
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
    
    ' Property: value
    ' 
    ' This property holds the current Template value.
    ' 
    ' Contains:
    ' 
    '   (string) - value
    ' 
    public value
    
    ' Property: separator
    ' 
    ' This property holds the standard Template separator.
    ' 
    ' Contains:
    ' 
    '   (char) - separator
    ' 
    public separator
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.1.0"
        
        value = ""
        separator = " "
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Subroutine: write
    ' 
    ' Subroutine to Response.write templates directly from arguments.
    ' 
    ' Parameters:
    ' 
    '   (string)   - Template with placeholders.
    '   (string[]) - Array with the replacements.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' ' The infamous hello world example using Template.write
    ' dim String : set String = new Template
    ' String.write "{0} {1}", array("Hello", "World")
    ' String.write "{0} {1}", null
    ' set String = nothing
    ' 
    ' (end code)
    ' 
    ' See also:
    ' 
    '   <substitute>
    ' 
    public sub write(sTemplate, saArgs)
        dim i
        if( isArray(saArgs) ) then
            for i = 0 to ubound(saArgs)
                sTemplate = Replace(sTemplate, "{" & i & "}", saArgs(i))
            next
        end if
        Response.write sTemplate
    end sub
    
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.substitute(array("Hello", "World"))
    ' Response.write String.substitute(null)
    ' set String = nothing
    ' 
    ' (end code)
    ' 
    ' See also:
    ' 
    '   <write>
    ' 
    public function substitute(saArgs)
        dim sTemplate, i
        sTemplate = value
        if( isArray(saArgs) ) then
            for i = 0 to ubound(saArgs)
                sTemplate = Replace(sTemplate, "{" & i & "}", saArgs(i))
            next
        end if
        substitute = sTemplate
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
    ' dim String : set String = new Template : String.value = "{0} {1}, I     like     ASP."
    ' Response.write String.clean(array("Hello", "World"))
    ' Response.write String.clean(null)
    ' set String = nothing
    ' 
    ' (end code)
    ' 
    public function clean(saArgs)
        dim sTemplate : sTemplate = substitute(saArgs)
        dim Re : set Re = new RegExp
        Re.pattern = sPattern
        Re.ignoreCase = bIgnoreCase
        Re.global = bGlogal
        clean = trim(Re.replace(sTemplate, " "))
        set Re = nothing
    end function
    
    ' Function: length
    ' 
    ' Compute the final length of an evaluated template with the given
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.length(array("Hello", "World"))
    ' Response.write String.length(null)
    ' set String = nothing
    ' 
    ' (end code)
    ' 
    public function length(saArgs)
        length = len(substitute(saArgs))
    end function
    
    ' Function: toLowerCase
    ' 
    ' Compute the template and retrieve a lowercased version of it's value.
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.toLowerCase(array("Hello", "World"))
    ' Response.write String.toLowerCase(null)
    ' set String = nothing
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
    ' Compute the template and retrieve an uppercased version of it's value.
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.toUpperCase(array("Hello", "World"))
    ' Response.write String.toUpperCase(null)
    ' set String = nothing
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
    ' Compute the template and retrieve a capitalized version of it's value.
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.capitalize(, array("hello", "world"))
    ' Response.write String.capitalize(" ", null)
    ' set String = nothing
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
    ' Compute the template and retrieve a camelized version of it's value.
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.camelize(, array("Hello", "World"))
    ' Response.write String.camelize(" ", null)
    ' set String = nothing
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
    
    ' Function: hyphenate
    ' 
    ' Compute the template and retrieve a hyphenated version of it's value.
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.hyphenate(, array("Hello", "World"))
    ' Response.write String.hyphenate(" ", null)
    ' set String = nothing
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
    ' Look up in the evaluated template for a character at the specified position.
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.contains(0, array("Hello", "World"))
    ' set String = nothing
    ' 
    ' (end code)
    ' 
    public function charAt(i, saArgs)
        charAt = mid(substitute(saArgs), ( i + 1 ), 1)
    end function
    
    ' Function: substr
    ' 
    ' Extracts a specified number of characters in a evaluated template, from a
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.substr(0, 5, array("Hello", "World"))
    ' Response.write String.substr(6,, array("Hello", "World"))
    ' set String = nothing
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
    ' Extracts the characters in a evaluated template between two specified
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.substr(0, 2, array("Hello", "World"))
    ' Response.write String.substr(6,, array("Hello", "World"))
    ' set String = nothing
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
    ' Compute the template and look for the position of the first occurence of a
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.contains("Hello", array("Hello", "World"))
    ' Response.write String.contains("World", array("Hello", "World"))
    ' Response.write String.contains("Lorem ipsum", array("Hello", "World"))
    ' set String = nothing
    ' 
    ' (end code)
    ' 
    public function contains(sFragment, saArgs)
        contains = instr(substitute(saArgs), sFragment) - 1
    end function
    
    ' Function: test
    ' 
    ' Compute the template and search for a match between the value and a
    ' relular expression.
    ' 
    ' Parameters:
    ' 
    '   (string)  - String representation of the regular expression
    '   (boolean) - true|false indicating to enable a CI search or not
    '   (boolean) - true|false indicating to match or not all occurrences of the pattern
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
    ' dim String : set String = new Template : String.value = "{0} {1}"
    ' Response.write String.test("Hello", false, true, array("Hello", "World"))
    ' Response.write String.test("hello", true, true, array("Hello", "World"))
    ' set String = nothing
    ' 
    ' (end code)
    ' 
    public function test(sPattern, bIgnoreCase, bGlobal, saArgs)
        dim sTemplate : sTemplate = substitute(saArgs)
        dim Re : set Re = new RegExp
        Re.pattern = sPattern
        Re.ignoreCase = bIgnoreCase
        Re.global = bGlogal
        dim M : set M = Re.Execute(sTemplate)
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
