<%

' File: translator.asp
' 
' AXE(ASP Xtreme Evolution) solution for multilingual applications.
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



' Class: Translator
' 
' Translator provides support for internationalization of text in code and templates.
' 
' Requires:
' 
'     - Translator_Interface implementation
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Oct 2010
' 
class Translator
    
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
    
    ' Property: [_δ]
    ' 
    ' {private} The multilingual dictionary.
    ' 
    ' Contains:
    ' 
    '     (Object) - the multilingual dictionary
    ' 
    private [_δ]
    
    ' Property: language
    ' 
    ' Holds the language scope to fetch for translations. Usually the language 
    ' is the composition of two or three letters as defined in the ISO-639 
    ' language codes and two other letters for country as defined in ISO-3166 
    ' country codes, with a hyphen as the separator.
    ' 
    ' Contains:
    ' 
    '     (string) - locale short string. Defaults to "en-us".
    ' 
    public language
    
    ' Property: Media
    ' 
    ' Translator_Interface implementation.
    ' 
    ' Contains:
    ' 
    '     (Translator_Interface) - Media implementing Translator_Interface
    ' 
    public Media
    
    ' Subroutine: [_ε]
    ' 
    ' {private} Checks for an adapter assignment.
    ' 
    private sub [_ε]
        if( isEmpty(Media) ) then
            Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Missing a Translator_Interface media."
        end if
    end sub
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        language = "en-us"
        set [_δ] = JSON.parse("{}")
    end sub
    
    private sub Class_terminate()
        set [_δ] = nothing
    end sub
    
    ' Function: getText
    ' 
    ' Looks up a message in the current configuration.
    ' 
    ' Parameters:
    ' 
    '     (string) - Message
    ' 
    ' Returns:
    ' 
    '     (string) - The translated value if on is found in the translation table, or the submitted messeage if not found
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim T : set T = new Translator
    ' set T.Media = new Translator_Media_Ini
    ' T.language = "pt-br"
    ' Response.write(T("Hello World"))
    ' ' prints: Olá Mundo
    ' set T = nothing
    ' 
    ' (end code)
    ' 
    public default function getText(message)
        getText = message
        if( not isEmpty( [_δ].get(language) ) ) then
            if( not isEmpty( [_δ].get(language).get(message) ) ) then
                getText = [_δ].get(language).get(message)
            end if
        end if
    end function
    
    ' Subroutine: load
    ' 
    ' Retrieves the Translator image from the persistence layer.
    ' 
    public sub load() : call [_ε]
        set [_δ] = JSON.parse( Media.load() )
    end sub
    
end class

%>
