<%

' File: translator.asp
' 
' AXE(ASP Xtreme Evolution) solution for multilingual applications.
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
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Oct 2009
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
    '   (float) - version
    ' 
    public classVersion
    
    ' Property: language
    ' 
    ' Holds the language scope to fetch for translations.
    ' 
    ' Contains:
    ' 
    '     (string) - locale short string. Defaults to "en-us".
    ' 
    public language
    
    ' Property: Adapter
    ' 
    ' Translator requires an implementation of Translator_Adapter_Interface.
    ' 
    ' Contains:
    ' 
    '     (Translator_Adapter_*) - Translator media adapter
    ' 
    public Adapter
    
    ' Subroutine: [_ε]
    ' 
    ' {private} Checks for an adapter assignment.
    ' 
    private sub [_ε]
        if( isEmpty(Adapter) ) then
            Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Missing an Adapter."
        end if
    end sub
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        language = "en-us"
    end sub
    
    private sub Class_terminate()
        set Adapter = nothing
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
    ' set T.Adapter = new Translator_Adapter_Ini
    ' T.language = "pt-br"
    ' Response.write(T("Hello World"))
    ' ' prints: Olá Mundo
    ' set T = nothing
    ' 
    ' (end code)
    ' 
    public default function getText(message) : call [_ε]
        getText = Adapter.getText(language, message)
    end function
    
end class

%>
