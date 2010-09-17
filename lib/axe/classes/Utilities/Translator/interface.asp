<%

' File: interface.asp
' 
' AXE(ASP Xtreme Evolution) translator object interface specification.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2010 Fabio Zendhi Nagao
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



' Class: Translator_Interface
' 
' Defines the common specifications required to implement a working adapter of 
' Translator class.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Jun 2010
' 
class Translator_Interface' extends Interface
    
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
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        set Parent = new Interface
        Parent.requireds = array("getText")
    end sub
    
    private sub Class_terminate()
        set Parent = nothing
    end sub
    
    ' Function: getText
    ' 
    ' Fetch for the translation of the message in the right language.
    ' 
    ' Parameters:
    ' 
    '     (lang)    - An identificator for the word/phrase.
    '     (message) - An identificator for the word/phrase.
    ' 
    ' Returns:
    ' 
    '     (string) - translated message
    ' 
    public default function getText(lang, message)
    end function
    
end class

%>
