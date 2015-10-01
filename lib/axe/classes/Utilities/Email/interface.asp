<%

' File: interface.asp
' 
' AXE(ASP Xtreme Evolution) email object interface specification.
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



' Class: Email_Interface
' 
' Defines the common specifications required to implement a working adapter of 
' Email class.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Jun 2010
' 
class Email_Interface' extends Interface
    
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
    '   (string) - version
    ' 
    public classVersion
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        set Parent = new Interface
        Parent.requireds = array("send")
    end sub
    
    private sub Class_terminate()
        set Parent = nothing
    end sub
    
    ' Subroutine: send
    ' 
    ' Procedure send is used to send an Email.
    ' 
    ' Parameters:
    ' 
    '     (Email) - An instance of Email class (id est an Email Object).
    ' 
    public sub send(Email)
    end sub
    
end class

%>
