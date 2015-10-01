<%

' File: auth.asp
' 
' AXE(ASP Xtreme Evolution) implementation of authentication utility.
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



' Class: Auth
' 
' Auth is concerned only with authentication and not with authorization. 
' Authentication is loosely defined as determining whether an entity actually 
' is what it purports to be (i.e., identification), based on some set of 
' credentials. Authorization, the process of deciding whether to allow an entity 
' access to, or to perform operations upon, other entities is outside the scope 
' of Auth.
' 
' Requires:
' 
'     - Auth_Interface implementation
' 
' See also:
' 
' <Acl>
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ April 2010
' 
class Auth
    
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
    
    ' Property: Adapter
    ' 
    ' Auth class requires an Auth_Interface implementation to work.
    ' 
    ' Contains:
    ' 
    '     (Auth_Interface) - Auth_Interface implementation to use.
    ' 
    public Adapter
    
    ' Subroutine: [_ε]
    ' 
    ' {private} Checks for an adapter assignment.
    ' 
    private sub [_ε]
        if( isEmpty(Adapter) ) then
            Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Missing an Auth_Interface implementation."
        end if
    end sub
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: authenticate
    ' 
    ' Performs an authentication attempt.
    ' 
    ' Parameters:
    ' 
    '     (string) - user
    '     (string) - password
    ' 
    ' Returns:
    ' 
    '     (int) - Auth_Interface.RESULT constant
    ' 
    ' Result constants:
    ' 
    '   SUCCESS
    '   FAILURE
    '   FAILURE_IDENTITY_NOT_FOUND
    '   FAILURE_IDENTITY_AMBIGUOUS
    '   FAILURE_CREDENTIAL_INVALID
    '   FAILURE_UNCATEGORIZED
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Authenticator : set Authenticator = new Auth
    ' set Authenticator.Adapter = new Auth_Adapter_MSSQL
    ' Authenticator.Adapter.connectionString = "Provider=SQLOLEDB; ..."
    ' 
    ' set [ƒ] = new MD5
    ' Response.write(Authenticator.authenticate("foo", [ƒ].encryptData("bar")) & vbNewline) ' prints FAILURE_IDENTITY_NOT_FOUND
    ' Response.write(Authenticator.authenticate("nagaozen", [ƒ].encryptData("wrong_password")) & vbNewline) ' prints SUCCESS
    ' Response.write(Authenticator.authenticate("nagaozen", [ƒ].encryptData("right_password")) & vbNewline) ' prints SUCCESS
    ' set [ƒ] = nothing
    ' 
    ' set Authenticator.Adapter = nothing
    ' set Authenticator = nothing
    ' 
    ' (end code)
    ' 
    public function authenticate(usr, pwd) : call [_ε]
        authenticate = Adapter.authenticate(usr, pwd)
    end function
    
end class

%>
