<%

' File: mssql.asp
' 
' AXE(ASP Xtreme Evolution) implementation of authentication utility.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2007-2010 Fabio Zendhi Nagao
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



' Class: Auth_Adapter_MSSQL
' 
' This class allows the application to test an authentication against a MSSQL
' database.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao (nagaozen) <http://zend.lojcomm.com.br> @ Nov 2010
' 
class Auth_Adapter_MSSQL' implements Auth_Interface
    
    ' --[ Interface ]-----------------------------------------------------------
    public Interface
    
    ' --[ Adapter definition ]--------------------------------------------------
    
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
    
    ' Property: connectionString
    ' 
    ' ADODB.Connection argument to specify either a connection string containing
    ' a series of argument= value statements separated by semicolons, or a file 
    ' or directory resource identified with a URL
    ' 
    ' Contains:
    ' 
    '     (string) - ADODB.Connection connectionString
    ' 
    public connectionString
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0"
        
        set Interface = new Auth_Interface
        set Interface.Implementation = Me
        if(not Interface.check) then
            Err.raise 17, "Evolved AXE runtime error", strsubstitute( _
                "Can't perform requested operation. '{0}' is a bad interface implementation of '{1}'", _
                array(classType, typename(Interface)) _
            )
        end if
    end sub
    
    private sub Class_terminate()
        set Interface.Implementation = nothing
        set Interface = nothing
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
    '     (int) - Auth_Adapter_Interface.RESULT constant
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
    public function authenticate(usr, pwd)
        if(varType(connectionString) = vbEmpty) then
            Err.raise 5, "Evolved AXE runtime error", strsubstitute( _
                "Invalid procedure call or argument. '{0}' property 'connectionString' isn't defined", _
                array(classType) _
            )
        end if
        
        authenticate = "FAILURE_UNCATEGORIZED"
        
        dim Conn, Cmd, Rs
        set Conn = Server.createObject("ADODB.Connection")
        Conn.open(connectionstring)
        
        set Cmd = Server.createObject("ADODB.Command")
        Cmd.activeConnection = Conn
        Cmd.commandType = adCmdStoredProc
        Cmd.commandText = "SP_axe_auth_authenticate"
        
        Cmd.Parameters.append Cmd.createParameter("@usr", adVarWChar, adParamInput, 128)
        Cmd.Parameters("@usr").value = usr
        
        set Rs = Cmd.execute()
        if( Rs.eof ) then
            authenticate = "FAILURE_IDENTITY_NOT_FOUND"
        else
            if( Rs(0).value <> pwd ) then
                authenticate = "FAILURE_CREDENTIAL_INVALID"
            else
                authenticate = "SUCCESS"
            end if
        end if
        set Rs = nothing
        
        set Cmd = nothing
        
        Conn.close()
        set Conn = nothing
    end function
    
end class

%>
