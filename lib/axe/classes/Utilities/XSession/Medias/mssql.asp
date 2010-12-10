<%

' File: mssql.asp
' 
' AXE(ASP Xtreme Evolution) mssql database media.
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



' Class: XSession_Media_MSSQL
' 
' This class enables XSession to persist data using MSSQL database.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao (nagaozen) <http://zend.lojcomm.com.br> @ Nov 2010
' 
class XSession_Media_MSSQL' implements XSession_Interface
    
    ' --[ Interface ]-----------------------------------------------------------
    public Interface
    
    ' --[ Media definition ]--------------------------------------------------
    
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
        classVersion = "1.0.0.0"
        
        set Interface = new XSession_Interface
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
    
    ' Function: load
    ' 
    ' Returns the serialized XSession from the json file.
    ' 
    ' Parameters:
    ' 
    '     (string) - XSession unique id
    ' 
    ' Returns:
    ' 
    '     (string) - XSession serialized content
    ' 
    public function load(id)
        if(varType(connectionString) = vbEmpty) then
            Err.raise 5, "Evolved AXE runtime error", strsubstitute( _
                "Invalid procedure call or argument. '{0}' property 'connectionString' isn't defined", _
                array(classType) _
            )
        end if
        
        dim Conn, Cmd, Rs
        set Conn = Server.createObject("ADODB.Connection")
        Conn.open(connectionstring)
        
        set Cmd = Server.createObject("ADODB.Command")
        Cmd.activeConnection = Conn
        Cmd.commandType = adCmdStoredProc
        Cmd.commandText = "SP_axe_xsession_load"
        
        Cmd.Parameters.append Cmd.createParameter("@uid", adGuid, adParamInput, 16)
        Cmd.Parameters("@uid").value = id
        
        set Rs = Cmd.execute()
        if( Rs.eof ) then
            load = "{}"
        else
            load = Rs(0).value
        end if
        set Rs = nothing
        
        set Cmd = nothing
        
        Conn.close()
        set Conn = nothing
    end function
    
    ' Subroutine: save
    ' 
    ' Media writing routine.
    ' 
    ' Parameters:
    ' 
    '     (string) - XSession unique id
    '     (string) - XSession serialized content
    ' 
    public sub save(id, content)
        if(varType(connectionString) = vbEmpty) then
            Err.raise 5, "Evolved AXE runtime error", strsubstitute( _
                "Invalid procedure call or argument. '{0}' property 'connectionString' isn't defined", _
                array(classType) _
            )
        end if
        
        dim Conn, Cmd, Rs
        set Conn = Server.createObject("ADODB.Connection")
        Conn.open(connectionstring)
        
        set Cmd = Server.createObject("ADODB.Command")
        Cmd.activeConnection = Conn
        Cmd.commandType = adCmdStoredProc
        Cmd.commandText = "SP_axe_xsession_save"
        
        Cmd.Parameters.append Cmd.createParameter("@uid", adGuid, adParamInput, 16)
        Cmd.Parameters.append Cmd.createParameter("@content", adLongVarWChar, adParamInput, 1073741823)
        Cmd.Parameters("@uid").value = id
        Cmd.Parameters("@content").value = content
        
        set Rs = Cmd.execute()
        set Rs = nothing
        
        set Cmd = nothing
        
        Conn.close()
        set Conn = nothing
    end sub
    
end class

%>
