<%

' File: xsession.asp
' 
' AXE(ASP Xtreme Evolution) re-implementation of sessions.
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



' Class: XSession
' 
' Since AXE changes the Session object default behaviour to work only in the 
' request scope, another persistence mechanism is required to share information 
' across the pages. XSession class offers a really simple way to store JSON 
' compliant value (read javascript Object, array, string, number, boolean).
' 
' Note:
' 
' - DO NOT try to store ASP objects or classes in XSession, it doesn't work.
' 
' Dependencies:
' 
'     - JSON2 class (/lib/axe/classes/Parsers/json2.asp)
' 
' Requires:
' 
'     - XSession_Interface implementation (for persistence)
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Nov 2010
' 
class XSession
    
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
    
    ' Property: id
    ' 
    ' XSession Unique Id.
    ' 
    ' Contains:
    ' 
    '     (uid) - unique id
    ' 
    public id
    
    ' Property: [_Ω]
    ' 
    ' {private}
    ' 
    ' Contains:
    ' 
    '     (Object) - in memory temporary xsession storage
    ' 
    private [_Ω]
    
    ' Property: Media
    ' 
    ' XSession_Interface implementation
    ' 
    ' Contains:
    ' 
    '     (XSession_Interface) - Media implementing XSession_Interface
    ' 
    public Media
    
    ' Subroutine: [_ε]
    ' 
    ' {private} Checks for an media assignment.
    ' 
    private sub [_ε]
        if( isEmpty(Media) ) then
            Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Missing a XSession_Interface media."
        end if
    end sub
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        id = Request.Cookies("XSESSION_UID")
        if( id = "" ) then
            id = guid()
            XCookies.setItem "XSESSION_UID", id, 8 * 3600, false, "/", false, true
        end if
        
        set [_Ω] = JSON.parse("{}")
    end sub
    
    private sub Class_terminate()
        set [_Ω] = nothing
    end sub
    
    ' Function: [get]
    ' 
    ' Reads the value of a XSession.
    ' 
    ' Parameters:
    ' 
    '     (string) - XSession identifier
    ' 
    ' Returns:
    ' 
    '     (mixed) - XSession value
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim [$]
    ' 
    ' set [$] = new XSession
    ' set [$].Media = new XSession_Media_MSSQL
    ' 
    ' call [$].set("foo", "bar")
    ' Response.write( [$].get("foo") )' prints bar
    ' 
    ' set [$].Media = nothing
    ' set [$] = nothing
    ' 
    ' (end code)
    ' 
    public function [get](key)
        [get] = [_Ω].get(key)
    end function
    
    ' Subroutine: [set]
    ' 
    ' Writes a XSession.
    ' 
    ' Parameters:
    ' 
    '     (string) - XSession identifier
    '     (mixed)  - XSession value
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim [$]
    ' 
    ' set [$] = new XSession
    ' set [$].Media = new XSession_Media_MSSQL
    ' 
    ' call [$].set("foo", "bar")
    ' Response.write( [$].get("foo") )' prints bar
    ' 
    ' set [$].Media = nothing
    ' set [$] = nothing
    ' 
    ' (end code)
    ' 
    public sub [set](key, value)
        [_Ω].set key, value
    end sub
    
    ' Subroutine: load
    ' 
    ' Retrieves the XSession image from the persistence layer.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim [$]
    ' 
    ' set [$] = new XSession
    ' set [$].Media = new XSession_Media_MSSQL
    ' [$].Media.connectionString = "Provider=SQLOLEDB; Initial Catalog=..."
    ' 
    ' call [$].load()
    ' Response.write( [$].get("foo") )
    ' 
    ' set [$].Media = nothing
    ' set [$] = nothing
    ' 
    ' (end code)
    ' 
    public sub load() : call [_ε]
        set [_Ω] = JSON.parse( Media.load(id) )
    end sub
    
    ' Subroutine: save
    ' 
    ' Writes the XSession image in the persistence layer.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim [$]
    ' 
    ' set [$] = new XSession
    ' set [$].Media = new XSession_Media_MSSQL
    ' [$].Media.connectionString = "Provider=SQLOLEDB; Initial Catalog=..."
    ' 
    ' call [$].set("foo", "bar")
    ' call [$].save()
    ' 
    ' set [$].Media = nothing
    ' set [$] = nothing
    ' 
    ' (end code)
    ' 
    public sub save() : call [_ε]
        call Media.save( id, JSON.stringify([_Ω]) )
    end sub
    
end class

%>
