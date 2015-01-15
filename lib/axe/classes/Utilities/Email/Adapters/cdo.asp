<%

' File: cdo.asp
' 
' AXE(ASP Xtreme Evolution) email adapter using CDO.Message.
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



' Class: Email_Adapter_CDO
' 
' This class provides an Email_Interface using Microsoft CDO.Message as it's 
' backend.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Jun 2010
' 
class Email_Adapter_CDO' implements Email_Interface
    
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
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        set Interface = new Email_Interface
        set Interface.Implementation = Me
        call Interface.check()
    end sub
    
    private sub Class_terminate()
        set Interface.Implementation = nothing
        set Interface = nothing
    end sub
    
    ' Function: [_getTo]
    ' 
    ' {private} Returns a csv version of receivers collection.
    ' 
    ' Parameters:
    ' 
    '     (Email) - An instance of Email class (id est an Email Object).
    ' 
    ' Returns:
    ' 
    '     (string) - serialized version of receivers collection
    ' 
    private function [_getTo](Email)
        [_getTo] = join(Email.Tos.keys(), ",")
    end function
    
    ' Function: [_getCC]
    ' 
    ' {private} Returns a csv version of carbon copy receivers collection.
    ' 
    ' Parameters:
    ' 
    '     (Email) - An instance of Email class (id est an Email Object).
    ' 
    ' Returns:
    ' 
    '     (string) - serialized version of carbon copy receivers collection
    ' 
    private function [_getCC](Email)
        [_getCC] = join(Email.Ccs.keys(), ",")
    end function
    
    ' Function: [_getBCC]
    ' 
    ' {private} Returns a csv version of blind carbon copy receivers collection collection.
    ' 
    ' Parameters:
    ' 
    '     (Email) - An instance of Email class (id est an Email Object).
    ' 
    ' Returns:
    ' 
    '     (string) - serialized version of blind carbon copy receivers collection collection
    ' 
    private function [_getBCC](Email)
        [_getBCC] = join(Email.Bccs.keys(), ",")
    end function
    
    ' Subroutine: send
    ' 
    ' Sends an Email using CDO (Collaboration Data Objects), a Microsoft
    ' technology that simplify the creation of messaging applications.
    ' 
    ' Parameters:
    ' 
    '     (Email) - An instance of Email class (id est an Email Object).
    ' 
    public sub send(Email)
        dim t, c, b
        t = [_getTo](Email)
        c = [_getCC](Email)
        b = [_getBCC](Email)
        
        dim Message : set Message = Server.createObject("CDO.Message")

        if( Email.Options.exists("Configuration") ) then
            dim Config, i, pair

            set Config = Server.createObject("CDO.Configuration")
            with Config.Fields
                for i = 0 to ubound( Email.Options("Configuration") )
                    pair = Email.Options("Configuration")(i)
                    .item(pair(0)) = pair(1)
                next
            end with

            set Message.Configuration = Config
        end if

        Message.from = Email.from
        Message.to = t
        if(len(c) > 0) then Message.cc = c
        if(len(b) > 0) then Message.bcc = b
        Message.subject = Email.subject
        if(Email.isHTML) then
            Message.htmlBody = Email.body
        else
            Message.textBody = Email.body
        end if
        if(Email.Attachments.count > 0) then
            dim path : for each path in Email.Attachments.keys()
                call Message.addAttachment(path)
            next
        end if
        call Message.send()

        if( Email.Options.exists("Configuration") ) then
            set Config = nothing
        end if

        set Message = nothing
    end sub
    
end class

%>
