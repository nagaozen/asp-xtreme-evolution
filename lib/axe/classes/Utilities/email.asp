<%

' File: email.asp
' 
' AXE(ASP Xtreme Evolution) email utility.
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



' Class: Email
' 
' Conceptual email object.
' 
' Requires:
' 
'     - Email_Interface implementation
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ May 2010
' 
class Email
    
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
    
    ' Property: Adapter
    ' 
    ' Email_Interface implementation
    ' 
    ' Contains:
    ' 
    '     (Email_Interface) - Adapter implementing Email_Interface
    ' 
    public Adapter
    
    ' Subroutine: [_ε]
    ' 
    ' {private} Checks for an adapter assignment.
    ' 
    private sub [_ε]
        if( isEmpty(Adapter) ) then
            Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Missing an Email_Interface adapter."
        end if
    end sub
    
    ' Property: from
    ' 
    ' From field is used to indicate source or origin.
    ' 
    ' Contains:
    ' 
    '     (string) - sender email
    ' 
    public from
    
    ' Property: Tos
    ' 
    ' This collection is used to populate the email to field. To field is used 
    ' for expressing destination or appointed end.
    ' 
    ' Contains:
    ' 
    '     (Scripting.Dictionary) - receivers email collection
    ' 
    public Tos
    
    ' Property: Ccs
    ' 
    ' This collection is used to populate the email carbon copy field. Cc field 
    ' is used to send a replica of the email to other users.
    ' 
    ' Contains:
    ' 
    '     (Scripting.Dictionary) - carbon copy receivers email collection
    ' 
    public Ccs
    
    ' Property: Bccs
    ' 
    ' This collection is used to populate the email blind carbon copy field. Bcc
    ' field is much like Cc field, but will not be seen by the recipients.
    ' 
    ' Contains:
    ' 
    '     (Scripting.Dictionary) - blind carbon copy receivers email collection
    ' 
    public Bccs
    
    ' Property: subject
    ' 
    ' Email subject, theme, topic...
    ' 
    ' Contains:
    ' 
    '     (string) - subject
    ' 
    public subject
    
    ' Property: body
    ' 
    ' Email body.
    ' 
    ' Contains:
    ' 
    '     (string) - email content
    ' 
    public body
    
    ' Property: isHTML
    ' 
    ' Switch to turn on/off the HTML body mode.
    ' 
    ' Contains:
    ' 
    '     (boolean) - Send in HTML mode? Defaults to true.
    ' 
    public isHTML
    
    ' Property: Attachments
    ' 
    ' This collection is used to hold the fully qualified file name of the 
    ' attachments. An email attachment is a computer file sent along with an 
    ' email message.
    ' 
    ' Contains:
    ' 
    '     (Scripting.Dictionary) - attachments collection
    ' 
    public Attachments
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0"
        
        set Tos = Server.createObject("Scripting.Dictionary")
        set Ccs = Server.createObject("Scripting.Dictionary")
        set Bccs = Server.createObject("Scripting.Dictionary")
        set Attachments = Server.createObject("Scripting.Dictionary")
        
        isHTML = true
    end sub
    
    private sub Class_terminate()
        set Attachments = nothing
        set Bccs = nothing
        set Ccs = nothing
        set Tos = nothing
    end sub
    
    ' Subroutine: addTo
    ' 
    ' Adds an email to the receivers collection.
    ' 
    ' Parameters:
    ' 
    '     (string) - email
    ' 
    public sub addTo(email)
        if( not Tos.exists(email) ) then
            call Tos.add(email, Tos.count)
        end if
    end sub
    
    ' Subroutine: removeTo
    ' 
    ' Removes an email from the receivers collection.
    ' 
    ' Parameters:
    ' 
    '     (string) - email
    ' 
    public sub removeTo(email)
        if( Tos.exists(email) ) then
            call Tos.remove(email)
        end if
    end sub
    
    ' Subroutine: addCC
    ' 
    ' Adds an email to the carbon copy receivers collection.
    ' 
    ' Parameters:
    ' 
    '     (string) - email
    ' 
    public sub addCC(email)
        if( not Ccs.exists(email) ) then
            call Ccs.add(email, Ccs.count)
        end if
    end sub
    
    ' Subroutine: removeCC
    ' 
    ' Removes an email from the carbon copy receivers collection.
    ' 
    ' Parameters:
    ' 
    '     (string) - email
    ' 
    public sub removeCC(email)
        if( Ccs.exists(email) ) then
            call Ccs.remove(email)
        end if
    end sub
    
    ' Subroutine: addBCC
    ' 
    ' Adds an email to the blind carbon copy receivers collection.
    ' 
    ' Parameters:
    ' 
    '     (string) - email
    ' 
    public sub addBCC(email)
        if( not Bccs.exists(email) ) then
            call Bccs.add(email, Bccs.count)
        end if
    end sub
    
    ' Subroutine: removeBCC
    ' 
    ' Removes an email from the blind carbon copy receivers collection.
    ' 
    ' Parameters:
    ' 
    '     (string) - email
    ' 
    public sub removeBCC(email)
        if( Bccs.exists(email) ) then
            call Bccs.remove(email)
        end if
    end sub
    
    ' Subroutine: addAttachment
    ' 
    ' Adds a fully qualified file name to the attachments collection.
    ' 
    public sub addAttachment(path)
        if( not Attachments.exists(path) ) then
            call Attachments.add(path, Attachments.count)
        end if
    end sub
    
    ' Subroutine: removeAttachment
    ' 
    ' Removes a fully qualified file name to the attachments collection.
    ' 
    public sub removeAttachment(path)
        if( Attachments.exists(path) ) then
            call Attachments.remove(path)
        end if
    end sub
    
    ' Subroutine: send
    ' 
    ' Procedure send is used to send an Email.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Eml : set Eml = new Email
    ' set Eml.Adapter = new Email_Adapter_CDO
    ' Eml.from = "sender@domain.com"
    ' Eml.addTo("receiver@domain.com")
    ' Eml.subject = "Email subject"
    ' Eml.body = "Hello World"
    ' Eml.isHTML = false
    ' call Eml.send()
    ' set Eml = nothing
    ' 
    ' (end code)
    ' 
    public sub send() : call [_ε]
        call Adapter.send(Me)
    end sub
    
end class

%>
