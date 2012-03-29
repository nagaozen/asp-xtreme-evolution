<script language="VBScript" runat="server">

' File: customevent.asp
' 
' AXE(ASP Xtreme Evolution) events factory.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2007-2011 Fabio Zendhi Nagao
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



' Class: CustomEvent
' 
' This class implements a generic way to work with Custom Events in ASP. It's
' the ASP Xtreme Evolution base for an Event-driven programming approach.
' 
' Known issues:
' 
'   - As you can notice by the first example, CustomEvent doesn't work well with
'     the VBScript standard Class events "Class_initialize" and "Class_terminate"
'   - It's odd but getref() is unable to get the reference of a method or class
'     procedure.
' 
' About:
' 
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ January 2009
' 
class CustomEvent
    
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
    
    ' Property: Owner
    ' 
    ' description
    ' 
    ' Contains:
    ' 
    '     (Class) - Should be set with the reference of the Caller.
    ' 
    public Owner
    
    private Handlers
    
    ' Property: Arguments
    ' 
    ' Arguments to be passed to the handlers.
    ' 
    ' Contains:
    ' 
    '     (Scripting.Dictionary) - A collection with the callback function arguments
    ' 
    public Arguments
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        set Handlers  = Server.createObject("Scripting.Dictionary")
        set Arguments = Server.createObject("Scripting.Dictionary")
    end sub
    
    private sub Class_terminate()
        Handlers.removeAll()
        Arguments.removeAll()
        set Handlers  = nothing
        set Arguments = nothing
    end sub
    
    ' Subroutine: addHandler
    ' 
    ' Adds a handler.
    ' 
    public sub addHandler(fn)
        select case typename(fn)
            case "JScriptTypeInfo"
                set Handlers.item(fn.toString()) = fn
            case else
                set Handlers.item(fn) = getRef(fn)
        end select
    end sub
    
    ' Subroutine: removeHandler
    ' 
    ' description
    ' 
    public sub removeHandler(fn)
        set Handlers.item(fn) = nothing
        Handlers.remove(fn)
    end sub
    
    ' Subroutine: fire
    ' 
    ' Fires all handlers attached to this event.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' class ClassWithEvents
    '     public classType
    '     public classVersion
    '     
    '     public onLoad
    '     public onUnload
    '     public onComplimentBefore
    '     public onComplimentAfter
    '     
    '     private sub Class_initialize()
    '         classType    = typename(Me)
    '         classVersion = "1.0.0.0"
    '         
    '         set onLoad = new CustomEvent : set onLoad.Owner = Me
    '         set onUnload = new CustomEvent : set onUnload.Owner = Me
    '         set onComplimentBefore = new CustomEvent : set onComplimentBefore.Owner = Me
    '         set onComplimentAfter = new CustomEvent : set onComplimentAfter.Owner = Me
    ' 
    '         onComplimentBefore.Arguments.item("firstname") = "Fabio"
    '         onComplimentBefore.Arguments.item("lastname") = "Nagao"
    '         onComplimentBefore.Arguments.item("nickname") = "nagaozen"
    '         call onLoad.fire()
    '     end sub
    '     
    '     private sub Class_terminate()
    '         call onUnload.fire()
    '         
    '         set onLoad = nothing
    '         set onUnload = nothing
    '         set onComplimentBefore = nothing
    '         set onComplimentAfter = nothing
    '     end sub
    '     
    '     public function compliment()
    '         call onComplimentBefore.fire()
    '         Response.write("Method compliment called." & vbNewline)
    '         call onComplimentAfter.fire()
    '     end function
    '     
    ' end class
    ' 
    ' 
    ' 
    ' sub ev_onLoad(ev)
    '     Response.write("Event onLoad has been fired" & vbNewline)
    ' end sub
    ' 
    ' sub ev_onUnLoad(ev)
    '     Response.write("Event onUnLoad has been fired" & vbNewline)
    ' end sub
    ' 
    ' sub ev_onComplimentBefore(ev)
    '     Response.write("Event onComplimentBefore has been fired. I was really expecting this method to say: 'Hello World " & ev.Arguments.item("firstname") & " " & ev.Arguments.item("lastname") & " (" & ev.Arguments.item("nickname") & ")'" & vbNewline)
    ' end sub
    ' 
    ' sub ev_onComplimentAfter(ev)
    '     Response.write("Event onComplimentAfter has been fired" & vbNewline)
    ' end sub
    ' 
    ' dim CwE : set CwE = new ClassWithEvents
    ' call CwE.onLoad.addHandler("ev_onLoad")
    ' call CwE.onUnLoad.addHandler("ev_onUnLoad")
    ' call CwE.onComplimentBefore.addHandler("ev_onComplimentBefore")
    ' call CwE.onComplimentBefore.addHandler(lambda("function(ev){ Response.write('Another handler attached to Event onComplimentAfter. This one is using a lambda function -- Yes, ' + ev.Arguments.item('firstname') + ' ' + ev.Arguments.item('lastname') + ' (' + ev.Arguments.item('nickname') + ') has implemented it for Classic ASP.\r\n') }"))
    ' call CwE.onComplimentAfter.addHandler("ev_onComplimentAfter")
    ' CwE.compliment()
    ' set CwE = nothing
    ' 
    ' Response.write("As you can see, this nothing @ line 74 doesn't work as expected. onLoad doesn't work either. But for user methods and procedures the CustomEvent works fine." & vbNewLine)
    ' 
    ' (end code)
    ' 
    public sub fire()
        dim fn : for each fn in Handlers
            Handlers.item(fn)(Me)
        next
    end sub
    
    ' Function: revealArguments
    ' 
    ' Reveals the event arguments.
    ' 
    ' Returns:
    ' 
    '     (string) - A list of the available event arguments
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' ' Using the class defined in the example above
    ' dim CwE : set CwE = new ClassWithEvents
    ' Response.write(CwE.onComplimentBefore.revealArguments())
    ' set CwE = nothing
    ' 
    ' (end code)
    ' 
    public function revealArguments()
        revealArguments = ""
        dim arg : for each arg in Arguments
            revealArguments = revealArguments & ", " & arg
        next
        revealArguments = mid(revealArguments, 3)
    end function
    
end class

</script>
