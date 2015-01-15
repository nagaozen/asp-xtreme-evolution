<%

' File: interface.class.asp
'
' AXE(ASP Xtreme Evolution) interfaces factory.
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



' Interface: Interface
'
' All interfaces should extend this class. This class enforces a strong, but not
' perfect, binding over an interface and it's adapters.
'
' About:
'
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
'
class Interface

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

    ' Property: Implementation
    '
    ' Interface implementation.
    '
    ' Contains:
    '
    '     (Adapter) - Interface implementation
    '
    public Implementation

    ' Property: requireds
    '
    ' Attributes, Methods and Procedures to be checked.
    '
    ' Contains:
    '
    '     (string[]) - Attributes, Methods and Procedures to be checked
    '
    public requireds

    private sub Class_initialize()
        classVersion = "1.0.0.0"
        classType    = typename(Me)
    end sub

    private sub Class_terminate()
    end sub

    ' Function: check
    '
    ' Performs a check in the Adapter for the Interface specifications.
    '
    ' Returns:
    '
    '     (boolean) - true, if adapter passes the test; false, otherwise.
    '
    ' Example:
    '
    ' (start code)
    '
    ' class Template_Adapter_Interface' extends Interface
    '
    '     ' --[ Inheritance ]---------------------------------------------------------
    '     public Parent
    '
    '     public property set Implementation(I)
    '         set Parent.Implementation = I
    '     end property
    '
    '     public property get Implementation
    '         set Implementation = Parent.Implementation
    '     end property
    '
    '     public property get requireds
    '         requireds = Parent.requireds
    '     end property
    '
    '     public function check()
    '         check = Parent.check()
    '     end function
    '
    '     ' --[ Interface definition ]------------------------------------------------
    '     public classVersion
    '     public classType
    '
    '     private sub Class_initialize()
    '         classVersion = "1.0.0.0"
    '         classType    = typename(Me)
    '
    '         set Parent = new Interface
    '         Parent.requireds = array("prop", "load", "save", "drop")
    '     end sub
    '
    '     private sub Class_terminate()
    '         set Parent = nothing
    '     end sub
    '
    '     public prop
    '
    '     public sub load()
    '     end sub
    '
    '     public sub save()
    '     end sub
    '
    '     public sub drop()
    '     end sub
    '
    ' end class
    '
    '
    '
    ' class Template_Adapter_Media' implements Template_Adapter_Interface
    '
    '     ' --[ Interface ]-----------------------------------------------------------
    '     public Interface
    '
    '     ' --[ Adapter definition ]--------------------------------------------------
    '     public classVersion
    '     public classType
    '
    '     private sub Class_initialize()
    '         classVersion = "1.0.0.0"
    '         classType    = typename(Me)
    '
    '         set Interface = new Template_Adapter_Interface
    '         set Interface.Implementation = Me
    '         if(not Interface.check()) then
    '             Core.addError( _
    '                 strsubstitute( _
    '                     "'{0}' is a bad interface implementation of '{1}'", _
    '                     array(classType, typename(Interface)) _
    '                 ) _
    '             )
    '         end if
    '     end sub
    '
    '     private sub Class_terminate()
    '         set Interface.Implementation = nothing
    '         set Interface = nothing
    '     end sub
    '
    '     public prop
    '
    '     public sub load()
    '     end sub
    '
    '     public sub save()
    '     end sub
    '
    '     public sub drop()
    '     end sub
    '
    ' end class
    '
    '
    '
    ' dim A : set A = new Template_Adapter_Media
    ' Response.write( A.Interface.check() )' prints true
    ' set A = nothing
    '
    ' (end code)
    '
    public function check()
        on error resume next

        check = true

        dim i : for i = 0 to ubound( requireds )
            dim r : r = typename( eval( "Implementation." & requireds(i) ) )
            if( _
                Err.number <> 0 and _
                Err.number <> 5 and _
                Err.number <> 450 _
            ) then
                check = false
                Err.clear()
                exit for
            end if
        next

        on error goto 0

        if(not check) then
            Err.raise 17, "Evolved AXE runtime error", strsubstitute( _
                "Can't perform requested operation. '{0}' is a bad interface implementation of '{1}'", _
                array(typename(Implementation), typename(Implementation.Interface)) _
            )
        end if
    end function

end class

%>
