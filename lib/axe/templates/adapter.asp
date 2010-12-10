<%

' Class: ${1:name}
' 
' ${3:description}
' 
' About:
' 
'     - Written by ${4:author} <http://${5:url}> @ ${6:month} ${7:year}
' 
class ${1:name}' implements ${2:Interface}
    
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
    
    $0
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0"
        
        set Interface = new $2
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
    
end class

%>
