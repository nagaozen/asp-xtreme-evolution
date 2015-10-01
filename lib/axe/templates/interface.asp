<%

' File: interface.asp
' 
' ${2:description}
' 
' About:
' 
'     - Written by ${3:author} <http://${4:url}> @ ${5:month} ${6:year}
' 
class ${1:Name}_Interface' extends Interface
    
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
    
    $0
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        set Parent = new Interface
        Parent.requireds = array()
    end sub
    
    private sub Class_terminate()
        set Parent = nothing
    end sub
    
end class

%>
