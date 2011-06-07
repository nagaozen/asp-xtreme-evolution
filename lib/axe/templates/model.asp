<%

' File: ${1:name}Model.asp
' 
' ${2:description}
' 
' About:
' 
'     - Written by ${3:author} <http://${4:url}> @ ${5:month} ${6:year}
' 
class $<[1]: return $1.capitalize() >Model
    
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
    ' Interface implementation
    ' 
    ' Contains:
    ' 
    '     (Interface) - Adapter implementing Interface
    ' 
    public Adapter
    
    ' Subroutine: [_ε]
    ' 
    ' {private} Checks for an adapter assignment.
    ' 
    private sub [_ε]
        if( isEmpty(Adapter) ) then
            Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Missing an Adapter."
        end if
    end sub
    
    $0
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: toXML
    ' 
    ' Returns a XML representation of this model.
    ' 
    ' Parameters:
    ' 
    '     (string) - root node of the recursion
    ' 
    ' Returns:
    ' 
    '     (string) - XML representation
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Model : set Model = new $<[1]: return $1.capitalize() >Model
    ' '.
    ' '. initialize it
    ' '.
    ' Response.write(Model.toXML())
    ' set Model = nothing
    ' 
    ' (end code)
    ' 
    public function toXML()
        
    end function
    
    ' Function: toJSON
    ' 
    ' Returns a JSON representation of this model.
    ' 
    ' Parameters:
    ' 
    '     (string) - root node of the recursion
    ' 
    ' Returns:
    ' 
    '     (string) - JSON representation
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Model : set Model = new $<[1]: return $1.capitalize() >Model
    ' '.
    ' '. initialize it
    ' '.
    ' Response.write(Model.toJSON())
    ' set Model = nothing
    ' 
    ' (end code)
    ' 
    public function toJSON()
        
    end function
    
end class

%>

