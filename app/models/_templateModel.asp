<%

' File: ${1:Name}Model.asp
' 
' ${2:Description}
' 
' About:
' 
'     - Written by ${3:author} <http://${4:url}> @ {5:month} {6:year}
' 
class $1Model
    
    $0
    
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
    ' dim Model : set Model = new $1Model
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
    ' dim Model : set Model = new $1Model
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
    
    ' Function: toHTML
    ' 
    ' Returns a HTML representation of this model.
    ' 
    ' Parameters:
    ' 
    '     (string) - root node of the recursion
    ' 
    ' Returns:
    ' 
    '     (string) - HTML representation
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Model : set Model = new $1Model
    ' '.
    ' '. initialize it
    ' '.
    ' Response.write(Model.toHTML())
    ' set Model = nothing
    ' 
    ' (end code)
    ' 
    public function toHTML()
        
    end function
    
end class

%>
