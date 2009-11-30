<!--#include virtual="/lib/axe/base.asp"-->
<!--#include virtual="/lib/axe/classes/kernel.class.asp"-->
<!--#include virtual="/lib/axe/singletons.initialize.asp"-->
<%

' File: ${1:controller}.asp
' 
' ${2:Description}
' 
' About:
' 
'     - Written by ${3:author} <http://${4:url}> @ {5:month} {6:year}
' 
class ${1:Template}Controller
    
    public sub defaultAction()
        $0
    end sub
    
end class



dim Controller : set Controller = new $1Controller
select case Session("action")
    
    case "defaultAction"
        call Controller.defaultAction()
    
    case else
        Core.addError("Action '" & Session("action") & "' is not available at controller '" & Session("controller") & "'")
    
end select
set Controller = nothing

%>
<!--#include virtual="/lib/axe/singletons.finalize.asp"-->
