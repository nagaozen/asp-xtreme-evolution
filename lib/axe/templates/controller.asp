<!--#include virtual="/lib/axe/base.asp"-->
<!--#include virtual="/lib/axe/classes/kernel.asp"-->
<!--#include virtual="/lib/axe/singletons.initialize.asp"-->
<%

' File: ${1:name}.asp
' 
' ${2:Description}
' 
' About:
' 
'     - Written by ${3:author} <http://${4:url}> @ ${5:month} ${6:year}
' 
class $<[1]: return $1.capitalize() >Controller
    
    public sub defaultAction()
        $0
    end sub
    
end class



dim Controller : set Controller = new $<[1]: return $1.capitalize() >Controller
select case Session("action")
    
    case "defaultAction"
        call Controller.defaultAction()
    
    case else
        call Core.exception("Action '{0}' is not available at controller '{1}'", array(Session("action"), Session("controller")))
    
end select
set Controller = nothing

%>
<!--#include virtual="/lib/axe/singletons.finalize.asp"-->
