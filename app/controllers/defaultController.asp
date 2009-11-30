<!--#include virtual="/lib/axe/base.asp"-->
<!--#include virtual="/lib/axe/classes/kernel.class.asp"-->
<!--#include virtual="/lib/axe/singletons.initialize.asp"-->
<!--#include virtual="/app/models/defaultModel.asp"-->
<%

' File: defaultController.asp
' 
' ASP Xtreme Evolution after install defaultController. *You should replace it with your own defaultController.*
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
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2007
' 
class DefaultController
    
    public sub defaultAction()
        dim Model, Xml, Xslt
        
        set Model = new DefaultModel
        set Xml = Model.introduce()
        set Model = nothing
        
        Session("this").add "title", "ASP Xtreme Evolution"
        Session("this").add "h1", "ASP Xtreme Evolution"
        
        set Xslt = Core.str2xml(Core.computeView())
        
        Session("this").add "Output.xml", Xml.xml
        Session("this").add "Output.xslt", Xslt.xml
        Session("this").add "Output.value" , Core.indentedTransform(Xml, Xslt, Application("Xslt.xhtml"), "  ")
        
        set Xslt = nothing
        set Xml = nothing
    end sub
    
    public sub another()
        dim Model, sHtml
        
        Session("this").add "h1", "I'm just anotherView created by 長尾 (Nagao)"
        
        set Model = new DefaultModel
        
        Session("this").add "defaultController.source", Model.read("controller", "defaultController")
        Session("this").add "defaultModel.source", Model.read("model", "defaultModel")
        Session("this").add "defaultView.source", Model.read("view", "defaultView")
        Session("this").add "anotherView.source", Model.read("view", "anotherView")
        
        set Model = nothing
        
        Session("view") = "anotherView"
        sHtml = Core.computeView()
        
        Session("this").add "Output.xml", "There isn't a xml layer."
        Session("this").add "Output.xslt", "No xslt layer either."
        Session("this").add "Output.value", sHtml
    end sub
    
end class



dim Controller : set Controller = new DefaultController
select case Session("action")
    
    case "defaultAction"
        call Controller.defaultAction()
    
    case "another"
        call Controller.another()
    
    case else
        Core.addError("Action '" & Session("action") & "' is not available at controller '" & Session("controller") & "'")
    
end select
set Controller = nothing

%>
<!--#include virtual="/lib/axe/singletons.finalize.asp"-->
