<%

' File: error.asp
' 
' When a error occurs processing this framework, it will be storage in an array.
' This errors array is then verified and if something is there, this view will
' be triggered instead of the current requested view. It can also be used as the
' default 500;100 Internal Server Error.
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

%>
<!--#include virtual="/lib/axe/base.asp"-->
<%

if Response.buffer then
    Response.clear
    Response.status = "500 Internal Server Error"
    Response.contentType = "text/html"
    Response.expires = 0
end if

' 500;100 : Internal Server Error - ASP Error
dim AspError, category, message
set AspError = Server.getLastError()
if((AspError.description <> "") and (AspError.file <> "") and (AspError.line > 0)) then
    category = strsubstitute( _
        "{0}{1} error (0x{2})", _
        array( _
            AspError.category, _
            iif(AspError.aspCode > "", Server.htmlEncode(", " & AspError.aspCode), ""), _
            Hex(AspError.number) _
        ) _
    )
    
    message = strsubstitute( _
        "<p><strong>{0}</strong> @ <code>{1}</code>{2}{3}</p>{4}", _
        array( _
            AspError.description, _
            AspError.file, _
            iif(AspError.line > 0, (", line: <code>" & AspError.line & "</code>"), ""), _
            iif(AspError.column > 0, (", column: " & AspError.column & "</code>"), ""), _
            iif(AspError.source > "", "<pre><code>" & Server.htmlEncode(AspError.source) & "</code></pre>", "") _
        ) _
    )
end if
set AspError = nothing

%>
<!--#include virtual="/lib/axe/classes/kernel.asp"-->
<!--#include virtual="/lib/axe/singletons.initialize.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html
    xmlns="http://www.w3.org/1999/xhtml"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xml:lang="en"
>
    <head>
        <title>ASP Xtreme Evolution :: runtime error</title>
        <meta http-equiv="content-type" content="text/html; charset=UTF-8"></meta>
        <link rel="icon" type="image/png" href="/lib/axe/assets/img/favicon.png" />
        <link rel="stylesheet" type="text/css" media="screen" href="/lib/axe/assets/css/reset-fonts-2.3.1.css" />
        <link rel="stylesheet" type="text/css" media="screen" href="/lib/axe/assets/css/asp-xtreme-evolution.css" />
    </head>
    <body id="errors">
        <div id="container">
            <div id="container-hd">
                <h1><%= category %></h1>
                <p>An error occurred processing the page ( <code>'<%= Request.ServerVariables("SCRIPT_NAME") & "' [" & Request.ServerVariables("REQUEST_METHOD") %>]</code> ) you requested. Please see the details below for more information.</p>
            </div>
            <div id="container-bd"><%= message %></div>
            <div id="container-ft">
                <hr />
                <ul>
                    <li><a href="http://validator.w3.org/check?uri=referer"><img src="http://www.w3.org/Icons/valid-xhtml11" alt="Valid XHTML 1.1" height="31" width="88" /></a></li>
                    <li><a href="http://jigsaw.w3.org/css-validator/check/referer"><img src="http://jigsaw.w3.org/css-validator/images/vcss" alt="Valid CSS!" /></a></li>
                    <li><a href="http://www.w3.org/WAI/WCAG1AA-Conformance" title="Explanation of Level Double-A Conformance"><img height="32" width="88" src="http://www.w3.org/WAI/wcag1AA" alt="Level Double-A conformance icon, W3C-WAI Web Content Accessibility Guidelines 1.0" /></a></li>
                </ul>
            </div>
        </div>
    </body>
</html>
<!--#include virtual="/lib/axe/singletons.finalize.asp"-->
<!--#include virtual="/lib/axe/sessions.finalize.asp"-->
