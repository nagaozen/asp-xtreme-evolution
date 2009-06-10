<%

'+-----------------------------------------------------------------------------+
'|This file is part of ASP Xtreme Evolution.                                   |
'|Copyright (C) 2007, 2008 Fabio Zendhi Nagao                                  |
'|                                                                             |
'|ASP Xtreme Evolution is free software: you can redistribute it and/or modify |
'|it under the terms of the GNU Lesser General Public License as published by  |
'|the Free Software Foundation, either version 3 of the License, or            |
'|(at your option) any later version.                                          |
'|                                                                             |
'|ASP Xtreme Evolution is distributed in the hope that it will be useful,      |
'|but WITHOUT ANY WARRANTY; without even the implied warranty of               |
'|MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                |
'|GNU Lesser General Public License for more details.                          |
'|                                                                             |
'|You should have received a copy of the GNU Lesser General Public License     |
'|along with ASP Xtreme Evolution.  If not, see <http://www.gnu.org/licenses/>.|
'+-----------------------------------------------------------------------------+

' File: error.asp
' 
' When a error occurs processing this framework, it will be storage in an array.
' This errors array is then verified and if something is there, this view will
' be triggered instead of the current requested view. It can also be used as the
' default 500;100 Internal Server Error.
' 
' About:
' 
'   Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2007

%>
<!--#include virtual="/app/core/base.asp"-->
<%

if Response.buffer then
    Response.clear
    Response.status = "500 Internal Server Error"
    Response.expires = 0
end if

%>
<!--#include virtual="/app/core/lib/kernel.class.asp"-->
<!--#include virtual="/app/core/singletons.initialize.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html
    xmlns="http://www.w3.org/1999/xhtml"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xml:lang="en">
    <head>
        <title>ASP Xtreme Evolution :: Runtime Error</title>
        <link rel="icon" type="image/png" href="/assets/default/img/favicon.png" />
        <link rel="stylesheet" type="text/css" media="screen" href="/assets/default/css/reset-fonts-2.3.1.css" />
        <link rel="stylesheet" type="text/css" media="screen" href="/assets/default/css/asp-xtreme-evolution.css" />
    </head>
    <body id="errors">
        <div id="container">
            <div id="container-hd">
                <h1>Framework Runtime Error</h1>
                <p>An error occurred processing the page you requested. Please see the details below for more information.</p>
            </div>
            <div id="container-bd">
                <table>
                    <tbody>
<%

' Framework Error
dim len, i, sClass
if(isArray(Session("errv"))) then
    len = uBound(Session("errv"))
    for i = 0 to len - 1
        if( i mod 2 = 0 ) then
            sClass = "even"
        else
            sClass = "odd"
        end if
        Response.write "<tr class=""" & sClass & """><td>" & Session("errv")(i) & "</td></tr>" & vbNewLine
    next
end if

' 500;100 : Internal Server Error - ASP Error
dim AspError
set AspError = Server.getLastError()
if((AspError.description <> "") AND (AspError.file <> "") AND (AspError.line > 0)) then
    if( i mod 2 = 0 ) then
        sClass = "even"
    else
        sClass = "odd"
    end if
    Response.write "<tr class=""" & sClass & """><td><b>" & AspError.description & "</b> @ " & AspError.file & ", line " & AspError.line & "</td></tr>" & vbNewLine
end if
set AspError = nothing

%>
                    </tbody>
                </table>
            </div>
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
<!--#include virtual="/app/core/singletons.finalize.asp"-->
<!--#include virtual="/app/core/sessions.finalize.asp"-->
