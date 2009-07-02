<%

'+-----------------------------------------------------------------------------+
'|This file is part of ASP Xtreme Evolution.                                   |
'|Copyright (C) 2007, 2009 Fabio Zendhi Nagao                                  |
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

' File: inspect.asp
' 
' This view provide resources for developers to analyse the current data source,
' transformation and output which is being sent to the user.
' 
' About:
' 
'   Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2007

%>
<!--#include virtual="/app/core/base.asp"-->
<!--#include virtual="/app/core/lib/kernel.class.asp"-->
<!--#include virtual="/app/core/singletons.initialize.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html
    xmlns="http://www.w3.org/1999/xhtml"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xml:lang="en">
    <head>
        <title>ASP Xtreme Evolution :: Inspect</title>
        <link rel="icon" type="image/png" href="/assets/default/img/favicon.png" />
        <link rel="stylesheet" type="text/css" media="screen" href="/assets/default/css/reset-fonts-2.3.1.css" />
        <link rel="stylesheet" type="text/css" media="screen" href="/assets/default/css/asp-xtreme-evolution.css" />
        <link rel="stylesheet" type="text/css" media="screen" href="/assets/default/css/asp-xtreme-evolution-syntax.css" />
    </head>
    <body id="inspect" class="frameworkView">
        <div id="container">
            <div id="container-hd">
                <h1>Inspecting :: /<% Response.write Session("controller") & "/" & Session("action") & "/" & join(Session("argv"), "/") %></h1>
                <ul>
                    <li><a id="inspect-sv-anchor" href="#" onclick="dE('inspect-sv-container');return false;">Server Variables</a></li>
                    <li><a id="inspect-xml-anchor" href="#" onclick="dE('inspect-xml-container');return false;">XML</a></li>
                    <li><a id="inspect-xslt-anchor" href="#" onclick="dE('inspect-xslt-container');return false;">XSLT</a></li>
                    <li><a id="inspect-output-anchor" href="#" onclick="dE('inspect-output-container');return false;">Output</a></li>
                </ul>
                <hr />
            </div>
            <div id="container-bd">
                <div id="inspect-sv-container">
                    <table>
                        <tbody>
<%

dim sVariable, i, sClass
i = 0
for each sVariable in Request.serverVariables
    if( i mod 2 = 0 ) then
        sClass = "even"
    else
        sClass = "odd"
    end if
    Response.write "                            <tr class=""" & sClass & """>" & vbNewLine
    Response.write "                                <th>" & sVariable & "</th>" & vbNewLine
    Response.write "                                <td>" & Request.serverVariables(sVariable) & "</td>" & vbNewLine
    Response.write "                            </tr>" & vbNewLine
    i = i + 1
next

%>
                        </tbody>
                    </table>
                </div>
                <div id="inspect-xml-container">
                    <pre name="code" class="xml"><% Response.write Core.printerFriendlyCode(Session("this").item("Output.xml")) %></pre>
                </div>
                <div id="inspect-xslt-container">
                    <pre name="code" class="xml"><% Response.write Core.printerFriendlyCode(Session("this").item("Output.xslt")) %></pre>
                </div>
                <div id="inspect-output-container">
                    <pre name="code" class="xhtml"><% Response.write Core.printerFriendlyCode(Session("this").item("Output.value")) %></pre>
                </div>
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
        <script type="text/javascript" src="/assets/default/js/dp.SyntaxHighlighter/Scripts/shCore.js"></script>
        <script type="text/javascript" src="/assets/default/js/dp.SyntaxHighlighter/Scripts/shBrushXml.js"></script>
        <script type="text/javascript">
            // <![CDATA[
            
            var idCurrentVisible = "inspect-sv-container";
            function dE(id) {
                if( idCurrentVisible == id ) { return; }
                document.getElementById(idCurrentVisible).style.display = "none";
                document.getElementById(id).style.display = "block";
                idCurrentVisible = id;
            }
            
            dp.SyntaxHighlighter.ClipboardSwf = "/assets/default/js/dp.SyntaxHighlighter/Scripts/clipboard.swf";
            dp.SyntaxHighlighter.HighlightAll("code");
            
            // ]]>
        </script>
    </body>
</html>
<!--#include virtual="/app/core/singletons.finalize.asp"-->
<!--#include virtual="/app/core/sessions.finalize.asp"-->
