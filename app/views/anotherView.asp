<!--#include virtual="/app/core/base.asp"-->
<!--#include virtual="/app/core/lib/kernel.class.asp"-->
<!--#include virtual="/app/core/singletons.initialize.asp"-->
<!--#include virtual="/app/core/shuttle.unload.initialize.asp"-->
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

' anotherView.asp
' 
' ASP Xtreme Evolution View: Default for Simple XHTML.
' 
' Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ December 2007

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xml:lang="en">
    <head>
        <title>ASP Xtreme Evolution: anotherView</title>
        <link rel="icon" type="image/png" href="/assets/default/img/favicon.png" />
        <link rel="stylesheet" type="text/css" media="screen" href="/assets/default/css/reset-fonts-2.3.1.css" />
        <link rel="stylesheet" type="text/css" media="screen" href="/assets/default/css/asp-xtreme-evolution.css" />
        <link rel="stylesheet" type="text/css" media="screen" href="/assets/default/css/asp-xtreme-evolution-syntax.css" />
    </head>
    <body>
        <div id="container">
            <div id="container-hd">
                <h1><%= Session("this").item("h1") %></h1>
            </div>
            <div id="container-bd">
                <p>As you can see, just changing <code>&lt;action&gt;</code> the http://www.domain.com/<code>&lt;Controller&gt;</code>/<code>&lt;action&gt;</code>/<code>&lt;arguments[]&gt;</code> pattern in the address bar, you call the <code>&lt;action&gt;</code> method of <code>&lt;Controller&gt;</code> passing the <code>&lt;arguments[]&gt;</code> to it. This means you are ready for a spectacular experince with friendly URLs and MVC approach. Check below the easy source code of the standard pages of the framework:</p>
                <h2>defaultController</h2>
                <p>A Controller processes and responds to events, typically user actions, and may invoke changes on the model.</p>
                <pre name="code" class="asp"><%= Session("this").item("defaultController.source") %></pre>
                <h2>defaultModel</h2>
                <p>A Model is the domain-specific representation of the information on which the application operates. Domain logic adds meaning to raw data (e.g., calculating whether today is the user's birthday, or the totals, taxes, and shipping charges for shopping cart items).</p>
                <pre name="code" class="asp"><%= Session("this").item("defaultModel.source") %></pre>
                <h2>defaultView</h2>
                <p>A View renders the model into a form suitable for interaction, typically a user interface element. As you can see here, multiple views can exist for different purposes in the same Controller. This view is attached with an action which enables the AXE(ASP Xtreme Evolution) full power, returning XML representation in the Model and using XSLT to dinamically generate the output.</p>
                <pre name="code" class="xml"><%= Session("this").item("defaultView.source") %></pre>
                <h2>anotherView</h2>
                <p>If you felt confused by the XML-XSLT pattern, you can use Models to bring raw data from databases and Views as HTML. It also tastes good!</p>
                <pre name="code" class="xml"><%= Session("this").item("anotherView.source") %></pre>
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
        <script type="text/javascript" src="/assets/default/js/dp.SyntaxHighlighter/Scripts/shBrushAsp.js"></script>
        <script type="text/javascript" src="/assets/default/js/dp.SyntaxHighlighter/Scripts/shBrushXml.js"></script>
        <script type="text/javascript">
            dp.SyntaxHighlighter.ClipboardSwf = "/assets/default/js/dp.SyntaxHighlighter/Scripts/clipboard.swf";
            dp.SyntaxHighlighter.HighlightAll("code");
        </script>

    </body>
</html>
<!--#include virtual="/app/core/shuttle.unload.finalize.asp"-->
<!--#include virtual="/app/core/singletons.finalize.asp"-->
