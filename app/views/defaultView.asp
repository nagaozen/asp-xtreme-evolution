<!--#include virtual="/lib/axe/base.asp"-->
<!--#include virtual="/lib/axe/classes/kernel.class.asp"-->
<!--#include virtual="/lib/axe/singletons.initialize.asp"-->
<!--#include virtual="/lib/axe/shuttle.unload.initialize.asp"-->
<%

' File: defaultView.asp
' 
' ASP Xtreme Evolution after install defaultView. This view uses the prototyped 
' DefaultModel.introduce() XML output to display the content using XSLT - which 
' is our recommendation. *You should replace it with your own defaultView.*
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
<xsl:transform version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
    <xsl:output method="xml" version="1.0" encoding="UTF-8" />
    <xsl:template match="/">
        <html xmlns="http://www.w3.org/1999/xhtml" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xml:lang="en">
            <head>
                <title><%= Session("this").item("title") %></title>
                <meta http-equiv="content-type" content="text/html; charset=UTF-8" />
                <link rel="icon" type="image/png" href="/lib/axe/assets/img/favicon.png" />
                <link rel="stylesheet" type="text/css" media="screen" href="/lib/axe/assets/css/reset-fonts-2.3.1.css" />
                <link rel="stylesheet" type="text/css" media="screen" href="/lib/axe/assets/css/asp-xtreme-evolution.css" />
            </head>
            <body id="defaultView">
                <div id="container">
                    <div id="container-hd">
                        <h1><%= Session("this").item("h1") %></h1>
                    </div>
                    <div id="container-bd">
                        <xsl:copy-of select="messages/message/*" />
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
    </xsl:template>
</xsl:transform>
<!--#include virtual="/lib/axe/shuttle.unload.finalize.asp"-->
<!--#include virtual="/lib/axe/singletons.finalize.asp"-->
