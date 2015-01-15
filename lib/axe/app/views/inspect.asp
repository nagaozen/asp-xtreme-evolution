<%

' File: inspect.asp
'
' This view provide resources for developers to analyse the current data source,
' transformation and output which is being sent to the user.
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
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ August 2013
'

%>
<!--#include virtual="/lib/axe/singletons.initialize.asp"-->
<!DOCTYPE html>
<html lang="en">
	<head>
		<meta charset="UTF-8">
		<meta name="robots" content="NOINDEX,NOFOLLOW"/>
		<meta name="description" content="ASP Xtreme Evolution : Inspect View">
		<meta name="keywords" content="axe, framework, inspect, view">
		<meta name="author" content="Evolved, Codrops">
		<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
		<title>ASP Xtreme Evolution : Inspect</title>
		<link rel="shortcut icon" href="/lib/axe/assets/inspect/favicon.ico">
		<link rel="stylesheet" type="text/css" href="/lib/axe/assets/inspect/css/normalize.css">
		<link rel="stylesheet" type="text/css" href="/lib/axe/assets/inspect/css/default.css">
		<link rel="stylesheet" type="text/css" href="/lib/highlight/styles/solarized_light.css">
	</head>
	<body>
		<div class="container">
			<nav>
				<ul class="nav">
					<li><a href="#" class="icon-logo">i</a></li>
					<li class='<%= iif( Request.Cookies("inspect_last_visible") = "inspect-sv-"    , "selected", "" ) %>'><a class="icon-cog"      id="inspect-sv-anchor"     href="#" onclick="dE(this, 'inspect-sv-');return false;">Server Variables</a></li>
					<li class='<%= iif( Request.Cookies("inspect_last_visible") = "inspect-ds-"    , "selected", "" ) %>'><a class="icon-database" id="inspect-ds-anchor"     href="#" onclick="dE(this, 'inspect-ds-');return false;">Data Source</a></li>
					<li class='<%= iif( Request.Cookies("inspect_last_visible") = "inspect-tpl-"   , "selected", "" ) %>'><a class="icon-beaker"   id="inspect-tpl-anchor"    href="#" onclick="dE(this, 'inspect-tpl-');return false;">Transformation</a></li>
					<li class='<%= iif( Request.Cookies("inspect_last_visible") = "inspect-output-", "selected", "" ) %>'><a class="icon-doc"      id="inspect-output-anchor" href="#" onclick="dE(this, 'inspect-output-');return false;">Output</a></li>
				</ul>
			</nav>
			<article class="main">
				<header class="clearfix">
					<span>ASP Xtreme Evolution</span>
					<h1>Inspector</h1>
				</header>
				<section id="inspect-sv-container">
					<h2>Server Variables</h2>
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
	Response.write "							<tr class=""" & sClass & """>" & vbNewLine
	Response.write "								<th>" & sVariable & "</th>" & vbNewLine
	Response.write "								<td>" & Request.serverVariables(sVariable) & "</td>" & vbNewLine
	Response.write "							</tr>" & vbNewLine
	i = i + 1
next

%>
						</tbody>
					</table>
				</section>
				<section id="inspect-ds-container">
					<h2>Data Source</h2>
					<pre><code><% Response.write Core.printerFriendlyCode( Session("this").item("Output.xml") ) %></code></pre>
				</section>
				<section id="inspect-tpl-container">
					<h2>Transformation</h2>
					<pre><code><% Response.write Core.printerFriendlyCode( Session("this").item("Output.xslt") ) %></code></pre>
				</section>
				<section id="inspect-output-container">
					<h2>Output</h2>
					<pre><code><% Response.write Core.printerFriendlyCode( Session("this").item("Output.value") ) %></code></pre>
				</section>
			</article>
		</div>
		<script type="text/javascript" src="/lib/highlight/highlight.pack.js"></script>
		<script>
// <![CDATA[

var idCurrentVisible = '<%= iif( Request.Cookies("inspect_last_visible") <> "", Request.Cookies("inspect_last_visible"), "inspect-sv-" ) %>';
function dE(e, id) {
    if( idCurrentVisible == id + "container" ) { return; }

    removeClass(document.getElementById(idCurrentVisible + "anchor").parentNode, "selected");
    addClass(e.parentNode, "selected");

    document.getElementById(idCurrentVisible + "container").style.display = "none";
    document.getElementById(id + "container").style.display = "block";

    idCurrentVisible = id;
    createCookie("inspect_last_visible", id, 1);
}

hljs.initHighlightingOnLoad();

document.getElementById(idCurrentVisible + "container").style.display = "block";





// Element classes helpers
function hasClass(e, cls) {
	return e.className.match(new RegExp("(^|\\s)" + cls + "(?:\\s|$)"));
}

function addClass(e, cls) {
	if(!this.hasClass(e,cls)) e.className += ' ' + cls;
}

function removeClass(e, cls) {
	if(hasClass(e,cls)) {
		var reg = new RegExp("(^|\\s)" + cls + "(?:\\s|$)");
		e.className = e.className.replace(reg, ' ');
	}
}

// Cookies helpers
function createCookie(name, value, days) {
	if(days) {
		var date = new Date();
		date.setTime(date.getTime()+(days*24*60*60*1000));
		var expires = "; expires="+date.toGMTString();
	}
	else var expires = "";
	document.cookie = name+"="+value+expires+"; path=/";
}

function readCookie(name) {
	var nameEQ = name + "=";
	var ca = document.cookie.split(';');
	for(var i=0;i < ca.length;i++) {
		var c = ca[i];
		while (c.charAt(0)==' ') c = c.substring(1,c.length);
		if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length);
	}
	return null;
}

function eraseCookie(name) {
	createCookie(name,"",-1);
}

// ]]>
		</script>
	</body>
</html>
<!--#include virtual="/lib/axe/singletons.finalize.asp"-->
<!--#include virtual="/lib/axe/sessions.finalize.asp"-->
