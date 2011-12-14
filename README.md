ASP Xtreme Evolution
====================

The ASP Xtreme Evolution goal is to be a versatile MVC URL-Friendly base for Classic ASP applications with some additional features that are not ASP native. It should implement things that are common to most applications removing the pain of starting a new software and helping you to structure it so that you get things right from the beginning. Our key concepts are choice and freedom over limiting conventions, polyglotism, sustained quality, extensibility which we try to implement in a clean, maintainable and extensible way.

License
-------
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Lesser General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Lesser General Public License for more details.

You should have received a copy of the GNU Lesser General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
A little about the binary extensions
------------------------------------

First of all, ASP Xtreme Evolution is based in some extensions, but don't worry
it's all free or open source and already comes with the package. Although, if you
want to check the original projects, here's a list of their sites:

* [Microsoft Core XML Services (MSXML) 6.0](http://www.microsoft.com/downloads/details.aspx?FamilyID=993c0bcf-3bcf-4009-be21-27e85e1857b1&displaylang=en)
* [Ionics Isapi Rewrite Filter](http://iirf.codeplex.com/)
* [Crazy Beavers Imager Resizer and Zip](http://www.crazybeavers.se/downloads/)

Installation
------------

1. Unzip the entire zip package in your HD and upload it to your domain FTP.

2. If the domain is not registered in the server yet, register it in IIS pointing it's root to the folder where default.asp is located. Don't forget to check "Run scripts (such as ASP)".

3. If your server does not provide a default cgi-bin folder, create a virtual directory named "cgi-bin" in your domain root pointing to /lib/axe/bin. Don't forget to check "Execute (such as ISAPI applications or CGI)"; Otherwise, move the contents of /lib/axe/bin to your cgi-bin folder.

4. Give only "Write" permission to your /app/writables folder. It's ONLY WRITE! DO NOT let users to Read or specially EXECUTE anything.

5. Go to your site properties and in ISAPI Filters tab, click [Add]. Fill the box with the following data:  
   Filter name: IIRF  
   [Browse...] to /lib/axe/bin/IsapiRewrite4.dll  

6. Go to your Web Service Extensions and Add the binaries that comes with the package. They are located at /lib/axe/bin and their names are:
    * CB Image Resizer (Imager.dll)
    * CB Zip (CBZIP.exe)
    * Ionics ISAPI Rewriting Filter (IsapiRewrite4.dll)
    
    Notes:
    
    * Don't forget to check "Set extension status to Allowed".
    * If Active Server Pages is not allowed yet, allow it too.

7. If your server does not support MSXML 6.0 yet, install it. It's available at: /lib/axe/bin/msxml6.msi

8. Create an application pool for the application views.

9. Set /app/views to the application pool created in step 9.

10. That's it! Open a Web browser and type your domain. The "Welcome to ASP Xtreme Evolution" page should open.

Version 1.4.9.999α
------------------

### Fixes

Mmmm, I really don't have a track of the fixes 8'( but changed a lot of things 
for sure.

### Added

* routing system based on /app/config.xml
* mvc bootstrapping for controllers
* new logo and favicon
* /app/singletons.initialize.asp and /app/singletons.finalize.asp to manage application singletons.
* classType and classVersion properties in all classes
* new View
    - genericView provides an easy shortcut for you to send an entire view content.
* [new functions](/lib/axe/docs/files/lib/axe/base-asp.html)
    - axeInfo()
    - iif(expr, truepart, falsepart)
    - strsubstitute(template, replacements)
    - sanitize(value, placeholders, replacements)
    - isOdd(n)
    - isEven(n)
    - max(a)
    - min(a)
    - floor(n)
    - ceiling(n)
    - hex2dec(value)
    - dec2hex(value)
    - guid()
    - dump(variable)
    - lambda(fn)
    - Kernel.exception(message) (reason: framework error system is now based on the same of asp.dll)
    - [and a lot more basic math functions...](/lib/axe/docs/files/lib/axe/base-math-asp.html)
* new classes
    - [ACL](/lib/axe/docs/files/lib/axe/classes/Utilities/acl-asp.html)
    - [Atom](/lib/axe/docs/files/lib/axe/classes/Feeds/atom-asp.html) _(not implemented yet)_
    - [Auth](/lib/axe/docs/files/lib/axe/classes/Utilities/auth-asp.html)
    - [Base64](/lib/axe/docs/files/lib/axe/classes/Utilities/base64-asp.html)
    - [Color](/lib/axe/docs/files/lib/axe/classes/Utilities/color-asp.html)
    - [CSV](/lib/axe/docs/files/lib/axe/classes/Parsers/csv-asp.html) _(not implemented yet)_
    - [CustomEvent](/lib/axe/docs/files/lib/axe/classes/customevent-asp.html)
    - [Email](/lib/axe/docs/files/lib/axe/classes/Utilities/email-asp.html)
    - [Interface](/lib/axe/docs/files/lib/axe/classes/interface-asp.html)
    - [JSON2](/lib/axe/docs/files/lib/axe/classes/Parsers/json2-asp.html)
    - [JSONSchema](/lib/axe/docs/files/lib/axe/classes/Parsers/jsonschema-asp.html)
    - [List](/lib/axe/docs/files/lib/axe/classes/Utilities/list-asp.html)
    - [Logger](/lib/axe/docs/files/lib/axe/classes/Utilities/logger-asp.html)
    - [Markdown](/lib/axe/docs/files/lib/axe/classes/Parsers/markdown-asp.html)
    - [Mustache](/lib/axe/docs/files/lib/axe/classes/Parsers/mustache-asp.html)
    - [Orderly](/lib/axe/docs/files/lib/axe/classes/Parsers/orderly-asp.html)
    - [Paginator](/lib/axe/docs/files/lib/axe/classes/Utilities/paginator-asp.html)
    - [RSS](/lib/axe/docs/files/lib/axe/classes/Feeds/rss-asp.html) _(work in progress)_
    - [StringBuilder](/lib/axe/docs/files/lib/axe/classes/Utilities/stringbuilder-asp.html)
    - [Template](/lib/axe/docs/files/lib/axe/classes/Utilities/template-asp.html)
    - [Textile](/lib/axe/docs/files/lib/axe/classes/Parsers/textile-asp.html)
    - [Translator](/lib/axe/docs/files/lib/axe/classes/Utilities/translator-asp.html)
    - [UnitTest](/lib/axe/docs/files/lib/axe/classes/unittest-asp.html)
    - [XSession](/lib/axe/docs/files/lib/axe/classes/Utilities/xsession-asp.html)
* new services
    - [Akismet](/lib/axe/docs/files/lib/axe/classes/Services/akismet-asp.html)
    - [reCaptcha](/lib/axe/docs/files/lib/axe/classes/Services/recaptcha-asp.html)

### Changes

* Changed folders structure
    * Moved /assets/default to /lib/axe/assets
    * Moved /app/docs to /lib/axe/docs
    * Moved /app/bin to /lib/axe/bin
    * Moved /app/core/lib to /lib/axe/classes
    * Moved /app/core to /lib/axe
* Changed classes filename structure
    * renamed filename.class.asp to filename.asp
* Moved everything from ANSI to UTF-8. It's now possible to write 長尾 (nagao). _Note:_ changing the encode is easy as updating default.asp and global.asa
* Changed both Kernel.loadTextFile and Kernel.createFile from FSO to Stream for better UTF-8 support.
* updated classes
    - [JSON](/lib/axe/docs/files/lib/axe/classes/Parsers/json-asp.html) (more methods and fixes!)
* Updated error.asp (better coding, new behavior).
* defaultView.asp displays both INSTALL.md and CHANGES.md content
* INSTALL.md and CHANGES.md are now written in Markdown syntax.
* removed Application("Msxml.version") and fixed the version to 6.0.
* removed iif() from classes.
* better templates!
* better documentation! It's updated to use [NaturalDocs v1.4](http://www.naturaldocs.org/ "NaturalDocs"). Of course it still uses [GeSHi](http://qbnz.com/highlighter/ "GeSHi") and [Tidy](http://tidy.sourceforge.net/ "Tidy") too.
* [SOAP Toolkit](http://www.microsoft.com/downloads/details.aspx?familyid=c943c0dd-ceec-4088-9753-86f052ec8450&displaylang=en "SOAP Toolkit 3.0")
    - [Security Update](http://www.microsoft.com/DownLoads/details.aspx?FamilyID=23d18fd1-34be-4123-ba56-9be2d4be1b23&displaylang=en "SOAP Toolkit 3.0 Security Update")
* [IIRF](http://iirf.codeplex.com/ "Ionic's ISAPI Rewrite Filter")
    - updated to 2.1 (current latest stable)
    - enhanced URL-Rewriting
* a lot of other minor updates to make the Framework better ...

### Removed

* removed functions
    - Kernel.urlDecode(s) (reason: that function was not compatible with all types of Encoding.)
    - Kernel.htmlDecode(s) (reason: not required for Kernel.)
    - Kernel.sanitize(s) (reason: using base.asp one instead.)
    - Kernel.addError(sError) (reason: framework error system is now based on the same of asp.dll)
    - Kernel.hasErrors(saErrors) (reason: same as above)



Version 1.0.1.1
---------------

### Fixes

* Fixed the empty shuttle bug. This error was occuring everytime you call a view without adding keys to the Session("this")
* Fixed a bug in the JSON Class getChildNodes method which wasn't working for objects with depth >= 3. Thanks for Sven Neumann for pointing it.

### Changes

* Enhanced the rewriting rules (last one was writing logs and executing 3 tests for all request.).
* Enhanced standardization to fit W3C Level Double-A Conformance to Web Content Accessibility Guidelines 1.0.
* Added [object Array] detection in JSON Class. It's required because if you define ['one','two,three','four'] in the old one getElement returns the "one,two,three,four" string.



Version 1.0.1.0
---------------

### Fixes

* Fixed last slash bug for complete URL Rewrite (/Controller/action/args/).
* Fixed no special chars bug (no '%') in urlDecode.
* A misspell in Imager Class.
* Fixed the evil "operation timed out" error creating another Application pool for the view folder. [More info](http://support.microsoft.com/default.aspx?scid=kb;en-us;Q316451).
* To fix the "80020009" error change AspMaxRequestEntityAllowed parameter in the C:\WINDOWS\system32\inetsrv\MetaBase.xml from 200000 to something between [0, 1073741824] bytes and iisreset.
* Fixed Memory Leak in the Server View Requests.

### Added

* Model, View and Controller Templates.
* My standard favicon in the Framework pages.

### Changes

* Changed Controllers Standard Structure to create scopes per action.
* Created a new Welcome page (changes in the defaultView and defaultModel).
* Removed all NON-ASCII characters and changed the standard encode from UTF-8 to ASCII.



Version 1.0.0.0
---------------

* Initial Release
