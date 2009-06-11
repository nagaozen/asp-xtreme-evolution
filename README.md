ASP Xtreme Evolution
====================

The ASP Xtreme Evolution goal is to be a versatile MVC URL-Friendly base for Classic ASP applications with some additional features that are not ASP native. It should implement things that are common to most applications removing the pain of starting a new software and helping you to structure it so that you get things right from the beginning.

License
-------
﻿This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Lesser General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Lesser General Public License for more details.

You should have received a copy of the GNU Lesser General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
A little about the extensions
-----------------------------

First of all, ASP Xtreme Evolution is based in some extensions, but don't worry it's all free and everything already comes with the package. Although, if you want to check the original projects, here's a list of their sites:

* [Microsoft Core XML Services (MSXML) 6.0](http://www.microsoft.com/downloads/details.aspx?FamilyID=993c0bcf-3bcf-4009-be21-27e85e1857b1&displaylang=en)
* [Ionics Isapi Rewrite Filter](http://www.codeplex.com/IIRF/)
* [Crazy Beavers : Imager Resizer and Zip](http://www.crazybeavers.se/downloads/)

Installation
------------

1. Unzip the entire zip package in your HD and upload it to your domain FTP.

2. If the domain is not registered in the server yet, register it in IIS pointing it's root to the folder where default.asp is located. Don't forget to check "Run scripts (such as ASP)".

3. If your server does not provide a default cgi-bin folder, create a virtual directory named "cgi-bin" in your domain root pointing to /app/bin. Don't forget to check "Execute (such as ISAPI applications or CGI)"; Otherwise, move the contents of /app/bin to your cgi-bin folder.

4. Give "Write" permission to your /app/cache folder.

5. Go to your site properties and in ISAPI Filters tab, click [Add]. Fill the box with the following data:  
   Filter name: IIRF  
   [Browse...] to /app/bin/IsapiRewrite4.dll  

6. Yet in the site properties, go to the Custom Errors tab and look for 500;100 Default "Internal Server Error - ASP Error". Click [Edit...]  
   Message type: URL  
   URL: /app/views/error.asp  

7. Go to your Web Service Extensions and Add the binaries that comes with the package. They are located at /app/bin and their names are:
    * CB Image Resizer (Imager.dll)
    * CB Zip (CBZIP.exe)
    * Ionics ISAPI Rewriting Filter (IsapiRewrite4.dll)
    
    Notes:
    
    * Don't forget to check "Set extension status to Allowed".
    * If Active Server Pages is not allowed yet, allow it too.

8. If your server does not support MSXML 6.0 yet, install it. It's available at: /app/bin/msxml6.msi

9. Create an application pool for views.

10. Create an application for /app/views and set it application pool to the one created in step 9.

11. That's it! Open a Web browser and type your domain. The "Welcome!" message should open.
    * To test the inspect mode, just append: ?inspect=true at the end of your domain. Ex.: http://localhost/?inspect=true
    * To test the URL Rewriting Filter, try to access the /defaultController/another/ page. Ex.: http://localhost/defaultController/another/
    * To test the other extensions, try to check if Imager.dll is responding with ?Test=True. Ex.: http://localhost/cgi-bin/Imager.dll?Test=True. If it displays: "The specified file does not exist." it's working. If you are prompted to download the file it's not.

Version 1.1.0.0
---------------

### Fixes

Mmmm, I really didn't have a track of the fixes 8'( but changed a lot of things 
for sure.

### Added

* brand-new logo and favicon
* classType and classVersion properties in all classes
* new View
    - genericView provides an easy shortcut for you to send an entire view content.
* [new functions](/app/docs/files/app/core/base-asp.html)
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
    - [and a lot more basic math functions...](/app/docs/files/app/core/base-math-asp.html)
* new classes
    - [Base64](/app/docs/files/app/core/lib/Utilities/base64-class-asp.html)
    - [Color](/app/docs/files/app/core/lib/Utilities/color-class-asp.html)
    - [CustomEvent](/app/docs/files/app/core/lib/customevent-class-asp.html)
    - [Markdown](/app/docs/files/app/core/lib/Parsers/markdown-class-asp.html)
    - [Paginator](/app/docs/files/app/core/lib/Utilities/paginator-class-asp.html)
    - [StringBuilder](/app/docs/files/app/core/lib/Utilities/stringbuilder-class-asp.html)
    - [Template](/app/docs/files/app/core/lib/Utilities/template-class-asp.html)
    - [UnitTest](/app/docs/files/app/core/lib/unittest-class-asp.html)
* new services
    - [Akismet](/app/docs/files/app/core/lib/Services/akismet-class-asp.html)
    - [reCaptcha](/app/docs/files/app/core/lib/Services/recaptcha-class-asp.html)

### Changes

* Moved everything from ANSI to UTF-8. Note: changing the encode is easy as updating `default.asp` and `global.asa`
* updated classes
    - [JSON](/app/docs/files/app/core/lib/Parsers/json-class-asp.html) (more methods and fixes!)
* defaultView display `CHANGES.md`
* `INSTALL.md` and `CHANGES.md` are now written in Markdown syntax.
* removed Application("Msxml.version") and fixed the version to 6.0.
* removed iif() from classes.
* better templates!
* better documentation! It's updated to use [NaturalDocs v1.4](http://www.naturaldocs.org/ "NaturalDocs"). Of course it still uses [GeSHi](http://qbnz.com/highlighter/ "GeSHi") and [Tidy](http://tidy.sourceforge.net/ "Tidy") too.
* [SOAP Toolkit](http://www.microsoft.com/downloads/details.aspx?familyid=c943c0dd-ceec-4088-9753-86f052ec8450&displaylang=en "SOAP Toolkit 3.0")
    - [Security Update](http://www.microsoft.com/DownLoads/details.aspx?FamilyID=23d18fd1-34be-4123-ba56-9be2d4be1b23&displaylang=en "SOAP Toolkit 3.0 Security Update")
* [IIRF](http://www.codeplex.com/IIRF "Ionic's ISAPI Rewrite Filter")
    - updated to 1.2.15 (current latest stable)
    - enhanced URL-Rewriting
* a lot of other minor updates to make the Framework better ...

### Removed

* removed functions
    - Kernel.urlDecode(s) (reason: that function was not compatible with all types of Encoding.)
    - Kernel.htmlDecode(s) (reason: not required for Kernel.)
    - Kernel.sanitize(s) (reason: using `base.asp` one instead.)



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
