Version 1.2.0.0
---------------

### Fixes

Mmmm, I really don't have a track of the fixes 8'( but changed a lot of things 
for sure.

### Added

* brand-new logo and favicon
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
    - Kernel.exception(message) (reason: framework error system is now based on the same of asp.dll)
    - [and a lot more basic math functions...](/lib/axe/docs/files/lib/axe/base-math-asp.html)
* new classes
    - [ACL](/lib/axe/docs/files/lib/axe/classes/Utilities/acl-asp.html) _(not implemented yet)_
    - [Atom](/lib/axe/docs/files/lib/axe/classes/Feeds/atom-asp.html) _(not implemented yet)_
    - [Auth](/lib/axe/docs/files/lib/axe/classes/Utilities/auth-asp.html) _(not implemented yet)_
    - [Base64](/lib/axe/docs/files/lib/axe/classes/Utilities/base64-asp.html)
    - [Color](/lib/axe/docs/files/lib/axe/classes/Utilities/color-asp.html)
    - [CSV](/lib/axe/docs/files/lib/axe/classes/Parsers/csv-asp.html) _(not implemented yet)_
    - [CustomEvent](/lib/axe/docs/files/lib/axe/classes/customevent-asp.html)
    - [Email](/lib/axe/docs/files/lib/axe/classes/Utilities/email-asp.html)
    - [Interface](/lib/axe/docs/files/lib/axe/classes/interface-asp.html)
    - [JSON2](/lib/axe/docs/files/lib/axe/classes/Parsers/json2-asp.html)
    - [List](/lib/axe/docs/files/lib/axe/classes/Utilities/list-asp.html)
    - [Logger](/lib/axe/docs/files/lib/axe/classes/Utilities/logger-asp.html) _(not implemented yet)_
    - [Markdown](/lib/axe/docs/files/lib/axe/classes/Parsers/markdown-asp.html)
    - [Orderly](/lib/axe/docs/files/lib/axe/classes/Parsers/orderly-asp.html)
    - [Paginator](/lib/axe/docs/files/lib/axe/classes/Utilities/paginator-asp.html)
    - [RSS](/lib/axe/docs/files/lib/axe/classes/Feeds/rss-asp.html) _(work in progress)_
    - [StringBuilder](/lib/axe/docs/files/lib/axe/classes/Utilities/stringbuilder-asp.html)
    - [Template](/lib/axe/docs/files/lib/axe/classes/Utilities/template-asp.html)
    - [Textile](/lib/axe/docs/files/lib/axe/classes/Parsers/textile-asp.html) _(not implemented yet)_
    - [Translator](/lib/axe/docs/files/lib/axe/classes/Utilities/translator-asp.html) _(work in progress)_
    - [UnitTest](/lib/axe/docs/files/lib/axe/classes/unittest-asp.html)
    - [XSession](/lib/axe/docs/files/lib/axe/classes/Utilities/xsession-asp.html) _(not implemented yet)_
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
    - updated to 1.2.15 (current latest stable)
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
