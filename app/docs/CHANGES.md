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
