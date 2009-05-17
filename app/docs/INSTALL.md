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

9. That's it! Open a Web browser and type your domain. The "Welcome!" message should open.
    * To test the inspect mode, just append: ?inspect=true at the end of your domain. Ex.: http://localhost/?inspect=true
    * To test the URL Rewriting Filter, try to access the /defaultController/another/ page. Ex.: http://localhost/defaultController/another/
    * To test the other extensions, try to check if Imager.dll is responding with ?Test=True. Ex.: http://localhost/cgi-bin/Imager.dll?Test=True. If it displays: "The specified file does not exist." it's working. If you are prompted to download the file it's not.