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

2. If the domain is not registered in the server yet, register it in IIS pointing it's root to the folder where `default.asp` is located. Don't forget to check "Run scripts (such as ASP)".

3. If your server does not provide a default `cgi-bin` folder, create a virtual directory named "cgi-bin" in your domain root pointing to `/lib/axe/cgi-bin`. Don't forget to check "Execute (such as ISAPI applications or CGI)"; Otherwise, move the contents of `/lib/axe/cgi-bin` to your cgi-bin folder.

4. Give only "Read and Write" permission to your `/instance/writables` folder.

5. Go to your site properties and in ISAPI Filters tab, click [Add]. Fill the box with the following data:  
   Filter name: IIRF  
   [Browse...] to `C:\Program Files\Ionic Shade\IIRF 2.1\IIRF.dll`  

6. Go to your Web Service Extensions and Add the binaries that comes with the package. They are located at `/lib/axe/cgi-bin` and their names are:
    * CB Image Resizer (Imager.dll)
    * CB Zip (CBZIP.exe)
    
    Notes:
    
    * Don't forget to check "Set extension status to Allowed".
    * If Active Server Pages is not allowed yet, allow it too.

7. If your server does not support MSXML 6.0 yet, install it. It's available at: `/lib/axe/bin/msxml6.msi`

8. Create an application pool for the application views.

9. Set `/app/views` to the application pool created in step 8.

10. That's it! Open a Web browser and type your domain. The "Welcome to ASP Xtreme Evolution" page should open.

