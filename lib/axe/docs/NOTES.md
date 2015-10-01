Important Notes
---------------

- Read **ALL** items of Installation Section, it's important. Really.

- If you are receiving an error: `80020009` without any additional information, increase the metabase `AspMaxRequestEntityAllowed` parameter size (Default is 204800, _id est_ 200KB). Maybe increate to 1MB: `cscript //NOLOGO adsutil.vbs SET W3SVC/AspMaxRequestEntityAllowed 1048576`.

- If you are receiving an error: `Response Buffer Limit Exceeded`, increase the metabase 'AspBufferingLimit' parameter size (Default is 4194304, _id est_ 4MB). Maybe increase to 64MB: `cscript //NOLOGO adsutil.vbs SET W3SVC/AspBufferingLimit 67108864`.

- There are some useful templates available at `/lib/axe/templates`.

- Never forget that this is Classic ASP! Be confident in your ASP skills and you will be fine. Also, both [AXE Documentation](/lib/axe/docs/) and [Microsoft Scripting](http://msdn.microsoft.com/en-us/library/ms950396.aspx) are always there to help you.

