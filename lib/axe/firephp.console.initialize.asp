<%

dim Firebug_Writer, Console

set Firebug_Writer = new Logger_Adapter_FirePHP

set Console = new Logger
Console.addAdapter Firebug_Writer

public sub Console_log(byVal l)
    if( ucase(Application("environment")) = "DEVELOPMENT" ) then Console.info l, null
end sub

public sub Console_info(byVal i)
    if( ucase(Application("environment")) = "DEVELOPMENT" ) then Console.info i, null
end sub

public sub Console_warn(byVal w)
    if( ucase(Application("environment")) = "DEVELOPMENT" ) then Console.warn w, null
end sub

public sub Console_error(byVal e)
    if( ucase(Application("environment")) = "DEVELOPMENT" ) then Console.error e, null
end sub

%>
