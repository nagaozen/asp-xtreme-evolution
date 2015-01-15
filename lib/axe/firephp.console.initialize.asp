<%

dim Console

set Console = new Logger
call Console.addAdapter( new Logger_Adapter_FirePHP )

public sub Console_log(byVal l)
    if( ucase( Application("environment") ) = "DEVELOPMENT" ) then _
        Console.debug l, null
end sub

public sub Console_info(byVal i)
    if( ucase( Application("environment") ) = "DEVELOPMENT" ) then _
        Console.info i, null
end sub

public sub Console_warn(byVal w)
    if( ucase( Application("environment") ) = "DEVELOPMENT" ) then _
        Console.warn w, null
end sub

public sub Console_error(byVal e)
    if( ucase( Application("environment") ) = "DEVELOPMENT" ) then _
        Console.error e, null
end sub

%>
