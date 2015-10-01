<%

dim console

set console = new Logger
call console.addAdapter( new Logger_Adapter_FirePHP )

%>
