<%

class JSONRPC

    public default function [new](byRef S, byVal e)
        set Service = S
        exports = e
        set [new] = Me
    end function

    public sub dispose()
        set Service = nothing
    end sub

    public sub process()
        dim request_object, Req

        request_object = Request.Form("request")()
        if( isEmpty(request_object) ) then
            handle_error -32600, "Invalid Request"
            exit sub
        end if

ON ERROR RESUME NEXT : Err.clear()
' try {
    set Req = [!](request_object)
' } catch(Err) {
if( Err.number <> 0 ) then
    Err.clear()
    handle_error -32700, "Parse error"
    exit sub
end if
' }
ON ERROR GOTO 0

        if( jsType(Req) = "array" ) then
            handle_batch Req' Batch mode
        else
            handle_simple Req' Simple mode
        end if

        set Req = nothing
    end sub

    ' --[ Private section ]-----------------------------------------------------

    private Service
    private exports

    private function get_response_object()
        ' Either the result member or error member MUST be included, but both
        ' members MUST NOT be included. If there was an error in detecting the
        ' id in the Request object (e.g. Parse error/Invalid Request), it MUST
        ' be Null.
        set get_response_object = [!!]("{'jsonrpc':'2.0','id':null}")
    end function

    private function get_error_object(byVal code, byVal message)
        set get_error_object = [!]("{}")
        with get_error_object
            .set "code", code
            .set "message", message
        end with
    end function

    private function get_response_result(byRef Req, byRef Res)
        if( isEmpty( Req.get("id") ) ) then
            set get_response_result = nothing
            exit function
        end if

        set get_response_result = get_response_object()
        with get_response_result
            .set "id", Req.id
            .set "result", Res
        end with
    end function

    private function get_response_error(byRef Req, byVal code, byVal message)
        if( not isNull(Req) ) then
            if( isEmpty( Req.get("id") ) ) then
                set get_response_error = nothing
                exit function
            end if
        end if

        set get_response_error = get_response_object()
        with get_response_error
            if( not isNull(Req) ) then _
                .set "id", Req.id
            .set "error", get_error_object(code, message)
        end with
    end function

    private sub handle_error(byVal code, byVal message)
        dim Res
        set Res = get_response_error(null, code, message)
        if( not Res is nothing ) then _
            Response.write JSON.stringify(Res)
        set Res = nothing
    end sub

    private sub handle_batch(byRef Req)
        dim Res _
          , i, Ei, Ei_Res
        set Res = [!]("[]")
        for each i in Req.enumerate()
            set Ei = Req.get(i)
            set Ei_Res = handle_request(Ei)
            if( not Ei_Res is nothing ) then _
                Res.push(Ei_Res)
            set Ei_Res = nothing
            set Ei = nothing
        next
        Response.write JSON.stringify(Res)
        set Res = nothing
    end sub

    private sub handle_simple(byRef Req)
        dim Res
        set Res = handle_request(Req)
        if( not Res is nothing ) then _
            Response.write JSON.stringify(Res)
        set Res = nothing
    end sub

    private function handle_request(byRef Req)
        if( inStr( 1, exports, [](",{0},",Req.method), vbTextCompare ) = 0 ) then
            set handle_request = get_response_error(Req, -32601, "Method not found")
            exit function
        end if

        dim cmd
        if( isEmpty( Req.get("params") ) ) then
            cmd = []("set handle_request = get_response_result( Req, .{0}() )", Req.method)
        else
            cmd = []("set handle_request = get_response_result( Req, .{0}(Req.params) )", Req.method)
        end if

        with Service

ON ERROR RESUME NEXT : Err.clear()
' try {
    execute cmd
' } catch(Err) {
if( Err.number <> 0 ) then
    set handle_request = get_response_error(Req, Err.number, Err.description)
    Err.clear()
end if
' }
ON ERROR GOTO 0

        end with
    end function

    private sub Class_initialize()
        Response.contentType = "application/json"' http://www.ietf.org/rfc/rfc4627.txt
    end sub

    private sub Class_terminate()
    end sub

end class

%>

