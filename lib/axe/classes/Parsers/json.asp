<%

' File: json.asp
' 
' AXE(ASP Xtreme Evolution) implementation of JSON parser.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2007-2009 Fabio Zendhi Nagao
' 
' ASP Xtreme Evolution is free software: you can redistribute it and/or modify
' it under the terms of the GNU Lesser General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
' 
' ASP Xtreme Evolution is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU Lesser General Public License for more details.
' 
' You should have received a copy of the GNU Lesser General Public License
' along with ASP Xtreme Evolution. If not, see <http://www.gnu.org/licenses/>.



' Class: Json
' 
' This Class goal is to provide a simple way to parse JSON (JavaScript Object 
' Notation) data directly from vbscript.
' 
' Notes:
' 
'   - JScript Class approach is inspired by CCB.JSONParser.asp <http://blog.crayoncowboy.com/?p=7> Copyright (c) 2007 Cliff Pruitt.
' 
' About:
' 
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ January 2008
' 
class Json

    ' Property: classType
    ' 
    ' Class type.
    ' 
    ' Contains:
    ' 
    '   (string) - type
    ' 
    public classType

    ' Property: classVersion
    ' 
    ' Class version.
    ' 
    ' Contains:
    ' 
    '   (float) - version
    ' 
    public classVersion

    private root

    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "1.2.0"
    end sub

    private sub Class_terminate()
        set root = nothing
    end sub

    ' Subroutine: loadJson
    ' 
    ' Since ASP Classes strangely doesn't accept parameters at it's initializa-
    ' tion. Use this subroutine to load the object.
    ' 
    ' Parameters:
    ' 
    '   (string) - The Json string representation
    ' 
    public sub loadJson(sJson)
        set root = new_JsonEngine(sJson)
    end sub

    ' Function: getElement
    ' 
    ' This function takes a dot separated path and look for the element value in
    ' the JSON.
    ' 
    ' Parameters:
    ' 
    '   (string) - Relative path from root
    ' 
    ' Returns:
    ' 
    '   (string) - The element value
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim sJson, oJson
    ' 
    ' sJson = join(array( _
    ' "{", _
    ' "    'hello' : 'Hello World !',", _
    ' "    'howdy' : 'How do you do ?',", _
    ' "    'fields' : {", _
    ' "       'one': 1,", _
    ' "       'two': 2,", _
    ' "       'three': 3,", _
    ' "       'four': [ 'one', 'two,three', 'four' ],", _
    ' "       'five': { 'one' : 'apple', 'two' : 'orange', 'three' : 'banana' }", _
    ' "    }", _
    ' "};" _
    ' ), vbNewLine)
    ' 
    ' set oJson = new Json
    ' oJson.loadJson(sJson)
    ' Response.write(oJson.getElement("hello") & "<br />" & vbNewLine)
    ' Response.write(oJson.getElement("howdy") & "<br />" & vbNewLine)
    ' set oJson = nothing
    ' 
    ' (end)
    ' 
    public function getElement(sPath)
        if(typename(root.getElement(sPath)) = "String") then
            getElement = string_decode( root.getElement(sPath) )
        else
            getElement = root.getElement(sPath)
        end if
    end function

    private function string_encode(value)
        string_encode = sanitize(value, array(vbCr, vbLf, """"), array("\r", "\n", "\"""))
    end function
    
    private function string_decode(value)
        string_decode = sanitize(value, array("\r", "\n", "\"""), array(vbCr, vbLf, """"))
    end function

    ' Subroutine: setElement
    ' 
    ' This subroutine augments the JSON by adding elements to it.
    ' 
    ' Notes:
    ' 
    '   - You can pass valid JSON Objects and Arrays as strings ( { ... } and [ ... ] ) respectively and this function will encode it to javascript.
    '   - The above note said valid JSON. Which means Objects and Arrays should escape " properly (\") since it's used as the JSON string delimiter.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' set oJson = new Json
    ' oJson.loadJson("{}")
    ' 
    ' call oJson.setElement("obj", "{'key':'value'}")
    ' call oJson.setElement("arr", "[1,2,3,4]")
    ' call oJson.setElement("str", "string")
    ' call oJson.setElement("int", 1982)
    ' call oJson.setElement("bool", true)
    ' call oJson.setElement("obj.augment", "I'm augmenting the first setElement element XD")
    ' Response.write(oJson.serialize("") & "<br />" & vbNewLine)
    ' 
    ' set oJson = nothing
    ' 
    ' (end)
    ' 
    public sub setElement(sPath, value)
        call root.setElement(sPath, value)
    end sub

    ' Subroutine: removeElement
    ' 
    ' This subroutine removes a node from the object.
    ' 
    ' Notes:
    ' 
    '   - Caution! It's recursive, removing a container ( Object or Array ) result in erasing it and all it's content.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' set oJson = new Json
    ' oJson.loadJson("{}")
    ' 
    ' call oJson.setElement("obj", "{'key':'value'}")
    ' call oJson.setElement("arr", "[1,2,3,4]")
    ' call oJson.setElement("str", "string")
    ' call oJson.setElement("int", 1982)
    ' call oJson.setElement("bool", true)
    ' call oJson.setElement("obj.augment", "I'm augmenting the first setElement element XD")
    ' Response.write(oJson.serialize("") & "<br />" & vbNewLine)
    ' 
    ' call oJson.removeElement("obj.augment")
    ' Response.write(oJson.serialize("") & "<br />" & vbNewLine)
    ' 
    ' set oJson = nothing
    ' 
    ' (end)
    ' 
    public sub removeElement(sPath)
        call root.removeElement(sPath)
    end sub

    ' Function: getChildNodes
    ' 
    ' Look for all element child keys and enumerate them.
    ' 
    ' Parameters:
    ' 
    '   (string) - Path to the parent element relative to root.
    ' 
    ' Returns:
    ' 
    '  (string[]) - With the child keys
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim sJson, oJson, key
    ' 
    ' sJson = join(array( _
    ' "{", _
    ' "    'hello' : 'Hello World !',", _
    ' "    'howdy' : 'How do you do ?',", _
    ' "    'fields' : {", _
    ' "       'one': 1,", _
    ' "       'two': 2,", _
    ' "       'three': 3,", _
    ' "       'four': [ 'one', 'two,three', 'four' ],", _
    ' "       'five': { 'one' : 'apple', 'two' : 'orange', 'three' : 'banana' }", _
    ' "    }", _
    ' "};" _
    ' ), vbNewLine)
    ' 
    ' set oJson = new Json
    ' oJson.loadJson(sJson)
    ' for each key in oJson.getChildNodes("")
    '     Response.write(key & " : " & oJson.getElement(key) & "<br />" & vbNewLine)
    ' next
    ' set oJson = nothing
    ' 
    ' (end)
    ' 
    public function getChildNodes(sPath)
        getChildNodes = split(root.getChildNodes(sPath), ",")
    end function
    
    ' Function: serialize
    ' 
    ' Converts the object into a JSON string.
    ' 
    ' Parameters:
    ' 
    '     (string) - starting path. "" means the entire object.
    ' 
    ' Returns:
    ' 
    '     (string) - a JSON string
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim sJson, oJson, key
    ' 
    ' sJson = join(array( _
    ' "{", _
    ' "    'hello' : 'Hello World !',", _
    ' "    'howdy' : 'How do you do ?',", _
    ' "    'fields' : {", _
    ' "       'one': 1,", _
    ' "       'two': 2,", _
    ' "       'three': 3,", _
    ' "       'four': [ 'one', 'two,three', 'four' ],", _
    ' "       'five': { 'one' : 'apple', 'two' : 'orange', 'three' : 'banana' }", _
    ' "    },", _
    ' "    'boolean' : true", _
    ' "};" _
    ' ), vbNewLine)
    ' 
    ' set oJson = new Json
    ' oJson.loadJson(sJson)
    ' Response.write oJson.serialize("")
    ' set oJson = nothing
    ' 
    ' (end code)
    ' 
    public function serialize(path)
        dim Stream, value : value = getElement(path)
        select case value
            case "[object Array]"
                set Stream = Server.createObject("ADODB.Stream")
                Stream.type = adTypeText
                Stream.mode = adModeReadWrite
                Stream.open()
                for each path in root.getChildNodes(path)
                    Stream.writeText(",")
                    Stream.writeText(serialize(path))
                next
                Stream.position = 0
                serialize = Stream.readText()
                Stream.close()
                set Stream = nothing
                
                if( len(serialize) > 0 ) then
                    serialize = "[" & mid(serialize, 2) & "]"
                else
                    serialize = "[]"
                end if
            case "[object Object]"
                dim a
                
                set Stream = Server.createObject("ADODB.Stream")
                Stream.type = adTypeText
                Stream.mode = adModeReadWrite
                Stream.open()
                for each path in root.getChildNodes(path)
                    a = split(path, ".")
                    Stream.writeText(",")
                    Stream.writeText("""")
                    Stream.writeText(a(ubound(a)))
                    Stream.writeText(""":")
                    Stream.writeText(serialize(path))
                next
                Stream.position = 0
                serialize = Stream.readText()
                Stream.close()
                set Stream = nothing
                
                if( len(serialize) > 0 ) then
                    serialize = "{" & mid(serialize, 2) & "}"
                else
                    serialize = "{}"
                end if
            case else
                select case lcase(typename(value))
                    case "string":
                        serialize = """" & string_encode(value) & """"
                    case "boolean":
                        serialize = lcase(value)
                    case else:
                        serialize = value
                end select
        end select
    end function

end class

%>
<script language="javascript" runat="server">

/* Function: new_JsonEngine
 * 
 * Private function used to create a new instance of the JsonEngine Class.
 * 
 * Parameters:
 * 
 *  (string) - The Json string representation
 * 
 * Returns:
 * 
 *  (object) - The Json root.
 */
function new_JsonEngine(sJson) {
    return new JsonEngine(sJson);
}

/* Class: JsonEngine
 * 
 * Since VBScript doesn't provide a native method to handle JSON, this class makes
 * the magic of wrapping JScript JSON to VBScript. Note that this class methods
 * are the same of the Json class above and just make what they were supposed to do.
 */
function JsonEngine(sJson) {
    var me = this;
    this.data = {};

    this.toJSON = function(value) {
        eval("var json = " + value);
        return json;
    };

    this.initialize = function(sJson) {
        this.data = me.toJSON(sJson);
    };

    this.isObject = function(value) {
        return /^{[^}]*}$/im.test(value);
    };

    this.isArray = function(value) {
        return /^\[[^]]*\]$/im.test(value);
    };

    this.getElement = function(sPath) {
        if(sPath === '') return me.data;
        var node = me.data;
        var aPath = sPath.split('.');
        for(var i = 0, len = aPath.length; i < len; i++) {
            if(!node[aPath[i]]) return "";
            node = node[aPath[i]];
        }
        return (typeof node == "object" && node.length)? "[object Array]" : node;
    };

    this.setElement = function(sPath, value) {
        if(me.isObject(value) || me.isArray(value)) value = me.toJSON(value);
        if(typeof value == "string") value = value.replace(/\r/g, "\\r").replace(/\n/g, "\\n").replace(/"/g, '\"');
        
        var parentNode = me.data;
        var aPath = sPath.split('.');
        for(var i = 0, len = (aPath.length - 1); i < len; i++) {
            if(!parentNode[aPath[i]]) parentNode[aPath[i]] = {};
            parentNode = parentNode[aPath[i]];
        }
        
        parentNode[aPath[aPath.length - 1]] = value;
    };

    this.removeElement = function(sPath) {
        var parentNode = me.data;
        var aPath = sPath.split('.');
        for(var i = 0, len = (aPath.length - 1); i < len; i++) {
            if(!parentNode[aPath[i]]) parentNode[aPath[i]] = {};
            parentNode = parentNode[aPath[i]];
        }
        delete parentNode[aPath[aPath.length - 1]];
    };

    this.getChildNodes = function(sPath) {
        var keys = [];
        var parentNode = me.data;
        if( sPath.length > 0 ) {
            var aPath = sPath.split('.');
            for(var i = 0, len = aPath.length; i < len; i++) {
                parentNode = parentNode[aPath[i]];
            }
        }
        for(var key in parentNode) {
            (sPath.length > 0)? keys.push(sPath + "." + key) : keys.push(key);
        }
        return keys;
    };

    this.initialize(sJson);
}

</script>
