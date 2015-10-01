<script language="Javascript" runat="server">

/*

File: json2.asp

AXE(ASP Xtreme Evolution) JSON parser based on Douglas Crockford json2.js.

This class is the result of Classic ASP JSON topic revisited by Fabio Zendhi 
Nagao (nagaozen). JSON2.ASP is a better option over JSON.ASP because it embraces
the AXE philosophy of real collaboration over the languages. It works under the
original json parser, so this class is strict in the standard rules, it also 
brings more of the Javascript json feeling to other ASP languages (eg. no more 
oJson.getElement("foo") stuff, just oJson.foo and you get it).

License:

This file is part of ASP Xtreme Evolution.
Copyright (C) 2007-2012 Fabio Zendhi Nagao

ASP Xtreme Evolution is free software: you can redistribute it and/or modify
it under the terms of the GNU Lesser General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

ASP Xtreme Evolution is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Lesser General Public License for more details.

You should have received a copy of the GNU Lesser General Public License
along with ASP Xtreme Evolution. If not, see <http://www.gnu.org/licenses/>.





Class: JSON

JSON (Javascript Object Notation) is a lightweight data-interchange format. It 
is easy for humans to read and write. It is easy for machines to parse and 
generate. It is based on a subset of the Javascript Programming Language, 
Standard ECMA-262 3rd Edition - December 1999. JSON is a text format that is 
completely language independent but uses conventions that are familiar to 
programmers of the C-family of languages, including C, C++, C#, Java, 
Javascript, Perl, Python, and many others. These properties make JSON an ideal 
data-interchange language.

Notes:

    - JSON parse/stringify from the Douglas Crockford json2.js <https://raw.githubusercontent.com/douglascrockford/JSON-js/master/json2.js>.
    - JSON.toXML is based on the Prof. Stefan Gössner "Converting Between XML and JSON" pragmatic approach <http://www.xml.com/pub/a/2006/05/31/converting-between-xml-and-json.html>.
    - JSON.minify is based on <https://github.com/getify/JSON.minify/blob/master/minify.json.js> and exists because of <https://plus.google.com/+DouglasCrockfordEsq/posts/RK8qyGVaGSr>.

About:

    - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ August 2010





Function: parse

This method parses a JSON text to produce an object or array. It can throw a SyntaxError exception.

Parameters:

    (string) - Valid JSON text.

Returns:

    (mixed) - a Javascript value, usually an object or array.

Example:

(start code)

dim Info : set Info = JSON.parse(join(array( _
    "{", _
    "  ""firstname"": ""Fabio"",", _
    "  ""lastname"": ""長尾"",", _
    "  ""alive"": true,", _
    "  ""age"": 27,", _
    "  ""nickname"": ""nagaozen"",", _
    "  ""fruits"": [", _
    "    ""banana"",", _
    "    ""orange"",", _
    "    ""apple"",", _
    "    ""papaya"",", _
    "    ""pineapple""", _
    "  ],", _
    "  ""complex"": {", _
    "    ""real"": 1,", _
    "    ""imaginary"": 2", _
    "  }", _
    "}" _
)))

Response.write(Info.firstname & vbNewline) ' prints Fabio
Response.write(Info.alive & vbNewline) ' prints True
Response.write(Info.age & vbNewline) ' prints 27
Response.write(Info.fruits.get(0) & vbNewline) ' prints banana
Response.write(Info.fruits.get(1) & vbNewline) ' prints orange
Response.write(Info.complex.real & vbNewline) ' prints 1
Response.write(Info.complex.imaginary & vbNewline) ' prints 2

' You can also enumerate object properties ...

dim key : for each key in Info.enumerate()
    Response.write( key & vbNewline )
next

' which prints:

' firstname
' lastname
' alive
' age
' nickname
' fruits
' complex

set Info = nothing

(end code)





Function: stringify

This method produces a JSON text from a Javascript value.

Parameters:

    (mixed) - any Javascript value, usually an object or array.
    (mixed) - an optional parameter that determines how object values are stringified for objects. It can be a function or an array of strings.
    (mixed) - an optional parameter that specifies the indentation of nested structures. If it is omitted, the text will be packed without extra whitespace. If it is a number, it will specify the number of spaces to indent at each level. If it is a string (such as '\t' or '&nbsp;'), it contains the characters used to indent at each level.

Returns:

    (string) - a string that contains the serialized JSON text.

Example:

(start code)

dim Info : set Info = JSON.parse("{""firstname"":""Fabio"", ""lastname"":""長尾""}")
Info.set "alive", true
Info.set "age", 27
Info.set "nickname", "nagaozen"
Info.set "fruits", array("banana","orange","apple","papaya","pineapple")
Info.set "complex", JSON.parse("{""real"":1, ""imaginary"":1}")

Response.write( JSON.stringify(Info, null, 2) & vbNewline ) ' prints the text below:
'{
'  "firstname": "Fabio",
'  "lastname": "長尾",
'  "alive": true,
'  "age": 27,
'  "nickname": "nagaozen",
'  "fruits": [
'    "banana",
'    "orange",
'    "apple",
'    "papaya",
'    "pineapple"
'  ],
'  "complex": {
'    "real": 1,
'    "imaginary": 1
'  }
'}

set Info = nothing

(end code)





Function: toXML

This method produces a XML text from a Javascript value.

Parameters:

    (mixed) - any Javascript value, usually an object or array.
    (string) - an optional parameter that determines what tag should be used as a container for the output. Defaults to none.

Returns:

    (string) - a string that contains the serialized XML text.

Example:

(start code)

dim Info : set Info = JSON.parse("{""firstname"":""Fabio"", ""lastname"":""長尾""}")
Info.set "alive", true
Info.set "age", 27
Info.set "nickname", "nagaozen"
Info.set "fruits", array("banana","orange","apple","papaya","pineapple")
Info.set "complex", JSON.parse("{""real"":1, ""imaginary"":1}")

Response.write( JSON.toXML(Info) & vbNewline ) ' prints the text below:
'<firstname>Fabio</firstname>
'<lastname>長尾</lastname>
'<alive>true</alive>
'<age>27</age>
'<nickname>nagaozen</nickname>
'<fruits>banana</fruits>
'<fruits>orange</fruits>
'<fruits>apple</fruits>
'<fruits>papaya</fruits>
'<fruits>pineapple</fruits>
'<complex>
'    <real>1</real>
'    <imaginary>1</imaginary>
'</complex>

set Info = nothing

(end code)





Function: minify

This method can be used as a helper to enable comments in json-like 
configuration files. According to Douglas Crockford, using comments are fine if
you pipe the code before handing it to your JSON parser. See 
<https://plus.google.com/118095276221607585885/posts/RK8qyGVaGSr>

Parameters:

    (string) - a json-like configuration string

Returns:

    (json) - valid minified json

*/

if(!Object.prototype.get) {
    Object.prototype.get = function(k) {
        return this[k];
    }
}

if(!Object.prototype.set) {
    Object.prototype.set = function(k,v) {
        if(typeof(v) === "unknown") {
            try {
                v = (new VBArray(v)).toArray();
            } catch(e) {
                return;
            }
        }
        this[k] = v;
    }
}

if(!Object.prototype.purge) {
    Object.prototype.purge = function(k) {
        delete this[k];
    }
}

if(!Object.prototype.enumerate) {
    Object.prototype.enumerate = function() {
        var d = new ActiveXObject("Scripting.Dictionary");
        for(var key in this) {
            if(this.hasOwnProperty(key)) {
                d.add(key, this[key]);
            }
        }
        return d.keys();
    }
}

if(!String.prototype.sanitize) {
    String.prototype.sanitize = function(a, b) {
        var len = a.length,
            s = this;
        if(len !== b.length) throw new TypeError('Invalid procedure call. Both arrays should have the same size.');
        for(var i = 0; i < len; i++) {
            var re = new RegExp(a[i],'g');
            s = s.replace(re, b[i]);
        }
        return s;
    }
}

if(!String.prototype.substitute) {
    String.prototype.substitute = function(object, regexp){
        return this.replace(regexp || (/\\?\{([^{}]+)\}/g), function(match, name){
            if(match.charAt(0) == '\\') return match.slice(1);
            return (object[name] != undefined) ? object[name] : '';
        });
    }
}










/*
    json2.js
    2014-02-04

    Public Domain.

    NO WARRANTY EXPRESSED OR IMPLIED. USE AT YOUR OWN RISK.

    See http://www.JSON.org/js.html


    This code should be minified before deployment.
    See http://javascript.crockford.com/jsmin.html

    USE YOUR OWN COPY. IT IS EXTREMELY UNWISE TO LOAD CODE FROM SERVERS YOU DO
    NOT CONTROL.


    This file creates a global JSON object containing two methods: stringify
    and parse.

        JSON.stringify(value, replacer, space)
            value       any JavaScript value, usually an object or array.

            replacer    an optional parameter that determines how object
                        values are stringified for objects. It can be a
                        function or an array of strings.

            space       an optional parameter that specifies the indentation
                        of nested structures. If it is omitted, the text will
                        be packed without extra whitespace. If it is a number,
                        it will specify the number of spaces to indent at each
                        level. If it is a string (such as '\t' or '&nbsp;'),
                        it contains the characters used to indent at each level.

            This method produces a JSON text from a JavaScript value.

            When an object value is found, if the object contains a toJSON
            method, its toJSON method will be called and the result will be
            stringified. A toJSON method does not serialize: it returns the
            value represented by the name/value pair that should be serialized,
            or undefined if nothing should be serialized. The toJSON method
            will be passed the key associated with the value, and this will be
            bound to the value

            For example, this would serialize Dates as ISO strings.

                Date.prototype.toJSON = function (key) {
                    function f(n) {
                        // Format integers to have at least two digits.
                        return n < 10 ? '0' + n : n;
                    }

                    return this.getUTCFullYear()   + '-' +
                         f(this.getUTCMonth() + 1) + '-' +
                         f(this.getUTCDate())      + 'T' +
                         f(this.getUTCHours())     + ':' +
                         f(this.getUTCMinutes())   + ':' +
                         f(this.getUTCSeconds())   + 'Z';
                };

            You can provide an optional replacer method. It will be passed the
            key and value of each member, with this bound to the containing
            object. The value that is returned from your method will be
            serialized. If your method returns undefined, then the member will
            be excluded from the serialization.

            If the replacer parameter is an array of strings, then it will be
            used to select the members to be serialized. It filters the results
            such that only members with keys listed in the replacer array are
            stringified.

            Values that do not have JSON representations, such as undefined or
            functions, will not be serialized. Such values in objects will be
            dropped; in arrays they will be replaced with null. You can use
            a replacer function to replace those with JSON values.
            JSON.stringify(undefined) returns undefined.

            The optional space parameter produces a stringification of the
            value that is filled with line breaks and indentation to make it
            easier to read.

            If the space parameter is a non-empty string, then that string will
            be used for indentation. If the space parameter is a number, then
            the indentation will be that many spaces.

            Example:

            text = JSON.stringify(['e', {pluribus: 'unum'}]);
            // text is '["e",{"pluribus":"unum"}]'


            text = JSON.stringify(['e', {pluribus: 'unum'}], null, '\t');
            // text is '[\n\t"e",\n\t{\n\t\t"pluribus": "unum"\n\t}\n]'

            text = JSON.stringify([new Date()], function (key, value) {
                return this[key] instanceof Date ?
                    'Date(' + this[key] + ')' : value;
            });
            // text is '["Date(---current time---)"]'


        JSON.parse(text, reviver)
            This method parses a JSON text to produce an object or array.
            It can throw a SyntaxError exception.

            The optional reviver parameter is a function that can filter and
            transform the results. It receives each of the keys and values,
            and its return value is used instead of the original value.
            If it returns what it received, then the structure is not modified.
            If it returns undefined then the member is deleted.

            Example:

            // Parse the text. Values that look like ISO date strings will
            // be converted to Date objects.

            myData = JSON.parse(text, function (key, value) {
                var a;
                if (typeof value === 'string') {
                    a =
/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:\.\d*)?)Z$/.exec(value);
                    if (a) {
                        return new Date(Date.UTC(+a[1], +a[2] - 1, +a[3], +a[4],
                            +a[5], +a[6]));
                    }
                }
                return value;
            });

            myData = JSON.parse('["Date(09/09/2001)"]', function (key, value) {
                var d;
                if (typeof value === 'string' &&
                        value.slice(0, 5) === 'Date(' &&
                        value.slice(-1) === ')') {
                    d = new Date(value.slice(5, -1));
                    if (d) {
                        return d;
                    }
                }
                return value;
            });


    This is a reference implementation. You are free to copy, modify, or
    redistribute.
*/

/*jslint evil: true, regexp: true */

/*members "", "\b", "\t", "\n", "\f", "\r", "\"", JSON, "\\", apply,
    call, charCodeAt, getUTCDate, getUTCFullYear, getUTCHours,
    getUTCMinutes, getUTCMonth, getUTCSeconds, hasOwnProperty, join,
    lastIndex, length, parse, prototype, push, replace, slice, stringify,
    test, toJSON, toString, valueOf
*/


// Create a JSON object only if one does not already exist. We create the
// methods in a closure to avoid creating global variables.

if (typeof JSON !== 'object') {
    JSON = {};
}

(function () {
    'use strict';

    function f(n) {
        // Format integers to have at least two digits.
        return n < 10 ? '0' + n : n;
    }

    if (typeof Date.prototype.toJSON !== 'function') {

        Date.prototype.toJSON = function () {

            return isFinite(this.valueOf())
                ? this.getUTCFullYear()     + '-' +
                    f(this.getUTCMonth() + 1) + '-' +
                    f(this.getUTCDate())      + 'T' +
                    f(this.getUTCHours())     + ':' +
                    f(this.getUTCMinutes())   + ':' +
                    f(this.getUTCSeconds())   + 'Z'
                : null;
        };

        String.prototype.toJSON      =
            Number.prototype.toJSON  =
            Boolean.prototype.toJSON = function () {
                return this.valueOf();
            };
    }

    var cx,
        escapable,
        gap,
        indent,
        meta,
        rep;


    function quote(string) {

// If the string contains no control characters, no quote characters, and no
// backslash characters, then we can safely slap some quotes around it.
// Otherwise we must also replace the offending characters with safe escape
// sequences.

        escapable.lastIndex = 0;
        return escapable.test(string) ? '"' + string.replace(escapable, function (a) {
            var c = meta[a];
            return typeof c === 'string'
                ? c
                : '\\u' + ('0000' + a.charCodeAt(0).toString(16)).slice(-4);
        }) + '"' : '"' + string + '"';
    }


    function str(key, holder) {

// Produce a string from holder[key].

        var i,          // The loop counter.
            k,          // The member key.
            v,          // The member value.
            length,
            mind = gap,
            partial,
            value = holder[key];

// If the value has a toJSON method, call it to obtain a replacement value.

        if (value && typeof value === 'object' &&
                typeof value.toJSON === 'function') {
            value = value.toJSON(key);
        }

// If we were called with a replacer function, then call the replacer to
// obtain a replacement value.

        if (typeof rep === 'function') {
            value = rep.call(holder, key, value);
        }

// What happens next depends on the value's type.

        switch (typeof value) {
        case 'string':
            return quote(value);

        case 'number':

// JSON numbers must be finite. Encode non-finite numbers as null.

            return isFinite(value) ? String(value) : 'null';

        case 'boolean':
        case 'null':

// If the value is a boolean or null, convert it to a string. Note:
// typeof null does not produce 'null'. The case is included here in
// the remote chance that this gets fixed someday.

            return String(value);

// If the type is 'object', we might be dealing with an object or an array or
// null.

        case 'object':

// Due to a specification blunder in ECMAScript, typeof null is 'object',
// so watch out for that case.

            if (!value) {
                return 'null';
            }

// Make an array to hold the partial results of stringifying this object value.

            gap += indent;
            partial = [];

// Is the value an array?

            if (Object.prototype.toString.apply(value) === '[object Array]') {

// The value is an array. Stringify every element. Use null as a placeholder
// for non-JSON values.

                length = value.length;
                for (i = 0; i < length; i += 1) {
                    partial[i] = str(i, value) || 'null';
                }

// Join all of the elements together, separated with commas, and wrap them in
// brackets.

                v = partial.length === 0
                    ? '[]'
                    : gap
                    ? '[\n' + gap + partial.join(',\n' + gap) + '\n' + mind + ']'
                    : '[' + partial.join(',') + ']';
                gap = mind;
                return v;
            }

// If the replacer is an array, use it to select the members to be stringified.

            if (rep && typeof rep === 'object') {
                length = rep.length;
                for (i = 0; i < length; i += 1) {
                    if (typeof rep[i] === 'string') {
                        k = rep[i];
                        v = str(k, value);
                        if (v) {
                            partial.push(quote(k) + (gap ? ': ' : ':') + v);
                        }
                    }
                }
            } else {

// Otherwise, iterate through all of the keys in the object.

                for (k in value) {
                    if (Object.prototype.hasOwnProperty.call(value, k)) {
                        v = str(k, value);
                        if (v) {
                            partial.push(quote(k) + (gap ? ': ' : ':') + v);
                        }
                    }
                }
            }

// Join all of the member texts together, separated with commas,
// and wrap them in braces.

            v = partial.length === 0
                ? '{}'
                : gap
                ? '{\n' + gap + partial.join(',\n' + gap) + '\n' + mind + '}'
                : '{' + partial.join(',') + '}';
            gap = mind;
            return v;
        }
    }

// If the JSON object does not yet have a stringify method, give it one.

    if (typeof JSON.stringify !== 'function') {
        escapable = /[\\\"\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;
        meta = {    // table of character substitutions
            '\b': '\\b',
            '\t': '\\t',
            '\n': '\\n',
            '\f': '\\f',
            '\r': '\\r',
            '"' : '\\"',
            '\\': '\\\\'
        };
        JSON.stringify = function (value, replacer, space) {

// The stringify method takes a value and an optional replacer, and an optional
// space parameter, and returns a JSON text. The replacer can be a function
// that can replace values, or an array of strings that will select the keys.
// A default replacer method can be provided. Use of the space parameter can
// produce text that is more easily readable.

            var i;
            gap = '';
            indent = '';

// If the space parameter is a number, make an indent string containing that
// many spaces.

            if (typeof space === 'number') {
                for (i = 0; i < space; i += 1) {
                    indent += ' ';
                }

// If the space parameter is a string, it will be used as the indent string.

            } else if (typeof space === 'string') {
                indent = space;
            }

// If there is a replacer, it must be a function or an array.
// Otherwise, throw an error.

            rep = replacer;
            if (replacer && typeof replacer !== 'function' &&
                    (typeof replacer !== 'object' ||
                    typeof replacer.length !== 'number')) {
                throw new Error('JSON.stringify');
            }

// Make a fake root object containing our value under the key of ''.
// Return the result of stringifying the value.

            return str('', {'': value});
        };
    }


// If the JSON object does not yet have a parse method, give it one.

    if (typeof JSON.parse !== 'function') {
        cx = /[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;
        JSON.parse = function (text, reviver) {

// The parse method takes a text and an optional reviver function, and returns
// a JavaScript value if the text is a valid JSON text.

            var j;

            function walk(holder, key) {

// The walk method is used to recursively walk the resulting structure so
// that modifications can be made.

                var k, v, value = holder[key];
                if (value && typeof value === 'object') {
                    for (k in value) {
                        if (Object.prototype.hasOwnProperty.call(value, k)) {
                            v = walk(value, k);
                            if (v !== undefined) {
                                value[k] = v;
                            } else {
                                delete value[k];
                            }
                        }
                    }
                }
                return reviver.call(holder, key, value);
            }


// Parsing happens in four stages. In the first stage, we replace certain
// Unicode characters with escape sequences. JavaScript handles many characters
// incorrectly, either silently deleting them, or treating them as line endings.

            text = String(text);
            cx.lastIndex = 0;
            if (cx.test(text)) {
                text = text.replace(cx, function (a) {
                    return '\\u' +
                        ('0000' + a.charCodeAt(0).toString(16)).slice(-4);
                });
            }

// In the second stage, we run the text against regular expressions that look
// for non-JSON patterns. We are especially concerned with '()' and 'new'
// because they can cause invocation, and '=' because it can cause mutation.
// But just to be safe, we want to reject all unexpected forms.

// We split the second stage into 4 regexp operations in order to work around
// crippling inefficiencies in IE's and Safari's regexp engines. First we
// replace the JSON backslash pairs with '@' (a non-JSON character). Second, we
// replace all simple value tokens with ']' characters. Third, we delete all
// open brackets that follow a colon or comma or that begin the text. Finally,
// we look to see that the remaining characters are only whitespace or ']' or
// ',' or ':' or '{' or '}'. If that is so, then the text is safe for eval.

            if (/^[\],:{}\s]*$/
                    .test(text.replace(/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g, '@')
                        .replace(/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g, ']')
                        .replace(/(?:^|:|,)(?:\s*\[)+/g, ''))) {

// In the third stage we use the eval function to compile the text into a
// JavaScript structure. The '{' operator is subject to a syntactic ambiguity
// in JavaScript: it can begin a block or an object literal. We wrap the text
// in parens to eliminate the ambiguity.

                j = eval('(' + text + ')');

// In the optional fourth stage, we recursively walk the new structure, passing
// each name/value pair to a reviver function for possible transformation.

                return typeof reviver === 'function'
                    ? walk({'': j}, '')
                    : j;
            }

// If the text is not JSON parseable, then a SyntaxError is thrown.

            throw new SyntaxError('JSON.parse');
        };
    }
}());










// Implement toXML
(function(){

/*
IMPORTANT: XML 1.0 name token is defined here <http://www.w3.org/TR/2008/REC-xml-20081126/#NT-Nmtoken>
            [#x10000-#xEFFFF] range is left out because ECMA 262 language specification (ECMAScript Edition 3)
            doesn't have a way to represent them as Regular Expressions Ranges.
*/
    function __tagize(value) {
        if( /^[_:A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF\uF900-\uFDCF\uFDF0-\uFFFD][-.·_:0-9A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF\uF900-\uFDCF\uFDF0-\uFFFD\u0300-\u036F\u203F-\u2040]*$/.test(value) ) {
            return value;
        } else {
            // NOTE: just naively assuming that `value` has an invalid NameStartChar
            return "_INVALID_NAMESTARTCHAR_" + value;
        }
    }

    function __sanitize(value) {
        return value.sanitize(
            ['&',    '<',   '>',    '\'',    '"'],
            ['&amp;','&lt;','&gt;', '&apos;','&quot;']
        );
    };

    function __toXML(o, t) {
        var xml = []
          , a   = []
          , p, i, len
        ;

        switch( typeof o ) {
            case "object":
                if( null === o ) {
                    xml.push("<{tag}/>".substitute({"tag":__tagize(t)}));
                } else if(o.length) {
                    a = o;
                    if(a.length === 0) {
                        xml.push("<{tag}/>".substitute({"tag":__tagize(t)}));
                    } else {
                        for(i = 0, len = a.length; i < len; i++) {
                            xml.push(__toXML(a[i], t));
                        }
                    }
                } else {
                    xml.push("<{tag}".substitute({"tag":__tagize(t)}));
                    a = [];
                    for(p in o) {
                        if(o.hasOwnProperty(p)) {
                            if(p.charAt(0) === "@") xml.push(" {param}='{content}'".substitute({"param":__tagize(p.substr(1)), "content":__sanitize(o[p].toString())}));
                            else a.push(p);
                        }
                    }
                    if(a.length === 0) {
                        xml.push("/>");
                    } else {
                        xml.push(">");
                        for(i = 0, len = a.length; i < len; i++) {
                            p = a[i];
                            if(p === "#text") {
                                xml.push(__sanitize(o[p]));
                            } else if(p === "#cdata") {
                                xml.push("<![CDATA[{code}]]>".substitute({"code": o[p]}));
                            } else {
                                xml.push(__toXML(o[p], p));
                            }
                        }
                        xml.push("</{tag}>".substitute({"tag":__tagize(t)}));
                    }
                }
                break;
            
            default:
                var s = String(o);
                if(s.length === 0) {
                    xml.push("<{tag}/>".substitute({"tag":__tagize(t)}));
                } else {
                    xml.push("<{tag}>{value}</{tag}>".substitute({"tag":__tagize(t), "value":__sanitize(s)}));
                }
        }
        return xml.join('');
    }

    if(typeof JSON.toXML !== 'function') {
        JSON.toXML = function(o, container){
            //container = container || "";
            var xml = [];
            if(container) xml.push("<{tag}>".substitute({"tag":container}));
            for(var p in o) {
                if(o.hasOwnProperty(p)) {
                    xml.push(__toXML(o[p], p));
                }
            }
            if(container) xml.push("</{tag}>".substitute({"tag":container}));
            return xml.join('');
        }
    }

})();










// Implement minify ( strip comments from json-like configuration files )
(function(){
    if(typeof JSON.minify !== 'function') {
        JSON.minify = function(json) {
            var tokenizer = /"|(\/\*)|(\*\/)|(\/\/)|\n|\r/g,
                in_string = false,
                in_multiline_comment = false,
                in_singleline_comment = false,
                tmp, tmp2, new_str = [], ns = 0, from = 0, lc, rc;
            
            tokenizer.lastIndex = 0;
            
            while( tmp = tokenizer.exec(json) ) {
                lc = RegExp.leftContext;
                rc = RegExp.rightContext;
                if(!in_multiline_comment && !in_singleline_comment) {
                    tmp2 = lc.substring(from);
                    if(!in_string) {
                        tmp2 = tmp2.replace(/(\n|\r|\s)*/g,"");
                    }
                    new_str[ns++] = tmp2;
                }
                from = tokenizer.lastIndex;
                
                if(tmp[0] == "\"" && !in_multiline_comment && !in_singleline_comment) {
                    tmp2 = lc.match(/(\\)*$/);
                    if(!in_string || !tmp2 || (tmp2[0].length % 2) == 0) { // start of string with ", or unescaped " character found to end string
                        in_string = !in_string;
                    }
                    from--; // include " character in next catch
                    rc = json.substring(from);
                } else if(tmp[0] == "/*" && !in_string && !in_multiline_comment && !in_singleline_comment) {
                    in_multiline_comment = true;
                } else if(tmp[0] == "*/" && !in_string && in_multiline_comment && !in_singleline_comment) {
                    in_multiline_comment = false;
                } else if(tmp[0] == "//" && !in_string && !in_multiline_comment && !in_singleline_comment) {
                    in_singleline_comment = true;
                } else if((tmp[0] == "\n" || tmp[0] == "\r") && !in_string && !in_multiline_comment && in_singleline_comment) {
                    in_singleline_comment = false;
                } else if(!in_multiline_comment && !in_singleline_comment && !(/\n|\r|\s/.test(tmp[0]))) {
                    new_str[ns++] = tmp[0];
                }
            }
            new_str[ns++] = rc;
            return new_str.join("");
        }
    }
})();

</script>
