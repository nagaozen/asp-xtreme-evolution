<%

' File: base64.asp
' 
' AXE(ASP Xtreme Evolution) base64 utility.
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



' Class: Base64
' 
' Base64 is a way of representing binary data using alphanumeric characters only,
' usually used for transmitting binary over a text channel such as email, xml
' and json.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Oct 2008
' 
class Base64
    
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
    
    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "1.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: getBinary
    ' 
    ' Fetch for a binary in the given path and return its value as a Stream.
    ' 
    ' Parameters:
    ' 
    '     (string) - Binary full path location
    ' 
    ' Returns:
    ' 
    '     (application/octet-stream) - Binary data
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Translator : set Translator = new Base64
    ' Response.contentType = "image/jpeg"
    ' Response.binaryWrite( Translator.getBinary( Server.mapPath("apple-ipod-touch.jpg") ) )
    ' set Translator = nothing
    ' 
    ' (end code)
    ' 
    public function getBinary(path)
        dim Stream : set Stream = Server.createObject("ADODB.Stream")
        Stream.type = 1' adTypeBinary
        Stream.open()
        Stream.loadFromFile(path)
        getBinary = Stream.read()
        Stream.close()
        set Stream = nothing
    end function
    
    ' Function: encode
    ' 
    ' Converts a binary into a base64 data.
    ' 
    ' Parameters:
    ' 
    '     (bytearray) - Data to encode
    ' 
    ' Returns:
    ' 
    '     (base64) - Encoded data
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim encoded, Translator : set Translator = new Base64
    ' encoded = Translator.encode( Translator.getBinary( Server.mapPath("apple-ipod-touch.jpg") ) )
    ' Response.write(encoded)
    ' set Translator = nothing
    ' 
    ' (end code)
    ' 
    public function encode(bin)
        dim Xml : set Xml = Server.createObject("MSXML2.DOMDocument.6.0")
        dim Node : set Node = Xml.createElement("data")
        Node.dataType = "bin.base64"
        Node.nodeTypedValue = bin
        encode = Node.text
        set Node = nothing
        set Xml = nothing
    end function
    
    ' Function: decode
    ' 
    ' Converts base64 into a binary data.
    ' 
    ' Parameters:
    ' 
    '     (base64) - Data do decode
    ' 
    ' Returns:
    ' 
    '     (bytearray) - Decoded data
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim encoded, decoded, Translator
    ' 
    ' set Translator = new Base64
    ' encoded = Translator.encode( Translator.getBinary( Server.mapPath("apple-ipod-touch.jpg") ) )
    ' decoded = Translator.decode( encoded )
    ' Response.contentType = "image/jpeg"
    ' Response.binaryWrite( decoded )
    ' set Translator = nothing
    ' 
    ' (end code)
    ' 
    public function decode(base64)
        dim Xml : set Xml = Server.createObject("MSXML2.DOMDocument.6.0")
        dim Node : set Node = Xml.createElement("data")
        Node.dataType = "bin.base64"
        Node.text = base64
        decode = Node.nodeTypedValue
        set Node = nothing
        set Xml = nothing
    end function
    
    ' Function: encodedSize
    ' 
    ' Computes the size of the base64 data in Kilobytes (KB).
    ' 
    ' Parameters:
    ' 
    '     (base64) - Encoded data
    ' 
    ' Returns:
    ' 
    '     (float) - Data size
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim encoded, Translator : set Translator = new Base64
    ' encoded = Translator.encode( Translator.getBinary( Server.mapPath("image.jpg") ) )
    ' Response.write(Translator.encodedSize(encoded) & " KB")
    ' set Translator = nothing
    ' 
    ' (end code)
    ' 
    public function encodedSize(base64)
        encodedSize = len(base64) / 1024
    end function
    
    ' Function: decodedSize
    ' 
    ' Computes the size of the binary data in Kilobytes (KB).
    ' 
    ' Parameters:
    ' 
    '     (bytearray) - Binary data
    ' 
    ' Returns:
    ' 
    '     (float) - Data size
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim encoded, decoded, Translator
    ' 
    ' set Translator = new Base64
    ' encoded = Translator.encode( Translator.getBinary( Server.mapPath("apple-ipod-touch.jpg") ) )
    ' decoded = Translator.decode( encoded )
    ' Response.write(Translator.decodedSize(decoded) & " KB")
    ' set Translator = nothing
    ' 
    ' (end code)
    ' 
    public function decodedSize(bin)
        decodedSize = lenb(bin) / 1024
    end function
    
end class

%>
