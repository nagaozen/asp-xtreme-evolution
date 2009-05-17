<%

'+-----------------------------------------------------------------------------+
'|This file is part of ASP Xtreme Evolution.                                   |
'|Copyright (C) 2007, 2008 Fabio Zendhi Nagao                                  |
'|                                                                             |
'|ASP Xtreme Evolution is free software: you can redistribute it and/or modify |
'|it under the terms of the GNU Lesser General Public License as published by  |
'|the Free Software Foundation, either version 3 of the License, or            |
'|(at your option) any later version.                                          |
'|                                                                             |
'|ASP Xtreme Evolution is distributed in the hope that it will be useful,      |
'|but WITHOUT ANY WARRANTY; without even the implied warranty of               |
'|MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                |
'|GNU Lesser General Public License for more details.                          |
'|                                                                             |
'|You should have received a copy of the GNU Lesser General Public License     |
'|along with ASP Xtreme Evolution.  If not, see <http://www.gnu.org/licenses/>.|
'+-----------------------------------------------------------------------------+

' Class: MD5
' 
' ASP VBScript code for generating an MD5 'digest' or 'signature' of a string.
' The MD5 algorithm is one of the industry standard methods for generating
' digital signatures. It is generically known as a digest, digital signature,
' one-way encryption, hash or checksum algorithm. A common use for MD5 is for
' password encryption as it is one-way in nature, that does not mean that your
' passwords are not free from a dictionary attack.
' 
' License:
' 
' This is 'free' software with the following restrictions
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are
' free to use the source code in your own code, but you may not claim that you
' created the sample code. It is expressly forbidden to sell or profit from this
' source code other than by the knowledge gained or the enhanced value added by
' your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on this code provided
' here, or any consultancy work, please do not hesitate to contact us.
'
' About:
' 
'   - Original work by Phil Fresle <http://www.frez.co.uk> @ 2001
'   - Class structure and Convert functions by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ February 2008
' 
class MD5
    
    private m_lOnBits(30)
    private m_l2Power(30)
    
    private BITS_TO_A_BYTE
    private BYTES_TO_A_WORD
    private BITS_TO_A_WORD
    
    private sub Class_initialize()
        BITS_TO_A_BYTE  = 8
        BYTES_TO_A_WORD = 4
        BITS_TO_A_WORD  = 32
        
        m_lOnBits(0)  = 1
        m_lOnBits(1)  = 3
        m_lOnBits(2)  = 7
        m_lOnBits(3)  = 15
        m_lOnBits(4)  = 31
        m_lOnBits(5)  = 63
        m_lOnBits(6)  = 127
        m_lOnBits(7)  = 255
        m_lOnBits(8)  = 511
        m_lOnBits(9)  = 1023
        m_lOnBits(10) = 2047
        m_lOnBits(11) = 4095
        m_lOnBits(12) = 8191
        m_lOnBits(13) = 16383
        m_lOnBits(14) = 32767
        m_lOnBits(15) = 65535
        m_lOnBits(16) = 131071
        m_lOnBits(17) = 262143
        m_lOnBits(18) = 524287
        m_lOnBits(19) = 1048575
        m_lOnBits(20) = 2097151
        m_lOnBits(21) = 4194303
        m_lOnBits(22) = 8388607
        m_lOnBits(23) = 16777215
        m_lOnBits(24) = 33554431
        m_lOnBits(25) = 67108863
        m_lOnBits(26) = 134217727
        m_lOnBits(27) = 268435455
        m_lOnBits(28) = 536870911
        m_lOnBits(29) = 1073741823
        m_lOnBits(30) = 2147483647
        
        m_l2Power(0)  = 1
        m_l2Power(1)  = 2
        m_l2Power(2)  = 4
        m_l2Power(3)  = 8
        m_l2Power(4)  = 16
        m_l2Power(5)  = 32
        m_l2Power(6)  = 64
        m_l2Power(7)  = 128
        m_l2Power(8)  = 256
        m_l2Power(9)  = 512
        m_l2Power(10) = 1024
        m_l2Power(11) = 2048
        m_l2Power(12) = 4096
        m_l2Power(13) = 8192
        m_l2Power(14) = 16384
        m_l2Power(15) = 32768
        m_l2Power(16) = 65536
        m_l2Power(17) = 131072
        m_l2Power(18) = 262144
        m_l2Power(19) = 524288
        m_l2Power(20) = 1048576
        m_l2Power(21) = 2097152
        m_l2Power(22) = 4194304
        m_l2Power(23) = 8388608
        m_l2Power(24) = 16777216
        m_l2Power(25) = 33554432
        m_l2Power(26) = 67108864
        m_l2Power(27) = 134217728
        m_l2Power(28) = 268435456
        m_l2Power(29) = 536870912
        m_l2Power(30) = 1073741824
    end sub
    
    private sub Class_terminate()
        erase m_lOnBits
        erase m_l2Power
    end sub
    
    private function LShift(lValue, iShiftBits)
        if iShiftBits = 0 then
            LShift = lValue
            exit function
        elseif iShiftBits = 31 then
            if lValue and 1 then
                LShift = &H80000000
            else
                LShift = 0
            end if
            exit function
        elseif iShiftBits < 0 or iShiftBits > 31 then
            Err.Raise 6
        end if
        
        if (lValue and m_l2Power(31 - iShiftBits)) then
            LShift = ((lValue and m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) or &H80000000
        else
            LShift = ((lValue and m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
        end if
    end function
    
    private function RShift(lValue, iShiftBits)
        if iShiftBits = 0 then
            RShift = lValue
            exit function
        elseif iShiftBits = 31 then
            if lValue and &H80000000 then
                RShift = 1
            else
                RShift = 0
            end if
            exit function
        elseif iShiftBits < 0 or iShiftBits > 31 then
            Err.Raise 6
        end if
        
        RShift = (lValue and &H7FFFFFFE) \ m_l2Power(iShiftBits)
        
        if (lValue and &H80000000) then
            RShift = (RShift or (&H40000000 \ m_l2Power(iShiftBits - 1)))
        end if
    end function
    
    private function RotateLeft(lValue, iShiftBits)
        RotateLeft = LShift(lValue, iShiftBits) or RShift(lValue, (32 - iShiftBits))
    end function
    
    private function AddUnsigned(lX, lY)
        dim lX4
        dim lY4
        dim lX8
        dim lY8
        dim lResult
        
        lX8 = lX and &H80000000
        lY8 = lY and &H80000000
        lX4 = lX and &H40000000
        lY4 = lY and &H40000000
        
        lResult = (lX and &H3FFFFFFF) + (lY and &H3FFFFFFF)
        
        if lX4 and lY4 then
            lResult = lResult xor &H80000000 xor lX8 xor lY8
        elseif lX4 or lY4 then
            if lResult and &H40000000 then
                lResult = lResult xor &HC0000000 xor lX8 xor lY8
            else
                lResult = lResult xor &H40000000 xor lX8 xor lY8
            end if
        else
            lResult = lResult xor lX8 xor lY8
        end if
        
        AddUnsigned = lResult
    end function
    
    private function F(x, y, z)
        F = (x and y) or ((not x) and z)
    end function
    
    private function G(x, y, z)
        G = (x and z) or (y and (not z))
    end function
    
    private function H(x, y, z)
        H = (x xor y xor z)
    end function
    
    private function I(x, y, z)
        I = (y xor (x or (not z)))
    end function
    
    private sub FF(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    end sub
    
    private sub GG(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    end sub
    
    private sub HH(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    end sub
    
    private sub II(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    end sub
    
    private function ConvertToWordArray(sMessage)
        dim lMessageLength
        dim lNumberOfWords
        dim lWordArray()
        dim lBytePosition
        dim lByteCount
        dim lWordCount
        
        const MODULUS_BITS = 512
        const CONGRUENT_BITS = 448
        
        lMessageLength = Len(sMessage)
        
        lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
        redim lWordArray(lNumberOfWords - 1)
        
        lBytePosition = 0
        lByteCount = 0
        do until lByteCount >= lMessageLength
            lWordCount = lByteCount \ BYTES_TO_A_WORD
            lBytePosition = (lByteCount mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
            lWordArray(lWordCount) = lWordArray(lWordCount) or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
            lByteCount = lByteCount + 1
        loop
        
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        
        lWordArray(lWordCount) = lWordArray(lWordCount) or LShift(&H80, lBytePosition)
        
        lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
        lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
        
        ConvertToWordArray = lWordArray
    end function
    
    private function WordToHex(lValue)
        dim lByte
        dim lCount
        
        for lCount = 0 to 3
            lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) and m_lOnBits(BITS_TO_A_BYTE - 1)
            WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
        next
    end function
    
    ' Function: encryptData
    ' 
    ' Use this method to encrypt your data.
    ' 
    ' Parameters:
    ' 
    '     (string) - Data to be encrypted
    ' 
    ' Returns:
    ' 
    '     (string) - Encrypted version of the input data
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim message : message = "This is a very secret message"
    ' dim Encryptor
    ' set Encryptor = new MD5
    ' Response.write("Plain: " & message & "<br />")
    ' Response.write("Encrypted: " & Encryptor.encryptData(message))
    ' set Encryptor = nothing
    ' 
    ' (end code)
    ' 
    public function encryptData(sMessage)
        dim x
        dim k
        dim AA
        dim BB
        dim CC
        dim DD
        dim a
        dim b
        dim c
        dim d
        
        const S11 = 7
        const S12 = 12
        const S13 = 17
        const S14 = 22
        const S21 = 5
        const S22 = 9
        const S23 = 14
        const S24 = 20
        const S31 = 4
        const S32 = 11
        const S33 = 16
        const S34 = 23
        const S41 = 6
        const S42 = 10
        const S43 = 15
        const S44 = 21
        
        x = ConvertToWordArray(sMessage)
        
        a = &H67452301
        b = &HEFCDAB89
        c = &H98BADCFE
        d = &H10325476
        
        for k = 0 to UBound(x) step 16
            AA = a
            BB = b
            CC = c
            DD = d
            
            FF a, b, c, d, x(k + 0), S11, &HD76AA478
            FF d, a, b, c, x(k + 1), S12, &HE8C7B756
            FF c, d, a, b, x(k + 2), S13, &H242070DB
            FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
            FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
            FF d, a, b, c, x(k + 5), S12, &H4787C62A
            FF c, d, a, b, x(k + 6), S13, &HA8304613
            FF b, c, d, a, x(k + 7), S14, &HFD469501
            FF a, b, c, d, x(k + 8), S11, &H698098D8
            FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
            FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
            FF b, c, d, a, x(k + 11), S14, &H895CD7BE
            FF a, b, c, d, x(k + 12), S11, &H6B901122
            FF d, a, b, c, x(k + 13), S12, &HFD987193
            FF c, d, a, b, x(k + 14), S13, &HA679438E
            FF b, c, d, a, x(k + 15), S14, &H49B40821
            
            GG a, b, c, d, x(k + 1), S21, &HF61E2562
            GG d, a, b, c, x(k + 6), S22, &HC040B340
            GG c, d, a, b, x(k + 11), S23, &H265E5A51
            GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
            GG a, b, c, d, x(k + 5), S21, &HD62F105D
            GG d, a, b, c, x(k + 10), S22, &H2441453
            GG c, d, a, b, x(k + 15), S23, &HD8A1E681
            GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
            GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
            GG d, a, b, c, x(k + 14), S22, &HC33707D6
            GG c, d, a, b, x(k + 3), S23, &HF4D50D87
            GG b, c, d, a, x(k + 8), S24, &H455A14ED
            GG a, b, c, d, x(k + 13), S21, &HA9E3E905
            GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
            GG c, d, a, b, x(k + 7), S23, &H676F02D9
            GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
            
            HH a, b, c, d, x(k + 5), S31, &HFFFA3942
            HH d, a, b, c, x(k + 8), S32, &H8771F681
            HH c, d, a, b, x(k + 11), S33, &H6D9D6122
            HH b, c, d, a, x(k + 14), S34, &HFDE5380C
            HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
            HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
            HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
            HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
            HH a, b, c, d, x(k + 13), S31, &H289B7EC6
            HH d, a, b, c, x(k + 0), S32, &HEAA127FA
            HH c, d, a, b, x(k + 3), S33, &HD4EF3085
            HH b, c, d, a, x(k + 6), S34, &H4881D05
            HH a, b, c, d, x(k + 9), S31, &HD9D4D039
            HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
            HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
            HH b, c, d, a, x(k + 2), S34, &HC4AC5665
            
            II a, b, c, d, x(k + 0), S41, &HF4292244
            II d, a, b, c, x(k + 7), S42, &H432AFF97
            II c, d, a, b, x(k + 14), S43, &HAB9423A7
            II b, c, d, a, x(k + 5), S44, &HFC93A039
            II a, b, c, d, x(k + 12), S41, &H655B59C3
            II d, a, b, c, x(k + 3), S42, &H8F0CCC92
            II c, d, a, b, x(k + 10), S43, &HFFEFF47D
            II b, c, d, a, x(k + 1), S44, &H85845DD1
            II a, b, c, d, x(k + 8), S41, &H6FA87E4F
            II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
            II c, d, a, b, x(k + 6), S43, &HA3014314
            II b, c, d, a, x(k + 13), S44, &H4E0811A1
            II a, b, c, d, x(k + 4), S41, &HF7537E82
            II d, a, b, c, x(k + 11), S42, &HBD3AF235
            II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
            II b, c, d, a, x(k + 9), S44, &HEB86D391
            
            a = AddUnsigned(a, AA)
            b = AddUnsigned(b, BB)
            c = AddUnsigned(c, CC)
            d = AddUnsigned(d, DD)
        next
        
        encryptData = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
    end function
    
end class

%>
