<%

'+-----------------------------------------------------------------------------+
'|This file is part of ASP Xtreme Evolution.                                   |
'|Copyright (C) 2007, 2009 Fabio Zendhi Nagao                                  |
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

' Class: SHA256
' 
' ASP VBScript code for generating a SHA256 'digest' or 'signature' of a string.
' The SHA256 algorithm is one of the industry standard methods for generating
' digital signatures. It is generically known as a digest, digital signature,
' one-way encryption, hash or checksum algorithm. A common use for SHA256 is for
' password encryption as it is one-way in nature, that does not mean that your
' passwords are not free from a dictionary attack. 
'
' if you are using the routine for passwords, you can make it a little more
' secure by concatenating some known random characters to the password before
' you generate the signature and on subsequent tests, so even if a hacker knows
' you are using SHA-256 for your passwords, the random characters will make it
' harder to dictionary attack.
'
' Due to the way in which the string is processed the routine assumes a
' single byte character set. VB passes unicode (2-byte) character strings, the
' ConvertToWordArray function uses on the first byte for each character. This
' has been done this way for ease of use, to make the routine truely portable
' you could accept a byte array instead, it would then be up to the calling
' routine to make sure that the byte array is generated from their string in
' a manner consistent with the string type.
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
'   - Class structure by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ February 2008
' 
class SHA256
    
    private m_lOnBits(30)
    private m_l2Power(30)
    private K(63)
    
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
        
        K(0)  = &H428A2F98
        K(1)  = &H71374491
        K(2)  = &HB5C0FBCF
        K(3)  = &HE9B5DBA5
        K(4)  = &H3956C25B
        K(5)  = &H59F111F1
        K(6)  = &H923F82A4
        K(7)  = &HAB1C5ED5
        K(8)  = &HD807AA98
        K(9)  = &H12835B01
        K(10) = &H243185BE
        K(11) = &H550C7DC3
        K(12) = &H72BE5D74
        K(13) = &H80DEB1FE
        K(14) = &H9BDC06A7
        K(15) = &HC19BF174
        K(16) = &HE49B69C1
        K(17) = &HEFBE4786
        K(18) = &HFC19DC6
        K(19) = &H240CA1CC
        K(20) = &H2DE92C6F
        K(21) = &H4A7484AA
        K(22) = &H5CB0A9DC
        K(23) = &H76F988DA
        K(24) = &H983E5152
        K(25) = &HA831C66D
        K(26) = &HB00327C8
        K(27) = &HBF597FC7
        K(28) = &HC6E00BF3
        K(29) = &HD5A79147
        K(30) = &H6CA6351
        K(31) = &H14292967
        K(32) = &H27B70A85
        K(33) = &H2E1B2138
        K(34) = &H4D2C6DFC
        K(35) = &H53380D13
        K(36) = &H650A7354
        K(37) = &H766A0ABB
        K(38) = &H81C2C92E
        K(39) = &H92722C85
        K(40) = &HA2BFE8A1
        K(41) = &HA81A664B
        K(42) = &HC24B8B70
        K(43) = &HC76C51A3
        K(44) = &HD192E819
        K(45) = &HD6990624
        K(46) = &HF40E3585
        K(47) = &H106AA070
        K(48) = &H19A4C116
        K(49) = &H1E376C08
        K(50) = &H2748774C
        K(51) = &H34B0BCB5
        K(52) = &H391C0CB3
        K(53) = &H4ED8AA4A
        K(54) = &H5B9CCA4F
        K(55) = &H682E6FF3
        K(56) = &H748F82EE
        K(57) = &H78A5636F
        K(58) = &H84C87814
        K(59) = &H8CC70208
        K(60) = &H90BEFFFA
        K(61) = &HA4506CEB
        K(62) = &HBEF9A3F7
        K(63) = &HC67178F2
    end sub
    
    private sub Class_terminate()
        erase m_lOnBits
        erase m_l2Power
        erase K
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
    
    private function Ch(x, y, z)
        Ch = ((x and y) xor ((not x) and z))
    end function
    
    private function Maj(x, y, z)
        Maj = ((x and y) xor (x and z) xor (y and z))
    end function
    
    private function S(x, n)
        S = (RShift(x, (n and m_lOnBits(4))) or LShift(x, (32 - (n and m_lOnBits(4)))))
    end function
    
    private function R(x, n)
        R = RShift(x, CInt(n and m_lOnBits(4)))
    end function
    
    private function Sigma0(x)
        Sigma0 = (S(x, 2) xor S(x, 13) xor S(x, 22))
    end function
    
    private function Sigma1(x)
        Sigma1 = (S(x, 6) xor S(x, 11) xor S(x, 25))
    end function
    
    private function Gamma0(x)
        Gamma0 = (S(x, 7) xor S(x, 18) xor R(x, 3))
    end function
    
    private function Gamma1(x)
        Gamma1 = (S(x, 17) xor S(x, 19) xor R(x, 10))
    end function
    
    private function ConvertToWordArray(sMessage)
        dim lMessageLength
        dim lNumberOfWords
        dim lWordArray()
        dim lBytePosition
        dim lByteCount
        dim lWordCount
        dim lByte
        
        const MODULUS_BITS = 512
        const CONGRUENT_BITS = 448
        
        lMessageLength = Len(sMessage)
        
        lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
        redim lWordArray(lNumberOfWords - 1)
        
        lBytePosition = 0
        lByteCount = 0
        do until lByteCount >= lMessageLength
            lWordCount = lByteCount \ BYTES_TO_A_WORD
            
            lBytePosition = (3 - (lByteCount mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
            
            lByte = AscB(Mid(sMessage, lByteCount + 1, 1))
            
            lWordArray(lWordCount) = lWordArray(lWordCount) or LShift(lByte, lBytePosition)
            lByteCount = lByteCount + 1
        loop
        
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (3 - (lByteCount mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        
        lWordArray(lWordCount) = lWordArray(lWordCount) or LShift(&H80, lBytePosition)
        
        lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
        lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
        
        ConvertToWordArray = lWordArray
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
    ' set Encryptor = new SHA256
    ' Response.write("Plain: " & message & "<br />")
    ' Response.write("Encrypted: " & Encryptor.encryptData(message))
    ' set Encryptor = nothing
    ' 
    ' (end code)
    ' 
    public function encryptData(sMessage)
        dim HASH(7)
        dim M
        dim W(63)
        dim a
        dim b
        dim c
        dim d
        dim e
        dim f
        dim g
        dim h
        dim i
        dim j
        dim T1
        dim T2
        
        HASH(0) = &H6A09E667
        HASH(1) = &HBB67AE85
        HASH(2) = &H3C6EF372
        HASH(3) = &HA54FF53A
        HASH(4) = &H510E527F
        HASH(5) = &H9B05688C
        HASH(6) = &H1F83D9AB
        HASH(7) = &H5BE0CD19
        
        M = ConvertToWordArray(sMessage)
        
        for i = 0 to UBound(M) step 16
            a = HASH(0)
            b = HASH(1)
            c = HASH(2)
            d = HASH(3)
            e = HASH(4)
            f = HASH(5)
            g = HASH(6)
            h = HASH(7)
            
            for j = 0 to 63
                if j < 16 then
                    W(j) = M(j + i)
                else
                    W(j) = AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(W(j - 2)), W(j - 7)), Gamma0(W(j - 15))), W(j - 16))
                end if
                    
                T1 = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(h, Sigma1(e)), Ch(e, f, g)), K(j)), W(j))
                T2 = AddUnsigned(Sigma0(a), Maj(a, b, c))
                
                h = g
                g = f
                f = e
                e = AddUnsigned(d, T1)
                d = c
                c = b
                b = a
                a = AddUnsigned(T1, T2)
            next
            
            HASH(0) = AddUnsigned(a, HASH(0))
            HASH(1) = AddUnsigned(b, HASH(1))
            HASH(2) = AddUnsigned(c, HASH(2))
            HASH(3) = AddUnsigned(d, HASH(3))
            HASH(4) = AddUnsigned(e, HASH(4))
            HASH(5) = AddUnsigned(f, HASH(5))
            HASH(6) = AddUnsigned(g, HASH(6))
            HASH(7) = AddUnsigned(h, HASH(7))
        next
        
        encryptData = LCase(Right("00000000" & Hex(HASH(0)), 8) & Right("00000000" & Hex(HASH(1)), 8) & Right("00000000" & Hex(HASH(2)), 8) & Right("00000000" & Hex(HASH(3)), 8) & Right("00000000" & Hex(HASH(4)), 8) & Right("00000000" & Hex(HASH(5)), 8) & Right("00000000" & Hex(HASH(6)), 8) & Right("00000000" & Hex(HASH(7)), 8))
    end function
    
end class

%>
