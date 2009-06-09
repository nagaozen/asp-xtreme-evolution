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

' Class: Rijndael
' 
' Implementation of the AES Rijndael Block Cipher. Inspired by Mike Scott's
' implementation in C.
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
class Rijndael
    
    private m_lOnBits(30)
    private m_l2Power(30)
    private m_bytOnBits(7)
    private m_byt2Power(7)
    
    private m_InCo(3)
    
    private m_fbsub(255)
    private m_rbsub(255)
    private m_ptab(255)
    private m_ltab(255)
    private m_ftable(255)
    private m_rtable(255)
    private m_rco(29)
    
    private m_Nk
    private m_Nb
    private m_Nr
    private m_fi(23)
    private m_ri(23)
    private m_fkey(119)
    private m_rkey(119)
    
    private sub Class_initialize()
        m_InCo(0) = &HB
        m_InCo(1) = &HD
        m_InCo(2) = &H9
        m_InCo(3) = &HE
        
        m_bytOnBits(0) = 1
        m_bytOnBits(1) = 3
        m_bytOnBits(2) = 7
        m_bytOnBits(3) = 15
        m_bytOnBits(4) = 31
        m_bytOnBits(5) = 63
        m_bytOnBits(6) = 127
        m_bytOnBits(7) = 255
        
        m_byt2Power(0) = 1
        m_byt2Power(1) = 2
        m_byt2Power(2) = 4
        m_byt2Power(3) = 8
        m_byt2Power(4) = 16
        m_byt2Power(5) = 32
        m_byt2Power(6) = 64
        m_byt2Power(7) = 128
        
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
        erase m_bytOnBits
        erase m_byt2Power
        
        erase m_InCo
        
        erase m_fbsub
        erase m_rbsub
        erase m_ptab
        erase m_ltab
        erase m_ftable
        erase m_rtable
        erase m_rco
        
        erase m_fi
        erase m_ri
        erase m_fkey
        erase m_rkey
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
    
    private function LShiftByte(bytValue, bytShiftBits)
        if bytShiftBits = 0 then
            LShiftByte = bytValue
            exit function
        elseif bytShiftBits = 7 then
            if bytValue and 1 then
                LShiftByte = &H80
            else
                LShiftByte = 0
            end if
            exit function
        elseif bytShiftBits < 0 or bytShiftBits > 7 then
            Err.Raise 6
        end if
        
        LShiftByte = ((bytValue and m_bytOnBits(7 - bytShiftBits)) * m_byt2Power(bytShiftBits))
    end function
    
    private function RShiftByte(bytValue, bytShiftBits)
        if bytShiftBits = 0 then
            RShiftByte = bytValue
            exit function
        elseif bytShiftBits = 7 then
            if bytValue and &H80 then
                RShiftByte = 1
            else
                RShiftByte = 0
            end if
            exit function
        elseif bytShiftBits < 0 or bytShiftBits > 7 then
            Err.Raise 6
        end if
        
        RShiftByte = bytValue \ m_byt2Power(bytShiftBits)
    end function
    
    private function RotateLeft(lValue, iShiftBits)
        RotateLeft = LShift(lValue, iShiftBits) or RShift(lValue, (32 - iShiftBits))
    end function
    
    private function RotateLeftByte(bytValue, bytShiftBits)
        RotateLeftByte = LShiftByte(bytValue, bytShiftBits) or RShiftByte(bytValue, (8 - bytShiftBits))
    end function
    
    private function Pack(b())
        dim lCount
        dim lTemp
        
        for lCount = 0 to 3
            lTemp = b(lCount)
            Pack = Pack or LShift(lTemp, (lCount * 8))
        next
    end function
    
    private function PackFrom(b(), k)
        dim lCount
        dim lTemp
        
        for lCount = 0 to 3
            lTemp = b(lCount + k)
            PackFrom = PackFrom or LShift(lTemp, (lCount * 8))
        next
    end function
    
    private sub Unpack(a, b())
        b(0) = a and m_lOnBits(7)
        b(1) = RShift(a, 8) and m_lOnBits(7)
        b(2) = RShift(a, 16) and m_lOnBits(7)
        b(3) = RShift(a, 24) and m_lOnBits(7)
    end sub
    
    private sub UnpackFrom(a, b(), k)
        b(0 + k) = a and m_lOnBits(7)
        b(1 + k) = RShift(a, 8) and m_lOnBits(7)
        b(2 + k) = RShift(a, 16) and m_lOnBits(7)
        b(3 + k) = RShift(a, 24) and m_lOnBits(7)
    end sub
    
    private function xtime(a)
        dim b
        
        if (a and &H80) then
            b = &H1B
        else
            b = 0
        end if
        
        xtime = LShiftByte(a, 1)
        xtime = xtime xor b
    end function
    
    private function bmul(x, y)
        if x <> 0 and y <> 0 then
            bmul = m_ptab((CLng(m_ltab(x)) + CLng(m_ltab(y))) mod 255)
        else
            bmul = 0
        end if
    end function
    
    private function subByte(a)
        dim b(3)
        
        Unpack a, b
        b(0) = m_fbsub(b(0))
        b(1) = m_fbsub(b(1))
        b(2) = m_fbsub(b(2))
        b(3) = m_fbsub(b(3))
        
        subByte = Pack(b)
    end function
    
    private function product(x, y)
        dim xb(3)
        dim yb(3)
        
        Unpack x, xb
        Unpack y, yb
        product = bmul(xb(0), yb(0)) xor bmul(xb(1), yb(1)) xor bmul(xb(2), yb(2)) xor bmul(xb(3), yb(3))
    end function
    
    private function InvMixCol(x)
        dim y
        dim m
        dim b(3)
        
        m = Pack(m_InCo)
        b(3) = product(m, x)
        m = RotateLeft(m, 24)
        b(2) = product(m, x)
        m = RotateLeft(m, 24)
        b(1) = product(m, x)
        m = RotateLeft(m, 24)
        b(0) = product(m, x)
        y = Pack(b)
        
        InvMixCol = y
    end function
    
    private function ByteSub(x)
        dim y
        dim z
        
        z = x
        y = m_ptab(255 - m_ltab(z))
        z = y
        z = RotateLeftByte(z, 1)
        y = y xor z
        z = RotateLeftByte(z, 1)
        y = y xor z
        z = RotateLeftByte(z, 1)
        y = y xor z
        z = RotateLeftByte(z, 1)
        y = y xor z
        y = y xor &H63
        
        ByteSub = y
    end function
    
    private sub gentables()
        dim i
        dim y
        dim b(3)
        dim ib
        
        m_ltab(0) = 0
        m_ptab(0) = 1
        m_ltab(1) = 0
        m_ptab(1) = 3
        m_ltab(3) = 1
        
        for i = 2 to 255
            m_ptab(i) = m_ptab(i - 1) xor xtime(m_ptab(i - 1))
            m_ltab(m_ptab(i)) = i
        next
        
        m_fbsub(0) = &H63
        m_rbsub(&H63) = 0
        
        for i = 1 to 255
            ib = i
            y = ByteSub(ib)
            m_fbsub(i) = y
            m_rbsub(y) = i
        next
        
        y = 1
        for i = 0 to 29
            m_rco(i) = y
            y = xtime(y)
        next
        
        for i = 0 to 255
            y = m_fbsub(i)
            b(3) = y xor xtime(y)
            b(2) = y
            b(1) = y
            b(0) = xtime(y)
            m_ftable(i) = Pack(b)
            
            y = m_rbsub(i)
            b(3) = bmul(m_InCo(0), y)
            b(2) = bmul(m_InCo(1), y)
            b(1) = bmul(m_InCo(2), y)
            b(0) = bmul(m_InCo(3), y)
            m_rtable(i) = Pack(b)
        next
    end sub
    
    private sub gkey(nb, nk, key())
        dim i
        dim j
        dim k
        dim m
        dim N
        dim C1
        dim C2
        dim C3
        dim CipherKey(7)
        
        m_Nb = nb
        m_Nk = nk
        
        if m_Nb >= m_Nk then
            m_Nr = 6 + m_Nb
        else
            m_Nr = 6 + m_Nk
        end if
        
        C1 = 1
        if m_Nb < 8 then
            C2 = 2
            C3 = 3
        else
            C2 = 3
            C3 = 4
        end if
        
        for j = 0 to nb - 1
            m = j * 3
            
            m_fi(m) = (j + C1) mod nb
            m_fi(m + 1) = (j + C2) mod nb
            m_fi(m + 2) = (j + C3) mod nb
            m_ri(m) = (nb + j - C1) mod nb
            m_ri(m + 1) = (nb + j - C2) mod nb
            m_ri(m + 2) = (nb + j - C3) mod nb
        next
        
        N = m_Nb * (m_Nr + 1)
        
        for i = 0 to m_Nk - 1
            j = i * 4
            CipherKey(i) = PackFrom(key, j)
        next
        
        for i = 0 to m_Nk - 1
            m_fkey(i) = CipherKey(i)
        next
        
        j = m_Nk
        k = 0
        do while j < N
            m_fkey(j) = m_fkey(j - m_Nk) xor _
                subByte(RotateLeft(m_fkey(j - 1), 24)) xor m_rco(k)
            if m_Nk <= 6 then
                i = 1
                do while i < m_Nk and (i + j) < N
                    m_fkey(i + j) = m_fkey(i + j - m_Nk) xor _
                        m_fkey(i + j - 1)
                    i = i + 1
                loop
            else
                i = 1
                do while i < 4 and (i + j) < N
                    m_fkey(i + j) = m_fkey(i + j - m_Nk) xor _
                        m_fkey(i + j - 1)
                    i = i + 1
                loop
                if j + 4 < N then
                    m_fkey(j + 4) = m_fkey(j + 4 - m_Nk) xor _
                        subByte(m_fkey(j + 3))
                end if
                i = 5
                do while i < m_Nk and (i + j) < N
                    m_fkey(i + j) = m_fkey(i + j - m_Nk) xor _
                        m_fkey(i + j - 1)
                    i = i + 1
                loop
            end if
            
            j = j + m_Nk
            k = k + 1
        loop
        
        for j = 0 to m_Nb - 1
            m_rkey(j + N - nb) = m_fkey(j)
        next
        
        i = m_Nb
        do while i < N - m_Nb
            k = N - m_Nb - i
            for j = 0 to m_Nb - 1
                m_rkey(k + j) = InvMixCol(m_fkey(i + j))
            next
            i = i + m_Nb
        loop
        
        j = N - m_Nb
        do while j < N
            m_rkey(j - N + m_Nb) = m_fkey(j)
            j = j + 1
        loop
    end sub
    
    private sub encrypt(buff())
        dim i
        dim j
        dim k
        dim m
        dim a(7)
        dim b(7)
        dim x
        dim y
        dim t
        
        for i = 0 to m_Nb - 1
            j = i * 4
            
            a(i) = PackFrom(buff, j)
            a(i) = a(i) xor m_fkey(i)
        next
        
        k = m_Nb
        x = a
        y = b
        
        for i = 1 to m_Nr - 1
            for j = 0 to m_Nb - 1
                m = j * 3
                y(j) = m_fkey(k) xor m_ftable(x(j) and m_lOnBits(7)) xor _
                    RotateLeft(m_ftable(RShift(x(m_fi(m)), 8) and m_lOnBits(7)), 8) xor _
                    RotateLeft(m_ftable(RShift(x(m_fi(m + 1)), 16) and m_lOnBits(7)), 16) xor _
                    RotateLeft(m_ftable(RShift(x(m_fi(m + 2)), 24) and m_lOnBits(7)), 24)
                k = k + 1
            next
            t = x
            x = y
            y = t
        next
        
        for j = 0 to m_Nb - 1
            m = j * 3
            y(j) = m_fkey(k) xor m_fbsub(x(j) and m_lOnBits(7)) xor _
                RotateLeft(m_fbsub(RShift(x(m_fi(m)), 8) and m_lOnBits(7)), 8) xor _
                RotateLeft(m_fbsub(RShift(x(m_fi(m + 1)), 16) and m_lOnBits(7)), 16) xor _
                RotateLeft(m_fbsub(RShift(x(m_fi(m + 2)), 24) and m_lOnBits(7)), 24)
            k = k + 1
        next
        
        for i = 0 to m_Nb - 1
            j = i * 4
            UnpackFrom y(i), buff, j
            x(i) = 0
            y(i) = 0
        next
    end sub
    
    private sub decrypt(buff())
        dim i
        dim j
        dim k
        dim m
        dim a(7)
        dim b(7)
        dim x
        dim y
        dim t
        
        for i = 0 to m_Nb - 1
            j = i * 4
            a(i) = PackFrom(buff, j)
            a(i) = a(i) xor m_rkey(i)
        next
        
        k = m_Nb
        x = a
        y = b
        
        for i = 1 to m_Nr - 1
            for j = 0 to m_Nb - 1
                m = j * 3
                y(j) = m_rkey(k) xor m_rtable(x(j) and m_lOnBits(7)) xor _
                    RotateLeft(m_rtable(RShift(x(m_ri(m)), 8) and m_lOnBits(7)), 8) xor _
                    RotateLeft(m_rtable(RShift(x(m_ri(m + 1)), 16) and m_lOnBits(7)), 16) xor _
                    RotateLeft(m_rtable(RShift(x(m_ri(m + 2)), 24) and m_lOnBits(7)), 24)
                k = k + 1
            next
            t = x
            x = y
            y = t
        next
        
        for j = 0 to m_Nb - 1
            m = j * 3
            
            y(j) = m_rkey(k) xor m_rbsub(x(j) and m_lOnBits(7)) xor _
                RotateLeft(m_rbsub(RShift(x(m_ri(m)), 8) and m_lOnBits(7)), 8) xor _
                RotateLeft(m_rbsub(RShift(x(m_ri(m + 1)), 16) and m_lOnBits(7)), 16) xor _
                RotateLeft(m_rbsub(RShift(x(m_ri(m + 2)), 24) and m_lOnBits(7)), 24)
            k = k + 1
        next
        
        for i = 0 to m_Nb - 1
            j = i * 4
            
            UnpackFrom y(i), buff, j
            x(i) = 0
            y(i) = 0
        next
    end sub
    
    private function IsInitialized(vArray)
        On Error Resume next
        
        IsInitialized = IsNumeric(UBound(vArray))
    end function
    
    private sub CopyBytesASP(bytDest, lDestStart, bytSource(), lSourceStart, lLength)
        dim lCount
        
        lCount = 0
        do
            bytDest(lDestStart + lCount) = bytSource(lSourceStart + lCount)
            lCount = lCount + 1
        loop until lCount = lLength
    end sub
    
    ' Function: encryptData
    ' 
    ' Use this method to encrypt your data.
    ' 
    ' Parameters:
    ' 
    '   (byte[]) - Data to be encrypted
    '   (byte[]) - Password
    ' 
    ' Returns:
    ' 
    '   (byte[]) - Encrypted version of the input data
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim sMessage : sMessage = "This is a very secret message"
    ' dim sPassword : sPassword = "Password"
    ' 
    ' dim E, byteMessage, bytePassword, encData
    ' set E = new Rijndael
    ' byteMessage  = E.string2bytes(sMessage)
    ' bytePassword = E.string2bytes(sPassword)
    ' encData = E.encryptData( byteMessage, bytePassword )
    ' 
    ' Response.write( "Plain Text: " & sMessage & "<br />" )
    ' Response.write( "Encrypted Hex: " & E.bytes2hex( encData ) & "<br />")
    ' set E = nothing
    ' 
    ' (end code)
    ' 
    public function encryptData(bytMessage, bytPassword)
        dim bytKey(31)
        dim bytIn()
        dim bytOut()
        dim bytTemp(31)
        dim lCount
        dim lLength
        dim lEncodedLength
        dim bytLen(3)
        dim lPosition
        
        if not IsInitialized(bytMessage) then
            exit function
        end if
        if not IsInitialized(bytPassword) then
            exit function
        end if
        
        for lCount = 0 to UBound(bytPassword)
            bytKey(lCount) = bytPassword(lCount)
            if lCount = 31 then
                exit For
            end if
        next
        
        gentables
        gkey 8, 8, bytKey
        
        lLength = UBound(bytMessage) + 1
        lEncodedLength = lLength + 4
        
        if lEncodedLength mod 32 <> 0 then
            lEncodedLength = lEncodedLength + 32 - (lEncodedLength mod 32)
        end if
        redim bytIn(lEncodedLength - 1)
        redim bytOut(lEncodedLength - 1)
        
        Unpack lLength, bytIn
        CopyBytesASP bytIn, 4, bytMessage, 0, lLength
    
        for lCount = 0 to lEncodedLength - 1 step 32
            CopyBytesASP bytTemp, 0, bytIn, lCount, 32
            Encrypt bytTemp
            CopyBytesASP bytOut, lCount, bytTemp, 0, 32
        next
        
        encryptData = bytOut
    end function
    
    ' Function: decryptData
    ' 
    ' Use this method to decrypt your data.
    ' 
    ' Parameters:
    ' 
    '   (byte[]) - Data to be decrypted
    '   (byte[]) - Password
    ' 
    ' Returns:
    ' 
    '   (byte[]) - Encrypted version of the input data
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim sMessage : sMessage = "This is a very secret message"
    ' dim sPassword : sPassword = "Password"
    ' 
    ' dim E, byteMessage, bytePassword, encData
    ' set E = new Rijndael
    ' byteMessage  = E.string2bytes(sMessage)
    ' bytePassword = E.string2bytes(sPassword)
    ' encData = E.encryptData( byteMessage, bytePassword )
    ' 
    ' Response.write( "Plain Text: " & sMessage & "<br />" )
    ' Response.write( "Encrypted Hex: " & E.bytes2hex( encData ) & "<br />")
    ' Response.write( "Decrypted: " &  E.bytes2string( E.decryptData( encData , bytePassword ) ) )
    ' set E = nothing
    ' 
    ' (end code)
    ' 
    public function decryptData(bytIn, bytPassword)
        dim bytMessage()
        dim bytKey(31)
        dim bytOut()
        dim bytTemp(31)
        dim lCount
        dim lLength
        dim lEncodedLength
        dim bytLen(3)
        dim lPosition
        
        if not IsInitialized(bytIn) then
            exit function
        end if
        if not IsInitialized(bytPassword) then
            exit function
        end if
        
        lEncodedLength = UBound(bytIn) + 1
        
        if lEncodedLength mod 32 <> 0 then
            exit function
        end if
        
        for lCount = 0 to UBound(bytPassword)
            bytKey(lCount) = bytPassword(lCount)
            if lCount = 31 then
                exit For
            end if
        next
        
        gentables
        gkey 8, 8, bytKey
        
        redim bytOut(lEncodedLength - 1)
        
        for lCount = 0 to lEncodedLength - 1 step 32
            CopyBytesASP bytTemp, 0, bytIn, lCount, 32
            Decrypt bytTemp
            CopyBytesASP bytOut, lCount, bytTemp, 0, 32
        next
        
        lLength = Pack(bytOut)
        
        if lLength > lEncodedLength - 4 then
            exit function
        end if
        
        redim bytMessage(lLength - 1)
        CopyBytesASP bytMessage, 0, bytOut, 4, lLength
        
        decryptData = bytMessage
    end function
    
    ' Function: string2bytes
    ' 
    ' Convert strings into an array of bytes.
    ' 
    ' Parameters:
    ' 
    '   (string) - Word or phrase to be converted
    ' 
    ' Returns:
    ' 
    '   (byte[]) - Converted data
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim E, a, i : set E = new Rijndael
    ' a = E.string2bytes("Lorem ipsum")
    ' Response.write "[ "
    ' for i = 0 to ubound(a)
    '     Response.write( a(i) & " " )
    ' next
    ' Response.write "]"
    ' erase a
    ' set E = nothing
    ' 
    ' (end code)
    ' 
    public function string2bytes(string)
        dim i, bytes, length
        length = len(string)
        redim bytes(length - 1)
        for i = 1 to length
            bytes(i - 1) = cbyte( ascb( mid( string, i, 1) ) )
        next
        string2bytes = bytes
    end function
    
    ' Function: bytes2string
    ' 
    ' Convert an array of bytes into strings.
    ' 
    ' Parameters:
    ' 
    '   (byte[]) - Byte array to be converted
    ' 
    ' Returns:
    ' 
    '   (string) - Converted data
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim E : set E = new Rijndael
    ' Response.write( E.bytes2string( E.string2bytes("Lorem ipsum") ) )
    ' set E = nothing
    ' 
    ' (end code)
    ' 
    public function bytes2string(bytes)
        dim i, string, length
        string = ""
        length = ubound(bytes) + 1
        for i = 0 to length - 1
            string = string & chr( bytes(i) )
        next
        bytes2string = string
    end function
    
    ' Function: bytes2hex
    ' 
    ' Convert an array of bytes into it's hex form.
    ' 
    ' Parameters:
    ' 
    '   (byte[]) - Byte array to be converted
    ' 
    ' Returns:
    ' 
    '   (hex) - Converted data
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim E : set E = new Rijndael
    ' Response.write( E.bytes2hex( E.string2bytes("Lorem ipsum") ) )
    ' set E = nothing
    ' 
    ' (end code)
    ' 
    public function bytes2hex(bytes)
        dim i, string, length
        string = ""
        length = ubound(bytes)
        for i = 0 To length
            string = string & right("0" & hex( bytes(i) ), 2)
        next
        bytes2hex = string
    end function
    
    ' Function: hex2bytes(hex)
    ' 
    ' Convert the hex form into an array of bytes
    ' 
    ' Parameters:
    ' 
    '     (hex) - Data to be converted
    ' 
    ' Returns:
    ' 
    '     (byte[]) - Byte array
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim E, a, i : set E = new Rijndael
    ' a = E.hex2bytes("FF00")
    ' Response.write "[ "
    ' for i = 0 to ubound(a)
    '     Response.write( a(i) & " " )
    ' next
    ' Response.write "]"
    ' erase a
    ' set E = nothing
    ' 
    ' (end code)
    ' 
    public function hex2bytes(hex)
        dim i, upper, bytes
        upper = ( len(hex) / 2 ) - 1
        redim bytes( upper )
        for i = 0 to upper
            bytes(i) = cbyte( "&H" & mid( hex, (i * 2 + 1), 2 ) )
        next
        hex2bytes = bytes
    end function
    
end class

%>
