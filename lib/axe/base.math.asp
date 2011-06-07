<%

' File: base.math.asp
' 
' More basic math functions derived from the built-in ones.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2007-2011 Fabio Zendhi Nagao
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
' 
' About:
' 
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ June 2009
' 

' Function: sec
' 
' Secant.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function sec(x)
    sec = 1 / cos(x)
end function

' Function: cosec
' 
' Cosecant.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function cosec(x)
    cosec = 1 / sin(x)
end function

' Function: cotan
' 
' Cotangent.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function cotan(x)
    cotan = 1 / tan(x)
end function

' Function: arcsin
' 
' Inverse Sine.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function arcsin(x)
    arcsin = atn(x / sqr(-x * x + 1))
end function

' Function: arccos
' 
' Inverse Cosine.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function arccos(x)
    arccos = atn(-x / sqr(-x * x + 1)) + 2 * atn(1)
end function

' Function: arcsec
' 
' Inverse Secant.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function arcsec(x)
    arcsec = atn(x / sqr(x * x - 1)) + sgn((x) -1) * (2 * atn(1))
end function

' Function: arccosec
' 
' Inverse Cosecant.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function arccosec(x)
    arccosec = atn(x / sqr(x * x - 1)) + (sgn(x) - 1) * (2 * atn(1))
end function

' Function: arccotan
' 
' Inverse Cotangent.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function arccotan(x)
    arccotan = atn(x) + 2 * atn(1)
end function

' Function: hsin
' 
' Hyperbolic Sine.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function hsin(x)
    hsin = (exp(x) - exp(-x)) / 2
end function

' Function: hcos
' 
' Hyperbolic Cosine.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function hcos(x)
    hcos = (exp(x) + exp(-x)) / 2
end function

' Function: htan
' 
' Hyperbolic Tangent.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function htan(x)
    htan = (exp(x) - exp(-x)) / (exp(x) + exp(-x))
end function

' Function: hsec
' 
' Hyperbolic Secant.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function hsec(x)
    hsec = 2 / (exp(x) + exp(-x))
end function

' Function: hcosec
' 
' Hyperbolic Cosecant.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function hcosec(x)
    hcosec = 2 / (exp(x) - exp(-x))
end function

' Function: hcotan
' 
' Hyperbolic Cotangent.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function hcotan(x)
    hcotan = (exp(x) + exp(-x)) / (exp(x) - exp(-x))
end function

' Function: harcsin
' 
' Inverse Hyperbolic Sine.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function harcsin(x)
    harcsin = log(x + sqr(x * x + 1))
end function

' Function: harccos
' 
' Inverse Hyperbolic Cosine.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function harccos(x)
    harccos = log(x + sqr(x * x - 1))
end function

' Function: harctan
' 
' Inverse Hyperbolic Tangent.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function harctan(x)
    harctan = log((1 + x) / (1 - x)) / 2
end function

' Function: harcsec
' 
' Inverse Hyperbolic Secant.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function harcsec(x)
    harcsec = log((sqr(-x * x + 1) + 1) / x)
end function

' Function: harccosec
' 
' Inverse Hyperbolic Cosecant.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function harccosec(x)
    harccosec = log((sgn(x) * sqr(x * x + 1) +1) / x)
end function

' Function: harccotan
' 
' Inverse Hyperbolic Cotangent.
' 
' Parameters:
' 
'     (double) - usual x
' 
public function harccotan(x)
    harccotan = log((x + 1) / (x - 1)) / 2
end function

' Function: logN
' 
' Logarithm to base N.
' 
' Parameters:
' 
'     (double) - usual x
'     (int)    - logarithm base
' 
public function logN(x, N)
    logN = log(x) / log(N)
end function

%>
