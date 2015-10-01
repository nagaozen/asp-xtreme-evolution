<%

' File: color.asp
' 
' AXE(ASP Xtreme Evolution) color utility.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2007-2012 Fabio Zendhi Nagao
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



' Class: Color
' 
' This class makes color manipulation a very fun task. It can mix, invert and
' convert colors between Hexadecimal, RGB and HSB.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao  @ November 2008
' 
class Color
    
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
    '   (string) - version
    ' 
    public classVersion
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: invert
    ' 
    ' Inverts the color.
    ' 
    ' Parameters:
    ' 
    '     (int) - red
    '     (int) - green
    '     (int) - blue
    ' 
    ' Returns:
    ' 
    '     (int[]) - [red, green, blue]
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Converter : set Converter = new Color
    ' Response.write join(Converter.invert(51, 102, 153), ",") ' prints 204,153,102
    ' set Converter = nothing
    ' 
    ' (end code)
    ' 
    public function invert(red, green, blue)
        dim a : a = array(red, green, blue)
        dim i
        for i = 0 to 2
            a(i) = 255 - a(i)
        next
        invert = a
    end function
    
    ' Function: mix
    ' 
    ' Mix two or more colors.
    ' 
    ' Parameters:
    ' 
    '     (int[])   - base RGB color
    '     (int[][]) - array  with colors in RGB to be mixed
    '     (int)     - number between [0,100] which is the amount of the colors to be mixed.
    ' 
    ' Returns:
    ' 
    '     (int[]) - Mixed color
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Converter : set Converter = new Color
    ' Response.write join(Converter.mix(array(0,0,0), array(Color.hex2rgb("FFF"), Color.hex2rgb("F0F")), 10), ",") ' prints 49,23,49
    ' set Converter = nothing
    ' 
    ' (end code)
    ' 
    public function mix(base, colors, percentage)
        dim i, u
        for i = 0 to ubound(colors)
            for u = 0 to 2
                base(u) = round( (base(u) / 100 * (100 - percentage)) + (colors(i)(u) / 100 * percentage) )
            next
        next
        mix = base
    end function
    
    ' Function: rgb2hex
    ' 
    ' Converts RGB to HEX.
    ' 
    ' Parameters:
    ' 
    '     (int) - red
    '     (int) - green
    '     (int) - blue
    ' 
    ' Returns:
    ' 
    '     (string) - Hexadecimal representation with 6 digits
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Converter : set Converter = new Color
    ' Response.write Converter.rgb2hex(51, 102, 153) ' prints 336699
    ' set Converter = nothing
    ' 
    ' (end code)
    ' 
    public function rgb2hex(red, green, blue)
        rgb2hex = join(array( right("00" & dec2hex(red), 2), right("00" & dec2hex(green), 2), right("00" & dec2hex(blue), 2) ),"")
    end function
    
    ' Function: hex2rgb
    ' 
    ' Converts HEX to RGB
    ' 
    ' Parameters:
    ' 
    '     (string) - Hexadecimal representation (both FFFFFF and FFF)
    ' 
    ' Returns:
    ' 
    '     (int[]) - [red, green, blue]
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Converter : set Converter = new Color
    ' Response.write join(Converter.hex2rgb("336699"), ",") ' prints 51,102,153
    ' set Converter = nothing
    ' 
    ' (end code)
    ' 
    public function hex2rgb(sharp)
        if(len(sharp) = 3) then
            sharp = mid(sharp, 1, 1) & mid(sharp, 1, 1) & mid(sharp, 2, 1) & mid(sharp, 2, 1) & mid(sharp, 3, 1) & mid(sharp, 3, 1)
        end if
        dim red : red = mid(sharp, 1, 2)
        dim green : green = mid(sharp, 3, 2)
        dim blue : blue = mid(sharp, 5, 2)
        hex2rgb = array(hex2dec(red), hex2dec(green), hex2dec(blue))
    end function
    
    ' Function: rgb2hsb
    ' 
    ' Converts RGB to HSB.
    ' 
    ' Parameters:
    ' 
    '     (int) - red
    '     (int) - green
    '     (int) - blue
    ' 
    ' Returns:
    ' 
    '     (int[]) - [hue, saturation, brightness]
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Converter : set Converter = new Color
    ' Response.write Converter.rgb2hsb(51, 102, 153) ' prints 210,67,60
    ' set Converter = nothing
    ' 
    ' (end code)
    ' 
    public function rgb2hsb(red, green, blue)
        dim hue, saturation, brightness
        
        dim maxi : maxi = max(array(red, green, blue))
        dim mini : mini = min(array(red, green, blue))
        dim delta : delta = maxi - mini
        
        brightness = (maxi / 255)
        saturation = iif( (maxi <> 0), (delta / maxi), 0)
        if( cdbl(saturation) = cdbl(0) ) then
            hue = 0
        else
            dim rr : rr = (maxi - red) / delta
            dim gr : gr = (maxi - green) / delta
            dim br : br = (maxi - blue) / delta
            if( red = maxi ) then
                hue = br - gr
            elseif( green = maxi ) then
                hue = 2 + rr - br
            else
                hue = 4 + gr - rr
            end if
            hue = hue / 6
            if( hue < 0 ) then
                hue = hue + 1
            end if
        end if
        
        rgb2hsb = array(round(hue * 360), round(saturation * 100), round(brightness * 100))
    end function
    
    ' Function: hsb2rgb
    ' 
    ' Converts HSB to RGB
    ' 
    ' Parameters:
    ' 
    '     (int) - hue
    '     (int) - saturation
    '     (int) - brightness
    ' 
    ' Returns:
    ' 
    '     (int[]) - [red, green, blue]
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Converter : set Converter = new Color
    ' Response.write Converter.hsb2rgb(210,67,60) ' prints 50,102,153
    ' set Converter = nothing
    ' 
    ' (end code)
    ' 
    public function hsb2rgb(hue, saturation, brightness)
        dim br : br = round( brightness / 100 * 255 )
        if( saturation = 0 ) then
            hsb2rgb = array(br, br, br)
            exit function
        else
            hue = hue mod 360
            dim f : f = hue mod 60
            dim p : p = round( ( brightness * (100 - saturation) ) / 10000 * 255 )
            dim q : q = round( ( brightness * (6000 - saturation * f) ) / 600000 * 255 )
            dim t : t = round( ( brightness * (6000 - saturation * (60 - f) ) ) / 600000 * 255 )
        end if
        select case floor(hue / 60)
            case 0
                hsb2rgb = array(br, t, p)
            case 1
                hsb2rgb = array(q, br, p)
            case 2
                hsb2rgb = array(p, br, t)
            case 3
                hsb2rgb = array(p, q, br)
            case 4
                hsb2rgb = array(t, p, br)
            case 5
                hsb2rgb = array(br, p, q)
        end select
    end function
    
    ' Function: hex2hsb
    ' 
    ' Converts HEX to HSB
    ' 
    ' Parameters:
    ' 
    '     (string) - Hexadecimal representation (both FFFFFF and FFF)
    ' 
    ' Returns:
    ' 
    '     (int[]) - [hue, saturation, brightness]
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Converter : set Converter = new Color
    ' Response.write Converter.hex2hsb("336699") ' prints 210,67,60
    ' set Converter = nothing
    ' 
    ' (end code)
    ' 
    public function hex2hsb(sharp)
        dim a : a = hex2rgb(sharp)
        hex2hsb = rgb2hsb(a(0), a(1), a(2))
    end function
    
    ' Function: hsb2hex
    ' 
    ' Converts HSB to HEX
    ' 
    ' Parameters:
    ' 
    '     (int) - hue
    '     (int) - saturation
    '     (int) - brightness
    ' 
    ' Returns:
    ' 
    '     (string) - Hexadecimal representation with 6 digits
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Converter : set Converter = new Color
    ' Response.write Converter.hsb2hex(210,67,60) ' prints 336699
    ' set Converter = nothing
    ' 
    ' (end code)
    ' 
    public function hsb2hex(hue, saturation, brightness)
        dim a : a = hsb2rgb(hue, saturation, brightness)
        hsb2hex = rgb2hex(a(0), a(1), a(2))
    end function
    
end class

%>
