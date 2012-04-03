<%

' File: paginator.asp
' 
' AXE(ASP Xtreme Evolution) paginator utility.
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



' Class: Paginator
' 
' Paginator is a very useful class to create general purpose page indexes.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Feb 2009
' 
class Paginator
    
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
    
    ' Property: all
    ' 
    ' The text to be used as "View all" in the index.
    ' 
    ' Contains:
    ' 
    '   (string) - Defaults to "View all"
    ' 
    public all
    
    ' Property: prv
    ' 
    ' The text to be used as "Previous" in the index.
    ' 
    ' Contains:
    ' 
    '   (string) - Defaults to "Previous"
    ' 
    public prv
    
    ' Property: nxt
    ' 
    ' The text to be used as "Next" in the index.
    ' 
    ' Contains:
    ' 
    '   (string) - Defaults to "Next"
    ' 
    public nxt
    
    ' Property: visibles
    ' 
    ' The number of indexes visible except First, Selected and Last.
    ' 
    ' Contains:
    ' 
    '   (int) - Defaults to 10
    ' 
    public visibles
    
    ' Property: page
    ' 
    ' The index of the current page.
    ' 
    ' Contains:
    ' 
    '   (int) - Defaults to 1
    ' 
    public page
    
    ' Property: pages
    ' 
    ' The number of pages.
    ' 
    ' Contains:
    ' 
    '   (int) - Defaults to 1
    ' 
    public pages
    
    ' Property: url
    ' 
    ' The url template to be used in the index anchors.
    ' 
    ' Contains:
    ' 
    '   (string) - url template
    ' 
    ' Note:
    ' 
    '   - The {page} placeholder will be replace with the page index.
    ' 
    public url
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        all = "View all"
        prv = "Previous"
        nxt = "Next"
        
        visibles = 10
        page     = 1
        pages    = 1
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: make
    ' 
    ' Builds the index
    ' 
    ' Returns:
    ' 
    '     (string) - indexes
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Indexes : set Indexes = new Paginator
    ' Indexes.page = 1
    ' Indexes.pages = 10
    ' Response.write( Indexes.make() )
    ' set Indexes = nothing
    ' 
    ' (end code)
    ' 
    public function make()
        if( isEmpty(url) ) then
            Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Please provide the Paginator 'url' property."
        end if
        
        dim a, b, half
        a = 1 : b = 1
        half = floor(visibles/2)
        
        if(pages > 1) then
            if(pages <= visibles) then
                b = pages
            elseif(page + half > pages) then
                a = pages - visibles : b = pages
            elseif(( page + half <= pages ) and ( page - half > 1 )) then
                a = page - half : b = page + half
            else
                a = 1 : b = visibles
            end if
        end if
        
        dim Stream : set Stream = Server.createObject("ADODB.Stream")
        Stream.type = adTypeText
        Stream.mode = adModeReadWrite
        Stream.open()
        
        Stream.writeText("<ul class='paginator'>" & vbNewline)
        
        if(page = 0) then
            Stream.writeText( strsubstitute("<li class='paginator-view-all'><span>{0}</span></li>" & vbNewline, array(all)) )
        else
            Stream.writeText( strsubstitute("<li class='paginator-view-all'><a href='{0}'><span>{1}</span></a></li>" & vbNewline, array(replace(url,"{page}","0"), all)) )
        end if
        
        if(page > 1) then
            Stream.writeText( strsubstitute("<li class='paginator-previous'><a href='{0}'><span>{1}</span></a></li>" & vbNewline, array(replace(url, "{page}", clng(page - 1)), prv)) )
        end if
        
        if(page - half > 1) then
            Stream.writeText( strsubstitute("<li class='paginator-first'><a href='{0}'><span>{1}</span></a></li>" & vbNewline, array(replace(url, "{page}", clng(1)), 1)) )
            Stream.writeText("<li class='paginator-gap'><span>...</span></li>" & vbNewline)
        end if
        
        dim i : for i = a to b
            if(i = page) then
                Stream.writeText( strsubstitute("<li class='paginator-index selected'><span>{0}</span></li>" & vbNewline, array(i)) )
            else
                Stream.writeText( strsubstitute("<li class='paginator-index'><a href='{0}'><span>{1}</span></a></li>" & vbNewline, array(replace(url, "{page}", clng(i)), i)) )
            end if
        next
        
        if(page + half < pages) then
            Stream.writeText("<li class='paginator-gap'><span>...</span></li>" & vbNewline)
            Stream.writeText( strsubstitute("<li class='paginator-last'><a href='{0}'><span>{1}</span></a></li>" & vbNewline, array(replace(url, "{page}", clng(pages)), pages)) )
        end if
        
        if( ( 0 < page ) and ( page < pages ) ) then
            Stream.writeText( strsubstitute("<li class='paginator-next'><a href='{0}'><span>{1}</span></a></li>" & vbNewline, array(replace(url, "{page}", clng(page + 1)), nxt)) )
        end if
        
        Stream.writeText("</ul>" & vbNewline)
        
        Stream.position = 0
        make = Stream.readText()
        
        Stream.close()
        set Stream = nothing
    end function
    
end class

%>
