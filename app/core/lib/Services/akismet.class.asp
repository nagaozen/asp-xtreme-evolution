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

' Class: Akismet
' 
' This class allows any ASP 3.0 application to use the Akismet anti-comment spam
' service. This service performs a number of checks on submitted data and
' returns whether or not the data is likely to be spam. Please note that in
' order to use this class, you must have a valid WordPress API key. They are
' free for non/small-profit types and getting one will only take couple of
' minutes. For commercial use, please visit the Akismet commercial licensing
' page.
' 
' About:
' 
'   - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ March 2008
' 
class Akismet
    
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
    
    private akismetServer
    private akismetVersion
    
    ' Property: wpApiKey
    ' 
    ' WordPress API key.
    ' 
    ' Contains:
    ' 
    '   (hex) - key
    ' 
    public wpApiKey
    
    ' Property: blog
    ' 
    ' Blog address.
    ' 
    ' Contains:
    ' 
    '   (uri) - blog address
    ' 
    public blog
    
    ' Property: comment
    ' 
    ' User comment data.
    ' 
    ' Contains:
    ' 
    '   (String Dictionary) - comment data
    ' 
    public comment
    
    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "1.1.0"
        
        akismetServer   = "rest.akismet.com"
        akismetVersion  = "1.1"
        
        set comment = Server.createObject("Scripting.Dictionary")
        comment.add "user_agent", Request.ServerVariables("HTTP_USER_AGENT")
        comment.add "referrer", Request.ServerVariables("HTTP_REFERER")
        comment.add "user_ip", Request.ServerVariables("REMOTE_ADDR")
        
        dim sv
        for each sv in Request.ServerVariables
            comment.add lcase(sv), Request.ServerVariables(sv)
        next
    end sub
    
    private sub Class_terminate()
        set comment = nothing
    end sub
    
    private function httpPost(url, data)
        dim Xhr : set Xhr = Server.createObject("MSXML2.ServerXMLHTTP.6.0")
        Xhr.open "POST", url, false
        Xhr.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        Xhr.setRequestHeader "Content-Length", len(data)
        Xhr.setRequestHeader "User-Agent", "Akismet ASP Class/" & classVersion & " | Akismet/1.11"
        Xhr.send(data)
        httpPost = Xhr.responseText
        set Xhr = nothing
    end function
    
    private function sd2querystring(sd)
        dim querystring, key
        querystring = ""
        for each key in sd.keys
            querystring = querystring & key & "=" & Server.urlEncode(sd.item(key)) & "&"
        next
        sd2querystring = mid(querystring, 1, ( len(querystring) - 1 ))
    end function
    
    ' Subroutine: initialize
    ' 
    ' ASP Classes doesn't have a constructor, so you must
    ' initialize the class manually.
    ' 
    ' Parameters:
    ' 
    '   (hex) - Wordpress API key
    '   (uri) - Blog address
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim SpamSentinel : set SpamSentinel = new Akismet
    ' SpamSentinel.initialize "123456789abc", "http://blog.domain.com"
    ' set SpamSentinel = nothing
    ' 
    ' (end code)
    ' 
    public sub initialize(key, uri)
        wpApiKey = key
        blog = uri
        comment.add "blog", blog
        
        dim akismet_response : akismet_response = httpPost("http://" & akismetServer & "/" & akismetVersion & "/verify-key", "key=" & wpApiKey & "&blog=" & blog)
        if( akismet_response = "invalid" ) then
            Response.write("Invalid API key. Please obtain one from http://wordpress.com/api-keys/")
            Response.end
        end if
    end sub
    
    ' Subroutine: setPermalink
    ' 
    ' Set Akismet permalink parameter.
    ' 
    ' Parameters:
    ' 
    '   (string) - value
    ' 
    public sub setPermalink(value)
        comment.add "permalink", value
    end sub
    
    ' Subroutine: setCommentType
    ' 
    ' Set Akismet comment_type parameter.
    ' 
    ' Parameters:
    ' 
    '   (string) - value
    ' 
    public sub setCommentType(value)
        comment.add "comment_type", value
    end sub
    
    ' Subroutine: setCommentAuthor
    ' 
    ' Set Akismet comment_author parameter.
    ' 
    ' Parameters:
    ' 
    '   (string) - value
    ' 
    public sub setCommentAuthor(value)
        comment.add "comment_author", value
    end sub
    
    ' Subroutine: setCommentAuthorEmail
    ' 
    ' Set Akismet comment_author_email parameter.
    ' 
    ' Parameters:
    ' 
    '   (string) - value
    ' 
    public sub setCommentAuthorEmail(value)
        comment.add "comment_author_email", value
    end sub
    
    ' Subroutine: setCommentAuthorUrl
    ' 
    ' Set Akismet comment_author_url parameter.
    ' 
    ' Parameters:
    ' 
    '   (string) - value
    ' 
    public sub setCommentAuthorUrl(value)
        comment.add "comment_author_url", value
    end sub
    
    ' Subroutine: setCommentContent
    ' 
    ' Set Akismet comment_content parameter.
    ' 
    ' Parameters:
    ' 
    '   (string) - value
    ' 
    public sub setCommentContent(value)
        comment.add "comment_content", value
    end sub
    
    ' Function: isSpam
    ' 
    ' Test for spam.
    ' 
    ' Returns:
    ' 
    '   true - if Akismet thinks this post is a spam
    '   false - otherwise
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim SpamSentinel : set SpamSentinel = new Akismet
    ' SpamSentinel.initialize "123456789abc", "http://blog.domain.com"
    ' SpamSentinel.setPermalink "http://blog.domain.com/entry-permalink/"
    ' SpamSentinel.setCommentType "comment"
    ' SpamSentinel.setCommentAuthor "Author"
    ' SpamSentinel.setCommentAuthorEmail "author@isp.com"
    ' SpamSentinel.setCommentAuthorUrl "http://author.isp.com"
    ' SpamSentinel.setCommentContent "The content that was submit"
    ' Response.write SpamSentinel.isSpam()
    ' set SpamSentinel = nothing
    ' 
    ' (end code)
    ' 
    public function isSpam()
        isSpam = CBool( httpPost("http://" & wpApiKey & "." & akismetServer & "/" & akismetVersion & "/comment-check", sd2querystring(comment)) )
    end function
    
    ' Function: submitSpam
    ' 
    ' This call is for submitting comments that weren't marked as spam but should have been.
    ' 
    ' Returns:
    ' 
    '   (string) - Akismet Service message. Something like: "Thanks for making the web a better place."
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim SpamSentinel : set SpamSentinel = new Akismet
    ' SpamSentinel.initialize "123456789abc", "http://blog.domain.com"
    ' SpamSentinel.setPermalink "http://blog.domain.com/entry-permalink/"
    ' SpamSentinel.setCommentType "comment"
    ' SpamSentinel.setCommentAuthor "Author"
    ' SpamSentinel.setCommentAuthorEmail "author@isp.com"
    ' SpamSentinel.setCommentAuthorUrl "http://author.isp.com"
    ' SpamSentinel.setCommentContent "The content that was submit"
    ' Response.write SpamSentinel.submitSpam()
    ' set SpamSentinel = nothing
    ' 
    ' (end code)
    ' 
    public function submitSpam()
        submitSpam = httpPost("http://" & wpApiKey & "." & akismetServer & "/" & akismetVersion & "/submit-spam", sd2querystring(comment))
    end function
    
    ' Function: submitHam
    ' 
    ' This call is intended for the marking of false positives, things that were incorrectly marked as spam.
    ' 
    ' Returns:
    ' 
    '   (string) - Akismet Service message. Something like: "Thanks for making the web a better place."
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim SpamSentinel : set SpamSentinel = new Akismet
    ' SpamSentinel.initialize "123456789abc", "http://blog.domain.com"
    ' SpamSentinel.setPermalink "http://blog.domain.com/entry-permalink/"
    ' SpamSentinel.setCommentType "comment"
    ' SpamSentinel.setCommentAuthor "Author"
    ' SpamSentinel.setCommentAuthorEmail "author@isp.com"
    ' SpamSentinel.setCommentAuthorUrl "http://author.isp.com"
    ' SpamSentinel.setCommentContent "The content that was submit"
    ' Response.write SpamSentinel.submitHam()
    ' set SpamSentinel = nothing
    ' 
    ' (end code)
    ' 
    public function submitHam()
        submitHam = httpPost("http://" & wpApiKey & "." & akismetServer & "/" & akismetVersion & "/submit-ham", sd2querystring(comment))
    end function
    
end class

%>
