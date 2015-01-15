<%

' File: rss.asp
' 
' AXE(ASP Xtreme Evolution) implementation of RSS 2.0 web feeds.
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



' Class: RSS
' 
' Really Simple Syndication Channel Model Abstraction. Implementation based on w3.org
' definition at <http://validator.w3.org/feed/docs/rss2.html>.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class RSS
    
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
    
    ' --[ Requireds ]-----------------------------------------------------------
    
    ' Property: title
    ' 
    ' It's how people refer to your service. If you have an HTML website that
    ' contains the same information as your RSS file, the title of your channel
    ' should be the same as the title of your website.
    ' 
    ' Contains:
    ' 
    '     (string) - The name of the channel
    ' 
    public title
    
    ' Property: link
    ' 
    ' The URL to the HTML website corresponding to the channel.
    ' 
    ' Contains:
    ' 
    '     (string) - The URL to the HTML website corresponding to the channel
    ' 
    public link
    
    ' Property: description
    ' 
    ' Phrase or sentence describing the channel.
    ' 
    ' Contains:
    ' 
    '     (string) - Phrase or sentence describing the channel
    ' 
    public description
    
    ' --[ Optionals ]-----------------------------------------------------------
    
    ' Property: language
    ' 
    ' This allows aggregators to group all Italian language sites, for example,
    ' on a single page. A list of allowable values for this element, as provided
    ' by Netscape (<http://backend.userland.com/stories/storyReader$16>). You
    ' may also use values defined by the W3C (<http://www.w3.org/TR/REC-html40/struct/dirlang.html#langcodes>).
    ' 
    ' Contains:
    ' 
    '     [(string)] - The language the channel is written in
    ' 
    public language
    
    ' Property: copyright
    ' 
    ' Copyright notice for content in the channel.
    ' 
    ' Contains:
    ' 
    '     [(string)] - Copyright notice for content in the channel
    ' 
    public copyright
    
    ' Property: managingEditor
    ' 
    ' Email address for person responsible for editorial content.
    ' 
    ' Contains:
    ' 
    '     [(string)] - Email address for person responsible for editorial content
    ' 
    public managingEditor
    
    ' Property: webMaster
    ' 
    ' Email address for person responsible for technical issues relating to channel.
    ' 
    ' Contains:
    ' 
    '     [(string)] - Email address for person responsible for technical issues relating to channel
    ' 
    public webMaster
    
    ' Property: pubDate
    ' 
    ' All date-times in RSS conform to the Date and Time Specification of
    ' <http://asg.web.cmu.edu/rfc/rfc822.html>, with the exception that the year
    ' may be expressed with two characters or four characters (four preferred).
    ' 
    ' Contains:
    ' 
    '     [(string)] - The publication date for the content in the channel
    ' 
    public pubDate
    
    ' Property: lastBuildDate
    ' 
    ' The last time the content of the channel changed.
    ' 
    ' Contains:
    ' 
    '     [(string)] - The last time the content of the channel changed.
    ' 
    public lastBuildDate
    
    ' Property: Categories
    ' 
    ' Specify one or more categories that the channel belongs to.
    ' 
    ' Contains:
    ' 
    '     [(RSS_Category[])] - Specify one or more categories that the channel belongs to
    ' 
    public Categories
    
    ' Property: generator
    ' 
    ' A string indicating the program used to generate the channel.
    ' 
    ' Contains:
    ' 
    '     [(string)] - A string indicating the program used to generate the channel
    ' 
    public generator
    
    ' Property: docs
    ' 
    ' A URL that points to the documentation for the format used in the RSS file.
    ' 
    ' Contains:
    ' 
    '     [(string)] - A URL that points to the documentation for the format used in the RSS file
    ' 
    public docs
    
    ' Property: Cloud
    ' 
    ' It specifies a web service that supports the rssCloud interface which can
    ' be implemented in HTTP-POST, XML-RPC or SOAP 1.1.
    ' 
    ' Contains:
    ' 
    '     [(RSS_Cloud)] - Cloud tag
    ' 
    public Cloud
    
    ' Property: ttl
    ' 
    ' It's a number of minutes that indicates how long a channel can be cached
    ' before refreshing from the source.
    ' 
    ' Contains:
    ' 
    '     [(int)] - TTL stands for time to live
    ' 
    public ttl
    
    ' Property: Image
    ' 
    ' Specifies a GIF, JPEG or PNG image that can be displayed with the channel.
    ' 
    ' Contains:
    ' 
    '     [(RSS_Image)] - Channel image
    ' 
    public Image
    
    ' Property: TextInput
    ' 
    ' The purpose of the <textInput> element is something of a mystery. You can
    ' use it to specify a search engine box. Or to allow a reader to provide
    ' feedback. Most aggregators ignore it.
    ' 
    ' Contains:
    ' 
    '     [(string)] - Specifies a text input box that can be displayed with the channel
    ' 
    public TextInput
    
    ' Property: SkipHours
    ' 
    ' An XML element that contains up to 24 <hour> sub-elements whose value is a
    ' number between 0 and 23, representing a time in GMT, when aggregators, if
    ' they support the feature, may not read the channel on hours listed in the
    ' skipHours element.
    ' 
    ' Contains:
    ' 
    '     [(RSS_SkipHours)] - A hint for aggregators telling them which hours they can skip
    ' 
    public SkipHours
    
    ' Property: SkipDays
    ' 
    ' An XML element that contains up to seven <day> sub-elements whose value is
    ' Monday, Tuesday, Wednesday, Thursday, Friday, Saturday or Sunday.
    ' Aggregators may not read the channel during days listed in the skipDays
    ' element.
    ' 
    ' Contains:
    ' 
    '     [(RSS_SkipDays)] - A hint for aggregators telling them which days they can skip
    ' 
    public SkipDays
    
    ' Property: Items
    ' 
    ' Specify one or more items that belongs to the channel.
    ' 
    ' Contains:
    ' 
    '     (List) - Specify one or more items that belongs to the channel
    ' 
    private Items
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        set Items = new List
    end sub
    
    private sub Class_terminate()
        set Items = nothing
    end sub
    
    ' Function: toString
    ' 
    ' Returns the RSS representation to be used inside <channel>.
    ' 
    ' Parameters:
    ' 
    '     (int) - Initial indentation level
    ' 
    ' Returns:
    ' 
    '     (string) - Item representation.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Item_item : set Item_item = new RSS_Item
    ' Item_item.title = "A Really Simple Syndication item"
    ' Item_item.link = "http://www.domain.com/friendly-url-as-permalink"
    ' Item_item.description = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla cursus nisl vel purus tristique eget venenatis erat gravida. Praesent nisi arcu, placerat at adipiscing sit amet, auctor tempus enim."
    ' Item_item.author = "name@domain.com (Your Name)"
    ' Item_item.comments = "http://www.domain.com/friendly-url-as-permalink#comments"
    ' Item_item.pubDate = "Sat, 22 August 2009 00:00:00 GMT"
    ' Item_item.source = "http://www.domain.com/feed/?type=rss"
    ' 
    ' dim Item_guid : set Item_guid = new RSS_Guid
    ' Item_guid.isPermaLink = true
    ' Item_guid.innerText = "http://www.domain.com/friendly-url-as-permalink"
    ' set Item_item.guid = Item_guid
    ' set Item_guid = nothing
    ' 
    ' Response.write(Item_item.toString(0))
    ' set Item_item = nothing
    ' 
    ' (end code)
    ' 
    public function toString()
        dim i, optionals : optionals = ""
        if(not isEmpty(language)) then optionals = optionals & ("{0}{1}<language>" & language & "</language>")
        if(not isEmpty(copyright)) then optionals = optionals & ("{0}{1}<copyright>" & copyright & "</copyright>")
        if(not isEmpty(managingEditor)) then optionals = optionals & ("{0}{1}<managingEditor>" & managingEditor & "</managingEditor>")
        if(not isEmpty(webMaster)) then optionals = optionals & ("{0}{1}<webMaster>" & webMaster & "</webMaster>")
        if(not isEmpty(pubDate)) then optionals = optionals & ("{0}{1}<pubDate>" & pubDate & "</pubDate>")
        if(not isEmpty(lastBuildDate)) then optionals = optionals & ("{0}{1}<lastBuildDate>" & lastBuildDate & "</lastBuildDate>")
        if(not isEmpty(Categories)) then
            for i = 0 to ubound(Categories)
                optionals = optionals & ("{0}" & Categories(i).toString(2))
            next
        end if
        if(not isEmpty(generator)) then optionals = optionals & ("{0}{1}<generator>" & generator & "</generator>")
        if(not isEmpty(docs)) then optionals = optionals & ("{0}{1}<docs>" & docs & "</docs>")
        if(not isEmpty(Cloud)) then optionals = optionals & ("{0}" & Cloud.toString(2))
        if(not isEmpty(ttl)) then optionals = optionals & ("{0}{1}<ttl>" & ttl & "</ttl>")
        if(not isEmpty(Image)) then optionals = optionals & ("{0}" & Image.toString(2))
        if(not isEmpty(TextInput)) then optionals = optionals & ("{0}" & TextInput.toString(2))
        if(not isEmpty(SkipHours)) then optionals = optionals & ("{0}" & SkipHours.toString(2))
        if(not isEmpty(SkipDays)) then optionals = optionals & ("{0}" & SkipDays.toString(2))
        if(not isEmpty(Items)) then
            for i = 0 to ubound(Items)
                optionals = optionals & ("{0}" & Items(i).toString(2))
            next
        end if
        optionals = strsubstitute(optionals, array(vbNewline, space(2 * 4)))
        
        toString = strsubstitute(join(array( _
            "<rss version='2.0'>", _
            "    <channel>", _
            "        <title>{0}</title>", _
            "        <link>{1}</link>", _
            "        <description>{2}</description>{3}", _
            "    </channel>", _
            "</rss>" _
        ), vbNewline), array(title, link, description))
    end function
    
end class



' Class: RSS_Item
' 
' Really Simple Syndication Item Model Abstraction. Implementation based on w3.org
' definition at <http://validator.w3.org/feed/docs/rss2.html#hrelementsOfLtitemgt>.
' 
' A channel may contain any number of <item>s. An item may represent a "story"
' -- much like a story in a newspaper or magazine; if so it's description is a
' synopsis of the story, and the link points to the full story. An item may also
' be complete in itself, if so, the description contains the text (entity-encoded
' HTML is allowed), and the link and title may be omitted. All elements of an item
' are optional, however at least one of title or description must be present.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class RSS_Item
    
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
    
    ' Property: title
    ' 
    ' The title of the item.
    ' 
    ' Contains:
    ' 
    '     [(string)] - The title of the item
    ' 
    public title
    
    ' Property: link
    ' 
    ' The URL of the item.
    ' 
    ' Contains:
    ' 
    '     [(string)] - The URL of the item
    ' 
    public link
    
    ' Property: description
    ' 
    ' The item synopsis.
    ' 
    ' Contains:
    ' 
    '     [(string)] - The item synopsis
    ' 
    public description
    
    ' Property: author
    ' 
    ' It's the email address of the author of the item. For newspapers and
    ' magazines syndicating via RSS, the author is the person who wrote the
    ' article that the <item> describes. For collaborative weblogs, the author
    ' of the item might be different from the managing editor or webmaster. For
    ' a weblog authored by a single individual it would make sense to omit the
    ' <author> element.
    ' 
    ' Contains:
    ' 
    '     [(string)] - Email address of the author of the item
    ' 
    public author
    
    ' Property: Categories
    ' 
    ' Includes the item in one or more categories. It has one optional attribute,
    ' domain, a string that identifies a categorization taxonomy. 
    ' 
    ' Contains:
    ' 
    '     [(RSS_Category[])] - Includes the item in one or more categories
    ' 
    public Categories
    
    ' Property: comments
    ' 
    ' URL of a page for comments relating to the item.
    ' 
    ' Contains:
    ' 
    '     [(string)] - URL of a page for comments relating to the item
    ' 
    public comments
    
    ' Property: Enclosure
    ' 
    ' Describes a media object that is attached to the item. It has three
    ' required attributes. url says where the enclosure is located, length says
    ' how big it is in bytes, and type says what it's type is, a standard MIME type.
    ' 
    ' Contains:
    ' 
    '     [(RSS_Enclosure)] - Describes a media object that is attached to the item
    ' 
    public Enclosure
    
    ' Property: Guid
    ' 
    ' It's a string that uniquely identifies the item. When present, an aggregator
    ' may choose to use this string to determine if an item is new. If the guid
    ' element has an attribute named "isPermaLink" with a value of true, the
    ' reader may assume that it is a permalink to the item, that is, a url that
    ' can be opened in a Web browser, that points to the full item described by 
    ' the <item> element.
    ' 
    ' Contains:
    ' 
    '     [(RSS_Guid)] - GUID stands for globally unique identifier
    ' 
    public Guid
    
    ' Property: pubDate
    ' 
    ' Indicates when the item was published.
    ' 
    ' Contains:
    ' 
    '     [(string)] - Indicates when the item was published
    ' 
    public pubDate
    
    ' Property: Source
    ' 
    ' It's value is the name of the RSS channel that the item came from, derived
    ' from it's <title>. It has one required attribute, url, which links to the 
    ' XMLization of the source.
    ' 
    ' Contains:
    ' 
    '     [(RSS_Source)] - The RSS channel that the item came from
    ' 
    public Source
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: toString
    ' 
    ' Returns the item representation to be used inside <channel>.
    ' 
    ' Parameters:
    ' 
    '     (int) - Initial indentation level
    ' 
    ' Returns:
    ' 
    '     (string) - Item representation.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Item_item : set Item_item = new RSS_Item
    ' Item_item.title = "A Really Simple Syndication item"
    ' Item_item.link = "http://www.domain.com/friendly-url-as-permalink"
    ' Item_item.description = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla cursus nisl vel purus tristique eget venenatis erat gravida. Praesent nisi arcu, placerat at adipiscing sit amet, auctor tempus enim."
    ' Item_item.author = "name@domain.com (Your Name)"
    ' Item_item.comments = "http://www.domain.com/friendly-url-as-permalink#comments"
    ' Item_item.pubDate = "Sat, 22 August 2009 00:00:00 GMT"
    ' Item_item.source = "http://www.domain.com/feed/?type=rss"
    ' 
    ' dim Item_guid : set Item_guid = new RSS_Guid
    ' Item_guid.isPermaLink = true
    ' Item_guid.innerText = "http://www.domain.com/friendly-url-as-permalink"
    ' set Item_item.guid = Item_guid
    ' set Item_guid = nothing
    ' 
    ' Response.write(Item_item.toString(0))
    ' set Item_item = nothing
    ' 
    ' (end code)
    ' 
    public function toString(indents)
        dim optionals : optionals = ""
        if(not isEmpty(title)) then optionals = optionals & ("{0}    <title>" & title & "</title>")
        if(not isEmpty(link)) then optionals = optionals & ("{0}    <link>" & link & "</link>")
        if(not isEmpty(description)) then optionals = optionals & ("{0}    <description>" & description & "</description>")
        if(not isEmpty(author)) then optionals = optionals & ("{0}    <author>" & author & "</author>")
        if(not isEmpty(Categories)) then
            dim i : for i = 0 to ubound(Categories)
                optionals = optionals & ("{0}" & Categories(i).toString(1))
            next
        end if
        if(not isEmpty(comments)) then optionals = optionals & ("{0}    <comments>" & comments & "</comments>")
        if(not isEmpty(Enclosure)) then optionals = optionals & ("{0}" & Enclosure.toString(1))
        if(not isEmpty(Guid)) then optionals = optionals & ("{0}" & Guid.toString(1))
        if(not isEmpty(pubDate)) then optionals = optionals & ("{0}    <pubDate>" & pubDate & "</pubDate>")
        if(not isEmpty(Source)) then optionals = optionals & ("{0}" & Source.toString(1))
        optionals = strsubstitute(optionals, array(vbNewline & space(indents * 4)))
        
        toString = strsubstitute(join(array( _
            "{0}<item>{1}", _
            "{0}</item>" _
        ), vbNewline), array(space(indents * 4), optionals))
    end function
    
end class



' Class: RSS_Category
' 
' Really Simple Syndication Category Model Abstraction. Implementation based on w3.org
' definition at <http://validator.w3.org/feed/docs/rss2.html#ltcategorygtSubelementOfLtitemgt>.
' 
' You may include as many category elements as you need to, for different
' domains, and to have an item cross-referenced in different parts of the same domain.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class RSS_Category
    
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
    
    ' Property: domain
    ' 
    ' A string that identifies a categorization taxonomy.
    ' 
    ' Contains:
    ' 
    '     [(string)] - A categorization taxonomy identifier
    ' 
    public domain
    
    ' Property: innerText
    ' 
    ' The value of the category.
    ' 
    ' Contains:
    ' 
    '     (string) - The value of the category
    ' 
    public innerText
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: toString
    ' 
    ' Returns the category representation to be used inside <channel> or <item>.
    ' 
    ' Parameters:
    ' 
    '     (int) - Initial indentation level
    ' 
    ' Returns:
    ' 
    '     (string) - Guid representation.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Channel_category : set Channel_category = new RSS_Category
    ' Channel_category.innerText = "Category"
    ' Response.write(Channel_category.toString(0))
    ' set Channel_category = nothing
    ' 
    ' (end code)
    ' 
    public function toString(indents)
        toString = strsubstitute("{0}<category{1}>{2}</category>", array(space(indents * 4), iif(isEmpty(domain), "", " domain='" & domain & "'"), innerText))
    end function
    
end class



' Class: RSS_Cloud
' 
' Really Simple Syndication Cloud Model Abstraction. Implementation based on w3.org
' definition at <http://validator.w3.org/feed/docs/rss2.html#ltcloudgtSubelementOfLtchannelgt>.
' 
' It specifies a web service that supports the rssCloud interface which can be
' implemented in HTTP-POST, XML-RPC or SOAP 1.1.
' 
' It's purpose is to allow processes to register with a cloud to be notified of
' updates to the channel, implementing a lightweight publish-subscribe protocol for RSS feeds.
' 
' A full explanation of this element and the rssCloud interface is <http://www.thetwowayweb.com/soapmeetsrss#rsscloudInterface>.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class RSS_Cloud
    
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
    
    ' Property: domain
    ' 
    ' A list of urls of RSS files to be watched.
    ' 
    ' Contains:
    ' 
    '     (string) - A list of urls of RSS files to be watched
    ' 
    public domain
    
    ' Property: port
    ' 
    ' The TCP port the workstation is listening on.
    ' 
    ' Contains:
    ' 
    '     (int) - The TCP port the workstation is listening on
    ' 
    public port
    
    ' Property: path
    ' 
    ' The path to it's responder.
    ' 
    ' Contains:
    ' 
    '     (string) - The path to it's responder
    ' 
    public path
    
    ' Property: procedure
    ' 
    ' The name of the procedure that the cloud should call to notify the
    ' workstation of changes.
    ' 
    ' Contains:
    ' 
    '     (string) - The name of the procedure that the cloud should call to notify the workstation of changes.
    ' 
    public procedure
    
    ' Property: protocol
    ' 
    ' A string indicating which protocol to use (xml-rpc or soap, case-sensitive)
    ' 
    ' Contains:
    ' 
    '     (xml-rpc or soap) - Protocol to use
    ' 
    public protocol
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: toString
    ' 
    ' Returns the cloud representation to be used inside <channel>.
    ' 
    ' Parameters:
    ' 
    '     (int) - Initial indentation level
    ' 
    ' Returns:
    ' 
    '     (string) - Cloud representation.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Channel_cloud : set Channel_cloud = new RSS_Cloud
    ' Channel_cloud.domain = "www.domain.com"
    ' Channel_cloud.port = 80
    ' Channel_cloud.path = "/webservices"
    ' Channel_cloud.procedure = "rssPleaseNotify"
    ' Channel_cloud.protocol = "xml-rpc"
    ' Response.write(Channel_cloud.toString(0))
    ' set Channel_cloud = nothing
    ' 
    ' (end code)
    ' 
    public function toString(indents)
        toString = strsubstitute("{0}<cloud domain='{1}' port='{2}' path='{3}' registerProcedure='{4}' protocol='{5}' />", array(space(indents * 4), domain, port, path, procedure, protocol))
    end function
    
end class



' Class: RSS_Image
' 
' Really Simple Syndication Image Model Abstraction. Implementation based on w3.org
' definition at <http://validator.w3.org/feed/docs/rss2.html#ltimagegtSubelementOfLtchannelgt>.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class RSS_Image
    
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
    
    ' --[ Requireds ]-----------------------------------------------------------
    
    ' Property: url
    ' 
    ' The URL of a GIF, JPEG or PNG image that represents the channel.
    ' 
    ' Contains:
    ' 
    '     (string) - The URL of a GIF, JPEG or PNG image that represents the channel
    ' 
    public url
    
    ' Property: title
    ' 
    ' Describes the image, it's used in the ALT attribute of the HTML <img> tag
    ' when the channel is rendered in HTML.
    ' 
    ' Contains:
    ' 
    '     (string) - Describes the image, it's used in the ALT attribute
    ' 
    public title
    
    ' Property: link
    ' 
    ' Is the URL of the site, when the channel is rendered, the image is a link
    ' to the site. (Note, in practice the image <title> and <link> should have
    ' the same value as the channel's <title> and <link>.
    ' 
    ' Contains:
    ' 
    '     (string) - The URL of the site
    ' 
    public link
    
    ' --[ Optionals ]-----------------------------------------------------------
    
    ' Property: description
    ' 
    ' Contains text that is included in the TITLE attribute of the link formed
    ' around the image in the HTML rendering.
    ' 
    ' Contains:
    ' 
    '     (string) - Contains text that is included in the TITLE attribute
    ' 
    public description
    
    ' Property: width
    ' 
    ' Indicates the width of the image in pixels(maximum value for width is 144).
    ' 
    ' Contains:
    ' 
    '     [(int)] - Indicates the width of the image in pixels. Defaults to 88
    ' 
    public width
    
    ' Property: height
    ' 
    ' Indicates the height of the image in pixels(maximum value for width is 400).
    ' 
    ' Contains:
    ' 
    '     [(int)] - Indicates the width of the image in pixels. Defaults to 31
    ' 
    public height
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: toString
    ' 
    ' Returns the image representation to be used inside <channel>.
    ' 
    ' Parameters:
    ' 
    '     (int) - Initial indentation level
    ' 
    ' Returns:
    ' 
    '     (string) - Image representation.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Channel_image : set Channel_image = new RSS_Image
    ' Channel_image.url = "http://www.domain.com/rss.png"
    ' Channel_image.title = "alt attribute value"
    ' Channel_image.link = "http://www.domain.com"
    ' Response.write(Channel_image.toString(0))
    ' set Channel_image = nothing
    ' 
    ' (end code)
    ' 
    public function toString(indents)
        dim optionals : optionals = ""
        if( not isEmpty(description) ) then optionals = optionals & ("{0}    <description>" & description & "</description>")
        if( not isEmpty(width) ) then optionals = optionals & ("{0}    <width>" & width & "</width>")
        if( not isEmpty(height) ) then optionals = optionals & ("{0}    <height>" & height & "</height>")
        optionals = strsubstitute(optionals, array(vbNewline & space(indents * 4)))
        
        toString = strsubstitute(join(array( _
            "{0}<image>", _
            "{0}    <url>{1}</url>", _
            "{0}    <title>{2}</title>", _
            "{0}    <link>{3}</link>{4}", _
            "{0}</image>" _
        ), vbNewline), array(space(indents * 4), url, title, link, optionals))
    end function
    
end class



' Class: RSS_TextInput
' 
' Really Simple Syndication TextInput Model Abstraction. Implementation based on w3.org
' definition at <http://validator.w3.org/feed/docs/rss2.html#lttextinputgtSubelementOfLtchannelgt>.
' 
' The purpose of the <textInput> element is something of a mystery. You can use 
' it to specify a search engine box. Or to allow a reader to provide feedback. 
' Most aggregators ignore it.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class RSS_TextInput
    
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
    
    ' Property: title
    ' 
    ' The label of the Submit button in the text input area.
    ' 
    ' Contains:
    ' 
    '     (string) - The label of the Submit button in the text input area
    ' 
    public title
    
    ' Property: description
    ' 
    ' Explains the text input area.
    ' 
    ' Contains:
    ' 
    '     (string) - Explains the text input area
    ' 
    public description
    
    ' Property: name
    ' 
    ' The name of the text object in the text input area.
    ' 
    ' Contains:
    ' 
    '     (string) - The name of the text object in the text input area
    ' 
    public name
    
    ' Property: link
    ' 
    ' The URL of the CGI script that processes text input requests.
    ' 
    ' Contains:
    ' 
    '     (string) - The URL of the CGI script that processes text input requests
    ' 
    public link
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: toString
    ' 
    ' Returns the textInput representation to be used inside <channel>.
    ' 
    ' Parameters:
    ' 
    '     (int) - Initial indentation level
    ' 
    ' Returns:
    ' 
    '     (string) - TextInput representation.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Channel_textInput : set Channel_textInput = new RSS_Category
    ' Channel_textInput.title = "Submit"
    ' Channel_textInput.description = "Please give your feedback"
    ' Channel_textInput.name = "Feedback"
    ' Channel_textInput.link = "http://www.domain.com/processTextInput/"
    ' Response.write(Channel_textInput.toString(0))
    ' set Channel_textInput = nothing
    ' 
    ' (end code)
    ' 
    public function toString(indents)
        toString = strsubstitute(join(array( _
            "{0}<textInput>", _
            "{0}    <title>{1}</title>", _
            "{0}    <description>{2}</description>", _
            "{0}    <name>{3}</name>", _
            "{0}    <link>{4}</link>", _
            "{0}</textInput>" _
        ), vbNewline), array(space(indents * 4), title, description, name, link))
    end function
    
end class



' Class: RSS_SkipHours
' 
' Really Simple Syndication SkipHours Model Abstraction. Implementation based on w3.org
' definition at <http://backend.userland.com/skipHoursDays#skiphours>.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class RSS_SkipHours
    
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
    
    ' Property: skip
    ' 
    ' An array from 0 to 23 which indicates the hour to skip.
    ' 
    ' Contains:
    ' 
    '     (boolean[]) - Set index i to true if you want to skip the hour starting in i
    ' 
    public skip
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        skip = array(false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false)
    end sub
    
    private sub Class_terminate()
        erase skip
    end sub
    
    ' Function: toString
    ' 
    ' Returns the skipHours representation to be used inside <channel>.
    ' 
    ' Parameters:
    ' 
    '     (int) - Initial indentation level
    ' 
    ' Returns:
    ' 
    '     (string) - SkipHours representation.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Channel_skipHours : set Channel_skipHours = new RSS_SkipHours
    ' dim i : for i = 0 to 5
    '     Channel_skipHours.skip(i) = true
    ' next
    ' Response.write(Channel_skipHours.toString(0))
    ' set Channel_skipHours = nothing
    ' 
    ' (end code)
    ' 
    public function toString(indents)
        dim hours : hours = ""
        dim i : for i = 0 to ubound(skip)
            if(skip(i)) then hours = hours & ("<hour>" & i & "</hour>")
        next
        
        toString = strsubstitute("{0}<skipHours>{1}</skipHours>", array(space(indents * 4), hours))
    end function
    
end class



' Class: RSS_SkipDays
' 
' Really Simple Syndication SkipHours Model Abstraction. Implementation based on w3.org
' definition at <http://backend.userland.com/skipHoursDays#skiphours>.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class RSS_SkipDays
    
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
    
    ' Property: skip
    ' 
    ' An array from "Sunday" to "Saturday" which indicates the weekday to skip.
    ' 
    ' Contains:
    ' 
    '     (boolean[]) - Set index i to true if you want to skip the hour starting in i
    ' 
    public skip
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        skip = array(false, false, false, false, false, false, false)
    end sub
    
    private sub Class_terminate()
        erase skip
    end sub
    
    ' Function: toString
    ' 
    ' Returns the skipDays representation to be used inside <channel>.
    ' 
    ' Parameters:
    ' 
    '     (int) - Initial indentation level
    ' 
    ' Returns:
    ' 
    '     (string) - SkipHours representation.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Channel_skipDays : set Channel_skipDays = new RSS_SkipDays
    ' Channel_skipDays.skip(0) = true
    ' Channel_skipDays.skip(6) = true
    ' Response.write(Channel_skipDays.toString(0))
    ' set Channel_skipDays = nothing
    ' 
    ' (end code)
    ' 
    public function toString(indents)
        dim days : days = ""
        if(skip(0)) then days = days & ("<day>Sunday</day>")
        if(skip(1)) then days = days & ("<day>Monday</day>")
        if(skip(2)) then days = days & ("<day>Tuesday</day>")
        if(skip(3)) then days = days & ("<day>Wednesday</day>")
        if(skip(4)) then days = days & ("<day>Thursday</day>")
        if(skip(5)) then days = days & ("<day>Friday</day>")
        if(skip(6)) then days = days & ("<day>Saturday</day>")
        
        toString = strsubstitute("{0}<skipDays>{1}</skipDays>", array(space(indents * 4), days))
    end function
    
end class



' Class: RSS_Enclosure
' 
' Really Simple Syndication Enclosure Model Abstraction. Implementation based on w3.org
' definition at <http://validator.w3.org/feed/docs/rss2.html#ltenclosuregtSubelementOfLtitemgt>.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class RSS_Enclosure
    
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
    
    ' Property: url
    ' 
    ' Where the enclosure is located.
    ' 
    ' Contains:
    ' 
    '     (string) - Where the enclosure is located
    ' 
    public url
    
    ' Property: length
    ' 
    ' Length in bytes.
    ' 
    ' Contains:
    ' 
    '     (int) - Length in bytes
    ' 
    public length
    
    ' Property: mime
    ' 
    ' A standard MIME type.
    ' 
    ' Contains:
    ' 
    '     (string) - A standard MIME type.
    ' 
    public mime
        
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: toString
    ' 
    ' Returns the enclosure representation to be used inside <item>.
    ' 
    ' Parameters:
    ' 
    '     (int) - Initial indentation level
    ' 
    ' Returns:
    ' 
    '     (string) - Enclosure representation.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Item_enclosure : set Item_enclosure = new RSS_Enclosure
    ' Item_enclosure.url = "http://www.domain.com/music.mp3"
    ' Item_enclosure.length = 1024
    ' Item_enclosure.mime = "audio/mpeg"
    ' Response.write(Item_enclosure.toString(0))
    ' set Item_enclosure = nothing
    ' 
    ' (end code)
    ' 
    public function toString(indents)
        toString = strsubstitute("{0}<enclosure url='{1}' length='{2}' type='{3}' />", array(space(indents * 4), url, length, mime))
    end function
    
end class



' Class: RSS_Guid
' 
' Really Simple Syndication Guid Model Abstraction. Implementation based on w3.org
' definition at <http://validator.w3.org/feed/docs/rss2.html#ltguidgtSubelementOfLtitemgt>.
' 
' Guid stands for globally unique identifier. It's a string that uniquely
' identifies the item. When present, an aggregator may choose to use this string
' to determine if an item is new.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class RSS_Guid
    
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
    
    ' Property: isPermaLink
    ' 
    ' isPermaLink is optional, it's default value is true. If it's value is false,
    ' the guid may not be assumed to be a url, or a url to anything in particular.
    ' 
    ' Contains:
    ' 
    '     [(boolean)] - Defines the type of the guid. Defaults to true &rArr; permalink.
    ' 
    public isPermaLink
    
    ' Property: innerText
    ' 
    ' The value of the guid.
    ' 
    ' Contains:
    ' 
    '     (string) - The value of the guid. It's usually an url.
    ' 
    public innerText
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
        
        isPermaLink = true
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: toString
    ' 
    ' Returns the guid representation to be used inside <item>.
    ' 
    ' Parameters:
    ' 
    '     (int) - Initial indentation level
    ' 
    ' Returns:
    ' 
    '     (string) - Guid representation.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Item_guid : set Item_guid = new RSS_Guid
    ' Item_guid.innerText = "http://www.domain.com/friendly-url-as-permalink"
    ' Response.write(Item_guid.toString(0))
    ' set Item_guid = nothing
    ' 
    ' (end code)
    ' 
    public function toString(indents)
        toString = strsubstitute("{0}<guid isPermaLink='{1}'>{2}</guid>", array(space(indents * 4), isPermaLink, innerText))
    end function
    
end class



' Class: RSS_Source
' 
' Really Simple Syndication Source Model Abstraction. Implementation based on w3.org
' definition at <http://validator.w3.org/feed/docs/rss2.html#ltsourcegtSubelementOfLtitemgt>.
' 
' The purpose of this element is to propogate credit for links, to publicize the
' sources of news items. It's used in the post command in the Radio UserLand
' aggregator. It should be generated automatically when forwarding an item from
' an aggregator to a weblog authoring tool.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class RSS_Source
    
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
    
    ' Property: url
    ' 
    ' Links to the XMLization of the source.
    ' 
    ' Contains:
    ' 
    '     (string) - Links to the XMLization of the source
    ' 
    public url
    
    ' Property: innerText
    ' 
    ' The value of the source.
    ' 
    ' Contains:
    ' 
    '     (string) - The value of the category
    ' 
    public innerText
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Function: toString
    ' 
    ' Returns the source representation to be used inside <item>.
    ' 
    ' Parameters:
    ' 
    '     (int) - Initial indentation level
    ' 
    ' Returns:
    ' 
    '     (string) - Source representation.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Item_source : set Item_source = new RSS_Source
    ' Item_source.url = "http://www.domain.com/links.xml"
    ' Item_source.innerText = "Domain Realm"
    ' Response.write(Item_source.toString(0))
    ' set Item_source = nothing
    ' 
    ' (end code)
    ' 
    public function toString(indents)
        toString = strsubstitute("{0}<source url='{1}'>{2}</source>", array(space(indents * 4), url, innerText))
    end function
    
end class

%>
<script language="javascript" runat="server">
function RSS_date() {
    return (new Date()).toString().replace(/UTC/, "GMT");
}
</script>
