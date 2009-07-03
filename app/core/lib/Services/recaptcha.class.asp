<%

' Class: ReCaptcha
' 
' Provides a client for the Carnegie Mellon University reCAPTCHA Web Service.
' Per the reCAPTCHA site, "reCAPTCHA is a free CAPTCHA service that helps to
' digitize books." Each reCAPTCHA requires the user to input two words, the
' first of which is the actual captcha, and the second of which is a word from
' some scanned text that Optical Character Recognition (OCR) software has been
' unable to identifiy. The assumption is that if a user correctly provides the
' first word, the second is likely correctly entered as well, and can be used to
' improve OCR software for digitizing books.
' 
' In order to use the reCAPTCHA service, you will need to sign up for an account
' and register one or more domains with the service in order to generate public
' and private keys.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao  @ October 2008
' 
class ReCaptcha
    
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
    
    ' Property: domainName
    ' 
    ' reCAPTCHA will only work on this domain and subdomains.
    ' 
    ' Contains:
    ' 
    '     (string) - Your site domain (no "http://")
    ' 
    public domainName
    
    ' Property: uriVerify
    ' 
    ' The URL of the reCAPTCHA_verify.asp
    ' 
    ' Contains:
    ' 
    '     (string) - The URL of the reCAPTCHA_verify.asp
    ' 
    public uriVerify
    
    ' Property: publicKey
    ' 
    ' reCAPTCHA provides a public key to be used in the javascript widget. If
    ' you don't have one yet go get it at <http://recaptcha.net/api/getkey>.
    ' 
    ' Contains:
    ' 
    '     (string) - Your public key
    ' 
    public publicKey
    
    ' Property: privateKey
    ' 
    ' reCAPTCHA provides a private key to be used in the javascript widget. If
    ' you don't have one yet go get it at <http://recaptcha.net/api/getkey>.
    ' 
    ' Contains:
    ' 
    '     (string) - Your private key
    ' 
    public privateKey
    
    ' Property: theme
    ' 
    ' Sets the interface theme.
    ' 
    ' Contains:
    ' 
    '     (string) - { "red", "white", "blackglass", "clean", "custom" }
    ' 
    public theme
    
    ' Property: lang
    ' 
    ' Sets the interface language.
    ' 
    ' Contains:
    ' 
    '     (string) - { "en", "nl", "fr", "de", "pt", "ru", "es", "tr" }
    ' 
    public lang
    
    ' Property: status
    ' 
    ' reCAPTCHA verify returns a second information which is the status id. For
    ' more details, please visit <http://recaptcha.net/apidocs/captcha/>
    ' 
    ' Contains:
    ' 
    '     (string) - status
    ' 
    public status
    
    private uriService
    
    private sub Class_initialize()
        classType    = typeName(Me)
        classVersion = "1.0.0"
        
        uriService   = "http://api-verify.recaptcha.net/verify"
        theme        = "red"
        lang         = "en"
        status       = ""
    end sub
    
    private sub Class_terminate()
    end sub
    
    ' Subroutine: insert
    ' 
    ' Use this subroutine in the document body to insert the reCAPTCHA in your page.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Captcha : set Captcha = new ReCaptcha
    ' Captcha.domainName = "<domain name>"
    ' Captcha.uriVerify  = "<reCAPTCHA_verify.asp URL>"
    ' Captcha.publicKey  = "<public key>"
    ' Captcha.privateKey = "<private key>"
    ' 
    ' call Captcha.insert()
    ' 
    ' set Captcha = nothing
    ' 
    ' (end code)
    ' 
    public sub insert()
        %>
<script type="text/javascript" src="http://api.recaptcha.net/js/recaptcha_ajax.js"></script>
<script type="text/javascript">
// <![CDATA[
(function() {
    if(window.ActiveXObject && !window.XMLHttpRequest) {
        window.XMLHttpRequest = function() {
            return new ActiveXObject("MSXML2.XMLHTTP");
        }
    }
    
    var sjat = function(url, data) {
        var R = new XMLHttpRequest();
        var method = (data == null) ? "GET" : "POST";
        R.open(method, url, false);
        if(data != null) R.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        R.send(data);
        if(R.status == 200) return R.responseText;
    }
    
    document.write('<div id="reCAPTCHA-container"></div>');
    
    Recaptcha.create("<%= publicKey %>", "reCAPTCHA-container", {
        theme: "<%= theme %>",
        lang: "<%= lang %>"
    });
    
    Recaptcha.test = function() {
        return sjat("<%= uriVerify %>", "challenge=" + Recaptcha.get_challenge() + "&response=" + Recaptcha.get_response());
    }
})();
// ]]>
</script>
        <%
    end sub
    
    ' Function: verify
    ' 
    ' Contacts the reCAPTCHA service and get it's response.
    ' 
    ' Parameters:
    ' 
    '     (string) - challenge hash
    '     (string) - user input
    ' 
    ' Returns:
    ' 
    '     (boolean) - true if the input is right, false otherwise
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim Captcha : set Captcha = new ReCaptcha
    ' Captcha.domainName = "<domain name>"
    ' Captcha.uriVerify  = "<reCAPTCHA_verify.asp URL>"
    ' Captcha.publicKey  = "<public key>"
    ' Captcha.privateKey = "<private key>"
    ' 
    ' if( Captcha.verify( cstr(Request.Form("challenge")), cstr(Request.Form("response")) ) ) then
    '     Response.write("Human access")
    ' else
    '     Response.write("Robot access")
    ' end if
    ' 
    ' set Captcha = nothing
    ' 
    ' (end code)
    ' 
    public function verify(sChallenge, sResponse)
        dim Xhr : set Xhr = Server.createObject("MSXML2.ServerXMLHTTP.6.0")
        Xhr.open "POST", uriService, false
        Xhr.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        Xhr.send( join(array( _
            "privatekey=" & privateKey, _
            "remoteip=" & Request.ServerVariables("REMOTE_ADDR"), _
            "challenge=" & sChallenge, _
            "response=" & sResponse _
        ), "&"))
        dim answer : answer = split(Xhr.responseText, vbLf)
        set Xhr = nothing
        
        verify = answer(0)
        status = answer(1)
    end function
    
end class

%>