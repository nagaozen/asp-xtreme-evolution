dp.sh.Brushes.Asp = function() {
    var keywords = 'AddHandler AddressOf AndAlso Alias And Ansi As Assembly Auto ' + 
                   'Boolean ByRef Byte ByVal Call Case Catch CBool CByte CChar CDate ' + 
                   'CDec CDbl Char CInt Class CLng CObj Const CShort CSng CStr CType ' + 
                   'Date Decimal Declare Default Delegate Dim DirectCast Do Double Each ' +
                   'Else ElseIf End Enum Erase Error Event Exit False Finally For Friend ' + 
                   'Function Get GetType GoSub GoTo Handles If Implements Imports In ' + 
                   'Inherits Integer Interface Is Let Lib Like Long Loop Me Mod Module ' + 
                   'MustInherit MustOverride MyBase MyClass Namespace New Next Not Nothing ' + 
                   'NotInheritable NotOverridable Object On Option Optional Or OrElse ' +
                   'Overloads Overridable Overrides ParamArray Preserve Private Property ' +
                   'Protected Public RaiseEvent ReadOnly ReDim REM RemoveHandler Resume ' + 
                   'Return Select Set Shadows Shared Short Single Static Step Stop String ' +
                   'Structure Sub SyncLock Then Throw To True Try TypeOf Unicode Until ' + 
                   'Variant When While With WithEvents WriteOnly Xor';
    
    var aResponse    = [".Cookies", ".Buffer", ".CacheControl", ".Charset", ".ContentType", ".Expires", ".ExpiresAbsolute","IsClientConnected", ".Pics", ".Status", ".AddHeader", ".AppendToLog", ".BinaryWrite", ".Clear", ".End", ".Flush", ".Redirect", ".Write"];
    var aRequest     = [".ClientCertificate", ".Cookies", ".Form", ".QueryString", ".ServerVariables", ".TotalBytes", ".BinaryRead"];
    var aApplication = [".Contents", ".StaticObjects", ".Contents.Remove", ".Contents.RemoveAll", ".Lock", ".Unlock", ".Application_OnEnd", ".Application_OnStart"];
    var aSession     = [".Contents", ".StaticObjects", ".CodePage", ".LCID", ".SessionID", ".Timeout", ".Abandon", ".Contents.Remove", ".Contents.RemoveAll", ".Session_OnEnd", ".Session_OnStart"];
    var aServer      = [".ScriptTimeOut", ".CreateObject", ".Execute", ".GetLastError", ".HTMLEncode", ".MapPath", ".Transfer", ".URLEncode"];
    var aError       = [".ASPCode", ".ASPDescription", ".Category", ".Column", ".Description", ".File", ".Line", ".Number", ".Source"];

    var asp_objects  = 'Response Request Application Session Server Error';
    var asp_methods  = aResponse.join(' ') + ' ' +
                       aRequest.join(' ') + ' ' +
                       aApplication.join(' ') + ' ' +
                       aSession.join(' ') + ' ' +
                       aServer.join(' ') + ' ' +
                       aError.join(' ');
    
    this.regexList = [
        {regex:new RegExp('\'.*$','gm'),css:'comment'},
        {regex:new RegExp('(\&lt;|<)!--\\s*.*?\\s*--(\&gt;|>)','gm'),css:'preprocessor'},
        {regex:dp.sh.RegexLib.DoubleQuotedString,css:'string'},
        {regex:new RegExp(this.GetKeywords(keywords),'gim'),css:'keyword'},
        {regex:new RegExp(this.GetKeywords(asp_objects),'gim'),css:'asp-object'},
        {regex:new RegExp(this.GetKeywords(asp_methods),'gim'),css:'asp-methods'}
    ];
    this.CssClass = 'dp-asp';
}
dp.sh.Brushes.Asp.prototype = new dp.sh.Highlighter();
dp.sh.Brushes.Asp.Aliases = ['asp'];
