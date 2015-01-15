(function(){

/* See: https://developer.mozilla.org/en-US/docs/Web/API/document.cookie */
var Cookies = {
  removeItem: function (sKey, sPath, sDomain) {
    if (!this.hasItem(sKey)) { return false; }
    document.cookie = encodeURIComponent(sKey) + "=; expires=Thu, 01 Jan 1970 00:00:00 GMT" + (sDomain ? "; domain=" + sDomain : "") + (sPath ? "; path=" + sPath : "/");
    return true;
  },
  hasItem: function (sKey) {
    if (!sKey) { return false; }
    return (new RegExp("(?:^|;\\s*)" + encodeURIComponent(sKey).replace(/[\-\.\+\*]/g, "\\$&") + "\\s*\\=")).test(document.cookie);
  }
};

/* clear client side aspsessionids */
var aspsessionids = document.cookie.match(/ASPSESSIONID[^=]+/g)
  , aspsessionid = null
  , i = 0
;

if( aspsessionids !== null ) {
  i = aspsessionids.length;
  while(i--) {
    aspsessionid = aspsessionids[i];
    Cookies.removeItem(aspsessionid,"/");
  }
}

}());
