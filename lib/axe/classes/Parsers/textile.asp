<%

' File: textile.asp
' 
' AXE(ASP Xtreme Evolution) implementation of Textile parser.
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



' Class: Textile
' 
' Textile is a lightweight markup language originally developed by Dean Allen 
' and billed as a "humane Web text generator". Textile converts its marked-up 
' text input to valid, well-formed XHTML and also inserts character entity 
' references for apostrophes, opening and closing single and double quotation 
' marks, ellipses and em dashes.
' 
' About:
' 
'     - This class uses the Textile javascript implementation of Ben Daglish
'     - More info about textile visit <http://textile.thresholdstate.com>
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Dec 2010
' 
class Textile
    
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
    
    private [_ζ]
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0"
        
        set [_ζ] = new_Textile()
    end sub
    
    private sub Class_terminate()
        set [_ζ] = nothing
    end sub
    
    ' Function: makeHtml
    ' 
    ' Converts Textile into XHTML.
    ' 
    ' Parameters:
    ' 
    '     (string) - textile
    ' 
    ' Returns:
    ' 
    '     (string) - html
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim sTextile : sTextile = "Textile _rocks_."
    ' dim Converter : set Converter = new Textile
    ' 
    ' Response.write Converter.makeHtml(sTextile)
    ' 
    ' set Converter = nothing
    ' 
    ' (end code)
    ' 
    public function makeHtml(text)
        makeHtml = [_ζ].makeHtml(text)
    end function
    
end class

%>
<script language="javascript" runat="server">

function new_Textile() {
    return ( new Textile_Js() );
}

var Textile_Js = (function(){
    
    
    
    
    
/* Ben Daglish <http://www.ben-daglish.net> code starts here */

/* Javascript Textile->HTML conversion
 *
 * ben@ben-daglish.net (with thanks to John Hughes for improvements)
 * Issued under the "do what you like with it - I take no respnsibility" licence
 * 
 * Minor fixed and enhancements by Fabio Zendhi Nagao (nagaozen)
 */

var inpr, inbq, inbqq, html,
    aliases = [],
    alg = {
        '>': 'right',
        '<': 'left',
        '=': 'center',
        '<>': 'justify',
        '~': 'bottom',
        '^': 'top'
    },
    ent = {
        "'": "&#8217;",
        " - ": " &#8211; ",
        "--": "&#8212;",
        " x ": " &#215; ",
        "\\.\\.\\.": "&#8230;",
        "\\(C\\)": "&#169;",
        "\\(R\\)": "&#174;",
        "\\(TM\\)": "&#8482;"
    },
    tags = {
        "b": "\\*\\*",
        "i": "__",
        "em": "_",
        "strong": "\\*",
        "cite": "\\?\\?",
        "sup": "\\^",
        "sub": "~",
        "span": "\\%",
        "del": "-",
        "code": "@",
        "ins": "\\+",
        "del": "-"
    },
    le = "\n\n",
    lstlev = 0,
    lst = "",
    elst = "",
    intable = 0,
    mm = "",
    para = /^p(\S*)\.\s*(.*)/,
    rfn = /^fn(\d+)\.\s*(.*)/,
    bq = /^bq\.(\.)?\s*/,
    table = /^table\s*{(.*)}\..*/,
    trstyle = /^\{(\S+)\}\.\s*\|/;

function convert(t) {
    var lines = t.split("\n");
    html = [];
    inpr = inbq = inbqq = 0;
    for(var i = 0; i < lines.length; i++) {
        if(lines[i].indexOf("[") == 0) {
            var m = lines[i].indexOf("]");
            aliases[lines[i].substring(1, m)] = lines[i].substring(m + 1);
        }
    }
    for(i = 0; i < lines.length; i++) {
        if(lines[i].indexOf("[") == 0) {
            continue;
        }
        if(mm = para.exec(lines[i])) {
            stp(1);
            inpr = 1;
            html.push( lines[i].replace(para, "<p" + make_attr(mm[1]) + ">" + prep(mm[2])) );
            continue;
        }
        if(mm = /^h(\d)(\S*)\.\s*(.*)/.exec(lines[i])) {
            stp(1);
            html.push( tag("h" + mm[1], make_attr(mm[2]), prep(mm[3])) + le );
            continue;
        }
        if(mm = rfn.exec(lines[i])) {
            stp(1);
            inpr = 1;
            html.push( lines[i].replace(rfn, '<p id="fn' + mm[1] + '"><sup>' + mm[1] + '<\/sup>' + prep(mm[2])) );
            continue;
        }
        if(lines[i].indexOf("*") == 0) {
            lst = "<ul>";
            elst = "<\/ul>";
        }
        else if(lines[i].indexOf("#") == 0) {
            lst = "<ol>";
            elst = "<\/ol>";
        }
        else {
            while (lstlev > 0) {
                html.push( elst );
                if(lstlev > 1) {
                    html.push( "<\/li>" );
                } else {
                    html.push( "\n" );
                }
                html.push( "\n" );
                lstlev--;
            }
            lst = "";
        }
        if(lst) {
            stp(1);
            var m = /^([*#]+)\s*(.*)/.exec(lines[i]);
            var lev = m[1].length;
            while (lev < lstlev) {
                html.push( elst + "<\/li>\n" );
                lstlev--;
            }
            while (lstlev < lev) {
                html = [ html.join("").replace(/<\/li>\n$/, "\n") ];
                html.push( lst );
                lstlev++;
            }
            html.push( tag("li", "", prep(m[2])) + "\n" );
            continue;
        }
        if(lines[i].match(table)) {
            stp(1);
            intable = 1;
            html.push( lines[i].replace(table, '<table style="$1;">\n') );
            continue;
        }
        if((lines[i].indexOf("|") == 0) || (lines[i].match(trstyle))) {
            stp(1);
            if(!intable) {
                html.push( "<table>\n" );
                intable = 1;
            }
            var rowst = "";
            var trow = "";
            var ts = trstyle.exec(lines[i]);
            if(ts) {
                rowst = qat('style', ts[1]);
                lines[i] = lines[i].replace(trstyle, "\|");
            }
            var cells = lines[i].split("|");
            for(j = 1; j < cells.length - 1; j++) {
                var ttag = "td";
                if(cells[j].indexOf("_.") == 0) {
                    ttag = "th";
                    cells[j] = cells[j].substring(2);
                }
                cells[j] = prep(cells[j]);
                var al = /^([<>=^~\/\\\{]+.*?)\.(.*)/.exec(cells[j]);
                var at = "",
                    st = "";
                if(al != null) {
                    cells[j] = al[2];
                    var cs = /\\(\d+)/.exec(al[1]);
                    if(cs != null) {
                        at += qat('colspan', cs[1]);
                    }
                    var rs = /\/(\d+)/.exec(al[1]);
                    if(rs != null) {
                        at += qat('rowspan', rs[1]);
                    }
                    var va = /([\^~])/.exec(al[1]);
                    if(va != null) {
                        st += "vertical-align:" + alg[va[1]] + ";";
                    }
                    var ta = /(<>|=|<|>)/.exec(al[1]);
                    if(ta != null) {
                        st += "text-align:" + alg[ta[1]] + ";";
                    }
                    var is = /\{([^\}]+)\}/.exec(al[1]);
                    if(is != null) {
                        st += is[1];
                    }
                    if(st != "") {
                        at += qat('style', st);
                    }
                }
                trow += tag(ttag, at, cells[j]);
            }
            html.push( "\t" + tag("tr", rowst, trow) + "\n" );
            continue;
        }
        if(intable) {
            html.push( "<\/table>" + le );
            intable = 0;
        }
        if(lines[i] == "") {
            stp();
        }
        else if(!inpr) {
            if(mm = bq.exec(lines[i])) {
                lines[i] = lines[i].replace(bq, "");
                html.push( "<blockquote>" );
                inbq = 1;
                if(mm[1]) {
                    inbqq = 1;
                }
            }
            html.push( "<p>" + prep(lines[i]) );
            inpr = 1;
        }
        else {
            html.push( prep(lines[i]) );
        }
    }
    stp();
    return html.join("");
}

function prep(m) {
    for(i in ent) {
        m = m.replace(new RegExp(i, "g"), ent[i]);
    }
    for(i in tags) {
        m = make_tag(m, RegExp("^" + tags[i] + "(.+?)" + tags[i]), i, "");
        m = make_tag(m, RegExp(" " + tags[i] + "(.+?)" + tags[i]), i, " ");
    }
    m = m.replace(/\[(\d+)\]/g, '<sup><a href="#fn$1">$1<\/a><\/sup>');
    m = m.replace(/([A-Z]+)\((.*?)\)/g, '<acronym title="$2">$1<\/acronym>');
    m = m.replace(/\"([^\"]+)\":((http|https|mailto):\S+)/g, '<a href="$2">$1<\/a>');
    m = make_image(m, /!([^!\s]+)!:(\S+)/);
    m = make_image(m, /!([^!\s]+)!/);
    m = m.replace(/"([^\"]+)":(\S+)/g, function($0, $1, $2) {
        return tag("a", qat('href', aliases[$2]), $1)
    });
    m = m.replace(/(=)?"([^\"]+)"/g, function($0, $1, $2) {
        return ($1) ? $0 : "&#8220;" + $2 + "&#8221;"
    });
    return m;
}

function make_tag(s, re, t, sp) {
    while (m = re.exec(s)) {
        var st = make_attr(m[1]);
        m[1] = m[1].replace(/^[\[\{\(]\S+[\]\}\)]/g, "");
        m[1] = m[1].replace(/^[<>=()]+/, "");
        s = s.replace(re, sp + tag(t, st, m[1]));
    }
    return s;
}

function make_image(m, re) {
    var ma = re.exec(m);
    if(ma != null) {
        var attr = "";
        var st = "";
        var at = /\((.*)\)$/.exec(ma[1]);
        if(at != null) {
            attr = qat('alt', at[1]) + qat("title", at[1]);
            ma[1] = ma[1].replace(/\((.*)\)$/, "");
        }
        if(ma[1].match(/^[><]/)) {
            st = "float:" + ((ma[1].indexOf(">") == 0) ? "right;" : "left;");
            ma[1] = ma[1].replace(/^[><]/, "");
        }
        var pdl = /(\(+)/.exec(ma[1]);
        if(pdl) {
            st += "padding-left:" + pdl[1].length + "em;";
        }
        var pdr = /(\)+)/.exec(ma[1]);
        if(pdr) {
            st += "padding-right:" + pdr[1].length + "em;";
        }
        if(st) {
            attr += qat('style', st);
        }
        var im = '<img src="' + ma[1] + '"' + attr + " /" + '>';
        if(ma.length > 2) {
            im = tag('a', qat('href', ma[2]), im);
        }
        m = m.replace(re, im);
    }
    return m;
}

function make_attr(s) {
    var st = "";
    var at = "";
    if(!s) {
        return "";
    }
    var l = /\[(\w\w)\]/.exec(s);
    if(l != null) {
        at += qat('lang', l[1]);
    }
    var ci = /\((\S+)\)/.exec(s);
    if(ci != null) {
        s = s.replace(/\((\S+)\)/, "");
        at += ci[1].replace(/#(.*)$/, ' id="$1"').replace(/^(\S+)/, ' class="$1"');
    }
    var ta = /(<>|=|<|>)/.exec(s);
    if(ta) {
        st += "text-align:" + alg[ta[1]] + ";";
    }
    var ss = /\{(\S+)\}/.exec(s);
    if(ss) {
        st += ss[1];
        if(!ss[1].match(/;$/)) {
            st += ";";
        }
    }
    var pdl = /(\(+)/.exec(s);
    if(pdl) {
        st += "padding-left:" + pdl[1].length + "em;";
    }
    var pdr = /(\)+)/.exec(s);
    if(pdr) {
        st += "padding-right:" + pdr[1].length + "em;";
    }
    if(st) {
        at += qat('style', st);
    }
    return at;
}

function qat(a, v) {
    return ' ' + a + '="' + v + '"';
}

function tag(t, a, c) {
    return "<" + t + a + ">" + c + "</" + t + ">";
}

function stp(b) {
    if(b) {
        inbqq = 0;
    }
    if(inpr) {
        html.push( "<\/p>" + le );
        inpr = 0;
    }
    if(inbq && !inbqq) {
        html.push( "<\/blockquote>" + le );
        inbq = 0;
    }
}
/* Code ends here */
    
    
    
    
    
    return {
        makeHtml: convert
    }
});

</script>
