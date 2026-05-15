VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.7
'Site:         http://sandsprite.com
'License:      MIT


Option Explicit

Public scJson As Object
Private m_init As Boolean

Private scBeautify As Object
Public tag 'do not use externally has to be public though..


'this is default property for class, thanks xiaoyao
Property Get getVal(ByVal path As String) As Variant
    Dim tmp
    GetVal2 path, tmp
    getVal = tmp
End Property

Property Let getVal(ByVal path As String, v As Variant)
    On Error Resume Next
    setVal path, v
End Property

'false means could not CreateObj(script control)
Property Get Initilized() As Boolean
    Initilized = m_init
End Property

Function getType(ByVal path As String) As String
    On Error Resume Next
    getType = scJson.Eval("typeof json." & path)
End Function

Function isArray(ByVal path As String) As Boolean
    On Error Resume Next
    Dim tmp As String
    isArray = scJson.Eval("json." & path & " instanceof Array")
    'tmp = scJson.eval("json." & path & ".constructor")
    'If InStr(tmp, "Array") > 0 Then isArray = True
End Function

Function stringify(Optional ByVal path As String = Empty) As String
    On Error Resume Next
    
    Static stringify_script As String
    
    'loaded once on demand (per class instance)
    If Len(stringify_script) = 0 Then stringify_script = getStringifyScript()
    
    If Len(path) > 0 Then
        If Left(path, 1) <> "." Then path = "." & path
    End If
    
    'ok if called twice without sc.reset
    scJson.AddCode stringify_script
    
    If Err.Number <> 0 Then
        Debug.Print "JSON object: " & jsErr
        Exit Function
    End If
        
    stringify = scJson.Eval("JSON.stringify(json" & path & ")")
    
End Function


'does this offer any advantage of LoadJson ? - also we single quote wrap it which could be a problem without checks
'Function parse(ByVal json As String) As Boolean
'    On Error Resume Next
'
'    If Len(stringify_script) = 0 Then
'        parse = LoadJson(json)
'    Else
'         scJson.AddCode "json = JSON.parse('" & json & "')"
'         parse = (Err.Number = 0)
'    End If
'
'End Function

Function enumProps(Optional ByVal path As String = Empty) As Collection

    On Error Resume Next
    Dim c As New Collection
    Dim js As String, tmp As String, t() As String, tt
    
    'scJson.addobject "c", c 'cant run twice?
    
    If Len(path) > 0 Then
        If Left(path, 1) <> "." Then path = "." & path
    End If
    
    js = "tmp=''; for(var propertyName in json" & path & ") {tmp+=propertyName+','}"
    scJson.AddCode js
    tmp = scJson.Eval("tmp")
    
    If Len(tmp) > 1 Then
        t = Split(tmp, ",")
        For Each tt In t
            If Len(tt) > 0 Then c.Add tt
        Next
    End If
    
    Set enumProps = c
    
End Function

Property Get jsErr() As String
    On Error Resume Next
    With scJson.Error
        jsErr = "Error: " & .Number & " Line:" & .Line & ":" & .Column & " - " & .Description & " " & .text
    End With
End Property

Private Sub Class_Initialize()
    On Error Resume Next

    InitSC scJson
    If Not scJson Is Nothing Then m_init = True
        
    'toVBSource App.path & "\json2_min.js"
    'toVBSource App.path & "\beautify_min.js"
        
End Sub

Private Function InitSC(o As Object) As Boolean
    On Error Resume Next
    Set o = CreateObject("MSScriptControl.ScriptControl")
    o.Language = "JScript"
    o.UseSafeSubset = True
    InitSC = (Err.Number = 0)
End Function

Function loadJson(ByVal json As String) As Boolean
    
    On Error Resume Next
    
    scJson.Reset
    
    json = "var None=''; var True=1; var False=0; var json = " & json
    
    scJson.AddCode json
    loadJson = (Err.Number = 0)
    
End Function

Function GetObjFields(obj As Object, fieldsCSV As String, Optional path As String) As Boolean
    
    Dim f() As String, ff, tmp
    
    f = Split(fieldsCSV, ",")
    For Each ff In f
        ff = Trim(ff)
        If Len(ff) > 0 Then
            If Len(path) > 0 Then
                If Not GetVal2(path & ff, tmp, False) Then Err.Raise 1, "GetObjFields", "Field not exist: " & ff
            Else
                If Not GetVal2(ff, tmp, True) Then Err.Raise 1, "GetObjFields", "Field not exist: " & ff
            End If
            CallByName obj, ff, VbLet, tmp
        End If
    Next
    
End Function

Function GetVal2(ByVal path As String, outVar, Optional CheckExists As Boolean = False) As Boolean
    
    On Error Resume Next
    
    Dim c As String
    
    If CheckExists Then
        If Not Me.exists(path) Then Exit Function
    End If
    
    c = Left(path, 1)
    If c <> "[" And c <> "." Then path = "." & path
    
    outVar = Empty
    outVar = scJson.Eval("json" & path)
    GetVal2 = (Err.Number = 0) And (scJson.Error.Number = 0) 'not working if doesnt exist...
    
End Function

Function setVal(ByVal path As String, value, Optional CheckExists As Boolean = False) As Boolean
    
    On Error Resume Next
    
    Dim c As String, skipQuote As Boolean
    
    If CheckExists Then
        If Not Me.exists(path) Then Exit Function
    End If
    
    c = Left(path, 1)
    If c <> "[" And c <> "." Then path = "." & path
    
    If TypeName(value) = "String" Then
        value = """" & Replace(value, """", "\""") & """"
    End If
    
    scJson.Eval "json" & path & " = " & value
    setVal = (Err.Number = 0) And (scJson.Error.Number = 0) 'not working if doesnt exist...
    
End Function

Function exists(ByVal path As String) As Boolean

    On Error Resume Next
    
    Dim tmp() As String, parentPath As String, field As String
    
    If Left(path, 1) = "." Then path = Mid(path, 2)
    If Right(path, 1) = "." Then path = Mid(path, 1, Len(path) - 1)
    
    If InStr(path, ".") > 0 Then
        tmp = Split(path, ".")
        If UBound(tmp) >= 1 Then
            field = tmp(UBound(tmp))
            ReDim Preserve tmp(UBound(tmp) - 1)
            parentPath = "." & Join(tmp, ".")
        End If
    Else
        field = path
    End If

    exists = scJson.Eval("json" & parentPath & ".hasOwnProperty('" & field & "')")
    
End Function


Function beautify(ByVal js As String) As String
    Dim tmp As String
    If beautify2(js, tmp) Then
        beautify = tmp
    Else
        beautify = js
    End If
End Function


Function beautify2(ByVal js As String, ByRef outVar As String) As Boolean

    On Error GoTo hell

    Dim tmp As String
    Static beautifier As String
    
    outVar = Empty
    If scBeautify Is Nothing Then InitSC scBeautify
    If scBeautify Is Nothing Then Exit Function
    
    'loaded once on demand (per class instance)
    If Len(beautifier) = 0 Then beautifier = getBeautifyScript()

    scBeautify.Reset
    Me.tag = js
    scBeautify.AddCode beautifier
    scBeautify.AddObject "sc", Me, True 'scriptcontrol.tag must be a property of the COM obj container not avail if CreateObj
    scBeautify.AddCode "sc.tag = js_beautify(sc.tag, {indent_size: 1, indent_char: '\t'}).split('\n').join('\r\n');"
    outVar = Me.tag
    Me.tag = Empty
    beautify2 = True

    Exit Function

hell:
    outVar = "Error: " & Err.Description

End Function


'--------------------------------------------------------------------------------------------------
'--------------------  private library functions and embedded sceripts below here  ----------------
'--------------------------------------------------------------------------------------------------

Private Function toVBSource(ByVal fname)
    On Error Resume Next
    Dim t() As String, o() As String, tt, i As Long, tmp As String, outFile As String
    Dim z As Long
    
    If Not FileExists(fname) Then Exit Function
    
    tmp = ReadFile(fname)
    outFile = fname & ".vb"
    
    t = Split(tmp, vbCrLf)
    If UBound(t) = 0 Then 'minimized one liner
        
        z = 1
        While z < Len(tmp)
            tt = Mid(tmp, z, 500)
            z = z + 500
            push o, vbTab & " push o, """ & Replace(tt, """", """""") & """"
        Wend
        
    Else
        For Each tt In t
            i = i + 1
            push o, i & vbTab & " push o, """ & Replace(tt, """", """""") & """"
        Next
    End If
    
    WriteFile outFile, Join(o, vbCrLf)

End Function

Private Function FileExists(ByVal path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Private Function ReadFile(filename)
  Dim f, temp
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Sub WriteFile(path, it)
    Dim f
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

'Minimizer: https://www.toptal.com/developers/javascript-minifier

'License: Public Domain
'Author:  douglas crockford
'Date:    2023-05-10
'Site:    https://github.com/douglascrockford/JSON-js
'
Private Function getStringifyScript()
     Dim o() As String
    
     push o, """object""!=typeof JSON&&(JSON={}),function(){""use strict"";var gap,indent,meta,rep,rx_one=/^[\],:{}\s]*$/,rx_two=/\\(?:[""\\\/bfnrt]|u[0-9a-fA-F]{4})/g,rx_three=/""[^""\\\n\r]*""|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g,rx_four=/(?:^|:|,)(?:\s*\[)+/g,rx_escapable=/[\\""\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,rx_dangerous=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufe"
     push o, "ff\ufff0-\uffff]/g;function f(t){return t<10?""0""+t:t}function this_value(){return this.valueOf()}function quote(t){return rx_escapable.lastIndex=0,rx_escapable.test(t)?'""'+t.replace(rx_escapable,function(t){var e=meta[t];return""string""==typeof e?e:""\\u""+(""0000""+t.charCodeAt(0).toString(16)).slice(-4)})+'""':'""'+t+'""'}function str(t,e){var n,o,f,u,r,$=gap,i=e[t];switch(i&&""object""==typeof i&&""function""==typeof i.toJSON&&(i=i.toJSON(t)),""function""==typeof rep&&(i=rep.call(e,t,i)),typeof i){case""str"
     push o, "ing"":return quote(i);case""number"":return isFinite(i)?String(i):""null"";case""boolean"":case""null"":return String(i);case""object"":if(!i)return""null"";if(gap+=indent,r=[],""[object Array]""===Object.prototype.toString.apply(i)){for(n=0,u=i.length;n<u;n+=1)r[n]=str(n,i)||""null"";return f=0===r.length?""[]"":gap?""[\n""+gap+r.join("",\n""+gap)+""\n""+$+""]"":""[""+r.join("","")+""]"",gap=$,f}if(rep&&""object""==typeof rep)for(n=0,u=rep.length;n<u;n+=1)""string""==typeof rep[n]&&(f=str(o=rep[n],i))&&r.push(quote(o)+(gap?"": "":"":"
     push o, """)+f);else for(o in i)Object.prototype.hasOwnProperty.call(i,o)&&(f=str(o,i))&&r.push(quote(o)+(gap?"": "":"":"")+f);return f=0===r.length?""{}"":gap?""{\n""+gap+r.join("",\n""+gap)+""\n""+$+""}"":""{""+r.join("","")+""}"",gap=$,f}}""function""!=typeof Date.prototype.toJSON&&(Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+""-""+f(this.getUTCMonth()+1)+""-""+f(this.getUTCDate())+""T""+f(this.getUTCHours())+"":""+f(this.getUTCMinutes())+"":""+f(this.getUTCSeconds())+""Z"":null},Boolean.proto"
     push o, "type.toJSON=this_value,Number.prototype.toJSON=this_value,String.prototype.toJSON=this_value),""function""!=typeof JSON.stringify&&(meta={""\b"":""\\b"",""   "":""\\t"",""\n"":""\\n"",""\f"":""\\f"",""\r"":""\\r"",'""':'\\""',""\\"":""\\\\""},JSON.stringify=function(t,e,n){var o;if(gap="""",indent="""",""number""==typeof n)for(o=0;o<n;o+=1)indent+="" "";else""string""==typeof n&&(indent=n);if(rep=e,e&&""function""!=typeof e&&(""object""!=typeof e||""number""!=typeof e.length))throw Error(""JSON.stringify"");return str("""",{"""":t})}),""function""!"
     push o, "=typeof JSON.parse&&(JSON.parse=function(text,reviver){var j;function walk(t,e){var n,o,f=t[e];if(f&&""object""==typeof f)for(n in f)Object.prototype.hasOwnProperty.call(f,n)&&(void 0!==(o=walk(f,n))?f[n]=o:delete f[n]);return reviver.call(t,e,f)}if(text=String(text),rx_dangerous.lastIndex=0,rx_dangerous.test(text)&&(text=text.replace(rx_dangerous,function(t){return""\\u""+(""0000""+t.charCodeAt(0).toString(16)).slice(-4)})),rx_one.test(text.replace(rx_two,""@"").replace(rx_three,""]"").replace(rx_four,"""""
     push o, ")))return j=eval(""(""+text+"")""),""function""==typeof reviver?walk({"""":j},""""):j;throw SyntaxError(""JSON.parse"")})}();"
     
     getStringifyScript = Join(o, "")
    
End Function

'License:   MIT
'Author:    Einar Lielmanis, <einar@jsbeautifier.org>, Originally converted to js by Vital, <vital76@gmail.com>
'Site:      http://jsbeautifier.org/
'
Private Function getBeautifyScript() As String

    Dim o() As String
    
1    push o, "function js_beautify(e,n){var r,t,a,i,T,E,o,f,c,O,K,l,s,N,R,h,_,d,u,S,A,v,$=!!(n=n||{}).braces_on_own_line&&n.braces_on_own_line,C=n.indent_size?n.indent_size:4,I=n.indent_char?n.indent_char:"" "",L=void 0===n.preserve_newlines||n.preserve_newlines,P=n.indent_level?n.indent_level:0,p=""undefined""!==n.space_after_anon_function&&n.space_after_anon_function,D=void 0===n.keep_array_indentation||n.keep_array_indentation;A=!1;var m=e.length;function b(){for(;t.length&&("" ""===t[t.length-1]||t[t.length-1]="
2    push o, "==O);)t.pop()}function X(e){return""[EXPRESSION]""===e||""[INDENTED-EXPRESSION]""===e}function k(e){if(f.eat_next_space=!1,!(D&&X(f.mode))){if(e=void 0===e||e,f.if_line=!1,b(),t.length){""\n""===t[t.length-1]&&e||(A=!0,t.push(""\n""));for(var n=0;n<f.indentation_level;n+=1)t.push(O);f.var_line&&f.var_line_reindented&&("" ""===I?t.push(""    ""):t.push(O))}}}function B(){if(f.eat_next_space){f.eat_next_space=!1;return}var e="" "";t.length&&(e=t[t.length-1]),"" ""!==e&&""\n""!==e&&e!==O&&t.push("" "")}function M(){A="
3    push o, "!1,f.eat_next_space=!1,t.push(a)}function W(){f.indentation_level+=1}function g(){t.length&&t[t.length-1]===O&&t.pop()}function w(e){f&&c.push(f),f={previous_mode:f?f.mode:""BLOCK"",mode:e,var_line:!1,var_line_tainted:!1,var_line_reindented:!1,in_html_comment:!1,if_line:!1,in_case:!1,eat_next_space:!1,indentation_baseline:-1,indentation_level:f?f.indentation_level+(f.var_line&&f.var_line_reindented?1:0):P}}function X(e){return""[EXPRESSION]""===e||""[INDENTED-EXPRESSION]""===e}function x(e){return""[EX"
4    push o, "PRESSION]""===e||""[INDENTED-EXPRESSION]""===e||""(EXPRESSION)""===e}function y(){u=""DO_BLOCK""===f.mode,c.length>0&&(f=c.pop())}function U(e,n){for(var r=0;r<n.length;r+=1)if(n[r]===e)return!0;return!1}function F(){for(var e=0,n=0,r=t.length-1;r>=0;r--)switch(t[r]){case"":"":0===e&&n++;break;case""?"":if(0===e){if(0===n)return!0;n--}break;case""{"":if(0===e)return!1;e--;break;case""("":case""["":e--;break;case"")"":case""]"":case""}"":e++}}function G(){if(v=0,N>=m)return["""",""TK_EOF""];S=!1;var e=r.charAt(N);if(N+=1,D"
5    push o, "&&X(f.mode)){for(var n,a=0;U(e,K);){if(""\n""===e?(b(),t.push(""\n""),A=!0,a=0):""    ""===e?a+=4:a+=1,N>=m)return["""",""TK_EOF""];e=r.charAt(N),N+=1}if(-1===f.indentation_baseline&&(f.indentation_baseline=a),A){for(n=0;n<f.indentation_level+1;n+=1)t.push(O);if(-1!==f.indentation_baseline)for(n=0;n<a-f.indentation_baseline;n++)t.push("" "")}}else{for(;U(e,K);){if(""\n""===e&&(v+=1),N>=m)return["""",""TK_EOF""];e=r.charAt(N),N+=1}if(L&&v>1)for(n=0;n<v;n+=1)k(0===n),A=!0;S=v>0}if(U(e,l)){if(N<m)for(;U(r.charAt(N),l)&"
6    push o, "&(e+=r.charAt(N),(N+=1)!==m););if(N!==m&&e.match(/^[0-9]+[Ee]$/)&&(""-""===r.charAt(N)||""+""===r.charAt(N))){var E=r.charAt(N),o=G(N+=1);return[e+=E+o[0],""TK_WORD""]}return""in""===e?[e,""TK_OPERATOR""]:(S&&""TK_OPERATOR""!==i&&!f.if_line&&(L||""var""!==T)&&k(),[e,""TK_WORD""])}if(""(""===e||""[""===e)return[e,""TK_START_EXPR""];if("")""===e||""]""===e)return[e,""TK_END_EXPR""];if(""{""===e)return[e,""TK_START_BLOCK""];if(""}""===e)return[e,""TK_END_BLOCK""];if("";""===e)return[e,""TK_SEMICOLON""];if(""/""===e){var c="""",R=!0;if(""*""==="
7    push o, "r.charAt(N)){if((N+=1)<m)for(;!(""*""===r.charAt(N)&&r.charAt(N+1)&&""/""===r.charAt(N+1))&&N<m&&(c+=e=r.charAt(N),(""\r""===e||""\n""===e)&&(R=!1),!((N+=1)>=m)););return(N+=2,R)?[""/*""+c+""*/"",""TK_INLINE_COMMENT""]:[""/*""+c+""*/"",""TK_BLOCK_COMMENT""]}if(""/""===r.charAt(N)){for(c=e;""\r""!==r.charAt(N)&&""\n""!==r.charAt(N)&&(c+=r.charAt(N),!((N+=1)>=m)););return N+=1,S&&k(),[c,""TK_COMMENT""]}}if(""'""===e||'""'===e||""/""===e&&(""TK_WORD""===i&&U(T,[""return"",""do""])||""TK_START_EXPR""===i||""TK_START_BLOCK""===i||""TK_END_BLOC"
8    push o, "K""===i||""TK_OPERATOR""===i||""TK_EQUALS""===i||""TK_EOF""===i||""TK_SEMICOLON""===i)){var _=e,d=!1,u=e;if(N<m){if(""/""===_){for(var $=!1;d||$||r.charAt(N)!==_;)if(u+=r.charAt(N),d?d=!1:(d=""\\""===r.charAt(N),""[""===r.charAt(N)?$=!0:""]""===r.charAt(N)&&($=!1)),(N+=1)>=m)return[u,""TK_STRING""]}else for(;d||r.charAt(N)!==_;)if(u+=r.charAt(N),d=!d&&""\\""===r.charAt(N),(N+=1)>=m)return[u,""TK_STRING""]}if(N+=1,u+=_,""/""===_)for(;N<m&&U(r.charAt(N),l);)u+=r.charAt(N),N+=1;return[u,""TK_STRING""]}if(""#""===e){var C=""#"";i"
9    push o, "f(N<m&&U(r.charAt(N),h)){do C+=e=r.charAt(N),N+=1;while(N<m&&""#""!==e&&""=""!==e);return""#""===e||(""[""===r.charAt(N)&&""]""===r.charAt(N+1)?(C+=""[]"",N+=2):""{""===r.charAt(N)&&""}""===r.charAt(N+1)&&(C+=""{}"",N+=2)),[C,""TK_WORD""]}}if(""<""===e&&""<!--""===r.substring(N-1,N+3))return N+=3,f.in_html_comment=!0,[""<!--"",""TK_COMMENT""];if(""-""===e&&f.in_html_comment&&""-->""===r.substring(N-1,N+2))return f.in_html_comment=!1,N+=2,S&&k(),[""-->"",""TK_COMMENT""];if(U(e,s)){for(;N<m&&U(e+r.charAt(N),s)&&(e+=r.charAt(N),!((N+"
10   push o, "=1)>=m)););return""=""===e?[e,""TK_EQUALS""]:[e,""TK_OPERATOR""]}return[e,""TK_UNKNOWN""]}for(O="""";C>0;)O+=I,C-=1;for(r=e,o="""",i=""TK_START_EXPR"",T="""",E="""",t=[],u=!1,K=""\n\r    "".split(""""),l=""abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_$"".split(""""),h=""0123456789"".split(""""),s=""+ - * / % & ++ -- = += -= *= /= %= == === != !== > < >= <= >> << >>> >>>= >>= <<= && &= | || ! !! , : ? ^ ^= |= ::"".split("" ""),R=""continue,try,throw,return,var,if,switch,case,default,for,while,break,function"".split("
11   push o, """,""),c=[],w(""BLOCK""),N=0;;){var Q=G(N);if(a=Q[0],""TK_EOF""===(d=Q[1]))break;switch(d){case""TK_START_EXPR"":if(""[""===a){if(""TK_WORD""===i||"")""===T){U(T,R)&&B(),w(""(EXPRESSION)""),M();break}""[EXPRESSION]""===f.mode||""[INDENTED-EXPRESSION]""===f.mode?""]""===E&&"",""===T?(""[EXPRESSION]""!==f.mode||(f.mode=""[INDENTED-EXPRESSION]"",D||W()),w(""[EXPRESSION]""),D||k()):""[""===T?(""[EXPRESSION]""!==f.mode||(f.mode=""[INDENTED-EXPRESSION]"",D||W()),w(""[EXPRESSION]""),D||k()):w(""[EXPRESSION]""):w(""[EXPRESSION]"")}else w(""(EXPR"
12   push o, "ESSION)"");"";""===T||""TK_START_BLOCK""===i?k():""TK_END_EXPR""===i||""TK_START_EXPR""===i||""TK_END_BLOCK""===i||"".""===T||(""TK_WORD""!==i&&""TK_OPERATOR""!==i?B():""function""===o?p&&B():(U(T,R)||""catch""===T)&&B()),M();break;case""TK_END_EXPR"":if(""]""===a){if(D){if(""}""===T){g(),M(),y();break}}else if(""[INDENTED-EXPRESSION]""===f.mode&&""]""===T){y(),k(),M();break}}y(),M();break;case""TK_START_BLOCK"":w(""do""===o?""DO_BLOCK"":""BLOCK""),$?(""TK_OPERATOR""!==i&&(""return""==T?B():k(!0)),M(),W()):(""TK_OPERATOR""!==i&&""TK_START_E"
13   push o, "XPR""!==i?""TK_START_BLOCK""===i?k():B():X(f.previous_mode)&&"",""===T&&k(),W(),M());break;case""TK_END_BLOCK"":y(),$?(k(),M()):(""TK_START_BLOCK""===i?A?g():b():k(),M());break;case""TK_WORD"":if(u){B(),M(),B(),u=!1;break}if(""function""===a&&(A||"";""===T)&&""{""!==T){v=A?v:0;for(var J=0;J<2-v;J++)k(!1)}if(""case""===a||""default""===a){"":""===T?g():(f.indentation_level--,k(),f.indentation_level++),M(),f.in_case=!0;break}_=""NONE"",""TK_END_BLOCK""===i?U(a.toLowerCase(),[""else"",""catch"",""finally""])?$?_=""NEWLINE"":(_=""SPAC"
14   push o, "E"",B()):_=""NEWLINE"":""TK_SEMICOLON""===i&&(""BLOCK""===f.mode||""DO_BLOCK""===f.mode)?_=""NEWLINE"":""TK_SEMICOLON""===i&&x(f.mode)?_=""SPACE"":""TK_STRING""===i?_=""NEWLINE"":""TK_WORD""===i?_=""SPACE"":""TK_START_BLOCK""===i?_=""NEWLINE"":""TK_END_EXPR""===i&&(B(),_=""NEWLINE""),""TK_END_BLOCK""!==i&&U(a.toLowerCase(),[""else"",""catch"",""finally""])?k():U(a,R)||""NEWLINE""===_?""else""===T?B():(""TK_START_EXPR""===i||""=""===T||"",""===T)&&""function""===a||(""return""===T||""throw""===T?B():""TK_END_EXPR""!==i?(""TK_START_EXPR""!==i||""var""!==a)&"
15   push o, "&"":""!==T&&(""if""===a&&""else""===o&&""{""!==T?B():k()):U(a,R)&&"")""!==T&&k()):X(f.mode)&&"",""===T&&""}""===E?k():""SPACE""===_&&B(),M(),o=a,""var""===a&&(f.var_line=!0,f.var_line_reindented=!1,f.var_line_tainted=!1),(""if""===a||""else""===a)&&(f.if_line=!0);break;case""TK_SEMICOLON"":M(),f.var_line=!1,f.var_line_reindented=!1;break;case""TK_STRING"":""TK_START_BLOCK""===i||""TK_END_BLOCK""===i||""TK_SEMICOLON""===i?k():""TK_WORD""===i&&B(),M();break;case""TK_EQUALS"":f.var_line&&(f.var_line_tainted=!0),B(),M(),B();break;case"
16   push o, """TK_OPERATOR"":var j=!0,z=!0;if(f.var_line&&"",""===a&&x(f.mode)&&(f.var_line_tainted=!1),f.var_line&&"",""===a){if(f.var_line_tainted){M(),f.var_line_reindented=!0,f.var_line_tainted=!1,k();break}f.var_line_tainted=!1}if(""return""===T||""throw""===T){B(),M();break}if("":""===a&&f.in_case){M(),k(),f.in_case=!1;break}if(""::""===a){M();break}if("",""===a){f.var_line?f.var_line_tainted?(M(),k(),f.var_line_tainted=!1):(M(),B()):""TK_END_BLOCK""===i&&""(EXPRESSION)""!==f.mode?(M(),""OBJECT""===f.mode&&""}""===T?k():B()):"
17   push o, """OBJECT""===f.mode?(M(),k()):(M(),B());break}U(a,[""--"",""++"",""!""])||U(a,[""-"",""+""])&&(U(i,[""TK_START_BLOCK"",""TK_START_EXPR"",""TK_EQUALS"",""TK_OPERATOR""])||U(T,R))?(j=!1,z=!1,"";""===T&&x(f.mode)&&(j=!0),""TK_WORD""===i&&U(T,R)&&(j=!0),""BLOCK""===f.mode&&(""{""===T||"";""===T)&&k()):"".""===a?j=!1:"":""!==a||(f.mode=""OBJECT"",F()||(j=!1)),j&&B(),M(),z&&B();break;case""TK_BLOCK_COMMENT"":var q=a.split(/\x0a|\x0d\x0a/);if(/^\/\*\*/.test(a))for(k(),t.push(q[0]),J=1;J<q.length;J++)k(),t.push("" ""),t.push(q[J].replace(/^\s"
18   push o, "\s*|\s\s*$/,""""));else for(q.length>1?(k(),b()):B(),J=0;J<q.length;J++)t.push(q[J]),t.push(""\n"");k();break;case""TK_INLINE_COMMENT"":B(),M(),x(f.mode)?B():k();break;case""TK_COMMENT"":S?k():B(),M(),k();break;case""TK_UNKNOWN"":M()}E=T,i=d,T=a}return t.join("""").replace(/[\n ]+$/,"""")}"

    getBeautifyScript = Join(o, "")
    
End Function





