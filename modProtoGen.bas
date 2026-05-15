Attribute VB_Name = "modProtoGen"
'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.7
'Site:         http://sandsprite.com
'License:      MIT

Option Explicit

' modProtoGen — parses a .vbp project file, walks every Class= and Form=
' entry, extracts the public surface (fields, methods, properties), and
' writes one proto file per class to ./protos/<ClassName>.txt
'
' Proto format mirrors VB6 syntax with bodies stripped, e.g.:
'
'   Class CUser
'     Public name As String
'     Public age As String
'     Public job As String
'     Public Function describeSelf() As String
'   End Class
'
' Forms emit the same way but with `Class Form1` (the runtime type name
' of a form is just its form name, and TypeName() returns "Form1").
'
' Usage from Form1:
'   modProtoGen.RegenerateProtos App.Path & "\Project1.vbp"

Private Const PROTOS_SUBDIR As String = "\protos\"

' --- Public entry point ---------------------------------------------------

' Returns the number of proto files written. Raises an error if the .vbp
' cannot be read or if the protos directory can't be created.
Public Function RegenerateProtos(ByVal vbpPath As String) As Long

    If Not FileExists(vbpPath) Then
        Err.Raise vbObjectError + 1, "modProtoGen", _
                  "vbp file not found: " & vbpPath
    End If

    Dim baseDir As String
    baseDir = ParentDir(vbpPath)

    Dim protosDir As String
    protosDir = baseDir & PROTOS_SUBDIR
    EnsureDir protosDir

    Dim entries() As String
    entries = ParseVbpEntries(vbpPath)

    Dim i As Long, written As Long
    Dim relPath As String, srcPath As String
    Dim isForm As Boolean
    Dim protoText As String, className As String

    For i = 0 To UBound(entries)
        ' entries(i) is like "Class=CFoo; CFoo.cls" or "Form=Form1; Form1.frm"
        isForm = (Left$(entries(i), 5) = "Form=")
        relPath = ExtractRelPath(entries(i))
        If LenB(relPath) > 0 Then
            srcPath = baseDir & "\" & relPath
            If FileExists(srcPath) Then
                className = ExtractClassName(srcPath, isForm)
                If LenB(className) > 0 Then
                    protoText = BuildProto(srcPath, className, isForm)
                    WriteTextFile protosDir & className & ".txt", protoText
                    written = written + 1
                End If
            End If
        End If
    Next i

    RegenerateProtos = written

End Function

' --- .vbp parsing ---------------------------------------------------------

Private Function ParseVbpEntries(ByVal vbpPath As String) As String()
    Dim content As String
    content = ReadTextFile(vbpPath)

    Dim lines() As String
    lines = Split(content, vbCrLf)

    Dim out() As String
    ReDim out(0 To 256)
    Dim n As Long, i As Long, ln As String

    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If Left$(ln, 6) = "Class=" Or Left$(ln, 5) = "Form=" Then
            If n > UBound(out) Then ReDim Preserve out(0 To n + 64)
            out(n) = ln
            n = n + 1
        End If
    Next i

    If n = 0 Then
        ReDim out(0 To -1)  ' empty
    Else
        ReDim Preserve out(0 To n - 1)
    End If

    ParseVbpEntries = out
End Function

' Given "Class=CFoo; CFoo.cls" or "Form=Form1; Form1.frm", return the
' file portion ("CFoo.cls" or "Form1.frm"). Class entries are
' "Class=Name; filename"; Form entries are "Form=filename" (no name).
Private Function ExtractRelPath(ByVal entry As String) As String
    Dim semi As Long
    semi = InStr(entry, ";")
    If semi > 0 Then
        ExtractRelPath = Trim$(Mid$(entry, semi + 1))
    Else
        ' Form= entries have just the filename
        Dim eq As Long
        eq = InStr(entry, "=")
        If eq > 0 Then ExtractRelPath = Trim$(Mid$(entry, eq + 1))
    End If
End Function

' --- Class name extraction ------------------------------------------------

' For .cls files: reads Attribute VB_Name = "ClassName"
' For .frm files: reads Begin VB.Form Form1 (form name follows the type)
Private Function ExtractClassName(ByVal srcPath As String, ByVal isForm As Boolean) As String
    Dim content As String
    content = ReadTextFile(srcPath)
    Dim lines() As String
    lines = Split(content, vbCrLf)

    Dim i As Long, ln As String
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If isForm Then
            ' "Begin VB.Form Form1"
            If LCase$(Left$(ln, 9)) = "begin vb." Then
                Dim toks() As String
                toks = Split(ln, " ")
                If UBound(toks) >= 2 Then
                    ExtractClassName = toks(2)
                    Exit Function
                End If
            End If
        Else
            ' Attribute VB_Name = "CFoo"
            If InStr(ln, "Attribute VB_Name") > 0 Then
                Dim q1 As Long, q2 As Long
                q1 = InStr(ln, """")
                If q1 > 0 Then
                    q2 = InStr(q1 + 1, ln, """")
                    If q2 > q1 Then
                        ExtractClassName = Mid$(ln, q1 + 1, q2 - q1 - 1)
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

' --- Proto building -------------------------------------------------------

' Walks the source file line-by-line, extracting Public field declarations
' and Public Sub/Function/Property signatures. Skips bodies entirely by
' tracking nesting. The output mirrors VB6 syntax with bodies stripped.
Private Function BuildProto(ByVal srcPath As String, ByVal className As String, ByVal isForm As Boolean) As String

    Dim content As String
    content = ReadTextFile(srcPath)
    Dim lines() As String
    lines = Split(content, vbCrLf)

    Dim sb As String
    sb = "Class " & className & vbCrLf

    Dim i As Long, ln As String, trimmed As String
    Dim inBody As Boolean
    Dim pastHeader As Boolean

    ' For .frm files, skip the visual designer block (everything until the
    ' line "Attribute VB_Name" or, failing that, the first Option/Public/etc).
    ' For .cls files, skip until past Attribute VB_Exposed.
    For i = 0 To UBound(lines)
        ln = lines(i)
        trimmed = Trim$(ln)

        If Not pastHeader Then
            If InStr(trimmed, "Attribute VB_Exposed") > 0 Then
                pastHeader = True
            ElseIf isForm And InStr(trimmed, "Attribute VB_Name") > 0 Then
                pastHeader = True
            End If
            GoTo NextLine
        End If

        ' Inside a Sub/Function/Property body? Skip until End Sub/Function/Property.
        If inBody Then
            If LCase$(Left$(trimmed, 8)) = "end sub" Or _
               LCase$(Left$(trimmed, 12)) = "end function" Or _
               LCase$(Left$(trimmed, 12)) = "end property" Then
                inBody = False
            End If
            GoTo NextLine
        End If

        ' Skip empty lines, comments, Option/Attribute/etc
        If LenB(trimmed) = 0 Then GoTo NextLine
        If Left$(trimmed, 1) = "'" Then GoTo NextLine
        If LCase$(Left$(trimmed, 7)) = "option " Then GoTo NextLine
        If LCase$(Left$(trimmed, 10)) = "attribute " Then GoTo NextLine
        If LCase$(Left$(trimmed, 8)) = "private " Then
            ' Could be Private Sub/Function/Property — still need to skip its body
            ' (unless one-liner)
            If LookLikeRoutine(trimmed) Then
                If Not LineContainsEnd(trimmed) Then inBody = True
            End If
            GoTo NextLine
        End If
        If LCase$(Left$(trimmed, 7)) = "friend " Then
            If LookLikeRoutine(trimmed) Then
                If Not LineContainsEnd(trimmed) Then inBody = True
            End If
            GoTo NextLine
        End If

        ' Public field?  e.g. "Public name As String" or "Public Users As New Collection"
        If LCase$(Left$(trimmed, 7)) = "public " Then
            If LookLikeRoutine(trimmed) Then
                ' Public Sub/Function/Property — record signature, then skip body
                ' UNLESS the line also contains its own End — then it's a
                ' one-liner and we don't enter body mode.
                sb = sb & "  Public " & SignatureLine(trimmed) & vbCrLf
                If Not LineContainsEnd(trimmed) Then inBody = True
            Else
                ' Public field
                sb = sb & "  " & FieldLine(trimmed) & vbCrLf
            End If
            GoTo NextLine
        End If

        ' Sub/Function/Property without explicit Public — VB6 defaults to Public
        If LookLikeRoutine(trimmed) Then
            sb = sb & "  Public " & SignatureLine(trimmed) & vbCrLf
            If Not LineContainsEnd(trimmed) Then inBody = True
            GoTo NextLine
        End If

NextLine:
    Next i

    sb = sb & "End Class" & vbCrLf
    BuildProto = sb
End Function

' Does this line start a Sub/Function/Property definition? Doesn't matter
' whether prefixed with Public/Private/Friend — caller has already handled that.
Private Function LookLikeRoutine(ByVal trimmed As String) As Boolean
    Dim lc As String
    lc = LCase$(trimmed)
    ' Strip leading Public/Private/Friend so we can test the keyword that follows
    If Left$(lc, 7) = "public " Then lc = Mid$(lc, 8)
    If Left$(lc, 8) = "private " Then lc = Mid$(lc, 9)
    If Left$(lc, 7) = "friend " Then lc = Mid$(lc, 8)
    If Left$(lc, 4) = "sub " Then LookLikeRoutine = True: Exit Function
    If Left$(lc, 9) = "function " Then LookLikeRoutine = True: Exit Function
    If Left$(lc, 9) = "property " Then LookLikeRoutine = True: Exit Function
End Function

' Does the trimmed line contain an "End Sub", "End Function", or "End Property"?
' Used to detect one-line routine definitions like
'   Public Property Get X() As Long: X = m_x: End Property
Private Function LineContainsEnd(ByVal trimmed As String) As Boolean
    Dim lc As String
    lc = LCase$(trimmed)
    If InStr(lc, "end sub") > 0 Then LineContainsEnd = True: Exit Function
    If InStr(lc, "end function") > 0 Then LineContainsEnd = True: Exit Function
    If InStr(lc, "end property") > 0 Then LineContainsEnd = True: Exit Function
End Function

' Returns the signature portion of a routine declaration, with any leading
' Public/Private/Friend stripped (the caller decides what prefix to write).
' If the line is a one-liner with ": body : End X" tail, trim that off so
' only the signature remains.
' Strips trailing line-continuation underscores.
Private Function SignatureLine(ByVal trimmed As String) As String
    Dim s As String
    s = trimmed
    If LCase$(Left$(s, 7)) = "public " Then s = Mid$(s, 8)
    If LCase$(Left$(s, 8)) = "private " Then s = Mid$(s, 9)
    If LCase$(Left$(s, 7)) = "friend " Then s = Mid$(s, 8)

    ' One-liner: cut at the first ':' that comes AFTER the closing paren of
    ' the parameter list (or after the As-clause if there's a return type).
    ' Quick heuristic: cut at first ':' that is not inside parens and not
    ' inside a string literal.
    Dim cut As Long
    cut = FirstColonOutsideParensAndStrings(s)
    If cut > 0 Then s = Trim$(Left$(s, cut - 1))

    ' If the signature continues with " _" on the next line we just take this
    ' line — good enough for 95% of cases.
    If Right$(s, 2) = " _" Then s = Trim$(Left$(s, Len(s) - 2)) & " ..."
    SignatureLine = s
End Function

Private Function FirstColonOutsideParensAndStrings(ByVal s As String) As Long
    Dim i As Long, depth As Long, inStr_ As Boolean, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then inStr_ = Not inStr_
        If Not inStr_ Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then depth = depth - 1
            If ch = ":" And depth = 0 Then
                FirstColonOutsideParensAndStrings = i
                Exit Function
            End If
        End If
    Next i
End Function

' For Public field declarations: returns the line as-is (already starts with "Public").
Private Function FieldLine(ByVal trimmed As String) As String
    ' Strip trailing comment if any
    Dim apos As Long
    apos = InStrApostropheOutsideString(trimmed)
    If apos > 0 Then
        FieldLine = Trim$(Left$(trimmed, apos - 1))
    Else
        FieldLine = trimmed
    End If
End Function

' Find an apostrophe that isn't inside a string literal. Simple state machine.
Private Function InStrApostropheOutsideString(ByVal s As String) As Long
    Dim i As Long, inStr_ As Boolean, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then inStr_ = Not inStr_
        If ch = "'" And Not inStr_ Then
            InStrApostropheOutsideString = i
            Exit Function
        End If
    Next i
End Function

' --- File helpers ---------------------------------------------------------

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(p)) > 0)
End Function

Private Function ParentDir(ByVal p As String) As String
    Dim i As Long
    For i = Len(p) To 1 Step -1
        If Mid$(p, i, 1) = "\" Then
            ParentDir = Left$(p, i - 1)
            Exit Function
        End If
    Next i
End Function

Private Sub EnsureDir(ByVal d As String)
    If Len(Dir$(d, vbDirectory)) = 0 Then MkDir d
End Sub

Private Function ReadTextFile(ByVal p As String) As String
    Dim fnum As Integer
    fnum = FreeFile
    Open p For Binary Access Read As #fnum
    Dim buf As String
    buf = Space$(LOF(fnum))
    Get #fnum, , buf
    Close #fnum
    ReadTextFile = buf
End Function

Private Sub WriteTextFile(ByVal p As String, ByVal content As String)
    Dim fnum As Integer
    fnum = FreeFile
    Open p For Output As #fnum
    Print #fnum, content;
    Close #fnum
End Sub
