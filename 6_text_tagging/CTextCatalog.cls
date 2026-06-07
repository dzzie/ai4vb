VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.8
'Site:         http://sandsprite.com
'License:      MIT
'
' The text sibling of CImageCatalog. Same two-phase, resumable design --
' only the payload changes: instead of a flat tag list per image, each
' text file is run through a SCHEMA-CONSTRAINED extraction prompt and the
' model returns a JSON object that we parse, VALIDATE, and store as columns.
'
'   Scan()            - mechanical: walk a folder, stat + MD5 each text file,
'                       upsert with classified=0. Fast, no AI. Idempotent:
'                       unchanged files are skipped; changed files are
'                       re-hashed and flagged for re-extraction.
'   ClassifyExtract() - the AI step: read each pending file, ask the model
'                       for the structured object, parse with CJSON, coerce
'                       any out-of-range enum field to a fallback, store the
'                       columns AND the raw JSON. RESUMABLE -- only touches
'                       rows where classified=0.
'
' classified flag values:
'   0 = pending, 1 = extracted OK, -1 = 200 received but JSON unparseable
'       (terminal -- needs your eyes; raw_json is stored so you can look).
'   A non-200 (transient HTTP/server) leaves the row at 0 so it retries.
'
' valid flag: 1 if every enum field came back inside its allowed set;
'             0 if any field had to be coerced to a fallback. This column
'             IS your eval metric -- COUNT(valid=0) is the schema-conformance
'             failure rate for a given model/prompt, the text analog of the
'             image A/B "did it follow the format" question.
'
' WHY store raw_json: same principle as keeping the transcript, the md5, the
' raw tags -- materialize the intermediate so you can re-parse or re-validate
' later WITHOUT re-calling the model, and so a bad field is auditable rather
' than a black box. Re-run validation by reading raw_json; never re-pay the
' inference cost to change a coercion rule.
'
' Dependencies: cSQLite (+ cSQLiteSchema, cSQLiteResults), CFileSystem2,
'               CJSON, modHash (FileMD5 -- shared with the image example).

Option Explicit

Public Event Progress(ByVal phase As String, ByVal done As Long, ByVal total As Long, ByVal currentPath As String)
Public Event LogMsg(ByVal msg As String)

' Set from the UI (e.g. a Cancel button) to stop a running phase. Checked
' between files; the in-flight AI call also gets ai.Cancel.
Public CancelRequested As Boolean

' Cap the text sent per file so one giant log doesn't blow the prompt out.
' qwen3 has plenty of context; this is about speed/cost, not a hard limit.
' Bump it if your documents are long and you need the whole thing.
Public MaxChars As Long

Private m_db As cSQLite
Private m_fso As CFileSystem2

' --- the schema, as allowed-value lists. Edit these to match YOUR domain. --
' This demo models generic "document triage". The two layers that make an
' extraction trustworthy are both here: a CONSTRAINED set (enum fields the
' model must pick from) and FREE TEXT (summary/keywords). Validation only
' polices the constrained fields -- that's where models drift.
Private Const ALLOWED_CATEGORY As String = "billing|technical|account|sales|feedback|legal|other"
Private Const ALLOWED_PRIORITY As String = "low|medium|high"
Private Const ALLOWED_SENTIMENT As String = "positive|neutral|negative"

Private Const FALLBACK_CATEGORY As String = "other"
Private Const FALLBACK_PRIORITY As String = "unknown"
Private Const FALLBACK_SENTIMENT As String = "unknown"

' The extraction prompt. Note the two load-bearing clauses, both lessons from
' this session: (1) "ONLY a JSON object ... no markdown" forces a parseable
' shape; (2) "Do not invent facts ... base every field only on what the
' document says" is the same anti-fabrication rule that stopped the chatroom
' agent making up statistics -- here it stops the model inventing a category
' or sentiment the text doesn't support.
Private Const EXTRACT_PROMPT As String = _
    "You are a document triage engine. Read the document and return ONLY a JSON " & _
    "object -- no markdown fences, no commentary, no preamble -- with EXACTLY these fields:" & vbLf & _
    "{""category"": one of [billing, technical, account, sales, feedback, legal, other]," & vbLf & _
    " ""priority"": one of [low, medium, high]," & vbLf & _
    " ""sentiment"": one of [positive, neutral, negative]," & vbLf & _
    " ""summary"": a single sentence under 20 words," & vbLf & _
    " ""keywords"": 3 to 8 comma-separated lowercase keywords}" & vbLf & _
    "Choose the single best category; if nothing fits, use ""other"". " & _
    "Do not invent facts not present in the text. Base every field only on what the document actually says. " & _
    "Output the JSON object and nothing else."

Private Sub Class_Initialize()
    Set m_fso = New CFileSystem2
    MaxChars = 8000
End Sub

' Open (creating if needed) the catalog DB and ensure the schema exists.
Public Sub OpenCatalog(ByVal dbPath As String)
    Set m_db = New cSQLite
    m_db.OpenDB dbPath, False, True

    Dim sch As cSQLiteSchema
    Set sch = New cSQLiteSchema
    sch.AttachDb m_db

    If Not sch.TableExists("docs") Then
        sch.CreateTable "docs", _
            "id INTEGER PRIMARY KEY, " & _
            "path TEXT UNIQUE NOT NULL, " & _
            "name TEXT, " & _
            "ext TEXT, " & _
            "size INTEGER, " & _
            "mtime TEXT, " & _
            "md5 TEXT, " & _
            "category TEXT, " & _
            "priority TEXT, " & _
            "sentiment TEXT, " & _
            "summary TEXT, " & _
            "keywords TEXT, " & _
            "raw_json TEXT, " & _
            "valid INTEGER, " & _
            "classified INTEGER DEFAULT 0, " & _
            "model TEXT, " & _
            "tagged_at TEXT"
        sch.CreateIndex "ix_docs_classified", "docs", "classified"
        sch.CreateIndex "ix_docs_category", "docs", "category"
        sch.CreateIndex "ix_docs_md5", "docs", "md5"
    End If
End Sub

Private Function IsTextExt(ByVal ext As String) As Boolean
    Select Case LCase$(ext)
        Case ".txt", ".md", ".log", ".csv", ".json", ".xml", ".eml", ".html", ".htm", ".ini", ".cfg"
            IsTextExt = True
    End Select
End Function

' Walk folderPath (optionally recursive), upserting text-file stats. Returns
' the number of rows inserted or updated (unchanged files are skipped).
Public Function Scan(ByVal folderPath As String, Optional ByVal recursive As Boolean = True) As Long
    CancelRequested = False

    Dim files() As String
    files = m_fso.GetFolderFiles(folderPath, "*", True, recursive)
    If Not GotAny(files) Then Exit Function

    Dim total As Long, i As Long, changed As Long
    total = UBound(files) + 1

    For i = 0 To UBound(files)
        If CancelRequested Then RaiseEvent LogMsg("Scan cancelled."): Exit For

        Dim p As String, ext As String
        p = files(i)
        ext = m_fso.GetExtension(p)

        RaiseEvent Progress("scan", i + 1, total, p)

        If IsTextExt(ext) Then
            Dim sz As Long, mt As String
            sz = fileLen(p)
            mt = Format$(FileDateTime(p), "yyyy-mm-dd hh:nn:ss")

            ' existence + change check (parameterized -- paths may contain quotes)
            Dim rs As cSQLiteResults, existId As Variant, existSig As Variant, newSig As String
            Set rs = m_db.Query("SELECT id, size || '|' || mtime FROM docs WHERE path = ?", p)
            newSig = CStr(sz) & "|" & mt

            If rs.MoveNext Then
                existId = rs.FieldInt(0)
                existSig = rs.FieldText(1)
                rs.CloseRS
                If CStr(existSig) <> newSig Then
                    ' content changed -> re-hash, wipe extracted fields, re-queue
                    m_db.ExecUpdate "docs", "where id = " & CLng(existId), _
                        "size,mtime,md5,category,priority,sentiment,summary,keywords,raw_json,valid,classified,model,tagged_at", _
                        sz, mt, FileMD5(p), Null, Null, Null, Null, Null, Null, Null, 0, Null, Null
                    changed = changed + 1
                End If
            Else
                rs.CloseRS
                m_db.ExecInsert "docs", "path,name,ext,size,mtime,md5,classified", _
                    p, m_fso.FileNameFromPath(p), LCase$(ext), sz, mt, FileMD5(p), 0
                changed = changed + 1
            End If
        End If
    Next

    RaiseEvent LogMsg("Scan complete. " & changed & " new/changed of " & total & " files.")
    Scan = changed
End Function

' Extract structured fields for every pending doc (classified=0) with the
' supplied text model. Resumable. Returns count extracted OK this run.
' ai must expose CreateResponse / ExtractOutput / LastStatus / LastError /
' Cancel / Model -- any of your provider classes set to a TEXT model
' (qwen3 locally, or a Gemini/Claude/OpenAI model).
Public Function ClassifyExtract(ByRef ai As Object, Optional ByVal maxFiles As Long = 0) As Long
    CancelRequested = False

    ' Collect the work list FIRST, then close the cursor, so we never write to
    ' the table while a SELECT is still open on it (the SQLite footgun).
    Dim ids As New Collection, paths As New Collection
    Dim lim As String
    If maxFiles > 0 Then lim = " LIMIT " & maxFiles

    Dim rs As cSQLiteResults
    Set rs = m_db.Query("SELECT id, path FROM docs WHERE classified = 0 ORDER BY id" & lim)
    Do While rs.MoveNext
        ids.Add rs.FieldInt("id")
        paths.Add rs.FieldText("path")
    Loop
    rs.CloseRS

    Dim total As Long, i As Long, ok As Long
    total = ids.Count
    If total = 0 Then RaiseEvent LogMsg("Nothing pending."): Exit Function

    Dim modelName As String
    On Error Resume Next: modelName = CStr(ai.Model): On Error GoTo 0

    For i = 1 To total
        If CancelRequested Then
            ai.Cancel
            RaiseEvent LogMsg("Extract cancelled at " & (i - 1) & "/" & total & ".")
            Exit For
        End If

        Dim id As Long, p As String
        id = ids.Item(i)
        p = paths.Item(i)

        RaiseEvent Progress("extract", i, total, p)

        If Not m_fso.FileExists(p) Then
            m_db.ExecUpdate "docs", "where id = " & id, "classified", -1
            RaiseEvent LogMsg("missing: " & p)
            GoTo NextDoc
        End If

        ' Read the file and bound its size. The document goes in the user turn;
        ' the schema/rules go in the system prompt -- keeps the instruction
        ' stable and separate from the (untrusted) document content.
        Dim body As String
        body = CStr(m_fso.ReadFile(p))
        If MaxChars > 0 And Len(body) > MaxChars Then body = Left$(body, MaxChars)

        ai.CreateResponse "DOCUMENT:" & vbLf & body, EXTRACT_PROMPT, False, True

        If ai.LastStatus <> 200 Then
            ' transient -> leave at 0 so a later run retries this one
            RaiseEvent LogMsg("AI fail id=" & id & " status=" & ai.LastStatus & " - " & ai.LastError)
            GoTo NextDoc
        End If

        Dim raw As String
        raw = ExtractJsonObject(CStr(ai.ExtractOutput))
        If LenB(raw) = 0 Then
            ' got a 200 but no JSON object in it -> terminal, store nothing parseable
            m_db.ExecUpdate "docs", "where id = " & id, "raw_json,classified", CStr(ai.ExtractOutput), -1
            RaiseEvent LogMsg("no JSON id=" & id & " (marked -1 for review)")
            GoTo NextDoc
        End If

        Dim j As New CJSON
        If Not j.Initilized Then RaiseEvent LogMsg("CJSON not initialized -- aborting."): Exit For
        If Not j.loadJson(raw) Then
            m_db.ExecUpdate "docs", "where id = " & id, "raw_json,classified", raw, -1
            RaiseEvent LogMsg("bad JSON id=" & id & " (" & j.jsErr & ") -- marked -1")
            GoTo NextDoc
        End If

        ' Pull fields, then VALIDATE the constrained ones. validFlag tracks
        ' whether anything had to be coerced -- that's the conformance metric.
        Dim cat As String, pri As String, sen As String, summ As String, kw As String
        Dim validFlag As Long
        validFlag = 1

        cat = Coerce(LCase$(Trim$(CStr(j.getVal("category")))), ALLOWED_CATEGORY, FALLBACK_CATEGORY, validFlag)
        pri = Coerce(LCase$(Trim$(CStr(j.getVal("priority")))), ALLOWED_PRIORITY, FALLBACK_PRIORITY, validFlag)
        sen = Coerce(LCase$(Trim$(CStr(j.getVal("sentiment")))), ALLOWED_SENTIMENT, FALLBACK_SENTIMENT, validFlag)
        summ = Trim$(CStr(j.getVal("summary")))
        kw = NormalizeKeywords(CStr(j.getVal("keywords")))

        m_db.ExecUpdate "docs", "where id = " & id, _
            "category,priority,sentiment,summary,keywords,raw_json,valid,classified,model,tagged_at", _
            cat, pri, sen, summ, kw, raw, validFlag, 1, modelName, Format$(Now, "yyyy-mm-dd hh:nn:ss")
        ok = ok + 1
        If validFlag = 0 Then RaiseEvent LogMsg("coerced id=" & id & " (enum field out of range)")

NextDoc:
    Next

    RaiseEvent LogMsg("Extract complete. " & ok & " extracted this run.")
    ClassifyExtract = ok
End Function

' Find the outermost { ... } in a model reply, stripping markdown fences and
' any chatty preamble in one shot. Models love to wrap JSON in ```json ... ```
' or add "Here is the object:" -- this is the structured-output analog of the
' chatroom's StripSpeakerPrefix: the harness sanitizes, it never trusts format.
Private Function ExtractJsonObject(ByVal s As String) As String
    Dim a As Long, b As Long
    a = InStr(s, "{")
    b = InStrRev(s, "}")
    If a > 0 And b > a Then ExtractJsonObject = Mid$(s, a, b - a + 1)
End Function

' Return val if it's in the pipe-delimited allowed list; otherwise return
' fallback and clear the caller's validFlag (it was coerced).
Private Function Coerce(ByVal val As String, ByVal allowed As String, ByVal fallback As String, ByRef validFlag As Long) As String
    If InList(val, allowed) Then
        Coerce = val
    Else
        Coerce = fallback
        validFlag = 0
    End If
End Function

Private Function InList(ByVal val As String, ByVal pipeList As String) As Boolean
    InList = (InStr("|" & pipeList & "|", "|" & val & "|") > 0) And LenB(val) <> 0
End Function

' Clean the free-text keyword field into a space-wrapped, lowercase, deduped
' string so LIKE '% term %' can match a whole keyword (same storage trick as
' the image catalog's tags). Drops empties and anything sentence-like.
Private Function NormalizeKeywords(ByVal rawKw As String) As String
    Dim s As String
    s = rawKw
    s = Replace$(s, vbCr, ",")
    s = Replace$(s, vbLf, ",")
    s = Replace$(s, ";", ",")

    Dim parts() As String, i As Long, t As String, seen As String, outp As String
    parts = Split(s, ",")
    seen = " "
    For i = 0 To UBound(parts)
        t = LCase$(Trim$(parts(i)))
        If LenB(t) <> 0 And Len(t) <= 30 And InStr(t, " ") <= 1 Then
            If InStr(seen, " " & t & " ") = 0 Then
                outp = outp & t & " "
                seen = seen & t & " "
            End If
        End If
    Next

    If LenB(outp) <> 0 Then NormalizeKeywords = " " & Trim$(outp) & " "
End Function

' Search extracted docs. Optional category filter (exact) plus AND-ed keyword
' terms matched against the keywords field. Caller iterates and CloseRS.
Public Function Search(Optional ByVal terms As String = "", Optional ByVal category As String = "") As cSQLiteResults
    Dim words() As String, i As Long
    Dim sql As String, params As New Collection

    sql = "SELECT path, category, priority, sentiment, summary, keywords FROM docs WHERE classified = 1"

    If LenB(Trim$(category)) <> 0 Then
        sql = sql & " AND category = ?"
        params.Add LCase$(Trim$(category))
    End If

    words = Split(Trim$(LCase$(terms)))
    For i = 0 To UBound(words)
        If LenB(words(i)) <> 0 Then
            sql = sql & " AND keywords LIKE ?"
            params.Add "%" & words(i) & "%"
        End If
    Next
    sql = sql & " ORDER BY category, path"

    Select Case params.Count
        Case 0: Set Search = m_db.Query(sql)
        Case 1: Set Search = m_db.Query(sql, params(1))
        Case 2: Set Search = m_db.Query(sql, params(1), params(2))
        Case 3: Set Search = m_db.Query(sql, params(1), params(2), params(3))
        Case Else: Set Search = m_db.Query(sql, params(1), params(2), params(3), params(4))
    End Select
End Function

' Breakdown of extracted docs by category -- handy for a quick dashboard.
Public Function CategoryCounts() As cSQLiteResults
    Set CategoryCounts = m_db.Query( _
        "SELECT category, COUNT(*) AS n FROM docs WHERE classified = 1 GROUP BY category ORDER BY n DESC")
End Function

Public Property Get PendingCount() As Long
    PendingCount = CLng(val(NzStr(m_db.Scalar("SELECT COUNT(*) FROM docs WHERE classified = 0"))))
End Property

Public Property Get DoneCount() As Long
    DoneCount = CLng(val(NzStr(m_db.Scalar("SELECT COUNT(*) FROM docs WHERE classified = 1"))))
End Property

' Rows where an enum field had to be coerced -- the schema-conformance miss
' rate for whatever model/prompt produced them. Your headline eval number.
Public Property Get CoercedCount() As Long
    CoercedCount = CLng(val(NzStr(m_db.Scalar("SELECT COUNT(*) FROM docs WHERE classified = 1 AND valid = 0"))))
End Property

' Rows that returned a 200 but produced no usable JSON -- parse-failure rate.
Public Property Get UnparseableCount() As Long
    UnparseableCount = CLng(val(NzStr(m_db.Scalar("SELECT COUNT(*) FROM docs WHERE classified = -1"))))
End Property

Public Property Get TotalCount() As Long
    TotalCount = CLng(val(NzStr(m_db.Scalar("SELECT COUNT(*) FROM docs"))))
End Property

Public Property Get DB() As cSQLite
    Set DB = m_db
End Property

Private Function NzStr(ByVal v As Variant) As String
    If IsNull(v) Then NzStr = "0" Else NzStr = CStr(v)
End Function

Private Function GotAny(ary() As String) As Boolean
    On Error GoTo no
    Dim x As Long: x = UBound(ary)
    GotAny = True
    Exit Function
no:
End Function


