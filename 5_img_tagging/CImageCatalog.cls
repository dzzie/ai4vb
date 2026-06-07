VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImageCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.8
'License:      MIT
'
' SQLite-backed image catalog with two decoupled phases:
'   Scan()           - mechanical: walk a folder, stat + MD5 each image, upsert
'                      into the DB with classified=0. Fast, no AI. Idempotent:
'                      unchanged files are skipped (not re-hashed); changed files
'                      are re-hashed and flagged for re-tagging.
'   ClassifyPending() - the AI step: feed each untagged image to a vision model,
'                      parse tags, store them. RESUMABLE -- only touches rows
'                      where classified=0, so a crash or cancel just resumes.
'                      Provider-agnostic: pass any class exposing
'                      CreateResponseWithImage / ExtractOutput / LastStatus.
'
' Search() does tag-based lookup. classified flag values:
'   0 = pending, 1 = tagged, -1 = file missing/unreadable at classify time.
'
' Dependencies: cSQLite (+ cSQLiteSchema, cSQLiteResults), CFileSystem2, modHash.

Option Explicit

Public Event Progress(ByVal phase As String, ByVal done As Long, ByVal total As Long, ByVal currentPath As String)
Public Event LogMsg(ByVal msg As String)

' Set from the UI (e.g. a Cancel button handler) to stop a running phase.
' Checked between files; the in-flight AI call also gets ai.Cancel.
Public CancelRequested As Boolean

Private m_db As cSQLite
Private m_fso As CFileSystem2
Private m_dbRs As cSQLiteResults   ' module-level so CloseRS is reachable in all paths

Private Const default_CLASSIFY_PROMPT As String = _
    "You are an image tagging engine. Look at the image and output ONLY a " & _
    "comma-separated list of lowercase tags describing what is actually visible: " & _
    "main subjects and objects, the scene or setting, and notable attributes " & _
    "(dominant colors, indoor or outdoor, time of day, overall style). Give 8 to 15 tags. " & _
    "Rules: output tags only -- no sentences, no explanation, no numbering, no preamble. " & _
    "Tag only what you can clearly see; do NOT guess at names, places, or text you cannot read. " & _
    "If it is a screenshot, document, diagram, or chart, tag it as such."

Dim CLASSIFY_PROMPT As String

Private Sub Class_Initialize()
    Set m_fso = New CFileSystem2
    
    Dim p As String
    p = App.path & "\prompt.txt"
    
    If m_fso.FileExists(p) Then
        CLASSIFY_PROMPT = m_fso.ReadFile(p)
    Else
        CLASSIFY_PROMPT = default_CLASSIFY_PROMPT
    End If
    
End Sub

' Open (creating if needed) the catalog database and ensure the schema exists.
Public Sub OpenCatalog(ByVal dbPath As String)
    Set m_db = New cSQLite
    m_db.OpenDB dbPath, False, True

    Dim sch As cSQLiteSchema
    Set sch = New cSQLiteSchema
    sch.AttachDb m_db

    If Not sch.TableExists("images") Then
        sch.CreateTable "images", _
            "id INTEGER PRIMARY KEY, " & _
            "path TEXT UNIQUE NOT NULL, " & _
            "name TEXT, " & _
            "ext TEXT, " & _
            "size INTEGER, " & _
            "mtime TEXT, " & _
            "md5 TEXT, " & _
            "tags TEXT, " & _
            "classified INTEGER DEFAULT 0, " & _
            "model TEXT, " & _
            "tagged_at TEXT"
        sch.CreateIndex "ix_images_classified", "images", "classified"
        sch.CreateIndex "ix_images_md5", "images", "md5"
    End If
End Sub

Private Function IsImageExt(ByVal ext As String) As Boolean
    Select Case LCase$(ext)
        Case ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff"
            IsImageExt = True
    End Select
End Function

' Walk folderPath (optionally recursive), upserting image file stats. Returns
' the number of rows inserted or updated (unchanged files are skipped).
Public Function Scan(ByVal folderPath As String, Optional ByVal recursive As Boolean = True) As Long
    CancelRequested = False

    Dim files() As String
    Dim existId As Variant, existSig As Variant, newSig As String, md5 As String
    
    files = m_fso.GetFolderFiles(folderPath, "*", True, recursive)
    If Not GotAny(files) Then Exit Function

    Dim total As Long, i As Long, changed As Long
    total = UBound(files) + 1

    For i = 0 To UBound(files)
        If CancelRequested Then
            RaiseEvent LogMsg("Scan cancelled.")
            Exit For
        End If

        Dim p As String, ext As String
        p = files(i)
        ext = m_fso.GetExtension(p)

        RaiseEvent Progress("scan", i + 1, total, p)

        If IsImageExt(ext) Then
            Dim sz As Long, mt As String
            sz = fileLen(p)
            mt = Format$(FileDateTime(p), "yyyy-mm-dd hh:nn:ss")
            md5 = FileMD5(p)
            ' existence + change check (parameterized -- paths may contain quotes)
            
            Set m_dbRs = m_db.Query("SELECT id, size || '|' || mtime FROM images WHERE path = ?", p)
            newSig = CStr(sz) & "|" & mt

            If m_dbRs.MoveNext Then
                existId = m_dbRs.FieldInt(0)
                existSig = m_dbRs.FieldText(1)
                m_dbRs.CloseRS
                If CStr(existSig) <> newSig Then
                    ' content changed -> re-hash, reset tags so it gets re-classified
                    m_db.ExecUpdate "images", "id = " & CLng(existId), _
                        "size,mtime,md5,tags,classified,model,tagged_at", _
                        sz, mt, md5, Null, 0, Null, Null
                    changed = changed + 1
                End If
            Else
                m_dbRs.CloseRS
                m_db.ExecInsert "images", "path,name,ext,size,mtime,md5,classified", _
                    p, m_fso.FileNameFromPath(p), LCase$(ext), sz, mt, md5, 0
                changed = changed + 1
            End If
        End If
    Next

    RaiseEvent LogMsg("Scan complete. " & changed & " new/changed of " & total & " files.")
    Scan = changed
End Function

' Tag every pending image (classified=0) with the supplied vision model.
' Resumable: re-running continues where it stopped. Returns count tagged.
' ai must expose CreateResponseWithImage / ExtractOutput / LastStatus /
' LastError / Cancel / Model -- i.e. any of your provider classes set to a
' VISION-capable model (qwen2.5-vl locally, or a Gemini/Claude model).
Public Function ClassifyPending(ByRef ai As Object, Optional ByVal maxFiles As Long = 0) As Long
    CancelRequested = False

    ' Collect the work list FIRST, then close the cursor, so we never write to
    ' the table while a SELECT statement is still open on it.
    Dim ids As New Collection, paths As New Collection
    Dim lim As String
    If maxFiles > 0 Then lim = " LIMIT " & maxFiles

    Dim rs As cSQLiteResults
    Set rs = m_db.Query("SELECT id, path FROM images WHERE classified = 0 ORDER BY id" & lim)
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
            RaiseEvent LogMsg("Classify cancelled at " & (i - 1) & "/" & total & ".")
            Exit For
        End If

        Dim id As Long, p As String
        id = ids.Item(i)
        p = paths.Item(i)

        RaiseEvent Progress("classify", i, total, p)

        If Not m_fso.FileExists(p) Then
            m_db.ExecUpdate "images", "id = " & id, "classified", -1
            RaiseEvent LogMsg("missing: " & p)
        Else
            ai.CreateResponseWithImage CLASSIFY_PROMPT, p, "", False, True

            If ai.LastStatus = 200 Then
                Dim tags As String
                tags = NormalizeTags(CStr(ai.ExtractOutput))
                If LenB(tags) <> 0 Then
                    m_db.ExecUpdate "images", "where id = " & id, _
                        "tags,classified,model,tagged_at", _
                        tags, 1, modelName, Format$(Now, "yyyy-mm-dd hh:nn:ss")
                    ok = ok + 1
                Else
                    RaiseEvent LogMsg("empty tags id=" & id & " (left pending)")
                End If
            Else
                ' leave classified=0 so a later run retries this one
                RaiseEvent LogMsg("AI fail id=" & id & " status=" & ai.LastStatus & " - " & ai.LastError)
            End If
        End If
    Next

    RaiseEvent LogMsg("Classify complete. " & ok & " tagged this run.")
    ClassifyPending = ok
End Function

' Turn raw model output into a clean, space-delimited, deduped tag string.
' Stored space-wrapped (" tag1 tag2 ") so LIKE '% term %' can match whole tags.
Private Function NormalizeTags(ByVal raw As String) As String
    Dim s As String
    s = raw
    s = Replace$(s, vbCr, ",")
    s = Replace$(s, vbLf, ",")
    s = Replace$(s, ";", ",")

    Dim parts() As String, i As Long, t As String, seen As String, outp As String
    parts = Split(s, ",")
    seen = " "
    For i = 0 To UBound(parts)
        t = LCase$(Trim$(parts(i)))
        ' drop empties, numbering, and anything sentence-like / too long to be a tag
        If LenB(t) <> 0 And Len(t) <= 30 And InStr(t, " ") <= 1 Then
            If InStr(seen, " " & t & " ") = 0 Then
                outp = outp & t & " "
                seen = seen & t & " "
            End If
        End If
    Next

    If LenB(outp) <> 0 Then NormalizeTags = " " & Trim$(outp) & " "
End Function

' Tag search. Returns a result set with path, tags, ext. Each whitespace-
' separated term is AND-ed (all must be present). Caller iterates and CloseRS.
Public Function Search(ByVal terms As String) As cSQLiteResults
    Dim words() As String, i As Long
    Dim sql As String, params As New Collection

    sql = "SELECT path, tags, ext FROM images WHERE classified = 1"
    words = Split(Trim$(LCase$(terms)))
    For i = 0 To UBound(words)
        If LenB(words(i)) <> 0 Then
            sql = sql & " AND tags LIKE ?"
            params.Add "%" & words(i) & "%"
        End If
    Next
    sql = sql & " ORDER BY path"

    ' marshal the collection into the ParamArray
    Select Case params.Count
        Case 0: Set Search = m_db.Query(sql)
        Case 1: Set Search = m_db.Query(sql, params(1))
        Case 2: Set Search = m_db.Query(sql, params(1), params(2))
        Case 3: Set Search = m_db.Query(sql, params(1), params(2), params(3))
        Case Else: Set Search = m_db.Query(sql, params(1), params(2), params(3), params(4))
    End Select
End Function

Public Property Get PendingCount() As Long
    PendingCount = CLng(Val(NzStr(m_db.Scalar("SELECT COUNT(*) FROM images WHERE classified = 0"))))
End Property

Public Property Get TaggedCount() As Long
    TaggedCount = CLng(Val(NzStr(m_db.Scalar("SELECT COUNT(*) FROM images WHERE classified = 1"))))
End Property

Public Property Get TotalCount() As Long
    TotalCount = CLng(Val(NzStr(m_db.Scalar("SELECT COUNT(*) FROM images"))))
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
