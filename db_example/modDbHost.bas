Attribute VB_Name = "modDbHost"
'Author:  David Zimmer <dzzie@yahoo.com>
'Site:    http://sandsprite.com
'License: MIT
'
' Database host for the ai4vb agent loop.
'
' Everything the agent touches at runtime goes through here:
'
'   DumpSchema(db)        -> DDL text spliced into the system prompt
'                            at startup. The agent sees the full table
'                            map in message #1 -- no discovery round
'                            trips wasted on PRAGMA / sqlite_master.
'
'   SafeSql(db, sql)      -> read-only SQL gate + result formatter.
'                            Rejects anything that isn't SELECT / PRAGMA
'                            / EXPLAIN / WITH. Returns either an ASCII
'                            grid (capped at MAX_ROWS) or a one-line
'                            error string the agent can self-correct
'                            from. This is what host.sql() ultimately
'                            calls from JScript.
'
'   SeedDb(db)            -> idempotent demo schema + data. users +
'                            projects + tasks, with deliberate twists
'                            (NULL assignee, paused project, lead who
'                            doesn't work tasks on their own project)
'                            so multi-stage questions have real shape.
'
' Why a module instead of folding this into Form1: keeps the form
' focused on the agent loop and UI, and lets modDbHost be reused if a
' future ai4vb spawns more than one form.

Option Explicit

' Output cap for SafeSql. The grid formatter truncates at this row count
' and emits a "... N more rows" footer. Keeps an unguarded "SELECT *"
' from blowing the agent's context window.
Private Const MAX_ROWS As Long = 50

' Per-cell display width cap. Long text values get truncated with an
' ellipsis. The agent can always ask for a narrower projection if it
' needs the full value.
Private Const MAX_CELL As Long = 60


' ============================================================
' Schema dump
' ============================================================

' Pull every CREATE TABLE / CREATE INDEX statement out of sqlite_master
' and emit them as a single readable block. Tables come first (sorted by
' name), then indexes. System tables (sqlite_*) are filtered out.
'
' Output shape:
'
'   -- tables --
'   CREATE TABLE users (...);
'   CREATE TABLE projects (...);
'   ...
'
'   -- indexes --
'   CREATE INDEX ix_projects_lead ON projects(lead_user_id);
'   ...
'
' This is the *authoritative* schema -- it's whatever the db actually
' has, not a hand-maintained copy that can drift. Splice into prompt.txt
' once at startup and the agent has a complete map.
Public Function DumpSchema(ByRef db As cSQLite) As String
    Dim t As New cSQLiteTable
    Dim sql As String
    Dim out As String
    Dim r As Long
    Dim kind As String, ddl As String

    sql = "SELECT type, sql FROM sqlite_master " & _
          "WHERE sql IS NOT NULL AND name NOT LIKE 'sqlite_%' " & _
          "ORDER BY CASE type WHEN 'table' THEN 0 " & _
          "                   WHEN 'index' THEN 1 " & _
          "                   WHEN 'view'  THEN 2 " & _
          "                   ELSE 3 END, name"

    t.LoadQuery db, sql

    Dim curKind As String
    curKind = ""

    For r = 0 To t.RowCount - 1
        kind = CStr(t.Item(r, 0))
        ddl = CStr(t.Item(r, 1))

        If kind <> curKind Then
            If Len(out) > 0 Then out = out & vbCrLf
            ' Hardcode section labels - "-- indexs --" is wrong, and a
            ' switch is clearer than IIf pluralization tricks.
            Dim sectionLabel As String
            Select Case kind
                Case "table":   sectionLabel = "-- tables --"
                Case "index":   sectionLabel = "-- indexes --"
                Case "view":    sectionLabel = "-- views --"
                Case "trigger": sectionLabel = "-- triggers --"
                Case Else:      sectionLabel = "-- " & kind & " --"
            End Select
            out = out & sectionLabel & vbCrLf
            curKind = kind
        End If

        ' Normalize whitespace in DDL so it reads as one statement per
        ' line block, terminated with a semicolon. sqlite_master stores
        ' the original text the user supplied, which may or may not
        ' have a trailing semicolon.
        out = out & Trim$(ddl)
        If Right$(Trim$(ddl), 1) <> ";" Then out = out & ";"
        out = out & vbCrLf
    Next r

    DumpSchema = out
End Function


' ============================================================
' Safe SQL gate
' ============================================================

' Single entry point for SQL the agent emits. Rejects writes, runs the
' query, returns an ASCII grid (or a one-line error). Errors are
' returned as strings beginning with "ERROR:" -- the agent loop
' surfaces them next stage so the model can self-correct.
'
' We do NOT raise on bad SQL; raising would force the JScript wrapper
' to try/catch and would surface as an MSScriptControl error rather
' than a clean tool result. Returning the error as data is more useful
' to the agent.
Public Function SafeSql(ByRef db As cSQLite, ByVal sql As String) As String
    Dim trimmed As String
    Dim verb As String

    trimmed = Trim$(sql)
    If Len(trimmed) = 0 Then
        SafeSql = "ERROR: empty SQL"
        Exit Function
    End If

    ' Strip a leading comment block so "-- comment\nSELECT ..." still
    ' classifies as a SELECT. Cheap heuristic, good enough for an
    ' agent-driven workload.
    Do While Left$(trimmed, 2) = "--"
        Dim nlPos As Long
        nlPos = InStr(trimmed, vbLf)
        If nlPos = 0 Then
            SafeSql = "ERROR: SQL is comment-only"
            Exit Function
        End If
        trimmed = Trim$(Mid$(trimmed, nlPos + 1))
    Loop

    verb = UCase$(FirstWord(trimmed))

    Select Case verb
        Case "SELECT", "PRAGMA", "EXPLAIN", "WITH"
            ' allowed
        Case Else
            SafeSql = "ERROR: only SELECT/PRAGMA/EXPLAIN/WITH allowed (got '" & verb & "')"
            Exit Function
    End Select

    ' Belt-and-braces: an attacker could chain "SELECT 1; DROP TABLE x".
    ' Our cSQLite wrapper prepares a single statement so the trailing
    ' DROP would be ignored, but checking explicitly makes intent clear.
    ' A semicolon at the very end is fine; one mid-query with content
    ' after it isn't.
    Dim semi As Long
    semi = InStr(trimmed, ";")
    If semi > 0 And semi < Len(Trim$(trimmed)) Then
        If Len(Trim$(Mid$(trimmed, semi + 1))) > 0 Then
            SafeSql = "ERROR: multi-statement SQL rejected; submit one statement"
            Exit Function
        End If
    End If

    ' Execute. We let cSQLiteTable.LoadQuery do the work; any prepare /
    ' step failure surfaces as a VB error that we catch and return as
    ' a string.
    Dim t As New cSQLiteTable
    On Error Resume Next
    t.LoadQuery db, trimmed
    If Err.Number <> 0 Then
        SafeSql = "ERROR: " & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    SafeSql = FormatGrid(t)
End Function


' ============================================================
' Grid formatter
' ============================================================

' Render a cSQLiteTable as a readable fixed-width ASCII grid:
'
'   id | name  | age | role
'   ---+-------+-----+----------------------
'    1 | Alice |  32 | Malware Analyst
'    2 | Bob   |  45 | Reverse Engineer
'   (2 rows)
'
' Empty result sets still show the header so the agent knows the query
' worked but matched nothing.
Private Function FormatGrid(ByRef t As cSQLiteTable) As String
    Dim nCols As Long, nRows As Long, displayRows As Long
    Dim c As Long, r As Long
    Dim widths() As Long
    Dim cellText As String
    Dim line As String, sep As String
    Dim out As String

    nCols = t.ColumnCount
    nRows = t.RowCount
    displayRows = nRows
    If displayRows > MAX_ROWS Then displayRows = MAX_ROWS

    If nCols = 0 Then
        FormatGrid = "(no columns) (" & nRows & " rows)"
        Exit Function
    End If

    ReDim widths(0 To nCols - 1)

    ' First pass: compute widths from headers and the rows we'll print.
    For c = 0 To nCols - 1
        widths(c) = Len(t.ColumnName(c))
    Next c

    For r = 0 To displayRows - 1
        For c = 0 To nCols - 1
            cellText = CellToText(t.Item(r, c))
            If Len(cellText) > MAX_CELL Then cellText = Left$(cellText, MAX_CELL - 1) & Chr$(133)
            If Len(cellText) > widths(c) Then widths(c) = Len(cellText)
        Next c
    Next r

    ' Header.
    For c = 0 To nCols - 1
        If c > 0 Then line = line & " | "
        line = line & PadRight(t.ColumnName(c), widths(c))
    Next c
    out = line & vbCrLf

    ' Separator.
    line = ""
    For c = 0 To nCols - 1
        If c > 0 Then line = line & "-+-"
        line = line & String$(widths(c), "-")
    Next c
    out = out & line & vbCrLf

    ' Rows.
    For r = 0 To displayRows - 1
        line = ""
        For c = 0 To nCols - 1
            cellText = CellToText(t.Item(r, c))
            If Len(cellText) > MAX_CELL Then cellText = Left$(cellText, MAX_CELL - 1) & Chr$(133)
            If c > 0 Then line = line & " | "
            line = line & PadRight(cellText, widths(c))
        Next c
        out = out & line & vbCrLf
    Next r

    ' Footer.
    If nRows = 0 Then
        out = out & "(0 rows)"
    ElseIf nRows > MAX_ROWS Then
        out = out & "(" & displayRows & " of " & nRows & " rows shown; " & _
              (nRows - displayRows) & " more rows truncated. Add LIMIT or " & _
              "narrow the WHERE clause.)"
    Else
        out = out & "(" & nRows & " row" & IIf(nRows = 1, "", "s") & ")"
    End If

    FormatGrid = out
End Function

' Stringify a cell value. NULL -> "NULL" (uppercase, distinguishable
' from an empty string). Everything else CStr'd.
Private Function CellToText(ByVal v As Variant) As String
    If IsNull(v) Then
        CellToText = "NULL"
    Else
        CellToText = CStr(v)
    End If
End Function

Private Function PadRight(ByVal s As String, ByVal width As Long) As String
    If Len(s) >= width Then
        PadRight = s
    Else
        PadRight = s & String$(width - Len(s), " ")
    End If
End Function

' First whitespace-delimited token, uppercase-friendly. Strips a leading
' open-paren so "(SELECT ...) UNION ..." classifies sanely -- we treat
' that as a SELECT-shaped read.
Private Function FirstWord(ByVal s As String) As String
    Dim i As Long, ch As String, started As Boolean
    Dim out As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or ch = "(" Then
            If started Then Exit For
        Else
            started = True
            out = out & ch
        End If
    Next i

    FirstWord = out
End Function


' ============================================================
' Seed
' ============================================================

' Build the demo db. Idempotent: drops tables first, recreates, seeds.
' Called once at startup against a :memory: or on-disk db so the agent
' has something interesting to chew on.
'
' Design notes (the twists are deliberate -- keep them):
'
'   * Sentinel has the highest budget but Carol (its lead) logs zero
'     hours on its tasks. So "who works most on the highest-budget
'     project" is NOT Carol -- it's whoever's assigned its tasks.
'
'   * One task has a NULL assignee. The agent has to handle NULL in
'     joins / aggregates correctly or it'll under-count.
'
'   * Hydra is status='paused' -- questions about active projects
'     need a status filter, not just "all projects."
Public Sub SeedDb(ByRef db As cSQLite)
    db.Execute "DROP TABLE IF EXISTS tasks"
    db.Execute "DROP TABLE IF EXISTS projects"
    db.Execute "DROP TABLE IF EXISTS users"

    db.Execute _
      "CREATE TABLE users (" & _
      "  id        INTEGER PRIMARY KEY," & _
      "  name      TEXT NOT NULL," & _
      "  age       INTEGER," & _
      "  role      TEXT," & _
      "  hire_date TEXT" & _
      ")"

    db.Execute _
      "CREATE TABLE projects (" & _
      "  id           INTEGER PRIMARY KEY," & _
      "  name         TEXT NOT NULL," & _
      "  lead_user_id INTEGER REFERENCES users(id)," & _
      "  budget       INTEGER," & _
      "  status       TEXT," & _
      "  start_date   TEXT" & _
      ")"

    db.Execute _
      "CREATE TABLE tasks (" & _
      "  id               INTEGER PRIMARY KEY," & _
      "  project_id       INTEGER NOT NULL REFERENCES projects(id)," & _
      "  assignee_user_id INTEGER REFERENCES users(id)," & _
      "  title            TEXT NOT NULL," & _
      "  hours            REAL," & _
      "  done             INTEGER" & _
      ")"

    db.Execute "CREATE INDEX ix_projects_lead ON projects(lead_user_id)"
    db.Execute "CREATE INDEX ix_tasks_project ON tasks(project_id)"
    db.Execute "CREATE INDEX ix_tasks_assignee ON tasks(assignee_user_id)"

    ' --- users ---
    db.Execute "INSERT INTO users(id,name,age,role,hire_date) VALUES " & _
        "(1,'Roger',50,'Security Researcher','2008-04-12')," & _
        "(2,'Alice',32,'Malware Analyst','2019-09-03')," & _
        "(3,'Bob',45,'Reverse Engineer','2014-01-20')," & _
        "(4,'Carol',38,'Threat Intelligence Lead','2017-06-15')"

    ' --- projects ---
    ' Carol leads Sentinel, the largest budget. Hydra is paused.
    db.Execute "INSERT INTO projects(id,name,lead_user_id,budget,status,start_date) VALUES " & _
        "(1,'Phoenix',  2, 150000,'active', '2023-01-15')," & _
        "(2,'Hydra',    3,  85000,'paused', '2022-07-01')," & _
        "(3,'Sentinel', 4, 420000,'active', '2024-03-22')," & _
        "(4,'Chimera',  1, 210000,'active', '2023-09-10')"

    ' --- tasks ---
    ' Sentinel tasks worked by Alice + Bob, NOT by Carol the lead.
    ' Task 7 has NULL assignee. Hydra has one done task and one open.
    db.Execute "INSERT INTO tasks(id,project_id,assignee_user_id,title,hours,done) VALUES " & _
        "(1, 1, 2, 'Unpack sample family A',     24.5, 1)," & _
        "(2, 1, 3, 'IDA scripting harness',      40.0, 1)," & _
        "(3, 1, 2, 'YARA rule tuning',           12.0, 0)," & _
        "(4, 2, 3, 'Stub decoder rewrite',       18.0, 1)," & _
        "(5, 2, 1, 'Triage backlog',              8.0, 0)," & _
        "(6, 3, 2, 'Sentinel: collector telemetry', 56.0, 1)," & _
        "(7, 3, NULL,'Sentinel: unassigned spike', 4.0, 0)," & _
        "(8, 3, 3, 'Sentinel: PE32+ corpus',      72.0, 1)," & _
        "(9, 3, 2, 'Sentinel: dashboard',         30.0, 0)," & _
        "(10,4, 1, 'Chimera: ham radio bridge',   22.0, 1)," & _
        "(11,4, 4, 'Chimera: threat intel feed',  16.5, 0)"
End Sub
