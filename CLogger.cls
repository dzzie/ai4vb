VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.7
'Site:         http://sandsprite.com
'License:      MIT


Option Explicit

' Tiny append-only logger. Open once, write many, close on terminate.
' Each line is timestamped. Sections get a header so the file is easy to skim.

Private m_path As String
Private m_fnum As Integer
Private m_open As Boolean

Public Sub OpenLog(ByVal filePath As String, Optional ByVal truncate As Boolean = False)
    CloseLog
    m_path = filePath
    m_fnum = FreeFile
    If truncate Then
        Open m_path For Output As #m_fnum
    Else
        Open m_path For Append As #m_fnum
    End If
    m_open = True
    WriteLine "================================================================"
    WriteLine "=== LOG OPENED " & Now & " ==="
    WriteLine "================================================================"
End Sub

Public Sub CloseLog()
    If m_open Then
        WriteLine "=== LOG CLOSED " & Now & " ==="
        Close #m_fnum
        m_open = False
    End If
End Sub

Public Sub Section(ByVal title As String)
    WriteLine ""
    WriteLine "---------- " & title & " ----------"
End Sub

' Multi-line aware: splits on CRLF so each line gets its own timestamp prefix.
' If you don't want that, use WriteRaw.
Public Sub log(ByVal label As String, ByVal value As String)
    Dim lines() As String
    Dim normalized As String
    normalized = Replace(value, vbCrLf, vbLf)
    lines = Split(normalized, vbLf)
    Dim i As Long
    If UBound(lines) = 0 Then
        WriteLine "[" & TS() & "] " & label & ": " & lines(0)
    Else
        WriteLine "[" & TS() & "] " & label & ":"
        For i = 0 To UBound(lines)
            WriteLine "    " & lines(i)
        Next
    End If
End Sub

Public Sub WriteLine(ByVal text As String)
    If Not m_open Then Exit Sub
    Print #m_fnum, text
End Sub

Public Sub WriteRaw(ByVal text As String)
    If Not m_open Then Exit Sub
    Print #m_fnum, text;
End Sub

' Force buffered data to disk so you can tail the file while the agent runs.
' VB does not expose fflush directly; closing+reopening append is the portable way.
Public Sub Flush()
    If Not m_open Then Exit Sub
    Close #m_fnum
    m_fnum = FreeFile
    Open m_path For Append As #m_fnum
End Sub

Private Function TS() As String
    TS = Format$(Now, "hh:nn:ss")
End Function

Private Sub Class_Terminate()
    CloseLog
End Sub
