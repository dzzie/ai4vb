VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COpenAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.7
'Site:         http://sandsprite.com
'License:      MIT



Option Explicit
 
Private m_apiKey As String
Private m_model As String
Private m_endpoint As String

' Timeout components (ms)
Private m_toResolve As Long
Private m_toConnect As Long
Private m_toSend As Long
Private m_toReceive As Long

Private m_lastStatus As Long
Private m_lastStatusText As String
Private m_lastResponseRaw As String
Private m_lastError As String
Private m_lastResponseId As String

' Cancel flag — set by Cancel() from the UI, checked by the polling loop
' in CreateResponse. Reset at the start of each CreateResponse so a stale
' cancel from a prior aborted call doesn't kill the next one.
Private m_cancelRequested As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long

Public Function GetApiKey()
     ShellExecute 0, "open", "https://platform.openai.com/api-keys", vbNullString, vbNullString, 1
End Function


Public Property Let ApiKey(ByVal value As String): m_apiKey = Trim$(value): End Property
Public Property Get ApiKey() As String: ApiKey = m_apiKey: End Property

Public Property Let Model(ByVal value As String): m_model = Trim$(value): End Property
Public Property Get Model() As String
    If LenB(m_model) = 0 Then Model = "gpt-4.1" Else Model = m_model
End Property

Public Property Let Endpoint(ByVal value As String): m_endpoint = Trim$(value): End Property
Public Property Get Endpoint() As String
    If LenB(m_endpoint) = 0 Then Endpoint = "https://api.openai.com/v1/responses" Else Endpoint = m_endpoint
End Property

' Defaults are reasonable; tweak in your UI
Public Sub SetTimeoutsMs( _
    Optional ByVal resolveMs As Long = 5000, _
    Optional ByVal connectMs As Long = 15000, _
    Optional ByVal sendMs As Long = 15000, _
    Optional ByVal receiveMs As Long = 60000 _
)
    m_toResolve = resolveMs
    m_toConnect = connectMs
    m_toSend = sendMs
    m_toReceive = receiveMs
End Sub

Public Property Get LastStatus() As Long: LastStatus = m_lastStatus: End Property
Public Property Get LastStatusText() As String: LastStatusText = m_lastStatusText: End Property
Public Property Get LastResponseRaw() As String: LastResponseRaw = m_lastResponseRaw: End Property
Public Property Get LastError() As String: LastError = m_lastError: End Property
Public Property Get LastResponseId() As String: LastResponseId = m_lastResponseId: End Property

' Call this to break the conversation chain and start fresh
Public Sub ResetContext()
    m_lastResponseId = vbNullString
End Sub

Private Sub ClearLast()
    m_lastStatus = 0
    m_lastStatusText = vbNullString
    m_lastResponseRaw = vbNullString
    m_lastError = vbNullString
    ' Note: m_lastResponseId is intentionally NOT cleared here so chaining survives ClearLast
End Sub

' asyncMode controls how the HTTP call is made:
'   False (default) — synchronous: http.send() blocks until response.
'                     Safe for non-UI callers. Cancel() is a no-op in
'                     this mode (nothing to abort between).
'   True            — asynchronous: http.send() returns immediately,
'                     we poll readyState with DoEvents and Sleep. UI
'                     stays responsive and Cancel() works. CALLERS
'                     USING THIS MODE MUST GUARD AGAINST RE-ENTRANCY
'                     because DoEvents will dispatch other UI events
'                     (including a second click of the same button).
Public Function CreateResponse(ByVal inputText As String, Optional ByVal instructions As String = "", Optional ByVal maintainContext As Boolean = False, Optional ByVal asyncMode As Boolean = False) As String

    ClearLast
    m_cancelRequested = False

    If LenB(m_apiKey) = 0 Then
        m_lastError = "ApiKey not set."
        Exit Function
    End If

    Dim body As String
    body = BuildRequestJson(inputText, instructions, maintainContext)

    On Error GoTo EH

    Dim http As Object
    ' ServerXMLHTTP has setTimeouts (XMLHTTP does not).
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Timeouts: resolve, connect, send, receive (all required)
    Dim r As Long, c As Long, s As Long, v As Long
    r = IIf(m_toResolve > 0, m_toResolve, 5000)
    c = IIf(m_toConnect > 0, m_toConnect, 15000)
    s = IIf(m_toSend > 0, m_toSend, 15000)
    v = IIf(m_toReceive > 0, m_toReceive, 60000)
    http.setTimeouts r, c, s, v

    ' Open with sync/async per asyncMode parameter.
    http.Open "POST", Me.Endpoint, asyncMode
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & m_apiKey

    http.send body

    If asyncMode Then
        ' Poll until readyState=4 (COMPLETED), with DoEvents so the UI breathes
        ' and the user can click Cancel. Sleep 30ms between polls to avoid
        ' pegging a CPU core (33 polls/sec is plenty for UI purposes).
        Do While http.readyState <> 4
            DoEvents
            Sleep 30
            If m_cancelRequested Then
                http.abort
                m_lastError = "Cancelled"
                m_lastStatus = 0
                Exit Function
            End If
        Loop
    End If
    ' In sync mode, http.send already blocked until completion; no poll needed.

    m_lastStatus = CLng(http.Status)
    m_lastStatusText = CStr(http.statusText)
    m_lastResponseRaw = CStr(http.responseText)

    ' Always capture the response id so the next chained call can use it
    m_lastResponseId = ParseResponseId(m_lastResponseRaw)

    CreateResponse = m_lastResponseRaw
    Form1.List1.AddItem m_lastResponseRaw 'for early testing only
    Exit Function

EH:
    m_lastError = "CreateResponse error: " & Err.Number & " - " & Err.Description
End Function

' Called from the UI (Cancel button) to abort an in-flight CreateResponse
' on the next poll cycle. Safe to call when no request is in flight — the
' flag is reset at the start of each CreateResponse.
Public Sub Cancel()
    m_cancelRequested = True
End Sub

Private Function BuildRequestJson(ByVal inputText As String, ByVal instructions As String, ByVal maintainContext As Boolean) As String
    Dim json As String
    json = "{""model"":""" & EscapeJsonString(Me.Model) & """,""input"":""" & EscapeJsonString(inputText) & """"
    If LenB(instructions) <> 0 Then
        json = json & ",""instructions"":""" & EscapeJsonString(instructions) & """"
    End If
    If maintainContext And LenB(m_lastResponseId) <> 0 Then
        json = json & ",""previous_response_id"":""" & EscapeJsonString(m_lastResponseId) & """"
    End If
    json = json & "}"
    BuildRequestJson = json
    Form1.List1.AddItem json 'jsut for testing...
End Function

' Lightweight id extractor � avoids spinning up CJSON just for one field.
' OpenAI always returns "id":"resp_..." near the top of the response.
Private Function ParseResponseId(ByVal json As String) As String
    Dim pos As Long, endPos As Long
    pos = InStr(json, """id"":""")
    If pos > 0 Then
        pos = pos + 6  ' skip past "id":"
    Else
        pos = InStr(json, """id"": """)
        If pos = 0 Then Exit Function
        pos = pos + 7  ' skip past "id": "
    End If
    endPos = InStr(pos, json, """")
    If endPos = 0 Then Exit Function
    ParseResponseId = Mid$(json, pos, endPos - pos)
End Function

Private Function EscapeJsonString(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace$(t, "\", "\\")
    t = Replace$(t, """", "\""")
    t = Replace$(t, vbCrLf, "\n")
    t = Replace$(t, vbCr, "\n")
    t = Replace$(t, vbLf, "\n")
    t = Replace$(t, vbTab, "\t")
    t = Replace$(t, Chr$(0), "")
    EscapeJsonString = t
End Function

Function ExtractOutput(Optional json As String) As String
    
    Dim j As New CJSON
    Dim i As Long, k As Long
    Dim outLen As Long, contentLen As Long
    Dim t As String, result As String
    Dim js As String
    
    If Not j.Initilized Then
        m_lastError = "CJSON not initialized (ScriptControl missing?)"
        Exit Function
    End If
    
    If Len(json) = 0 Then js = m_lastResponseRaw Else js = json
    
    If Not j.loadJson(js) Then
        m_lastError = "loadJson failed: " & j.jsErr
        Exit Function
    End If
    
    outLen = CLng(j.getVal("output.length"))
    
    For i = 0 To outLen - 1
        If CStr(j.getVal("output[" & i & "].type")) = "message" Then
            
            contentLen = CLng(j.getVal("output[" & i & "].content.length"))
            
            For k = 0 To contentLen - 1
                If CStr(j.getVal("output[" & i & "].content[" & k & "].type")) = "output_text" Then
                    t = CStr(j.getVal("output[" & i & "].content[" & k & "].text"))
                    result = result & t
                End If
            Next
            
        End If
    Next

    ExtractOutput = unixToDOS(result)
    
End Function

Function unixToDOS(ByVal tmp As String)
    Dim isMixed As Boolean
    isMixed = (InStr(tmp, vbCrLf) > 0)
    If isMixed Then tmp = VBA.Replace(tmp, vbCrLf, Chr(5))
    tmp = VBA.Replace(tmp, vbLf, vbCrLf)
    If isMixed Then tmp = VBA.Replace(tmp, Chr(5), vbCrLf)
    unixToDOS = tmp
End Function

Private Sub Class_Initialize()
    SetTimeoutsMs
End Sub

