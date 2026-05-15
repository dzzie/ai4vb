VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClaudeAI"
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
Private m_maxTokens As Long

' Timeout components (ms)
Private m_toResolve As Long
Private m_toConnect As Long
Private m_toSend As Long
Private m_toReceive As Long

Private m_lastStatus As Long
Private m_lastStatusText As String
Private m_lastResponseRaw As String
Private m_lastError As String

' Cancel flag — set by Cancel() from the UI, checked by the polling loop
' in CreateResponse. Reset at the start of each CreateResponse so a stale
' cancel from a prior aborted call doesn't kill the next one.
Private m_cancelRequested As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Conversation history for maintainContext=True. Each entry is a single
' message; entries alternate user/assistant starting with user. We store
' the JSON-escaped content (already escaped) so building the request body
' is just string concatenation. Anthropic's /v1/messages is stateless —
' every request must include the full history, so we maintain it here.
Private m_historyRoles As Collection   ' "user" | "assistant"
Private m_historyTexts As Collection   ' escaped content strings

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long

Public Function GetApiKey()
    ShellExecute 0, "open", "https://console.anthropic.com/settings/keys", vbNullString, vbNullString, 1
End Function

Public Property Let ApiKey(ByVal value As String): m_apiKey = Trim$(value): End Property
Public Property Get ApiKey() As String: ApiKey = m_apiKey: End Property

Public Property Let Model(ByVal value As String): m_model = Trim$(value): End Property
Public Property Get Model() As String
    If LenB(m_model) = 0 Then Model = "claude-sonnet-4-6" Else Model = m_model
End Property

Public Property Let Endpoint(ByVal value As String): m_endpoint = Trim$(value): End Property
Public Property Get Endpoint() As String
    If LenB(m_endpoint) = 0 Then Endpoint = "https://api.anthropic.com/v1/messages" Else Endpoint = m_endpoint
End Property

' Anthropic requires max_tokens in every request. Default 1024.
Public Property Let MaxTokens(ByVal value As Long): m_maxTokens = value: End Property
Public Property Get MaxTokens() As Long
    If m_maxTokens <= 0 Then MaxTokens = 1024 Else MaxTokens = m_maxTokens
End Property

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

' Number of messages currently in the history (user + assistant turns).
Public Property Get HistoryLength() As Long
    HistoryLength = m_historyRoles.Count
End Property

' Clear the conversation history. Call before starting a new conversation
' chain. Equivalent in semantics to COpenAI.ResetContext.
Public Sub ResetContext()
    Set m_historyRoles = New Collection
    Set m_historyTexts = New Collection
End Sub

Private Sub ClearLast()
    m_lastStatus = 0
    m_lastStatusText = vbNullString
    m_lastResponseRaw = vbNullString
    m_lastError = vbNullString
End Sub

' inputText      - the new user message for this turn
' systemPrompt   - optional system prompt (sent as top-level "system" field)
' maintainContext - if True, the request includes all prior user/assistant
'                   messages from the in-memory history AND the new user
'                   message; on a successful response we capture the
'                   assistant reply for the next call. If False, the
'                   request is a single-message request and history is
'                   neither read nor modified.
' asyncMode      - if False (default): synchronous, http.send() blocks.
'                  Safe for non-UI callers; Cancel() is a no-op.
'                  If True: async with DoEvents-pumping poll loop.
'                  UI stays responsive and Cancel() works. CALLERS USING
'                  THIS MODE MUST GUARD AGAINST RE-ENTRANCY because
'                  DoEvents dispatches other UI events.
Public Function CreateResponse(ByVal inputText As String, Optional ByVal systemPrompt As String = "", Optional ByVal maintainContext As Boolean = False, Optional ByVal asyncMode As Boolean = False) As String

    ClearLast
    m_cancelRequested = False

    If LenB(m_apiKey) = 0 Then
        m_lastError = "ApiKey not set."
        Exit Function
    End If

    ' Pre-escape the new user message; we'll need it both for the request
    ' body and (on success) for the history.
    Dim escapedUser As String
    escapedUser = EscapeJsonString(inputText)

    Dim body As String
    body = BuildRequestJson(escapedUser, systemPrompt, maintainContext)

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    Dim r As Long, c As Long, s As Long, v As Long
    r = IIf(m_toResolve > 0, m_toResolve, 5000)
    c = IIf(m_toConnect > 0, m_toConnect, 15000)
    s = IIf(m_toSend > 0, m_toSend, 15000)
    v = IIf(m_toReceive > 0, m_toReceive, 60000)
    http.setTimeouts r, c, s, v

    http.Open "POST", Me.Endpoint, asyncMode
    http.setRequestHeader "Content-Type", "application/json"
    ' Anthropic uses x-api-key, not Bearer
    http.setRequestHeader "x-api-key", m_apiKey
    ' Required version header
    http.setRequestHeader "anthropic-version", "2023-06-01"

    http.send body

    If asyncMode Then
        ' Poll for completion. DoEvents keeps the UI alive so the user can
        ' click Cancel; Sleep 30 avoids pegging a CPU core.
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
    ' In sync mode, http.send already blocked until completion.

    m_lastStatus = CLng(http.Status)
    m_lastStatusText = CStr(http.statusText)
    m_lastResponseRaw = CStr(http.responseText)

    ' Only commit to history on success and only when context is being
    ' maintained. We append BOTH the user message we just sent AND the
    ' assistant reply we just received, so the next call's history is
    ' complete. A failed request leaves history unchanged so retries
    ' don't accumulate duplicate user messages.
    If maintainContext And m_lastStatus = 200 Then
        m_historyRoles.Add "user"
        m_historyTexts.Add escapedUser

        Dim assistantText As String
        assistantText = ExtractOutput()    ' uses m_lastResponseRaw
        If LenB(assistantText) > 0 Then
            m_historyRoles.Add "assistant"
            m_historyTexts.Add EscapeJsonString(assistantText)
        End If
    End If

    CreateResponse = m_lastResponseRaw
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

' Builds the /v1/messages request body. When maintainContext is True, the
' messages array is the entire history followed by the new user message.
' When False, it's just the new user message.
'
' escapedUser  - already-escaped JSON content of the new user message
' systemPrompt - raw (unescaped) system prompt; we escape here
' maintainContext - whether to include the historical messages
Private Function BuildRequestJson(ByVal escapedUser As String, ByVal systemPrompt As String, ByVal maintainContext As Boolean) As String
    Dim json As String
    json = "{""model"":""" & EscapeJsonString(Me.Model) & """"
    json = json & ",""max_tokens"":" & Me.MaxTokens
    If LenB(systemPrompt) <> 0 Then
        json = json & ",""system"":""" & EscapeJsonString(systemPrompt) & """"
    End If

    json = json & ",""messages"":["

    Dim i As Long, first As Boolean
    first = True

    If maintainContext And Not (m_historyRoles Is Nothing) Then
        For i = 1 To m_historyRoles.Count
            If Not first Then json = json & ","
            json = json & "{""role"":""" & m_historyRoles.Item(i) & """,""content"":""" & m_historyTexts.Item(i) & """}"
            first = False
        Next i
    End If

    ' Append the new user message (escapedUser is already JSON-escaped).
    If Not first Then json = json & ","
    json = json & "{""role"":""user"",""content"":""" & escapedUser & """}"

    json = json & "]}"
    BuildRequestJson = json
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

' Extracts the assistant text from Anthropic's response.
' Response shape: { "content": [ { "type": "text", "text": "..." }, ... ] }
Function ExtractOutput(Optional json As String) As String

    Dim j As New CJSON
    Dim i As Long
    Dim contentLen As Long
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

    ' Anthropic error responses include a top-level "error" object
    Dim errType As String
    errType = CStr(j.getVal("error.type"))
    If LenB(errType) <> 0 Then
        m_lastError = "API error [" & errType & "]: " & CStr(j.getVal("error.message"))
        Exit Function
    End If

    contentLen = CLng(j.getVal("content.length"))

    For i = 0 To contentLen - 1
        If CStr(j.getVal("content[" & i & "].type")) = "text" Then
            t = CStr(j.getVal("content[" & i & "].text"))
            result = result & t
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
    Set m_historyRoles = New Collection
    Set m_historyTexts = New Collection
End Sub
