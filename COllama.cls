VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COllama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.8
'Site:         http://sandsprite.com
'License:      MIT
'
' Public interface mirrors CClaudeAI exactly so the three providers can be
' late bound interchangeably (Dim ai As Object : Set ai = New COllama, etc).
'
' Targets Ollama's NATIVE /api/chat endpoint (NOT the OpenAI-compat
' /v1/chat/completions). Native is the closer mirror of the Claude design:
' a stateless messages array of {role,content}, with conversation history
' kept client-side and re-sent every request -- identical to CClaudeAI.
'
' KEY DIFFERENCES vs hosted providers:
'   - No API key required for local use. ApiKey is kept for interface
'     parity and only sent as Bearer if you actually set one (e.g. when
'     Ollama sits behind an auth proxy). Locally it is ignored.
'   - "stream":false is MANDATORY. Native /api/chat defaults to NDJSON
'     streaming, which CJSON cannot parse. We force a single JSON object.
'   - Response text lives at message.content (an object, not an array).
'   - Token usage: prompt_eval_count (input), eval_count (output).
'     Ollama has no cached-token concept, so LastCachedTokens stays 0.
'   - MaxTokens maps to options.num_predict, emitted only when you set it.


Option Explicit

Public LastInputTokens As Long
Public LastOutputTokens As Long
Public LastCachedTokens As Long

Public TotalInputTokens As Long
Public TotalOutputTokens As Long
Public TotalCachedTokens As Long
Public TotalRequests As Long
Public LastReasoning As String

Private cjs As New CJSON
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
Private m_keepAlive As String     'default 30m, -1 to persist until unload, 0 to always unload (all cold starts)

' Cancel flag -- set by Cancel() from the UI, checked by the polling loop
' in CreateResponse. Reset at the start of each CreateResponse so a stale
' cancel from a prior aborted call doesn't kill the next one.
Private m_cancelRequested As Boolean
Private m_think As Boolean   ' VB6 defaults this to False, so thinking is off unless you flip it

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Conversation history for maintainContext=True. Each entry is a single
' message; entries alternate user/assistant starting with user. We store
' the JSON-escaped content (already escaped) so building the request body
' is just string concatenation. Ollama's /api/chat is stateless -- every
' request must include the full history, so we maintain it here.
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

Private Declare Function CryptBinaryToStringA Lib "crypt32.dll" ( _
    ByRef pbBinary As Byte, _
    ByVal cbBinary As Long, _
    ByVal dwFlags As Long, _
    ByVal pszString As String, _
    ByRef pcchString As Long _
) As Long

Private Const CRYPT_STRING_BASE64 As Long = 1
Private Const CRYPT_STRING_NOCRLF As Long = &H40000000

Public Property Let Think(ByVal value As Boolean): m_think = value: End Property
Public Property Get Think() As Boolean: Think = m_think: End Property

Public Sub ResetTokenTotals()
    TotalInputTokens = 0
    TotalOutputTokens = 0
    TotalCachedTokens = 0
    TotalRequests = 0
End Sub

Private Function StripThinkReasoning(ByVal s As String) As String
    Dim p As Long
    LastReasoning = Empty
    p = InStrRev(s, "</think>", -1, vbTextCompare)
    If p > 0 Then
        LastReasoning = Mid(s, 1, p)
        s = Mid$(s, p + 8)    ' 8 = Len("</think>")
    End If
    StripThinkReasoning = Trim$(s)
End Function

' No key needed for local Ollama. Open the model library so the method
' does something useful and the signature matches the other providers.
Public Function GetApiKey()
    ShellExecute 0, "open", "https://ollama.com/library", vbNullString, vbNullString, 1
End Function

Public Property Let ApiKey(ByVal value As String): m_apiKey = Trim$(value): End Property
Public Property Get ApiKey() As String: ApiKey = m_apiKey: End Property

Public Property Let Model(ByVal value As String): m_model = Trim$(value): End Property
Public Property Get Model() As String
    If LenB(m_model) = 0 Then Model = "llama3.2" Else Model = m_model
End Property

Public Property Let KeepAlive(ByVal value As String): m_keepAlive = Trim$(value): End Property
Public Property Get KeepAlive() As String
    If LenB(m_keepAlive) = 0 Then KeepAlive = "30m" Else KeepAlive = m_keepAlive
End Property

Property Let RemoteIP(ip As String)
    m_endpoint = "http://" & ip & ":11434/api/chat"
End Property

Public Property Let Endpoint(ByVal value As String): m_endpoint = Trim$(value): End Property
Public Property Get Endpoint() As String
    If LenB(m_endpoint) = 0 Then Endpoint = "http://localhost:11434/api/chat" Else Endpoint = m_endpoint
End Property

' Maps to options.num_predict. 0 (default) = model decides; positive caps
' output tokens. Property kept for interface parity with CClaudeAI; note
' the Get returns the raw value (no 1024 floor) so 0 means "model default".
Public Property Let MaxTokens(ByVal value As Long): m_maxTokens = value: End Property
Public Property Get MaxTokens() As Long: MaxTokens = m_maxTokens: End Property

Public Sub SetTimeoutsMs( _
    Optional ByVal resolveMs As Long = 5000, _
    Optional ByVal connectMs As Long = 15000, _
    Optional ByVal sendMs As Long = 15000, _
    Optional ByVal receiveMs As Long = 120000 _
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

Public Property Get HistoryLength() As Long
    HistoryLength = m_historyRoles.Count
End Property

' Clear the conversation history. Call before starting a new conversation
' chain. Equivalent in semantics to CClaudeAI.ResetContext.
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

' inputText       - the new user message for this turn
' systemPrompt    - optional system prompt (sent as a leading system-role
'                   message in the messages array)
' maintainContext - if True, the request includes all prior user/assistant
'                   messages from the in-memory history AND the new user
'                   message; on a successful response we capture the
'                   assistant reply for the next call. If False, the
'                   request is a single-message request and history is
'                   neither read nor modified.
' asyncMode       - if False (default): synchronous, http.send() blocks.
'                   Safe for non-UI callers; Cancel() is a no-op.
'                   If True: async with DoEvents-pumping poll loop.
'                   UI stays responsive and Cancel() works. CALLERS USING
'                   THIS MODE MUST GUARD AGAINST RE-ENTRANCY because
'                   DoEvents dispatches other UI events.
Public Function CreateResponse(ByVal inputText As String, Optional ByVal systemPrompt As String = "", Optional ByVal maintainContext As Boolean = False, Optional ByVal asyncMode As Boolean = False) As String

    ClearLast
    m_cancelRequested = False

    ' Pre-escape the new user message; we'll need it both for the request
    ' body and (on success) for the history.
    Dim escapedUser As String
    escapedUser = cjs.EscapeJsonString(inputText)

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
    ' Local Ollama needs no auth. Send Bearer only if a key was set, so the
    ' same class works against an auth-fronted Ollama without code changes.
    If LenB(m_apiKey) <> 0 Then
        http.setRequestHeader "Authorization", "Bearer " & m_apiKey
    End If

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
    AccumulateUsage

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
            m_historyTexts.Add cjs.EscapeJsonString(assistantText)
        End If
    End If

    CreateResponse = m_lastResponseRaw
    Exit Function

EH:
    m_lastError = "CreateResponse error: " & Err.Number & " - " & Err.Description
End Function

' Called from the UI (Cancel button) to abort an in-flight CreateResponse
' on the next poll cycle. Safe to call when no request is in flight -- the
' flag is reset at the start of each CreateResponse.
Public Sub Cancel()
    m_cancelRequested = True
End Sub

' Returns the model's reasoning from the last response (when Think was on).
' Empty if thinking was off or the model is non-reasoning. Call after CreateResponse.
Public Property Get LastThinking() As String
    Dim j As New CJSON
    If Not j.Initilized Then Exit Property
    If LenB(m_lastResponseRaw) = 0 Then Exit Property
    If Not j.loadJson(m_lastResponseRaw) Then Exit Property
    LastThinking = unixToDOS(CStr(j.getVal("message.thinking")))
End Property

' Builds the /api/chat request body. When maintainContext is True, the
' messages array is the entire history followed by the new user message.
' When False, it's just the new user message. A system-role message is
' prepended when systemPrompt is supplied (re-sent every request since the
' endpoint is stateless). "stream":false is REQUIRED for a single JSON body.
'
' escapedUser  - already-escaped JSON content of the new user message
' systemPrompt - raw (unescaped) system prompt; we escape here
' maintainContext - whether to include the historical messages
Private Function BuildRequestJson(ByVal escapedUser As String, ByVal systemPrompt As String, ByVal maintainContext As Boolean) As String
    Dim json As String
    
    json = "{""model"":""" & cjs.EscapeJsonString(Me.Model) & """"
    json = json & ",""stream"":false"
    json = json & ",""think"":" & IIf(m_think, "true", "false")
    json = json & ",""keep_alive"":""" & Me.KeepAlive & """"  ' or "-1" to pin it in VRAM until you unload
    
    json = json & ",""messages"":["

    Dim first As Boolean
    first = True

    If LenB(systemPrompt) <> 0 Then
        json = json & "{""role"":""system"",""content"":""" & cjs.EscapeJsonString(systemPrompt) & """}"
        first = False
    End If

    Dim i As Long
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

    json = json & "]"

    ' options.num_predict only when MaxTokens was explicitly set.
    If m_maxTokens > 0 Then
        json = json & ",""options"":{""num_predict"":" & m_maxTokens & "}"
    End If

    json = json & "}"
    BuildRequestJson = json
End Function

' Extracts the assistant text from Ollama's native /api/chat response.
' Response shape (non-streaming):
'   { "message": { "role":"assistant", "content":"..." }, "done":true,
'     "prompt_eval_count":N, "eval_count":M, ... }
' Error shape: { "error":"..." }  (HTTP non-200)
Function ExtractOutput(Optional json As String) As String

    Dim j As New CJSON
    Dim t As String
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

    ' Ollama error responses are a top-level "error" string.
    Dim errMsg As String
    errMsg = CStr(j.getVal("error"))
    If LenB(errMsg) <> 0 Then
        m_lastError = "API error: " & errMsg
        Exit Function
    End If

    t = CStr(j.getVal("message.content"))
    t = StripThinkReasoning(t)
    ExtractOutput = unixToDOS(t)

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


Private Function Base64EncodeFile(ByVal path As String) As String

    Dim hFile As Integer
    Dim bytes() As Byte
    Dim fileLen As Long

    On Error GoTo EH

    If LenB(Dir$(path)) = 0 Then
        m_lastError = "Base64EncodeFile: file not found: " & path
        Exit Function
    End If

    hFile = FreeFile
    Open path For Binary Access Read As #hFile
    fileLen = LOF(hFile)
    If fileLen = 0 Then
        Close #hFile
        m_lastError = "Base64EncodeFile: file is empty: " & path
        Exit Function
    End If
    ReDim bytes(0 To fileLen - 1)
    Get #hFile, , bytes
    Close #hFile

    Dim needed As Long
    Dim flags As Long
    flags = CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF

    If CryptBinaryToStringA(bytes(0), fileLen, flags, vbNullString, needed) = 0 Then
        m_lastError = "Base64EncodeFile: CryptBinaryToStringA (sizing) failed"
        Exit Function
    End If

    Dim buf As String
    buf = String$(needed, vbNullChar)
    If CryptBinaryToStringA(bytes(0), fileLen, flags, buf, needed) = 0 Then
        m_lastError = "Base64EncodeFile: CryptBinaryToStringA (encode) failed"
        Exit Function
    End If

    Base64EncodeFile = Left$(buf, needed)
    Exit Function

EH:
    If hFile <> 0 Then Close #hFile
    m_lastError = "Base64EncodeFile error: " & Err.Number & " - " & Err.Description
End Function

' Ollama native /api/chat takes images as a raw-base64 array on the message
' (no data-URI prefix, no media type). MediaTypeForPath is kept only for
' interface parity with the other providers; the value is unused here.
Private Function MediaTypeForPath(ByVal path As String) As String
    Dim ext As String
    Dim dotPos As Long
    dotPos = InStrRev(path, ".")
    If dotPos = 0 Then
        MediaTypeForPath = "image/png"
        Exit Function
    End If
    ext = LCase$(Mid$(path, dotPos + 1))
    Select Case ext
        Case "png":           MediaTypeForPath = "image/png"
        Case "jpg", "jpeg":   MediaTypeForPath = "image/jpeg"
        Case "gif":           MediaTypeForPath = "image/gif"
        Case "webp":          MediaTypeForPath = "image/webp"
        Case Else:            MediaTypeForPath = "image/png"
    End Select
End Function

' Same semantics as CreateResponse, but the user turn carries an image.
' The image is sent as raw base64 in the message's "images" array. Ollama's
' /api/chat is stateless, so on maintainContext we store the text portion of
' the turn (not the image) plus the assistant reply -- follow-up text turns
' can refer to what the model already SAID about the image, but a new visual
' inspection would require re-sending the bytes via another image call.
Public Function CreateResponseWithImage( _
    ByVal inputText As String, _
    ByVal imagePath As String, _
    Optional ByVal systemPrompt As String = "", _
    Optional ByVal maintainContext As Boolean = False, _
    Optional ByVal asyncMode As Boolean = False _
) As String

    ClearLast
    m_cancelRequested = False

    Dim b64 As String
    b64 = Base64EncodeFile(imagePath)
    If LenB(b64) = 0 Then
        Exit Function
    End If

    Dim escapedUser As String
    escapedUser = cjs.EscapeJsonString(inputText)

    Dim body As String
    body = BuildRequestJsonWithImage(escapedUser, systemPrompt, maintainContext, b64)

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
    If LenB(m_apiKey) <> 0 Then
        http.setRequestHeader "Authorization", "Bearer " & m_apiKey
    End If

    http.send body

    If asyncMode Then
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

    m_lastStatus = CLng(http.Status)
    m_lastStatusText = CStr(http.statusText)
    m_lastResponseRaw = CStr(http.responseText)
    AccumulateUsage

    If maintainContext And m_lastStatus = 200 Then
        m_historyRoles.Add "user"
        m_historyTexts.Add escapedUser

        Dim assistantText As String
        assistantText = ExtractOutput()
        If LenB(assistantText) > 0 Then
            m_historyRoles.Add "assistant"
            m_historyTexts.Add cjs.EscapeJsonString(assistantText)
        End If
    End If

    CreateResponseWithImage = m_lastResponseRaw
    Exit Function

EH:
    m_lastError = "CreateResponseWithImage error: " & Err.Number & " - " & Err.Description
End Function

' Builds an /api/chat body where the new user turn carries an "images"
' array (raw base64). Logs a sanitized copy -- the base64 payload is
' replaced with the first 10 chars + "..." so the log stays readable.
Private Function BuildRequestJsonWithImage( _
    ByVal escapedUser As String, _
    ByVal systemPrompt As String, _
    ByVal maintainContext As Boolean, _
    ByVal b64 As String _
) As String

    Dim json As String
    
    json = "{""model"":""" & cjs.EscapeJsonString(Me.Model) & """"
    json = json & ",""stream"":false"
    json = json & ",""think"":" & IIf(m_think, "true", "false")
    json = json & ",""keep_alive"":""" & Me.KeepAlive & """"  ' or "-1" to pin it in VRAM until you unload
    
    json = json & ",""messages"":["

    Dim first As Boolean
    first = True

    If LenB(systemPrompt) <> 0 Then
        json = json & "{""role"":""system"",""content"":""" & cjs.EscapeJsonString(systemPrompt) & """}"
        first = False
    End If

    Dim i As Long
    If maintainContext And Not (m_historyRoles Is Nothing) Then
        For i = 1 To m_historyRoles.Count
            If Not first Then json = json & ","
            json = json & "{""role"":""" & m_historyRoles.Item(i) & """,""content"":""" & m_historyTexts.Item(i) & """}"
            first = False
        Next i
    End If

    ' New user turn: content + images array (raw base64, no data URI).
    If Not first Then json = json & ","
    json = json & "{""role"":""user"",""content"":""" & escapedUser & """,""images"":[""" & b64 & """]}"

    json = json & "]"

    If m_maxTokens > 0 Then
        json = json & ",""options"":{""num_predict"":" & m_maxTokens & "}"
    End If

    json = json & "}"
    BuildRequestJsonWithImage = json

    ' Sanitized log copy
    Dim preview As String
    If Len(b64) > 10 Then
        preview = Left$(b64, 10) & "..."
    Else
        preview = b64
    End If
    Form1.List1.AddItem Replace$(json, b64, preview)

End Function

Private Sub AccumulateUsage()

    If m_lastStatus <> 200 Then Exit Sub
    If LenB(m_lastResponseRaw) = 0 Then Exit Sub

    Dim j As New CJSON
    If Not j.Initilized Then Exit Sub
    If Not j.loadJson(m_lastResponseRaw) Then Exit Sub
    If LenB(CStr(j.getVal("error"))) <> 0 Then Exit Sub

    LastInputTokens = CLng(j.getVal("prompt_eval_count"))
    LastOutputTokens = CLng(j.getVal("eval_count"))
    LastCachedTokens = 0   ' Ollama has no cached-token concept

    TotalInputTokens = TotalInputTokens + LastInputTokens
    TotalOutputTokens = TotalOutputTokens + LastOutputTokens
    TotalCachedTokens = TotalCachedTokens + LastCachedTokens
    TotalRequests = TotalRequests + 1

End Sub
