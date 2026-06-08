VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGemini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.8
'Site:         http://sandsprite.com
'License:      MIT
'
' Public interface mirrors CClaudeAI / COllama exactly so all providers can
' be late bound interchangeably (Dim ai As Object : Set ai = New CGemini).
'
' Targets Google's NATIVE generateContent API:
'   https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent
'
' Gemini differs from the other providers on the wire; the public interface
' hides all of it:
'   - Auth is a header, x-goog-api-key, NOT an Authorization: Bearer.
'   - The model name goes in the URL PATH, not the request body. Endpoint
'     holds the base; CreateResponse splices in Model + ":generateContent".
'   - Messages are "contents" of {role, parts:[{text}]}; the assistant role
'     is "model" (not "assistant"). History is kept client-side and re-sent.
'   - A system prompt is a top-level "systemInstruction", not a role.
'   - The answer is candidates[0].content.parts[*].text; usage is in
'     usageMetadata (promptTokenCount / candidatesTokenCount /
'     cachedContentTokenCount).
'   - MaxTokens maps to generationConfig.maxOutputTokens (emitted only if set).
'
' THINKING: Gemini 3.x models always reason; there is no clean off switch
' like Ollama's. The Think property here controls whether thought SUMMARIES
' are returned (includeThoughts), which is what makes LastThinking populate.
' To constrain reasoning DEPTH on 3.x models, set ThinkingLevel ("low"/"high")
' -- it emits generationConfig.thinkingConfig.thinkingLevel. NOTE: 2.5-series
' models use thinkingBudget instead of thinkingLevel; this property targets
' the 3.x idiom (matching the default gemini-3.5-flash). Leave it empty to
' use the model default.


Option Explicit

Public LastInputTokens As Long
Public LastOutputTokens As Long
Public LastCachedTokens As Long

Public TotalInputTokens As Long
Public TotalOutputTokens As Long
Public TotalCachedTokens As Long
Public TotalRequests As Long

Private cjs As New CJSON
Private m_apiKey As String
Private m_model As String
Private m_endpoint As String        ' base; model + ":generateContent" appended per call
Private m_maxTokens As Long
Private m_think As Boolean           ' return thought summaries (populates LastThinking)
Private m_thinkingLevel As String    ' "" | "low" | "high"  (3.x idiom)

' Timeout components (ms)
Private m_toResolve As Long
Private m_toConnect As Long
Private m_toSend As Long
Private m_toReceive As Long

Private m_lastStatus As Long
Private m_lastStatusText As String
Private m_lastResponseRaw As String
Private m_lastError As String

Private m_cancelRequested As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Conversation history for maintainContext=True. Roles are "user" | "model"
' (Gemini's assistant role is "model"). generateContent is stateless, so the
' full history is re-sent every request.
Private m_historyRoles As Collection
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

Private m_modelId As gemAiModel
Private m_customModel As String

Public Enum gemAiModel
    mdl_UserDefined = -1
    gem_2_5_flash_lite = 0
    gem_3_5_flash = 1
    gem_3_1_pro = 2
End Enum

' Resolve an enum member to its canonical API/tag string.
' mdl_UserDefined (and anything unrecognized) returns "" so callers know to
' fall back to a manually-set Model string.
Public Function ModelStr(ByVal m As gemAiModel) As String
    Select Case m
        Case gem_3_5_flash:       ModelStr = "gemini-3.5-flash"
        Case gem_3_1_pro:         ModelStr = "gemini-3.1-pro-preview"
        Case gem_2_5_flash_lite:  ModelStr = "gemini-2.5-flash-lite"
        ' --- escape hatch / unknown ---
        Case Else:                ModelStr = ""
    End Select
End Function

Sub SetModel(ByVal m As gemAiModel, Optional customModel As String)
    m_modelId = m
    ' Resolve known enum members to their wire string. On mdl_UserDefined we
    ' leave m_model untouched so a manually-set Model string survives.
    If m <> mdl_UserDefined Then
        m_model = ModelStr(m)
    Else
        m_customModel = Trim(customModel)
    End If
End Sub

Public Property Get ModelId() As gemAiModel: ModelId = m_modelId: End Property

Public Property Get Model()
    If m_modelId = mdl_UserDefined Then
        Model = m_customModel
    Else
        Model = Trim$(m_model)
    End If
End Property

Public Sub ResetTokenTotals()
    TotalInputTokens = 0
    TotalOutputTokens = 0
    TotalCachedTokens = 0
    TotalRequests = 0
End Sub

' Open the Google AI Studio key page so the signature matches the other
' providers' GetApiKey (where it opens the relevant console).
Public Function GetApiKey()
    ShellExecute 0, "open", "https://aistudio.google.com/apikey", vbNullString, vbNullString, 1
End Function

Function isApiKeySet() As Boolean
    isApiKeySet = (Len(m_apiKey) > 0)
End Function

Function LoadSavedApiKey() As String
    m_apiKey = GetSetting("ai4vb", "gemini", "key")
    LoadSavedApiKey = m_apiKey
End Function

Function SaveApiKey(Optional ByVal key As String)
    If Len(key) = 0 Then
        If Len(m_apiKey) = 0 Then
            key = InputBox("Enter Gemini API Key")
        Else
            key = m_apiKey
        End If
    End If
    Call SaveSetting("ai4vb", "gemini", "key", key)
    m_apiKey = key
End Function

Public Property Let ApiKey(ByVal value As String): m_apiKey = Trim$(value): End Property
Public Property Get ApiKey() As String: ApiKey = m_apiKey: End Property

' Endpoint is the BASE (model is appended per-request). Override only if you
' need a different API version or a proxy.
Public Property Let Endpoint(ByVal value As String): m_endpoint = Trim$(value): End Property
Public Property Get Endpoint() As String
    If LenB(m_endpoint) = 0 Then
        Endpoint = "https://generativelanguage.googleapis.com/v1beta/models"
    Else
        Endpoint = m_endpoint
    End If
End Property

' Full per-request URL: base/{model}:generateContent
Private Function BuildUrl() As String
    Dim b As String
    b = Me.Endpoint
    If Right$(b, 1) = "/" Then b = Left$(b, Len(b) - 1)
    BuildUrl = b & "/" & Me.Model & ":generateContent"
End Function

' Maps to generationConfig.maxOutputTokens. 0 (default) = model default.
Public Property Let MaxTokens(ByVal value As Long): m_maxTokens = value: End Property
Public Property Get MaxTokens() As Long: MaxTokens = m_maxTokens: End Property

' When True, request thought summaries so LastThinking is populated.
' Does NOT disable reasoning (Gemini 3.x always reasons) -- see ThinkingLevel.
Public Property Let Think(ByVal value As Boolean): m_think = value: End Property
Public Property Get Think() As Boolean: Think = m_think: End Property

' "" | "low" | "high". Emits generationConfig.thinkingConfig.thinkingLevel
' (3.x idiom). For 2.5-series models swap to thinkingBudget. Empty = default.
Public Property Let ThinkingLevel(ByVal value As String): m_thinkingLevel = LCase$(Trim$(value)): End Property
Public Property Get ThinkingLevel() As String: ThinkingLevel = m_thinkingLevel: End Property

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
' systemPrompt    - optional system prompt (top-level systemInstruction)
' maintainContext - include prior turns from history AND capture the reply
' asyncMode       - False (default) sync; True = DoEvents poll loop, Cancel works
Public Function CreateResponse(ByVal inputText As String, Optional ByVal systemPrompt As String = "", Optional ByVal maintainContext As Boolean = False, Optional ByVal asyncMode As Boolean = False) As String

    ClearLast
    m_cancelRequested = False

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

    http.Open "POST", BuildUrl(), asyncMode
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "x-goog-api-key", m_apiKey

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

        Dim modelText As String
        modelText = ExtractOutput()
        If LenB(modelText) > 0 Then
            m_historyRoles.Add "model"
            m_historyTexts.Add cjs.EscapeJsonString(modelText)
        End If
    End If

    CreateResponse = m_lastResponseRaw
    Exit Function

EH:
    m_lastError = "CreateResponse error: " & Err.Number & " - " & Err.Description
End Function

Public Sub Cancel()
    m_cancelRequested = True
End Sub

' Builds the generateContent body. contents is the history (when maintained)
' plus the new user turn. systemInstruction is top-level. generationConfig is
' emitted only when MaxTokens / Think / ThinkingLevel are in play.
Private Function BuildRequestJson(ByVal escapedUser As String, ByVal systemPrompt As String, ByVal maintainContext As Boolean) As String

    Dim json As String
    
    json = "{""contents"":["

    Dim first As Boolean
    first = True

    Dim i As Long
    If maintainContext And Not (m_historyRoles Is Nothing) Then
        For i = 1 To m_historyRoles.Count
            If Not first Then json = json & ","
            json = json & "{""role"":""" & m_historyRoles.Item(i) & """,""parts"":[{""text"":""" & m_historyTexts.Item(i) & """}]}"
            first = False
        Next i
    End If

    If Not first Then json = json & ","
    json = json & "{""role"":""user"",""parts"":[{""text"":""" & escapedUser & """}]}"
    json = json & "]"

    If LenB(systemPrompt) <> 0 Then
        json = json & ",""systemInstruction"":{""parts"":[{""text"":""" & cjs.EscapeJsonString(systemPrompt) & """}]}"
    End If

    json = json & GenerationConfigFragment()

    json = json & "}"
    BuildRequestJson = json
End Function

' Returns ",""generationConfig"":{...}" or "" if nothing to configure.
Private Function GenerationConfigFragment() As String
    Dim inner As String
    Dim need As Boolean

    If m_maxTokens > 0 Then
        inner = inner & """maxOutputTokens"":" & m_maxTokens
        need = True
    End If

    Dim Think As String
    If m_think Then Think = """includeThoughts"":true"
    If LenB(m_thinkingLevel) <> 0 Then
        If LenB(Think) <> 0 Then Think = Think & ","
        Think = Think & """thinkingLevel"":""" & m_thinkingLevel & """"
    End If
    If LenB(Think) <> 0 Then
        If need Then inner = inner & ","
        inner = inner & """thinkingConfig"":{" & Think & "}"
        need = True
    End If

    If need Then
        GenerationConfigFragment = ",""generationConfig"":{" & inner & "}"
    Else
        GenerationConfigFragment = ""
    End If
End Function

' Extracts the answer text from candidates[0].content.parts, skipping any
' part flagged "thought":true (those are reasoning summaries -> LastThinking).
' Response shape:
'   { "candidates":[{ "content":{ "parts":[{"text":"..."}], "role":"model" },
'      "finishReason":"STOP" }],
'     "usageMetadata":{...} }
' Error shape: { "error":{ "code":N, "message":"...", "status":"..." } }
Function ExtractOutput(Optional json As String) As String

    Dim j As New CJSON
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

    Dim errMsg As String
    errMsg = CStr(j.getVal("error.message"))
    If LenB(errMsg) <> 0 Then
        m_lastError = "API error: " & errMsg
        Exit Function
    End If

    Dim out As String, i As Long
    Dim partText As String, thoughtFlag As String
    For i = 0 To 31
        thoughtFlag = LCase$(CStr(j.getVal("candidates[0].content.parts[" & i & "].thought")))
        partText = CStr(j.getVal("candidates[0].content.parts[" & i & "].text"))
        ' empty text and not a thought part => past the end of the array
        If LenB(partText) = 0 And thoughtFlag <> "true" Then Exit For
        If thoughtFlag <> "true" Then out = out & partText
    Next i

    If LenB(out) = 0 Then
        ' No answer text -- surface a block reason if present.
        Dim blocked As String
        blocked = CStr(j.getVal("promptFeedback.blockReason"))
        If LenB(blocked) <> 0 Then m_lastError = "Blocked: " & blocked
    End If

    ExtractOutput = unixToDOS(out)

End Function

' Returns the model's reasoning summaries from the last response (when Think
' was on). Empty if Think was off. Call after CreateResponse.
Public Property Get LastThinking() As String
    Dim j As New CJSON
    If Not j.Initilized Then Exit Property
    If LenB(m_lastResponseRaw) = 0 Then Exit Property
    If Not j.loadJson(m_lastResponseRaw) Then Exit Property

    Dim out As String, i As Long
    Dim partText As String, thoughtFlag As String
    For i = 0 To 31
        thoughtFlag = LCase$(CStr(j.getVal("candidates[0].content.parts[" & i & "].thought")))
        partText = CStr(j.getVal("candidates[0].content.parts[" & i & "].text"))
        If LenB(partText) = 0 And thoughtFlag <> "true" Then Exit For
        If thoughtFlag = "true" Then out = out & partText
    Next i

    LastThinking = unixToDOS(out)
End Property

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
    SetModel gem_3_5_flash
    LoadSavedApiKey
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

' Gemini wants the mime type alongside the base64 (inline_data.mime_type),
' so unlike the Ollama class this helper is actually used.
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
        Case "heic":          MediaTypeForPath = "image/heic"
        Case "heif":          MediaTypeForPath = "image/heif"
        Case Else:            MediaTypeForPath = "image/png"
    End Select
End Function

' Same semantics as CreateResponse, but the user turn carries an image as an
' inline_data part (base64 + mime_type). Gemini models are natively
' multimodal, so no special model selection is required -- gemini-3.5-flash
' accepts images directly.
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

    Dim mt As String
    mt = MediaTypeForPath(imagePath)

    Dim escapedUser As String
    escapedUser = cjs.EscapeJsonString(inputText)

    Dim body As String
    body = BuildRequestJsonWithImage(escapedUser, systemPrompt, maintainContext, b64, mt)

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    Dim r As Long, c As Long, s As Long, v As Long
    r = IIf(m_toResolve > 0, m_toResolve, 5000)
    c = IIf(m_toConnect > 0, m_toConnect, 15000)
    s = IIf(m_toSend > 0, m_toSend, 15000)
    v = IIf(m_toReceive > 0, m_toReceive, 60000)
    http.setTimeouts r, c, s, v

    http.Open "POST", BuildUrl(), asyncMode
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "x-goog-api-key", m_apiKey

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

        Dim modelText As String
        modelText = ExtractOutput()
        If LenB(modelText) > 0 Then
            m_historyRoles.Add "model"
            m_historyTexts.Add cjs.EscapeJsonString(modelText)
        End If
    End If

    CreateResponseWithImage = m_lastResponseRaw
    Exit Function

EH:
    m_lastError = "CreateResponseWithImage error: " & Err.Number & " - " & Err.Description
End Function

' New user turn = text part + inline_data part. Logs a sanitized copy with the
' base64 payload truncated.
Private Function BuildRequestJsonWithImage( _
    ByVal escapedUser As String, _
    ByVal systemPrompt As String, _
    ByVal maintainContext As Boolean, _
    ByVal b64 As String, _
    ByVal mimeType As String _
) As String

    Dim json As String
    
    json = "{""contents"":["

    Dim first As Boolean
    first = True

    Dim i As Long
    If maintainContext And Not (m_historyRoles Is Nothing) Then
        For i = 1 To m_historyRoles.Count
            If Not first Then json = json & ","
            json = json & "{""role"":""" & m_historyRoles.Item(i) & """,""parts"":[{""text"":""" & m_historyTexts.Item(i) & """}]}"
            first = False
        Next i
    End If

    If Not first Then json = json & ","
    json = json & "{""role"":""user"",""parts"":["
    json = json & "{""text"":""" & escapedUser & """},"
    json = json & "{""inline_data"":{""mime_type"":""" & mimeType & """,""data"":""" & b64 & """}}"
    json = json & "]}"
    json = json & "]"

    If LenB(systemPrompt) <> 0 Then
        json = json & ",""systemInstruction"":{""parts"":[{""text"":""" & cjs.EscapeJsonString(systemPrompt) & """}]}"
    End If

    json = json & GenerationConfigFragment()

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
    If LenB(CStr(j.getVal("error.message"))) <> 0 Then Exit Sub

    LastInputTokens = CLng(j.getVal("usageMetadata.promptTokenCount"))
    LastOutputTokens = CLng(j.getVal("usageMetadata.candidatesTokenCount"))
    LastCachedTokens = CLng(j.getVal("usageMetadata.cachedContentTokenCount"))

    TotalInputTokens = TotalInputTokens + LastInputTokens
    TotalOutputTokens = TotalOutputTokens + LastOutputTokens
    TotalCachedTokens = TotalCachedTokens + LastCachedTokens
    TotalRequests = TotalRequests + 1

End Sub

