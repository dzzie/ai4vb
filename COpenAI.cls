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
 
Public LastInputTokens As Long
Public LastOutputTokens As Long
Public LastCachedTokens As Long
Public LastReasoningTokens As Long

Public TotalInputTokens As Long
Public TotalOutputTokens As Long
Public TotalCachedTokens As Long
Public TotalReasoningTokens As Long
Public TotalRequests As Long

Private cjs As New CJSON
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

' Cancel flag â€” set by Cancel() from the UI, checked by the polling loop
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

Private Declare Function CryptBinaryToStringA Lib "crypt32.dll" ( _
    ByRef pbBinary As Byte, _
    ByVal cbBinary As Long, _
    ByVal dwFlags As Long, _
    ByVal pszString As String, _
    ByRef pcchString As Long _
) As Long

' CRYPT_STRING_BASE64    = 1   (with CR/LF line breaks every 64 chars)
' CRYPT_STRING_NOCRLF    = &H40000000  (suppress line breaks)
Private Const CRYPT_STRING_BASE64 As Long = 1
Private Const CRYPT_STRING_NOCRLF As Long = &H40000000


Public Sub ResetTokenTotals()
    TotalInputTokens = 0
    TotalOutputTokens = 0
    TotalCachedTokens = 0
    TotalReasoningTokens = 0
    TotalRequests = 0
End Sub

Public Function GetApiKey()
     ShellExecute 0, "open", "https://platform.openai.com/api-keys", vbNullString, vbNullString, 1
End Function

Function isApiKeySet() As Boolean
    isApiKeySet = (Len(m_apiKey) > 0)
End Function

Function LoadSavedApiKey() As String
    m_apiKey = GetSetting("ai4vb", "chatgpt", "key")
    LoadSavedApiKey = m_apiKey
End Function

Function SaveApiKey(Optional ByVal key As String)
    If Len(key) = 0 Then
        If Len(m_apiKey) = 0 Then
            key = InputBox("Enter Chatgpt API Key")
        Else
            key = m_apiKey
        End If
    End If
    Call SaveSetting("ai4vb", "chatgpt", "key", key)
    m_apiKey = key
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
'   False (default) â€” synchronous: http.send() blocks until response.
'                     Safe for non-UI callers. Cancel() is a no-op in
'                     this mode (nothing to abort between).
'   True            â€” asynchronous: http.send() returns immediately,
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
    AccumulateUsage
    
    ' Always capture the response id so the next chained call can use it
    m_lastResponseId = ParseResponseId(m_lastResponseRaw)

    CreateResponse = m_lastResponseRaw
    Form1.List1.AddItem m_lastResponseRaw 'for early testing only
    Exit Function

EH:
    m_lastError = "CreateResponse error: " & Err.Number & " - " & Err.Description
End Function

' Same semantics as CreateResponse, but the user turn carries an image
' alongside the text. The image is sent inline as a base64 data URI in
' an input_image content part. For continuity, OpenAI's previous_response_id
' carries the image forward server-side — re-sending the bytes on subsequent
' chained turns is NOT required.
'
' detail: "low" | "high" | "auto" (default). "low" downsamples to ~512px
' for cheaper, faster vision calls when fine detail isn't needed.
Public Function CreateResponseWithImage( _
    ByVal inputText As String, _
    ByVal imagePath As String, _
    Optional ByVal instructions As String = "", _
    Optional ByVal maintainContext As Boolean = False, _
    Optional ByVal asyncMode As Boolean = False, _
    Optional ByVal detail As String = "auto" _
) As String

    ClearLast
    m_cancelRequested = False

    If LenB(m_apiKey) = 0 Then
        m_lastError = "ApiKey not set."
        Exit Function
    End If

    Dim b64 As String
    b64 = Base64EncodeFile(imagePath)
    If LenB(b64) = 0 Then
        ' Base64EncodeFile already populated m_lastError
        Exit Function
    End If

    Dim body As String
    body = BuildRequestJsonWithImage(inputText, instructions, maintainContext, _
                                     b64, MediaTypeForPath(imagePath), detail)

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
    http.setRequestHeader "Authorization", "Bearer " & m_apiKey

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
    m_lastResponseId = ParseResponseId(m_lastResponseRaw)
    AccumulateUsage
    
    CreateResponseWithImage = m_lastResponseRaw
    Form1.List1.AddItem m_lastResponseRaw    ' early-testing log, mirrors CreateResponse
    Exit Function

EH:
    m_lastError = "CreateResponseWithImage error: " & Err.Number & " - " & Err.Description
End Function


' Builds the /v1/responses body with a multipart input message.
' Logs a sanitized copy to Form1.List1 — the base64 payload is replaced
' with "<first10>..." so the listbox doesn't get hosed by a multi-MB blob.
Private Function BuildRequestJsonWithImage( _
    ByVal inputText As String, _
    ByVal instructions As String, _
    ByVal maintainContext As Boolean, _
    ByVal b64 As String, _
    ByVal mediaType As String, _
    ByVal detail As String _
) As String

    Dim json As String
 
    json = "{""model"":""" & cjs.EscapeJsonString(Me.Model) & """"
    json = json & ",""input"":[{""role"":""user"",""content"":["
    json = json & "{""type"":""input_text"",""text"":""" & cjs.EscapeJsonString(inputText) & """},"
    json = json & "{""type"":""input_image"",""image_url"":""data:" & mediaType & ";base64," & b64 & """"
    If LenB(detail) <> 0 And detail <> "auto" Then
        json = json & ",""detail"":""" & detail & """"
    End If
    json = json & "}]}]"
    If LenB(instructions) <> 0 Then
        json = json & ",""instructions"":""" & cjs.EscapeJsonString(instructions) & """"
    End If
    If maintainContext And LenB(m_lastResponseId) <> 0 Then
        json = json & ",""previous_response_id"":""" & cjs.EscapeJsonString(m_lastResponseId) & """"
    End If
    json = json & "}"
    BuildRequestJsonWithImage = json

    ' Sanitized log copy: replace the actual base64 payload with a short
    ' preview so the listbox stays usable.
    Dim preview As String
    If Len(b64) > 10 Then
        preview = Left$(b64, 10) & "..."
    Else
        preview = b64
    End If
    Form1.List1.AddItem Replace$(json, b64, preview)

End Function

' Called from the UI (Cancel button) to abort an in-flight CreateResponse
' on the next poll cycle. Safe to call when no request is in flight â€” the
' flag is reset at the start of each CreateResponse.
Public Sub Cancel()
    m_cancelRequested = True
End Sub

Private Function BuildRequestJson(ByVal inputText As String, ByVal instructions As String, ByVal maintainContext As Boolean) As String
    Dim json As String
    
    json = "{""model"":""" & cjs.EscapeJsonString(Me.Model) & """,""input"":""" & cjs.EscapeJsonString(inputText) & """"
    If LenB(instructions) <> 0 Then
        json = json & ",""instructions"":""" & cjs.EscapeJsonString(instructions) & """"
    End If
    If maintainContext And LenB(m_lastResponseId) <> 0 Then
        json = json & ",""previous_response_id"":""" & cjs.EscapeJsonString(m_lastResponseId) & """"
    End If
    json = json & "}"
    BuildRequestJson = json
    Form1.List1.AddItem json 'jsut for testing...
End Function

' Lightweight id extractor — avoids spinning up CJSON just for one field.
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

'Private Function EscapeJsonString(ByVal s As String) As String
'    Dim t As String
'    t = s
'    t = Replace$(t, "\", "\\")
'    t = Replace$(t, """", "\""")
'    t = Replace$(t, vbCrLf, "\n")
'    t = Replace$(t, vbCr, "\n")
'    t = Replace$(t, vbLf, "\n")
'    t = Replace$(t, vbTab, "\t")
'    t = Replace$(t, Chr$(0), "")
'    EscapeJsonString = t
'End Function

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
    LoadSavedApiKey
End Sub

' Reads a file into a byte array and base64-encodes it via crypt32.
' Returns the encoded string with no line breaks.
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

    ' Two-call pattern: first call with NULL string to learn the required
    ' buffer length, then second call to actually fill it. CryptBinaryToStringA
    ' returns the length INCLUDING the null terminator, so we trim one off.
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

    ' needed now reflects chars written, NOT including the trailing null.
    ' (On the second call crypt32 updates pcchString to the actual length.)
    Base64EncodeFile = Left$(buf, needed)
    Exit Function

EH:
    If hFile <> 0 Then Close #hFile
    m_lastError = "Base64EncodeFile error: " & Err.Number & " - " & Err.Description
End Function

Private Function MediaTypeForPath(ByVal path As String) As String
    Dim ext As String
    Dim dotPos As Long
    dotPos = InStrRev(path, ".")
    If dotPos = 0 Then
        MediaTypeForPath = "image/png"   ' fallback
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

Private Sub AccumulateUsage()

    If m_lastStatus <> 200 Then Exit Sub
    If LenB(m_lastResponseRaw) = 0 Then Exit Sub

    Dim j As New CJSON
    If Not j.Initilized Then Exit Sub
    If Not j.loadJson(m_lastResponseRaw) Then Exit Sub
    If LenB(CStr(j.getVal("error.type"))) <> 0 Then Exit Sub

    LastInputTokens = CLng(j.getVal("usage.input_tokens"))
    LastOutputTokens = CLng(j.getVal("usage.output_tokens"))
    LastCachedTokens = 0
    LastReasoningTokens = 0
    If j.exists("usage.input_tokens_details.cached_tokens") Then
        LastCachedTokens = CLng(j.getVal("usage.input_tokens_details.cached_tokens"))
    End If
    If j.exists("usage.output_tokens_details.reasoning_tokens") Then
        LastReasoningTokens = CLng(j.getVal("usage.output_tokens_details.reasoning_tokens"))
    End If

    TotalInputTokens = TotalInputTokens + LastInputTokens
    TotalOutputTokens = TotalOutputTokens + LastOutputTokens
    TotalCachedTokens = TotalCachedTokens + LastCachedTokens
    TotalReasoningTokens = TotalReasoningTokens + LastReasoningTokens
    TotalRequests = TotalRequests + 1

End Sub


