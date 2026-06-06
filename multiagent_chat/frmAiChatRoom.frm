VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAiChatRoom 
   Caption         =   "AI Chatroom"
   ClientHeight    =   9495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   420
      Left            =   2430
      TabIndex        =   13
      Top             =   3825
      Width           =   1365
   End
   Begin VB.CommandButton cmdInterject 
      Caption         =   "Interject"
      Height          =   375
      Left            =   7380
      TabIndex        =   12
      Top             =   2610
      Width           =   1320
   End
   Begin VB.TextBox txtInterject 
      Height          =   615
      Left            =   1215
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2520
      Width           =   5940
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   4050
      TabIndex        =   9
      Top             =   3240
      Width           =   2940
   End
   Begin RichTextLib.RichTextBox txtLog 
      Height          =   5145
      Left            =   90
      TabIndex        =   8
      Top             =   4320
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   9075
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAiChatRoom.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7380
      TabIndex        =   5
      Top             =   3735
      Width           =   1335
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   7380
      TabIndex        =   4
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   7425
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtUserQ 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1800
      Width           =   7065
   End
   Begin VB.TextBox txtScenario 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   8640
   End
   Begin VB.Label Label1 
      Caption         =   "User interject"
      Height          =   330
      Left            =   180
      TabIndex        =   10
      Top             =   2565
      Width           =   1410
   End
   Begin VB.Label lblLog 
      Caption         =   "Transcript:"
      Height          =   240
      Left            =   135
      TabIndex        =   6
      Top             =   4005
      Width           =   3000
   End
   Begin VB.Label lblQ 
      Caption         =   "Opening question (USER):"
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   3000
   End
   Begin VB.Label lblScenario 
      Caption         =   "Scenario / ground rules:"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "frmAiChatRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.8
'License:      MIT
'
' Round-robin multi-provider AI chatroom. All providers share one interface
' (CreateResponse / ExtractOutput / LastStatus / LastError / Cancel), so the
' room is just a loop over a collection of agent objects.
'
' KEY DESIGN: a single shared transcript (gTranscript) is the conversation.
' Every agent is fed the WHOLE transcript fresh each turn with
' maintainContext:=False -- we do NOT use the providers' internal per-class
' history, because each keeps its own private thread in its own format and
' they would diverge. The shared transcript is the context; agents are
' stateless responders to it.

Option Explicit

Private gAgents As Collection      ' provider objects (shared interface, late bound)
Private gNames As Collection       ' parallel display names: "AGENT_1", ...
Private gRoles As Collection

Private gTranscript As String      ' the running room conversation
Private gRound As Long
Private gCancel As Boolean
Private gBusy As Boolean
Private gLastUser As String        ' last USER line posted (detect new injections)

Dim fso As New CFileSystem2


' ----------------------------------------------------------------------------
' Roster. Fill in your keys/models. Order here = round-robin order.
' Models are set by STRING (late-bind safe). Add/remove agents freely.
' ----------------------------------------------------------------------------
Private Sub InitAgents()

    Set gAgents = New Collection
    Set gNames = New Collection
    Set gRoles = New Collection
    
    Dim claude As New CClaudeAI
    Dim chat As New COpenAI
    Dim gem As New CGemini
    Dim lama As New COllama
'
    ' --- AGENT_1: Claude ---
    claude.ApiKey = GetSetting("ai4vb", "claude", "key")
    'claude.Model = "claude-opus-4-8"
    claude.SetTimeoutsMs 5000, 15000, 15000, 120000
    AddAgent claude, "Claude", "the PROPONENT -- argue FOR the idea, surface its strongest upside"
    If Len(claude.ApiKey) = 0 Then List1.AddItem "Claude API key not set!"
    
    
    ' --- AGENT_2: Chatgpt ---
    chat.ApiKey = GetSetting("ai4vb", "chatgpt", "key")
    'chat.Model = "gemini-3.5-flash"
    chat.SetTimeoutsMs 5000, 15000, 15000, 120000
    AddAgent chat, "chatGPT", "the SKEPTIC -- attack the weakest part of whatever has been said"
    If Len(chat.ApiKey) = 0 Then List1.AddItem "ChatGPT API key not set!"
        
    ' --- AGENT_2: Gemini ---
    gem.ApiKey = GetSetting("ai4vb", "gemini", "key")
    gem.SetModel gem_2_5_flash_lite
    gem.SetTimeoutsMs 5000, 15000, 15000, 120000
    AddAgent gem, "gemini", "the PRAGMATIST -- only concrete what-would-you-actually-do specifics"
    If Len(gem.ApiKey) = 0 Then List1.AddItem "Gemini API key not set!"

    ' --- AGENT_3: Ollama (local) ---
    lama.RemoteIP = "192.168.0.51"              ' <-- your AI box IP (or omit for localhost)
    lama.Model = "qwen3:30b"
    lama.Think = False
    lama.SetTimeoutsMs 5000, 15000, 15000, 180000
    AddAgent lama, "Ollama", "the AGITATOR -- be contrarian, aggressive, and push hard on weak points, but argue ONLY from real reasoning. Never invent a statistic to win."

End Sub

Private Sub AddAgent(ByVal obj As Object, ByVal nm As String, ByVal role As String)
    gAgents.Add obj
    gNames.Add nm
    gRoles.Add role
End Sub

' ----------------------------------------------------------------------------
' Per-agent system prompt: identity + the three rules that stop a model from
' hallucinating the whole room.
' ----------------------------------------------------------------------------
Private Function BuildSystemPrompt(ByVal myName As String, roleText As String) As String
    Dim s As String
    s = "You are " & myName & ", one of several AI participants in a group chat with a human USER." & vbLf
    s = s & vbLf & "SCENARIO / GROUND RULES:" & vbLf & txtScenario.text & vbLf
    s = s & vbLf & "You are shown the full transcript so far. Add ONE short contribution as " & myName & "." & vbLf
    s = s & "Rules:" & vbLf
    s = s & "- Respond ONLY as " & myName & ". Never write lines for USER or any other AGENT." & vbLf
    s = s & "- Do NOT prefix your reply with '" & myName & " says:'. That label is added for you. Output only your message." & vbLf
    s = s & "- A few sentences. React to what others said; do not repeat points already made." & vbLf
    s = s & vbLf & "Your fixed role in this room is " & roleText & ". Stay in it, every round." & vbLf
    BuildSystemPrompt = s
End Function

' ----------------------------------------------------------------------------
' One full round: every agent speaks once, in order, each seeing everything
' said so far INCLUDING earlier agents this same round.
' ----------------------------------------------------------------------------
Private Sub RunRound()
    If gBusy Then Exit Sub
    If gAgents Is Nothing Then Exit Sub
    If gAgents.Count = 0 Then Exit Sub

    gBusy = True
    gCancel = False
    cmdStart.Enabled = False
    cmdContinue.Enabled = False
    cmdCancel.Enabled = True

    ' Optional human injection: if the user typed something new, post it first.
'    Dim uq As String
'    uq = Trim$(txtUserQ.text)
'    If LenB(uq) <> 0 And uq <> gLastUser Then
'        PostLine "USER", uq
'        gLastUser = uq
'    End If

    Dim inj As String
    inj = Trim$(txtInterject.text)
    If LenB(inj) <> 0 Then
        PostLine "USER", inj
        txtInterject.text = ""
    End If

    gRound = gRound + 1
    LogRaw "----- round " & gRound & " -----"

    Dim i As Long, ai As Object, nm As String, sys As String, reply As String
    For i = 1 To gAgents.Count
        If gCancel Then
            LogRaw "[cancelled]"
            Exit For
        End If

        Set ai = gAgents.Item(i)
        nm = gNames.Item(i)

        Me.Caption = "running " & nm & " ..."
        DoEvents

        sys = BuildSystemPrompt(nm, gRoles.Item(i))

        ' Feed the WHOLE shared transcript. maintainContext = False (each agent
        ' sees the room, not its own private thread). asyncMode = True keeps the
        ' UI alive and lets Cancel work mid-call.
        ai.CreateResponse gTranscript, sys, False, True

        If ai.LastStatus = 200 Then
            reply = Trim$(ai.ExtractOutput)
        Else
            reply = ""
        End If

        If LenB(reply) = 0 Then
            ' Log the failure but DON'T add error noise to the shared transcript.
            LogRaw nm & " (error): " & ai.LastError & " [status " & ai.LastStatus & "]"
        Else
            PostLine nm, reply
        End If
    Next i

    Me.Caption = "AI Chatroom - round " & gRound & " complete"
    gBusy = False
    cmdStart.Enabled = True
    cmdContinue.Enabled = True
    cmdCancel.Enabled = False
End Sub

' Append a line to BOTH the shared transcript (what agents see) and the log.
Private Sub PostLine(ByVal who As String, ByVal text As String)
    Dim ln As String
    ln = who & " says: " & text
    gTranscript = gTranscript & ln & vbCrLf
    LogRaw ln
End Sub

' Append to the visible log and autoscroll to the bottom.
Private Sub LogRaw(ByVal s As String)
    txtLog.text = txtLog.text & s & vbCrLf & vbCrLf
    txtLog.SelStart = Len(txtLog.text)
    txtLog.SelLength = 0
    DoEvents
End Sub

Private Sub cmdCopy_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText txtLog.text
End Sub

Private Sub cmdInterject_Click()
    If gBusy Then Exit Sub
    If LenB(gTranscript) = 0 Then MsgBox "Press Start first.": Exit Sub
    Dim inj As String
    inj = Trim$(txtInterject.text)
    If LenB(inj) = 0 Then Exit Sub
    PostLine "USER", inj
    txtInterject.text = ""
End Sub

' ----------------------------------------------------------------------------
' UI
' ----------------------------------------------------------------------------
Private Sub Form_Load()
    Set form1 = Me 'lazy
    Me.Caption = "AI Chatroom"
    cmdCancel.Enabled = False
    
    Dim p As String
    p = App.path & "\system_prompt.txt"
    If fso.FileExists(p) Then
       txtScenario.text = fso.ReadFile(p)
    Else
        txtScenario.text = "You are experts debating the best approach to a problem. Be concise, disagree where warranted, and build toward a useful answer."
    End If
    
End Sub

Private Sub cmdStart_Click()
    If gBusy Then Exit Sub
    If LenB(Trim$(txtScenario.text)) = 0 Then MsgBox "Enter a scenario.": Exit Sub
    If LenB(Trim$(txtUserQ.text)) = 0 Then MsgBox "Enter an opening question.": Exit Sub

    InitAgents
    If gAgents.Count = 0 Then MsgBox "No agents configured (see InitAgents).": Exit Sub

    txtLog.text = ""
    gTranscript = ""
    gRound = 0
    gLastUser = ""

    PostLine "USER", Trim$(txtUserQ.text)
    gLastUser = Trim$(txtUserQ.text)

    RunRound
End Sub

Private Sub cmdContinue_Click()
    If gBusy Then Exit Sub
    If gAgents Is Nothing Then MsgBox "Press Start first.": Exit Sub
    ' If you type a new line in the question box before Continue, it gets
    ' posted as a fresh USER turn at the top of the round.
    RunRound
End Sub

Private Sub cmdCancel_Click()
    gCancel = True
    ' Nudge whichever agent is mid-call to abort; harmless on the idle ones.
    On Error Resume Next
    Dim i As Long
    For i = 1 To gAgents.Count
        gAgents.Item(i).Cancel
    Next i
End Sub
