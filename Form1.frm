VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Give AI Access to your programs object model to answer questions for you.  http://sandsprite.com"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   465
      Left            =   9270
      TabIndex        =   22
      Top             =   4140
      Width           =   870
   End
   Begin VB.TextBox txtOllamaModel 
      Height          =   330
      Left            =   3420
      TabIndex        =   21
      Text            =   "qwen3:30b"
      Top             =   810
      Width           =   1365
   End
   Begin VB.TextBox txtOllamaIP 
      Height          =   285
      Left            =   1170
      TabIndex        =   18
      Text            =   "192.168.0.51"
      Top             =   855
      Width           =   1320
   End
   Begin VB.OptionButton optOLlama 
      Caption         =   "ollama"
      Height          =   285
      Left            =   3915
      TabIndex        =   17
      Top             =   3330
      Width           =   915
   End
   Begin VB.CheckBox chkAsync 
      Caption         =   "async"
      Height          =   255
      Left            =   900
      TabIndex        =   16
      Top             =   3360
      Value           =   1  'Checked
      Width           =   750
   End
   Begin VB.TextBox txtAgentPrompt 
      Height          =   2025
      Left            =   1125
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "Form1.frx":0000
      Top             =   1260
      Width           =   8235
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   3360
      Width           =   1515
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   3720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text2 
      Height          =   1575
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   9420
      Width           =   8415
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   780
      TabIndex        =   10
      Top             =   7560
      Width           =   8355
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   9540
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      Language        =   "jscript"
   End
   Begin VB.CommandButton cmdAgentTest 
      Caption         =   "Agentic Test"
      Height          =   255
      Left            =   6030
      TabIndex        =   9
      Top             =   3375
      Width           =   1575
   End
   Begin VB.OptionButton optClaude 
      Caption         =   "Claude"
      Height          =   255
      Left            =   2835
      TabIndex        =   8
      Top             =   3375
      Width           =   1335
   End
   Begin VB.OptionButton OptChatGpt 
      Caption         =   "ChatGpt"
      Height          =   255
      Left            =   1755
      TabIndex        =   7
      Top             =   3360
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.CommandButton cmdSetClaudeKey 
      Caption         =   "Set"
      Height          =   315
      Left            =   9300
      TabIndex        =   6
      Top             =   660
      Width           =   915
   End
   Begin VB.TextBox txtClaudeKey 
      Height          =   255
      Left            =   1170
      TabIndex        =   5
      Top             =   540
      Width           =   8055
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   4020
      Width           =   8295
   End
   Begin VB.CommandButton cmdSetApiKey 
      Caption         =   "Set"
      Height          =   315
      Left            =   9300
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox txtApiKey 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   7995
   End
   Begin VB.Label Label2 
      Caption         =   "Model: "
      Height          =   240
      Left            =   2790
      TabIndex        =   20
      Top             =   855
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Ollama IP:"
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   900
      Width           =   960
   End
   Begin VB.Label Label4 
      Caption         =   "Prompt"
      Height          =   255
      Left            =   270
      TabIndex        =   15
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Claude Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lblApiKey 
      Caption         =   "ChatGpt Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuRegen 
         Caption         =   "Regen Prototypes"
      End
      Begin VB.Menu mnuJsontest 
         Caption         =   "Json Test"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:       David Zimmer <dzzie@yahoo.com>
'Ai Assistant: Claude Opus 4.7
'Site:         http://sandsprite.com
'License:      MIT

Option Explicit

Dim lama As New COllama
Dim ai As New COpenAI
Dim claude As New CClaudeAI
Dim mgr As New CManager
Dim log As New CLogger
Dim fso As New CFileSystem2

' Re-entry / cancel flags for the agent loop.
'   m_agentRunning   — True while cmdAgentTest_Click is mid-loop. Blocks
'                      re-entry from another click during DoEvents pumps.
'   m_agentCancelled — set by cmdCancel_Click; checked between stages
'                      and after each HTTP call.
Private m_agentRunning As Boolean
Private m_agentCancelled As Boolean

 
Private Sub cmdCopy_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText txtOut.text
End Sub

 Private Sub Form_Load()
    On Error Resume Next
    txtApiKey = GetSetting("ai4vb", "chatgpt", "key")
    ai.ApiKey = txtApiKey.text
    txtClaudeKey = GetSetting("ai4vb", "claude", "key")
    claude.ApiKey = txtClaudeKey.text

    ' --- build some fake data ---
    mgr.addUser "Roger", "50", "Security Researcher"
    mgr.addUser "Alice", "32", "Malware Analyst"
    mgr.addUser "Bob", "45", "Reverse Engineer"
    mgr.addUser "Carol", "38", "Threat Intelligence Lead"

    ' --- projects: leadName is a soft FK to Users.name ---
    ' Budgets deliberately out-of-order; highest is NOT the first declared.
    ' Carol leads "Sentinel" with the largest budget — the right answer
    ' to "who leads the highest-budget project, and what's their job?"
    ' is Carol / Threat Intelligence Lead.
    mgr.addProject "Phoenix", "Alice", "150000"
    mgr.addProject "Hydra", "Bob", "85000"
    mgr.addProject "Sentinel", "Carol", "420000"
    mgr.addProject "Chimera", "Roger", "210000"

    ' --- setup ScriptControl ---
    sc.Language = "JScript"

End Sub

' Called from JScript via host.answer(x) so AI output surfaces in the UI.
' Returns x so the eval result is non-empty — helps the AI see its own
' output reflected back next turn.
Public Function answer(ByVal x As String) As String
    txtOut.text = txtOut.text & "[answer] " & x & vbCrLf
    log.log "host.answer", x
    answer = x
End Function

' Called from JScript via host.describe(obj). Returns the contents of
' ./protos/<TypeName(obj)>.txt — a generated description of the object's
' public surface. If the proto file is missing, returns a clear error
' string so the AI can see why it failed.
'
' NOTE: parameter typed as Variant rather than Object. MSScriptControl
' marshals JScript values as VARIANTs and VB6 binds these to Variant
' parameters reliably. "As Object" with ByVal sometimes fails to bind
' from JScript with an empty "Line 1:" error.
Public Function describe(ByRef obj As Variant) As String
    Dim tn As String, p As String
    On Error GoTo EH
    tn = TypeName(obj)
    p = App.path & "\protos\" & tn & ".txt"
    If Not fso.FileExists(p) Then
        describe = "ERROR: no proto file for class '" & tn & "' at " & p & _
                  ". Regenerate protos via the 'Regen Protos' button."
        Exit Function
    End If
    describe = fso.ReadFile(p)
    Exit Function
EH:
    describe = "ERROR: describe() failed: " & Err.Number & " " & Err.Description
End Function

Private Sub cmdAgentTest_Click()

    Const MAX_STAGES As Integer = 20    ' << bumped from 10
    Dim stage As Integer
    Dim userMsg As String
    Dim jsScript As String
    Dim result As String
    Dim errText As String
    Dim hadError As Boolean
    Dim lastResult As String
    Dim startTime As Date
    Dim outcome As String

    ' Re-entry guard: if a run is already in progress, don't start another.
    ' DoEvents in the polling loop means the button could fire again
    ' before the first run finishes.
    If m_agentRunning Then Exit Sub
    m_agentRunning = True
    m_agentCancelled = False

    ' UI: lock the agent button, unlock cancel, prime the progress bar.
    cmdAgentTest.Enabled = False
    cmdCancel.Enabled = True
    ProgressBar1.Min = 0
    ProgressBar1.Max = MAX_STAGES
    ProgressBar1.value = 0

    startTime = Now
    Me.Caption = "running " & Format$(startTime, "hh:nn:ss")
    outcome = "incomplete"

    ' --- open log file fresh for this run ---
    log.OpenLog App.path & "\agent.log", False    ' append mode
    log.Section "AGENT RUN START"

    ' --- ScriptControl setup happens per-stage now (see below).
    ' --- Each stage gets a fresh JScript engine, so variables from one
    '     stage cannot leak into the next. This trades cross-stage memo
    '     for hermetic, debuggable scripts.

    ' --- build the system prompt ---
    Dim sysPrompt As String, p As String

    p = App.path & "\prompt.txt"
    If Not fso.FileExists(p) Then
        MsgBox "sysPrompt not found: " & p
        log.log "FATAL", "prompt.txt missing at " & p
        log.CloseLog
        Exit Sub
    End If

    sysPrompt = fso.ReadFile(p)
    log.log "sysPrompt", sysPrompt

    ' --- set up AI: honor the radio buttons (OptChatGpt vs Option2/Claude) ---
    Dim agentAI As Object     'we hold this late bound so we can switch between the two arbitrarily
    Dim agentLabel As String  'they have the same interface on purpose, downside no intellisense
    
    If OptChatGpt.value Then
        Set agentAI = ai
        agentLabel = "ChatGPT"
    ElseIf optClaude.value Then
        Set agentAI = claude
        agentLabel = "Claude"
    Else
        Set agentAI = lama
        agentLabel = "Ollama"
        lama.RemoteIP = txtOllamaIP
        lama.Model = txtOllamaModel
        lama.Think = True
    End If
    
    agentAI.ResetContext   ' new conversation chain per run
    log.log "agent backend", agentLabel & " (model: " & agentAI.Model & ")"

    userMsg = txtAgentPrompt
    log.log "initial task", userMsg

    List1.Clear
    txtOut.text = Empty
    hadError = False
    lastResult = ""

    For stage = 1 To MAX_STAGES

        log.Section "STAGE " & stage
        List1.AddItem "--- Stage " & stage & " ---"
        ProgressBar1.value = stage
        Me.Caption = "running stage " & stage & "/" & MAX_STAGES & " (" & _
                     DateDiff("s", startTime, Now) & "s)"

        ' Inter-stage cancel check. The polling loop inside CreateResponse
        ' also checks the cancel flag, but if the user clicks Cancel between
        ' stages (after a fast eval, before the next HTTP call), we catch it
        ' here too.
        If m_agentCancelled Then
            log.log "CANCELLED", "user cancelled before stage " & stage
            outcome = "cancelled at stage " & stage
            Exit For
        End If

        ' Build message: on error, send error + the bad script; otherwise
        ' send last result. After stage 1 we always rebuild userMsg here.
        If stage > 1 Then
            If hadError Then
                userMsg = "Your previous JScript produced an error." & vbCrLf & _
                          "ERROR: " & errText & vbCrLf & _
                          "SCRIPT THAT FAILED:" & vbCrLf & jsScript & vbCrLf & _
                          "Return corrected JScript. Do not repeat the same script."
                hadError = False
            Else
                userMsg = "Result of your last script: [" & lastResult & "]" & vbCrLf & _
                          "If the task is fully answered, return DONE. " & _
                          "Otherwise emit the next JScript."
            End If
        End If

        log.log "user -> AI", userMsg

        ' Call AI
        ' maintainContext=True so the AI sees the full conversation chain.
        ' if asyncMode=True so the polling loop pumps DoEvents and the user can
        ' click Cancel; we've already guarded against re-entrancy with m_agentRunning above.
        agentAI.CreateResponse userMsg, sysPrompt, True, (chkAsync.value = 1)
        log.log "AI HTTP status", CStr(agentAI.LastStatus) & " " & agentAI.LastStatusText

        If agentAI.LastStatus <> 200 Then
            If m_agentCancelled Then
                log.log "CANCELLED", "user cancelled during HTTP at stage " & stage
                outcome = "cancelled at stage " & stage
            Else
                List1.AddItem "HTTP Error " & agentAI.LastStatus & ": " & agentAI.LastStatusText
                log.log "AI raw response", agentAI.LastResponseRaw
                log.log "ABORT", "non-200 status, exiting loop"
                outcome = "HTTP error " & agentAI.LastStatus
            End If
            Exit For
        End If

        jsScript = Trim$(agentAI.ExtractOutput())
        log.log "AI -> script", jsScript
  
        List1.AddItem "AI JS: " & Left$(jsScript, 80) & IIf(Len(jsScript) > 80, "...", "")

        ' Check for completion signal
        If jsScript = "DONE" Or jsScript = "" Then
            List1.AddItem "(AI signalled completion)"
            log.log "AI signal", "DONE (empty or literal)"
            outcome = "DONE in " & stage & " stages"
            Exit For
        End If

        ' --- Execute the JS ---
        ' Reset the engine so prior-stage variables can't leak into this
        ' script. sc.Reset clears the language too, so we set it again,
        ' then re-bind the objects the AI is allowed to call.
        On Error Resume Next
        sc.Reset
        sc.Language = "JScript"
        sc.AddObject "manager", mgr, True
        sc.AddObject "host", Me, True
        Err.Clear
        sc.Error.Clear    ' belt-and-braces: clear any prior sc error too
        result = CStr(sc.Eval(unixToDOS(jsScript)))
        errText = ""

        ' Check ScriptControl-level error first
        If sc.Error.Number <> 0 Then
            errText = "Line " & sc.Error.Line & ": " & sc.Error.text
        ElseIf Err.Number <> 0 Then
            errText = Err.Description
        End If
        On Error GoTo 0

        If errText <> "" Then
            List1.AddItem "  ERROR: " & errText
            log.log "JS ERROR", errText
            hadError = True
        Else
            List1.AddItem "  Result: " & result
            log.log "JS result", result
            txtOut.text = txtOut.text & "Stage " & stage & ": " & result & vbCrLf
            lastResult = result
            
            
            If optOLlama.value Then
               txtOut.text = txtOut.text & "OLLAMA REASONING: " & lama.LastThinking
               log.log "OLLAMA REASONING: ", lama.LastThinking
            End If
          
        End If

        log.Flush   ' so you can tail agent.log while it runs

        DoEvents

    Next stage

    If stage > MAX_STAGES Then
        List1.AddItem "(MAX_STAGES reached without DONE)"
        log.log "OUTCOME", "max stages reached, no DONE"
        outcome = "max stages, no DONE"
    End If

    Dim elapsed As Long
    elapsed = DateDiff("s", startTime, Now)
    Me.Caption = "complete secs: " & elapsed & " (" & outcome & ")"
    log.log "ELAPSED", elapsed & "s (" & outcome & ")"

    List1.AddItem "=== Complete (" & elapsed & "s, " & outcome & ") ==="
    log.Section "AGENT RUN END"
    log.CloseLog

    ProgressBar1.value = 0
    cmdAgentTest.Enabled = True
    cmdCancel.Enabled = False
    m_agentRunning = False

End Sub

' Cancel handler. Sets two flags: one on the AI class (so its polling
' loop bails out of the in-flight HTTP request) and one on the form
' (so the stage loop bails before starting the next stage).
Private Sub cmdCancel_Click()
    m_agentCancelled = True
    ' Cancel both possible backends — only one is actually running but
    ' calling Cancel on an idle one is a no-op.
    On Error Resume Next
    ai.Cancel
    claude.Cancel
    On Error GoTo 0
End Sub

Private Sub cmdSetApiKey_Click()
    SaveSetting "ai4vb", "chatgpt", "key", txtApiKey
    ai.ApiKey = txtApiKey.text
End Sub

Private Sub cmdSetClaudeKey_Click()
    ' Was saving txtApiKey (ChatGPT box) into the claude slot — fixed.
    SaveSetting "ai4vb", "claude", "key", txtClaudeKey
    claude.ApiKey = txtClaudeKey.text
End Sub
 

Function alert(x)
    MsgBox x
End Function

Private Sub Form_Unload(Cancel As Integer)
     SaveSetting "ai4vb", "chatgpt", "key", txtApiKey
     SaveSetting "ai4vb", "claude", "key", txtClaudeKey
End Sub

Private Sub Label3_Click()
    claude.GetApiKey
End Sub

Private Sub lblApiKey_Click()
    ai.GetApiKey
End Sub

Function unixToDOS(ByVal tmp As String)
    Dim isMixed As Boolean
    isMixed = (InStr(tmp, vbCrLf) > 0)
    If isMixed Then tmp = VBA.Replace(tmp, vbCrLf, Chr(5))
    tmp = VBA.Replace(tmp, vbLf, vbCrLf)
    If isMixed Then tmp = VBA.Replace(tmp, Chr(5), vbCrLf)
    unixToDOS = tmp
End Function

Private Sub List1_Click()
    On Error Resume Next
    Dim s As String, js As New CJSON
    s = List1.List(List1.ListIndex)
    Text2.text = js.beautify(s)
End Sub

Private Sub mnuJsontest_Click()
    Dim js As String
    js = fso.ReadFile(App.path & "\output.txt")
    txtOut.text = ai.ExtractOutput(js)
End Sub

Private Sub mnuRegen_Click()
     modProtoGen.RegenerateProtos App.path & "\Project1.vbp"
End Sub
