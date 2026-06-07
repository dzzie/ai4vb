VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLamaModel 
      Height          =   285
      Left            =   810
      TabIndex        =   20
      Text            =   "qwen3:30b"
      Top             =   405
      Width           =   1230
   End
   Begin VB.OptionButton optGemini 
      Caption         =   "Gemini"
      Height          =   330
      Left            =   5265
      TabIndex        =   19
      Top             =   405
      Width           =   1050
   End
   Begin VB.OptionButton optClaude 
      Caption         =   "Claude"
      Height          =   330
      Left            =   4275
      TabIndex        =   18
      Top             =   405
      Width           =   1185
   End
   Begin VB.OptionButton optChatGpt 
      Caption         =   "ChatGPT"
      Height          =   285
      Left            =   3240
      TabIndex        =   17
      Top             =   405
      Width           =   1230
   End
   Begin VB.TextBox txtCategory 
      Height          =   285
      Left            =   4050
      TabIndex        =   14
      Top             =   5175
      Width           =   1905
   End
   Begin VB.OptionButton optQwen 
      Caption         =   "lama"
      Height          =   285
      Left            =   135
      TabIndex        =   13
      Top             =   405
      Value           =   -1  'True
      Width           =   645
   End
   Begin VB.TextBox txtQwenIp 
      Height          =   285
      Left            =   2025
      TabIndex        =   12
      Text            =   "192.168.0.51"
      Top             =   405
      Width           =   1140
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   285
      Left            =   7425
      TabIndex        =   10
      Top             =   45
      Width           =   915
   End
   Begin VB.TextBox txtFolder 
      Height          =   330
      Left            =   45
      TabIndex        =   9
      Top             =   0
      Width           =   6585
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   6840
      TabIndex        =   8
      Top             =   45
      Width           =   465
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Classify"
      Height          =   285
      Left            =   6840
      TabIndex        =   7
      Top             =   495
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   8415
      TabIndex        =   6
      Top             =   495
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   285
      Left            =   6210
      TabIndex        =   5
      Top             =   5220
      Width           =   1140
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   45
      TabIndex        =   4
      Top             =   5175
      Width           =   3840
   End
   Begin VB.ListBox lstResults 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   5625
      Width           =   7125
   End
   Begin VB.TextBox txtLog 
      Height          =   3390
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1485
      Width           =   11400
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   7875
      TabIndex        =   0
      Top             =   5625
      Width           =   3750
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   45
      TabIndex        =   3
      Top             =   1125
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "AI Debugout"
      Height          =   285
      Left            =   7965
      TabIndex        =   16
      Top             =   5310
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Catagory"
      Height          =   195
      Left            =   4050
      TabIndex        =   15
      Top             =   4950
      Width           =   645
   End
   Begin VB.Label lblStatus 
      Caption         =   "lblStatus"
      Height          =   240
      Left            =   45
      TabIndex        =   11
      Top             =   855
      Width           =   11490
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cat As CTextCatalog
Attribute cat.VB_VarHelpID = -1
Private lama As New COllama
Private chat As New COpenAI
Private gem As New CGemini
Private claude As New CClaudeAI

Private Sub cmdBrowse_Click()
    Dim f As String
    f = BrowseForFolder(Me.hwnd, "Select the image folder to scan")
    If LenB(f) <> 0 Then txtFolder.text = f
End Sub

Private Sub cmdCancel_Click()
    cat.CancelRequested = True             ' stops after the current file
End Sub

Private Sub Form_Load()
    Set cat = New CTextCatalog
    txtFolder = App.path & "\test"
    cat.OpenCatalog App.path & "\docs.db"
End Sub

Private Sub cmdScan_Click()
    cat.Scan txtFolder.text, True
    lblStatus.Caption = cat.TotalCount & " docs, " & cat.PendingCount & " to extract"
End Sub

Private Sub cmdExtract_Click()
    
    Dim ai As Object
    
    If optChatGpt.value Then
        If Not chat.isApiKeySet Then chat.SaveApiKey
        Set ai = chat
    End If
    
    If optClaude.value Then
        If Not claude.isApiKeySet Then claude.SaveApiKey
        Set ai = claude
    End If
    
    If optGemini.value Then
        If Not gem.isApiKeySet Then gem.SaveApiKey
        Set ai = gem
    End If
    
    If optQwen.value Then
        lama.Model = txtLamaModel
        lama.RemoteIP = txtQwenIp
        Set ai = lama
    End If
    
    List1.Clear
    List1.AddItem "Using " & TypeName(ai)
    
    cat.ClassifyExtract ai                ' resumable; pass maxFiles for a batch
    lblStatus.Caption = cat.DoneCount & " done, " & cat.CoercedCount & " coerced, " & cat.UnparseableCount & " bad"
    ProgressBar1.value = 0
    
End Sub

Private Sub cmdSearch_Click()
    Dim rs As cSQLiteResults
    Set rs = cat.Search(txtSearch.text, txtCategory.text)   ' either arg optional
    lstResults.Clear
    Do While rs.MoveNext
        lstResults.AddItem rs.FieldText("category") & " / " & rs.FieldText("priority") & "  " & rs.FieldText("summary")
    Loop
    rs.CloseRS
End Sub

Private Sub cat_Progress(ByVal phase As String, ByVal done As Long, ByVal total As Long, ByVal currentPath As String)
    Me.Caption = phase & " " & done & "/" & total
    ProgressBar1.value = (done * 100&) \ IIf(total = 0, 1, total)
    DoEvents
End Sub

Private Sub cat_LogMsg(ByVal msg As String)
    txtLog.text = txtLog.text & msg & vbCrLf
End Sub
