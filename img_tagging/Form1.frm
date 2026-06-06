VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optGemini 
      Caption         =   "Gemini"
      Height          =   195
      Left            =   3645
      TabIndex        =   14
      Top             =   630
      Width           =   780
   End
   Begin VB.TextBox txtQwenIp 
      Height          =   285
      Left            =   1890
      TabIndex        =   13
      Text            =   "192.168.0.51"
      Top             =   585
      Width           =   1545
   End
   Begin VB.OptionButton optQwen 
      Caption         =   "qwen2.5v1:7b"
      Height          =   285
      Left            =   405
      TabIndex        =   12
      Top             =   585
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   7875
      TabIndex        =   11
      Top             =   5265
      Width           =   3750
   End
   Begin VB.TextBox txtLog 
      Height          =   3390
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   1620
      Width           =   11400
   End
   Begin VB.ListBox lstResults 
      Height          =   1815
      Left            =   45
      TabIndex        =   9
      Top             =   5760
      Width           =   7125
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   1260
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   5310
      Width           =   5865
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   285
      Left            =   6255
      TabIndex        =   6
      Top             =   5355
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   8460
      TabIndex        =   5
      Top             =   630
      Width           =   1095
   End
   Begin VB.CommandButton cmdClassify 
      Caption         =   "Classify"
      Height          =   285
      Left            =   6885
      TabIndex        =   4
      Top             =   630
      Width           =   1005
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   6885
      TabIndex        =   2
      Top             =   180
      Width           =   465
   End
   Begin VB.TextBox txtFolder 
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   6585
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   285
      Left            =   7470
      TabIndex        =   0
      Top             =   180
      Width           =   915
   End
   Begin VB.Label lblStatus 
      Caption         =   "lblStatus"
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   990
      Width           =   11490
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents cat As CImageCatalog
Attribute cat.VB_VarHelpID = -1

Private Sub cmdBrowse_Click()
    Dim f As String
    f = BrowseForFolder(Me.hwnd, "Select the image folder to scan")
    If LenB(f) <> 0 Then txtFolder.Text = f
End Sub

Private Sub Form_Load()

    'Debug.Print FileMD5(App.path & "\Project1.vbp")
    'End
    
    Set cat = New CImageCatalog
    cat.OpenCatalog App.path & "\images.db"
    txtFolder = App.path & "\test"
        
End Sub

Private Sub cmdScan_Click()
    cat.Scan txtFolder.Text, True          ' recursive
    lblStatus.Caption = cat.TotalCount & " imgs, " & cat.PendingCount & " to tag"
End Sub

Private Sub cmdClassify_Click()

    Dim lama As New COllama
    Dim gem As New CGemini
    Dim ai As Object
    Dim k As String
    
    List1.Clear
    
    If optQwen.value Then
        ' Vision model. LOCAL = private (images never leave the box) -- your use case.
        lama.RemoteIP = txtQwenIp              ' or omit for localhost
        lama.Model = "qwen2.5vl:7b"            ' ollama pull qwen2.5vl:32b  (vision-capable)
        lama.SetTimeoutsMs 5000, 15000, 15000, 120000
        Set ai = lama
        List1.AddItem "Using ollama on: " & txtQwenIp
    Else
        If Not gem.isApiKeySet() Then
           k = InputBox("Enter your gemini Api key (saved it to reg)", , Clipboard.GetText)
           If Len(k) = 0 Then Exit Sub
           gem.SaveApiKey k
        End If
        gem.SetTimeoutsMs 5000, 15000, 15000, 120000
        Set ai = gem
        List1.AddItem "Using Gemini"
    End If
    
    cat.ClassifyPending ai                   ' resumable; pass maxFiles to do a batch
    lblStatus.Caption = cat.TaggedCount & " tagged, " & cat.PendingCount & " left"
    ProgressBar1.value = 0
    
End Sub

Private Sub cmdCancel_Click()
    cat.CancelRequested = True             ' stops after the current file
End Sub

Private Sub cmdSearch_Click()
    Dim rs As cSQLiteResults
    Set rs = cat.Search(txtSearch.Text)    ' "cat outdoor" = both tags required
    lstResults.Clear
    Do While rs.MoveNext
        lstResults.AddItem rs.FieldText("path") & "  [" & Trim$(rs.FieldText("tags")) & "]"
    Loop
    rs.CloseRS
End Sub

' --- events from the catalog ---
Private Sub cat_Progress(ByVal phase As String, ByVal done As Long, ByVal total As Long, ByVal currentPath As String)
    Me.Caption = phase & " " & done & "/" & total
    ProgressBar1.value = (done * 100&) \ IIf(total = 0, 1, total)
    DoEvents                               ' lets the UI repaint AND Cancel register
End Sub

Private Sub cat_LogMsg(ByVal msg As String)
    txtLog.Text = txtLog.Text & msg & vbCrLf
End Sub
