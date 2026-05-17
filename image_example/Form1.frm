VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDebug 
      Height          =   2775
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   5580
      Width           =   8055
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   1080
      TabIndex        =   15
      Top             =   4440
      Width           =   8055
   End
   Begin VB.OptionButton optClaude 
      Caption         =   "Claude"
      Height          =   255
      Left            =   9360
      TabIndex        =   14
      Top             =   1740
      Width           =   1155
   End
   Begin VB.OptionButton optOpenAi 
      Caption         =   "OpenAi"
      Height          =   255
      Left            =   9360
      TabIndex        =   13
      Top             =   1380
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtOut 
      Height          =   1755
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2520
      Width           =   8115
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   315
      Left            =   9360
      TabIndex        =   10
      Top             =   2100
      Width           =   1155
   End
   Begin VB.TextBox txtReq 
      Height          =   1095
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1320
      Width           =   8115
   End
   Begin VB.TextBox txtImage 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   900
      Width           =   8115
   End
   Begin VB.TextBox txtApiKey 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   0
      Width           =   7995
   End
   Begin VB.CommandButton cmdSetApiKey 
      Caption         =   "Set"
      Height          =   315
      Left            =   9180
      TabIndex        =   2
      Top             =   0
      Width           =   915
   End
   Begin VB.TextBox txtClaudeKey 
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   8055
   End
   Begin VB.CommandButton cmdSetClaudeKey 
      Caption         =   "Set"
      Height          =   315
      Left            =   9180
      TabIndex        =   0
      Top             =   540
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "Debug Log"
      Height          =   435
      Left            =   120
      TabIndex        =   16
      Top             =   4500
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "Response"
      Height          =   435
      Left            =   120
      TabIndex        =   11
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Request"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Image"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   675
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
      Left            =   0
      TabIndex        =   5
      Top             =   60
      Width           =   1035
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
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ai As New COpenAI
Dim claude As New CClaudeAI
Dim log As New CLogger
Dim fso As New CFileSystem2

Private Sub cmdSubmit_Click()
    
    Dim o As Object
    Dim startTime As Date
    
    List1.Clear
    txtOut.text = Empty
    
    If optOpenAi.value Then Set o = ai Else Set o = claude
    
    If Not fso.FileExists(txtImage) Then
        MsgBox "image not found"
        Exit Sub
    End If
    
    startTime = Now
    Me.Caption = IIf(optOpenAi.value, "OpenAI", "Claude") & " running " & Format$(startTime, "hh:nn:ss")
    
    o.CreateResponseWithImage txtReq, txtImage
    txtOut = unixToDOS(o.ExtractOutput())
    
    Me.Caption = "complete secs: " & DateDiff("s", startTime, Now) & _
             " · " & Format$(o.LastInputTokens, "#,##0") & _
             "/" & Format$(o.LastOutputTokens, "#,##0") & " tok  Total:" & _
             o.TotalOutputTokens
             
End Sub

Private Sub lblApiKey_Click()
    ai.GetApiKey
End Sub

Private Sub Label3_Click()
    claude.GetApiKey
End Sub

Private Sub cmdSetApiKey_Click()
    SaveSetting "ai4vb", "chatgpt", "key", txtApiKey
    ai.ApiKey = txtApiKey.text
End Sub

Private Sub cmdSetClaudeKey_Click()
    SaveSetting "ai4vb", "claude", "key", txtClaudeKey
    claude.ApiKey = txtClaudeKey.text
End Sub

Private Sub Form_Load()
    On Error Resume Next
    txtImage = App.path & "\image1.jpg"
    txtReq = "what is this? what color is it? Should I eat it if a witch gave it to me?"
    txtApiKey = GetSetting("ai4vb", "chatgpt", "key")
    ai.ApiKey = txtApiKey.text
    txtClaudeKey = GetSetting("ai4vb", "claude", "key")
    claude.ApiKey = txtClaudeKey.text
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
    txtDebug = unixToDOS(List1.List(List1.ListIndex))
End Sub
