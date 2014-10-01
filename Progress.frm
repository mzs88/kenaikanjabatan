VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Progress 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4680
      Top             =   240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Proses . . . . ."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    Me.ProgressBar1.Move 120, Me.ProgressBar1.Top, Me.ScaleWidth - 240
End Sub

Private Sub Timer1_Timer()
Me.Show
ProgressBar1.Value = ProgressBar1.Value + 10
If ProgressBar1.Value = 50 Then
ProgressBar1.Value = ProgressBar1 + 50
If ProgressBar1.Value >= ProgressBar1.Max Then
Timer1.Enabled = False
Unload Me
End If
End If
End Sub
