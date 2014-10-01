VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnTutup 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3600
      Picture         =   "Login.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Tutup"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton BtnLogin 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2760
      Picture         =   "Login.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Login"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox TxtPass 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "12345"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox TxtUser 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "fadili"
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Login Acsess"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4920
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    cNDb
End Sub

Private Sub BtnLogin_Click()
    If TxtUser.Text = "" Then
        MsgBox "Username Masih kosong"
        TxtUser.SetFocus
    ElseIf TxtPass.Text = "" Then
        MsgBox "Password Masih kosong"
        TxtPass.SetFocus
    Else
        MySql = "SELECT user_name, password, level FROM login " & _
        "WHERE user_name = '" & TxtUser.Text & "' and password = '" & TxtPass.Text & "'"
        Set SdR = ConN.Execute(MySql)
        If Not SdR.BOF Then
            
            MenuUtama.StatusBar1.Panels.Item(2).Text = SdR.Fields(0)
            MenuUtama.StatusBar1.Panels.Item(4).Text = SdR.Fields(2)
            If MenuUtama.StatusBar1.Panels.Item(4).Text = "Admin" Then
                With MenuUtama
                    .Toolbar1.Buttons(1).Enabled = False
                    .Toolbar1.Buttons(2).Enabled = True
                    .Toolbar1.Buttons(4).Enabled = True
                    .Toolbar1.Buttons(5).Enabled = True
                    .Toolbar1.Buttons(7).Enabled = True
        
                    .MnOperator.Enabled = True
                    .MnPenilaian.Enabled = True
                    .RptDataPegawai.Enabled = True
                    .MnDataPegawai.Enabled = True
                End With
                Unload Me
            Else
                With MenuUtama
                    .Toolbar1.Buttons(1).Enabled = False
                    .Toolbar1.Buttons(2).Enabled = True
                    .Toolbar1.Buttons(4).Enabled = False
                    .Toolbar1.Buttons(5).Enabled = False
                    .Toolbar1.Buttons(7).Enabled = False
        
                    .MnOperator.Enabled = False
                    .MnPenilaian.Enabled = False
                    .RptDataPegawai.Enabled = True
                    .RptPenilaian.Enabled = True
                    .MnDataPegawai.Enabled = False
                End With
                Unload Me
            End If
            
        Else
            MsgBox "Username dn Password Anda Salah"
            TxtUser.SetFocus
        End If
    End If
End Sub

Private Sub BtnTutup_Click()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.WindowState <> 0 Then Exit Sub

        Static l As Integer, t As Integer
        If Button = 1 Then
        Me.Left = (Me.Left + X) - l
        Me.Top = (Me.Top + Y) - t
        Else
        l = X
        t = Y
    End If
End Sub

