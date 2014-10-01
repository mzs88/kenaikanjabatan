VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MenuUtama 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Menu Utama"
   ClientHeight    =   6720
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11985
   Icon            =   "MenuUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   11925
      TabIndex        =   0
      Top             =   630
      Visible         =   0   'False
      Width           =   11985
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   240
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   2
         Top             =   120
         Width           =   960
      End
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   2040
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   1
         Top             =   600
         Width           =   4095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   2760
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MenuUtama.frx":0FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MenuUtama.frx":1F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MenuUtama.frx":2F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MenuUtama.frx":3EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MenuUtama.frx":4E6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6465
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Picture         =   "MenuUtama.frx":5404
            Text            =   "Username :"
            TextSave        =   "Username :"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Uname"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Picture         =   "MenuUtama.frx":5E16
            Text            =   "Status :"
            TextSave        =   "Status :"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Status"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Picture         =   "MenuUtama.frx":6828
            Text            =   "Tanggal :"
            TextSave        =   "Tanggal :"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Tgl"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Picture         =   "MenuUtama.frx":723A
            Text            =   "Jam :"
            TextSave        =   "Jam :"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Jam"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   33020
            Picture         =   "MenuUtama.frx":7C4C
            Text            =   $"MenuUtama.frx":865E
            TextSave        =   $"MenuUtama.frx":8710
            Object.ToolTipText     =   $"MenuUtama.frx":87C2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Li"
            Object.ToolTipText     =   "Login"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Lg"
            Object.ToolTipText     =   "Logout"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Dp"
            Object.ToolTipText     =   "Data Pegawai"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DoP"
            Object.ToolTipText     =   "Data Operator"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PnL"
            Object.ToolTipText     =   "Penilaian"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "File"
   End
   Begin VB.Menu MnData 
      Caption         =   "Data"
      Begin VB.Menu MnDataPegawai 
         Caption         =   "Data Pegawai"
      End
      Begin VB.Menu MnOperator 
         Caption         =   "Data Operator"
      End
      Begin VB.Menu MnPenilaian 
         Caption         =   "Perhitungan"
      End
   End
   Begin VB.Menu MnLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu RptDataPegawai 
         Caption         =   "Laporan Data Pegawai"
      End
      Begin VB.Menu RptPenilaian 
         Caption         =   "Laporan Penilaian"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "MenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Sub MDIForm_Resize()
Dim client_rect As RECT
Dim client_hwnd As Long

    picStretched.Move 0, 0, _
        ScaleWidth, ScaleHeight

    picStretched.PaintPicture _
        picOriginal.Picture, _
        0, 0, _
        picStretched.ScaleWidth, _
        picStretched.ScaleHeight, _
        0, 0, _
        picOriginal.ScaleWidth, _
        picOriginal.ScaleHeight

    Picture = picStretched.Image

    client_hwnd = FindWindowEx(Me.hwnd, 0, "MDIClient", vbNullChar)
    GetClientRect client_hwnd, client_rect
    InvalidateRect client_hwnd, client_rect, 1
End Sub

Private Sub MDIForm_Load()
    With Toolbar1
        .Buttons(2).Enabled = False
        .Buttons(4).Enabled = False
        .Buttons(5).Enabled = False
        .Buttons(7).Enabled = False
    End With
    
    MnOperator.Enabled = False
    MnPenilaian.Enabled = False
    RptDataPegawai.Enabled = False
    MnDataPegawai.Enabled = False
    StatusBar1.Panels.Item(6).Text = Date
   
    picOriginal.Picture = LoadPicture(App.Path & "\Image\41telkom.jpg")
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, _
UnloadMode As Integer)
    Cancel = 1
    pesan = MsgBox("Keluar?", _
    vbExclamation + vbYesNo + vbDefaultButton2)
    If pesan = vbYes Then
        Cancel = 0
        End
    Else
        Cancel = 1
    End If
End Sub


Private Sub MnDataPegawai_Click()
    Load DataPegawai
    DataPegawai.Show
    DataPegawai.ZOrder
End Sub

Private Sub MnLogin_Click()
    Login.Show 1
End Sub

Private Sub MnOperator_Click()
    DataOperator.Show 1
End Sub

Private Sub MnPenilaian_Click()
    Penilaian.Show
    Penilaian.ZOrder
End Sub


Private Sub RptDataPegawai_Click()
    ReportPegawai.Show 1
End Sub

Private Sub RptPenilaian_Click()
    ReportRanking.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            Toolbar1.Buttons(2).Enabled = True
            Login.Show 1
        Case 2
            
            With Toolbar1
                .Buttons(1).Enabled = True
                .Buttons(2).Enabled = False
                .Buttons(4).Enabled = False
                .Buttons(5).Enabled = False
                .Buttons(7).Enabled = False
            End With
            
            MnOperator.Enabled = False
            MnPenilaian.Enabled = False
            RptDataPegawai.Enabled = False
            MnDataPegawai.Enabled = False
            MenuUtama.StatusBar1.Panels.Item(2).Text = ""
            MenuUtama.StatusBar1.Panels.Item(4).Text = ""
            Dim Cancel As Integer, _
                UnloadMode As Integer
                Cancel = 1
                pesan = MsgBox("Logout ?", _
                vbExclamation + vbYesNo + vbDefaultButton2)
            If pesan = vbYes Then
                Cancel = 0
                End
            Else
                Cancel = 1
            End If
        Case 4
            MnDataPegawai_Click
        Case 5
            MnOperator_Click
        Case 7
            MnPenilaian_Click
    End Select
End Sub



Private Sub Timer1_Timer()
Dim a, b, c As String
    StatusBar1.Panels.Item(8).Text = Time
    a = StatusBar1.Panels.Item(9).Text
    b = Left(a, 1)
    c = Right(a, Len(a) - 1)
    StatusBar1.Panels.Item(9).Text = c + b
End Sub
