VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form DataOperator 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Operator"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10545
   Icon            =   "DataOperator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   4080
      TabIndex        =   8
      Top             =   720
      Width           =   6375
      Begin MSComctlLib.ListView LvDataOperator 
         Height          =   2655
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4683
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nik"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Username"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Level"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   3855
      Begin VB.CommandButton BtnBatal 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3000
         Picture         =   "DataOperator.frx":0FA2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Batal"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton BtnHapus 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2160
         Picture         =   "DataOperator.frx":19A4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Hapus"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton BtnUbah 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   960
         Picture         =   "DataOperator.frx":23A6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ubah"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton BtnSimpan 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Picture         =   "DataOperator.frx":2DA8
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Tambah"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame TxtIDTxtID 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tampilkan Password"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox CmbLevel 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "DataOperator.frx":37AA
         Left            =   1320
         List            =   "DataOperator.frx":37AC
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1320
         Width           =   2415
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
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox TxtNama 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox TxtID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Level"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nama"
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
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nik"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Data Operator"
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
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   10320
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "DataOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        TxtPass.PasswordChar = ""
    Else
        TxtPass.PasswordChar = "*"
    End If
End Sub


Private Sub Form_Load()
cNDb
LoadOp
LoadLevel
End Sub

Private Sub LoadOp()
    MySql = "SELECT nik, user_name, level FROM login ORDER BY user_name"
    Set SdR = New ADODB.Recordset
    SdR.CursorLocation = adUseClient
    SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LvDataOperator.View = lvwReport
    LvDataOperator.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LvDataOperator.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
            SdR.MoveNext
        Loop
End Sub

Private Sub BtnBatal_Click()
    Batal
End Sub

Private Sub Batal()
    TxtID.Text = ""
    TxtNama.Text = ""
    TxtPass.Text = ""
    BtnSimpan.ToolTipText = "Tambah"
    BtnHapus.ToolTipText = "Hapus"
    BtnUbah.ToolTipText = "Ubah"
    BtnUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
    BtnSimpan.Picture = LoadPicture(App.Path & "\Button\Create.ico")
End Sub

Private Sub BtnHapus_Click()
    MySql = "DELETE FROM `login` WHERE " & _
    "nik = '" & LvDataOperator.ListItems(LvDataOperator.SelectedItem.Index) & "'"
    ConN.Execute MySql
    MsgBox ("Data Berhasil Dihapus")
    LoadOp
End Sub

Private Sub BtnSimpan_Click()
    If BtnSimpan.ToolTipText = "Tambah" Then
        BtnSimpan.ToolTipText = "Simpan"
        BtnSimpan.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        ListPegawai.Label6.Caption = "Pilih Operator"
        ListPegawai.Show 1
        
    Else
        If CmbLevel.Text = "" Or TxtPass.Text = "" Then
            MsgBox "Data belum lengkap"
        Else
        
            MySql = "INSERT INTO login(nik, user_name, password, level) " & _
            "VALUES ('" & TxtID.Text & "','" & TxtNama.Text & "', '" & TxtPass.Text & "', '" & CmbLevel.Text & "')"
            ConN.Execute MySql
            MsgBox ("Data Berhasil Disimpan")
            Batal
            LoadOp
        End If
    End If
End Sub

Private Sub BtnUbah_Click()
If BtnUbah.ToolTipText = "Ubah" Then
    MySql = "SELECT nik, user_name, password FROM login WHERE " & _
    "nik = '" & LvDataOperator.ListItems(LvDataOperator.SelectedItem.Index) & "'"
    Set SdR = New ADODB.Recordset
    SdR.CursorLocation = adUseClient
    SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic
    SdR.Requery
    With SdR
        If .EOF Then
        TxtNama.Text = ""
        Else
        TxtID.Text = .Fields(0)
        TxtNama.Text = .Fields(1)
        TxtPass.Text = .Fields(2)
        End If
    End With
    BtnUbah.ToolTipText = "Simpan"
    BtnUbah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
Else
    MySql = "UPDATE `login` SET `user_name`=" & _
    "'" & TxtNama.Text & "',`password`= " & _
    "'" & TxtPass.Text & "',`level`= " & _
    "'" & CmbLevel.Text & "' WHERE nik = '" & TxtID.Text & "'"
    ConN.Execute MySql
    MsgBox ("Data Berhasil Dirubah")
    BtnUbah.ToolTipText = "Ubah"
    Batal
    TxtID.Enabled = True
    BtnUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
    LoadOp
End If
End Sub



Private Sub cID()
Dim cKode As String * 4
Dim cHitung As Long
     MySql = "SELECT nik FROM login  Where nik In(Select Max(nik)From login)Order By nik Desc"
     Set SdR = New ADODB.Recordset
     SdR.CursorLocation = adUseClient
     SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic
    SdR.Requery
    With SdR
        If .EOF Then
            cKode = "0001"
            TxtNik = cKode
        Else
            cHitung = Str(!nik) + 1
            cKode = Right("0000" & cHitung, 4)
        End If
        TxtID.Text = cKode
    End With
    SdR.Close
End Sub

Private Sub LoadLevel()
With CmbLevel
    .AddItem ("Admin")
    .AddItem ("User")
End With
End Sub

Private Sub LvDataOperator_Click()
    Batal
End Sub
