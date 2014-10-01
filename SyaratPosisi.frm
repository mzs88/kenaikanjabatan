VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SyaratPosisi 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Syarat Posisi"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6480
      Picture         =   "SyaratPosisi.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Tutup"
      Top             =   4920
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   7335
      Begin VB.ComboBox CmbJabatan 
         Height          =   315
         ItemData        =   "SyaratPosisi.frx":0A02
         Left            =   2280
         List            =   "SyaratPosisi.frx":0A1B
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
      Begin VB.ComboBox CmbSyarat 
         Height          =   315
         ItemData        =   "SyaratPosisi.frx":0B2D
         Left            =   2280
         List            =   "SyaratPosisi.frx":0B46
         TabIndex        =   7
         Top             =   960
         Width           =   3855
      End
      Begin VB.CommandButton BtnDelete 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6240
         Picture         =   "SyaratPosisi.frx":0C58
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Batal"
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton BtnHapus 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6240
         Picture         =   "SyaratPosisi.frx":165A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Hapus"
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton BtnUbah 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6240
         Picture         =   "SyaratPosisi.frx":205C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Simpan"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox CmbJumlah 
         Height          =   315
         ItemData        =   "SyaratPosisi.frx":2A5E
         Left            =   2280
         List            =   "SyaratPosisi.frx":2A77
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jabatan Yang Ditawarkan"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jumlah"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2040
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Syarat Posisi"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2040
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   7335
      Begin MSComctlLib.ListView LvSyaratPosisi 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Jabatan yang ditawarkan"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Jumlah"
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Syarat Posisi"
            Object.Width           =   6068
         EndProperty
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Syarat Posisi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   12
      Top             =   120
      Width           =   7320
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "SyaratPosisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
cNDb
LoadPenempatan
End Sub

Private Sub LoadPenempatan()
MySql = "Select * from tb_penempatan"
Set SdR = New ADODB.Recordset
SdR.CursorLocation = adUseClient
SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LvSyaratPosisi.View = lvwReport
    LvSyaratPosisi.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LvSyaratPosisi.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
            SdR.MoveNext
        Loop
End Sub

Private Sub Form_Initialize()

    Me.Left = (MenuUtama.ScaleWidth - Me.Width) / 2
    Me.Top = (MenuUtama.ScaleHeight - Me.Height) / 2

End Sub

Private Sub BtnDelete_Click()
CmbJabatan.Text = ""
CmbJumlah.Text = ""
CmbSyarat.Text = ""
End Sub

Private Sub BtnHapus_Click()
MySql = "DELETE FROM tb_penempatan"
ConN.Execute MySql
MsgBox ("Data Berhasil Dihapus")
LoadPenempatan
End Sub

Private Sub BtnUbah_Click()
MySql = "INSERT INTO tb_penempatan (jabatan, jumlah, syarat_posisi) VALUES ('" & CmbJabatan.Text & "', '" & CmbJumlah.Text & "', '" & CmbSyarat.Text & "')"
ConN.Execute MySql
MsgBox ("Data Berhasil Disimpan")
LoadPenempatan
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

