VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Kriteria 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kriteria"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7020
   Icon            =   "Kriteria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6120
      Picture         =   "Kriteria.frx":0FA2
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   5040
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pendidikan"
      TabPicture(0)   =   "Kriteria.frx":19A4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Pengalaman"
      TabPicture(1)   =   "Kriteria.frx":19C0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Karakter"
      TabPicture(2)   =   "Kriteria.frx":19DC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame8"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Kinerja"
      TabPicture(3)   =   "Kriteria.frx":19F8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "Frame9"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Masa Kerja"
      TabPicture(4)   =   "Kriteria.frx":1A14
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame10"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame6"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame4 
         BackColor       =   &H80000004&
         Height          =   1815
         Left            =   -74880
         TabIndex        =   63
         Top             =   480
         Width           =   6495
         Begin VB.CommandButton BtnKarTambah 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":1A30
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Tambah"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton BtnKarUbah 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":2432
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Ubah"
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton BtnKarHapus 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":2E34
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Hapus"
            Top             =   960
            Width           =   735
         End
         Begin VB.CommandButton BtnKarBatal 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":3836
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Batal"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox TxtKarNik 
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
            Left            =   1440
            TabIndex        =   0
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox CmbLeaderShip 
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
            ItemData        =   "Kriteria.frx":4238
            Left            =   1440
            List            =   "Kriteria.frx":4245
            TabIndex        =   1
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox TxtAHPLearning 
            Alignment       =   2  'Center
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
            Left            =   3720
            TabIndex        =   68
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtAHPLeader 
            Alignment       =   2  'Center
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
            Left            =   3720
            TabIndex        =   4
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox TxtSAWLeader 
            Alignment       =   2  'Center
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
            Left            =   4560
            TabIndex        =   67
            Top             =   600
            Width           =   735
         End
         Begin VB.ComboBox CmbLearning 
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
            ItemData        =   "Kriteria.frx":425E
            Left            =   1440
            List            =   "Kriteria.frx":426B
            TabIndex        =   2
            Top             =   960
            Width           =   2175
         End
         Begin VB.ComboBox CmbAttention 
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
            ItemData        =   "Kriteria.frx":4284
            Left            =   1440
            List            =   "Kriteria.frx":4291
            TabIndex        =   3
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox TxtSAWLearning 
            Alignment       =   2  'Center
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
            Left            =   4560
            TabIndex        =   66
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtAHPAttention 
            Alignment       =   2  'Center
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
            Left            =   3720
            TabIndex        =   65
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox TxtSAWAttention 
            Alignment       =   2  'Center
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
            Left            =   4560
            TabIndex        =   64
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Nik Pegawai"
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
            TabIndex        =   74
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Leadership"
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
            TabIndex        =   73
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Learning"
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
            TabIndex        =   72
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Attention"
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
            TabIndex        =   71
            Top             =   1320
            Width           =   1200
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "AHP"
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
            Left            =   3720
            TabIndex        =   70
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "SAW"
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
            Left            =   4560
            TabIndex        =   69
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000004&
         Height          =   1815
         Left            =   -74880
         TabIndex        =   58
         Top             =   480
         Width           =   6495
         Begin VB.TextBox TxtPengNik 
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
            Left            =   1440
            TabIndex        =   5
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox CmbPengalaman 
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
            ItemData        =   "Kriteria.frx":42AA
            Left            =   1440
            List            =   "Kriteria.frx":42B7
            TabIndex        =   6
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox TxtPengSAW 
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
            Left            =   1440
            TabIndex        =   8
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox TxtPengAHP 
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
            Left            =   1440
            TabIndex        =   7
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton BtnPengTambah 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":42DE
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Tambah"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton BtnPengUbah 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":4CE0
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Ubah"
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton BtnPengHapus 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":56E2
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Hapus"
            Top             =   960
            Width           =   735
         End
         Begin VB.CommandButton BtnPengBatal 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":60E4
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Batal"
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Nik Pegawai"
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
            TabIndex        =   62
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Pengalaman"
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
            TabIndex        =   61
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "AHP"
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
            TabIndex        =   60
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "SAW"
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
            TabIndex        =   59
            Top             =   1320
            Width           =   1200
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
         Height          =   2295
         Left            =   -74880
         TabIndex        =   57
         Top             =   2400
         Width           =   6495
         Begin MSComctlLib.ListView LsPendidikan 
            Height          =   1935
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nik Pegawai"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Pendidikan"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "AHP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "SAW"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000004&
         Height          =   1815
         Left            =   -74880
         TabIndex        =   52
         Top             =   480
         Width           =   6495
         Begin VB.CommandButton BtnPenBatal 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":6AE6
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Batal"
            Top             =   1320
            Width           =   735
         End
         Begin VB.CommandButton BtnPenHapus 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":74E8
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Hapus"
            Top             =   960
            Width           =   735
         End
         Begin VB.CommandButton BtnPenUbah 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":7EEA
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Ubah"
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton BtnPenTambah 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":88EC
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Tambah"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtPendAHP 
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
            Left            =   1440
            TabIndex        =   15
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox TxtPendSAW 
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
            Left            =   1440
            TabIndex        =   16
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ComboBox CmbPendidikan 
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
            ItemData        =   "Kriteria.frx":92EE
            Left            =   1440
            List            =   "Kriteria.frx":9301
            TabIndex        =   14
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox TxtPendNik 
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
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "SAW"
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
            TabIndex        =   56
            Top             =   1320
            Width           =   1200
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "AHP"
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
            TabIndex        =   55
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Pendidikan"
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
            TabIndex        =   54
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Nik Pegawai"
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
            TabIndex        =   53
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000004&
         Height          =   1815
         Left            =   -74880
         TabIndex        =   39
         Top             =   480
         Width           =   6495
         Begin VB.CommandButton BtnKinBatal 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":9320
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Batal"
            Top             =   1320
            Width           =   735
         End
         Begin VB.CommandButton BtnKinHapus 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":9D22
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Hapus"
            Top             =   960
            Width           =   735
         End
         Begin VB.CommandButton BtnKinUbah 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":A724
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Ubah"
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton BtnKinTambah 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":B126
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Tambah"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtKinAHP 
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
            Left            =   1440
            TabIndex        =   43
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox TxtKinSAW 
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
            Left            =   1440
            TabIndex        =   42
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ComboBox CmbKinerja 
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
            ItemData        =   "Kriteria.frx":BB28
            Left            =   1440
            List            =   "Kriteria.frx":BB35
            TabIndex        =   41
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox TxtKinNik 
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
            Left            =   1440
            TabIndex        =   40
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "SAW"
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
            TabIndex        =   51
            Top             =   1320
            Width           =   1200
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "AHP"
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
            TabIndex        =   50
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Kinerja"
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
            TabIndex        =   49
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Nik Pegawai"
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
            TabIndex        =   48
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   1815
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   6495
         Begin VB.CommandButton BtnMsBatal 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":BB4E
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Batal"
            Top             =   1320
            Width           =   735
         End
         Begin VB.CommandButton BtnMsHapus 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":C550
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Hapus"
            Top             =   960
            Width           =   735
         End
         Begin VB.CommandButton BtnMsUbah 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":CF52
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Ubah"
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton BtnMsTambah 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            Picture         =   "Kriteria.frx":D954
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Tambah"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtMsAHP 
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
            Left            =   1440
            TabIndex        =   30
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox TxtMsSAW 
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
            Left            =   1440
            TabIndex        =   29
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ComboBox CmbMsKerja 
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
            ItemData        =   "Kriteria.frx":E356
            Left            =   1440
            List            =   "Kriteria.frx":E363
            TabIndex        =   28
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox TxtMsNik 
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
            Left            =   1440
            TabIndex        =   27
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SAW"
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
            TabIndex        =   38
            Top             =   1320
            Width           =   1200
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "AHP"
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
            TabIndex        =   37
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Masa Kerja"
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
            TabIndex        =   36
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nik Pegawai"
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
            TabIndex        =   35
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H80000004&
         Height          =   2295
         Left            =   -74880
         TabIndex        =   25
         Top             =   2400
         Width           =   6495
         Begin MSComctlLib.ListView LsPengalaman 
            Height          =   1935
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nik Pegawai"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Pengalaman"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "AHP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "SAW"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H80000004&
         Height          =   2295
         Left            =   -74880
         TabIndex        =   24
         Top             =   2400
         Width           =   6495
         Begin MSComctlLib.ListView LsKarakter 
            Height          =   1935
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nik Pegawai"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Leadership"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Learning"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Attention"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "AHP Leadership"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "SAW Leadership"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "AHP Learning"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "SAW Learning"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "AHP Attention"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "SAW Attention"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000004&
         Height          =   2295
         Left            =   -74880
         TabIndex        =   23
         Top             =   2400
         Width           =   6495
         Begin MSComctlLib.ListView LsKinerja 
            Height          =   1935
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nik Pegawai"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Kinerja"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "AHP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "SAW"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H80000004&
         Height          =   2295
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   6495
         Begin MSComctlLib.ListView LsMsKerja 
            Height          =   1935
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nik Pegawai"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Masa Kerja "
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "AHP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "SAW"
               Object.Width           =   2540
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "Kriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
cNDb
LoadPendidikan
LoadPengalaman
LoadKarakter
LoadMasaKerja
LoadKinerja
End Sub

'============================ Pendidikan ==============================
Private Sub LoadPendidikan()
    Dim cList As ListItem

    MySql = "SELECT nik_pegawai, level_pendidikan, AHP, SAW FROM tb_pendidikan ORDER BY nik_pegawai ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LsPendidikan.View = lvwReport
    LsPendidikan.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LsPendidikan.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
                cList.SubItems(3) = SdR.Fields(3)
            SdR.MoveNext
        Loop
End Sub

Private Sub CmbPendidikan_Click()
    Dim CmB As String
    CmB = CmbPendidikan.Text
    Select Case CmB
        Case "SMA"
            TxtPendAHP.Text = "1"
            TxtPendSAW.Text = "0"
        Case "D1,D2,D3"
            TxtPendAHP.Text = "3"
            TxtPendSAW.Text = "0.25"
        Case "S1"
            TxtPendAHP.Text = "5"
            TxtPendSAW.Text = "0.5"
        Case "S2"
            TxtPendAHP.Text = "7"
            TxtPendSAW.Text = "0.7"
        Case "S3"
            TxtPendAHP.Text = "9"
            TxtPendSAW.Text = "1"
    End Select
End Sub

Private Sub BtnPenTambah_Click()
    If BtnPenTambah.ToolTipText = "Tambah" Then
        BtnPenTambah.ToolTipText = "Simpan"
        BtnPenTambah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        BtnPenUbah.Enabled = False
        ListPegawai.Label6.Caption = "Pendidikan"
        ListPegawai.Show vbModal
    Else
        MySql = "INSERT INTO tb_pendidikan (nik_pegawai, level_pendidikan, AHP, SAW) VALUES ( " & _
        "'" & TxtPendNik.Text & "', " & _
        "'" & CmbPendidikan.Text & "', " & _
        "'" & TxtPendAHP.Text & "', " & _
        "'" & TxtPendSAW.Text & "')"
        ConN.Execute MySql
        MsgBox ("Data Berhasil ditambah")
        BtnPenTambah.ToolTipText = "Tambah"
        BtnPenTambah.Picture = LoadPicture(App.Path & "\Button\Create.ico")
        BtnPenUbah.Enabled = True
        LoadPendidikan
        BtlPendidikan
    End If
End Sub

Private Sub BtnPenUbah_Click()
    If BtnPenUbah.ToolTipText = "Ubah" Then
        BtnPenUbah.ToolTipText = "Update"
        BtnPenUbah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        BtnPenTambah.Enabled = False
        BcPendidikan
    Else
        MySql = "UPDATE tb_pendidikan SET level_pendidikan = " & _
        "'" & CmbPendidikan.Text & "', ahp = " & _
        "'" & TxtPendAHP.Text & "', saw = " & _
        "'" & TxtPendSAW.Text & "' WHERE nik_pegawai = " & _
        "'" & LsPendidikan.ListItems(LsPendidikan.SelectedItem.Index) & "'"
        ConN.Execute MySql
        MsgBox "Data Sudah Dirubah"
        BtnPenUbah.ToolTipText = "Ubah"
        BtnPenUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
        BtnPenTambah.Enabled = True
        LoadPendidikan
        BtlPendidikan
    End If
End Sub

Private Sub BcPendidikan()
    TxtPendNik.Text = LsPendidikan.ListItems(LsPendidikan.SelectedItem.Index)
    CmbPendidikan.Text = LsPendidikan.ListItems(LsPendidikan.SelectedItem.Index).SubItems(1)
    TxtPendAHP.Text = LsPendidikan.ListItems(LsPendidikan.SelectedItem.Index).SubItems(2)
    TxtPendSAW.Text = LsPendidikan.ListItems(LsPendidikan.SelectedItem.Index).SubItems(3)
End Sub


Private Sub BtnPenHapus_Click()
    MySql = "DELETE FROM tb_pendidikan WHERE nik_pegawai = " & _
    "'" & LsPendidikan.ListItems(LsPendidikan.SelectedItem.Index) & "'"
    ConN.Execute MySql
    MsgBox "Data Berhasil Dihapus"
    LoadPendidikan
    BtlPendidikan
End Sub

Private Sub BtnPenBatal_Click()
    BtlPendidikan
End Sub

Private Sub BtlPendidikan()
    TxtPendNik.Text = ""
    CmbPendidikan.Text = ""
    TxtPendAHP.Text = ""
    TxtPendSAW.Text = ""
    BtnPenTambah.ToolTipText = "Tambah"
    BtnPenTambah.Picture = LoadPicture(App.Path & "\Button\Create.ico")
    BtnPenUbah.ToolTipText = "Ubah"
    BtnPenUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
    BtnPenTambah.Enabled = True
    BtnPenUbah.Enabled = True
End Sub

'============================ Pengalaman ==============================
Private Sub LoadPengalaman()
    Dim cList As ListItem

    MySql = "SELECT nik_pegawai, lama_pengalaman, AHP, SAW FROM tb_pengalaman ORDER BY nik_pegawai ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LsPengalaman.View = lvwReport
    LsPengalaman.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LsPengalaman.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
                cList.SubItems(3) = SdR.Fields(3)
            SdR.MoveNext
        Loop
End Sub

Private Sub CmbPengalaman_Click()
    Dim CmB As String
    CmB = CmbPengalaman.Text
    Select Case CmB
        Case "<1 Tahun"
            TxtPengAHP.Text = "3"
            TxtPengSAW.Text = "0.25"
        Case "2 s/d 5 Tahun"
            TxtPengAHP.Text = "5"
            TxtPengSAW.Text = "0.5"
        Case ">5 Tahun"
            TxtPengAHP.Text = "7"
            TxtPengSAW.Text = "0.75"
    End Select
End Sub

Private Sub BtnPengTambah_Click()
    If BtnPengTambah.ToolTipText = "Tambah" Then
        BtnPengTambah.ToolTipText = "Simpan"
        BtnPengTambah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        BtnPengUbah.Enabled = False
        ListPegawai.Label6.Caption = "Pengalaman"
        ListPegawai.Show vbModal
    Else
        MySql = "INSERT INTO tb_pengalaman (nik_pegawai, lama_pengalaman, AHP, SAW) VALUES ( " & _
        "'" & TxtPengNik.Text & "', " & _
        "'" & CmbPengalaman.Text & "', " & _
        "'" & TxtPengAHP.Text & "', " & _
        "'" & TxtPengSAW.Text & "')"
        ConN.Execute MySql
        MsgBox ("Data Berhasil ditambah")
        BtnPengTambah.ToolTipText = "Tambah"
        BtnPengTambah.Picture = LoadPicture(App.Path & "\Button\Create.ico ")
        BtnPengUbah.Enabled = True
        LoadPengalaman
        BtlPengalaman
    End If
End Sub

Private Sub BtnPengUbah_Click()
    If BtnPengUbah.ToolTipText = "Ubah" Then
        BtnPengUbah.ToolTipText = "Update"
        BtnPengUbah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        BtnPengTambah.Enabled = False
        BcPengalaman
    Else
        MySql = "UPDATE tb_pengalaman SET lama_pengalaman = " & _
        "'" & CmbPengalaman.Text & "', AHP = " & _
        "'" & TxtPengAHP.Text & "', SAW = " & _
        "'" & TxtPengSAW.Text & "' WHERE nik_pegawai = " & _
        "'" & LsPengalaman.ListItems(LsPengalaman.SelectedItem.Index) & "'"
        ConN.Execute MySql
        MsgBox "Data Sudah Dirubah"
        BtnPengUbah.ToolTipText = "Ubah"
        BtnPengUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
        LoadPengalaman
        BtlPengalaman
    End If
End Sub

Private Sub BcPengalaman()
    TxtPengNik.Text = LsPengalaman.ListItems(LsPengalaman.SelectedItem.Index)
    CmbPengalaman.Text = LsPengalaman.ListItems(LsPengalaman.SelectedItem.Index).SubItems(1)
    TxtPengAHP.Text = LsPengalaman.ListItems(LsPengalaman.SelectedItem.Index).SubItems(2)
    TxtPengSAW.Text = LsPengalaman.ListItems(LsPengalaman.SelectedItem.Index).SubItems(3)
End Sub

Private Sub BtnPengHapus_Click()
    MySql = "DELETE FROM tb_pengalaman WHERE nik_pegawai = " & _
    "'" & LsPengalaman.ListItems(LsPengalaman.SelectedItem.Index) & "'"
    ConN.Execute MySql
    MsgBox "Data Berhasil Dihapus"
    LoadPengalaman
    BtlPengalaman
End Sub

Private Sub BtnPengBatal_Click()
    BtlPengalaman
End Sub

Private Sub BtlPengalaman()
    TxtPengNik.Text = ""
    CmbPengalaman.Text = ""
    TxtPengAHP.Text = ""
    TxtPengSAW.Text = ""
    BtnPengTambah.ToolTipText = "Tambah"
    BtnPengTambah.Picture = LoadPicture(App.Path & "\Button\Create.ico")
    BtnPengUbah.ToolTipText = "Ubah"
    BtnPengUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
    BtnPengTambah.Enabled = True
    BtnPengUbah.Enabled = True
End Sub
'============================ Karakter ==============================
Private Sub LoadKarakter()
    Dim cList As ListItem

    MySql = "SELECT nik_pegawai, leadership_abilitiy, learning_abilitiy, " & _
    "attention_to_detail, AHP1, AHP2, AHP3, SAW1, SAW2, SAW3 FROM tb_karakter ORDER BY nik_pegawai ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LsKarakter.View = lvwReport
    LsKarakter.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LsKarakter.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
                cList.SubItems(3) = SdR.Fields(3)
                cList.SubItems(4) = SdR.Fields(4)
                cList.SubItems(5) = SdR.Fields(5)
                cList.SubItems(6) = SdR.Fields(6)
                cList.SubItems(7) = SdR.Fields(7)
                cList.SubItems(8) = SdR.Fields(8)
                cList.SubItems(9) = SdR.Fields(9)
            SdR.MoveNext
        Loop
End Sub

Private Sub CmbLeaderShip_Click()
    Dim CmB As String
    CmB = CmbLeaderShip.Text
    Select Case CmB
        Case "Baik"
            TxtAHPLeader.Text = "7"
            TxtSAWLeader.Text = "0.75"
        Case "Sedang"
            TxtAHPLeader.Text = "5"
            TxtSAWLeader.Text = "0.5"
        Case "Jelek"
            TxtAHPLeader.Text = "3"
            TxtSAWLeader.Text = "0.25"
    End Select
End Sub

Private Sub CmbLearning_Click()
    Dim CmB As String
    CmB = CmbLearning.Text
    Select Case CmB
        Case "Baik"
            TxtAHPLearning.Text = "7"
            TxtSAWLearning.Text = "0.75"
        Case "Sedang"
            TxtAHPLearning.Text = "5"
            TxtSAWLearning.Text = "0.5"
        Case "Jelek"
            TxtAHPLearning.Text = "3"
            TxtSAWLearning.Text = "0.25"
    End Select
End Sub

Private Sub CmbAttention_Click()
    Dim CmB As String
    CmB = CmbAttention.Text
    Select Case CmB
        Case "Baik"
            TxtAHPAttention.Text = "7"
            TxtSAWAttention.Text = "0.75"
        Case "Sedang"
            TxtAHPAttention.Text = "5"
            TxtSAWAttention.Text = "0.5"
        Case "Jelek"
            TxtAHPAttention.Text = "3"
            TxtSAWAttention.Text = "0.25"
    End Select
End Sub

Private Sub BtnKarTambah_Click()
    If BtnKarTambah.ToolTipText = "Tambah" Then
        BtnKarTambah.ToolTipText = "Simpan"
        BtnKarTambah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        BtnKarUbah.Enabled = False
        ListPegawai.Label6.Caption = "Karakter"
        ListPegawai.Show vbModal
    Else
        MySql = "INSERT INTO tb_karakter (nik_pegawai, leadership_abilitiy, learning_abilitiy, attention_to_detail, AHP1, AHP2, AHP3, SAW1, SAW2, SAW3) VALUES ( " & _
        "'" & TxtKarNik.Text & "', " & _
        "'" & CmbLeaderShip.Text & "', " & _
        "'" & CmbLearning.Text & "', " & _
        "'" & CmbAttention.Text & "', " & _
        "'" & TxtAHPLeader.Text & "', " & _
        "'" & TxtAHPLearning.Text & "', " & _
        "'" & TxtAHPAttention.Text & "', " & _
        "'" & TxtSAWLeader.Text & "'," & _
        "'" & TxtSAWLearning.Text & "', " & _
        "'" & TxtSAWAttention.Text & "')"
        ConN.Execute MySql
        MsgBox ("Data Berhasil ditambah")
        BtnKarTambah.ToolTipText = "Tambah"
        BtnKarTambah.Picture = LoadPicture(App.Path & "\Button\Create.ico")
        LoadKarakter
        BtlKarakter
    End If
End Sub

Private Sub BtnKarUbah_Click()
    If BtnKarUbah.ToolTipText = "Ubah" Then
        BtnKarUbah.ToolTipText = "Update"
        BtnKarUbah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        BtnKarTambah.Enabled = False
        BcKarakter
    Else
        MySql = "UPDATE tb_karakter SET leadership_abilitiy = " & _
        "'" & CmbLeaderShip.Text & "', learning_abilitiy = " & _
        "'" & CmbLearning.Text & "', attention_to_detail = " & _
        "'" & CmbAttention.Text & "', AHP1 = " & _
        "'" & TxtAHPLeader.Text & "', ahp2 = " & _
        "'" & TxtAHPLearning.Text & "', ahp3 = " & _
        "'" & TxtAHPAttention.Text & "', saw1 = " & _
        "'" & TxtSAWLeader.Text & "', saw2 = " & _
        "'" & TxtSAWLearning.Text & "', saw3 = " & _
        "'" & TxtSAWAttention.Text & "' WHERE nik_pegawai = " & _
        "'" & LsKarakter.ListItems(LsKarakter.SelectedItem.Index) & "'"
        ConN.Execute MySql
        MsgBox "Data Sudah Dirubah"
        BtnKarUbah.ToolTipText = "Ubah"
        BtnKarUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
        LoadKarakter
        BtlKarakter
    End If
End Sub

Private Sub BcKarakter()
    TxtKarNik.Text = LsKarakter.ListItems(LsKarakter.SelectedItem.Index)
    CmbLeaderShip.Text = LsKarakter.ListItems(LsKarakter.SelectedItem.Index).SubItems(1)
    CmbLearning.Text = LsKarakter.ListItems(LsKarakter.SelectedItem.Index).SubItems(2)
    CmbAttention.Text = LsKarakter.ListItems(LsKarakter.SelectedItem.Index).SubItems(3)
    TxtAHPLeader.Text = LsKarakter.ListItems(LsKarakter.SelectedItem.Index).SubItems(4)
    TxtAHPLearning.Text = LsKarakter.ListItems(LsKarakter.SelectedItem.Index).SubItems(5)
    TxtAHPAttention.Text = LsKarakter.ListItems(LsKarakter.SelectedItem.Index).SubItems(6)
    TxtSAWLeader.Text = LsKarakter.ListItems(LsKarakter.SelectedItem.Index).SubItems(7)
    TxtSAWLearning.Text = LsKarakter.ListItems(LsKarakter.SelectedItem.Index).SubItems(8)
    TxtSAWAttention.Text = LsKarakter.ListItems(LsKarakter.SelectedItem.Index).SubItems(9)
End Sub

Private Sub BtnKarHapus_Click()
    MySql = "DELETE FROM tb_karakter WHERE nik_pegawai = " & _
    "'" & LsKarakter.ListItems(LsKarakter.SelectedItem.Index) & "'"
    ConN.Execute MySql
    MsgBox "Data Berhasil Dihapus"
    LoadKarakter
    BtlKarakter
End Sub

Private Sub BtnKarBatal_Click()
    BtlKarakter
End Sub

Private Sub BtlKarakter()
    TxtKarNik.Text = ""
    CmbLeaderShip.Text = ""
    CmbLearning.Text = ""
    CmbAttention.Text = ""
    TxtAHPLeader.Text = ""
    TxtAHPLearning.Text = ""
    TxtAHPAttention.Text = ""
    TxtSAWLeader.Text = ""
    TxtSAWLearning.Text = ""
    TxtSAWAttention.Text = ""
    BtnKarTambah.ToolTipText = "Tambah"
    BtnKarTambah.Picture = LoadPicture(App.Path & "\Button\Create.ico")
    BtnKarUbah.ToolTipText = "Ubah"
    BtnKarUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
    BtnKarTambah.Enabled = True
    BtnKarUbah.Enabled = True
End Sub

'============================ Kinerja ==============================
Private Sub LoadKinerja()
    Dim cList As ListItem

    MySql = "SELECT nik_pegawai, kinerja, AHP, SAW FROM tb_kinerja ORDER BY nik_pegawai ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LsKinerja.View = lvwReport
    LsKinerja.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LsKinerja.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
                cList.SubItems(3) = SdR.Fields(3)
            SdR.MoveNext
        Loop
End Sub

Private Sub CmbKinerja_Click()
    Dim CmB As String
    CmB = CmbKinerja.Text
    Select Case CmB
        Case "Baik"
            TxtKinAHP.Text = "7"
            TxtKinSAW.Text = "0.75"
        Case "Sedang"
            TxtKinAHP.Text = "5"
            TxtKinSAW.Text = "0.5"
        Case "Jelek"
            TxtKinAHP.Text = "3"
            TxtKinSAW.Text = "0.25"
    End Select
End Sub

Private Sub BtnKinTambah_Click()
    If BtnKinTambah.ToolTipText = "Tambah" Then
        BtnKinTambah.ToolTipText = "Simpan"
        BtnKinTambah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        BtnKinUbah.Enabled = False
        ListPegawai.Label6.Caption = "Kinerja"
        ListPegawai.Show vbModal
    Else
        MySql = "INSERT INTO tb_kinerja (nik_pegawai, kinerja, AHP, SAW) VALUES ( " & _
        "'" & TxtKinNik.Text & "', " & _
        "'" & CmbKinerja.Text & "', " & _
        "'" & TxtKinAHP.Text & "', " & _
        "'" & TxtKinSAW.Text & "')"
        ConN.Execute MySql
        MsgBox ("Data Berhasil ditambah")
        BtnKinTambah.ToolTipText = "Tambah"
        BtnKinTambah.Picture = LoadPicture(App.Path & "\Button\Create.ico")
        LoadKinerja
        BtlKinerja
    End If
End Sub

Private Sub BtnKinUbah_Click()
    If BtnKinUbah.ToolTipText = "Ubah" Then
        BtnKinUbah.ToolTipText = "Update"
        BtnKinUbah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        BtnKinTambah.Enabled = False
        BcKinerja
    Else
        MySql = "UPDATE tb_kinerja SET kinerja = " & _
        "'" & CmbKinerja.Text & "', AHP = " & _
        "'" & TxtKinAHP.Text & "', SAW = " & _
        "'" & TxtKinSAW.Text & "' WHERE nik_pegawai = " & _
        "'" & LsKinerja.ListItems(LsKinerja.SelectedItem.Index) & "'"
        ConN.Execute MySql
        MsgBox "Data Sudah Dirubah"
        BtnKinUbah.ToolTipText = "Ubah"
        BtnKinUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
        LoadKinerja
        BtlKinerja
    End If
End Sub

Private Sub BcKinerja()
    TxtKinNik.Text = LsKinerja.ListItems(LsKinerja.SelectedItem.Index)
    CmbKinerja.Text = LsKinerja.ListItems(LsKinerja.SelectedItem.Index).SubItems(1)
    TxtKinAHP.Text = LsKinerja.ListItems(LsKinerja.SelectedItem.Index).SubItems(2)
    TxtKinSAW.Text = LsKinerja.ListItems(LsKinerja.SelectedItem.Index).SubItems(3)
End Sub

Private Sub BtnKinHapus_Click()
    MySql = "DELETE FROM tb_kinerja WHERE nik_pegawai = " & _
    "'" & LsKinerja.ListItems(LsKinerja.SelectedItem.Index) & "'"
    ConN.Execute MySql
    MsgBox "Data Berhasil Dihapus"
    LoadKinerja
    BtlKinerja
End Sub

Private Sub BtnKinBatal_Click()
    BtlKinerja
End Sub

Private Sub BtlKinerja()
    TxtKinNik.Text = ""
    CmbKinerja.Text = ""
    TxtKinAHP.Text = ""
    TxtKinSAW.Text = ""
    BtnKinTambah.ToolTipText = "Tambah"
    BtnKinTambah.Picture = LoadPicture(App.Path & "\Button\Create.ico")
    BtnKinUbah.ToolTipText = "Ubah"
    BtnKinUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
    BtnKinTambah.Enabled = True
    BtnKinUbah.Enabled = True
End Sub
'============================ MasaKerja ==============================
Private Sub LoadMasaKerja()
    Dim cList As ListItem

    MySql = "SELECT nik_pegawai, masa_kerja, AHP, SAW FROM tb_masakerja ORDER BY nik_pegawai ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LsMsKerja.View = lvwReport
    LsMsKerja.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LsMsKerja.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
                cList.SubItems(3) = SdR.Fields(3)
            SdR.MoveNext
        Loop
End Sub

Private Sub CmbMsKerja_Click()
    Dim CmB As String
    CmB = CmbMsKerja.Text
    Select Case CmB
        Case "1 Tahun"
            TxtMsAHP.Text = "3"
            TxtMsSAW.Text = "0.25"
        Case "2 s/d 3 Tahun"
            TxtMsAHP.Text = "5"
            TxtMsSAW.Text = "0.5"
        Case "4 Tahun"
            TxtMsAHP.Text = "7"
            TxtMsSAW.Text = "0.75"
    End Select
End Sub

Private Sub BtnMsTambah_Click()
    If BtnMsTambah.ToolTipText = "Tambah" Then
        BtnMsTambah.ToolTipText = "Simpan"
        BtnMsTambah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        BtnMsUbah.Enabled = False
        ListPegawai.Label6.Caption = "Masa Kerja"
        ListPegawai.Show vbModal
    Else
        MySql = "INSERT INTO tb_masakerja (nik_pegawai, masa_kerja, AHP, SAW) VALUES ( " & _
        "'" & TxtMsNik.Text & "', " & _
        "'" & CmbMsKerja.Text & "', " & _
        "'" & TxtMsAHP.Text & "', " & _
        "'" & TxtMsSAW.Text & "')"
        ConN.Execute MySql
        MsgBox ("Data Berhasil ditambah")
        BtnMsTambah.ToolTipText = "Tambah"
        BtnMsTambah.Picture = LoadPicture(App.Path & "\Button\Create.ico")
        LoadMasaKerja
        BtlMsKerja
    End If
End Sub

Private Sub BtnMsUbah_Click()
    If BtnMsUbah.ToolTipText = "Ubah" Then
        BtnMsUbah.ToolTipText = "Update"
        BtnMsUbah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        BtnMsTambah.Enabled = False
        BcMsKerja
    Else
        MySql = "UPDATE tb_masakerja SET masa_kerja = " & _
        "'" & CmbMsKerja.Text & "', AHP = " & _
        "'" & TxtMsAHP.Text & "', SAW = " & _
        "'" & TxtMsSAW.Text & "' WHERE nik_pegawai = " & _
        "'" & LsMsKerja.ListItems(LsMsKerja.SelectedItem.Index) & "'"
        ConN.Execute MySql
        MsgBox "Data Sudah Dirubah"
        BtnMsUbah.ToolTipText = "Ubah"
        BtnMsUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
        LoadMasaKerja
        BtlMsKerja
    End If
End Sub

Private Sub BcMsKerja()
    TxtMsNik.Text = LsMsKerja.ListItems(LsMsKerja.SelectedItem.Index)
    CmbMsKerja.Text = LsMsKerja.ListItems(LsMsKerja.SelectedItem.Index).SubItems(1)
    TxtMsAHP.Text = LsMsKerja.ListItems(LsMsKerja.SelectedItem.Index).SubItems(2)
    TxtMsSAW.Text = LsMsKerja.ListItems(LsMsKerja.SelectedItem.Index).SubItems(3)
End Sub

Private Sub BtnMsHapus_Click()
    MySql = "DELETE FROM tb_masakerja WHERE nik_pegawai = " & _
    "'" & LsMsKerja.ListItems(LsMsKerja.SelectedItem.Index) & "'"
    ConN.Execute MySql
    MsgBox "Data Berhasil Dihapus"
    LoadMasaKerja
    BtlMsKerja
End Sub

Private Sub BtnMsBatal_Click()
    BtlMsKerja
End Sub

Private Sub BtlMsKerja()
    TxtMsNik.Text = ""
    CmbMsKerja.Text = ""
    TxtMsAHP.Text = ""
    TxtMsSAW.Text = ""
    BtnMsTambah.ToolTipText = "Tambah"
    BtnMsTambah.Picture = LoadPicture(App.Path & "\Button\Create.ico")
    BtnMsUbah.ToolTipText = "Ubah"
    BtnMsUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
    BtnMsTambah.Enabled = True
    BtnMsUbah.Enabled = True

End Sub

Private Sub LsKarakter_Click()
    BtlKarakter
End Sub

Private Sub LsKinerja_Click()
    BtlKinerja
End Sub

Private Sub LsMsKerja_Click()
    BtlMsKerja
End Sub

Private Sub LsPendidikan_Click()
    BtlPendidikan
End Sub

Private Sub LsPengalaman_Click()
    BtlPengalaman
End Sub
