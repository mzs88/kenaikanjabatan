VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Penghitungan 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Penghitungan"
   ClientHeight    =   10845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18150
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10845
   ScaleWidth      =   18150
   StartUpPosition =   2  'CenterScreen
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18150
      _ExtentX        =   32015
      _ExtentY        =   741
      BandCount       =   2
      BackColor       =   16777215
      _CBWidth        =   18150
      _CBHeight       =   420
      _Version        =   "6.0.8169"
      MinHeight1      =   360
      Width1          =   1440
      NewRow1         =   0   'False
      MinHeight2      =   360
      Width2          =   1440
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   450
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   794
         ButtonWidth     =   820
         ButtonHeight    =   794
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Param"
               Object.ToolTipText     =   "Parameter"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bobot"
               Object.ToolTipText     =   "Bobot Kriteria"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "KtR"
               Object.ToolTipText     =   "Kriteria"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Proses"
               Object.ToolTipText     =   "Prosess Data"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Rst"
               Object.ToolTipText     =   "Reset"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "CtkLpr"
               Object.ToolTipText     =   "Cetak Rangking"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AHP"
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   17895
      Begin VB.Frame DTahp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   5160
         TabIndex        =   151
         Top             =   1800
         Visible         =   0   'False
         Width           =   9495
         Begin VB.Frame Frame38 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            TabIndex        =   152
            Top             =   6360
            Width           =   3255
            Begin MSComctlLib.ListView ListView1 
               Height          =   2295
               Left            =   120
               TabIndex        =   153
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4048
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Rangking"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin MSComctlLib.ListView LsKarakterAHP 
            Height          =   3135
            Left            =   120
            TabIndex        =   155
            Top             =   1200
            Width           =   13695
            _ExtentX        =   24156
            _ExtentY        =   5530
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
            NumItems        =   7
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
               Text            =   "Leadership"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Learning"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Attention"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Data Mentah AHP"
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
            TabIndex        =   154
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MnLs 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9120
         TabIndex        =   63
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame19 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   67
            Top             =   1080
            Width           =   10335
            Begin MSComctlLib.ListView LvLeader4 
               Height          =   2295
               Left            =   120
               TabIndex        =   68
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvLeader5 
               Height          =   735
               Left            =   120
               TabIndex        =   69
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame18 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   10560
            TabIndex        =   64
            Top             =   1080
            Width           =   3255
            Begin MSComctlLib.ListView LvLeader7 
               Height          =   735
               Left            =   120
               TabIndex        =   65
               Top             =   2640
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1296
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LvLeader6 
               Height          =   2295
               Left            =   120
               TabIndex        =   66
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4048
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Normalisasi Leadership"
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
            TabIndex        =   70
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MnK 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7800
         TabIndex        =   41
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame25 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Frame2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   45
            Top             =   1065
            Width           =   10335
            Begin MSComctlLib.ListView LvKinerja4 
               Height          =   2295
               Left            =   120
               TabIndex        =   46
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvKinerja5 
               Height          =   735
               Left            =   120
               TabIndex        =   47
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame26 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   10560
            TabIndex        =   42
            Top             =   1065
            Width           =   3255
            Begin MSComctlLib.ListView LvKinerja6 
               Height          =   2295
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4048
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LvKinerja7 
               Height          =   735
               Left            =   120
               TabIndex        =   44
               Top             =   2640
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1296
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Normalisasi Kinerja"
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
            TabIndex        =   48
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MnPe 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   10560
            TabIndex        =   15
            Top             =   1065
            Width           =   3255
            Begin MSComctlLib.ListView LvPengalaman6 
               Height          =   2295
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4048
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LvPengalaman7 
               Height          =   735
               Left            =   120
               TabIndex        =   17
               Top             =   2640
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1296
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   12
            Top             =   1065
            Width           =   10335
            Begin MSComctlLib.ListView LvPengalaman4 
               Height          =   2295
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvPengalaman5 
               Height          =   735
               Left            =   120
               TabIndex        =   14
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Normalisasi Pengalaman"
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
            TabIndex        =   18
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MpPe 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   23
            Top             =   1020
            Width           =   3255
            Begin MSComctlLib.ListView LvPengalaman1 
               Height          =   3135
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   5530
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Skala"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   3480
            TabIndex        =   20
            Top             =   1020
            Width           =   10335
            Begin MSComctlLib.ListView LvPengalaman2 
               Height          =   2295
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvPengalaman3 
               Height          =   735
               Left            =   120
               TabIndex        =   22
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Perbandingan Pengalaman"
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
            TabIndex        =   25
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MnP 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   30
            Top             =   1065
            Width           =   10335
            Begin MSComctlLib.ListView LvPendidikan4 
               Height          =   2295
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvPendidikan5 
               Height          =   735
               Left            =   120
               TabIndex        =   32
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   10560
            TabIndex        =   27
            Top             =   1065
            Width           =   3255
            Begin MSComctlLib.ListView LvPendidikan6 
               Height          =   2295
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4048
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LvPendidikan7 
               Height          =   735
               Left            =   120
               TabIndex        =   29
               Top             =   2640
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1296
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Normalisasi Pendidikan"
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
            TabIndex        =   33
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MpP 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame13 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   3480
            TabIndex        =   37
            Top             =   1020
            Width           =   10335
            Begin MSComctlLib.ListView LvPendidikan2 
               Height          =   2295
               Left            =   120
               TabIndex        =   38
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvPendidikan3 
               Height          =   735
               Left            =   120
               TabIndex        =   39
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   35
            Top             =   1020
            Width           =   3255
            Begin MSComctlLib.ListView LvPendidikan1 
               Height          =   3135
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   5530
               View            =   3
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Skala"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Perbandingan Pendidikan"
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
            TabIndex        =   40
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame PnB 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   124
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame41 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   128
            Top             =   1065
            Width           =   10335
            Begin MSComctlLib.ListView LvBobot4 
               Height          =   2295
               Left            =   120
               TabIndex        =   129
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvBobot5 
               Height          =   735
               Left            =   120
               TabIndex        =   130
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame40 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   10560
            TabIndex        =   125
            Top             =   1065
            Width           =   3255
            Begin MSComctlLib.ListView LvBobot6 
               Height          =   2295
               Left            =   120
               TabIndex        =   126
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4048
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Kriteria"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LvBobot7 
               Height          =   735
               Left            =   120
               TabIndex        =   127
               Top             =   2640
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1296
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Kriteria"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Normalisasi Berpasangan"
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
            TabIndex        =   131
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MnMk 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13080
         TabIndex        =   108
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame35 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   112
            Top             =   1065
            Width           =   10335
            Begin MSComctlLib.ListView LvMsKerja4 
               Height          =   2295
               Left            =   120
               TabIndex        =   113
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvMsKerja5 
               Height          =   735
               Left            =   120
               TabIndex        =   114
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame34 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   10560
            TabIndex        =   109
            Top             =   1065
            Width           =   3255
            Begin MSComctlLib.ListView LvMsKerja6 
               Height          =   2295
               Left            =   120
               TabIndex        =   110
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4048
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LvMsKerja7 
               Height          =   735
               Left            =   120
               TabIndex        =   111
               Top             =   2640
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1296
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Normalisasi Masa Kerja"
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
            TabIndex        =   115
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame PpB 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   49
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame15 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   53
            Top             =   1020
            Width           =   3255
            Begin MSComctlLib.ListView LvBobot1 
               Height          =   3135
               Left            =   120
               TabIndex        =   54
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   5530
               View            =   3
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Kriteria"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "ahp"
                  Object.Width           =   882
               EndProperty
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   3480
            TabIndex        =   50
            Top             =   1020
            Width           =   10335
            Begin MSComctlLib.ListView LvBobot2 
               Height          =   2295
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvBobot3 
               Height          =   735
               Left            =   120
               TabIndex        =   52
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Perbandingan Berpasangan"
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
            TabIndex        =   55
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1455
         Left            =   17640
         Max             =   0
         TabIndex        =   144
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame MpMk 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13080
         TabIndex        =   101
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame33 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   3480
            TabIndex        =   104
            Top             =   1020
            Width           =   10335
            Begin MSComctlLib.ListView LvMsKerja2 
               Height          =   2295
               Left            =   120
               TabIndex        =   105
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvMsKerja3 
               Height          =   735
               Left            =   120
               TabIndex        =   106
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame32 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   102
            Top             =   1020
            Width           =   3255
            Begin MSComctlLib.ListView LvMsKerja1 
               Height          =   3135
               Left            =   120
               TabIndex        =   103
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   5530
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Skala"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Perbandingan Masa Kerja"
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
            TabIndex        =   107
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MnA 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11760
         TabIndex        =   93
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame31 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   97
            Top             =   1080
            Width           =   10335
            Begin MSComctlLib.ListView LvAttention4 
               Height          =   2295
               Left            =   120
               TabIndex        =   98
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvAttention5 
               Height          =   735
               Left            =   120
               TabIndex        =   99
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame30 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   10560
            TabIndex        =   94
            Top             =   1080
            Width           =   3255
            Begin MSComctlLib.ListView LvAttention6 
               Height          =   2295
               Left            =   120
               TabIndex        =   95
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4048
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LvAttention7 
               Height          =   735
               Left            =   120
               TabIndex        =   96
               Top             =   2640
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1296
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Normalisasi Attention"
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
            TabIndex        =   100
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MpA 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11760
         TabIndex        =   86
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame29 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   3480
            TabIndex        =   89
            Top             =   960
            Width           =   10335
            Begin MSComctlLib.ListView LvAttention2 
               Height          =   2295
               Left            =   120
               TabIndex        =   90
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvAttention3 
               Height          =   735
               Left            =   120
               TabIndex        =   91
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame24 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Frame1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   87
            Top             =   960
            Width           =   3255
            Begin MSComctlLib.ListView LvAttention1 
               Height          =   3135
               Left            =   120
               TabIndex        =   88
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   5530
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Skala"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Perbandingan Attention"
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
            TabIndex        =   92
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MnLr 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10440
         TabIndex        =   78
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame23 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   82
            Top             =   1080
            Width           =   10335
            Begin MSComctlLib.ListView LvLearning4 
               Height          =   2295
               Left            =   120
               TabIndex        =   83
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvLearning5 
               Height          =   735
               Left            =   120
               TabIndex        =   84
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame22 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   10560
            TabIndex        =   79
            Top             =   1080
            Width           =   3255
            Begin MSComctlLib.ListView LvLearning6 
               Height          =   2295
               Left            =   120
               TabIndex        =   80
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4048
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LvLearning7 
               Height          =   735
               Left            =   120
               TabIndex        =   81
               Top             =   2640
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1296
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Jumlah"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Normalisasi Learning"
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
            TabIndex        =   85
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MpLr 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10440
         TabIndex        =   71
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame21 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   3480
            TabIndex        =   74
            Top             =   960
            Width           =   10335
            Begin MSComctlLib.ListView LvLearning2 
               Height          =   2295
               Left            =   120
               TabIndex        =   75
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvLearning3 
               Height          =   735
               Left            =   120
               TabIndex        =   76
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame20 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Frame1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   3255
            Begin MSComctlLib.ListView LvLearning1 
               Height          =   3135
               Left            =   120
               TabIndex        =   73
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   5530
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Skala"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Perbandingan Learning"
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
            TabIndex        =   77
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MpLs 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9120
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame17 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   3480
            TabIndex        =   59
            Top             =   960
            Width           =   10335
            Begin MSComctlLib.ListView LvLeader2 
               Height          =   2295
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvLeader3 
               Height          =   735
               Left            =   120
               TabIndex        =   61
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   57
            Top             =   960
            Width           =   3255
            Begin MSComctlLib.ListView LvLeader1 
               Height          =   3135
               Left            =   120
               TabIndex        =   58
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   5530
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Skala"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Perbandingan Leadership"
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
            Height          =   435
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame MpK 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7800
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Frame Frame27 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   3480
            TabIndex        =   7
            Top             =   1020
            Width           =   10335
            Begin MSComctlLib.ListView LvKinerja2 
               Height          =   2295
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   4048
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
               NumItems        =   0
            End
            Begin MSComctlLib.ListView LvKinerja3 
               Height          =   735
               Left            =   120
               TabIndex        =   9
               Top             =   2640
               Width           =   10095
               _ExtentX        =   17806
               _ExtentY        =   1296
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
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame28 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   5
            Top             =   1020
            Width           =   3255
            Begin MSComctlLib.ListView LvKinerja1 
               Height          =   3135
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   5530
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Skala"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Perbandingan Kinerja"
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
            TabIndex        =   10
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4335
         Left            =   240
         TabIndex        =   136
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   7646
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
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
      End
      Begin VB.Frame RnKAHP 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   3840
         TabIndex        =   116
         Top             =   1800
         Visible         =   0   'False
         Width           =   8175
         Begin VB.Frame Frame39 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            TabIndex        =   121
            Top             =   6360
            Width           =   3255
            Begin MSComctlLib.ListView LvAkumulasi3 
               Height          =   2295
               Left            =   120
               TabIndex        =   122
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4048
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Rangking"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Frame Frame37 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            TabIndex        =   119
            Top             =   3720
            Width           =   11775
            Begin MSComctlLib.ListView LvAkumulasi2 
               Height          =   2295
               Left            =   120
               TabIndex        =   120
               Top             =   240
               Width           =   11535
               _ExtentX        =   20346
               _ExtentY        =   4048
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
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Pendidikan"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Pengalaman"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Leader Ship"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Learning"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "Attention"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   6
                  Text            =   "Kinerja"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   7
                  Text            =   "Masa Kerja"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Frame Frame36 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            TabIndex        =   117
            Top             =   1080
            Width           =   11775
            Begin MSComctlLib.ListView LvAkumulasi1 
               Height          =   2295
               Left            =   120
               TabIndex        =   118
               Top             =   240
               Width           =   11535
               _ExtentX        =   20346
               _ExtentY        =   4048
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
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Pendidikan"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Pengalaman"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Leader Ship"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Learning"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "Attention"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   6
                  Text            =   "Kinerja"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   7
                  Text            =   "Masa Kerja"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Nilai Perbandingan && Normalisasi Penghitungan Global"
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
            TabIndex        =   123
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SAW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   17895
      Begin VB.Frame DTsaw 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3840
         TabIndex        =   156
         Top             =   2520
         Visible         =   0   'False
         Width           =   11415
         Begin VB.Frame Frame42 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            TabIndex        =   157
            Top             =   6360
            Width           =   3255
            Begin MSComctlLib.ListView ListView2 
               Height          =   2295
               Left            =   120
               TabIndex        =   158
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4048
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Rangking"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin MSComctlLib.ListView LsKarakterSAW 
            Height          =   3135
            Left            =   120
            TabIndex        =   159
            Top             =   1200
            Width           =   13695
            _ExtentX        =   24156
            _ExtentY        =   5530
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
            NumItems        =   7
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
               Text            =   "Leadership"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Learning"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Attention"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Data Mentah SAW"
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
            TabIndex        =   160
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame BpS1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   132
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         Begin VB.Frame Frame44 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            TabIndex        =   133
            Top             =   960
            Width           =   13695
            Begin MSComctlLib.ListView LvBobotSAW3 
               Height          =   2055
               Left            =   3240
               TabIndex        =   134
               Top             =   240
               Width           =   10335
               _ExtentX        =   18230
               _ExtentY        =   3625
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
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Pendidikan"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Pengalaman"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Leadership"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Learning"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "Attention"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   6
                  Text            =   "Kinerja"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   7
                  Text            =   "Masa Kerja"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LvBobotSAW4 
               Height          =   975
               Left            =   3240
               TabIndex        =   135
               Top             =   2400
               Width           =   10335
               _ExtentX        =   18230
               _ExtentY        =   1720
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
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   6
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   7
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LvBobotSAW2 
               Height          =   3135
               Left            =   120
               TabIndex        =   150
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   5530
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Kriteria"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Bobot"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Matrix Perbandingan"
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
            TabIndex        =   146
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame RnkSAW 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   138
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            TabIndex        =   142
            Top             =   960
            Width           =   13695
            Begin MSComctlLib.ListView LvBobotSAW6 
               Height          =   2295
               Left            =   120
               TabIndex        =   143
               Top             =   240
               Width           =   10815
               _ExtentX        =   19076
               _ExtentY        =   4048
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
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Pendidikan"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Pengalaman"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Leadership"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Learning"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "Attention"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   6
                  Text            =   "Kinerja"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   7
                  Text            =   "Masa Kerja"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LvBobotSAW7 
               Height          =   2295
               Left            =   11040
               TabIndex        =   149
               Top             =   240
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   4048
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "SAW"
                  Object.Width           =   1764
               EndProperty
            End
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Hasil Proses Normalisasi"
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
            TabIndex        =   147
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame BnS1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   139
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            TabIndex        =   140
            Top             =   960
            Width           =   13695
            Begin MSComctlLib.ListView LvBobotSAW5 
               Height          =   2295
               Left            =   120
               TabIndex        =   141
               Top             =   240
               Width           =   13455
               _ExtentX        =   23733
               _ExtentY        =   4048
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
               NumItems        =   8
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Pendidikan"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Pengalaman"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Leadership"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   4
                  Text            =   "Learning"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "Attention"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   6
                  Text            =   "Kinerja"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   7
                  Text            =   "Masa Kerja"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "Proses Normalisasi"
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
            TabIndex        =   148
            Top             =   240
            Width           =   13680
            WordWrap        =   -1  'True
         End
      End
      Begin MSComctlLib.TreeView TreeView2 
         Height          =   4335
         Left            =   240
         TabIndex        =   137
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   7646
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
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
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   15840
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penghitungan.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penghitungan.frx":06FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penghitungan.frx":0DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penghitungan.frx":14EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penghitungan.frx":1BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penghitungan.frx":22E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LvBobotSAW1 
      Height          =   1815
      Left            =   6240
      TabIndex        =   145
      Top             =   7680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kriteria"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "ahp"
         Object.Width           =   882
      EndProperty
   End
End
Attribute VB_Name = "Penghitungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TSkala As Double
Dim TKarakter As Double
Dim TKinerja As Double
Dim TMasaKerja As Double
Dim TPendidikan As Double
Dim TPengalaman As Double

Dim TNmKarakter As Double
Dim TNmKinerja As Double
Dim TNmMasaKerja As Double
Dim TNmPendidikan As Double
Dim TNmPengalaman As Double
Dim TNmJumlah As Double

Dim BpPrioritas As Double
Dim BpPendidikan As Double
Dim BpPengalaman As Double
Dim BpKarakter As Double
Dim BpKinerja As Double
Dim BpLeadership As Double
Dim BpLearning As Double
Dim BpAttention As Double
Dim BpMsKerja As Double

Dim Total As Double

Dim a1 As Double
Dim b1 As Double
Dim c1 As Double
Dim d1 As Double
Dim e1 As Double


Dim nodxAHP As Node
Dim nodrAHP As Node
Dim nodyAHP As Node

Dim nodxSAW As Node
Dim nodrSAW As Node
Dim nodySAW As Node

Dim TotJml As Double

Private Sub Form_Load()
cNDb

'=============== Collbar =======================
    CoolBar1.Align = vbAlignTop

    CoolBar1.Bands.Clear
    CoolBar1.Bands.Add , , , , , Toolbar1
    'CoolBar1.Bands.Add , , , , , Toolbar2

    CoolBar1.Bands(1).Width = CoolBar1.Width / 2
    'CoolBar1.Bands(2).Width = CoolBar1.Width / 2
    
    TNmKarakter = 0
    TNmKinerja = 0
    TNmMasaKerja = 0
    TNmPendidikan = 0
    TNmPengalaman = 0
    TNmJumlah = 0
    
    HapusRanking
    
End Sub

Private Sub LoadListAHP()
    TreeView1.LineStyle = tvwRootLines
    Set nodxAHP = TreeView1.Nodes.Add(, , , "Matrix Perbandingan Berpasangan")
        nodxAHP.Expanded = True
            Set nodrAHP = TreeView1.Nodes.Add(nodxAHP, tvwChild, , "Matrix Normalisasi Berpasangan")
    Set nodxAHP = TreeView1.Nodes.Add(, , , "Data Mentah AHP")
    Set nodxAHP = TreeView1.Nodes.Add(, , , "Matrix Perbandingan Pendidikan")
        nodxAHP.Expanded = True
            Set nodrAHP = TreeView1.Nodes.Add(nodxAHP, tvwChild, , "Matrix Normalisasi Pendidikan")
    
    Set nodxAHP = TreeView1.Nodes.Add(, , , "Matrix Perbandingan Pengalaman")
        nodxAHP.Expanded = True
            Set nodrAHP = TreeView1.Nodes.Add(nodxAHP, tvwChild, , "Matrix Normalisasi Pengalaman")
    
    Set nodxAHP = TreeView1.Nodes.Add(, , , "Matrix Perbandingan Karakter")
        nodxAHP.Expanded = False
            Set nodrAHP = TreeView1.Nodes.Add(nodxAHP, tvwChild, , "Matrix Perbandingan Leadership")
                nodrAHP.Expanded = True
                    Set nodyAHP = TreeView1.Nodes.Add(nodrAHP, tvwChild, , "Matrix Normalisasi Leadership")
    
            Set nodrAHP = TreeView1.Nodes.Add(nodxAHP, tvwChild, , "Matrix Perbandingan Learning")
                nodrAHP.Expanded = True
                    Set nodyAHP = TreeView1.Nodes.Add(nodrAHP, tvwChild, , "Matrix Normalisasi Learning")
            
            Set nodrAHP = TreeView1.Nodes.Add(nodxAHP, tvwChild, , "Matrix Perbandingan Attention")
                nodrAHP.Expanded = True
                    Set nodyAHP = TreeView1.Nodes.Add(nodrAHP, tvwChild, , "Matrix Normalisasi Attention")
    
    Set nodxAHP = TreeView1.Nodes.Add(, , , "Matrix Perbandingan Kinerja")
        nodxAHP.Expanded = True
            Set nodrAHP = TreeView1.Nodes.Add(nodxAHP, tvwChild, , "Matrix Normalisasi Kinerja")
        
    Set nodxAHP = TreeView1.Nodes.Add(, , , "Matrix Perbandingan Masa Kerja")
        nodxAHP.Expanded = True
            Set nodrAHP = TreeView1.Nodes.Add(nodxAHP, tvwChild, , "Matrix Normalisasi Masa Kerja")
        
    Set nodxAHP = TreeView1.Nodes.Add(, , , "Nilai Perbandingan & Normalisasi Penghitungan Global")
        nodxAHP.Expanded = True
End Sub

Private Sub LoadListSAW()
    Set nodxSAW = TreeView2.Nodes.Add(, , , "Data Mentah SAW")
    Set nodxSAW = TreeView2.Nodes.Add(, , , "Matrix Perbandingan")
        nodxSAW.Expanded = True
            Set nodrSAW = TreeView2.Nodes.Add(nodxSAW, tvwChild, , "Proses Normalisasi")
    Set nodxSAW = TreeView2.Nodes.Add(, , , "Hasil Proses Normalisasi")
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Param"
            SyaratPosisi.Show 1
        Case "Bobot"
            BK.Show 1
        Case "KtR"
            Kriteria.Show 1
        Case "Proses"
            
            LoadListAHP
            LoadDatalistAHP
            BtnProsesBobot_Click
            BtnProsesPendidikan_Click
            BtnProsesPengalaman_Click
            BtnLeader_Click
            BtnLearning_Click
            BtnAttention_Click
            BtnKinerja_Click
            BtnMsKerja_Click
            LoadKarakterAHP
            LoadKarakterSAW
            'BobotKriteriaSAW
            Toolbar1.Buttons(5).Enabled = False
            Toolbar1.Buttons(6).Enabled = True
            Toolbar1.Buttons(8).Enabled = True
            LoadBobot
            Command1_Click
            RangkingAHP
            

            
            LoadListSAW
            LoadKriteriaSAW
            BtnProsesBobotSAW_Click
            
        Case "Rst"
            HapusRanking
            BersihAHP
            BersihSAW
            Toolbar1.Buttons(5).Enabled = True
            Toolbar1.Buttons(6).Enabled = False
            Toolbar1.Buttons(8).Enabled = False

        Case "CtkLpr"
            ReportRanking.Show 1
    End Select
End Sub

Private Sub BersihAHP()
'=============== Set 0 =======================
    TNmKarakter = 0
    TNmKinerja = 0
    TNmMasaKerja = 0
    TNmPendidikan = 0
    TNmPengalaman = 0
    TNmJumlah = 0
    
'=============== Treeview Nodes Clear =======================
    TreeView1.Nodes.Clear
    
'=============== List Items Clear =======================
    LvBobot2.ListItems.Clear
    LvBobot3.ListItems.Clear
    LvBobot4.ListItems.Clear
    LvBobot5.ListItems.Clear
    LvBobot6.ListItems.Clear
    LvBobot7.ListItems.Clear

    LvPendidikan2.ListItems.Clear
    LvPendidikan3.ListItems.Clear
    LvPendidikan4.ListItems.Clear
    LvPendidikan5.ListItems.Clear
    LvPendidikan6.ListItems.Clear
    LvPendidikan7.ListItems.Clear
    
    LvPengalaman2.ListItems.Clear
    LvPengalaman3.ListItems.Clear
    LvPengalaman4.ListItems.Clear
    LvPengalaman5.ListItems.Clear
    LvPengalaman6.ListItems.Clear
    LvPengalaman7.ListItems.Clear

    LvLeader2.ListItems.Clear
    LvLeader3.ListItems.Clear
    LvLeader4.ListItems.Clear
    LvLeader5.ListItems.Clear
    LvLeader6.ListItems.Clear
    LvLeader7.ListItems.Clear

    LvLearning2.ListItems.Clear
    LvLearning3.ListItems.Clear
    LvLearning4.ListItems.Clear
    LvLearning5.ListItems.Clear
    LvLearning6.ListItems.Clear
    LvLearning7.ListItems.Clear

    LvAttention2.ListItems.Clear
    LvAttention3.ListItems.Clear
    LvAttention4.ListItems.Clear
    LvAttention5.ListItems.Clear
    LvAttention6.ListItems.Clear
    LvAttention7.ListItems.Clear

    LvKinerja2.ListItems.Clear
    LvKinerja3.ListItems.Clear
    LvKinerja4.ListItems.Clear
    LvKinerja5.ListItems.Clear
    LvKinerja6.ListItems.Clear
    LvKinerja7.ListItems.Clear

    LvMsKerja2.ListItems.Clear
    LvMsKerja3.ListItems.Clear
    LvMsKerja4.ListItems.Clear
    LvMsKerja5.ListItems.Clear
    LvMsKerja6.ListItems.Clear
    LvMsKerja7.ListItems.Clear
    
    LvBobot2.ColumnHeaders.Clear
    LvBobot3.ColumnHeaders.Clear
    LvBobot4.ColumnHeaders.Clear
    LvBobot5.ColumnHeaders.Clear
    
'=============== ColumHeaders Clear =======================

    LvPendidikan2.ColumnHeaders.Clear
    LvPendidikan3.ColumnHeaders.Clear
    LvPendidikan4.ColumnHeaders.Clear
    LvPendidikan5.ColumnHeaders.Clear

    LvPengalaman2.ColumnHeaders.Clear
    LvPengalaman3.ColumnHeaders.Clear
    LvPengalaman4.ColumnHeaders.Clear
    LvPengalaman5.ColumnHeaders.Clear

    LvLeader2.ColumnHeaders.Clear
    LvLeader3.ColumnHeaders.Clear
    LvLeader4.ColumnHeaders.Clear
    LvLeader5.ColumnHeaders.Clear

    LvLearning2.ColumnHeaders.Clear
    LvLearning3.ColumnHeaders.Clear
    LvLearning4.ColumnHeaders.Clear
    LvLearning5.ColumnHeaders.Clear

    LvAttention2.ColumnHeaders.Clear
    LvAttention3.ColumnHeaders.Clear
    LvAttention4.ColumnHeaders.Clear
    LvAttention5.ColumnHeaders.Clear

    LvKinerja2.ColumnHeaders.Clear
    LvKinerja3.ColumnHeaders.Clear
    LvKinerja4.ColumnHeaders.Clear
    LvKinerja5.ColumnHeaders.Clear

    LvMsKerja2.ColumnHeaders.Clear
    LvMsKerja3.ColumnHeaders.Clear
    LvMsKerja4.ColumnHeaders.Clear
    LvMsKerja5.ColumnHeaders.Clear
    
'=============== Visible All =======================
    
    PpB.Visible = False
    PnB.Visible = False
    MpP.Visible = False
    MnP.Visible = False
    MpPe.Visible = False
    MnPe.Visible = False
    MpLs.Visible = False
    MnLs.Visible = False
    MpLr.Visible = False
    MnLr.Visible = False
    MpA.Visible = False
    MnA.Visible = False
    MpK.Visible = False
    MnK.Visible = False
    MpMk.Visible = False
    MnMk.Visible = False
    RnKAHP.Visible = False
    VScroll1.Visible = False
End Sub

Private Sub BersihSAW()
    TreeView2.Nodes.Clear
    BpS1.Visible = False
    BnS1.Visible = False
    RnkSAW.Visible = False
End Sub

'=============== Treeview =======================


Private Sub LoadDatalistAHP()
    '========== Bobot =======================
    LvBobot2.ColumnHeaders.Clear
    LvBobot2.ColumnHeaders.Add , , "Kriteria"
    LvBobot2.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    
    LvBobot3.ColumnHeaders.Clear
    LvBobot3.ColumnHeaders.Add , , "Kriteria"
    LvBobot3.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    
    LvBobot4.ColumnHeaders.Clear
    LvBobot4.ColumnHeaders.Add , , "Kriteria"
    
    LvBobot5.ColumnHeaders.Clear
    LvBobot5.ColumnHeaders.Add , , "Kriteria"
    
    LoadKriteria
    Bobot
    '========== Pendidikan =======================
    LvPendidikan2.ColumnHeaders.Add , , "Nama"
    LvPendidikan2.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvPendidikan3.ColumnHeaders.Add , , "Nama"
    LvPendidikan3.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvPendidikan4.ColumnHeaders.Add , , "Nama"
    LvPendidikan5.ColumnHeaders.Add , , "Nama"
    LoadNamaPendidikan
    KlmPendidikan
    '========== Pengalaman =======================
    LvPengalaman2.ColumnHeaders.Add , , "Nama"
    LvPengalaman2.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvPengalaman3.ColumnHeaders.Add , , "Nama"
    LvPengalaman3.ColumnHeaders.Add , , "Skala"
    LvPengalaman4.ColumnHeaders.Add , , "Nama"
    LvPengalaman5.ColumnHeaders.Add , , "Nama"
    LoadNamaPengalaman
    KlmPengalaman
    '========== Leadership =======================
    LvLeader2.ColumnHeaders.Add , , "Nama"
    LvLeader2.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvLeader3.ColumnHeaders.Add , , "Nama"
    LvLeader3.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvLeader4.ColumnHeaders.Add , , "Nama"
    LvLeader5.ColumnHeaders.Add , , "Nama"
    LoadNamaLeadership
    KlmLeadership
    '========== Learning =========================
    LvLearning2.ColumnHeaders.Add , , "Nama"
    LvLearning2.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvLearning3.ColumnHeaders.Add , , "Nama"
    LvLearning3.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvLearning4.ColumnHeaders.Add , , "Nama"
    LvLearning5.ColumnHeaders.Add , , "Nama"
    LoadNamaLearning
    KlmLearning
    '========== Attention ========================
    LvAttention2.ColumnHeaders.Add , , "Nama"
    LvAttention2.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvAttention3.ColumnHeaders.Add , , "Nama"
    LvAttention3.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvAttention4.ColumnHeaders.Add , , "Nama"
    LvAttention5.ColumnHeaders.Add , , "Nama"
    LoadNamaAttention
    KlmAttention
    '========== Kinerja ==========================
    LvKinerja2.ColumnHeaders.Add , , "Nama"
    LvKinerja2.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvKinerja3.ColumnHeaders.Add , , "Nama"
    LvKinerja3.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvKinerja4.ColumnHeaders.Add , , "Nama"
    LvKinerja5.ColumnHeaders.Add , , "Nama"
    LoadNamaKinerja
    KlmKinerja
    '========== Masa Kerja =======================
    LvMsKerja2.ColumnHeaders.Add , , "Nama"
    LvMsKerja2.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvMsKerja3.ColumnHeaders.Add , , "Nama"
    LvMsKerja3.ColumnHeaders.Add , , "Skala", , lvwColumnCenter
    LvMsKerja4.ColumnHeaders.Add , , "Nama"
    LvMsKerja5.ColumnHeaders.Add , , "Nama"
    LoadNamaMsKerja
    KlmMsKerja
End Sub


'========== Bobot =======================
Private Sub Bobot()
 
    Dim c As Integer
    For c = 1 To LvBobot1.ListItems.count
        LvBobot2.ColumnHeaders.Add , , LvBobot1.ListItems(c).Text, , lvwColumnRight
        LvBobot3.ColumnHeaders.Add , , LvBobot1.ListItems(c).Text, , lvwColumnRight
        LvBobot4.ColumnHeaders.Add , , LvBobot1.ListItems(c).Text, , lvwColumnRight
        LvBobot5.ColumnHeaders.Add , , LvBobot1.ListItems(c).Text, , lvwColumnRight
    Next
    
    LvBobot4.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
    LvBobot5.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
End Sub

Private Sub LoadKriteria()
    Dim cList As ListItem

    MySql = "SELECT nama_kriteria, ahp FROM tb_kriteria ORDER BY nama_kriteria ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LvBobot1.View = lvwReport
    LvBobot1.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LvBobot1.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
End Sub

Private Sub BobotKriteria()
    Dim cList As ListItem
    Dim cIndex As Integer
    Dim Karakter As Double
    Dim Kinerja As Double
    Dim MasaKerja As Double
    Dim Pendidikan As Double
    Dim Pengalaman As Double
    Dim Kriteria As String
    Dim Skala As Double
    LvBobot2.ListItems.Clear
    With LvBobot1
        For cIndex = 1 To .ListItems.count
            Karakter = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(1).SubItems(1))
            Kinerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(2).SubItems(1))
            MasaKerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(3).SubItems(1))
            Pendidikan = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(4).SubItems(1))
            Pengalaman = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(5).SubItems(1))
            Kriteria = .ListItems(cIndex)
            Skala = .ListItems(cIndex).SubItems(1)
            
            Set cList = LvBobot2.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Skala
                cList.SubItems(2) = Format(Karakter, "0.000")
                cList.SubItems(3) = Format(Kinerja, "0.000")
                cList.SubItems(4) = Format(MasaKerja, "0.000")
                cList.SubItems(5) = Format(Pendidikan, "0.000")
                cList.SubItems(6) = Format(Pengalaman, "0.000")
        Next
    End With
End Sub

Private Sub TotKrK()
    Dim cList As ListItem
    Dim lngIndex As Integer
        
        TSkala = 0
        TKarakter = 0
        TKinerja = 0
        TMasaKerja = 0
        TPendidikan = 0
        TPengalaman = 0
        
        For lngIndex = 1 To LvBobot2.ListItems.count
            TSkala = TSkala + LvBobot2.ListItems(lngIndex).SubItems(1)
            TKarakter = TKarakter + LvBobot2.ListItems(lngIndex).SubItems(2)
            TKinerja = TKinerja + LvBobot2.ListItems(lngIndex).SubItems(3)
            TMasaKerja = TMasaKerja + LvBobot2.ListItems(lngIndex).SubItems(4)
            TPendidikan = TPendidikan + LvBobot2.ListItems(lngIndex).SubItems(5)
            TPengalaman = TPengalaman + LvBobot2.ListItems(lngIndex).SubItems(6)
        Next
        LvBobot3.ListItems.Clear
         Set cList = LvBobot3.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = TSkala
                cList.SubItems(2) = Format(TKarakter, "0.000")
                cList.SubItems(3) = Format(TKinerja, "0.000")
                cList.SubItems(4) = Format(TMasaKerja, "0.000")
                cList.SubItems(5) = Format(TPendidikan, "0.000")
                cList.SubItems(6) = Format(TPengalaman, "0.000")
End Sub

Private Sub Normalisasi()
    Dim cListNm As ListItem
        Dim i As Integer
        Dim NmKarakter As Double
        Dim NmKinerja As Double
        Dim NmMasaKerja As Double
        Dim NmPendidikan As Double
        Dim NmPengalaman As Double
        Dim NmKriteria As String
        
        NmKinerja = 0
        NmMasaKerja = 0
        NmPendidikan = 0
        NmPengalaman = 0
        NmKriteria = 0
        LvBobot4.ListItems.Clear
        With LvBobot2
        
            For i = 1 To .ListItems.count
                NmKarakter = .ListItems(i).SubItems(2) / TKarakter
                NmKinerja = Val(.ListItems(i).SubItems(3)) / TKinerja
                NmMasaKerja = Val(.ListItems(i).SubItems(4)) / TMasaKerja
                NmPendidikan = Val(.ListItems(i).SubItems(5)) / TPendidikan
                NmPengalaman = Val(.ListItems(i).SubItems(6)) / TPengalaman
                NmKriteria = .ListItems(i)
        
                Set cListNm = LvBobot4.ListItems.Add(, , NmKriteria)
                    cListNm.SubItems(1) = Format(NmKarakter, "0.000")
                    cListNm.SubItems(2) = Format(NmKinerja, "0.000")
                    cListNm.SubItems(3) = Format(NmMasaKerja, "0.000")
                    cListNm.SubItems(4) = Format(NmPendidikan, "0.000")
                    cListNm.SubItems(5) = Format(NmPengalaman, "0.000")
                    cListNm.SubItems(6) = Format(NmKarakter + NmKinerja + NmMasaKerja + NmPendidikan + NmPengalaman, "0.000")
            Next
        End With
End Sub

Private Sub TotNm()
    Dim cList As ListItem
    Dim i As Integer
    
    TNmKarakter = 0
    TNmKinerja = 0
    TNmMasaKerja = 0
    TNmPendidikan = 0
    TNmPengalaman = 0
    TNmJumlah = 0

        For i = 1 To LvBobot4.ListItems.count
            TNmKarakter = TNmKarakter + LvBobot4.ListItems(i).SubItems(1)
            TNmKinerja = TNmKinerja + LvBobot4.ListItems(i).SubItems(2)
            TNmMasaKerja = TNmMasaKerja + LvBobot4.ListItems(i).SubItems(3)
            TNmPendidikan = TNmPendidikan + LvBobot4.ListItems(i).SubItems(4)
            TNmPengalaman = TNmPengalaman + LvBobot4.ListItems(i).SubItems(5)
            TNmJumlah = TNmJumlah + LvBobot4.ListItems(i).SubItems(6)
        Next
        LvBobot5.ListItems.Clear
         Set cList = LvBobot5.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = Format(TNmKarakter, "0.000")
                cList.SubItems(2) = Format(TNmKinerja, "0.000")
                cList.SubItems(3) = Format(TNmMasaKerja, "0.000")
                cList.SubItems(4) = Format(TNmPendidikan, "0.000")
                cList.SubItems(5) = Format(TNmPengalaman, "0.000")
                cList.SubItems(6) = Format(TNmJumlah, "0.000")
                
        BobotPrioritas
End Sub

Private Sub BobotPrioritas()
    Dim cIndexBp As Integer
    Dim cList As ListItem
    Dim Kriteria As String
    LvBobot6.ListItems.Clear
        For cIndexBp = 1 To LvBobot4.ListItems.count
            Kriteria = LvBobot4.ListItems(cIndexBp)
            BpPrioritas = LvBobot4.ListItems(cIndexBp).SubItems(6) / TNmJumlah
            
            Set cList = LvBobot6.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Format(BpPrioritas, "0.000")
        Next
       JmlBP
End Sub

Private Sub JmlBP()
    Dim cIndexJMlBP As Integer
    Dim jml As Double
    Dim cList As ListItem
        For cIndexJMlBP = 1 To LvBobot4.ListItems.count
            jml = jml + LvBobot6.ListItems(cIndexJMlBP).SubItems(1)
        Next
        LvBobot7.ListItems.Clear
        Set cList = LvBobot7.ListItems.Add(, , "Jumlah")
            cList.SubItems(1) = Format(jml, "0.000")
End Sub

Private Sub BtnProsesBobot_Click()
    BobotKriteria
    TotKrK
    Normalisasi
    TotNm
End Sub


'========== Pendidikan =======================
Private Sub KlmPendidikan()
    Dim c As Integer
    For c = 1 To LvPendidikan1.ListItems.count
        LvPendidikan2.ColumnHeaders.Add , , LvPendidikan1.ListItems(c).Text, , lvwColumnRight
        LvPendidikan3.ColumnHeaders.Add , , LvPendidikan1.ListItems(c).Text, , lvwColumnRight
        LvPendidikan4.ColumnHeaders.Add , , LvPendidikan1.ListItems(c).Text, , lvwColumnRight
        LvPendidikan5.ColumnHeaders.Add , , LvPendidikan1.ListItems(c).Text, , lvwColumnRight
    Next
    
    LvPendidikan4.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
    LvPendidikan5.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
End Sub

Private Sub LoadNamaPendidikan()
    Dim cList As ListItem

    MySql = "SELECT tb_pegawai.nama, tb_pendidikan.AHP FROM tb_pendidikan, tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_pendidikan.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LvPendidikan1.View = lvwReport
    LvPendidikan1.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LvPendidikan1.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
End Sub

Private Sub Pendidikan()
    Dim cList As ListItem
    Dim cIndex As Integer
    Dim Karakter As Double
    Dim Kinerja As Double
    Dim MasaKerja As Double
    Dim Pendidikan As Double
    Dim Pengalaman As Double
    Dim Kriteria As String
    Dim Skala As Double
        
    For cIndex = 1 To LvPendidikan1.ListItems.count
        Karakter = Val(LvPendidikan1.ListItems(cIndex).SubItems(1)) / Val(LvPendidikan1.ListItems(1).SubItems(1))
        Kinerja = Val(LvPendidikan1.ListItems(cIndex).SubItems(1)) / Val(LvPendidikan1.ListItems(2).SubItems(1))
        MasaKerja = Val(LvPendidikan1.ListItems(cIndex).SubItems(1)) / Val(LvPendidikan1.ListItems(3).SubItems(1))
        Pendidikan = Val(LvPendidikan1.ListItems(cIndex).SubItems(1)) / Val(LvPendidikan1.ListItems(4).SubItems(1))
        Pengalaman = Val(LvPendidikan1.ListItems(cIndex).SubItems(1)) / Val(LvPendidikan1.ListItems(5).SubItems(1))
        Kriteria = LvPendidikan1.ListItems(cIndex)
        Skala = LvPendidikan1.ListItems(cIndex).SubItems(1)
            
        Set cList = LvPendidikan2.ListItems.Add(, , Kriteria)
            cList.SubItems(1) = Skala
            cList.SubItems(2) = Format(Karakter, "0.000")
            cList.SubItems(3) = Format(Kinerja, "0.000")
            cList.SubItems(4) = Format(MasaKerja, "0.000")
            cList.SubItems(5) = Format(Pendidikan, "0.000")
            cList.SubItems(6) = Format(Pengalaman, "0.000")
        Next
        
End Sub

Private Sub NrmPendidikan()
    Dim cListNm As ListItem
        Dim cIndexNm As Integer
        Dim NmKarakter As Double
        Dim NmKinerja As Double
        Dim NmMasaKerja As Double
        Dim NmPendidikan As Double
        Dim NmPengalaman As Double
        Dim NmKriteria As String
        
        For cIndexNm = 1 To LvPendidikan2.ListItems.count
           NmKarakter = Val(LvPendidikan2.ListItems(cIndexNm).SubItems(2)) / TKarakter
           NmKinerja = Val(LvPendidikan2.ListItems(cIndexNm).SubItems(3)) / TKinerja
           NmMasaKerja = Val(LvPendidikan2.ListItems(cIndexNm).SubItems(4)) / TMasaKerja
           NmPendidikan = Val(LvPendidikan2.ListItems(cIndexNm).SubItems(5)) / TPendidikan
           NmPengalaman = Val(LvPendidikan2.ListItems(cIndexNm).SubItems(6)) / TPengalaman
           NmKriteria = LvPendidikan2.ListItems(cIndexNm)
           
           Set cListNm = LvPendidikan4.ListItems.Add(, , NmKriteria)
               cListNm.SubItems(1) = Format(NmKarakter, "0.000")
               cListNm.SubItems(2) = Format(NmKinerja, "0.000")
               cListNm.SubItems(3) = Format(NmMasaKerja, "0.000")
               cListNm.SubItems(4) = Format(NmPendidikan, "0.000")
               cListNm.SubItems(5) = Format(NmPengalaman, "0.000")
               cListNm.SubItems(6) = Format(NmKarakter + NmKinerja + NmMasaKerja + NmPendidikan + NmPengalaman, "0.000")
      Next
End Sub

Private Sub TotPendidikan()
    Dim cList As ListItem
    Dim lngIndex As Integer
        TSkala = 0
        TKarakter = 0
        TKinerja = 0
        TMasaKerja = 0
        TPendidikan = 0
        TPengalaman = 0
        For lngIndex = 1 To LvPendidikan2.ListItems.count
            TSkala = TSkala + LvPendidikan2.ListItems(lngIndex).SubItems(1)
            TKarakter = TKarakter + LvPendidikan2.ListItems(lngIndex).SubItems(2)
            TKinerja = TKinerja + LvPendidikan2.ListItems(lngIndex).SubItems(3)
            TMasaKerja = TMasaKerja + LvPendidikan2.ListItems(lngIndex).SubItems(4)
            TPendidikan = TPendidikan + LvPendidikan2.ListItems(lngIndex).SubItems(5)
            TPengalaman = TPengalaman + LvPendidikan2.ListItems(lngIndex).SubItems(6)
        Next
        
         Set cList = LvPendidikan3.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = TSkala
                cList.SubItems(2) = Format(TKarakter, "0.000")
                cList.SubItems(3) = Format(TKinerja, "0.000")
                cList.SubItems(4) = Format(TMasaKerja, "0.000")
                cList.SubItems(5) = Format(TPendidikan, "0.000")
                cList.SubItems(6) = Format(TPengalaman, "0.000")
        
End Sub

Private Sub TotNmPendidikan()
    Dim cList As ListItem
    Dim cIndexTNm As Integer
    
    TNmKarakter = 0
    TNmKinerja = 0
    TNmMasaKerja = 0
    TNmPendidikan = 0
    TNmPengalaman = 0
    TNmJumlah = 0
            
    With LvPendidikan4
        For cIndexTNm = 1 To .ListItems.count
            TNmKarakter = TNmKarakter + .ListItems(cIndexTNm).SubItems(1)
            TNmKinerja = TNmKinerja + .ListItems(cIndexTNm).SubItems(2)
            TNmMasaKerja = TNmMasaKerja + .ListItems(cIndexTNm).SubItems(3)
            TNmPendidikan = TNmPendidikan + .ListItems(cIndexTNm).SubItems(4)
            TNmPengalaman = TNmPengalaman + .ListItems(cIndexTNm).SubItems(5)
            TNmJumlah = TNmJumlah + .ListItems(cIndexTNm).SubItems(6)
        Next
    End With
         Set cList = LvPendidikan5.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = Format(TNmKarakter, "0.000")
                cList.SubItems(2) = Format(TNmKinerja, "0.000")
                cList.SubItems(3) = Format(TNmMasaKerja, "0.000")
                cList.SubItems(4) = Format(TNmPendidikan, "0.000")
                cList.SubItems(5) = Format(TNmPengalaman, "0.000")
                cList.SubItems(6) = Format(TNmJumlah, "0.000")
                
        BbPendidikan
End Sub

Private Sub BbPendidikan()
    Dim cIndexBp As Integer
    Dim cList As ListItem
    Dim Kriteria As String
        For cIndexBp = 1 To LvPendidikan4.ListItems.count
            Kriteria = LvPendidikan4.ListItems(cIndexBp)
            BpPendidikan = LvPendidikan4.ListItems(cIndexBp).SubItems(6) / TNmJumlah
            
            Set cList = LvPendidikan6.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Format(BpPendidikan, "0.000")
        Next
       
       BbPend
End Sub

Private Sub BbPend()
    Dim cIndexJMlBP As Integer
    Dim jml As Double
    Dim cList As ListItem
        For cIndexJMlBP = 1 To LvPendidikan4.ListItems.count
            jml = jml + LvPendidikan4.ListItems(cIndexJMlBP).SubItems(1)
        Next
        Set cList = LvPendidikan7.ListItems.Add(, , "Jumlah")
            cList.SubItems(1) = Format(jml, "0.000")
        
End Sub

Private Sub DellBbtPend()
    Dim KD As String
    KD = "PND"
    MySql = "DELETE FROM tb_bbt_ahp WHERE kode = '" & KD & "'"
    ConN.Execute MySql
End Sub

Private Sub SaveBbtPend()
    Dim KD As String
    KD = "PND"
    Dim i As Integer
    For i = 1 To LvPendidikan6.ListItems.count
        MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
        "'" & LvPendidikan6.ListItems(i) & "', " & _
        "'" & LvPendidikan6.ListItems(i).SubItems(1) & "'," & _
        "'" & KD & "')"
        ConN.Execute MySql
    Next
End Sub

Private Sub BtnProsesPendidikan_Click()
    Pendidikan
    TotPendidikan
    NrmPendidikan
    TotNmPendidikan
    DellBbtPend
    SaveBbtPend
End Sub

'========== Pengalaman =======================
Private Sub KlmPengalaman()
    Dim c As Integer
    For c = 1 To LvPengalaman1.ListItems.count
        LvPengalaman2.ColumnHeaders.Add , , LvPengalaman1.ListItems(c).Text, , lvwColumnRight
        LvPengalaman3.ColumnHeaders.Add , , LvPengalaman1.ListItems(c).Text, , lvwColumnRight
        LvPengalaman4.ColumnHeaders.Add , , LvPengalaman1.ListItems(c).Text, , lvwColumnRight
        LvPengalaman5.ColumnHeaders.Add , , LvPengalaman1.ListItems(c).Text, , lvwColumnRight
    Next
    
    LvPengalaman4.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
    LvPengalaman5.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
End Sub

Private Sub LoadNamaPengalaman()
    Dim cList As ListItem

    MySql = "SELECT tb_pegawai.nama, tb_pengalaman.AHP FROM tb_pengalaman , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_pengalaman.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LvPengalaman1.View = lvwReport
    LvPengalaman1.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LvPengalaman1.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
End Sub

Private Sub Pengalaman()
    Dim cList As ListItem
    Dim cIndex As Integer
    Dim Karakter As Double
    Dim Kinerja As Double
    Dim MasaKerja As Double
    Dim Pendidikan As Double
    Dim Pengalaman As Double
    Dim Kriteria As String
    Dim Skala As Double
        
    For cIndex = 1 To LvPengalaman1.ListItems.count
        Karakter = Val(LvPengalaman1.ListItems(cIndex).SubItems(1)) / Val(LvPengalaman1.ListItems(1).SubItems(1))
        Kinerja = Val(LvPengalaman1.ListItems(cIndex).SubItems(1)) / Val(LvPengalaman1.ListItems(2).SubItems(1))
        MasaKerja = Val(LvPengalaman1.ListItems(cIndex).SubItems(1)) / Val(LvPengalaman1.ListItems(3).SubItems(1))
        Pendidikan = Val(LvPengalaman1.ListItems(cIndex).SubItems(1)) / Val(LvPengalaman1.ListItems(4).SubItems(1))
        Pengalaman = Val(LvPengalaman1.ListItems(cIndex).SubItems(1)) / Val(LvPengalaman1.ListItems(5).SubItems(1))
        Kriteria = LvPengalaman1.ListItems(cIndex)
        Skala = LvPengalaman1.ListItems(cIndex).SubItems(1)
            
        Set cList = LvPengalaman2.ListItems.Add(, , Kriteria)
            cList.SubItems(1) = Skala
            cList.SubItems(2) = Format(Karakter, "0.000")
            cList.SubItems(3) = Format(Kinerja, "0.000")
            cList.SubItems(4) = Format(MasaKerja, "0.000")
            cList.SubItems(5) = Format(Pendidikan, "0.000")
            cList.SubItems(6) = Format(Pengalaman, "0.000")
        Next
        
End Sub

Private Sub NrmPengalaman()
    Dim cListNm As ListItem
        Dim cIndexNm As Integer
        Dim NmKarakter As Double
        Dim NmKinerja As Double
        Dim NmMasaKerja As Double
        Dim NmPendidikan As Double
        Dim NmPengalaman As Double
        Dim NmKriteria As String
        
        For cIndexNm = 1 To LvPengalaman2.ListItems.count
           NmKarakter = Val(LvPengalaman2.ListItems(cIndexNm).SubItems(2)) / TKarakter
           NmKinerja = Val(LvPengalaman2.ListItems(cIndexNm).SubItems(3)) / TKinerja
           NmMasaKerja = Val(LvPengalaman2.ListItems(cIndexNm).SubItems(4)) / TMasaKerja
           NmPendidikan = Val(LvPengalaman2.ListItems(cIndexNm).SubItems(5)) / TPendidikan
           NmPengalaman = Val(LvPengalaman2.ListItems(cIndexNm).SubItems(6)) / TPengalaman
           NmKriteria = LvPengalaman2.ListItems(cIndexNm)
           
           Set cListNm = LvPengalaman4.ListItems.Add(, , NmKriteria)
               cListNm.SubItems(1) = Format(NmKarakter, "0.000")
               cListNm.SubItems(2) = Format(NmKinerja, "0.000")
               cListNm.SubItems(3) = Format(NmMasaKerja, "0.000")
               cListNm.SubItems(4) = Format(NmPendidikan, "0.000")
               cListNm.SubItems(5) = Format(NmPengalaman, "0.000")
               cListNm.SubItems(6) = Format(NmKarakter + NmKinerja + NmMasaKerja + NmPendidikan + NmPengalaman, "0.000")
      Next
End Sub

Private Sub TotPengalaman()
    Dim cList As ListItem
    Dim lngIndex As Integer
        TSkala = 0
        TKarakter = 0
        TKinerja = 0
        TMasaKerja = 0
        TPendidikan = 0
        TPengalaman = 0
        For lngIndex = 1 To LvPengalaman2.ListItems.count
            TSkala = TSkala + LvPengalaman2.ListItems(lngIndex).SubItems(1)
            TKarakter = TKarakter + LvPengalaman2.ListItems(lngIndex).SubItems(2)
            TKinerja = TKinerja + LvPengalaman2.ListItems(lngIndex).SubItems(3)
            TMasaKerja = TMasaKerja + LvPengalaman2.ListItems(lngIndex).SubItems(4)
            TPendidikan = TPendidikan + LvPengalaman2.ListItems(lngIndex).SubItems(5)
            TPengalaman = TPengalaman + LvPengalaman2.ListItems(lngIndex).SubItems(6)
        Next
        
         Set cList = LvPengalaman3.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = TSkala
                cList.SubItems(2) = Format(TKarakter, "0.000")
                cList.SubItems(3) = Format(TKinerja, "0.000")
                cList.SubItems(4) = Format(TMasaKerja, "0.000")
                cList.SubItems(5) = Format(TPendidikan, "0.000")
                cList.SubItems(6) = Format(TPengalaman, "0.000")
        
End Sub

Private Sub TotNmPengalaman()
    Dim cList As ListItem
    
    TNmKarakter = 0
    TNmKinerja = 0
    TNmMasaKerja = 0
    TNmPendidikan = 0
    TNmPengalaman = 0
    TNmJumlah = 0

    Dim cIndexTNm As Integer
        For cIndexTNm = 1 To LvPengalaman4.ListItems.count
            TNmKarakter = TNmKarakter + LvPengalaman4.ListItems(cIndexTNm).SubItems(1)
            TNmKinerja = TNmKinerja + LvPengalaman4.ListItems(cIndexTNm).SubItems(2)
            TNmMasaKerja = TNmMasaKerja + LvPengalaman4.ListItems(cIndexTNm).SubItems(3)
            TNmPendidikan = TNmPendidikan + LvPengalaman4.ListItems(cIndexTNm).SubItems(4)
            TNmPengalaman = TNmPengalaman + LvPengalaman4.ListItems(cIndexTNm).SubItems(5)
            TNmJumlah = TNmJumlah + LvPengalaman4.ListItems(cIndexTNm).SubItems(6)
        Next
        
         Set cList = LvPengalaman5.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = Format(TNmKarakter, "0.000")
                cList.SubItems(2) = Format(TNmKinerja, "0.000")
                cList.SubItems(3) = Format(TNmMasaKerja, "0.000")
                cList.SubItems(4) = Format(TNmPendidikan, "0.000")
                cList.SubItems(5) = Format(TNmPengalaman, "0.000")
                cList.SubItems(6) = Format(TNmJumlah, "0.000")
                
        BbPengalaman
End Sub

Private Sub BbPengalaman()
    Dim cIndexBp As Integer
    Dim cList As ListItem
    Dim Kriteria As String
        For cIndexBp = 1 To LvPengalaman4.ListItems.count
            Kriteria = LvPengalaman4.ListItems(cIndexBp)
            BpPengalaman = LvPengalaman4.ListItems(cIndexBp).SubItems(6) / TNmJumlah
            
            Set cList = LvPengalaman6.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Format(BpPengalaman, "0.000")
        Next
       BbPeng
End Sub

Private Sub BbPeng()
    Dim cIndexJMlBP As Integer
    Dim jml As Double
    Dim cList As ListItem
        For cIndexJMlBP = 1 To LvPengalaman4.ListItems.count
            jml = jml + LvPengalaman4.ListItems(cIndexJMlBP).SubItems(1)
        Next
        Set cList = LvPengalaman7.ListItems.Add(, , "Jumlah")
            cList.SubItems(1) = Format(jml, "0.000")
End Sub

Private Sub DellBbtPeng()
    Dim KD As String
    KD = "PNG"
    MySql = "DELETE FROM tb_bbt_ahp WHERE kode = '" & KD & "'"
    ConN.Execute MySql
End Sub

Private Sub SaveBbtPeng()
    Dim KD As String
    KD = "PNG"
    Dim i As Integer
    With LvPengalaman6
        For i = 1 To .ListItems.count
            MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
            "'" & .ListItems(i) & "', " & _
            "'" & .ListItems(i).SubItems(1) & "'," & _
            "'" & KD & "')"
            ConN.Execute MySql
        Next
    End With
End Sub

Private Sub BtnProsesPengalaman_Click()
    Pengalaman
    TotPengalaman
    NrmPengalaman
    TotNmPengalaman
    DellBbtPeng
    SaveBbtPeng
End Sub

'========== Leadership =======================
Private Sub KlmLeadership()
    Dim c As Integer
    For c = 1 To LvLeader1.ListItems.count
        LvLeader2.ColumnHeaders.Add , , LvLeader1.ListItems(c).Text, , lvwColumnRight
        LvLeader3.ColumnHeaders.Add , , LvLeader1.ListItems(c).Text, , lvwColumnRight
        LvLeader4.ColumnHeaders.Add , , LvLeader1.ListItems(c).Text, , lvwColumnRight
        LvLeader5.ColumnHeaders.Add , , LvLeader1.ListItems(c).Text, , lvwColumnRight
    Next
    
    LvLeader4.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
    LvLeader5.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
End Sub

Private Sub LoadNamaLeadership()
    Dim cList As ListItem

    MySql = "SELECT tb_pegawai.nama, tb_karakter.AHP1 FROM tb_karakter , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LvLeader1.View = lvwReport
    LvLeader1.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LvLeader1.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
End Sub

Private Sub Leadership()
    Dim cList As ListItem
    Dim cIndex As Integer
    Dim Karakter As Double
    Dim Kinerja As Double
    Dim MasaKerja As Double
    Dim Pendidikan As Double
    Dim Pengalaman As Double
    Dim Kriteria As String
    Dim Skala As Double
    With LvLeader1
        For cIndex = 1 To .ListItems.count
            Karakter = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(1).SubItems(1))
            Kinerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(2).SubItems(1))
            MasaKerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(3).SubItems(1))
            Pendidikan = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(4).SubItems(1))
            Pengalaman = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(5).SubItems(1))
            Kriteria = .ListItems(cIndex)
            Skala = .ListItems(cIndex).SubItems(1)
   
            Set cList = LvLeader2.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Skala
                cList.SubItems(2) = Format(Karakter, "0.000")
                cList.SubItems(3) = Format(Kinerja, "0.000")
                cList.SubItems(4) = Format(MasaKerja, "0.000")
                cList.SubItems(5) = Format(Pendidikan, "0.000")
                cList.SubItems(6) = Format(Pengalaman, "0.000")
        Next
    End With
End Sub

Private Sub NrmLeadership()
    Dim cListNm As ListItem
        Dim cIndexNm As Integer
        Dim NmKarakter As Double
        Dim NmKinerja As Double
        Dim NmMasaKerja As Double
        Dim NmPendidikan As Double
        Dim NmPengalaman As Double
        Dim NmKriteria As String
        With LvLeader2
            For cIndexNm = 1 To .ListItems.count
                NmKriteria = .ListItems(cIndexNm)
                NmKarakter = .ListItems(cIndexNm).SubItems(2) / TKarakter
                NmKinerja = .ListItems(cIndexNm).SubItems(3) / TKinerja
                NmMasaKerja = .ListItems(cIndexNm).SubItems(4) / TMasaKerja
                NmPendidikan = .ListItems(cIndexNm).SubItems(5) / TPendidikan
                NmPengalaman = .ListItems(cIndexNm).SubItems(6) / TPengalaman
        
                Set cListNm = LvLeader4.ListItems.Add(, , NmKriteria)
                    cListNm.SubItems(1) = Format(NmKarakter, "0.000")
                    cListNm.SubItems(2) = Format(NmKinerja, "0.000")
                    cListNm.SubItems(3) = Format(NmMasaKerja, "0.000")
                    cListNm.SubItems(4) = Format(NmPendidikan, "0.000")
                    cListNm.SubItems(5) = Format(NmPengalaman, "0.000")
                    cListNm.SubItems(6) = Format(NmKarakter + NmKinerja + NmMasaKerja + NmPendidikan + NmPengalaman, "0.000")
            Next
        End With
End Sub

Private Sub TotLeadership()
    Dim cList As ListItem
    Dim lngIndex As Integer
        
        TSkala = 0
        TKarakter = 0
        TKinerja = 0
        TMasaKerja = 0
        TPendidikan = 0
        TPengalaman = 0
        
        With LvLeader2
        For lngIndex = 1 To .ListItems.count
            TSkala = TSkala + .ListItems(lngIndex).SubItems(1)
            TKarakter = TKarakter + .ListItems(lngIndex).SubItems(2)
            TKinerja = TKinerja + .ListItems(lngIndex).SubItems(3)
            TMasaKerja = TMasaKerja + .ListItems(lngIndex).SubItems(4)
            TPendidikan = TPendidikan + .ListItems(lngIndex).SubItems(5)
            TPengalaman = TPengalaman + .ListItems(lngIndex).SubItems(6)
        Next
        End With
         Set cList = LvLeader3.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = TSkala
                cList.SubItems(2) = Format(TKarakter, "0.000")
                cList.SubItems(3) = Format(TKinerja, "0.000")
                cList.SubItems(4) = Format(TMasaKerja, "0.000")
                cList.SubItems(5) = Format(TPendidikan, "0.000")
                cList.SubItems(6) = Format(TPengalaman, "0.000")
        
End Sub

Private Sub TotNrmLeadership()
    Dim cList As ListItem
    
    TNmKarakter = 0
    TNmKinerja = 0
    TNmMasaKerja = 0
    TNmPendidikan = 0
    TNmPengalaman = 0
    TNmJumlah = 0

    Dim cIndexTNm As Integer
    With LvLeader4
        For cIndexTNm = 1 To LvLeader4.ListItems.count
            TNmKarakter = TNmKarakter + LvLeader4.ListItems(cIndexTNm).SubItems(1)
            TNmKinerja = TNmKinerja + LvLeader4.ListItems(cIndexTNm).SubItems(2)
            TNmMasaKerja = TNmMasaKerja + LvLeader4.ListItems(cIndexTNm).SubItems(3)
            TNmPendidikan = TNmPendidikan + LvLeader4.ListItems(cIndexTNm).SubItems(4)
            TNmPengalaman = TNmPengalaman + LvLeader4.ListItems(cIndexTNm).SubItems(5)
            TNmJumlah = TNmJumlah + LvLeader4.ListItems(cIndexTNm).SubItems(6)
        Next
    End With
         Set cList = LvLeader5.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = Format(TNmKarakter, "0.000")
                cList.SubItems(2) = Format(TNmKinerja, "0.000")
                cList.SubItems(3) = Format(TNmMasaKerja, "0.000")
                cList.SubItems(4) = Format(TNmPendidikan, "0.000")
                cList.SubItems(5) = Format(TNmPengalaman, "0.000")
                cList.SubItems(6) = Format(TNmJumlah, "0.000")
                
        LeaderPrioritas
End Sub

Private Sub LeaderPrioritas()
    Dim cIndexBp As Integer
    Dim cList As ListItem
    Dim Kriteria As String
        For cIndexBp = 1 To LvLeader4.ListItems.count
            Kriteria = LvLeader4.ListItems(cIndexBp)
            BpLeadership = LvLeader4.ListItems(cIndexBp).SubItems(6) / TNmJumlah
            
            Set cList = LvLeader6.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Format(BpLeadership, "0.000")
        Next
       JmlLP
End Sub

Private Sub JmlLP()
    Dim cIndexJMlBP As Integer
    Dim jml As Double
    Dim cList As ListItem
        For cIndexJMlBP = 1 To LvLeader4.ListItems.count
            jml = jml + LvLeader6.ListItems(cIndexJMlBP).SubItems(1)
        Next
        Set cList = LvLeader7.ListItems.Add(, , "Jumlah")
            cList.SubItems(1) = Format(jml, "0.000")
End Sub

Private Sub DellBbtLdr()
    Dim KD As String
    KD = "LDR"
    MySql = "DELETE FROM tb_bbt_ahp WHERE kode = '" & KD & "'"
    ConN.Execute MySql
End Sub

Private Sub SaveBbtLdr()
    Dim KD As String
    KD = "LDR"
    Dim i As Integer
    With LvLeader6
        For i = 1 To .ListItems.count
            MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
            "'" & .ListItems(i) & "', " & _
            "'" & .ListItems(i).SubItems(1) & "'," & _
            "'" & KD & "')"
            ConN.Execute MySql
        Next
        'BtnLeader.Enabled = False
    End With
End Sub


Private Sub BtnLeader_Click()
    Leadership
    TotLeadership
    NrmLeadership
    TotNrmLeadership
    DellBbtLdr
    SaveBbtLdr
End Sub

'========== Learning =========================
Private Sub KlmLearning()
    Dim c As Integer
    For c = 1 To LvLearning1.ListItems.count
        LvLearning2.ColumnHeaders.Add , , LvLearning1.ListItems(c).Text, , lvwColumnRight
        LvLearning3.ColumnHeaders.Add , , LvLearning1.ListItems(c).Text, , lvwColumnRight
        LvLearning4.ColumnHeaders.Add , , LvLearning1.ListItems(c).Text, , lvwColumnRight
        LvLearning5.ColumnHeaders.Add , , LvLearning1.ListItems(c).Text, , lvwColumnRight
    Next
    
    LvLearning4.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
    LvLearning5.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
End Sub

Private Sub LoadNamaLearning()
    Dim cList As ListItem

    MySql = "SELECT tb_pegawai.nama, tb_karakter.AHP2 FROM tb_karakter , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LvLearning1.View = lvwReport
    LvLearning1.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LvLearning1.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
End Sub

Private Sub Learning()
    Dim cList As ListItem
    Dim cIndex As Integer
    Dim Karakter As Double
    Dim Kinerja As Double
    Dim MasaKerja As Double
    Dim Pendidikan As Double
    Dim Pengalaman As Double
    Dim Kriteria As String
    Dim Skala As Double
    With LvLearning1
        For cIndex = 1 To .ListItems.count
            Karakter = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(1).SubItems(1))
            Kinerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(2).SubItems(1))
            MasaKerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(3).SubItems(1))
            Pendidikan = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(4).SubItems(1))
            Pengalaman = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(5).SubItems(1))
            Kriteria = .ListItems(cIndex)
            Skala = .ListItems(cIndex).SubItems(1)
   
            Set cList = LvLearning2.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Skala
                cList.SubItems(2) = Format(Karakter, "0.000")
                cList.SubItems(3) = Format(Kinerja, "0.000")
                cList.SubItems(4) = Format(MasaKerja, "0.000")
                cList.SubItems(5) = Format(Pendidikan, "0.000")
                cList.SubItems(6) = Format(Pengalaman, "0.000")
        Next
    End With
End Sub

Private Sub NrmLearning()
    Dim cListNm As ListItem
        Dim cIndexNm As Integer
        Dim NmKarakter As Double
        Dim NmKinerja As Double
        Dim NmMasaKerja As Double
        Dim NmPendidikan As Double
        Dim NmPengalaman As Double
        Dim NmKriteria As String
        With LvLearning2
            For cIndexNm = 1 To .ListItems.count
                NmKarakter = Val(.ListItems(cIndexNm).SubItems(2)) / TKarakter
                NmKinerja = Val(.ListItems(cIndexNm).SubItems(3)) / TKinerja
                NmMasaKerja = Val(.ListItems(cIndexNm).SubItems(4)) / TMasaKerja
                NmPendidikan = Val(.ListItems(cIndexNm).SubItems(5)) / TPendidikan
                NmPengalaman = Val(.ListItems(cIndexNm).SubItems(6)) / TPengalaman
                NmKriteria = .ListItems(cIndexNm)
        
                Set cListNm = LvLearning4.ListItems.Add(, , NmKriteria)
                    cListNm.SubItems(1) = Format(NmKarakter, "0.000")
                    cListNm.SubItems(2) = Format(NmKinerja, "0.000")
                    cListNm.SubItems(3) = Format(NmMasaKerja, "0.000")
                    cListNm.SubItems(4) = Format(NmPendidikan, "0.000")
                    cListNm.SubItems(5) = Format(NmPengalaman, "0.000")
                    cListNm.SubItems(6) = Format(NmKarakter + NmKinerja + NmMasaKerja + NmPendidikan + NmPengalaman, "0.000")
            Next
        End With
End Sub

Private Sub TotLearning()
    Dim cList As ListItem
    Dim lngIndex As Integer
        
        TSkala = 0
        TKarakter = 0
        TKinerja = 0
        TMasaKerja = 0
        TPendidikan = 0
        TPengalaman = 0
        With LvLearning2
        For lngIndex = 1 To .ListItems.count
            TSkala = TSkala + .ListItems(lngIndex).SubItems(1)
            TKarakter = TKarakter + .ListItems(lngIndex).SubItems(2)
            TKinerja = TKinerja + .ListItems(lngIndex).SubItems(3)
            TMasaKerja = TMasaKerja + .ListItems(lngIndex).SubItems(4)
            TPendidikan = TPendidikan + .ListItems(lngIndex).SubItems(5)
            TPengalaman = TPengalaman + .ListItems(lngIndex).SubItems(6)
        Next
        End With
         Set cList = LvLearning3.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = TSkala
                cList.SubItems(2) = Format(TKarakter, "0.000")
                cList.SubItems(3) = Format(TKinerja, "0.000")
                cList.SubItems(4) = Format(TMasaKerja, "0.000")
                cList.SubItems(5) = Format(TPendidikan, "0.000")
                cList.SubItems(6) = Format(TPengalaman, "0.000")
        
End Sub

Private Sub TotNrmLearning()
    Dim cList As ListItem
    
    TNmKarakter = 0
    TNmKinerja = 0
    TNmMasaKerja = 0
    TNmPendidikan = 0
    TNmPengalaman = 0
    TNmJumlah = 0
    
    Dim cIndexTNm As Integer
    With LvLearning4
        For cIndexTNm = 1 To .ListItems.count
            TNmKarakter = TNmKarakter + .ListItems(cIndexTNm).SubItems(1)
            TNmKinerja = TNmKinerja + .ListItems(cIndexTNm).SubItems(2)
            TNmMasaKerja = TNmMasaKerja + .ListItems(cIndexTNm).SubItems(3)
            TNmPendidikan = TNmPendidikan + .ListItems(cIndexTNm).SubItems(4)
            TNmPengalaman = TNmPengalaman + .ListItems(cIndexTNm).SubItems(5)
            TNmJumlah = TNmJumlah + .ListItems(cIndexTNm).SubItems(6)
        Next
    End With
         Set cList = LvLearning5.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = Format(TNmKarakter, "0.000")
                cList.SubItems(2) = Format(TNmKinerja, "0.000")
                cList.SubItems(3) = Format(TNmMasaKerja, "0.000")
                cList.SubItems(4) = Format(TNmPendidikan, "0.000")
                cList.SubItems(5) = Format(TNmPengalaman, "0.000")
                cList.SubItems(6) = Format(TNmJumlah, "0.000")
                
        LearningPrioritas
End Sub

Private Sub LearningPrioritas()
    Dim cIndexBp As Integer
    Dim cList As ListItem
    Dim Kriteria As String
        For cIndexBp = 1 To LvLearning4.ListItems.count
            Kriteria = LvLearning4.ListItems(cIndexBp)
            BpLearning = LvLearning4.ListItems(cIndexBp).SubItems(6) / TNmJumlah
            
            Set cList = LvLearning6.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Format(BpLearning, "0.000")
        Next
       JmlLeP
End Sub

Private Sub JmlLeP()
    Dim cIndexJMlBP As Integer
    Dim jml As Double
    Dim cList As ListItem
        For cIndexJMlBP = 1 To LvLearning4.ListItems.count
            jml = jml + LvLearning6.ListItems(cIndexJMlBP).SubItems(1)
        Next
        Set cList = LvLearning7.ListItems.Add(, , "Jumlah")
            cList.SubItems(1) = Format(jml, "0.000")
End Sub

Private Sub DellBbtLnr()
    Dim KD As String
    KD = "LNR"
    MySql = "DELETE FROM tb_bbt_ahp WHERE kode = '" & KD & "'"
    ConN.Execute MySql
End Sub

Private Sub SaveBbtLnr()
    Dim KD As String
    KD = "LNR"
    Dim i As Integer
    With LvLearning6
        For i = 1 To .ListItems.count
            MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
            "'" & .ListItems(i) & "', " & _
            "'" & .ListItems(i).SubItems(1) & "'," & _
            "'" & KD & "')"
            ConN.Execute MySql
        Next
        'BtnLearning.Enabled = False
    End With
End Sub

Private Sub BtnLearning_Click()
    Learning
    TotLearning
    NrmLearning
    TotNrmLearning
    DellBbtLnr
    SaveBbtLnr
End Sub
'========== Attention ========================
Private Sub KlmAttention()
    Dim c As Integer
    For c = 1 To LvAttention1.ListItems.count
        LvAttention2.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
        LvAttention3.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
        LvAttention4.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
        LvAttention5.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
    Next
    
    LvAttention4.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
    LvAttention5.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
End Sub

Private Sub LoadNamaAttention()
    Dim cList As ListItem

    MySql = "SELECT tb_pegawai.nama, tb_karakter.AHP3 FROM tb_karakter , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LvAttention1.View = lvwReport
    LvAttention1.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LvAttention1.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
End Sub

Private Sub Attention()
    Dim cList As ListItem
    Dim cIndex As Integer
    Dim Karakter As Double
    Dim Kinerja As Double
    Dim MasaKerja As Double
    Dim Pendidikan As Double
    Dim Pengalaman As Double
    Dim Kriteria As String
    Dim Skala As Double
    With LvAttention1
        For cIndex = 1 To .ListItems.count
            Karakter = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(1).SubItems(1))
            Kinerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(2).SubItems(1))
            MasaKerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(3).SubItems(1))
            Pendidikan = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(4).SubItems(1))
            Pengalaman = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(5).SubItems(1))
            Kriteria = .ListItems(cIndex)
            Skala = .ListItems(cIndex).SubItems(1)
   
            Set cList = LvAttention2.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Skala
                cList.SubItems(2) = Format(Karakter, "0.000")
                cList.SubItems(3) = Format(Kinerja, "0.000")
                cList.SubItems(4) = Format(MasaKerja, "0.000")
                cList.SubItems(5) = Format(Pendidikan, "0.000")
                cList.SubItems(6) = Format(Pengalaman, "0.000")
        Next
    End With
End Sub

Private Sub NrmAttention()
    Dim cListNm As ListItem
        Dim cIndexNm As Integer
        Dim NmKarakter As Double
        Dim NmKinerja As Double
        Dim NmMasaKerja As Double
        Dim NmPendidikan As Double
        Dim NmPengalaman As Double
        Dim NmKriteria As String
        With LvAttention2
            For cIndexNm = 1 To .ListItems.count
                NmKarakter = Val(.ListItems(cIndexNm).SubItems(2)) / TKarakter
                NmKinerja = Val(.ListItems(cIndexNm).SubItems(3)) / TKinerja
                NmMasaKerja = Val(.ListItems(cIndexNm).SubItems(4)) / TMasaKerja
                NmPendidikan = Val(.ListItems(cIndexNm).SubItems(5)) / TPendidikan
                NmPengalaman = Val(.ListItems(cIndexNm).SubItems(6)) / TPengalaman
                NmKriteria = .ListItems(cIndexNm)
        
                Set cListNm = LvAttention4.ListItems.Add(, , NmKriteria)
                    cListNm.SubItems(1) = Format(NmKarakter, "0.000")
                    cListNm.SubItems(2) = Format(NmKinerja, "0.000")
                    cListNm.SubItems(3) = Format(NmMasaKerja, "0.000")
                    cListNm.SubItems(4) = Format(NmPendidikan, "0.000")
                    cListNm.SubItems(5) = Format(NmPengalaman, "0.000")
                    cListNm.SubItems(6) = Format(NmKarakter + NmKinerja + NmMasaKerja + NmPendidikan + NmPengalaman, "0.000")
            Next
        End With
End Sub

Private Sub TotAttention()
    Dim cList As ListItem
    Dim lngIndex As Integer
        
        TSkala = 0
        TKarakter = 0
        TKinerja = 0
        TMasaKerja = 0
        TPendidikan = 0
        TPengalaman = 0
        With LvAttention2
        For lngIndex = 1 To .ListItems.count
            TSkala = TSkala + .ListItems(lngIndex).SubItems(1)
            TKarakter = TKarakter + .ListItems(lngIndex).SubItems(2)
            TKinerja = TKinerja + .ListItems(lngIndex).SubItems(3)
            TMasaKerja = TMasaKerja + .ListItems(lngIndex).SubItems(4)
            TPendidikan = TPendidikan + .ListItems(lngIndex).SubItems(5)
            TPengalaman = TPengalaman + .ListItems(lngIndex).SubItems(6)
        Next
        End With
         Set cList = LvAttention3.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = TSkala
                cList.SubItems(2) = Format(TKarakter, "0.000")
                cList.SubItems(3) = Format(TKinerja, "0.000")
                cList.SubItems(4) = Format(TMasaKerja, "0.000")
                cList.SubItems(5) = Format(TPendidikan, "0.000")
                cList.SubItems(6) = Format(TPengalaman, "0.000")
        
End Sub

Private Sub TotNrmAttention()
    Dim cList As ListItem
    
    TNmKarakter = 0
    TNmKinerja = 0
    TNmMasaKerja = 0
    TNmPendidikan = 0
    TNmPengalaman = 0
    TNmJumlah = 0

    Dim cIndexTNm As Integer
    With LvAttention4
        For cIndexTNm = 1 To .ListItems.count
            TNmKarakter = TNmKarakter + .ListItems(cIndexTNm).SubItems(1)
            TNmKinerja = TNmKinerja + .ListItems(cIndexTNm).SubItems(2)
            TNmMasaKerja = TNmMasaKerja + .ListItems(cIndexTNm).SubItems(3)
            TNmPendidikan = TNmPendidikan + .ListItems(cIndexTNm).SubItems(4)
            TNmPengalaman = TNmPengalaman + .ListItems(cIndexTNm).SubItems(5)
            TNmJumlah = TNmJumlah + .ListItems(cIndexTNm).SubItems(6)
        Next
    End With
         Set cList = LvAttention5.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = Format(TNmKarakter, "0.000")
                cList.SubItems(2) = Format(TNmKinerja, "0.000")
                cList.SubItems(3) = Format(TNmMasaKerja, "0.000")
                cList.SubItems(4) = Format(TNmPendidikan, "0.000")
                cList.SubItems(5) = Format(TNmPengalaman, "0.000")
                cList.SubItems(6) = Format(TNmJumlah, "0.000")
                
        AttentionPrioritas
End Sub

Private Sub AttentionPrioritas()
    Dim cIndexBp As Integer
    Dim cList As ListItem
    Dim Kriteria As String
        For cIndexBp = 1 To LvAttention4.ListItems.count
            Kriteria = LvAttention4.ListItems(cIndexBp)
            BpAttention = LvAttention4.ListItems(cIndexBp).SubItems(6) / TNmJumlah
            
            Set cList = LvAttention6.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Format(BpAttention, "0.000")
        Next
       JmlAP
End Sub

Private Sub JmlAP()
    Dim cIndexJMlBP As Integer
    Dim jml As Double
    Dim cList As ListItem
        For cIndexJMlBP = 1 To LvAttention4.ListItems.count
            jml = jml + LvAttention6.ListItems(cIndexJMlBP).SubItems(1)
        Next
        Set cList = LvAttention7.ListItems.Add(, , "Jumlah")
            cList.SubItems(1) = Format(jml, "0.000")
End Sub

Private Sub DellBbtAtn()
    Dim KD As String
    KD = "ATN"
    MySql = "DELETE FROM tb_bbt_ahp WHERE kode = '" & KD & "'"
    ConN.Execute MySql
End Sub

Private Sub SaveBbtAtn()
    Dim KD As String
    KD = "ATN"
    Dim i As Integer
    With LvAttention6
        For i = 1 To .ListItems.count
            MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
            "'" & .ListItems(i) & "', " & _
            "'" & .ListItems(i).SubItems(1) & "'," & _
            "'" & KD & "')"
            ConN.Execute MySql
        Next
       ' BtnAttention.Enabled = False
    End With
End Sub

Private Sub BtnAttention_Click()
    Attention
    TotAttention
    NrmAttention
    TotNrmAttention
    DellBbtAtn
    SaveBbtAtn
End Sub
'========== Kinerja ==========================
Private Sub KlmKinerja()
    Dim c As Integer
    For c = 1 To LvKinerja1.ListItems.count
        LvKinerja2.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
        LvKinerja3.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
        LvKinerja4.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
        LvKinerja5.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
    Next
    
    LvKinerja4.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
    LvKinerja5.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
End Sub

Private Sub LoadNamaKinerja()
    Dim cList As ListItem

    MySql = "SELECT tb_pegawai.nama, tb_kinerja.AHP FROM tb_kinerja , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_kinerja.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LvKinerja1.View = lvwReport
    LvKinerja1.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LvKinerja1.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
End Sub

Private Sub Kinerja()
    Dim cList As ListItem
    Dim cIndex As Integer
    Dim Karakter As Double
    Dim Kinerja As Double
    Dim MasaKerja As Double
    Dim Pendidikan As Double
    Dim Pengalaman As Double
    Dim Kriteria As String
    Dim Skala As Double
    With LvKinerja1
        For cIndex = 1 To .ListItems.count
            Karakter = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(1).SubItems(1))
            Kinerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(2).SubItems(1))
            MasaKerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(3).SubItems(1))
            Pendidikan = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(4).SubItems(1))
            Pengalaman = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(5).SubItems(1))
            Kriteria = .ListItems(cIndex)
            Skala = .ListItems(cIndex).SubItems(1)
   
            Set cList = LvKinerja2.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Skala
                cList.SubItems(2) = Format(Karakter, "0.000")
                cList.SubItems(3) = Format(Kinerja, "0.000")
                cList.SubItems(4) = Format(MasaKerja, "0.000")
                cList.SubItems(5) = Format(Pendidikan, "0.000")
                cList.SubItems(6) = Format(Pengalaman, "0.000")
        Next
    End With
End Sub

Private Sub NrmKinerja()
    Dim cListNm As ListItem
        Dim cIndexNm As Integer
        Dim NmKarakter As Double
        Dim NmKinerja As Double
        Dim NmMasaKerja As Double
        Dim NmPendidikan As Double
        Dim NmPengalaman As Double
        Dim NmKriteria As String
        With LvKinerja2
            For cIndexNm = 1 To .ListItems.count
                NmKarakter = Val(.ListItems(cIndexNm).SubItems(2)) / TKarakter
                NmKinerja = Val(.ListItems(cIndexNm).SubItems(3)) / TKinerja
                NmMasaKerja = Val(.ListItems(cIndexNm).SubItems(4)) / TMasaKerja
                NmPendidikan = Val(.ListItems(cIndexNm).SubItems(5)) / TPendidikan
                NmPengalaman = Val(.ListItems(cIndexNm).SubItems(6)) / TPengalaman
                NmKriteria = .ListItems(cIndexNm)
        
                Set cListNm = LvKinerja4.ListItems.Add(, , NmKriteria)
                    cListNm.SubItems(1) = Format(NmKarakter, "0.000")
                    cListNm.SubItems(2) = Format(NmKinerja, "0.000")
                    cListNm.SubItems(3) = Format(NmMasaKerja, "0.000")
                    cListNm.SubItems(4) = Format(NmPendidikan, "0.000")
                    cListNm.SubItems(5) = Format(NmPengalaman, "0.000")
                    cListNm.SubItems(6) = Format(NmKarakter + NmKinerja + NmMasaKerja + NmPendidikan + NmPengalaman, "0.000")
            Next
        End With
End Sub

Private Sub TotKinerja()
    Dim cList As ListItem
    Dim lngIndex As Integer
        
        TSkala = 0
        TKarakter = 0
        TKinerja = 0
        TMasaKerja = 0
        TPendidikan = 0
        TPengalaman = 0
        With LvKinerja2
        For lngIndex = 1 To .ListItems.count
            TSkala = TSkala + .ListItems(lngIndex).SubItems(1)
            TKarakter = TKarakter + .ListItems(lngIndex).SubItems(2)
            TKinerja = TKinerja + .ListItems(lngIndex).SubItems(3)
            TMasaKerja = TMasaKerja + .ListItems(lngIndex).SubItems(4)
            TPendidikan = TPendidikan + .ListItems(lngIndex).SubItems(5)
            TPengalaman = TPengalaman + .ListItems(lngIndex).SubItems(6)
        Next
        End With
         Set cList = LvKinerja3.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = TSkala
                cList.SubItems(2) = Format(TKarakter, "0.000")
                cList.SubItems(3) = Format(TKinerja, "0.000")
                cList.SubItems(4) = Format(TMasaKerja, "0.000")
                cList.SubItems(5) = Format(TPendidikan, "0.000")
                cList.SubItems(6) = Format(TPengalaman, "0.000")
        
End Sub

Private Sub TotNrmKinerja()
    Dim cList As ListItem
    
    TNmKarakter = 0
    TNmKinerja = 0
    TNmMasaKerja = 0
    TNmPendidikan = 0
    TNmPengalaman = 0
    TNmJumlah = 0

    Dim cIndexTNm As Integer
    With LvKinerja4
        For cIndexTNm = 1 To .ListItems.count
            TNmKarakter = TNmKarakter + .ListItems(cIndexTNm).SubItems(1)
            TNmKinerja = TNmKinerja + .ListItems(cIndexTNm).SubItems(2)
            TNmMasaKerja = TNmMasaKerja + .ListItems(cIndexTNm).SubItems(3)
            TNmPendidikan = TNmPendidikan + .ListItems(cIndexTNm).SubItems(4)
            TNmPengalaman = TNmPengalaman + .ListItems(cIndexTNm).SubItems(5)
            TNmJumlah = TNmJumlah + .ListItems(cIndexTNm).SubItems(6)
        Next
    End With
         Set cList = LvKinerja5.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = Format(TNmKarakter, "0.000")
                cList.SubItems(2) = Format(TNmKinerja, "0.000")
                cList.SubItems(3) = Format(TNmMasaKerja, "0.000")
                cList.SubItems(4) = Format(TNmPendidikan, "0.000")
                cList.SubItems(5) = Format(TNmPengalaman, "0.000")
                cList.SubItems(6) = Format(TNmJumlah, "0.000")
                
        KinerjaPrioritas
End Sub

Private Sub KinerjaPrioritas()
    Dim cIndexBp As Integer
    Dim cList As ListItem
    Dim Kriteria As String
        For cIndexBp = 1 To LvKinerja4.ListItems.count
            Kriteria = LvKinerja4.ListItems(cIndexBp)
            BpKinerja = LvKinerja4.ListItems(cIndexBp).SubItems(6) / TNmJumlah
            
            Set cList = LvKinerja6.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Format(BpKinerja, "0.000")
        Next
       JmlKP
End Sub

Private Sub JmlKP()
    Dim cIndexJMlBP As Integer
    Dim jml As Double
    Dim cList As ListItem
        For cIndexJMlBP = 1 To LvKinerja4.ListItems.count
            jml = jml + LvKinerja6.ListItems(cIndexJMlBP).SubItems(1)
        Next
        Set cList = LvKinerja7.ListItems.Add(, , "Jumlah")
            cList.SubItems(1) = Format(jml, "0.000")
End Sub

Private Sub DellBbtKnj()
    Dim KD As String
    KD = "KNJ"
    MySql = "DELETE FROM tb_bbt_ahp WHERE kode = '" & KD & "'"
    ConN.Execute MySql
End Sub

Private Sub SaveBbtKnj()
    Dim KD As String
    KD = "KNJ"
    Dim i As Integer
    With LvKinerja6
        For i = 1 To .ListItems.count
            MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
            "'" & .ListItems(i) & "', " & _
            "'" & .ListItems(i).SubItems(1) & "'," & _
            "'" & KD & "')"
            ConN.Execute MySql
        Next
       ' BtnKinerja.Enabled = False
    End With
End Sub

Private Sub BtnKinerja_Click()
    Kinerja
    TotKinerja
    NrmKinerja
    TotNrmKinerja
    DellBbtKnj
    SaveBbtKnj
End Sub
'========== Masa Kerja =======================
Private Sub KlmMsKerja()
    Dim c As Integer
    For c = 1 To LvMsKerja1.ListItems.count
        LvMsKerja2.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
        LvMsKerja3.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
        LvMsKerja4.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
        LvMsKerja5.ColumnHeaders.Add , , LvAttention1.ListItems(c).Text, , lvwColumnRight
    Next
    
    LvMsKerja4.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
    LvMsKerja5.ColumnHeaders.Add , , "Jumlah", , lvwColumnRight
End Sub

Private Sub LoadNamaMsKerja()
    Dim cList As ListItem

    MySql = "SELECT tb_pegawai.nama, tb_masakerja.AHP FROM tb_masakerja , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_masakerja.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LvMsKerja1.View = lvwReport
    LvMsKerja1.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LvMsKerja1.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
End Sub

Private Sub MsKerja()
    Dim cList As ListItem
    Dim cIndex As Integer
    Dim Karakter As Double
    Dim Kinerja As Double
    Dim MasaKerja As Double
    Dim Pendidikan As Double
    Dim Pengalaman As Double
    Dim Kriteria As String
    Dim Skala As Double
    With LvMsKerja1
        For cIndex = 1 To .ListItems.count
            Karakter = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(1).SubItems(1))
            Kinerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(2).SubItems(1))
            MasaKerja = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(3).SubItems(1))
            Pendidikan = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(4).SubItems(1))
            Pengalaman = Val(.ListItems(cIndex).SubItems(1)) / Val(.ListItems(5).SubItems(1))
            Kriteria = .ListItems(cIndex)
            Skala = .ListItems(cIndex).SubItems(1)
   
            Set cList = LvMsKerja2.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Skala
                cList.SubItems(2) = Format(Karakter, "0.000")
                cList.SubItems(3) = Format(Kinerja, "0.000")
                cList.SubItems(4) = Format(MasaKerja, "0.000")
                cList.SubItems(5) = Format(Pendidikan, "0.000")
                cList.SubItems(6) = Format(Pengalaman, "0.000")
        Next
    End With
End Sub

Private Sub NrmMsKerja()
    Dim cListNm As ListItem
        Dim cIndexNm As Integer
        Dim NmKarakter As Double
        Dim NmKinerja As Double
        Dim NmMasaKerja As Double
        Dim NmPendidikan As Double
        Dim NmPengalaman As Double
        Dim NmKriteria As String
        With LvMsKerja2
            For cIndexNm = 1 To .ListItems.count
                NmKarakter = Val(.ListItems(cIndexNm).SubItems(2)) / TKarakter
                NmKinerja = Val(.ListItems(cIndexNm).SubItems(3)) / TKinerja
                NmMasaKerja = Val(.ListItems(cIndexNm).SubItems(4)) / TMasaKerja
                NmPendidikan = Val(.ListItems(cIndexNm).SubItems(5)) / TPendidikan
                NmPengalaman = Val(.ListItems(cIndexNm).SubItems(6)) / TPengalaman
                NmKriteria = .ListItems(cIndexNm)
        
                Set cListNm = LvMsKerja4.ListItems.Add(, , NmKriteria)
                    cListNm.SubItems(1) = Format(NmKarakter, "0.000")
                    cListNm.SubItems(2) = Format(NmKinerja, "0.000")
                    cListNm.SubItems(3) = Format(NmMasaKerja, "0.000")
                    cListNm.SubItems(4) = Format(NmPendidikan, "0.000")
                    cListNm.SubItems(5) = Format(NmPengalaman, "0.000")
                    cListNm.SubItems(6) = Format(NmKarakter + NmKinerja + NmMasaKerja + NmPendidikan + NmPengalaman, "0.000")
            Next
        End With
End Sub

Private Sub TotMsKerja()
    Dim cList As ListItem
    Dim lngIndex As Integer
        
        TSkala = 0
        TKarakter = 0
        TKinerja = 0
        TMasaKerja = 0
        TPendidikan = 0
        TPengalaman = 0
        With LvMsKerja2
        For lngIndex = 1 To .ListItems.count
            TSkala = TSkala + .ListItems(lngIndex).SubItems(1)
            TKarakter = TKarakter + .ListItems(lngIndex).SubItems(2)
            TKinerja = TKinerja + .ListItems(lngIndex).SubItems(3)
            TMasaKerja = TMasaKerja + .ListItems(lngIndex).SubItems(4)
            TPendidikan = TPendidikan + .ListItems(lngIndex).SubItems(5)
            TPengalaman = TPengalaman + .ListItems(lngIndex).SubItems(6)
        Next
        End With
         Set cList = LvMsKerja3.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = TSkala
                cList.SubItems(2) = Format(TKarakter, "0.000")
                cList.SubItems(3) = Format(TKinerja, "0.000")
                cList.SubItems(4) = Format(TMasaKerja, "0.000")
                cList.SubItems(5) = Format(TPendidikan, "0.000")
                cList.SubItems(6) = Format(TPengalaman, "0.000")
        
End Sub

Private Sub TotNrmMsKerja()
    Dim cList As ListItem
    
    TNmKarakter = 0
    TNmKinerja = 0
    TNmMasaKerja = 0
    TNmPendidikan = 0
    TNmPengalaman = 0
    TNmJumlah = 0

    Dim cIndexTNm As Integer
    With LvMsKerja4
        For cIndexTNm = 1 To .ListItems.count
            TNmKarakter = TNmKarakter + .ListItems(cIndexTNm).SubItems(1)
            TNmKinerja = TNmKinerja + .ListItems(cIndexTNm).SubItems(2)
            TNmMasaKerja = TNmMasaKerja + .ListItems(cIndexTNm).SubItems(3)
            TNmPendidikan = TNmPendidikan + .ListItems(cIndexTNm).SubItems(4)
            TNmPengalaman = TNmPengalaman + .ListItems(cIndexTNm).SubItems(5)
            TNmJumlah = TNmJumlah + .ListItems(cIndexTNm).SubItems(6)
        Next
    End With
         Set cList = LvMsKerja5.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = Format(TNmKarakter, "0.000")
                cList.SubItems(2) = Format(TNmKinerja, "0.000")
                cList.SubItems(3) = Format(TNmMasaKerja, "0.000")
                cList.SubItems(4) = Format(TNmPendidikan, "0.000")
                cList.SubItems(5) = Format(TNmPengalaman, "0.000")
                cList.SubItems(6) = Format(TNmJumlah, "0.000")
                
        MsKerjaPrioritas
End Sub

Private Sub MsKerjaPrioritas()
    Dim cIndexBp As Integer
    Dim cList As ListItem
    Dim Kriteria As String
        For cIndexBp = 1 To LvMsKerja4.ListItems.count
            Kriteria = LvMsKerja4.ListItems(cIndexBp)
            BpMsKerja = LvMsKerja4.ListItems(cIndexBp).SubItems(6) / TNmJumlah
            
            Set cList = LvMsKerja6.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Format(BpMsKerja, "0.000")
        Next
       JmlMsP
End Sub

Private Sub JmlMsP()
    Dim cIndexJMlBP As Integer
    Dim jml As Double
    Dim cList As ListItem
        For cIndexJMlBP = 1 To LvMsKerja4.ListItems.count
            jml = jml + LvMsKerja6.ListItems(cIndexJMlBP).SubItems(1)
        Next
        Set cList = LvMsKerja7.ListItems.Add(, , "Jumlah")
            cList.SubItems(1) = Format(jml, "0.000")
End Sub

Private Sub DellBbtMsk()
    Dim KD As String
    KD = "MSK"
    MySql = "DELETE FROM tb_bbt_ahp WHERE kode = '" & KD & "'"
    ConN.Execute MySql
End Sub

Private Sub SaveBbtMsk()
    Dim KD As String
    KD = "MSK"
    Dim i As Integer
    With LvMsKerja6
        For i = 1 To .ListItems.count
            MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
            "'" & .ListItems(i) & "', " & _
            "'" & .ListItems(i).SubItems(1) & "'," & _
            "'" & KD & "')"
            ConN.Execute MySql
        Next
    End With
End Sub

Private Sub BtnMsKerja_Click()
    MsKerja
    TotMsKerja
    NrmMsKerja
    TotNrmMsKerja
    DellBbtMsk
    SaveBbtMsk
End Sub
'============================ Akumulasi =================================================
Private Sub LoadBobot()
    Dim cList As ListItem
    MySql = "SELECT DISTINCT nama, " & _
    "SUM(CASE WHEN kode = 'PND' THEN nilai ELSE 0 END) AS 'Pendidikan', " & _
    "SUM(CASE WHEN kode = 'PNG' THEN nilai ELSE 0 END) AS 'Pengalaman', " & _
    "SUM(CASE WHEN kode = 'LDR' THEN nilai ELSE 0 END) AS 'Leadership', " & _
    "SUM(CASE WHEN kode = 'LNR' THEN nilai ELSE 0 END) AS 'learning', " & _
    "SUM(CASE WHEN kode = 'ATN' THEN nilai ELSE 0 END) AS 'Attention', " & _
    "SUM(CASE WHEN kode = 'KNJ' THEN nilai ELSE 0 END) AS 'Kinerja', " & _
    "SUM(CASE WHEN kode = 'MSK' THEN nilai ELSE 0 END) AS 'MasaKerja' " & _
    "FROM tb_bbt_ahp GROUP BY nama ORDER BY nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic
    With LvAkumulasi1
        .View = lvwReport
        .ListItems.Clear
            Do Until SdR.EOF
                Set cList = .ListItems.Add(, , SdR.Fields(0))
                    cList.SubItems(1) = Format(SdR.Fields(1), "0.000")
                    cList.SubItems(2) = Format(SdR.Fields(2), "0.000")
                    cList.SubItems(3) = Format(SdR.Fields(3), "0.000")
                    cList.SubItems(4) = Format(SdR.Fields(4), "0.000")
                    cList.SubItems(5) = Format(SdR.Fields(5), "0.000")
                    cList.SubItems(6) = Format(SdR.Fields(6), "0.000")
                    cList.SubItems(7) = Format(SdR.Fields(7), "0.000")
                SdR.MoveNext
            Loop
    End With
End Sub

Private Sub Command1_Click()
    Dim cIndex As Integer
    Dim cList As ListItem
    Dim cList1 As ListItem
    Dim Kriteria As String
    Dim Pendidikan As Double
    Dim Pengalaman As Double
    Dim Kinerja As Double
    Dim Leader As Double
    Dim Learning As Double
    Dim Attention As Double
    Dim MsKerja As Double
    LvAkumulasi2.ListItems.Clear
    LvAkumulasi3.ListItems.Clear
    For cIndex = 1 To LvAkumulasi1.ListItems.count
        Kriteria = LvAkumulasi1.ListItems(cIndex)
        Leader = LvBobot6.ListItems(1).SubItems(1) * LvAkumulasi1.ListItems(cIndex).SubItems(3)
        Learning = LvBobot6.ListItems(1).SubItems(1) * LvAkumulasi1.ListItems(cIndex).SubItems(4)
        Attention = LvBobot6.ListItems(1).SubItems(1) * LvAkumulasi1.ListItems(cIndex).SubItems(5)
        Kinerja = LvBobot6.ListItems(2).SubItems(1) * LvAkumulasi1.ListItems(cIndex).SubItems(6)
        MsKerja = LvBobot6.ListItems(3).SubItems(1) * LvAkumulasi1.ListItems(cIndex).SubItems(7)
        Pendidikan = LvBobot6.ListItems(4).SubItems(1) * LvAkumulasi1.ListItems(cIndex).SubItems(1)
        Pengalaman = LvBobot6.ListItems(5).SubItems(1) * LvAkumulasi1.ListItems(cIndex).SubItems(2)
        
        With LvAkumulasi2
            Set cList = LvAkumulasi2.ListItems.Add(, , Kriteria)
                With cList
                    .SubItems(1) = Format(Pendidikan, "0.000")
                    .SubItems(2) = Format(Pengalaman, "0.000")
                    .SubItems(3) = Format(Leader, "0.000")
                    .SubItems(4) = Format(Learning, "0.000")
                    .SubItems(5) = Format(Attention, "0.000")
                    .SubItems(6) = Format(Kinerja, "0.000")
                    .SubItems(7) = Format(MsKerja, "0.000")
                End With
        End With
        Dim list As ListItem
        Set list = LvAkumulasi3.ListItems.Add(, , Kriteria)
            list.SubItems(1) = Format(Pendidikan + Pengalaman + Leader + Learning + Attention + Kinerja + MsKerja, "0.000")
            Next
End Sub

Private Sub RangkingAHP()
    Dim kode As String
    kode = "AHP"
    Dim i As Integer
    With LvAkumulasi3
    For i = 1 To .ListItems.count
    
        MySql = "INSERT INTO rangking (nama, nilai, kode) VALUES ( " & _
            "'" & .ListItems(i) & "', " & _
            "'" & .ListItems(i).SubItems(1) & "'," & _
            "'" & kode & "')"
            ConN.Execute MySql
    Next
    End With
End Sub

Private Sub Command2_Click()
    LoadBobot
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
          
    If Node.Text = "Matrix Perbandingan Berpasangan" Then
        'PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With PpB
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Normalisasi Berpasangan" Then
        PpB.Visible = False
        'PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With PnB
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Data Mentah AHP" Then
        PpB.Visible = False
        PnB.Visible = False
        'DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With DTahp
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Perbandingan Pendidikan" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        'MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With MpP
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Normalisasi Pendidikan" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        'MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With MnP
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Perbandingan Pengalaman" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        'MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With MpPe
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Normalisasi Pengalaman" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        'MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With MnPe
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Perbandingan Leadership" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        'MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With MpLs
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Normalisasi Leadership" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        'MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With MnLs
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Perbandingan Learning" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        'MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With MpLr
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Normalisasi Learning" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        'MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With MnLr
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Perbandingan Attention" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        'MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With MpA
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Normalisasi Attention" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        'MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With MnA
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Perbandingan Kinerja" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        'MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
        With MpK
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Normalisasi Kinerja" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        'MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
         With MnK
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Perbandingan Masa Kerja" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        'MpMk.Visible = False
        MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False
         With MpMk
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Normalisasi Masa Kerja" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        'MnMk.Visible = False
        RnKAHP.Visible = False
        VScroll1.Visible = False

         With MnMk
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Nilai Perbandingan & Normalisasi Penghitungan Global" Then
        PpB.Visible = False
        PnB.Visible = False
        DTahp.Visible = False
        MpP.Visible = False
        MnP.Visible = False
        MpPe.Visible = False
        MnPe.Visible = False
        MpLs.Visible = False
        MnLs.Visible = False
        MpLr.Visible = False
        MnLr.Visible = False
        MpA.Visible = False
        MnA.Visible = False
        MpK.Visible = False
        MnK.Visible = False
        MpMk.Visible = False
        MnMk.Visible = False
        'RnKAHP.Visible = False
         With RnKAHP
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 9135
            .Width = 13695
        End With
        VScroll1.Visible = True
    End If
End Sub

Private Sub AlignVScroll1()
    
    VScroll1.Height = Frame1.Height
    VScroll1.Top = 0
    VScroll1.Left = Frame1.Height - VScroll1.Width + 12780
    VScroll1.Max = RnKAHP.Height - Frame1.Height
    
   RnKAHP.Top = (-1 * VScroll1)
End Sub

Private Sub Form_Resize()
 AlignVScroll1
End Sub

Private Sub VScroll1_Change()
    AlignVScroll1
End Sub

'***************************** SAW **************************************************************
Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Text = "Data Mentah SAW" Then
        'DTsaw.Visible = False
        BpS1.Visible = False
        BnS1.Visible = False
        RnkSAW.Visible = False
        With DTsaw
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Matrix Perbandingan" Then
        DTsaw.Visible = False
        'BpS1.Visible = False
        BnS1.Visible = False
        RnkSAW.Visible = False
        With BpS1
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Proses Normalisasi" Then
        DTsaw.Visible = False
        BpS1.Visible = False
        'BnS1.Visible = False
        RnkSAW.Visible = False
        With BnS1
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    ElseIf Node.Text = "Hasil Proses Normalisasi" Then
        DTsaw.Visible = False
        BpS1.Visible = False
        BnS1.Visible = False
        'RnkSAW.Visible = False
        With RnkSAW
            .Visible = True
            .Left = 3840
            .Top = 240
            .Height = 4695
            .Width = 13935
        End With
    End If
End Sub

Private Sub BtnProsesBobotSAW_Click()
    JmlBPSAW
    BobotKriteriaSAW
    LoadSAW
    NMax
End Sub


Private Sub LoadKriteriaSAW()
    Dim cList As ListItem

    MySql = "SELECT nama_kriteria, ahp FROM tb_kriteria ORDER BY nama_kriteria ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LvBobotSAW1.View = lvwReport
    LvBobotSAW1.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LvBobotSAW1.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
End Sub

Private Sub JmlBPSAW()
    Dim cJml As Integer
    Dim jml As Double
    
        For cJml = 1 To LvBobotSAW1.ListItems.count
            jml = jml + LvBobotSAW1.ListItems(cJml).SubItems(1)
        Next
        TotJml = jml
End Sub


Private Sub BobotKriteriaSAW()
    Dim cList As ListItem
    Dim cIndex As Integer
    Dim Nyamah As String
    a1 = 0
    b1 = 0
    c1 = 0
    d1 = 0
    e1 = 0
    
    Dim Skala As Double
    With LvBobotSAW1
        For cIndex = 1 To .ListItems.count
            Nyamah = .ListItems(cIndex)
            a1 = Val(.ListItems(cIndex).SubItems(1)) / TotJml
        
            With LvBobotSAW2
            Set cList = .ListItems.Add(, , Nyamah)
                cList.SubItems(1) = Format(a1, "0.000")
            End With
        Next
    End With
End Sub

Private Sub LoadSAW()
    Dim cList As ListItem

    MySql = "SELECT DISTINCT tb_pegawai.nama, tb_pendidikan.SAW AS Pendidikan, " & _
    "tb_pengalaman.SAW AS Pengalaman, tb_karakter.SAW1 AS Leadership, " & _
    "tb_karakter.SAW2 AS Learning, tb_karakter.SAW3 AS Attention, tb_kinerja.SAW AS Kinerja, " & _
    "tb_masakerja.SAW AS Masakerja FROM tb_pegawai, tb_karakter, tb_kinerja, tb_masakerja, " & _
    "tb_pengalaman, tb_pendidikan WHERE tb_pegawai.nik_pegawai = tb_pendidikan.nik_pegawai " & _
    "AND tb_pegawai.nik_pegawai = tb_pengalaman.nik_pegawai AND tb_pegawai.nik_pegawai = " & _
    "tb_karakter.nik_pegawai AND tb_pegawai.nik_pegawai = tb_kinerja.nik_pegawai AND " & _
    "tb_pegawai.nik_pegawai = tb_masakerja.nik_pegawai GROUP BY tb_pegawai.nama ORDER BY  tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    With LvBobotSAW3
        .View = lvwReport
        .ListItems.Clear
        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = Format(SdR.Fields(1), "0.000")
                cList.SubItems(2) = Format(SdR.Fields(2), "0.000")
                cList.SubItems(3) = Format(SdR.Fields(3), "0.000")
                cList.SubItems(4) = Format(SdR.Fields(4), "0.000")
                cList.SubItems(5) = Format(SdR.Fields(5), "0.000")
                cList.SubItems(6) = Format(SdR.Fields(6), "0.000")
                cList.SubItems(7) = Format(SdR.Fields(7), "0.000")
            SdR.MoveNext
        Loop
    End With
End Sub

Private Sub NMax()
Dim i As Integer
Dim cList As ListItem
Dim maxvalueA As Double
Dim maxvalueB As Double
Dim maxvalueC As Double
Dim maxvalueD As Double
Dim maxvalueE As Double
Dim maxvalueF As Double
Dim maxvalueG As Double
Dim maxvalueH As Double
    '================= MAX VALUE =============================================
    With LvBobotSAW3
    '.ListItems.Clear
     maxvalueA = 0
        For i = 1 To .ListItems.count
                If CDbl(.ListItems(i).SubItems(1)) > maxvalueA Then
                    maxvalueA = CDbl(.ListItems(i).SubItems(1))
                End If
                
                If CDbl(.ListItems(i).SubItems(2)) > maxvalueB Then
                    maxvalueB = CDbl(.ListItems(i).SubItems(2))
                End If
                
                If CDbl(.ListItems(i).SubItems(3)) > maxvalueC Then
                    maxvalueC = .ListItems(i).SubItems(3)
                End If
                
                If CDbl(.ListItems(i).SubItems(4)) > maxvalueD Then
                    maxvalueD = .ListItems(i).SubItems(4)
                End If
                
                If CDbl(.ListItems(i).SubItems(5)) > maxvalueE Then
                    maxvalueE = .ListItems(i).SubItems(5)
                End If
                
                If CDbl(.ListItems(i).SubItems(6)) > maxvalueF Then
                    maxvalueF = .ListItems(i).SubItems(6)
                End If
                
                If CDbl(.ListItems(i).SubItems(7)) > maxvalueG Then
                    maxvalueG = .ListItems(i).SubItems(7)
                End If
        Next
    End With
    
        With LvBobotSAW4
            .ListItems.Clear
            Set cList = .ListItems.Add(, , "Max")
                With cList
                .SubItems(1) = Format(maxvalueA, "0.000")
                .SubItems(2) = Format(maxvalueB, "0.000")
                .SubItems(3) = Format(maxvalueC, "0.000")
                .SubItems(4) = Format(maxvalueD, "0.000")
                .SubItems(5) = Format(maxvalueE, "0.000")
                .SubItems(6) = Format(maxvalueF, "0.000")
                .SubItems(7) = Format(maxvalueG, "0.000")
                End With
        End With
            '================================ NORMALISASI ================================================
            Dim ii As Integer
            Dim cList1 As ListItem
            Dim Nama As String
            
            Dim NormalisasiA As Double
            Dim NormalisasiB As Double
            Dim NormalisasiC As Double
            Dim NormalisasiD As Double
            Dim NormalisasiE As Double
            Dim NormalisasiF As Double
            Dim NormalisasiG As Double
            
            LvBobotSAW5.ListItems.Clear
                For ii = 1 To LvBobotSAW3.ListItems.count
                    Nama = LvBobotSAW3.ListItems(ii)
                    NormalisasiA = LvBobotSAW3.ListItems(ii).SubItems(1) / maxvalueA
                    NormalisasiB = LvBobotSAW3.ListItems(ii).SubItems(2) / maxvalueB
                    NormalisasiC = LvBobotSAW3.ListItems(ii).SubItems(3) / maxvalueC
                    NormalisasiD = LvBobotSAW3.ListItems(ii).SubItems(4) / maxvalueD
                    NormalisasiE = LvBobotSAW3.ListItems(ii).SubItems(5) / maxvalueE
                    NormalisasiF = LvBobotSAW3.ListItems(ii).SubItems(6) / maxvalueF
                    NormalisasiG = LvBobotSAW3.ListItems(ii).SubItems(7) / maxvalueG
                
                    With LvBobotSAW5
                        Set cList1 = .ListItems.Add(, , Nama)
                        With cList1
                            .SubItems(1) = Format(NormalisasiA, "0.000")
                            .SubItems(2) = Format(NormalisasiB, "0.000")
                            .SubItems(3) = Format(NormalisasiC, "0.000")
                            .SubItems(4) = Format(NormalisasiD, "0.000")
                            .SubItems(5) = Format(NormalisasiE, "0.000")
                            .SubItems(6) = Format(NormalisasiF, "0.000")
                            .SubItems(7) = Format(NormalisasiG, "0.000")
                        End With
                    End With
                    
                Next
    RankingSAW
    SaveRankingSAW

End Sub

Private Sub RankingSAW()
    '============================ RANKING ==============================
    Dim i As Integer
    Dim cList As ListItem
    Dim cList1 As ListItem
    Dim Jeneng As String
    
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim e As Double
    Dim f As Double
    Dim g As Double
    
    a = 0
    b = 0
    c = 0
    d = 0
    e = 0
    f = 0
    g = 0
    
    LvBobotSAW6.ListItems.Clear
    LvBobotSAW7.ListItems.Clear
    
    For i = 1 To LvBobotSAW5.ListItems.count
        Jeneng = LvBobotSAW5.ListItems(i)
        a = Round(CDec(LvBobotSAW2.ListItems(4).SubItems(1)) * CDec(LvBobotSAW5.ListItems(i).SubItems(1)), 3)
        b = Round(CDec(LvBobotSAW2.ListItems(5).SubItems(1)) * CDec(LvBobotSAW5.ListItems(i).SubItems(2)), 3)
        c = Round(CDec(LvBobotSAW2.ListItems(1).SubItems(1)) * CDec(LvBobotSAW5.ListItems(i).SubItems(3)), 3)
        d = Round(CDec(LvBobotSAW2.ListItems(1).SubItems(1)) * CDec(LvBobotSAW5.ListItems(i).SubItems(4)), 3)
        e = Round(CDec(LvBobotSAW2.ListItems(1).SubItems(1)) * CDec(LvBobotSAW5.ListItems(i).SubItems(5)), 3)
        f = Round(CDec(LvBobotSAW2.ListItems(2).SubItems(1)) * CDec(LvBobotSAW5.ListItems(i).SubItems(6)), 3)
        g = Round(CDec(LvBobotSAW2.ListItems(3).SubItems(1)) * CDec(LvBobotSAW5.ListItems(i).SubItems(7)), 3)
        
        With LvBobotSAW6
            Set cList = .ListItems.Add(, , Jeneng)
            With cList
                .SubItems(1) = Format(a, "0.000")
                .SubItems(2) = Format(b, "0.000")
                .SubItems(3) = Format(c, "0.000")
                .SubItems(4) = Format(d, "0.000")
                .SubItems(5) = Format(e, "0.000")
                .SubItems(6) = Format(f, "0.000")
                .SubItems(7) = Format(g, "0.000")
            End With
        End With
        
        With LvBobotSAW7
            Set cList1 = .ListItems.Add(, , Jeneng)
                cList1.SubItems(1) = Format(a + b + c + d + e + f + g, "0.000")
        End With
    Next
End Sub

Private Sub HapusRanking()
    MySql = "DELETE FROM rangking"
    ConN.Execute MySql
End Sub

Private Sub SaveRankingSAW()
    Dim kode As String
    kode = "SAW"
    Dim i As Integer
    With LvBobotSAW7
        For i = 1 To .ListItems.count
    
            MySql = "INSERT INTO rangking (nama, nilai, kode) VALUES ( " & _
            "'" & .ListItems(i) & "', " & _
            "'" & .ListItems(i).SubItems(1) & "'," & _
            "'" & kode & "')"
            ConN.Execute MySql
        Next
    End With
End Sub

Private Sub LoadKarakterAHP()
    Dim cList As ListItem

    MySql = "SELECT tb_pegawai.nama, tb_karakter.leadership_abilitiy, tb_karakter.learning_abilitiy, " & _
    "tb_karakter.attention_to_detail, tb_karakter.AHP1, tb_karakter.AHP2, tb_karakter.AHP3 FROM tb_pegawai, " & _
    "tb_karakter WHERE tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai ORDER BY tb_pegawai.nama ASC "
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LsKarakterAHP.View = lvwReport
    LsKarakterAHP.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LsKarakterAHP.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
                cList.SubItems(3) = SdR.Fields(3)
                cList.SubItems(4) = SdR.Fields(4)
                cList.SubItems(5) = SdR.Fields(5)
                cList.SubItems(6) = SdR.Fields(6)
            SdR.MoveNext
        Loop
End Sub

Private Sub LoadKarakterSAW()
    Dim cList As ListItem

    MySql = "SELECT tb_pegawai.nama, tb_karakter.leadership_abilitiy, tb_karakter.learning_abilitiy," & _
    "tb_karakter.attention_to_detail, tb_karakter.SAW1, tb_karakter.SAW2, " & _
    "tb_karakter.SAW3 FROM tb_pegawai, tb_karakter WHERE " & _
    "tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai ORDER BY tb_pegawai.nama ASC "
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LsKarakterSAW.View = lvwReport
    LsKarakterSAW.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LsKarakterSAW.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
                cList.SubItems(3) = SdR.Fields(3)
                cList.SubItems(4) = SdR.Fields(4)
                cList.SubItems(5) = SdR.Fields(5)
                cList.SubItems(6) = SdR.Fields(6)
            SdR.MoveNext
        Loop
End Sub


