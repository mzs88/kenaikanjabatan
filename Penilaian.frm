VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Penilaian 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   13590
   Begin VB.Frame frame 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   13335
      Begin MSComctlLib.ListView LsDataSAW 
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   50
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.TreeView ListSAW 
         Height          =   3255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5741
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView LsDataSAW 
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   51
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataSAW 
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   52
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataSAW 
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   53
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataSAW 
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   54
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label LblSAW 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   13335
      Begin MSComctlLib.ListView TmpLDataAHP 
         Height          =   255
         Index           =   3
         Left            =   7320
         TabIndex        =   34
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Tampungan 
         Height          =   255
         Index           =   3
         Left            =   10440
         TabIndex        =   42
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Tampungan 
         Height          =   255
         Index           =   2
         Left            =   10440
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView TmpLDataAHP 
         Height          =   255
         Index           =   2
         Left            =   7320
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView TmpLDataAHP 
         Height          =   255
         Index           =   1
         Left            =   7320
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   7
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.TreeView ListAHP 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5741
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView TmpLDataAHP 
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Tampungan 
         Height          =   255
         Index           =   0
         Left            =   10440
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Tampungan 
         Height          =   255
         Index           =   1
         Left            =   10440
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   14
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   15
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   19
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   7
         Left            =   4200
         TabIndex        =   20
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   8
         Left            =   4200
         TabIndex        =   21
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   22
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   10
         Left            =   4200
         TabIndex        =   23
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   11
         Left            =   4200
         TabIndex        =   24
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   12
         Left            =   4200
         TabIndex        =   25
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   13
         Left            =   4200
         TabIndex        =   26
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   14
         Left            =   5760
         TabIndex        =   27
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   15
         Left            =   5760
         TabIndex        =   28
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   16
         Left            =   5760
         TabIndex        =   29
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   17
         Left            =   5760
         TabIndex        =   30
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   18
         Left            =   5760
         TabIndex        =   31
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   19
         Left            =   5760
         TabIndex        =   32
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LsDataAHP 
         Height          =   255
         Index           =   20
         Left            =   5760
         TabIndex        =   33
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView TmpLDataAHP 
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   35
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView TmpLDataAHP 
         Height          =   255
         Index           =   5
         Left            =   7320
         TabIndex        =   36
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView TmpLDataAHP 
         Height          =   255
         Index           =   6
         Left            =   7320
         TabIndex        =   37
         Top             =   3000
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView TmpLDataAHP 
         Height          =   255
         Index           =   7
         Left            =   8880
         TabIndex        =   38
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView TmpLDataAHP 
         Height          =   255
         Index           =   8
         Left            =   8880
         TabIndex        =   39
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView TmpLDataAHP 
         Height          =   255
         Index           =   9
         Left            =   8880
         TabIndex        =   40
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView TmpLDataAHP 
         Height          =   255
         Index           =   10
         Left            =   8880
         TabIndex        =   41
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Tampungan 
         Height          =   255
         Index           =   4
         Left            =   10440
         TabIndex        =   43
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Tampungan 
         Height          =   255
         Index           =   5
         Left            =   10440
         TabIndex        =   44
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Tampungan 
         Height          =   255
         Index           =   6
         Left            =   10440
         TabIndex        =   45
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Tampungan 
         Height          =   255
         Index           =   7
         Left            =   11760
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Tampungan 
         Height          =   255
         Index           =   8
         Left            =   11760
         TabIndex        =   47
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Tampungan 
         Height          =   255
         Index           =   9
         Left            =   11760
         TabIndex        =   48
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView Tampungan 
         Height          =   255
         Index           =   10
         Left            =   11760
         TabIndex        =   49
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label LblAHP 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penilaian.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penilaian.frx":06FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penilaian.frx":0DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penilaian.frx":14EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penilaian.frx":1BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penilaian.frx":285A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penilaian.frx":2BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penilaian.frx":2F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Penilaian.frx":39A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Friedman Test"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "CheSquere"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Penilaian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nodxAHP As Node
Dim nodrAHP As Node
Dim nodyAHP As Node

Dim nodxSAW As Node
Dim nodrSAW As Node
Dim nodySAW As Node

Dim PmbMnB As Double
Dim PmbMnP As Double
Dim PmbMnPe As Double
Dim PmbMnL As Double
Dim PmbMnLe As Double
Dim PmbMnT As Double
Dim PmbMnK As Double
Dim PmbMnMsK As Double

Private Sub Form_Load()
cNDb
    'On Error Resume Next
    Me.Move 0, 0, MenuUtama.ScaleWidth, _
    MenuUtama.ScaleHeight
    Dim i As Integer, j As Integer
    
    For i = 0 To 20
        With LsDataAHP(i)
        .Visible = False
        .View = lvwReport
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .GridLines = True
        End With
    Next
    
    For j = 0 To 4
        With LsDataSAW(j)
        .Visible = False
        .View = lvwReport
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .GridLines = True
        End With
    Next
LoadListAHP
LoadListSAW

With Toolbar1
    .Buttons(5).Enabled = False
    .Buttons(7).Enabled = False
End With


ListAHP.Enabled = False
ListSAW.Enabled = False
End Sub

Private Sub Form_Resize()
Dim i As Integer, j As Integer
 With Me.frame(0)
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, _
        Me.ScaleHeight / 2 - 500
 End With
 
 With Me.frame(1)
        .Move .Left, (Me.ScaleHeight / 2) + 300, Me.ScaleWidth - .Left * 2, _
        Me.ScaleHeight - .Top - 200
 End With
 
 With ListAHP
    .Move .Left, .Top, 4000, frame(0).Height - 400
 End With

 With ListSAW
    .Move .Left, .Top, 4000, frame(1).Height - 400
 End With

With LblAHP
    .Move ListAHP.Width + 200, .Top, frame(0).Width - ListAHP.Width - 300, 400
End With

With LblSAW
    .Move ListSAW.Width + 200, .Top, frame(1).Width - ListSAW.Width - 300, 400
End With

For i = 0 To 20
LsDataAHP(i).Move ListAHP.Width + 200, LblAHP.Height + 300, frame(0).Width - ListAHP.Width - 300, (frame(0).Height - LblAHP.Height) - 500
Next

For j = 0 To 4
LsDataSAW(j).Move ListAHP.Width + 200, LblSAW.Height + 300, frame(1).Width - ListSAW.Width - 300, (frame(1).Height - LblSAW.Height) - 500
Next

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Dim i As Integer, j As Integer
    Case 1
    BK.Show 1
    Case 2
    Kriteria.Show 1
    Case 4
    Progress.Show
    Bersih
    
    LoadKriteria
    MatrikPerbandinganBerpasangan
    MatrikNormalisasiBerpasangan
    
    LoadKarakterAHP
    
    LoadNamaPendidikan
    MatrikPerbandinganPendidikan
    MatrikNormalisasiPendidikan
    
    LoadNamaPengalaman
    MatrikPerbandinganPengalaman
    MatrikNormalisasiPengalaman
    
    LoadNamaLeadership
    MatrikPerbandinganLeadership
    MatrikNormalisasiLeadership
    
    LoadNamaLearning
    MatrikPerbandinganLearning
    MatrikNormalisasiLearning
    
    LoadNamaAttention
    MatrikPerbandinganAttention
    MatrikNormalisasiAttention
    
    LoadNamaKinerja
    MatrikPerbandinganKinerja
    MatrikNormalisasiKinerja
    
    LoadNamaMsKerja
    MatrikPerbandinganMsKerja
    MatrikNormalisasiMsKerja
    
    NilaiPerbandingan
    NilaiNormalisasiGlobal
    
    LoadKarakterSAW
    MatrikPerbandinganSAW
    MatrikNormalisasiSAW
    MatrikNormalisasiGlobalSAW
    
    ListAHP.Enabled = True
    ListSAW.Enabled = True
    With Toolbar1
        .Buttons(4).Enabled = False
        .Buttons(5).Enabled = True
        .Buttons(7).Enabled = True
    End With
    Case 5
    Bersih
    Progress.Show
    For i = 0 To 20
        With LsDataAHP(i)
            .ListItems.Clear
            .ColumnHeaders.Clear
            .Visible = False
        End With
    Next

    For j = 0 To 4
        With LsDataSAW(j)
            .ListItems.Clear
            .ColumnHeaders.Clear
            .Visible = False
        End With
    Next
    With Toolbar1
        .Buttons(4).Enabled = True
        .Buttons(5).Enabled = False
        .Buttons(7).Enabled = False
    End With
    LblAHP.Caption = ""
    LblSAW.Caption = ""
    ListAHP.Enabled = False
    ListSAW.Enabled = False
    Case 7
    ReportRanking.Show
    Case 9
    FredmanTest.Show
    Case 10
    ChiSquare.Show
End Select
End Sub


Private Sub LoadListAHP()
    With ListAHP
    .LineStyle = tvwTreeLines
    .ImageList = ImageList1
    Set nodxAHP = ListAHP.Nodes.Add(, , "a", "Matrik Perbandingan Berpasangan", 6, 7)
    Set nodrAHP = ListAHP.Nodes.Add(, , "b", "Matrik Normalisasi Berpasangan", 6, 7)
    Set nodxAHP = ListAHP.Nodes.Add(, , "c", "Data Mentah AHP", 6, 7)
    Set nodxAHP = ListAHP.Nodes.Add(, , "d", "Matrik Perbandingan Pendidikan", 6, 7)
    Set nodrAHP = ListAHP.Nodes.Add(, , "e", "Matrik Normalisasi Pendidikan", 6, 7)
    Set nodxAHP = ListAHP.Nodes.Add(, , "f", "Matrik Perbandingan Pengalaman", 6, 7)
    Set nodrAHP = ListAHP.Nodes.Add(, , "g", "Matrik Normalisasi Pengalaman", 6, 7)
    Set nodxAHP = ListAHP.Nodes.Add(, , , "Matrik Perbandingan Karakter", 7)
        nodxAHP.Expanded = True
            Set nodrAHP = ListAHP.Nodes.Add(nodxAHP, tvwChild, "h", "Matrik Perbandingan Leadership", 6, 7)
            Set nodyAHP = ListAHP.Nodes.Add(nodxAHP, tvwChild, "i", "Matrik Normalisasi Leadership", 6, 7)
            Set nodrAHP = ListAHP.Nodes.Add(nodxAHP, tvwChild, "j", "Matrik Perbandingan Learning", 6, 7)
            Set nodyAHP = ListAHP.Nodes.Add(nodxAHP, tvwChild, "k", "Matrik Normalisasi Learning", 6, 7)
            Set nodrAHP = ListAHP.Nodes.Add(nodxAHP, tvwChild, "l", "Matrik Perbandingan Attention", 6, 7)
            Set nodyAHP = ListAHP.Nodes.Add(nodxAHP, tvwChild, "m", "Matrik Normalisasi Attention", 6, 7)
    Set nodxAHP = ListAHP.Nodes.Add(, , "n", "Matrik Perbandingan Kinerja", 6, 7)
    Set nodrAHP = ListAHP.Nodes.Add(, , "o", "Matrik Normalisasi Kinerja", 6, 7)
    Set nodxAHP = ListAHP.Nodes.Add(, , "p", "Matrik Perbandingan Masa Kerja", 6, 7)
    Set nodrAHP = ListAHP.Nodes.Add(, , "q", "Matrik Normalisasi Masa Kerja", 6, 7)
    Set nodxAHP = ListAHP.Nodes.Add(, , "r", "Nilai Perbandingan", 6, 7)
    Set nodrAHP = ListAHP.Nodes.Add(, , "s", "Normalisasi Penghitungan Global", 6, 7)
    End With
End Sub

Private Sub ListAHP_NodeClick(ByVal Node As MSComctlLib.Node)
Select Case Node.Key
Case "a"
LblAHP.Caption = "Matrik Perbandingan Berpasangan"
Dim a As Integer
For a = 0 To 20
    If LsDataAHP(a).Visible = True Then
        LsDataAHP(a).Visible = False
    Else
        LsDataAHP(0).Visible = True
    End If
Next

Case "b"
LblAHP.Caption = "Matrik Normalisasi Berpasangan"
Dim b As Integer
For b = 0 To 20
    If LsDataAHP(b).Visible = True Then
        LsDataAHP(b).Visible = False
    Else
        LsDataAHP(1).Visible = True
    End If
Next

Case "c"
LblAHP.Caption = "Data Mentah AHP"
Dim c As Integer
For c = 0 To 20
    If LsDataAHP(c).Visible = True Then
        LsDataAHP(c).Visible = False
    Else
        LsDataAHP(2).Visible = True
    End If
Next

Case "d"
LblAHP.Caption = "Matrik Perbandingan Pendidikan"
Dim d As Integer
For d = 0 To 20
    If LsDataAHP(d).Visible = True Then
        LsDataAHP(d).Visible = False
    Else
        LsDataAHP(3).Visible = True
    End If
Next

Case "e"
LblAHP.Caption = "Matrik Normalisasi Pendidikan"
Dim e As Integer
For e = 0 To 20
    If LsDataAHP(e).Visible = True Then
        LsDataAHP(e).Visible = False
    Else
        LsDataAHP(4).Visible = True
    End If
Next

Case "f"
LblAHP.Caption = "Matrik Perbandingan Pengalaman"
Dim f As Integer
For f = 0 To 20
    If LsDataAHP(f).Visible = True Then
        LsDataAHP(f).Visible = False
    Else
        LsDataAHP(5).Visible = True
    End If
Next

Case "g"
LblAHP.Caption = "Matrik Normalisasi Pengalaman"
Dim g As Integer
For g = 0 To 20
    If LsDataAHP(g).Visible = True Then
        LsDataAHP(g).Visible = False
    Else
        LsDataAHP(6).Visible = True
    End If
Next

Case "h"
LblAHP.Caption = "Matrik Perbandingan Leadership"
Dim h As Integer
For h = 0 To 20
    If LsDataAHP(h).Visible = True Then
        LsDataAHP(h).Visible = False
    Else
        LsDataAHP(7).Visible = True
    End If
Next

Case "i"
LblAHP.Caption = "Matrik Normalisasi Leadership"
Dim i As Integer
For i = 0 To 20
    If LsDataAHP(i).Visible = True Then
        LsDataAHP(i).Visible = False
    Else
        LsDataAHP(8).Visible = True
    End If
Next

Case "j"
LblAHP.Caption = "Matrik Perbandingan Learning"
Dim j As Integer
For j = 0 To 20
    If LsDataAHP(j).Visible = True Then
        LsDataAHP(j).Visible = False
    Else
        LsDataAHP(9).Visible = True
    End If
Next

Case "k"
LblAHP.Caption = "Matrik Normalisasi Learning"
Dim k As Integer
For k = 0 To 20
    If LsDataAHP(k).Visible = True Then
        LsDataAHP(k).Visible = False
    Else
        LsDataAHP(10).Visible = True
    End If
Next

Case "l"
LblAHP.Caption = "Matrik Perbandingan Attention"
Dim l As Integer
For l = 0 To 20
    If LsDataAHP(l).Visible = True Then
        LsDataAHP(l).Visible = False
    Else
        LsDataAHP(11).Visible = True
    End If
Next

Case "m"
LblAHP.Caption = "Matrik Normalisasi Attention"
Dim m As Integer
For m = 0 To 20
    If LsDataAHP(m).Visible = True Then
        LsDataAHP(m).Visible = False
    Else
        LsDataAHP(12).Visible = True
    End If
Next

Case "n"
LblAHP.Caption = "Matrik Perbandingan Kinerja"
Dim n As Integer
For n = 0 To 20
    If LsDataAHP(n).Visible = True Then
        LsDataAHP(n).Visible = False
    Else
        LsDataAHP(13).Visible = True
    End If
Next

Case "o"
LblAHP.Caption = "Matrik Normalisasi Kinerja"
Dim o As Integer
For o = 0 To 20
    If LsDataAHP(o).Visible = True Then
        LsDataAHP(o).Visible = False
    Else
        LsDataAHP(14).Visible = True
    End If
Next

Case "p"
LblAHP.Caption = "Matrik Perbandingan Masa Kerja"
Dim p As Integer
For p = 0 To 20
    If LsDataAHP(p).Visible = True Then
        LsDataAHP(p).Visible = False
    Else
        LsDataAHP(15).Visible = True
    End If
Next

Case "q"
LblAHP.Caption = "Matrik Normalisasi Masa Kerja"
Dim q As Integer
For q = 0 To 20
    If LsDataAHP(q).Visible = True Then
        LsDataAHP(q).Visible = False
    Else
        LsDataAHP(16).Visible = True
    End If
Next

Case "r"
LblAHP.Caption = "Nilai Perbandingan"
Dim r As Integer
For r = 0 To 20
    If LsDataAHP(r).Visible = True Then
        LsDataAHP(r).Visible = False
    Else
        LsDataAHP(17).Visible = True
    End If
Next

Case "s"
LblAHP.Caption = "Normalisasi Penghitungan Global"
Dim s As Integer
For s = 0 To 20
    If LsDataAHP(s).Visible = True Then
        LsDataAHP(s).Visible = False
    Else
        LsDataAHP(18).Visible = True
    End If
Next


End Select
End Sub

Private Sub LoadListSAW()
    ListSAW.LineStyle = tvwTreeLines
    ListSAW.ImageList = ImageList1
    Set nodxSAW = ListSAW.Nodes.Add(, , "a", "Data Mentah SAW", 6, 7)
    Set nodxSAW = ListSAW.Nodes.Add(, , "b", "Matrik Perbandingan", 6, 7)
    Set nodrSAW = ListSAW.Nodes.Add(, , "c", "Proses Normalisasi", 6, 7)
    Set nodxSAW = ListSAW.Nodes.Add(, , "d", "Hasil Proses Normalisasi", 6, 7)
End Sub



Private Sub ListSAW_NodeClick(ByVal Node As MSComctlLib.Node)
    Select Case Node.Key
    Case "a"
    LblSAW.Caption = "Data Mentah SAW"
    Dim a As Integer
    For a = 0 To 4
    If LsDataSAW(a).Visible = True Then
        LsDataSAW(a).Visible = False
    Else
        LsDataSAW(0).Visible = True
    End If
    Next

    Case "b"
    LblSAW.Caption = "Matrik Perbandingan"
    Dim b As Integer
    For b = 0 To 4
    If LsDataSAW(b).Visible = True Then
        LsDataSAW(b).Visible = False
    Else
        LsDataSAW(1).Visible = True
    End If
    Next

    Case "c"
    LblSAW.Caption = "Proses Normalisasi"
    Dim c As Integer
    For c = 0 To 4
    If LsDataSAW(c).Visible = True Then
        LsDataSAW(c).Visible = False
    Else
        LsDataSAW(2).Visible = True
    End If
    Next

    Case "d"
    LblSAW.Caption = "Hasil Proses Normalisasi"
    Dim d As Integer
    For d = 0 To 4
    If LsDataSAW(d).Visible = True Then
        LsDataSAW(d).Visible = False
    Else
        LsDataSAW(3).Visible = True
    End If
    Next

    End Select
End Sub

Private Sub LoadKriteria()
    Dim cList As ListItem
    With TmpLDataAHP(0)
    .View = lvwReport
    .FullRowSelect = True
    
    .ColumnHeaders.Add , , ""
    .ColumnHeaders.Add , , ""
    
    MySql = "SELECT nama_kriteria, ahp FROM tb_kriteria ORDER BY nama_kriteria ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    .View = lvwReport
    .ListItems.Clear
        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
    End With
End Sub

'MatrikPerbandinganBerpasangan
Private Sub MatrikPerbandinganBerpasangan()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(0).ListItems.count + 1
With LsDataAHP(0)
    .View = lvwReport
    .FullRowSelect = True
    .GridLines = True
Tampungan(0).View = lvwReport

.ColumnHeaders.Add , , "Kriteria", 2000
.ColumnHeaders.Add , , "AHP", (.Width - 2100) / a, lvwColumnCenter
Tampungan(0).ColumnHeaders.Add , , "jumlah"
For i = 1 To TmpLDataAHP(0).ListItems.count
    .ColumnHeaders.Add , , TmpLDataAHP(0).ListItems(i).Text, (.Width - 2100) / a, lvwColumnRight
    Tampungan(0).ColumnHeaders.Add , , TmpLDataAHP(0).ListItems(i).Text
Next
End With
HMpB
JMpB
End Sub

Private Sub HMpB()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(0).ListItems.Clear
For Each X In TmpLDataAHP(0).ListItems
    Set cList = LsDataAHP(0).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(0).ListItems.count
        b = TmpLDataAHP(0).ListItems(i).SubItems(1)
        c = a / b
        d = d + c
        cList.SubItems(i + 1) = Format(c, "0.000")
    Next
    c = 0
    d = 0
Next
End Sub

Private Sub JMpB()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double

a = 0
Set cList = LsDataAHP(0).ListItems.Add(, , "")
Set cList = LsDataAHP(0).ListItems.Add(, , "Jumlah")
Set cList1 = Tampungan(0).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To LsDataAHP(0).ColumnHeaders.count - 1
    For Each X In LsDataAHP(0).ListItems
    a = a + X.SubItems(i)
    b = b + X.SubItems(i + 1)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    cList1.SubItems(i) = Format(b, "0.000")
    a = 0
    b = 0
Next
End Sub

'MatrikNormalisasiBerpasangan
Private Sub MatrikNormalisasiBerpasangan()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(0).ListItems.count + 2
With LsDataAHP(1)
    .View = lvwReport
    .FullRowSelect = True
    .GridLines = True
    .ColumnHeaders.Add , , "Kriteria", 2000

For i = 1 To TmpLDataAHP(0).ListItems.count
    .ColumnHeaders.Add , , TmpLDataAHP(0).ListItems(i).Text, (.Width - 2000) / a, lvwColumnRight
Next
.ColumnHeaders.Add , , "Jumlah", (.Width - 2000) / a, lvwColumnRight
.ColumnHeaders.Add , , "Bobot", (.Width - 2000) / a, lvwColumnRight
End With
HMnB
PMnB
HMnB1
JMnB
End Sub

Private Sub HMnB()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(1).ListItems.Clear
For Each X In TmpLDataAHP(0).ListItems
    Set cList = LsDataAHP(1).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(0).ListItems.count
        b = TmpLDataAHP(0).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(0).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
    Next
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub PMnB()
Dim i As Integer
Dim a As Double
Dim JklM As Integer
JklM = LsDataAHP(1).ColumnHeaders.count - 2
For i = 1 To LsDataAHP(1).ListItems.count
    a = a + LsDataAHP(1).ListItems(i).SubItems(JklM)
Next
PmbMnB = Format(a, "0.000")
End Sub

Private Sub HMnB1()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(1).ListItems.Clear
For Each X In TmpLDataAHP(0).ListItems
    Set cList = LsDataAHP(1).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(0).ListItems.count
        b = TmpLDataAHP(0).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(0).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
        g = f / PmbMnB
        cList.SubItems(i + 2) = Format(g, "0.000")
    Next
    MySql = "INSERT INTO coba (nama, nilai, kode) VALUES ( " & _
    "'" & X & "', " & _
    "'" & g & "'," & _
    "'" & "BnB" & "')"
    ConN.Execute MySql
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub JMnB()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double

a = 0
Set cList = LsDataAHP(1).ListItems.Add(, , "")
Set cList = LsDataAHP(1).ListItems.Add(, , "Jumlah")
Set cList1 = Tampungan(0).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To LsDataAHP(1).ColumnHeaders.count - 1
    For Each X In LsDataAHP(1).ListItems
    a = a + X.SubItems(i)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    a = 0
Next
End Sub

Private Sub LoadKarakterAHP()
    Dim cList As ListItem
    With LsDataAHP(2)
    .ColumnHeaders.Add , , "Nama", 4000
    .ColumnHeaders.Add , , "Leadership", (.Width - 4300) / 6
    .ColumnHeaders.Add , , "Nilai", (.Width - 4300) / 6, lvwColumnRight
    .ColumnHeaders.Add , , "Learning", (.Width - 4300) / 6
    .ColumnHeaders.Add , , "Nilai", (.Width - 4300) / 6, lvwColumnRight
    .ColumnHeaders.Add , , "Attention", (.Width - 4300) / 6
    .ColumnHeaders.Add , , "Nilai", (.Width - 4300) / 6, lvwColumnRight
    
    MySql = "SELECT tb_pegawai.nama, tb_karakter.leadership_abilitiy,tb_karakter.AHP1, " & _
    "tb_karakter.learning_abilitiy, tb_karakter.AHP2, tb_karakter.attention_to_detail, " & _
    "tb_karakter.AHP3 FROM tb_pegawai, tb_karakter WHERE " & _
    "tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    .View = lvwReport
    .ListItems.Clear
    .FullRowSelect = True
    
        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
                cList.SubItems(3) = SdR.Fields(3)
                cList.SubItems(4) = SdR.Fields(4)
                cList.SubItems(5) = SdR.Fields(5)
                cList.SubItems(6) = SdR.Fields(6)
            SdR.MoveNext
        Loop
    End With
End Sub

Private Sub LoadNamaPendidikan()
    Dim cList As ListItem
    With TmpLDataAHP(1)
    .View = lvwReport
    .FullRowSelect = True
    
    .ColumnHeaders.Add , , ""
    .ColumnHeaders.Add , , ""
    
    MySql = "SELECT tb_pegawai.nama, tb_pendidikan.AHP FROM tb_pendidikan, tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_pendidikan.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    .View = lvwReport
    .ListItems.Clear
        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
    End With
End Sub

'MatrikPerbandinganPendidikan
Private Sub MatrikPerbandinganPendidikan()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(1).ListItems.count + 1
With LsDataAHP(3)
    .View = lvwReport
    .FullRowSelect = True
    .GridLines = True
Tampungan(1).View = lvwReport

.ColumnHeaders.Add , , "Kriteria", 2000
.ColumnHeaders.Add , , "AHP", (.Width - 2300) / a, lvwColumnCenter
Tampungan(1).ColumnHeaders.Add , , "jumlah"
For i = 1 To TmpLDataAHP(1).ListItems.count
    .ColumnHeaders.Add , , TmpLDataAHP(1).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
    Tampungan(1).ColumnHeaders.Add , , TmpLDataAHP(1).ListItems(i).Text
Next
End With
HMpP
JMpP
End Sub

Private Sub HMpP()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double
With TmpLDataAHP(1)
LsDataAHP(3).ListItems.Clear
For Each X In .ListItems
    Set cList = LsDataAHP(3).ListItems.Add(, , X)
        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
        a = X.ListSubItems(1)
    For i = 1 To .ListItems.count
        b = .ListItems(i).SubItems(1)
        c = a / b
        cList.SubItems(i + 1) = Format(c, "0.000")
    Next
    c = 0
Next
End With
End Sub

Private Sub JMpP()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double
With LsDataAHP(3)
a = 0
Set cList = .ListItems.Add(, , "")
Set cList = .ListItems.Add(, , "Jumlah")
Set cList1 = Tampungan(1).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To .ColumnHeaders.count - 1
    For Each X In LsDataAHP(3).ListItems
    a = a + X.SubItems(i)
    b = b + X.SubItems(i + 1)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    cList1.SubItems(i) = Format(b, "0.000")
    a = 0
    b = 0
Next
End With
End Sub

'MatrikNormalisasiPendidikan
Private Sub MatrikNormalisasiPendidikan()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(1).ListItems.count + 2
With LsDataAHP(4)
    .View = lvwReport
    .FullRowSelect = True
    .GridLines = True
    .ColumnHeaders.Add , , "Kriteria", 2000

For i = 1 To TmpLDataAHP(1).ListItems.count
    .ColumnHeaders.Add , , TmpLDataAHP(1).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
Next
.ColumnHeaders.Add , , "Jumlah", (.Width - 2300) / a, lvwColumnRight
.ColumnHeaders.Add , , "Bobot", (.Width - 2300) / a, lvwColumnRight
End With
HMnP
PMnP
HMnP1
JMnP
End Sub

Private Sub HMnP()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(4).ListItems.Clear
For Each X In TmpLDataAHP(1).ListItems
    Set cList = LsDataAHP(4).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(1).ListItems.count
        b = TmpLDataAHP(1).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(1).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
    Next
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub PMnP()
Dim i As Integer
Dim a As Double
Dim JklM As Integer
JklM = LsDataAHP(4).ColumnHeaders.count - 2
For i = 1 To LsDataAHP(4).ListItems.count
    a = a + LsDataAHP(4).ListItems(i).SubItems(JklM)
Next
PmbMnP = Format(a, "0.000")
End Sub

Private Sub HMnP1()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(4).ListItems.Clear
For Each X In TmpLDataAHP(1).ListItems
    Set cList = LsDataAHP(4).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(1).ListItems.count
        b = TmpLDataAHP(1).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(1).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
        g = f / PmbMnP
        cList.SubItems(i + 2) = Format(g, "0.000")
    Next
    MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
    "'" & X & "', " & _
    "'" & g & "'," & _
    "'" & "PND" & "')"
    ConN.Execute MySql

    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub JMnP()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double

a = 0
Set cList = LsDataAHP(4).ListItems.Add(, , "")
Set cList = LsDataAHP(4).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To LsDataAHP(4).ColumnHeaders.count - 1
    For Each X In LsDataAHP(4).ListItems
    a = a + X.SubItems(i)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    a = 0
Next
End Sub

Private Sub LoadNamaPengalaman()
    Dim cList As ListItem
    With TmpLDataAHP(2)
    .View = lvwReport
    .FullRowSelect = True
    
    .ColumnHeaders.Add , , ""
    .ColumnHeaders.Add , , ""
    
    MySql = "SELECT tb_pegawai.nama, tb_pengalaman.AHP FROM tb_pengalaman , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_pengalaman.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    .View = lvwReport
    .ListItems.Clear
        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
    End With
End Sub


'MatrikPerbandinganPengalaman
Private Sub MatrikPerbandinganPengalaman()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(2).ListItems.count + 1
    With LsDataAHP(5)
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        Tampungan(2).View = lvwReport
        .ColumnHeaders.Add , , "Kriteria", 2000
        .ColumnHeaders.Add , , "AHP", (.Width - 2300) / a, lvwColumnCenter
        Tampungan(2).ColumnHeaders.Add , , "jumlah"
        For i = 1 To TmpLDataAHP(2).ListItems.count
            .ColumnHeaders.Add , , TmpLDataAHP(2).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
            Tampungan(2).ColumnHeaders.Add , , TmpLDataAHP(2).ListItems(i).Text
        Next
    End With
HMpPe
JMpPe
End Sub

Private Sub HMpPe()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double
With TmpLDataAHP(2)
LsDataAHP(5).ListItems.Clear
For Each X In .ListItems
    Set cList = LsDataAHP(5).ListItems.Add(, , X)
        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
        a = X.ListSubItems(1)
    For i = 1 To .ListItems.count
        b = .ListItems(i).SubItems(1)
        c = a / b
        cList.SubItems(i + 1) = Format(c, "0.000")
    Next
    c = 0
Next
End With
End Sub

Private Sub JMpPe()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double
With LsDataAHP(5)
a = 0
Set cList = .ListItems.Add(, , "")
Set cList = .ListItems.Add(, , "Jumlah")
Set cList1 = Tampungan(2).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To .ColumnHeaders.count - 1
    For Each X In LsDataAHP(5).ListItems
    a = a + X.SubItems(i)
    b = b + X.SubItems(i + 1)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    cList1.SubItems(i) = Format(b, "0.000")
    a = 0
    b = 0
Next
End With
End Sub

'MatrikNormalisasiPengalaman
Private Sub MatrikNormalisasiPengalaman()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(2).ListItems.count + 2
With LsDataAHP(6)
    .View = lvwReport
    .FullRowSelect = True
    .GridLines = True
    .ColumnHeaders.Add , , "Kriteria", 2000

For i = 1 To TmpLDataAHP(2).ListItems.count
    .ColumnHeaders.Add , , TmpLDataAHP(2).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
Next
.ColumnHeaders.Add , , "Jumlah", (.Width - 2300) / a, lvwColumnRight
.ColumnHeaders.Add , , "Bobot", (.Width - 2300) / a, lvwColumnRight
End With
HMnPe
PMnPe
HMnPe1
JMnPe
End Sub

Private Sub HMnPe()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(6).ListItems.Clear
For Each X In TmpLDataAHP(2).ListItems
    Set cList = LsDataAHP(6).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(2).ListItems.count
        b = TmpLDataAHP(2).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(2).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
    Next
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub PMnPe()
Dim i As Integer
Dim a As Double
Dim JklM As Integer
JklM = LsDataAHP(6).ColumnHeaders.count - 2
For i = 1 To LsDataAHP(6).ListItems.count
    a = a + LsDataAHP(6).ListItems(i).SubItems(JklM)
Next
PmbMnPe = Format(a, "0.000")
End Sub

Private Sub HMnPe1()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(6).ListItems.Clear
For Each X In TmpLDataAHP(2).ListItems
    Set cList = LsDataAHP(6).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(2).ListItems.count
        b = TmpLDataAHP(2).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(2).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
        g = f / PmbMnPe
        cList.SubItems(i + 2) = Format(g, "0.000")
    Next
    MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
    "'" & X & "', " & _
    "'" & g & "'," & _
    "'" & "PNG" & "')"
    ConN.Execute MySql
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub JMnPe()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double

a = 0
Set cList = LsDataAHP(6).ListItems.Add(, , "")
Set cList = LsDataAHP(6).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To LsDataAHP(6).ColumnHeaders.count - 1
    For Each X In LsDataAHP(6).ListItems
    a = a + X.SubItems(i)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    a = 0
Next
End Sub

Private Sub LoadNamaLeadership()
    Dim cList As ListItem
    With TmpLDataAHP(3)
    .View = lvwReport
    .FullRowSelect = True
    
    .ColumnHeaders.Add , , ""
    .ColumnHeaders.Add , , ""
    
    MySql = "SELECT tb_pegawai.nama, tb_karakter.AHP1 FROM tb_karakter , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    .View = lvwReport
    .ListItems.Clear
        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
    End With
End Sub


'MatrikPerbandinganLeadership
Private Sub MatrikPerbandinganLeadership()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(3).ListItems.count + 1
    With LsDataAHP(7)
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        Tampungan(3).View = lvwReport
        .ColumnHeaders.Add , , "Kriteria", 2000
        .ColumnHeaders.Add , , "AHP", (.Width - 2300) / a, lvwColumnCenter
        Tampungan(3).ColumnHeaders.Add , , "jumlah"
        For i = 1 To TmpLDataAHP(3).ListItems.count
            .ColumnHeaders.Add , , TmpLDataAHP(3).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
            Tampungan(3).ColumnHeaders.Add , , TmpLDataAHP(3).ListItems(i).Text
        Next
    End With
HMpL
JMpL
End Sub

Private Sub HMpL()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double
With TmpLDataAHP(3)
LsDataAHP(7).ListItems.Clear
For Each X In .ListItems
    Set cList = LsDataAHP(7).ListItems.Add(, , X)
        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
        a = X.ListSubItems(1)
    For i = 1 To .ListItems.count
        b = .ListItems(i).SubItems(1)
        c = a / b
        cList.SubItems(i + 1) = Format(c, "0.000")
    Next
    c = 0
Next
End With
End Sub

Private Sub JMpL()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double
With LsDataAHP(7)
a = 0
Set cList = .ListItems.Add(, , "")
Set cList = .ListItems.Add(, , "Jumlah")
Set cList1 = Tampungan(3).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To .ColumnHeaders.count - 1
    For Each X In LsDataAHP(7).ListItems
    a = a + X.SubItems(i)
    b = b + X.SubItems(i + 1)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    cList1.SubItems(i) = Format(b, "0.000")
    a = 0
    b = 0
Next
End With
End Sub

'MatrikNormalisasiLeadership
Private Sub MatrikNormalisasiLeadership()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(3).ListItems.count + 2
With LsDataAHP(8)
    .View = lvwReport
    .FullRowSelect = True
    .GridLines = True
    .ColumnHeaders.Add , , "Kriteria", 2000

For i = 1 To TmpLDataAHP(3).ListItems.count
    .ColumnHeaders.Add , , TmpLDataAHP(3).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
Next
.ColumnHeaders.Add , , "Jumlah", (.Width - 2300) / a, lvwColumnRight
.ColumnHeaders.Add , , "Bobot", (.Width - 2300) / a, lvwColumnRight
End With
HMnL
PMnL
HMnL1
JMnL
End Sub

Private Sub HMnL()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(8).ListItems.Clear
For Each X In TmpLDataAHP(3).ListItems
    Set cList = LsDataAHP(8).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(3).ListItems.count
        b = TmpLDataAHP(3).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(3).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
    Next
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub PMnL()
Dim i As Integer
Dim a As Double
Dim JklM As Integer
JklM = LsDataAHP(8).ColumnHeaders.count - 2
For i = 1 To LsDataAHP(8).ListItems.count
    a = a + LsDataAHP(8).ListItems(i).SubItems(JklM)
Next
PmbMnL = Format(a, "0.000")
End Sub

Private Sub HMnL1()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(8).ListItems.Clear
For Each X In TmpLDataAHP(3).ListItems
    Set cList = LsDataAHP(8).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(3).ListItems.count
        b = TmpLDataAHP(3).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(3).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
        g = f / PmbMnL
        cList.SubItems(i + 2) = Format(g, "0.000")
    Next
    MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
    "'" & X & "', " & _
    "'" & g & "'," & _
    "'" & "LDR" & "')"
    ConN.Execute MySql
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub JMnL()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double

a = 0
Set cList = LsDataAHP(8).ListItems.Add(, , "")
Set cList = LsDataAHP(8).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To LsDataAHP(8).ColumnHeaders.count - 1
    For Each X In LsDataAHP(8).ListItems
    a = a + X.SubItems(i)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    a = 0
Next
End Sub

Private Sub LoadNamaLearning()
    Dim cList As ListItem
    With TmpLDataAHP(4)
    .View = lvwReport
    .FullRowSelect = True
    
    .ColumnHeaders.Add , , ""
    .ColumnHeaders.Add , , ""
    
    MySql = "SELECT tb_pegawai.nama, tb_karakter.AHP2 FROM tb_karakter , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    .View = lvwReport
    .ListItems.Clear
        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
    End With
End Sub


'MatrikPerbandinganLearning
Private Sub MatrikPerbandinganLearning()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(4).ListItems.count + 1
    With LsDataAHP(9)
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        Tampungan(4).View = lvwReport
        .ColumnHeaders.Add , , "Kriteria", 2000
        .ColumnHeaders.Add , , "AHP", (.Width - 2300) / a, lvwColumnCenter
        Tampungan(4).ColumnHeaders.Add , , "jumlah"
        For i = 1 To TmpLDataAHP(4).ListItems.count
            .ColumnHeaders.Add , , TmpLDataAHP(4).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
            Tampungan(4).ColumnHeaders.Add , , TmpLDataAHP(4).ListItems(i).Text
        Next
    End With
HMpLe
JMpLe
End Sub

Private Sub HMpLe()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double
With TmpLDataAHP(4)
LsDataAHP(9).ListItems.Clear
For Each X In .ListItems
    Set cList = LsDataAHP(9).ListItems.Add(, , X)
        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
        a = X.ListSubItems(1)
    For i = 1 To .ListItems.count
        b = .ListItems(i).SubItems(1)
        c = a / b
        cList.SubItems(i + 1) = Format(c, "0.000")
    Next
    c = 0
Next
End With
End Sub

Private Sub JMpLe()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double
With LsDataAHP(9)
a = 0
Set cList = .ListItems.Add(, , "")
Set cList = .ListItems.Add(, , "Jumlah")
Set cList1 = Tampungan(4).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To .ColumnHeaders.count - 1
    For Each X In LsDataAHP(9).ListItems
    a = a + X.SubItems(i)
    b = b + X.SubItems(i + 1)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    cList1.SubItems(i) = Format(b, "0.000")
    a = 0
    b = 0
Next
End With
End Sub

'MatrikNormalisasiLearning
Private Sub MatrikNormalisasiLearning()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(4).ListItems.count + 2
With LsDataAHP(10)
    .View = lvwReport
    .FullRowSelect = True
    .GridLines = True
    .ColumnHeaders.Add , , "Kriteria", 2000

For i = 1 To TmpLDataAHP(4).ListItems.count
    .ColumnHeaders.Add , , TmpLDataAHP(4).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
Next
.ColumnHeaders.Add , , "Jumlah", (.Width - 2300) / a, lvwColumnRight
.ColumnHeaders.Add , , "Bobot", (.Width - 2300) / a, lvwColumnRight
End With
HMnLe
PMnLe
HMnLe1
JMnLe
End Sub

Private Sub HMnLe()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(10).ListItems.Clear
For Each X In TmpLDataAHP(4).ListItems
    Set cList = LsDataAHP(10).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(4).ListItems.count
        b = TmpLDataAHP(4).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(4).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
    Next
    
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub PMnLe()
Dim i As Integer
Dim a As Double
Dim JklM As Integer
JklM = LsDataAHP(10).ColumnHeaders.count - 2
For i = 1 To LsDataAHP(10).ListItems.count
    a = a + LsDataAHP(10).ListItems(i).SubItems(JklM)
Next
PmbMnLe = Format(a, "0.000")
End Sub

Private Sub HMnLe1()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(10).ListItems.Clear
For Each X In TmpLDataAHP(4).ListItems
    Set cList = LsDataAHP(10).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(4).ListItems.count
        b = TmpLDataAHP(4).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(4).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
        g = f / PmbMnLe
        cList.SubItems(i + 2) = Format(g, "0.000")
    Next
    MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
    "'" & X & "', " & _
    "'" & g & "'," & _
    "'" & "LNR" & "')"
    ConN.Execute MySql
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub JMnLe()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double

a = 0
Set cList = LsDataAHP(10).ListItems.Add(, , "")
Set cList = LsDataAHP(10).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To LsDataAHP(10).ColumnHeaders.count - 1
    For Each X In LsDataAHP(10).ListItems
    a = a + X.SubItems(i)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    a = 0
Next
End Sub

Private Sub LoadNamaAttention()
    Dim cList As ListItem
    With TmpLDataAHP(5)
    .View = lvwReport
    .FullRowSelect = True
    
    .ColumnHeaders.Add , , ""
    .ColumnHeaders.Add , , ""
    
    MySql = "SELECT tb_pegawai.nama, tb_karakter.AHP3 FROM tb_karakter , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    .View = lvwReport
    .ListItems.Clear
        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
    End With
End Sub


'MatrikPerbandinganAttention
Private Sub MatrikPerbandinganAttention()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(5).ListItems.count + 1
    With LsDataAHP(11)
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        Tampungan(5).View = lvwReport
        .ColumnHeaders.Add , , "Kriteria", 2000
        .ColumnHeaders.Add , , "AHP", (.Width - 2300) / a, lvwColumnCenter
        Tampungan(5).ColumnHeaders.Add , , "jumlah"
        For i = 1 To TmpLDataAHP(5).ListItems.count
            .ColumnHeaders.Add , , TmpLDataAHP(5).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
            Tampungan(5).ColumnHeaders.Add , , TmpLDataAHP(5).ListItems(i).Text
        Next
    End With
HMpT
JMpT
End Sub

Private Sub HMpT()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double
With TmpLDataAHP(5)
LsDataAHP(11).ListItems.Clear
For Each X In .ListItems
    Set cList = LsDataAHP(11).ListItems.Add(, , X)
        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
        a = X.ListSubItems(1)
    For i = 1 To .ListItems.count
        b = .ListItems(i).SubItems(1)
        c = a / b
        cList.SubItems(i + 1) = Format(c, "0.000")
    Next
    c = 0
Next
End With
End Sub

Private Sub JMpT()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double
With LsDataAHP(11)
a = 0
Set cList = .ListItems.Add(, , "")
Set cList = .ListItems.Add(, , "Jumlah")
Set cList1 = Tampungan(5).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To .ColumnHeaders.count - 1
    For Each X In LsDataAHP(11).ListItems
    a = a + X.SubItems(i)
    b = b + X.SubItems(i + 1)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    cList1.SubItems(i) = Format(b, "0.000")
    a = 0
    b = 0
Next
End With
End Sub

'MatrikNormalisasiAttention
Private Sub MatrikNormalisasiAttention()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(5).ListItems.count + 2
With LsDataAHP(12)
    .View = lvwReport
    .FullRowSelect = True
    .GridLines = True
    .ColumnHeaders.Add , , "Kriteria", 2000

For i = 1 To TmpLDataAHP(5).ListItems.count
    .ColumnHeaders.Add , , TmpLDataAHP(5).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
Next
.ColumnHeaders.Add , , "Jumlah", (.Width - 2300) / a, lvwColumnRight
.ColumnHeaders.Add , , "Bobot", (.Width - 2300) / a, lvwColumnRight
End With
HMnT
PMnT
HMnT1
JMnT
End Sub

Private Sub HMnT()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(12).ListItems.Clear
For Each X In TmpLDataAHP(5).ListItems
    Set cList = LsDataAHP(12).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(5).ListItems.count
        b = TmpLDataAHP(5).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(5).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
    Next
    
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub PMnT()
Dim i As Integer
Dim a As Double
Dim JklM As Integer
JklM = LsDataAHP(12).ColumnHeaders.count - 2
For i = 1 To LsDataAHP(12).ListItems.count
    a = a + LsDataAHP(12).ListItems(i).SubItems(JklM)
Next
PmbMnT = Format(a, "0.000")
End Sub

Private Sub HMnT1()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(12).ListItems.Clear
For Each X In TmpLDataAHP(5).ListItems
    Set cList = LsDataAHP(12).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(5).ListItems.count
        b = TmpLDataAHP(5).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(5).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
        g = f / PmbMnT
        cList.SubItems(i + 2) = Format(g, "0.000")
    Next
    MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
    "'" & X & "', " & _
    "'" & g & "'," & _
    "'" & "ATN" & "')"
    ConN.Execute MySql
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub JMnT()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double

a = 0
Set cList = LsDataAHP(12).ListItems.Add(, , "")
Set cList = LsDataAHP(12).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To LsDataAHP(12).ColumnHeaders.count - 1
    For Each X In LsDataAHP(10).ListItems
    a = a + X.SubItems(i)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    a = 0
Next
End Sub

Private Sub LoadNamaKinerja()
    Dim cList As ListItem
    With TmpLDataAHP(6)
    .View = lvwReport
    .FullRowSelect = True
    
    .ColumnHeaders.Add , , ""
    .ColumnHeaders.Add , , ""
    
    MySql = "SELECT tb_pegawai.nama, tb_kinerja.AHP FROM tb_kinerja , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_kinerja.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    .View = lvwReport
    .ListItems.Clear
        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
    End With
End Sub


'MatrikPerbandinganKinerja
Private Sub MatrikPerbandinganKinerja()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(6).ListItems.count + 1
    With LsDataAHP(13)
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        Tampungan(6).View = lvwReport
        .ColumnHeaders.Add , , "Kriteria", 2000
        .ColumnHeaders.Add , , "AHP", (.Width - 2300) / a, lvwColumnCenter
        Tampungan(6).ColumnHeaders.Add , , "jumlah"
        For i = 1 To TmpLDataAHP(6).ListItems.count
            .ColumnHeaders.Add , , TmpLDataAHP(6).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
            Tampungan(6).ColumnHeaders.Add , , TmpLDataAHP(6).ListItems(i).Text
        Next
    End With
HMpK
JMpK
End Sub

Private Sub HMpK()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double
With TmpLDataAHP(6)
LsDataAHP(13).ListItems.Clear
For Each X In .ListItems
    Set cList = LsDataAHP(13).ListItems.Add(, , X)
        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
        a = X.ListSubItems(1)
    For i = 1 To .ListItems.count
        b = .ListItems(i).SubItems(1)
        c = a / b
        cList.SubItems(i + 1) = Format(c, "0.000")
    Next
    c = 0
Next
End With
End Sub

Private Sub JMpK()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double
With LsDataAHP(13)
a = 0
Set cList = .ListItems.Add(, , "")
Set cList = .ListItems.Add(, , "Jumlah")
Set cList1 = Tampungan(6).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To .ColumnHeaders.count - 1
    For Each X In LsDataAHP(13).ListItems
    a = a + X.SubItems(i)
    b = b + X.SubItems(i + 1)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    cList1.SubItems(i) = Format(b, "0.000")
    a = 0
    b = 0
Next
End With
End Sub

'MatrikNormalisasiKinerja
Private Sub MatrikNormalisasiKinerja()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(6).ListItems.count + 2
With LsDataAHP(14)
    .View = lvwReport
    .FullRowSelect = True
    .GridLines = True
    .ColumnHeaders.Add , , "Kriteria", 2000

For i = 1 To TmpLDataAHP(6).ListItems.count
    .ColumnHeaders.Add , , TmpLDataAHP(6).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
Next
.ColumnHeaders.Add , , "Jumlah", (.Width - 2300) / a, lvwColumnRight
.ColumnHeaders.Add , , "Bobot", (.Width - 2300) / a, lvwColumnRight
End With
HMnK
PMnK
HMnK1
JMnK
End Sub

Private Sub HMnK()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(14).ListItems.Clear
For Each X In TmpLDataAHP(6).ListItems
    Set cList = LsDataAHP(14).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(6).ListItems.count
        b = TmpLDataAHP(6).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(6).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
    Next
    
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub PMnK()
Dim i As Integer
Dim a As Double
Dim JklM As Integer
JklM = LsDataAHP(14).ColumnHeaders.count - 2
For i = 1 To LsDataAHP(14).ListItems.count
    a = a + LsDataAHP(14).ListItems(i).SubItems(JklM)
Next
PmbMnK = Format(a, "0.000")
End Sub

Private Sub HMnK1()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(14).ListItems.Clear
For Each X In TmpLDataAHP(6).ListItems
    Set cList = LsDataAHP(14).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(6).ListItems.count
        b = TmpLDataAHP(6).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(6).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
        g = f / PmbMnK
        cList.SubItems(i + 2) = Format(g, "0.000")
    Next
    MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
    "'" & X & "', " & _
    "'" & g & "'," & _
    "'" & "KNJ" & "')"
    ConN.Execute MySql
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub JMnK()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double

a = 0
Set cList = LsDataAHP(14).ListItems.Add(, , "")
Set cList = LsDataAHP(14).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To LsDataAHP(14).ColumnHeaders.count - 1
    For Each X In LsDataAHP(10).ListItems
    a = a + X.SubItems(i)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    a = 0
Next
End Sub

'>
Private Sub LoadNamaMsKerja()
    Dim cList As ListItem
    With TmpLDataAHP(7)
    .View = lvwReport
    .FullRowSelect = True
    
    .ColumnHeaders.Add , , ""
    .ColumnHeaders.Add , , ""
    
    MySql = "SELECT tb_pegawai.nama, tb_masakerja.AHP FROM tb_masakerja , tb_pegawai " & _
    "WHERE tb_pegawai.nik_pegawai = tb_masakerja.nik_pegawai ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    .View = lvwReport
    .ListItems.Clear
        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
    End With
End Sub


'MatrikPerbandinganMsKerja
Private Sub MatrikPerbandinganMsKerja()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(7).ListItems.count + 1
    With LsDataAHP(15)
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        Tampungan(7).View = lvwReport
        .ColumnHeaders.Add , , "Kriteria", 2000
        .ColumnHeaders.Add , , "AHP", (.Width - 2300) / a, lvwColumnCenter
        Tampungan(7).ColumnHeaders.Add , , "jumlah"
        For i = 1 To TmpLDataAHP(7).ListItems.count
            .ColumnHeaders.Add , , TmpLDataAHP(7).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
            Tampungan(7).ColumnHeaders.Add , , TmpLDataAHP(7).ListItems(i).Text
        Next
    End With
HMpMsK
JMpMsK
End Sub

Private Sub HMpMsK()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double
With TmpLDataAHP(7)
LsDataAHP(15).ListItems.Clear
For Each X In .ListItems
    Set cList = LsDataAHP(15).ListItems.Add(, , X)
        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
        a = X.ListSubItems(1)
    For i = 1 To .ListItems.count
        b = .ListItems(i).SubItems(1)
        c = a / b
        cList.SubItems(i + 1) = Format(c, "0.000")
    Next
    c = 0
Next
End With
End Sub

Private Sub JMpMsK()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double
With LsDataAHP(15)
a = 0
Set cList = .ListItems.Add(, , "")
Set cList = .ListItems.Add(, , "Jumlah")
Set cList1 = Tampungan(7).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To .ColumnHeaders.count - 1
    For Each X In LsDataAHP(15).ListItems
    a = a + X.SubItems(i)
    b = b + X.SubItems(i + 1)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    cList1.SubItems(i) = Format(b, "0.000")
    a = 0
    b = 0
Next
End With
End Sub

'MatrikNormalisasiMsKerja
Private Sub MatrikNormalisasiMsKerja()
Dim i As Integer
Dim a As Integer
a = TmpLDataAHP(7).ListItems.count + 2
With LsDataAHP(16)
    .View = lvwReport
    .FullRowSelect = True
    .GridLines = True
    .ColumnHeaders.Add , , "Kriteria", 2000

For i = 1 To TmpLDataAHP(7).ListItems.count
    .ColumnHeaders.Add , , TmpLDataAHP(7).ListItems(i).Text, (.Width - 2300) / a, lvwColumnRight
Next
.ColumnHeaders.Add , , "Jumlah", (.Width - 2300) / a, lvwColumnRight
.ColumnHeaders.Add , , "Bobot", (.Width - 2300) / a, lvwColumnRight
End With
HMnMsK
PMnMsK
HMnMsK1
JMnMsK
End Sub

Private Sub HMnMsK()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(16).ListItems.Clear
For Each X In TmpLDataAHP(7).ListItems
    Set cList = LsDataAHP(16).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(7).ListItems.count
        b = TmpLDataAHP(7).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(7).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
    Next
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub PMnMsK()
Dim i As Integer
Dim a As Double
Dim JklM As Integer
JklM = LsDataAHP(16).ColumnHeaders.count - 2
For i = 1 To LsDataAHP(16).ListItems.count
    a = a + LsDataAHP(16).ListItems(i).SubItems(JklM)
Next
PmbMnMsK = Format(a, "0.000")
End Sub

Private Sub HMnMsK1()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double, d As Double, f As Double, g As Double

LsDataAHP(16).ListItems.Clear
For Each X In TmpLDataAHP(7).ListItems
    Set cList = LsDataAHP(16).ListItems.Add(, , X)

        cList.SubItems(1) = Format(X.ListSubItems(1), "0.000")
    a = X.ListSubItems(1)
    For i = 1 To TmpLDataAHP(7).ListItems.count
        b = TmpLDataAHP(7).ListItems(i).SubItems(1)
        c = a / b
        d = Tampungan(7).ListItems(1).SubItems(i)
        e = c / d
        f = f + e
        cList.SubItems(i) = Format(e, "0.000")
        cList.SubItems(i + 1) = Format(f, "0.000")
        g = f / PmbMnMsK
        cList.SubItems(i + 2) = Format(g, "0.000")
    Next
        MySql = "INSERT INTO tb_bbt_ahp (nama, nilai, kode) VALUES ( " & _
    "'" & X & "', " & _
    "'" & g & "'," & _
    "'" & "MSK" & "')"
    ConN.Execute MySql
    c = 0
    e = 0
    f = 0
Next
End Sub

Private Sub JMnMsK()
Dim i As Integer
Dim X As ListItem
Dim cList As ListItem, cList1 As ListItem
Dim a As Double, b As Double, c As Double

a = 0
Set cList = LsDataAHP(16).ListItems.Add(, , "")
Set cList = LsDataAHP(16).ListItems.Add(, , "Jumlah")
cList.Bold = True
cList.ForeColor = vbBlue
On Error Resume Next
For i = 1 To LsDataAHP(16).ColumnHeaders.count - 1
    For Each X In LsDataAHP(10).ListItems
    a = a + X.SubItems(i)
    Next
    cList.SubItems(i) = Format(a, "0.000")
    cList.ListSubItems(i).Bold = True
    cList.ListSubItems(i).ForeColor = vbBlue
    a = 0
Next
End Sub

'>
Private Sub NilaiPerbandingan()
    Dim cList As ListItem
    With LsDataAHP(17)
        .ColumnHeaders.Add , , "Nama", 4000
        .ColumnHeaders.Add , , "Leadership", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Learning", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Attention", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Kinerja", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Masa Kerja", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Pendidikan", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Pengalaman", (.Width - 4300) / 7, lvwColumnRight

    MySql = "SELECT DISTINCT nama," & _
    "SUM(CASE WHEN kode = 'LDR' THEN nilai ELSE 0 END) AS 'Leadership', " & _
    "SUM(CASE WHEN kode = 'LNR' THEN nilai ELSE 0 END) AS 'Learning', " & _
    "SUM(CASE WHEN kode = 'ATN' THEN nilai ELSE 0 END) AS 'Attention', " & _
    "SUM(CASE WHEN kode = 'KNJ' THEN nilai ELSE 0 END) AS 'Kinerja', " & _
    "SUM(CASE WHEN kode = 'MSK' THEN nilai ELSE 0 END) AS 'MsKerja', " & _
    "SUM(CASE WHEN kode = 'PND' THEN nilai ELSE 0 END) AS 'Pendidikan', " & _
    "SUM(CASE WHEN kode = 'PNG' THEN nilai ELSE 0 END) AS 'Pengalaman' " & _
    "FROM tb_bbt_ahp GROUP BY nama ORDER BY nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic
        .View = lvwReport
        .FullRowSelect = True
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

Private Sub NilaiNormalisasiGlobal()
    Dim cList As ListItem
    With LsDataAHP(18)
        .ColumnHeaders.Add , , "Nama", 4000
        .ColumnHeaders.Add , , "Leadership", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Learning", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Attention", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Kinerja", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Masa Kerja", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Pendidikan", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Pengalaman", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Rangking", (.Width - 4300) / 8, lvwColumnRight

    MySql = "SELECT DISTINCT nama,      SUM(CASE WHEN kode = 'LDR' THEN nilai * bnb.Leader ELSE 0 END) AS 'Leadership',      SUM(CASE WHEN kode = 'LNR' THEN nilai * bnb.Learn ELSE 0 END) AS 'Learning',      SUM(CASE WHEN kode = 'ATN' THEN nilai * bnb.Atten ELSE 0 END) AS 'Attention',      SUM(CASE WHEN kode = 'KNJ' THEN convert(nilai * bnb.Kiner,decimal(4,3)) ELSE 0 END) AS 'Kinerja',      SUM(CASE WHEN kode = 'MSK' THEN nilai * bnb.MsKer ELSE 0 END) AS 'MsKerja',      SUM(CASE WHEN kode = 'PND' THEN nilai * bnb.Pend ELSE 0 END) AS 'Pendidikan',      SUM(CASE WHEN kode = 'PNG' THEN convert(nilai * bnb.Peng,decimal(4,3)) ELSE 0 END) AS 'Pengalaman'      FROM tb_bbt_ahp, bnb GROUP BY nama ORDER BY nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic
        .View = lvwReport
        .FullRowSelect = True
        .ListItems.Clear
        Dim clm1 As Double
        Dim clm2 As Double
        Dim clm3 As Double
        Dim clm4 As Double
        Dim clm5 As Double
        Dim clm6 As Double
        Dim clm7 As Double
        Dim clm8 As Double
        
        Do Until SdR.EOF
         clm1 = SdR.Fields(1)
         clm2 = SdR.Fields(2)
         clm3 = SdR.Fields(3)
         clm4 = SdR.Fields(4)
         clm5 = SdR.Fields(5)
         clm6 = SdR.Fields(6)
         clm7 = SdR.Fields(7)
         Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = Format(clm1, "0.000")
                cList.SubItems(2) = Format(clm2, "0.000")
                cList.SubItems(3) = Format(clm3, "0.000")
                cList.SubItems(4) = Format(clm4, "0.000")
                cList.SubItems(5) = Format(clm5, "0.000")
                cList.SubItems(6) = Format(clm6, "0.000")
                cList.SubItems(7) = Format(clm7, "0.000")
                cList.SubItems(8) = Format(clm1 + _
                                            clm2 + _
                                            clm3 + _
                                            clm4 + _
                                            clm5 + _
                                            clm6 + _
                                            clm7, "0.000")
                cList.ListSubItems(8).Bold = True
                cList.ListSubItems(8).ForeColor = vbBlue
            SdR.MoveNext
        Loop
                Dim i As Integer
        For i = 1 To .ListItems.count
            MySql = "INSERT INTO rangking (nama, nilai, kode) VALUES ( " & _
            "'" & .ListItems(i) & "', " & _
            "'" & .ListItems(i).SubItems(8) & "'," & _
            "'" & "AHP" & "')"
            ConN.Execute MySql
        Next
    End With
End Sub
Private Sub Bersih()
    MySql = "DELETE FROM coba"
    ConN.Execute MySql
    
    MySql = "DELETE FROM tb_bbt_ahp"
    ConN.Execute MySql
    
    MySql = "DELETE FROM rangking"
    ConN.Execute MySql
End Sub


Private Sub LoadKarakterSAW()
    Dim cList As ListItem
    With LsDataSAW(0)
    .ColumnHeaders.Add , , "Nama", 4000
    .ColumnHeaders.Add , , "Leadership", (.Width - 4300) / 6
    .ColumnHeaders.Add , , "Nilai", (.Width - 4300) / 6, lvwColumnRight
    .ColumnHeaders.Add , , "Learning", (.Width - 4300) / 6
    .ColumnHeaders.Add , , "Nilai", (.Width - 4300) / 6, lvwColumnRight
    .ColumnHeaders.Add , , "Attention", (.Width - 4300) / 6
    .ColumnHeaders.Add , , "Nilai", (.Width - 4300) / 6, lvwColumnRight
    
    MySql = "SELECT tb_pegawai.nama, " & _
    "tb_karakter.leadership_abilitiy, " & _
    "tb_karakter.SAW1, " & _
    "tb_karakter.learning_abilitiy, " & _
    "tb_karakter.SAW2, " & _
    "tb_karakter.attention_to_detail, " & _
    "tb_karakter.SAW3 " & _
    "FROM tb_pegawai, tb_karakter " & _
    "WHERE tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai " & _
    "ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    .View = lvwReport
    .ListItems.Clear
    .FullRowSelect = True
    
        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
                cList.SubItems(3) = SdR.Fields(3)
                cList.SubItems(4) = SdR.Fields(4)
                cList.SubItems(5) = SdR.Fields(5)
                cList.SubItems(6) = SdR.Fields(6)
            SdR.MoveNext
        Loop
    End With
End Sub

Private Sub MatrikPerbandinganSAW()
    Dim cList As ListItem
    With LsDataSAW(1)
        .ColumnHeaders.Add , , "Nama", 4000
        .ColumnHeaders.Add , , "Leadership", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Learning", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Attention", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Kinerja", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Masa Kerja", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Pendidikan", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Pengalaman", (.Width - 4300) / 7, lvwColumnRight

    MySql = "SELECT DISTINCT tb_pegawai.nama, " & _
    "tb_karakter.SAW1 AS Leadership, " & _
    "tb_karakter.SAW2 AS Learning, " & _
    "tb_karakter.SAW3 AS Attention, " & _
    "tb_kinerja.SAW AS Kinerja, " & _
    "tb_masakerja.SAW AS Masakerja, " & _
    "tb_pendidikan.SAW AS Pendidikan, " & _
    "tb_pengalaman.SAW AS Pengalaman " & _
    "FROM tb_pegawai, tb_karakter, tb_kinerja, tb_masakerja, tb_pengalaman, tb_pendidikan " & _
    "WHERE tb_pegawai.nik_pegawai = tb_pendidikan.nik_pegawai " & _
    "AND tb_pegawai.nik_pegawai = tb_pengalaman.nik_pegawai " & _
    "AND tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai " & _
    "AND tb_pegawai.nik_pegawai = tb_kinerja.nik_pegawai " & _
    "AND tb_pegawai.nik_pegawai = tb_masakerja.nik_pegawai " & _
    "GROUP BY tb_pegawai.nama " & _
    "ORDER BY tb_pegawai.nama ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

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
        
    MySql = "SELECT DISTINCT " & _
    "max(tb_karakter.SAW1) AS Leadership, " & _
    "max(tb_karakter.SAW2) AS Learning, " & _
    "max(tb_karakter.SAW3) AS Attention, " & _
    "max(tb_kinerja.SAW) AS Kinerja, " & _
    "max(tb_masakerja.SAW) AS Masakerja, " & _
    "max(tb_pendidikan.SAW) AS Pendidikan, " & _
    "max(tb_pengalaman.SAW) AS Pengalaman " & _
    "FROM tb_pegawai, tb_karakter, tb_kinerja, tb_masakerja, tb_pengalaman, tb_pendidikan " & _
    "WHERE tb_pegawai.nik_pegawai = tb_pendidikan.nik_pegawai " & _
    "AND tb_pegawai.nik_pegawai = tb_pengalaman.nik_pegawai " & _
    "AND tb_pegawai.nik_pegawai = tb_karakter.nik_pegawai " & _
    "AND tb_pegawai.nik_pegawai = tb_kinerja.nik_pegawai " & _
    "AND tb_pegawai.nik_pegawai = tb_masakerja.nik_pegawai "
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , "")
             Set cList = .ListItems.Add(, , "Nilai Max")
                cList.Bold = True
                cList.ForeColor = vbBlue
                cList.SubItems(1) = Format(SdR.Fields(0), "0.000")
                cList.ListSubItems(1).Bold = True
                cList.ListSubItems(1).ForeColor = vbBlue
                cList.SubItems(2) = Format(SdR.Fields(1), "0.000")
                cList.ListSubItems(2).Bold = True
                cList.ListSubItems(2).ForeColor = vbBlue
                cList.SubItems(3) = Format(SdR.Fields(2), "0.000")
                cList.ListSubItems(3).Bold = True
                cList.ListSubItems(3).ForeColor = vbBlue
                cList.SubItems(4) = Format(SdR.Fields(3), "0.000")
                cList.ListSubItems(4).Bold = True
                cList.ListSubItems(4).ForeColor = vbBlue
                cList.SubItems(5) = Format(SdR.Fields(4), "0.000")
                cList.ListSubItems(5).Bold = True
                cList.ListSubItems(5).ForeColor = vbBlue
                cList.SubItems(6) = Format(SdR.Fields(5), "0.000")
                cList.ListSubItems(6).Bold = True
                cList.ListSubItems(6).ForeColor = vbBlue
                cList.SubItems(7) = Format(SdR.Fields(6), "0.000")
                cList.ListSubItems(7).Bold = True
                cList.ListSubItems(7).ForeColor = vbBlue
            SdR.MoveNext
        Loop
    End With
End Sub

Private Sub MatrikNormalisasiSAW()
    Dim cList As ListItem
    With LsDataSAW(2)
        .ColumnHeaders.Add , , "Nama", 4000
        .ColumnHeaders.Add , , "Leadership", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Learning", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Attention", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Kinerja", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Masa Kerja", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Pendidikan", (.Width - 4300) / 7, lvwColumnRight
        .ColumnHeaders.Add , , "Pengalaman", (.Width - 4300) / 7, lvwColumnRight

    MySql = " SELECT DISTINCT `tb_pegawai`.`nama` AS `nama`, " & _
            "(`tb_karakter`.`SAW1` / `nilaimax`.`Leadership`) AS `Leadership`, " & _
            "(`tb_karakter`.`SAW2` / `nilaimax`.`Learning`) AS `Learning`, " & _
            "(`tb_karakter`.`SAW3` / `nilaimax`.`Attention`) AS `Attention`, " & _
            "(`tb_kinerja`.`SAW` / `nilaimax`.`Kinerja`) AS `Kinerja`, " & _
            "(`tb_masakerja`.`SAW` / `nilaimax`.`Masakerja`) AS `Masakerja`, " & _
            "(`tb_pendidikan`.`SAW` / `nilaimax`.`Pendidikan`) AS `Pendidikan`, " & _
            "(`tb_pengalaman`.`SAW` / `nilaimax`.`Pengalaman`) AS `Pengalaman` " & _
            "FROM((((((`tb_pegawai` JOIN `tb_karakter`) " & _
            "JOIN `tb_kinerja`) " & _
            "JOIN `tb_masakerja`) " & _
            "JOIN `nilaimax`)" & _
            "JOIN `tb_pengalaman`) " & _
            "JOIN `tb_pendidikan`) " & _
            "Where((`tb_pegawai`.`nik_pegawai` = `tb_pendidikan`.`nik_pegawai`) " & _
            "AND (`tb_pegawai`.`nik_pegawai` = `tb_pengalaman`.`nik_pegawai`) " & _
            "AND (`tb_pegawai`.`nik_pegawai` = `tb_karakter`.`nik_pegawai`) " & _
            "AND (`tb_pegawai`.`nik_pegawai` = `tb_kinerja`.`nik_pegawai`) " & _
            "AND (`tb_pegawai`.`nik_pegawai` = `tb_masakerja`.`nik_pegawai`)) " & _
            "GROUP BY `tb_pegawai`.`nama` " & _
            "ORDER BY `tb_pegawai`.`nama`"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

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

Private Sub MatrikNormalisasiGlobalSAW()
    Dim cList As ListItem
    With LsDataSAW(3)
        .ColumnHeaders.Add , , "Nama", 4000
        .ColumnHeaders.Add , , "Leadership", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Learning", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Attention", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Kinerja", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Masa Kerja", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Pendidikan", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Pengalaman", (.Width - 4300) / 8, lvwColumnRight
        .ColumnHeaders.Add , , "Rangking", (.Width - 4300) / 8, lvwColumnRight

    MySql = "SELECT normalisasimax.nama, normalisasimax.Leadership * hasilbagisaw.leadership AS Leadership, " & _
    "normalisasimax.Learning * hasilbagisaw.Learning AS Learning, " & _
    "normalisasimax.Attention * hasilbagisaw.Attenton AS Attenton, " & _
    "normalisasimax.Kinerja * hasilbagisaw.Kinerja AS Kinerja, " & _
    "normalisasimax.Masakerja * hasilbagisaw.MasaKerja AS MasaKerja, " & _
    "normalisasimax.Pendidikan * hasilbagisaw.Pendidikan AS Pendidikan, " & _
    "normalisasimax.Pengalaman * hasilbagisaw.Pengalaman AS Pengalaman " & _
    "FROM normalisasimax , hasilbagisaw"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic
        Dim clm1 As Double
        Dim clm2 As Double
        Dim clm3 As Double
        Dim clm4 As Double
        Dim clm5 As Double
        Dim clm6 As Double
        Dim clm7 As Double
        Dim clm8 As Double
        
        Do Until SdR.EOF
         clm1 = SdR.Fields(1)
         clm2 = SdR.Fields(2)
         clm3 = SdR.Fields(3)
         clm4 = SdR.Fields(4)
         clm5 = SdR.Fields(5)
         clm6 = SdR.Fields(6)
         clm7 = SdR.Fields(7)
         Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = Format(clm1, "0.000")
                cList.SubItems(2) = Format(clm2, "0.000")
                cList.SubItems(3) = Format(clm3, "0.000")
                cList.SubItems(4) = Format(clm4, "0.000")
                cList.SubItems(5) = Format(clm5, "0.000")
                cList.SubItems(6) = Format(clm6, "0.000")
                cList.SubItems(7) = Format(clm7, "0.000")
                cList.SubItems(8) = Format(clm1 + _
                                            clm2 + _
                                            clm3 + _
                                            clm4 + _
                                            clm5 + _
                                            clm6 + _
                                            clm7, "0.000")
                cList.ListSubItems(8).Bold = True
                cList.ListSubItems(8).ForeColor = vbBlue
            SdR.MoveNext
        Loop
        
        Dim i As Integer
        For i = 1 To .ListItems.count
    
            MySql = "INSERT INTO rangking (nama, nilai, kode) VALUES ( " & _
            "'" & .ListItems(i) & "', " & _
            "'" & .ListItems(i).SubItems(8) & "'," & _
            "'" & "SAW" & "')"
            ConN.Execute MySql
        Next
    End With
End Sub

