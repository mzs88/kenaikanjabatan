VERSION 5.00
Begin VB.Form Splash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   5280
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1200
      Left            =   6240
      Top             =   1680
   End
   Begin VB.Label lblDisp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checking Database ..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6360
      TabIndex        =   0
      Top             =   4920
      Width           =   3360
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
        Static count As Integer
        count = count + 1
        
        If count = 1 Then
            lblDisp = "Software Initialized ..."
                
        ElseIf count = 2 Then
            lblDisp = "Menyiapkan Database ..."
            
        ElseIf count = 3 Then
            lblDisp = "Menyiapkan Aplikasi..."
        
        ElseIf count = 4 Then
            lblDisp = "Wait..."
          
        ElseIf count = 5 Then
            Timer1.Enabled = False
            Unload Me
            MenuUtama.Show
        End If
End Sub

