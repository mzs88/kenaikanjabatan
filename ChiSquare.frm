VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ChiSquare 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "ChiSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    
    MySql = "SELECT nilai FROM tb_chisquare WHERE ve= '" & Combo2.Text & "'  AND alfa = '" & Combo1.Text & "'"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic
        SdR.Requery
        If SdR.EOF = False Then
            Text1.Text = SdR(0)
        End If
        SdR.Close
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Move 0, 0, Penilaian.ScaleWidth, _
    Penilaian.ScaleHeight - Penilaian.Toolbar1.Height
cNDb
LoadChisquare
LoadAlfa
LoadVe
End Sub

Private Sub LoadChisquare()
    Dim cList As ListItem
    With ListView1
        .ColumnHeaders.Add , , "", 500
        .ColumnHeaders.Add , , "0.995", (.Width - 530) / 11, lvwColumnCenter
        .ColumnHeaders.Add , , "0.975", (.Width - 530) / 11, lvwColumnCenter
        .ColumnHeaders.Add , , "0.2", (.Width - 530) / 11, lvwColumnCenter
        .ColumnHeaders.Add , , "0.1", (.Width - 530) / 11, lvwColumnCenter
        .ColumnHeaders.Add , , "0.05", (.Width - 530) / 11, lvwColumnCenter
        .ColumnHeaders.Add , , "0.025", (.Width - 530) / 11, lvwColumnCenter
        .ColumnHeaders.Add , , "0.02", (.Width - 530) / 11, lvwColumnCenter
        .ColumnHeaders.Add , , "0.01", (.Width - 530) / 11, lvwColumnCenter
        .ColumnHeaders.Add , , "0.005", (.Width - 530) / 11, lvwColumnCenter
        .ColumnHeaders.Add , , "0.002", (.Width - 530) / 11, lvwColumnCenter
        .ColumnHeaders.Add , , "0.001", (.Width - 530) / 11, lvwColumnCenter
        
    MySql = " SELECT DISTINCT ve, sum(CASE WHEN alfa = 0.995 THEN nilai ELSE 0 END) AS '0.995', sum(CASE WHEN alfa = 0.975 THEN nilai ELSE 0 END) AS '0.975', sum(CASE WHEN alfa = 0.2 THEN nilai ELSE 0 END) AS '0.2', sum(CASE WHEN alfa = 0.1 THEN nilai ELSE 0 END) AS '0.1', sum(CASE WHEN alfa = 0.05 THEN nilai ELSE 0 END) AS '0.05', sum(CASE WHEN alfa = 0.025 THEN nilai ELSE 0 END) AS '0.025', sum(CASE WHEN alfa = 0.02 THEN nilai ELSE 0 END) AS '0.02', sum(CASE WHEN alfa = 0.01 THEN nilai ELSE 0 END) AS '0.01', sum(CASE WHEN alfa = 0.005 THEN nilai ELSE 0 END) AS '0.005', sum(CASE WHEN alfa = 0.002 THEN nilai ELSE 0 END) AS '0.002', sum(CASE WHEN alfa = 0.001 THEN nilai ELSE 0 END) AS '0.001' FROM tb_chisquare GROUP BY ve"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

        Do Until SdR.EOF
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
                cList.SubItems(3) = SdR.Fields(3)
                cList.SubItems(4) = SdR.Fields(4)
                cList.SubItems(5) = SdR.Fields(5)
                cList.SubItems(6) = SdR.Fields(6)
                cList.SubItems(7) = SdR.Fields(7)
                cList.SubItems(8) = SdR.Fields(8)
                cList.SubItems(9) = SdR.Fields(9)
                cList.SubItems(10) = SdR.Fields(10)
                cList.SubItems(11) = SdR.Fields(11)
        SdR.MoveNext
        Loop
    End With
End Sub

Private Sub LoadAlfa()
    MySql = " SELECT DISTINCT tb_chisquare.alfa FROM `tb_chisquare`"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

        Do Until SdR.EOF
             Combo1.AddItem (SdR.Fields(0))
        SdR.MoveNext
        Loop

End Sub

Private Sub LoadVe()
    MySql = " SELECT DISTINCT tb_chisquare.ve FROM `tb_chisquare` ORDER BY ve ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

        Do Until SdR.EOF
             Combo2.AddItem (SdR.Fields(0))
        SdR.MoveNext
        Loop

End Sub

Private Sub Form_Resize()
ListView1.Move 0, Me.Text1.Height + 200, Me.ScaleWidth, Me.ScaleHeight
End Sub
