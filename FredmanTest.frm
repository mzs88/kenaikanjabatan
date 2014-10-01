VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FredmanTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fredman Test"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmb_alfa 
      Height          =   315
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   0
      Width           =   1455
   End
   Begin VB.ComboBox cmb_volume 
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Result"
      Height          =   315
      Left            =   7080
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4440
      Width           =   7935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox txt_result 
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Alfa"
      Height          =   315
      Left            =   5040
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Volume"
      Height          =   315
      Left            =   2640
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Perangkingan"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "FredmanTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim k1 As Double
Dim n1 As Double

Dim k2 As Double
Dim n2 As Double

Dim hasil1 As Double
Dim hasil2 As Double
Dim hasil3 As Double


Private Sub Command1_Click()
    MySql = "SELECT nilai FROM tb_chisquare WHERE ve= '" & cmb_volume.Text & "'  AND alfa = '" & cmb_alfa.Text & "'"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic
        SdR.Requery
        If SdR.EOF = False Then
            txt_result.Text = SdR(0)
        End If
        SdR.Close
        ListView1.ColumnHeaders.Clear
        Kolom
        Text1.Text = "Critical Value tabel Chi-Square dengan nilai derajat kebebasan (df)2 dan alfa = 5% adalah '" & txt_result.Text & "', sehingga H0" & _
                        " ditolak. jadi, kesimpulannya dari hasil peringkingan kedua metode mempunyai hasil yang sama"
    End Sub

Private Sub Combo1_Click()
'k1 = 0
'n1 = 0
'hasil2 = 0
'ListView1.ColumnHeaders.Clear

'Kolom
End Sub

Private Sub Form_Load()
cNDb
Rangking
LoadAlfa
LoadVe
End Sub

Private Sub Kolom()
MySql = "SELECT rangking.nama, sum(case when kode='ahp' then nilai else 0 end) as 'AHP', " & _
        "sum(case when kode='saw' then nilai else 0 end) as 'SAW' FROM `rangking` GROUP BY nama ORDER BY nilai desc LIMIT " & Combo1.Text & ""
    Set SdR = New ADODB.Recordset
    SdR.CursorLocation = adUseClient
    SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    
    With ListView1
        .View = lvwReport
        
        .ColumnHeaders.Add , , "Pemohon", 2000
        .ColumnHeaders.Add , , "AHP", (.Width - 2100) / 2, lvwColumnRight
        .ColumnHeaders.Add , , "SAW", (.Width - 2100) / 2, lvwColumnRight
        .ListItems.Clear
        Do Until SdR.EOF
        Dim data1 As Double
        Dim data2 As Double
        data1 = SdR.Fields(1)
        data2 = SdR.Fields(2)
             Set cList = .ListItems.Add(, , SdR.Fields(0))
                Dim i As Integer
                Dim a As Integer
                For i = 1 To .ListItems.count
                     cList.SubItems(1) = a + i
                     cList.SubItems(2) = a + i
                Next
                
            SdR.MoveNext
        Loop
        
            Dim h As Integer
            For h = 1 To .ListItems.count
                Dim b As Integer
                b = b + .ListItems(h).SubItems(1)
            Next
            
            Set cList = .ListItems.Add(, , "Rj")
                cList.SubItems(1) = b
                cList.SubItems(2) = b
                
            Set cList = .ListItems.Add(, , "R" & ChrW$(178) & "j")
                cList.SubItems(1) = b ^ 2
                cList.SubItems(2) = b ^ 2
                
            Set cList = .ListItems.Add(, , "Jumlah Kolom (K)")
                cList.SubItems(1) = .ColumnHeaders.count - 1
                
            Set cList = .ListItems.Add(, , "Jumlah Barin (N)")
                cList.SubItems(1) = .ListItems.count - 4
                
            Set cList = .ListItems.Add(, , ChrW$(207) & "R" & ChrW$(178))
                cList.SubItems(1) = b ^ 2 & "+" & b ^ 2 & "=" & b ^ 2 + b ^ 2
                                    hasil1 = b ^ 2 + b ^ 2
                                    
            Set cList = .ListItems.Add(, , "12/nk(k +1)")
                k1 = .ColumnHeaders.count - 1
                n1 = .ListItems.count - 6
                hasil2 = 12 / ((n1) * (k1) * (k1 + 1))
                cList.SubItems(1) = ""
                cList.SubItems(1) = "12/" & n1 & "*" & k1 & "*" & k1 + 1 & "=" & Format(hasil2, "0.0")
                
            Set cList = .ListItems.Add(, , "3n(k+1)")
                k2 = .ColumnHeaders.count - 1
                n2 = .ListItems.count - 7
                hasil3 = 3 * (n2) * (k2 + 1)
                cList.SubItems(1) = "3 *" & (n2) & "*" & (k2 + 1) & "=" & hasil3
                                    
            Set cList = .ListItems.Add(, , "Tes Statistic M")
                cList.SubItems(1) = Format(hasil2, "0.0") & "*" & hasil1 & "-" & hasil3 & "=" & (hasil2 * hasil1) - hasil3
            Set cList = .ListItems.Add(, , "Hasil Table Chisquare")
                cList.SubItems(1) = txt_result.Text
        End With
    End Sub

Private Sub Rangking()
    Dim i As Integer
    For i = 1 To 10
        Combo1.AddItem (i)
    Next
End Sub

Private Sub LoadAlfa()
    MySql = " SELECT DISTINCT tb_chisquare.alfa FROM `tb_chisquare`"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

        Do Until SdR.EOF
             cmb_alfa.AddItem (SdR.Fields(0))
        SdR.MoveNext
        Loop

End Sub

Private Sub LoadVe()
    MySql = " SELECT DISTINCT tb_chisquare.ve FROM `tb_chisquare` ORDER BY ve ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

        Do Until SdR.EOF
             cmb_volume.AddItem (SdR.Fields(0))
        SdR.MoveNext
        Loop

End Sub

