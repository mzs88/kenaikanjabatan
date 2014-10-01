VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FredmanTest 
   Caption         =   "Fredman Test"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9975
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
   Begin VB.Label Label1 
      Caption         =   "Perangkingan"
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   1455
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
Dim AscChar As Integer
AscChar = 228
MsgBox "My Ascii character is " & Chr(AscChar)
End Sub

Private Sub Combo1_Click()
k1 = 0
n1 = 0
hasil2 = 0
ListView1.ColumnHeaders.Clear

Kolom
End Sub

Private Sub Form_Load()
cNDb
Rangking
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
        End With
    End Sub

Private Sub Rangking()
    Dim i As Integer
    For i = 1 To 10
        Combo1.AddItem (i)
    Next
End Sub

