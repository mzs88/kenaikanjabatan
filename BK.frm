VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BK 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bobot Kriteria"
   ClientHeight    =   4845
   ClientLeft      =   225
   ClientTop       =   675
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnProses 
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
      Height          =   315
      Left            =   4800
      Picture         =   "BK.frx":0FA2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   735
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
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   5415
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   3
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
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "saw"
            Object.Width           =   882
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5415
      Begin VB.CommandButton BtnSimpan 
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
         Height          =   315
         Left            =   4560
         Picture         =   "BK.frx":19A4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton BtnHapus 
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
         Height          =   315
         Left            =   4560
         Picture         =   "BK.frx":23A6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox CmbSkala 
         Height          =   330
         ItemData        =   "BK.frx":2DA8
         Left            =   1080
         List            =   "BK.frx":2DBB
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox CmbKriteria 
         Height          =   330
         ItemData        =   "BK.frx":2DCE
         Left            =   1080
         List            =   "BK.frx":2DE1
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Skala"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Kriteria"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Bobot Kriteria"
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
      TabIndex        =   8
      Top             =   120
      Width           =   5400
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "BK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TKarakter As Single
Dim TKinerja As Single
Dim TMasaKerja As Single
Dim TPendidikan As Single
Dim TPengalaman As Single

Dim TNmKarakter As Single
Dim TNmKinerja As Single
Dim TNmMasaKerja As Single
Dim TNmPendidikan As Single
Dim TNmPengalaman As Single

Private Sub Command1_Click()
    TotColum
End Sub

Private Sub BtnProses_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cNDb
    LoadKriteria
    End Sub
Private Sub Kolom()
 
    Dim c As Integer
    For c = 1 To ListView1.ListItems.count
        With ListView2
            .ColumnHeaders.Add , , ListView1.ListItems(c).Text
        End With
        
        With ListView3
            .ColumnHeaders.Add , , ListView1.ListItems(c).Text
        End With
    
        With ListView4
            .ColumnHeaders.Add , , ListView1.ListItems(c).Text
        End With
    
        With ListView5
            .ColumnHeaders.Add , , ListView1.ListItems(c).Text
        End With
    Next
   
End Sub

Private Sub LoadKriteria()
    Dim cList As ListItem

    MySql = "SELECT nama_kriteria, ahp, saw FROM tb_kriteria ORDER BY nama_kriteria ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    ListView1.View = lvwReport
    ListView1.ListItems.Clear
        Do Until SdR.EOF
             Set cList = ListView1.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
                cList.SubItems(2) = SdR.Fields(2)
            SdR.MoveNext
        Loop
End Sub

Private Sub BtnSimpan_Click()
    MySql = "INSERT INTO tb_kriteria (nama_kriteria, ahp) VALUES ('" & CmbKriteria.Text & "', '" & CmbSkala.Text & "')"
    ConN.Execute MySql
    MsgBox ("Data Berhasil Disimpan")
    LoadKriteria
End Sub

Private Sub BtnHapus_Click()
    MySql = "DELETE FROM tb_kriteria"
    ConN.Execute MySql
    MsgBox ("Data Berhasil Dihapus")
    LoadKriteria
End Sub


Private Sub BobotKriteria()
Dim cList As ListItem
        Dim cIndex As Long
        Dim Karakter As Single
        Dim Kinerja As Single
        Dim MasaKerja As Single
        Dim Pendidikan As Single
        Dim Pengalaman As Single
        Dim Kriteria As String
        
        For cIndex = 1 To ListView1.ListItems.count
            Karakter = Val(ListView1.ListItems(cIndex).SubItems(1)) / Val(ListView1.ListItems(1).SubItems(1))
            Kinerja = Val(ListView1.ListItems(cIndex).SubItems(1)) / Val(ListView1.ListItems(2).SubItems(1))
            MasaKerja = Val(ListView1.ListItems(cIndex).SubItems(1)) / Val(ListView1.ListItems(3).SubItems(1))
            Pendidikan = Val(ListView1.ListItems(cIndex).SubItems(1)) / Val(ListView1.ListItems(4).SubItems(1))
            Pengalaman = Val(ListView1.ListItems(cIndex).SubItems(1)) / Val(ListView1.ListItems(5).SubItems(1))
            Kriteria = ListView1.ListItems(cIndex)
            
            Set cList = ListView2.ListItems.Add(, , Kriteria)
                cList.SubItems(1) = Format(Karakter, "0.000")
                cList.SubItems(2) = Format(Kinerja, "0.000")
                cList.SubItems(3) = Format(MasaKerja, "0.000")
                cList.SubItems(4) = Format(Pendidikan, "0.000")
                cList.SubItems(5) = Format(Pengalaman, "0.000")
        Next
        
End Sub

Private Sub Normalisasi()
    Dim cListNm As ListItem
        Dim cIndexNm As Long
        Dim NmKarakter As Single
        Dim NmKinerja As Single
        Dim NmMasaKerja As Single
        Dim NmPendidikan As Single
        Dim NmPengalaman As Single
        Dim NmKriteria As String
        
        For cIndexNm = 1 To ListView2.ListItems.count
           NmKarakter = Val(ListView2.ListItems(cIndexNm).SubItems(1)) / Val(TKarakter)
           NmKinerja = Val(ListView2.ListItems(cIndexNm).SubItems(2)) / Val(TKinerja)
           NmMasaKerja = Val(ListView2.ListItems(cIndexNm).SubItems(3)) / Val(TMasaKerja)
           NmPendidikan = Val(ListView2.ListItems(cIndexNm).SubItems(4)) / Val(TPendidikan)
           NmPengalaman = Val(ListView2.ListItems(cIndexNm).SubItems(5)) / Val(TPengalaman)
           NmKriteria = ListView2.ListItems(cIndexNm)
           
           Set cListNm = ListView4.ListItems.Add(, , NmKriteria)
                cListNm.SubItems(1) = Format(NmKarakter, "0.000")
                cListNm.SubItems(2) = Format(NmKinerja, "0.000")
                cListNm.SubItems(3) = Format(NmMasaKerja, "0.000")
                cListNm.SubItems(4) = Format(NmPendidikan, "0.000")
                cListNm.SubItems(5) = Format(NmPengalaman, "0.000")
      Next
      
      Dim NjmL As ListItem
      Dim cIndexTNm As Long
        For cIndexTNm = 1 To ListView4.ListItems.count
            TNmKarakter = TNmKarakter + ListView4.ListItems(cIndexTNm).SubItems(1)
            TNmKinerja = TNmKinerja + ListView4.ListItems(cIndexTNm).SubItems(2)
            TNmMasaKerja = TNmMasaKerja + ListView4.ListItems(cIndexTNm).SubItems(3)
            TNmPendidikan = TNmPendidikan + ListView4.ListItems(cIndexTNm).SubItems(4)
            TNmPengalaman = TNmPengalaman + ListView4.ListItems(cIndexTNm).SubItems(5)
        Next
        
        
    
End Sub

Private Sub TotKrK()
    Dim lngIndex As Long
        For lngIndex = 1 To ListView2.ListItems.count
            TKarakter = TKarakter + ListView2.ListItems(lngIndex).SubItems(1)
            TKinerja = TKinerja + ListView2.ListItems(lngIndex).SubItems(2)
            TMasaKerja = TMasaKerja + ListView2.ListItems(lngIndex).SubItems(3)
            TPendidikan = TPendidikan + ListView2.ListItems(lngIndex).SubItems(4)
            TPengalaman = TPengalaman + ListView2.ListItems(lngIndex).SubItems(5)
        Next
        
         Set cList = ListView3.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = Format(TKarakter, "0.000")
                cList.SubItems(2) = Format(TKinerja, "0.000")
                cList.SubItems(3) = Format(TMasaKerja, "0.000")
                cList.SubItems(4) = Format(TPendidikan, "0.000")
                cList.SubItems(5) = Format(TPengalaman, "0.000")
End Sub



Private Sub TotNm()
    Dim cIndexTNm As Long
        For cIndexTNm = 1 To ListView4.ListItems.count
            TNmKarakter = TNmKarakter + ListView4.ListItems(cIndexTNm).SubItems(1)
            TNmKinerja = TNmKinerja + ListView4.ListItems(cIndexTNm).SubItems(2)
            TNmMasaKerja = TNmMasaKerja + ListView4.ListItems(cIndexTNm).SubItems(3)
            TNmPendidikan = TNmPendidikan + ListView4.ListItems(cIndexTNm).SubItems(4)
            TNmPengalaman = TNmPengalaman + ListView4.ListItems(cIndexTNm).SubItems(5)
        Next
        
         Set cList = ListView5.ListItems.Add(, , "Jumlah")
                cList.SubItems(1) = Format(TNmKarakter, "0.000")
                cList.SubItems(2) = Format(TNmKinerja, "0.000")
                cList.SubItems(3) = Format(TNmMasaKerja, "0.000")
                cList.SubItems(4) = Format(TNmPendidikan, "0.000")
                cList.SubItems(5) = Format(TNmPengalaman, "0.000")
End Sub

Private Sub TotColum()
    Dim a As Single
    Dim b As Single
    Dim c As Single
    Dim d As Single
    Dim e As Single
    Dim cttl As Single
            a = ListView4.ListItems(1).SubItems(1)
            b = ListView4.ListItems(1).SubItems(2)
            c = ListView4.ListItems(1).SubItems(3)
            d = ListView4.ListItems(1).SubItems(4)
            e = ListView4.ListItems(1).SubItems(5)
            
            cctl = a + b + c + d + e
            ListView6.ListItems.Add , , "Jumlah"
            ListView6.ListItems.Add , , Format(cctl, "0.000")
End Sub






