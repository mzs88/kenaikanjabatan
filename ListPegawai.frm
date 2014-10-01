VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ListPegawai 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Pegawai"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   Icon            =   "ListPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4335
      Begin MSComctlLib.ListView LsPegawai 
         Height          =   4215
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   7435
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nik Pegawai"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nama Pegawai"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "List Pegawai"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4320
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ListPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    cNDb
    LoadPendidikan
End Sub

Private Sub LoadPendidikan()
    Dim cList As ListItem

    MySql = "SELECT nik_pegawai, nama FROM tb_pegawai ORDER BY nik_pegawai ASC"
    Set SdR = New ADODB.Recordset
        SdR.CursorLocation = adUseClient
        SdR.Open MySql, ConN, adOpenDynamic, adLockOptimistic

    LsPegawai.View = lvwReport
    LsPegawai.ListItems.Clear
        Do Until SdR.EOF
             Set cList = LsPegawai.ListItems.Add(, , SdR.Fields(0))
                cList.SubItems(1) = SdR.Fields(1)
            SdR.MoveNext
        Loop
End Sub

Private Sub LsPegawai_Click()
    Dim Ls As String
    Ls = Label6.Caption
    If Ls = "Pendidikan" Then
        Dim NikPendidikan As String
        NikPendidikan = LsPegawai.ListItems(LsPegawai.SelectedItem.Index)
        MySql = "SELECT nik_pegawai FROM tb_pendidikan WHERE nik_pegawai = '" & NikPendidikan & "'"
        Set SdR = New ADODB.Recordset
        With SdR
            .CursorLocation = adUseClient
            .Open MySql, ConN, adOpenDynamic, adLockOptimistic
            If Not .EOF Then
                MsgBox "Data Sudah Ada"
            Else
                Kriteria.TxtPendNik.Text = NikPendidikan
                Unload Me
            End If
        End With

    ElseIf Ls = "Pengalaman" Then
        Dim NikPengalaman As String
        NikPengalaman = LsPegawai.ListItems(LsPegawai.SelectedItem.Index)
        MySql = "SELECT nik_pegawai FROM tb_pengalaman WHERE nik_pegawai = '" & NikPengalaman & "'"
        Set SdR = New ADODB.Recordset
        With SdR
            .CursorLocation = adUseClient
            .Open MySql, ConN, adOpenDynamic, adLockOptimistic
            If Not .EOF Then
                MsgBox "Data Sudah Ada"
            Else
                Kriteria.TxtPengNik.Text = NikPengalaman
                Unload Me
            End If
        End With
    
    ElseIf Ls = "Karakter" Then
        Dim NikKarakter As String
        NikKarakter = LsPegawai.ListItems(LsPegawai.SelectedItem.Index)
        MySql = "SELECT nik_pegawai FROM tb_karakter WHERE nik_pegawai = '" & NikKarakter & "'"
        Set SdR = New ADODB.Recordset
        With SdR
            .CursorLocation = adUseClient
            .Open MySql, ConN, adOpenDynamic, adLockOptimistic
            If Not .EOF Then
                MsgBox "Data Sudah Ada"
            Else
                Kriteria.TxtKarNik.Text = NikKarakter
                Unload Me
            End If
        End With
    
    ElseIf Ls = "Kinerja" Then
        Dim NikKinerja As String
        NikKinerja = LsPegawai.ListItems(LsPegawai.SelectedItem.Index)
        MySql = "SELECT nik_pegawai FROM tb_kinerja WHERE nik_pegawai = '" & NikKinerja & "'"
        Set SdR = New ADODB.Recordset
        With SdR
            .CursorLocation = adUseClient
            .Open MySql, ConN, adOpenDynamic, adLockOptimistic
            If Not .EOF Then
                MsgBox "Data Sudah Ada"
            Else
                Kriteria.TxtKinNik.Text = NikKinerja
                Unload Me
            End If
        End With
        
    ElseIf Ls = "Masa Kerja" Then
        Dim NikMsKerja As String
        NikMsKerja = LsPegawai.ListItems(LsPegawai.SelectedItem.Index)
        MySql = "SELECT nik_pegawai FROM tb_masakerja WHERE nik_pegawai = '" & NikMsKerja & "'"
        Set SdR = New ADODB.Recordset
        With SdR
            .CursorLocation = adUseClient
            .Open MySql, ConN, adOpenDynamic, adLockOptimistic
            If Not .EOF Then
                MsgBox "Data Sudah Ada"
            Else
                Kriteria.TxtMsNik.Text = NikMsKerja
                Unload Me
            End If
        End With
    ElseIf Ls = "Pilih Operator" Then
        Dim NikOperator As String
        NikOperator = LsPegawai.ListItems(LsPegawai.SelectedItem.Index)
        MySql = "SELECT nik FROM login WHERE nik = '" & NikOperator & "'"
        Set SdR = New ADODB.Recordset
        With SdR
            .CursorLocation = adUseClient
            .Open MySql, ConN, adOpenDynamic, adLockOptimistic
            If Not .EOF Then
                MsgBox "Data Sudah Ada"
            Else
                With DataOperator
                    .TxtID.Text = NikOperator
                    .TxtNama.Text = LsPegawai.ListItems(LsPegawai.SelectedItem.Index).SubItems(1)
                End With
                Unload Me
            End If
        End With
    
    End If
End Sub
