VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form DataPegawai 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Data Pegawai"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   14310
   Icon            =   "DataPegawai.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   14310
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   14055
      Begin MSComCtl2.DTPicker DTPLahir 
         Height          =   315
         Left            =   8880
         TabIndex        =   22
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116523009
         CurrentDate     =   41739
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6120
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox CmbJabatan 
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
         ItemData        =   "DataPegawai.frx":0FA2
         Left            =   7080
         List            =   "DataPegawai.frx":0FC4
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   5295
      End
      Begin VB.ComboBox CmbAgama 
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
         ItemData        =   "DataPegawai.frx":1151
         Left            =   7080
         List            =   "DataPegawai.frx":1167
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   3255
      End
      Begin VB.CommandButton BtnBatal 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2640
         Picture         =   "DataPegawai.frx":11B1
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Batal"
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton BtnHapus 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1800
         Picture         =   "DataPegawai.frx":1BB3
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Hapus"
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton BtnUbah 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   960
         Picture         =   "DataPegawai.frx":25B5
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ubah"
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton BtnSimpan 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Picture         =   "DataPegawai.frx":2FB7
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Input"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox TxtID 
         Enabled         =   0   'False
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
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtNama 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox TxtAlamat 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox TxtTempat 
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
         Left            =   7080
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox CmbKelamin 
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
         ItemData        =   "DataPegawai.frx":39B9
         Left            =   7080
         List            =   "DataPegawai.frx":39C3
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   12480
         TabIndex        =   9
         Top             =   120
         Width           =   1455
         Begin VB.PictureBox foto 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   120
            ScaleHeight     =   1695
            ScaleWidth      =   1215
            TabIndex        =   21
            Top             =   240
            Width           =   1215
            Begin VB.Image Photo 
               Height          =   1695
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1215
            End
         End
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NIK"
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
         TabIndex        =   16
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nama"
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
         TabIndex        =   15
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alamat"
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
         TabIndex        =   14
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tempat/Tgl Lahir"
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
         Left            =   5160
         TabIndex        =   13
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Kelamin"
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
         Left            =   5160
         TabIndex        =   12
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Agama"
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
         Left            =   5160
         TabIndex        =   11
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Posisi Jabatan"
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
         Left            =   5160
         TabIndex        =   10
         Top             =   1320
         Width           =   1800
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5055
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   8916
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Data Pegawai"
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
      TabIndex        =   7
      Top             =   0
      Width           =   14040
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "DataPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFileName As String

Private Sub BtnHapus_Click()
    On Error Resume Next
    Dim FileName As String
    FileName = App.Path & "\Foto\" & ListView1.ListItems(ListView1.SelectedItem.Index) & ".jpg"
    pesan = MsgBox("Anda yakin akan menghapus data '" & ListView1.ListItems(ListView1.SelectedItem.Index) & " ' ", _
                    vbExclamation + vbYesNo + vbDefaultButton2)
    If pesan = vbYes Then
    MySql = "DELETE FROM tb_pegawai WHERE nik_pegawai = ?"
    Set CmD = New ADODB.Command
    With CmD
        .ActiveConnection = ConN
        .CommandType = adCmdText
        .CommandText = MySql
        .Parameters.Append .CreateParameter("p1", adChar, adParamInput, 129, ListView1.ListItems(ListView1.SelectedItem.Index))
        .Execute
    End With
    Kill FileName
    MsgBox ("Data Sudah Dihapus")
    LoadDatPegawai
    End If
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With Me.ListView1
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, _
        Me.ScaleHeight - .Top - 200
        'BuatTabel
    End With
    Kolom
    With Me.Label6
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2
        'Me.ScaleHeight -.Top - 200
    End With
    
    With Me.Frame1
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2
    End With
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Move 0, 0, MenuUtama.ScaleWidth, _
    MenuUtama.ScaleHeight
    cNDb
    Kolom
    LoadDatPegawai
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MenuUtama.Toolbar1.Buttons(4).Value = tbrUnpressed
End Sub

Private Sub Kolom()
    With Me.ListView1
        With .ColumnHeaders
            .Clear
            .Add , , "Nik", 1500
            .Add , , "Nama", 1400
            .Add , , "Tempat", 2500
            .Add , , "TTL", 2000
            .Add , , "Kelamin", 1500
            .Add , , "Alamat", 2000
            .Add , , "Agama, Tgl Lahir", 2800
            .Add , , "Jabatan", _
            IIf(Me.ListView1.Width - 12200 > 0, _
            Me.ListView1.Width - 12200, 4000)
        End With
    End With
End Sub


    'With ListView1
    '    .ColumnHeaders.Add , , "Nik"
    '    .ColumnHeaders.Add , , "Nama"
    '    .ColumnHeaders.Add , , "Tempat"
    '    .ColumnHeaders.Add , , "Tanggal Lahir"
     '   .ColumnHeaders.Add , , "Jenis Kelamin"
     '   .ColumnHeaders.Add , , "Alamat"
     '   .ColumnHeaders.Add , , "Agama"
     '   .ColumnHeaders.Add , , "Jabatan", 3700
    'End With

Private Sub LoadDatPegawai()
     Dim cList As ListItem

    MySql = "SELECT nik_pegawai, nama, tmpt_lahir, tgl_lahir, jenis_kelamin, alamat, agama, posisi_jabatan FROM tb_pegawai ORDER BY nik_pegawai ASC"
    Set SdR = New ADODB.Recordset
    With SdR
        .CursorLocation = adUseClient
        .Open MySql, ConN, adOpenDynamic, adLockOptimistic

        ListView1.View = lvwReport
        ListView1.ListItems.Clear
            Do Until .EOF
                Set cList = ListView1.ListItems.Add(, , .Fields(0))
                    cList.SubItems(1) = .Fields(1)
                    cList.SubItems(2) = .Fields(2)
                    cList.SubItems(3) = .Fields(3)
                    cList.SubItems(4) = .Fields(4)
                    cList.SubItems(5) = .Fields(5)
                    cList.SubItems(6) = .Fields(6)
                    cList.SubItems(7) = .Fields(7)
                    .MoveNext
            Loop
    End With
End Sub

Private Sub cNikPegawai()
Dim cKode As String * 8
Dim cHitung As Long
     MySql = "SELECT nik_pegawai FROM tb_pegawai Where nik_pegawai In(Select Max(nik_pegawai)From tb_pegawai)Order By nik_pegawai Desc"
     Set SdR = New ADODB.Recordset
     With SdR
        .CursorLocation = adUseClient
        .Open MySql, ConN, adOpenDynamic, adLockOptimistic
        .Requery
    
            If .EOF Then
                cKode = "0001"
                TxtNik = cKode
            Else
                cHitung = Str(!nik_pegawai) + 1
                cKode = Right("0000" & cHitung, 8)
            End If
        TxtID.Text = cKode
        .Close
    End With
End Sub

Private Sub BtnSimpan_Click()

    If BtnSimpan.ToolTipText = "Input" Then
        BtnSimpan.ToolTipText = "Simpan"
        BtnSimpan.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        cNikPegawai
    Else
        If TxtNama.Text = "" Or TxtAlamat.Text = "" Or TxtTempat.Text = "" _
                                Or Photo.Picture = Empty Then
           pesan = MsgBox("Data kurang lengkap", vbCritical)
        Else
        MySql = "INSERT INTO tb_pegawai (nik_pegawai, nama, alamat, tmpt_lahir, tgl_lahir, jenis_kelamin, agama, posisi_jabatan, foto) VALUES (" & _
            "?, " & _
            "?," & _
            "?, " & _
            "?, " & _
            "?, " & _
            "?, " & _
            "?, " & _
            "?, " & _
            "?)"
       Set CmD = New ADODB.Command
       With CmD
            .ActiveConnection = ConN
            .CommandType = adCmdText
            .CommandText = MySql
            .Parameters.Append .CreateParameter("p1", adChar, adParamInput, 129, Me.TxtID.Text)
            .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, 200, Me.TxtNama.Text)
            .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, 200, Me.TxtAlamat.Text)
            .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, 200, Me.TxtTempat.Text)
            .Parameters.Append .CreateParameter("p5", adDate, adParamInput, 7, Me.DTPLahir.Value)
            .Parameters.Append .CreateParameter("p6", adVarChar, adParamInput, 200, Me.CmbKelamin.Text)
            .Parameters.Append .CreateParameter("p7", adVarChar, adParamInput, 200, Me.CmbAgama.Text)
            .Parameters.Append .CreateParameter("p8", adVarChar, adParamInput, 200, Me.CmbJabatan.Text)
            .Parameters.Append .CreateParameter("p9", adVarChar, adParamInput, 100, TxtID.Text)
            
            .Execute
       End With
       BtnSimpan.ToolTipText = "Input"
       BtnSimpan.Picture = LoadPicture(App.Path & "\Button\Create.ico")
       SimpanFoto
       LoadDatPegawai
       End If
    End If
        
End Sub

Private Sub SimpanFoto()
    SavePicture Photo.Picture, App.Path & "\Foto\" & TxtID.Text & ".jpg"
End Sub

Private Sub Baca()
        On Error Resume Next
        TxtID.Text = ListView1.ListItems(ListView1.SelectedItem.Index)
        TxtNama.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
        TxtAlamat.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(5)
        TxtTempat.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
        CmbKelamin.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(4)
        CmbAgama.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(6)
        CmbJabatan.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(7)
        Photo.Picture = LoadPicture(App.Path & "\Foto\" & TxtID.Text & ".jpg")
End Sub

Private Sub BtnUbah_Click()
    If BtnUbah.ToolTipText = "Ubah" Then
        BtnUbah.ToolTipText = "Update"
        BtnUbah.Picture = LoadPicture(App.Path & "\Button\Apply.ico")
        Baca
    Else
        SimpanFoto
        MySql = "UPDATE tb_pegawai SET nama = " & _
        "'" & TxtNama.Text & "', tmpt_lahir = " & _
        "'" & TxtTempat.Text & "', tgl_lahir = " & _
        "'" & Format(DTPLahir.Value, "yyyy-MM-dd") & "', jenis_kelamin = " & _
        "'" & CmbKelamin.Text & "', alamat = " & _
        "'" & TxtAlamat.Text & "', agama = " & _
        "'" & CmbAgama.Text & "', posisi_jabatan = " & _
        "'" & CmbJabatan.Text & "' WHERE nik_pegawai = " & _
        "'" & TxtID.Text & "'"
        ConN.Execute MySql
        MsgBox ("Data Berhasil Dirubah")
        LoadDatPegawai
        BtnUbah.ToolTipText = "Ubah"
        BtnUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
    End If
End Sub

Private Sub ListView1_Click()
    Batal
End Sub

Private Sub Photo_Click()
    CommonDialog1.Filter = "JPEG (*.jpg)|*.jpg|"
    CommonDialog1.ShowOpen
    Photo.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub BtnBatal_Click()
    Batal
End Sub

Private Sub Batal()
    TxtID.Text = ""
    TxtNama.Text = ""
    TxtAlamat.Text = ""
    TxtTempat.Text = ""
    BtnUbah.ToolTipText = "Ubah"
    BtnSimpan.ToolTipText = "Input"
    Photo.Picture = Nothing
    BtnSimpan.Picture = LoadPicture(App.Path & "\Button\Create.ico")
    BtnUbah.Picture = LoadPicture(App.Path & "\Button\Modify.ico")
End Sub
