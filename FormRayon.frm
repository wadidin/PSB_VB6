VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.1#0"; "Codejock.Controls.v17.1.0.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormRayon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form Data Sekolah Asal"
   ClientHeight    =   6840
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8400
   Icon            =   "FormRayon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport LaporanRayon 
      Left            =   360
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   6960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "tbl_rayon"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin XtremeSuiteControls.GroupBox GB_Pencarian 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   6615
      _Version        =   1114113
      _ExtentX        =   11668
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Pencarian &Data"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox TB_Cari 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         MaxLength       =   40
         TabIndex        =   18
         Text            =   "Pencarian"
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox Combo_Pencarian 
         Height          =   315
         ItemData        =   "FormRayon.frx":25CA
         Left            =   360
         List            =   "FormRayon.frx":25D7
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   405
         Width           =   1575
      End
      Begin XtremeSuiteControls.PushButton PB_Cari 
         Height          =   380
         Left            =   5160
         TabIndex        =   8
         Top             =   340
         Width           =   1095
         _Version        =   1114113
         _ExtentX        =   1931
         _ExtentY        =   670
         _StockProps     =   79
         Caption         =   "   &Cari"
         UseVisualStyle  =   -1  'True
         Picture         =   "FormRayon.frx":25FB
      End
   End
   Begin XtremeSuiteControls.GroupBox GB_Isi 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5295
      _Version        =   1114113
      _ExtentX        =   9340
      _ExtentY        =   2566
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox TB_Rayon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   17
         Text            =   "Rayon"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TB_Sekolah 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         MaxLength       =   40
         TabIndex        =   16
         Text            =   "Nama Sekolah"
         Top             =   360
         Width           =   2895
      End
      Begin XtremeSuiteControls.Label Label_DeskripsiRayon 
         Height          =   240
         Left            =   3120
         TabIndex        =   5
         Top             =   900
         Width           =   1815
         _Version        =   1114113
         _ExtentX        =   3201
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "3 &Huru Kota Sekolah"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label_Rayon 
         Height          =   240
         Left            =   360
         TabIndex        =   4
         Top             =   900
         Width           =   1335
         _Version        =   1114113
         _ExtentX        =   2355
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Rayon/Daerah"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label_Sekolah 
         Height          =   240
         Left            =   360
         TabIndex        =   3
         Top             =   400
         Width           =   1350
         _Version        =   1114113
         _ExtentX        =   2381
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&Nama Sekolah"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GB_Judul 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _Version        =   1114113
      _ExtentX        =   14420
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.Label Label_Judul 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   4500
         _Version        =   1114113
         _ExtentX        =   7938
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "FORM DATA &SEKOLAH ASAL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton PB_Tambah 
      Height          =   825
      Left            =   6960
      TabIndex        =   9
      Top             =   1150
      Width           =   1320
      _Version        =   1114113
      _ExtentX        =   2328
      _ExtentY        =   1455
      _StockProps     =   79
      Caption         =   " &Tambah"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormRayon.frx":2736
   End
   Begin XtremeSuiteControls.PushButton PB_SU 
      Height          =   825
      Left            =   6960
      TabIndex        =   10
      Top             =   2160
      Width           =   1320
      _Version        =   1114113
      _ExtentX        =   2328
      _ExtentY        =   1455
      _StockProps     =   79
      Caption         =   " &Simpan"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormRayon.frx":2805
   End
   Begin XtremeSuiteControls.PushButton PB_Edit 
      Height          =   825
      Left            =   6960
      TabIndex        =   11
      Top             =   3120
      Width           =   1320
      _Version        =   1114113
      _ExtentX        =   2328
      _ExtentY        =   1455
      _StockProps     =   79
      Caption         =   "      &Edit"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormRayon.frx":29C2
   End
   Begin XtremeSuiteControls.PushButton PB_Hapus 
      Height          =   825
      Left            =   6960
      TabIndex        =   12
      Top             =   4080
      Width           =   1320
      _Version        =   1114113
      _ExtentX        =   2328
      _ExtentY        =   1455
      _StockProps     =   79
      Caption         =   "   &Hapus"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormRayon.frx":2B6C
   End
   Begin XtremeSuiteControls.PushButton PB_Batal 
      Height          =   825
      Left            =   6960
      TabIndex        =   13
      Top             =   5060
      Width           =   1320
      _Version        =   1114113
      _ExtentX        =   2328
      _ExtentY        =   1455
      _StockProps     =   79
      Caption         =   "     &Batal"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormRayon.frx":2C39
   End
   Begin XtremeSuiteControls.PushButton PB_Keluar 
      Height          =   705
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Width           =   8160
      _Version        =   1114113
      _ExtentX        =   14393
      _ExtentY        =   1244
      _StockProps     =   79
      Caption         =   "       &Keluar"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormRayon.frx":2CD8
   End
   Begin XtremeSuiteControls.PushButton PB_Cetak 
      Height          =   1380
      Left            =   5520
      TabIndex        =   15
      Top             =   1155
      Width           =   1200
      _Version        =   1114113
      _ExtentX        =   2117
      _ExtentY        =   2434
      _StockProps     =   79
      Caption         =   " &Cetak"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormRayon.frx":2E7A
   End
   Begin MSDataGridLib.DataGrid DG_Rayon 
      Height          =   2020
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3572
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "sekolah"
         Caption         =   "NAMA SEKOLAH"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "rayon"
         Caption         =   "RAYON/DAERAH"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            DividerStyle    =   4
            ColumnWidth     =   4004,788
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2505,26
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormRayon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Import Library Database
Dim Koneksi As ADODB.Connection
Dim RSRayon As ADODB.Recordset

'Untuk Tombol Simpan dan Update
Private BTFungsi As Integer

Option Explicit

Private Sub BukaDB()
    'Untuk Koneksi adodb
    Set Koneksi = New ADODB.Connection
    
    'Untuk Memanggil database access yang sudah di buat
    Set RSRayon = New ADODB.Recordset
    Koneksi.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\db_siswabaru.mdb;Persist Security Info=False"
End Sub

Private Sub DataTable()
    'Memanggil Database dan Table ke Adodc
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\db_siswabaru.mdb;Persist Security Info=False"
    Adodc1.RecordSource = "tbl_rayon"
    Adodc1.Refresh
    DG_Rayon.Refresh
    Set DG_Rayon.DataSource = Adodc1
End Sub

Private Sub TB_Rayon_KeyPress(KeyAscii As Integer)
    'Huruf Kapital pada TextBox
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub IsianAwal()
    'Isian Bersih
    TB_Sekolah.Text = "Nama Sekolah"
    TB_Rayon.Text = "Rayon"
    TB_Cari.Text = ""
    Combo_Pencarian.ListIndex = 0
End Sub

Private Sub IsianBersih()
    'Isian Bersih
    TB_Sekolah.Text = ""
    TB_Rayon.Text = ""
End Sub

Private Sub IsianTidakBisa()
    'Isian Tidak Bisa
    TB_Sekolah.Enabled = False
    TB_Rayon.Enabled = False
End Sub

Private Sub IsianBisa()
    'Isian Bisa
    TB_Sekolah.Enabled = True
    TB_Rayon.Enabled = True
End Sub

Private Sub TombolAwal()
    'Tombol Awal
    PB_Tambah.Enabled = True
    PB_SU.Caption = "&Simpan"
    PB_SU.Enabled = False
    PB_Edit.Enabled = False
    PB_Hapus.Enabled = False
    PB_Batal.Enabled = True
End Sub

Private Sub Form_Load()
    'StartupPosition
    Me.Left = (MenuUtama.ScaleWidth - Me.Width) / 2
    Me.Top = (MenuUtama.ScaleHeight - Me.Height) / 2
    
    'Memanggil Method Koneksi Database yang sudah di buat
    Call BukaDB
    
    'Memanggil Method DataTable yang sudah di buat
    Call DataTable
    
    'Memanggil Method yang telah dibuat
    Call IsianAwal
    Call IsianTidakBisa
    Call TombolAwal
End Sub

Private Sub PB_Tambah_Click()
    'Untuk Funsi Tombol simpan 1
    BTFungsi = 1
    'Tombol Berganti nama jadi simpan
    PB_SU.Caption = "&Simpan"
    'Tombol Simpan bisa digunakan
    PB_SU.Enabled = True
    'Tombol Tambah tidak bisa digunakan
    PB_Tambah.Enabled = False
    
    'memanggil method isianbersih, isianbisa agar textbox bisa diisi
    Call IsianBersih
    Call IsianBisa
    'Ketika tombol Tambah di klik kursor langsung diarahkan ke Textbox Nama Sekolah
    TB_Sekolah.SetFocus
End Sub

Private Sub PB_SU_Click()
    If BTFungsi = 1 Then
        Adodc1.Recordset.Find "sekolah = '" & TB_Sekolah.Text & "'"
        
        If Not Adodc1.Recordset.EOF Then
            MsgBox "Maaf, Nama Sekolah Sudah Ada !", vbOKOnly, "Informasi"
            TB_Sekolah.Text = ""
            TB_Sekolah.SetFocus
        ElseIf TB_Sekolah.Text = "" Or TB_Rayon.Text = "" Then
            MsgBox "Harap di isi data yang masih kosong !", vbOKOnly, "Informasi"
            TB_Sekolah.SetFocus
        Else
            Adodc1.Recordset.AddNew
            Adodc1.Recordset!sekolah = TB_Sekolah.Text
            Adodc1.Recordset!rayon = TB_Rayon.Text
            Adodc1.Recordset.Update
            
            MsgBox "Data berhasil di simpan...", vbOKOnly, "Informasi"
            DG_Rayon.Refresh
            
            'Memanggil Method IsianAwal, IsianTidakBisa dan TombolAwal yang sudah di buat
            Call IsianAwal
            Call IsianTidakBisa
            Call TombolAwal
        End If
        
    ElseIf BTFungsi = 2 Then
        If TB_Sekolah.Text = "" Or TB_Rayon.Text = "" Then
            MsgBox "Harap di isi data yang masih kosong !", vbOKOnly, "Informasi"
            TB_Sekolah.SetFocus
        Else
            Adodc1.Recordset.Update
            Adodc1.Recordset!sekolah = TB_Sekolah.Text
            Adodc1.Recordset!rayon = TB_Rayon.Text
            Adodc1.Recordset.Update
        
            MsgBox "Data berhasil di update...", vbOKOnly, "Informasi"
            DG_Rayon.Refresh
            
            'Memanggil Method IsianAwal, IsianTidakBisa dan TombolAwal yang sudah di buat
            Call IsianAwal
            Call IsianTidakBisa
            Call TombolAwal
        End If
    End If
End Sub

Private Sub DG_Rayon_Click()
    'Untuk Menampilkan data dari DataGrid ke TextBox
    If Adodc1.Recordset.RecordCount > 0 Then
        With Adodc1.Recordset
            TB_Sekolah.Text = !sekolah
            TB_Rayon.Text = !rayon
        End With
    End If
    
    'Untuk Tombol yang digunakan saat data di klik
    PB_Edit.Enabled = True
    PB_Hapus.Enabled = True
    PB_Tambah.Enabled = False
End Sub

Private Sub PB_Edit_Click()
    'pilihan untung pesan
    Dim Edit As String
    Edit = MsgBox("Kamu yakin mau Edit data ini ?", vbYesNo, "Konfirmasi")
    
    'pilihan pesan 1
    If Edit = vbYes Then
        'Untuk Funsi Tombol simpan 2
        BTFungsi = 2
        'Tombol Berganti nama jadi simpan
        PB_SU.Caption = "&Update"
        'Tombol Simpan bisa digunakan
        PB_SU.Enabled = True
        'Tombol yang tidak bisa digunakan
        PB_Hapus.Enabled = False
        PB_Tambah.Enabled = False
        PB_Edit.Enabled = False
        
        'Memanggil Method Isianbisa yang sudah di buat
        Call IsianBisa
        
        'Untuk focus ke textbox sekolah
        TB_Sekolah.SetFocus
        
    'pilihan pesan 2
    Else
        'Memanggil Method IsianAwal, IsianTidakBisa dan TombolAwal yang sudah di buat
        Call IsianAwal
        Call IsianTidakBisa
        Call TombolAwal
    End If
End Sub

Private Sub PB_Hapus_Click()
    'pilihan untuk pesan
    Dim Hapus As String
    Hapus = MsgBox("Kamu yakin mau hapus data ini ?", vbYesNo, "Konfirmasi")
    
    'pilihan pesan 1
    If Hapus = vbYes Then
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveFirst
        
        MsgBox "Data berhasil di hapus...", vbOKOnly, "Informasi"
        DG_Rayon.Refresh
        
        'Memanggil Method, IsianAwal IsianTidakBisa dan TombolAwal yang sudah di buat
        Call IsianAwal
        Call IsianTidakBisa
        Call TombolAwal
    
    'pilihan pesan 2
    Else
        MsgBox "Data gagal di hapus !", vbOKOnly, "Informasi"
    End If
End Sub

Private Sub PB_Batal_Click()
    'Memanggil method isianawal, isiantidakbisa, dan tombolawal
    Call IsianAwal
    Call IsianTidakBisa
    Call TombolAwal
    
    'DataGrid di refresh
    DG_Rayon.Refresh
End Sub

Private Sub PB_Cari_Click()
    If Combo_Pencarian.ListIndex = 0 Then
        MsgBox "Maaf, Silahkan pilih pencarian berdasarkan kriteria yang disiapkan !", vbOKOnly, "Informasi"
        TB_Cari.Text = ""
        Combo_Pencarian.SetFocus
    ElseIf Combo_Pencarian.ListIndex = 1 Then
        Adodc1.RecordSource = "SELECT * FROM tbl_rayon WHERE sekolah like '%" & TB_Cari.Text & "%'"
        Adodc1.Refresh
    ElseIf Combo_Pencarian.ListIndex = 2 Then
        Adodc1.RecordSource = "SELECT * FROM tbl_rayon WHERE rayon like '%" & TB_Cari.Text & "%'"
        Adodc1.Refresh
    End If
End Sub

Private Sub PB_Cetak_Click()
    LaporanRayon.ReportFileName = App.Path + "\LapRayon.Rpt"
    LaporanRayon.DiscardSavedData = True
    LaporanRayon.Destination = crptToWindow
    LaporanRayon.WindowState = crptMaximized
    LaporanRayon.PrintReport
End Sub

Private Sub PB_Keluar_Click()
    Unload Me
End Sub
