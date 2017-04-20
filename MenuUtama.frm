VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MenuUtama 
   BackColor       =   &H8000000C&
   Caption         =   "APLIKASI Penerimaan Siswa Baru (PSB)"
   ClientHeight    =   9720
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   19755
   Icon            =   "MenuUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MenuUtama.frx":25CA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbars 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19755
      _ExtentX        =   34846
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "ImageList_TB"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Jeda1"
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Jeda2"
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tbSekolah"
            Object.ToolTipText     =   "Data Sekolah Asal"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tbCalon"
            Object.ToolTipText     =   "Data Calon Siswa"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tbSiswa"
            Object.ToolTipText     =   "Data Siswa Baru"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Jeda3"
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tbLapSekolah"
            Object.ToolTipText     =   "Laporan Sekolah Asal"
            Object.Tag             =   ""
            ImageIndex      =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tbLapCalon"
            Object.ToolTipText     =   "Laporan Calon Siswa"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tbLapSiswa"
            Object.ToolTipText     =   "Laporan Siswa Baru"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Jeda4"
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tbTentang"
            Object.ToolTipText     =   "Tentang Aplikasi"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Jeda5"
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBars 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   9345
      Width           =   19755
      _ExtentX        =   34846
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Text            =   "Jam"
            TextSave        =   "19/04/2017"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Tanggal"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "21.37"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Jam"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport LaporanRayon 
      Left            =   720
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport LaporanCalon 
      Left            =   1200
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport LaporanSiswa 
      Left            =   1680
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ImageList ImageList_TB 
      Left            =   1920
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuUtama.frx":9D4F2
            Key             =   "lgSekolah"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuUtama.frx":9DD0C
            Key             =   "lgCalon"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuUtama.frx":9E526
            Key             =   "lgSiswa"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuUtama.frx":9ED40
            Key             =   "lgLapSekolah"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuUtama.frx":9F55A
            Key             =   "lgLapCalon"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuUtama.frx":9FD74
            Key             =   "lgLapSiswa"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuUtama.frx":A058E
            Key             =   "lgTentang"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnLogOut 
         Caption         =   "Log &Out"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnJeda1 
         Caption         =   "-"
      End
      Begin VB.Menu mnKeluar 
         Caption         =   "&Keluar"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnData 
      Caption         =   "&Master Data"
      Begin VB.Menu mnSekolah 
         Caption         =   "Data &Sekolah Asal"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnJeda3 
         Caption         =   "-"
      End
      Begin VB.Menu mnCalon 
         Caption         =   "Data &Calon Siswa"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnJeda4 
         Caption         =   "-"
      End
      Begin VB.Menu mnSiswa 
         Caption         =   "Data &Siswa Baru"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnLaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnLapSekolah 
         Caption         =   "Laporan &Sekolah Asal"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnJeda5 
         Caption         =   "-"
      End
      Begin VB.Menu mnLapCalon 
         Caption         =   "Laporan &Calon Siswa"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnJeda6 
         Caption         =   "-"
      End
      Begin VB.Menu mnLapSiswa 
         Caption         =   "Laporan &Siswa Baru"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnBantuan 
      Caption         =   "&Bantuan"
      Begin VB.Menu mnTentang 
         Caption         =   "&Tentang"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "MenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub MDIForm_Unload(Cancel As Integer)
    'pilihan pesan untuk batal
    
    If (MsgBox("Kamu yakin mau keluar aplikasi ?", vbYesNo, "Konfirmasi") = vbYes) Then
        'pilihan 1
        End
    Else
        'pilihan 2
        Cancel = True
    End If
End Sub

Private Sub mnCalon_Click()
    'Form Calon tampil dengan new tab
    Dim FCalon As New FormCalon
    FCalon.Show
End Sub

Private Sub mnKeluar_Click()
    'pilihan pesan untuk batal
    
    If (MsgBox("Kamu yakin mau keluar aplikasi ?", vbYesNo, "Konfirmasi") = vbYes) Then
        'pilihan 1
        End
    Else
        'pilihan 2
    End If
End Sub

Private Sub mnLapCalon_Click()
    LaporanCalon.ReportFileName = App.Path + "\LapCalon.Rpt"
    LaporanCalon.DiscardSavedData = True
    LaporanCalon.Destination = crptToWindow
    LaporanCalon.WindowState = crptMaximized
    LaporanCalon.PrintReport
End Sub

Private Sub mnLapSekolah_Click()
    LaporanRayon.ReportFileName = App.Path + "\LapRayon.Rpt"
    LaporanRayon.DiscardSavedData = True
    LaporanRayon.Destination = crptToWindow
    LaporanRayon.WindowState = crptMaximized
    LaporanRayon.PrintReport
End Sub

Private Sub mnLapSiswa_Click()
    LaporanSiswa.ReportFileName = App.Path + "\LapSiswa.Rpt"
    LaporanSiswa.DiscardSavedData = True
    LaporanSiswa.Destination = crptToWindow
    LaporanSiswa.WindowState = crptMaximized
    LaporanSiswa.PrintReport
End Sub

Private Sub mnLogOut_Click()
    'MenuUTama tidak bisa digunakan
    MenuUtama.Enabled = False

    'Form Login tampil dengan new tab
    Dim FLogin As New FormLogin
    FLogin.Show
End Sub

Private Sub mnSekolah_Click()
    'Form Login tampil dengan new tab
    Dim FRayon As New FormRayon
    FRayon.Show
End Sub

Private Sub mnSiswa_Click()
    'Form Siswa tampil dengan new tab
    Dim FSiswa As New FormSiswa
    FSiswa.Show
End Sub

Private Sub mnTentang_Click()
    'MenuUTama tidak bisa digunakan
    MenuUtama.Enabled = False
    
    'Form Tentang tampil dengan new tab
    Dim FTentang As New FormTentang
    FTentang.Show
End Sub


Private Sub Toolbars_ButtonClick(ByVal Toolbars As Button)
    Select Case Toolbars.Key
        Case Is = "tbSekolah"
              'Form Login tampil dengan new tab
              Dim FRayon As New FormRayon
              FRayon.Show
        Case Is = "tbCalon"
              'Form Login tampil dengan new tab
              Dim FCalon As New FormCalon
              FCalon.Show
        Case Is = "tbSiswa"
              'Form Login tampil dengan new tab
              Dim FSiswa As New FormSiswa
              FSiswa.Show
        Case Is = "tbLapSekolah"
              LaporanRayon.ReportFileName = App.Path + "\LapRayon.Rpt"
              LaporanRayon.DiscardSavedData = True
              LaporanRayon.Destination = crptToWindow
              LaporanRayon.WindowState = crptMaximized
              LaporanRayon.PrintReport
        Case Is = "tbLapCalon"
              LaporanCalon.ReportFileName = App.Path + "\LapCalon.Rpt"
              LaporanRayon.DiscardSavedData = True
              LaporanCalon.Destination = crptToWindow
              LaporanCalon.WindowState = crptMaximized
              LaporanCalon.PrintReport
        Case Is = "tbLapSiswa"
              LaporanSiswa.ReportFileName = App.Path + "\LapSiswa.Rpt"
              LaporanRayon.DiscardSavedData = True
              LaporanSiswa.Destination = crptToWindow
              LaporanSiswa.WindowState = crptMaximized
              LaporanSiswa.PrintReport
        Case Is = "tbTentang"
              'MenuUTama tidak bisa digunakan
              MenuUtama.Enabled = False
              
              'Form Tentang tampil dengan new tab
              Dim FTentang As New FormTentang
              FTentang.Show
    End Select
End Sub
