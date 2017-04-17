VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.1#0"; "Codejock.Controls.v17.1.0.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormCalon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form Data Calon Siswa Baru"
   ClientHeight    =   8505
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   13995
   Icon            =   "FormCalon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport LaporanCalon 
      Left            =   360
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   7080
      Top             =   8520
      Width           =   6735
      _ExtentX        =   11880
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   8520
      Width           =   6615
      _ExtentX        =   11668
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
      Caption         =   "tbl_calon"
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
   Begin XtremeSuiteControls.GroupBox GB_Judul 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      _Version        =   1114113
      _ExtentX        =   24156
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.Label Label_Judul 
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Top             =   360
         Width           =   5265
         _Version        =   1114113
         _ExtentX        =   9287
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "FORM DATA &CALON SISWA BARU"
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
   Begin XtremeSuiteControls.GroupBox GB_Isi 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   13695
      _Version        =   1114113
      _ExtentX        =   24156
      _ExtentY        =   5106
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin VB.ComboBox Combo_Sekolah 
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
         ItemData        =   "FormCalon.frx":25CA
         Left            =   2520
         List            =   "FormCalon.frx":25D1
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1320
         Width           =   3855
      End
      Begin VB.ComboBox Combo_Kelamin 
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
         ItemData        =   "FormCalon.frx":25E9
         Left            =   9120
         List            =   "FormCalon.frx":25F6
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox TB_Lahir 
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
         Left            =   9120
         MaxLength       =   17
         TabIndex        =   20
         Text            =   "Tempat Lahir"
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox TB_Alamat 
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
         Left            =   9120
         MaxLength       =   25
         TabIndex        =   18
         Text            =   "Alamat Pendaftar"
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox TB_Nama 
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
         Left            =   9120
         MaxLength       =   25
         TabIndex        =   17
         Text            =   "Nama Pendatar"
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox TB_NEM 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
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
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "NEM"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TB_Rayon 
         Enabled         =   0   'False
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
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   15
         Text            =   "Rayon"
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox TB_Nomor 
         Enabled         =   0   'False
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
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   13
         Text            =   "No. Pendaftaran"
         Top             =   360
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTPicker_Daftar 
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   95485953
         CurrentDate     =   42820
      End
      Begin MSComCtl2.DTPicker DTPicker_Lahir 
         Height          =   375
         Left            =   9120
         TabIndex        =   19
         Top             =   2280
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   95485953
         CurrentDate     =   42820
      End
      Begin XtremeSuiteControls.Label Label_DeskripsiNEM 
         Height          =   240
         Left            =   3960
         TabIndex        =   35
         Top             =   2325
         Width           =   2370
         _Version        =   1114113
         _ExtentX        =   4180
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Isi dengan &Angka dan Titik"
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
      Begin XtremeSuiteControls.Label Label_NEM 
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   2340
         Width           =   450
         _Version        =   1114113
         _ExtentX        =   794
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&NEM"
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
         TabIndex        =   11
         Top             =   1880
         Width           =   1335
         _Version        =   1114113
         _ExtentX        =   2355
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&Rayon/Daerah"
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
         TabIndex        =   10
         Top             =   1400
         Width           =   1200
         _Version        =   1114113
         _ExtentX        =   2117
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&Asal Sekolah"
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
      Begin XtremeSuiteControls.Label Label_TglLahir 
         Height          =   240
         Left            =   7320
         TabIndex        =   9
         Top             =   2355
         Width           =   1245
         _Version        =   1114113
         _ExtentX        =   2196
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&Tanggal Lahir"
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
      Begin XtremeSuiteControls.Label Label_TmptLahir 
         Height          =   240
         Left            =   7320
         TabIndex        =   8
         Top             =   1860
         Width           =   1185
         _Version        =   1114113
         _ExtentX        =   2090
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&Tempat Lahir"
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
      Begin XtremeSuiteControls.Label Label_Kelamin 
         Height          =   240
         Left            =   7320
         TabIndex        =   7
         Top             =   1395
         Width           =   1245
         _Version        =   1114113
         _ExtentX        =   2196
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&Jenis Kelamin"
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
      Begin XtremeSuiteControls.Label Label_Alamat 
         Height          =   240
         Left            =   7320
         TabIndex        =   6
         Top             =   900
         Width           =   1545
         _Version        =   1114113
         _ExtentX        =   2725
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&Alamat Pendaftar"
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
      Begin XtremeSuiteControls.Label Label_Nama 
         Height          =   240
         Left            =   7320
         TabIndex        =   5
         Top             =   405
         Width           =   1470
         _Version        =   1114113
         _ExtentX        =   2593
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&Nama Pendaftar"
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
      Begin XtremeSuiteControls.Label Label_Nomor 
         Height          =   240
         Left            =   360
         TabIndex        =   4
         Top             =   400
         Width           =   1455
         _Version        =   1114113
         _ExtentX        =   2566
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&No. Pendaftaran"
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
      Begin XtremeSuiteControls.Label Label_TglDaftar 
         Height          =   240
         Left            =   360
         TabIndex        =   3
         Top             =   900
         Width           =   1905
         _Version        =   1114113
         _ExtentX        =   3360
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&Tanggal Pendaftaran"
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
   Begin XtremeSuiteControls.GroupBox GB_Pencarian 
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   8895
      _Version        =   1114113
      _ExtentX        =   15690
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Pencarian &Data"
      UseVisualStyle  =   -1  'True
      Begin VB.ComboBox Combo_Pencarian 
         Height          =   315
         ItemData        =   "FormCalon.frx":2625
         Left            =   360
         List            =   "FormCalon.frx":263B
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   405
         Width           =   2775
      End
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
         Left            =   3360
         MaxLength       =   40
         TabIndex        =   22
         Text            =   "Pencarian"
         Top             =   360
         Width           =   3735
      End
      Begin XtremeSuiteControls.PushButton PB_Cari 
         Height          =   375
         Left            =   7440
         TabIndex        =   24
         Top             =   345
         Width           =   1095
         _Version        =   1114113
         _ExtentX        =   1931
         _ExtentY        =   670
         _StockProps     =   79
         Caption         =   "   &Cari"
         UseVisualStyle  =   -1  'True
         Picture         =   "FormCalon.frx":2694
      End
   End
   Begin MSDataGridLib.DataGrid DG_Calon 
      Bindings        =   "FormCalon.frx":27CF
      Height          =   2145
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   3784
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "no_daftar"
         Caption         =   "NO DAFTAR"
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
         DataField       =   "tanggal_daftar"
         Caption         =   "TGL DAFTAR"
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
      BeginProperty Column02 
         DataField       =   "nama"
         Caption         =   "NAMA PENDAFTAR"
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
      BeginProperty Column03 
         DataField       =   "alamat"
         Caption         =   "ALAMAT"
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
      BeginProperty Column04 
         DataField       =   "jenis_kelamin"
         Caption         =   "JENIS KELAMIN"
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
      BeginProperty Column05 
         DataField       =   "tempat_lahir"
         Caption         =   "TEMPAT LAHIR"
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
      BeginProperty Column06 
         DataField       =   "tanggal_lahir"
         Caption         =   "TGL LAHIR"
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
      BeginProperty Column07 
         DataField       =   "sekolah"
         Caption         =   "SEKOLAH ASAL"
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
      BeginProperty Column08 
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
      BeginProperty Column09 
         DataField       =   "NEM"
         Caption         =   "NEM"
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
            ColumnWidth     =   1500,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton PB_Tambah 
      Height          =   945
      Left            =   11040
      TabIndex        =   26
      Top             =   4280
      Width           =   1320
      _Version        =   1114113
      _ExtentX        =   2328
      _ExtentY        =   1667
      _StockProps     =   79
      Caption         =   " &Tambah"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormCalon.frx":27E4
   End
   Begin XtremeSuiteControls.PushButton PB_SU 
      Height          =   945
      Left            =   12480
      TabIndex        =   27
      Top             =   4280
      Width           =   1320
      _Version        =   1114113
      _ExtentX        =   2328
      _ExtentY        =   1667
      _StockProps     =   79
      Caption         =   " &Simpan"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormCalon.frx":28B3
   End
   Begin XtremeSuiteControls.PushButton PB_Edit 
      Height          =   945
      Left            =   11040
      TabIndex        =   28
      Top             =   5380
      Width           =   1320
      _Version        =   1114113
      _ExtentX        =   2328
      _ExtentY        =   1667
      _StockProps     =   79
      Caption         =   "      &Edit"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormCalon.frx":2A70
   End
   Begin XtremeSuiteControls.PushButton PB_Hapus 
      Height          =   945
      Left            =   12480
      TabIndex        =   29
      Top             =   5380
      Width           =   1320
      _Version        =   1114113
      _ExtentX        =   2328
      _ExtentY        =   1667
      _StockProps     =   79
      Caption         =   "   &Hapus"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormCalon.frx":2C1A
   End
   Begin XtremeSuiteControls.PushButton PB_Batal 
      Height          =   945
      Left            =   11040
      TabIndex        =   30
      Top             =   6500
      Width           =   2790
      _Version        =   1114113
      _ExtentX        =   4921
      _ExtentY        =   1667
      _StockProps     =   79
      Caption         =   "     &Batal"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormCalon.frx":2CE7
   End
   Begin XtremeSuiteControls.PushButton PB_Keluar 
      Height          =   825
      Left            =   120
      TabIndex        =   31
      Top             =   7560
      Width           =   13725
      _Version        =   1114113
      _ExtentX        =   24209
      _ExtentY        =   1455
      _StockProps     =   79
      Caption         =   "       &Keluar"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormCalon.frx":2D86
   End
   Begin XtremeSuiteControls.PushButton PB_Cetak 
      Height          =   900
      Left            =   9240
      TabIndex        =   32
      Top             =   4280
      Width           =   1580
      _Version        =   1114113
      _ExtentX        =   2787
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   " &Cetak"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormCalon.frx":2F28
   End
End
Attribute VB_Name = "FormCalon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Import Library Database
Dim Koneksi As ADODB.Connection
Dim RSCalon As ADODB.Recordset
Dim RSRayon As ADODB.Recordset

'Untuk Tombol Simpan dan Update
Private BTFungsi As Integer

Option Explicit

Private Sub BukaDB()
    'Untuk Koneksi adodb
    Set Koneksi = New ADODB.Connection
    
    'Untuk Memanggil database access yang sudah di buat
    Set RSCalon = New ADODB.Recordset
    Koneksi.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\db_siswabaru.mdb;Persist Security Info=False"
End Sub

Private Sub DataTable()
    'Memanggil Database dan Table ke Adodc
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\db_siswabaru.mdb;Persist Security Info=False"
    Adodc1.RecordSource = "tbl_calon"
    Adodc1.Refresh
    DG_Calon.Refresh
    Set DG_Calon.DataSource = Adodc1
    
    Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\db_siswabaru.mdb;Persist Security Info=False"
    Adodc2.RecordSource = "tbl_rayon"
    Adodc2.Refresh
    With Adodc2.Recordset
        Do Until .EOF
            Combo_Sekolah.AddItem !sekolah
            .MoveNext
        Loop
    End With
End Sub

Private Sub KodeOtomatis()
    Call BukaDB
    
    RSCalon.Open ("SELECT * FROM tbl_calon WHERE no_daftar IN(SELECT MAX(no_daftar)FROM tbl_calon) order by no_daftar DESC"), Koneksi
    RSCalon.Requery
    
    Dim Urutan As String * 10
    Dim Hitung As Long
    
    With RSCalon
        If .EOF Then
            Urutan = "ND-" + "0001"
            TB_Nomor.Text = Urutan
        Else
            Hitung = Right(!no_daftar, 4) + 1
            Urutan = "ND-" + Right("0000" & Hitung, 4)
        End If
        TB_Nomor.Text = Urutan
    End With
End Sub

Private Sub Combo_Sekolah_Click()
    Call BukaDB
    
    Set RSRayon = New ADODB.Recordset
    RSRayon.Open "SELECT * FROM tbl_rayon WHERE sekolah='" & Combo_Sekolah & "'", Koneksi
    RSRayon.Requery
    With RSRayon
        If .EOF And .BOF Then
            Exit Sub
        Else
            TB_Rayon.Text = !rayon
        End If
    End With
    RSRayon.Close
End Sub

Private Sub TB_Rayon_KeyPress(KeyAscii As Integer)
    'Huruf Kapital pada TextBox
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TB_NEM_KeyPress(KeyAscii As Integer)
    'Hanya Angka dan titik saja yang bisa di input textbox
    Select Case KeyAscii
    Case Is < 32
    Case 46
        If InStr(TB_NEM.Text, ".") <> 0 Then KeyAscii = 0
    Case 48 To 57
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub IsianAwal()
    'Isian Awal
    TB_Nomor.Text = "No. Pendaftaran"
    DTPicker_Daftar.Value = Now
    TB_Nama.Text = "Nama Pendaftar"
    TB_Alamat.Text = "Alamat Pendaftar"
    Combo_Kelamin.ListIndex = 0
    TB_Lahir.Text = "Tempat Lahir"
    DTPicker_Lahir.Value = Now
    Combo_Sekolah.ListIndex = 0
    TB_Rayon.Text = "Rayon"
    TB_NEM.Text = "NEM"
    TB_Cari.Text = ""
    Combo_Pencarian.ListIndex = 0
End Sub

Private Sub IsianBersih()
    'Isian Bersih
    DTPicker_Daftar.Value = Now
    TB_Nama.Text = ""
    TB_Alamat.Text = ""
    Combo_Kelamin.ListIndex = 0
    TB_Lahir.Text = ""
    DTPicker_Lahir.Value = Now
    Combo_Sekolah.ListIndex = 0
    TB_Rayon.Text = "Rayon"
    TB_NEM.Text = ""
End Sub

Private Sub IsianTidakBisa()
    'Isian Tidak Bisa
    DTPicker_Daftar.Enabled = False
    TB_Nama.Enabled = False
    TB_Alamat.Enabled = False
    Combo_Kelamin.Enabled = False
    TB_Lahir.Enabled = False
    DTPicker_Lahir.Enabled = False
    Combo_Sekolah.Enabled = False
    TB_NEM.Enabled = False
End Sub

Private Sub IsianBisa()
    'Isian Bisa
    DTPicker_Daftar.Enabled = True
    TB_Nama.Enabled = True
    TB_Alamat.Enabled = True
    Combo_Kelamin.Enabled = True
    TB_Lahir.Enabled = True
    DTPicker_Lahir.Enabled = True
    Combo_Sekolah.Enabled = True
    TB_NEM.Enabled = True
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
    
    'memanggil methor isianbersih, isianbisa agar textbox bisa diisi
    Call IsianBersih
    Call IsianBisa
    
    'Membuat Nomor Pendaftaran Otomatis
    Call KodeOtomatis
    
    'Ketika tombol Tambah di klik kursor langsung diarahkan ke Textbox Nama Sekolah
    TB_Nama.SetFocus
End Sub

Private Sub PB_SU_Click()
    If BTFungsi = 1 Then
        If TB_Nama.Text = "" Or TB_Alamat.Text = "" Or TB_NEM.Text = "" Or Combo_Sekolah.ListIndex = 0 Or Combo_Kelamin.ListIndex = 0 Then
            MsgBox "Harap di isi data yang masih kosong !", vbInformation + vbOKOnly, "Informasi"
            TB_Nama.SetFocus
        Else
            Adodc1.Recordset.AddNew
            Adodc1.Recordset!no_daftar = TB_Nomor.Text
            Adodc1.Recordset!tanggal_daftar = DTPicker_Daftar.Value
            Adodc1.Recordset!nama = TB_Nama.Text
            Adodc1.Recordset!alamat = TB_Alamat.Text
            Adodc1.Recordset!jenis_kelamin = Combo_Kelamin.Text
            Adodc1.Recordset!tempat_lahir = TB_Lahir.Text
            Adodc1.Recordset!tanggal_lahir = DTPicker_Lahir.Value
            Adodc1.Recordset!sekolah = Combo_Sekolah.Text
            Adodc1.Recordset!rayon = TB_Rayon.Text
            Adodc1.Recordset!NEM = TB_NEM.Text
            Adodc1.Recordset.Update
            
            MsgBox "Data berhasil di simpan...", vbInformation + vbOKOnly, "Informasi"
            DG_Calon.Refresh
            
            'Memanggil Method IsianAwal, IsianTidakBisa dan TombolAwal yang sudah di buat
            Call IsianAwal
            Call IsianTidakBisa
            Call TombolAwal
        End If
        
    ElseIf BTFungsi = 2 Then
        If TB_Nama.Text = "" Or TB_Alamat.Text = "" Or TB_NEM.Text = "" Or Combo_Sekolah.ListIndex = 0 Or Combo_Kelamin.ListIndex = 0 Then
            MsgBox "Harap di isi data yang masih kosong !", vbInformation + vbOKOnly, "Informasi"
            TB_Nama.SetFocus
        Else
            Adodc1.Recordset.Update
            Adodc1.Recordset!no_daftar = TB_Nomor.Text
            Adodc1.Recordset!tanggal_daftar = DTPicker_Daftar.Value
            Adodc1.Recordset!nama = TB_Nama.Text
            Adodc1.Recordset!alamat = TB_Alamat.Text
            Adodc1.Recordset!jenis_kelamin = Combo_Kelamin.Text
            Adodc1.Recordset!tempat_lahir = TB_Lahir.Text
            Adodc1.Recordset!tanggal_lahir = DTPicker_Lahir.Value
            Adodc1.Recordset!sekolah = Combo_Sekolah.Text
            Adodc1.Recordset!rayon = TB_Rayon.Text
            Adodc1.Recordset!NEM = TB_NEM.Text
            Adodc1.Recordset.Update
        
            MsgBox "Data berhasil di update...", vbInformation + vbOKOnly, "Informasi"
            DG_Calon.Refresh
        
            'Memanggil Method IsianAwal, IsianTidakBisa dan TombolAwal yang sudah di buat
            Call IsianAwal
            Call IsianTidakBisa
            Call TombolAwal
        End If
    End If
End Sub

Private Sub DG_Calon_Click()
    'Untuk Menampilkan data dari DataGrid ke TextBox
    If Adodc1.Recordset.RecordCount > 0 Then
        With Adodc1.Recordset
            TB_Nomor.Text = !no_daftar
            DTPicker_Daftar.Value = !tanggal_daftar
            TB_Nama.Text = !nama
            TB_Alamat.Text = !alamat
            Combo_Kelamin.Text = !jenis_kelamin
            TB_Lahir.Text = !tempat_lahir
            DTPicker_Lahir.Value = !tanggal_lahir
            Combo_Sekolah.Text = !sekolah
            TB_Rayon.Text = !rayon
            TB_NEM.Text = !NEM
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
    Edit = MsgBox("Kamu yakin mau Edit data ini ?", vbYesNo + vbQuestion, "Konfirmasi")
    
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
        Combo_Sekolah.SetFocus
        
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
    Hapus = MsgBox("Kamu yakin mau hapus data ini ?", vbYesNo + vbQuestion, "Konfirmasi")
    
    'pilihan pesan 1
    If Hapus = vbYes Then
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveFirst
        
        MsgBox "Data berhasil di hapus...", vbInformation + vbOKOnly, "Informasi"
        DG_Calon.Refresh
        
        'Memanggil Method IsianAwal, IsianTidakBisa dan TombolAwal yang sudah di buat
        Call IsianAwal
        Call IsianTidakBisa
        Call TombolAwal
    
    'pilihan pesan 2
    Else
        MsgBox "Data gagal di hapus !", vbInformation + vbOKOnly, "Informasi"
    End If
End Sub

Private Sub PB_Batal_Click()
    'Memanggil method isianawal, isiantidakbisa, dan tombolawal
    Call IsianAwal
    Call IsianTidakBisa
    Call TombolAwal
    
    'DataGrid di refresh
    DG_Calon.Refresh
End Sub

Private Sub PB_Cari_Click()
    If Combo_Pencarian.ListIndex = 0 Then
        MsgBox "Maaf, Silahkan pilih pencarian berdasarkan kriteria yang telah disiapkan !", vbInformation + vbOKOnly, "Informasi"
        TB_Cari.Text = ""
        Combo_Pencarian.SetFocus
    ElseIf Combo_Pencarian.ListIndex = 1 Then
        Adodc1.RecordSource = "SELECT * FROM tbl_calon WHERE no_daftar like '%" & TB_Cari.Text & "%'"
        Adodc1.Refresh
    ElseIf Combo_Pencarian.ListIndex = 2 Then
        Adodc1.RecordSource = "SELECT * FROM tbl_calon WHERE nama like '%" & TB_Cari.Text & "%'"
        Adodc1.Refresh
    ElseIf Combo_Pencarian.ListIndex = 3 Then
        Adodc1.RecordSource = "SELECT * FROM tbl_calon WHERE jenis_kelamin like '%" & TB_Cari.Text & "%'"
        Adodc1.Refresh
    ElseIf Combo_Pencarian.ListIndex = 4 Then
        Adodc1.RecordSource = "SELECT * FROM tbl_calon WHERE sekolah like '%" & TB_Cari.Text & "%'"
        Adodc1.Refresh
    ElseIf Combo_Pencarian.ListIndex = 5 Then
        Adodc1.RecordSource = "SELECT * FROM tbl_calon WHERE rayon like '%" & TB_Cari.Text & "%'"
        Adodc1.Refresh
    End If
End Sub

Private Sub PB_Cetak_Click()
    LaporanCalon.ReportFileName = App.Path + "\LapCalon.Rpt"
    LaporanCalon.DiscardSavedData = True
    LaporanCalon.Destination = crptToWindow
    LaporanCalon.WindowState = crptMaximized
    LaporanCalon.PrintReport
End Sub

Private Sub PB_Keluar_Click()
    Unload Me
End Sub

