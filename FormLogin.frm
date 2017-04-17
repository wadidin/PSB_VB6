VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.1#0"; "Codejock.Controls.v17.1.0.ocx"
Begin VB.Form FormLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login Aplikasi"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5940
   Icon            =   "FormLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   81
   ScaleMode       =   0  'User
   ScaleWidth      =   297
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton PB_Login 
      Height          =   555
      Left            =   4200
      TabIndex        =   5
      Top             =   240
      Width           =   1575
      _Version        =   1114113
      _ExtentX        =   2778
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "   &Login"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormLogin.frx":25CA
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.GroupBox GB_Login 
      Height          =   1300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _Version        =   1114113
      _ExtentX        =   7011
      _ExtentY        =   2293
      _StockProps     =   79
      Caption         =   "&Login"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit FE_Username 
         Height          =   345
         Left            =   1920
         TabIndex        =   3
         Top             =   300
         Width           =   1815
         _Version        =   1114113
         _ExtentX        =   3201
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Text            =   "Username"
      End
      Begin XtremeSuiteControls.FlatEdit FE_Password 
         Height          =   350
         Left            =   1920
         TabIndex        =   4
         Top             =   750
         Width           =   1815
         _Version        =   1114113
         _ExtentX        =   3201
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Text            =   "Password"
         PasswordChar    =   "*"
      End
      Begin XtremeSuiteControls.Label L_Password 
         Height          =   240
         Left            =   720
         TabIndex        =   2
         Top             =   795
         Width           =   900
         _Version        =   1114113
         _ExtentX        =   1588
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&Password"
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
      Begin XtremeSuiteControls.Label L_Username 
         Height          =   240
         Left            =   720
         TabIndex        =   1
         Top             =   350
         Width           =   945
         _Version        =   1114113
         _ExtentX        =   1667
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "&Username"
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
      Begin VB.Image I_Password 
         Height          =   250
         Left            =   240
         Picture         =   "FormLogin.frx":26D8
         Stretch         =   -1  'True
         Top             =   795
         Width           =   250
      End
      Begin VB.Image I_Username 
         Height          =   250
         Left            =   240
         Picture         =   "FormLogin.frx":2895
         Stretch         =   -1  'True
         Top             =   350
         Width           =   250
      End
   End
   Begin XtremeSuiteControls.PushButton PB_Batal 
      Cancel          =   -1  'True
      Height          =   555
      Left            =   4200
      TabIndex        =   6
      Top             =   885
      Width           =   1575
      _Version        =   1114113
      _ExtentX        =   2778
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "    &Batal"
      UseVisualStyle  =   -1  'True
      Picture         =   "FormLogin.frx":2B30
      ImageAlignment  =   0
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Sub Isian()
    'Isian Kosong
    FE_Username.Text = ""
    FE_Password.Text = ""
End Sub

Private Sub Form_Load()
    'Memanggil Method Isian
    Call Isian
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'pilihan pesan untuk batal
    
    If (MsgBox("Kamu yakin mau keluar aplikasi ?", vbYesNo, "Konfirmasi") = vbYes) Then
        'pilihan 1
        End
    Else
        'pilihan 2
        Cancel = True
    End If
End Sub

Private Sub PB_Login_Click()
    'Username dan Password Jika Benar
    If FE_Username.Text = "wadin" And FE_Password.Text = "nidaw" Then
        MenuUtama.Enabled = True
        Me.Hide
        
    'Username Jika Salah
    ElseIf FE_Username.Text <> "wadin" And FE_Password.Text = "nidaw" Then
        MsgBox "Maaf, Username Kamu Salah !"
        FE_Username.Text = ""
        FE_Username.SetFocus
        
    'Password Jika Salah
    ElseIf FE_Username.Text = "wadin" And FE_Password.Text <> "nidaw" Then
        MsgBox "Maaf, Password Kamu Salah !"
        FE_Password.Text = ""
        FE_Password.SetFocus
        
    'Username dan Password Jika Salah
    ElseIf FE_Username.Text <> "wadin" And FE_Password.Text <> "nidaw" Then
        MsgBox "Maaf, Username Atau Password Kamu Salah !"
        FE_Username.Text = ""
        FE_Password.Text = ""
        FE_Username.SetFocus
    End If
End Sub


Private Sub PB_Batal_Click()
    'pilihan pesan untuk batal
    
    If (MsgBox("Kamu yakin mau keluar aplikasi ?", vbYesNo, "Konfirmasi") = vbYes) Then
        'pilihan 1
        End
    Else
        'pilihan 2
    End If
End Sub

