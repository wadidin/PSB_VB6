VERSION 5.00
Begin VB.Form FormTentang 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tentang Aplikasi"
   ClientHeight    =   3150
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "FormTentang.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2174.186
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   780
      Left            =   240
      Picture         =   "FormTentang.frx":25CA
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   1
      Top             =   240
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "wApps Software License"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2160
      LinkItem        =   "http://www.wadidin.copaster.com/"
      TabIndex        =   6
      Top             =   2683
      Width           =   1755
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   $"FormTentang.frx":4B94
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   1290
      TabIndex        =   2
      Top             =   1245
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "APLIKASI Penerimaan Siswa Baru (PSB)"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1290
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versi BETA 1"
      Height          =   225
      Left            =   1290
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      AutoSize        =   -1  'True
      Caption         =   "Aplikasi ini berlisensi dari"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   255
      TabIndex        =   3
      Top             =   2683
      Width           =   1725
   End
End
Attribute VB_Name = "FormTentang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  'Form Tentang di sembunyikan
  Unload Me
  'MenuUtama akan tampil
  MenuUtama.Show
  'MenuUTama tidak bisa digunakan
  MenuUtama.Enabled = True
End Sub

Private Sub Form_Load()
    Me.Caption = "Tentang Aplikasi " & App.Title
    lblVersion.Caption = "Versi " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'MenuUTama tidak bisa digunakan
    MenuUtama.Enabled = True
End Sub
