VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.1#0"; "Codejock.Controls.v17.1.0.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#17.1#0"; "Codejock.SkinFramework.v17.1.0.ocx"
Begin VB.Form FormLoading 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3705
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FormLoading.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_Loading 
      Interval        =   100
      Left            =   720
      Top             =   600
   End
   Begin XtremeSuiteControls.ProgressBar L_ProgressBar 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2520
      Width           =   4935
      _Version        =   1114113
      _ExtentX        =   8705
      _ExtentY        =   661
      _StockProps     =   93
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   240
      Top             =   600
      _Version        =   1114113
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.Label Label_Loading 
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
      _Version        =   1114113
      _ExtentX        =   1931
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Loading ..."
      ForeColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label_Persen 
      Height          =   300
      Left            =   2280
      TabIndex        =   1
      Top             =   3120
      Width           =   540
      _Version        =   1114113
      _ExtentX        =   953
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "00 %"
      ForeColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      AutoSize        =   -1  'True
   End
   Begin VB.Image Image_Loading 
      Height          =   3720
      Left            =   0
      Picture         =   "FormLoading.frx":000C
      Top             =   0
      Width           =   6585
   End
End
Attribute VB_Name = "FormLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Xtreme Framework
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Option Explicit

Public Sub TemaLoading()
    'Tema Form
    SkinFramework1.LoadSkin App.Path + "\styles\Codejock.cjstyles", ""
    SkinFramework1.ApplyWindow Me.hWnd
End Sub

Private Sub Form_Load()
    'Memanggil Method Tema
    Call TemaLoading
End Sub

Private Sub Timer_Loading_Timer()
    'ProgressBar Loading +2
    Me.L_ProgressBar.Value = Me.L_ProgressBar.Value + 2
    Me.Label_Persen.Caption = Me.L_ProgressBar.Value & "%"
    
    'ProgressBar Jika Penuh
    If Me.L_ProgressBar.Value = 102 Then
        'FormLoading tidak akan tampil
        Me.Hide
        
        'MenuUtama akan tampil
        MenuUtama.Show
        'MenuUTama tidak bisa digunakan
        MenuUtama.Enabled = False
        
        'Form Login tampil dengan new tab
        Dim FLogin As New FormLogin
        FLogin.Show
    End If
End Sub

