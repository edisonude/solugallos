VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmBloqueo 
   Caption         =   "Seguridad"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   240
      Picture         =   "frmBloqueo.frx":0000
      ScaleHeight     =   2250
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   360
      Width           =   1515
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmBloqueo.frx":3941
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Desbloqueo de edición"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   5295
      Begin VB.CommandButton btnDesbloqueo 
         Caption         =   "Desbloquear"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   3
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox tPass 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmBloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDesbloqueo_Click()
If Me.tPass.Text = "9004215532" Then
    frmmenu.Show
    Unload Me
Else
    MsgBox "Contraseña de desbloqueo incorrecta", vbCritical
    Me.tPass = ""
    Me.tPass.SetFocus
End If
End Sub

Private Sub Form_Load()
Cooper_skin Me
End Sub

Private Sub tPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call btnDesbloqueo_Click
End If
End Sub
