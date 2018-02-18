VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frmerrores 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ERRORES Y ADVERTENCIAS"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   10065
      TabIndex        =   6
      Top             =   150
      Width           =   10065
   End
   Begin VB.Frame frmerrores 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   300
      TabIndex        =   2
      Top             =   255
      Width           =   9495
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2070
         Left            =   240
         ScaleHeight     =   2070
         ScaleWidth      =   2115
         TabIndex        =   0
         Top             =   345
         Width           =   2115
         Begin VB.Image Imgerror 
            Height          =   2100
            Left            =   -15
            Picture         =   "Frmerrores.frx":0000
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00808080&
         Height          =   2460
         Left            =   2520
         TabIndex        =   5
         ToolTipText     =   "Haga click para corregir"
         Top             =   225
         Width           =   6825
         Begin ACTIVESKINLibCtl.SkinLabel lberror 
            Height          =   1860
            Left            =   255
            OleObjectBlob   =   "Frmerrores.frx":EEB2
            TabIndex        =   1
            Top             =   330
            Width           =   6345
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00F8F4EE&
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   -60
         ScaleHeight     =   165
         ScaleWidth      =   210
         TabIndex        =   4
         Top             =   -75
         Width           =   210
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00F8F4EE&
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   1815
         ScaleHeight     =   165
         ScaleWidth      =   7725
         TabIndex        =   3
         Top             =   -75
         Width           =   7725
      End
   End
   Begin VB.Timer Temporizador 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4080
      Top             =   2595
   End
   Begin VB.Timer Tiempo 
      Interval        =   2000
      Left            =   4485
      Top             =   2595
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Frmerrores.frx":EF10
      Top             =   0
   End
End
Attribute VB_Name = "Frmerrores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valor As Integer
Private Sub Form_Click()
Temporizador.Enabled = True
End Sub

Private Sub Form_Load()
Uno_skin Me
valor = 225
Call Aplicar_Transparencia(Me.hwnd, CByte(valor))

Me.Left = 15015 + 800
End Sub

Private Sub Frame2_Click()
Temporizador.Enabled = True
End Sub

Private Sub Image1_Click()
Temporizador.Enabled = True
End Sub

Private Sub Frame1_Click()
Temporizador.Enabled = True
End Sub
Private Sub frmerrores_Click()
Temporizador.Enabled = True
End Sub
Private Sub Imgerror_Click()
Temporizador.Enabled = True
End Sub

Private Sub Temporizador_Timer()
valor = valor - 20
If valor <= 0 Then
    Unload Me
    Exit Sub
End If
Call Aplicar_Transparencia(Me.hwnd, CByte(valor))
End Sub

Private Sub Tiempo_Timer()
If Tiempo.Interval = 2000 Then
    Temporizador.Enabled = True
    Tiempo.Enabled = False
End If
End Sub
