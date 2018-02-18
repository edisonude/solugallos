VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmGano 
   Caption         =   "Ganador"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   13200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      Begin VB.PictureBox Picture1 
         Height          =   3870
         Left            =   165
         ScaleHeight     =   3810
         ScaleWidth      =   3435
         TabIndex        =   3
         Top             =   360
         Width           =   3495
         Begin VB.Image Image1 
            Height          =   3825
            Left            =   0
            Picture         =   "FrmGano.frx":0000
            Top             =   0
            Width           =   3435
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   1440
         Top             =   240
      End
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   285
         Top             =   240
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   720
         OleObjectBlob   =   "FrmGano.frx":4925
         Top             =   195
      End
      Begin ACTIVESKINLibCtl.SkinLabel titulo 
         Height          =   1020
         Left            =   5640
         OleObjectBlob   =   "FrmGano.frx":4B59
         TabIndex        =   1
         Top             =   120
         Width           =   5715
      End
      Begin ACTIVESKINLibCtl.SkinLabel TWin 
         Height          =   2820
         Left            =   3960
         OleObjectBlob   =   "FrmGano.frx":4BBD
         TabIndex        =   2
         Top             =   1440
         Width           =   8595
      End
   End
End
Attribute VB_Name = "FrmGano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hh, mm, ss As Integer

'Para voz
Dim WithEvents speaking As SpeechLib.SpVoice
Attribute speaking.VB_VarHelpID = -1
Dim WithEvents speaking2 As SpeechLib.SpVoice
Attribute speaking2.VB_VarHelpID = -1
Private speakingvoice As SpeechLib.ISpeechObjectToken
Dim m, n As Long
Dim RecdTime As Boolean


Private Sub Form_Load()
Uno_skin Me
Me.Left = 25015 + 600
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Unload Me
Unload FrmInfo
'FrmInfo.TValor = 0
FrmPantalla.Show
'FrmInfo.Show
End Sub

Private Sub Timer2_Timer()
Dim mensaje As String


If Me.TWin = "EMPATE" Then
    Me.titulo.Visible = True
    mensaje = "No hay ganador, pelea empatada"
Else
    mensaje = "La pelea fue ganada por " & Me.TWin.Caption
End If

Set speaking2 = New SpeechLib.SpVoice
Set speaking2.Voice = speaking2.GetVoices().Item(voz)
    speaking2.Rate = -1
    speaking2.Volume = 100
    
    m = speaking2.Speak(mensaje, SVSFlagsAsync)
    
    Me.Timer2.Enabled = False
End Sub

