VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmNoCorrido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nadie fue gallina ni se rindio"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9570
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   9375
      Begin VB.PictureBox Picture1 
         Height          =   3870
         Left            =   120
         ScaleHeight     =   3810
         ScaleWidth      =   3435
         TabIndex        =   2
         Top             =   345
         Width           =   3495
         Begin VB.Image Image1 
            Height          =   3825
            Left            =   0
            Picture         =   "FrmNoCorrido.frx":0000
            Top             =   0
            Width           =   3435
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   720
         OleObjectBlob   =   "FrmNoCorrido.frx":4925
         Top             =   195
      End
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   285
         Top             =   240
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   2460
         Left            =   3600
         OleObjectBlob   =   "FrmNoCorrido.frx":4B59
         TabIndex        =   1
         Top             =   1080
         Width           =   5715
      End
   End
End
Attribute VB_Name = "FrmNoCorrido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Me.Left = 15015 + 800

Set speaking2 = New SpeechLib.SpVoice
Set speaking2.Voice = speaking2.GetVoices().Item(voz)
    speaking2.Rate = -1
    speaking2.Volume = 100
    
    m = speaking2.Speak("No hay gallo corrido ni caido", SVSFlagsAsync)
End Sub


Private Sub Timer1_Timer()
Unload Me
End Sub
