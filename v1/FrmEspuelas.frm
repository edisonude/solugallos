VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmEspuelas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "No cambian espuelas"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8895
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   8655
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   720
         OleObjectBlob   =   "FrmEspuelas.frx":0000
         Top             =   195
      End
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   285
         Top             =   240
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   2940
         Left            =   3765
         OleObjectBlob   =   "FrmEspuelas.frx":0234
         TabIndex        =   2
         Top             =   525
         Width           =   4635
      End
      Begin VB.PictureBox Picture1 
         Height          =   3870
         Left            =   105
         ScaleHeight     =   3810
         ScaleWidth      =   3435
         TabIndex        =   1
         Top             =   225
         Width           =   3495
         Begin VB.Image Image1 
            Height          =   3825
            Left            =   0
            Picture         =   "FrmEspuelas.frx":02BC
            Top             =   0
            Width           =   3435
         End
      End
   End
End
Attribute VB_Name = "FrmEspuelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ss As Integer, mm As Integer, hh As Integer
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
    
    m = speaking2.Speak("No hay cambio de espuelas", SVSFlagsAsync)
End Sub


Private Sub Timer1_Timer()
Unload Me
End Sub
