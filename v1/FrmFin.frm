VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmFin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Esto se acabo"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8865
   Icon            =   "FrmFin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.PictureBox Picture1 
         Height          =   3870
         Left            =   75
         ScaleHeight     =   3810
         ScaleWidth      =   3435
         TabIndex        =   2
         Top             =   90
         Width           =   3495
         Begin VB.Image Image1 
            Height          =   3825
            Left            =   0
            Picture         =   "FrmFin.frx":08CA
            Top             =   0
            Width           =   3435
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   5040
         Top             =   240
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   6840
         OleObjectBlob   =   "FrmFin.frx":51EF
         Top             =   120
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   1980
         Left            =   3720
         OleObjectBlob   =   "FrmFin.frx":5423
         TabIndex        =   1
         Top             =   1080
         Width           =   4635
      End
   End
End
Attribute VB_Name = "FrmFin"
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

Dim consecutivo As Integer

Private Sub Form_Load()
Uno_skin Me

Me.Left = 25015 + 800

''Calculo Duracion de la pelea
'Duracion = Format(TimeValue("00:15:00") - TimeValue(Duracion), "hh:mm:ss")
'
'consecutivo = HallaConsecutivo("Select Max(Consecutivo) As NConsecutivo from Informacion where Fecha=#" & Month(FechaTrabajo) & "/" & Day(FechaTrabajo) & "/" & Year(FechaTrabajo) & "#")
''MsgBox "conse fin" & consecutivo
'Dim qry2 As New rdoQuery
'SQL = "Insert into Informacion values(" & consecutivo & ",'" & FechaTrabajo & "','" & HorInicio & "','" & Format(Time, "hh:mm:ss") & "','" & VCuerda1 & "','" & VColor1 & "','" & VCuerda2 & "','" & VColor2 & "'," & VValor & ",'" & Duracion & "','')"
'Set qry2.ActiveConnection = RDOCONEXION
'    qry2.SQL = SQL
'    qry2.Execute
'    qry2.Close
'
Set speaking2 = New SpeechLib.SpVoice
Set speaking2.Voice = speaking2.GetVoices().Item(voz)
    speaking2.Rate = -1
    speaking2.Volume = 100

    m = speaking2.Speak("Fin de la pelea", SVSFlagsAsync)
End Sub
Private Sub Timer1_Timer()

If fin <> 1 Then
    Unload Me
    FrmGanador.Show
Else
    Unload Me
End If

End Sub

