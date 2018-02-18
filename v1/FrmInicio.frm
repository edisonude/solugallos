VERSION 5.00
Begin VB.Form FrmInicio 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmInicio.frx":0000
   ScaleHeight     =   5430
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   4560
   End
End
Attribute VB_Name = "FrmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Para voz
Dim WithEvents speaking2 As SpeechLib.SpVoice
Attribute speaking2.VB_VarHelpID = -1
Dim m As Long

Private Sub Form_Load()
Me.Left = 15015 + 1000

'CALCULA EL PROXIMO SABADO O EL ANTERIOR
Dim FechaInicial As Date
FechaInicial = Format(Date, "dd/mm/yyyy")

If Weekday(Format(Date, "dd/mm/yyyy")) <= 4 Then
    While Weekday(FechaInicial) <> 7
        FechaInicial = FechaInicial - 1
    Wend
Else
    While Weekday(FechaInicial) <> 7
        FechaInicial = FechaInicial + 1
    Wend
End If

FechaTrabajo = FechaInicial

Set speaking2 = New SpeechLib.SpVoice
Set speaking2.Voice = speaking2.GetVoices().Item(voz)
    speaking2.Rate = -1
    speaking2.Volume = 100

    m = speaking2.Speak("Galleros Colombia    COMIGA", SVSFlagsAsync)

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Unload Me

frmMenu.Show
FrmPantalla.Show
'FrmInfo.Show

End Sub
