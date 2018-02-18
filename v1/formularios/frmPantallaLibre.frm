VERSION 5.00
Begin VB.Form FrmPantallaLibre 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   16515
   ClientLeft      =   225
   ClientTop       =   225
   ClientWidth     =   25365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   16515
   ScaleWidth      =   25365
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   2760
   End
   Begin VB.PictureBox picEspuelas 
      Appearance      =   0  'Flat
      BackColor       =   &H0088AAEA&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4485
      Left            =   9240
      ScaleHeight     =   4485
      ScaleWidth      =   11580
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   11580
      Begin VB.Timer tEnredados 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   720
         Top             =   360
      End
      Begin VB.Timer tEspuelas 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   360
      End
      Begin VB.Label tiempoEspuelas 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "03:00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   219.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4245
         Left            =   -240
         TabIndex        =   3
         Top             =   -240
         Width           =   12105
      End
   End
   Begin VB.PictureBox picArena1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   960
      ScaleHeight     =   6015
      ScaleWidth      =   6180
      TabIndex        =   0
      Top             =   7410
      Visible         =   0   'False
      Width           =   6180
      Begin VB.Timer tiempoDobleArena1 
         Enabled         =   0   'False
         Interval        =   800
         Left            =   600
         Top             =   360
      End
      Begin VB.Timer TArena1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   135
         Top             =   375
      End
      Begin VB.Label tiempoArena1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "57"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   219.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4245
         Left            =   -2790
         TabIndex        =   1
         Top             =   420
         Width           =   12105
      End
      Begin VB.Image Image1 
         Height          =   6045
         Left            =   0
         Picture         =   "frmPantallaLibre.frx":0000
         Top             =   0
         Width           =   6390
      End
   End
   Begin VB.PictureBox picWall 
      Height          =   18215
      Left            =   0
      Picture         =   "frmPantallaLibre.frx":7DF42
      ScaleHeight     =   18150
      ScaleWidth      =   28155
      TabIndex        =   4
      Top             =   0
      Width           =   28215
      Begin VB.Timer timeTeclas 
         Interval        =   10
         Left            =   2280
         Top             =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   735
         Left            =   4920
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Timer tiempoDobleFin 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   480
         Top             =   480
      End
      Begin VB.Timer tiempoDobleB 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   0
         Top             =   480
      End
      Begin VB.PictureBox picArena2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5865
         Left            =   21585
         ScaleHeight     =   5865
         ScaleWidth      =   5925
         TabIndex        =   5
         Top             =   7605
         Visible         =   0   'False
         Width           =   5925
         Begin VB.Timer tiempoDobleArena2 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   600
            Top             =   360
         End
         Begin VB.Timer TArena2 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   120
            Top             =   360
         End
         Begin VB.Label tiempoArena2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "57"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   219.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   4245
            Left            =   -3045
            TabIndex        =   6
            Top             =   300
            Width           =   12105
         End
         Begin VB.Image iArena2 
            Height          =   5895
            Left            =   -150
            Picture         =   "frmPantallaLibre.frx":C35F84
            Top             =   -15
            Width           =   6075
         End
      End
      Begin VB.Label tPeso 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "$ 10.000.000"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   80.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2250
         Left            =   15960
         TabIndex        =   12
         Top             =   840
         Width           =   9135
      End
      Begin VB.Label tValor 
         BackStyle       =   0  'Transparent
         Caption         =   "$ 1.000.000"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   99.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2535
         Left            =   11400
         TabIndex        =   24
         Top             =   13200
         Width           =   10710
      End
      Begin VB.Label etiquetaValor 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   60
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1455
         Left            =   7680
         TabIndex        =   23
         Top             =   13680
         Width           =   3015
      End
      Begin VB.Label tCiudad1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ASOGAL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   80.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1815
         Left            =   480
         TabIndex        =   22
         Top             =   3000
         Width           =   12255
      End
      Begin VB.Label tCiudad2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ASOGAL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   80.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   15360
         TabIndex        =   21
         Top             =   3000
         Width           =   12255
      End
      Begin VB.Label TCuerda2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ASOGAL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   99.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3360
         Left            =   15360
         TabIndex        =   20
         Top             =   4800
         Width           =   12255
      End
      Begin VB.Label TColor2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   60
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1455
         Left            =   15360
         TabIndex        =   19
         Top             =   7320
         Width           =   12480
      End
      Begin VB.Label TCuerda1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ASOGAL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   90
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3360
         Left            =   240
         TabIndex        =   18
         Top             =   4800
         Width           =   12480
      End
      Begin VB.Label TColor1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   60
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1455
         Left            =   240
         TabIndex        =   17
         Top             =   7320
         Width           =   12480
      End
      Begin VB.Label tIdGallo2 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label tIdGallo1 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label tIdPelea 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblLetrero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "d"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   80.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00089400&
         Height          =   1965
         Left            =   4920
         TabIndex        =   13
         Top             =   13320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image imgVersus 
         Height          =   3120
         Left            =   12690
         Picture         =   "frmPantallaLibre.frx":CAAA86
         Top             =   4635
         Width           =   2865
      End
      Begin VB.Image imgLogo 
         Height          =   4230
         Left            =   11175
         Picture         =   "frmPantallaLibre.frx":CC7EC8
         Top             =   195
         Width           =   5985
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   39.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   975
         Index           =   2
         Left            =   19680
         TabIndex        =   11
         Top             =   120
         Width           =   4095
      End
      Begin VB.Label TPelea 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2535
         Left            =   5400
         TabIndex        =   10
         Top             =   -75
         Width           =   4830
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Pelea:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   60
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1455
         Index           =   3
         Left            =   1770
         TabIndex        =   9
         Top             =   945
         Width           =   3015
      End
      Begin VB.Label LTurno 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   300
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   6045
         Left            =   7140
         TabIndex        =   8
         Top             =   6825
         Width           =   14655
      End
      Begin VB.Label LTurno2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   300
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   6045
         Left            =   6960
         TabIndex        =   25
         Top             =   6960
         Width           =   14655
      End
   End
   Begin VB.Timer TimeMarquesina2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer TimeMarquesina 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   555
      Top             =   0
   End
   Begin VB.Timer MyMarquesina 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Ap"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   120
      TabIndex        =   26
      Top             =   7080
      Width           =   12480
   End
End
Attribute VB_Name = "FrmPantallaLibre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim hh, mm, ss As Integer
Dim mmE, ssE As Integer
'Contadores para los relojes de arena
Dim arena1, arena2 As Integer
Dim dobleArena1 As Integer
Dim dobleArena2 As Integer
Dim dobleFin As Integer
Dim dobleB As Integer

'Para voz
Dim WithEvents speaking As SpeechLib.SpVoice
Attribute speaking.VB_VarHelpID = -1
Dim WithEvents speaking2 As SpeechLib.SpVoice
Attribute speaking2.VB_VarHelpID = -1
Private speakingvoice As SpeechLib.ISpeechObjectToken
Dim m, n As Long
Dim RecdTime As Boolean

'Declaramos el Api GetAsyncKeyState
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Inicia la pelea desde el Start
Public Sub ini()
If Timer1.Enabled = True Then
    Timer1.Enabled = False
Else
    Timer1.Enabled = True
    HorInicio = Format(Time, "hh:mm:ss")
End If
End Sub

Private Sub Form_Load()
'Carga los valores de congiracion
Call cargarValoresConfig

'Verifico modo prueba
If confModoPrueba = "Si" Then
    Me.Timer1.Interval = 150
    Me.TArena1.Interval = 150
    Me.TArena2.Interval = 150
End If

'Posicionamiento de objetos
Me.picWall.Top = 0
Me.Top = 400
Me.Left = 21100
Me.Width = 27495
Me.Height = 16000
'Ubicacion reloj de arena
Me.picArena2.Top = 7545
Me.picArena2.Left = 21285

'Configuracion de tiempos
ss = 0: mm = 0
ssE = 0: mmE = confTiempoEspuelas


'Preparación de la ventana
Me.TCuerda1.ForeColor = confColor1
Me.tCiudad1.ForeColor = confColor1
Me.TCuerda2.ForeColor = confColor2
Me.tCiudad2.ForeColor = confColor2

If confVerValor = "si" Then
    Me.etiquetaValor.Visible = True
    Me.tValor.Visible = True
Else
    Me.etiquetaValor.Visible = False
    Me.tValor.Visible = False
End If

If confVerColorGallo = "Si" Then
    Me.TColor1.Visible = True
    Me.TColor2.Visible = True
Else
    Me.TColor1.Visible = False
    Me.TColor2.Visible = False
End If

'Inicializacion de relojes
dobleArena1 = 0
dobleArena2 = 0
dobleFin = 0
dobleB = 0
End Sub

Private Sub LTurno_Change()
'frmpantalla
LTurno2.Caption = LTurno.Caption
End Sub

Private Sub MyMarquesina_Timer()
Call Marquesina
End Sub

Private Sub TArena2_Timer()
arena2 = arena2 - 1
Me.tiempoArena2 = Format(arena2, "00")

If arena2 = 0 And Me.TArena1.Enabled = True Then
    Me.TArena2.Enabled = False
Else
    If arena2 = 0 Then
        Me.TArena2.Enabled = False
        FinPelea
    End If
End If
End Sub
Private Sub TArena1_Timer()
arena1 = arena1 - 1
Me.tiempoArena1 = Format(arena1, "00")

If arena1 = 0 And Me.TArena2.Enabled = True Then
    Me.TArena1.Enabled = False
Else
    If arena1 = 0 Then
        Me.TArena1.Enabled = False
        FinPelea
    End If
End If
End Sub

Private Sub tEspuelas_Timer()
If ssE = 0 Then
    mmE = mmE - 1
    ssE = 60
End If
ssE = ssE - 1

Me.tiempoEspuelas.Caption = Format(mmE, "00") & ":" & Format(ssE, "00")

If mmE = 0 And ssE = 0 Then
    Call hablar("Finaliza tiempo para cambio de espuelas")
    Me.picEspuelas.Visible = False
    Me.tEspuelas.Enabled = False
End If
End Sub

Private Sub tiempoDobleArena1_Timer()
dobleArena1 = 0
tiempoDobleArena1.Enabled = False
End Sub

Private Sub tiempoDobleArena2_Timer()
dobleArena2 = 0
tiempoDobleArena2.Enabled = False
End Sub

Private Sub tiempoDobleB_Timer()
dobleB = 0
tiempoDobleB.Enabled = False
End Sub

Private Sub tiempoDobleFin_Timer()
dobleFin = 0
tiempoDobleFin.Enabled = False
End Sub

Private Sub TimeMarquesina_Timer()
    TimeMarquesina.Enabled = False
    TimeMarquesina2.Enabled = True
End Sub

Private Sub TimeMarquesina2_Timer()
    TimeMarquesina2.Enabled = False
    MyMarquesina.Enabled = False
    Me.lblLetrero.Visible = False
End Sub

Private Sub Timer1_Timer()
Dim retraso As Long

If ss = 59 Then
    mm = mm + 1
    ss = -1
End If
ss = ss + 1

LTurno.Caption = Format(mm, "00") & ":" & Format(ss, "00")

'Invocar espuelas
If mm = 5 And ss = 0 Then
    FrmEspuelas.Show
Else
    'Detengo el tiempo en el ultimo segundo por reloj de arena
    If (mm = confTiempo - 1 And ss = 59 And (Me.picArena1.Visible = True Or Me.picArena2.Visible = True)) Then
        Me.Timer1.Enabled = False
    Else
        If mm = confTiempo And ss = 0 Then
            FinPelea
        End If
    End If
End If

'
'Select Case LTurno.Caption
'Case "05:00"
'    FrmEspuelas.Show
'Case "09:59"
'    If Me.picArena1.Visible = True Or Me.picArena2.Visible = True Then
'        Me.Timer1.Enabled = False
'    End If
'Case "10:00"
'    FinPelea
'End Select


End Sub


Private Function FinPelea()
'    FrmGanador.tIdPelea = Me.tIdPelea
'    FrmGanador.TPelea = Me.TPelea
'    FrmGanador.tCuerda1 = Me.tCuerda1
'    FrmGanador.tCuerda2 = Me.tCuerda2
'    FrmGanador.tIdGallo1 = Me.tIdGallo1
'    FrmGanador.tIdGallo2 = Me.tIdGallo2
'
'    'Manejo de tiempo incluyendo o no incluyendo el tiempo de reloj de arena
'    If confArenaIncluido = "no" Then
'        Dim tiem As Integer
'        tiem = tiempoaEntero(Me.LTurno)
'        tiem = tiem - confTiempoArena
'        Me.LTurno = pasarAHora(tiem)
'    End If
fin = 1

    FrmFin.Show
    Me.Timer1.Enabled = False
End Function

Private Sub Marquesina()
    'Realiza la funcion de marquecina en el formulario
    lblLetrero.Left = lblLetrero.Left - 70
    If lblLetrero.Left < -1 * (lblLetrero.Width) Then
        lblLetrero.Left = Me.Width
    End If
End Sub



Public Function Mejores()

               ' Me.lblLetrero.Caption = "La pelea más rápida con una duración de: " & Right(Rs(0), 5) & " fue ganada por:" & Rs(1) & ""
                TimeMarquesina.Enabled = True
                MyMarquesina.Enabled = True
                Me.lblLetrero.Visible = True

End Function
Public Function QuitarMejores()

               ' Me.lblLetrero.Caption = "La pelea más rápida con una duración de: " & Right(Rs(0), 5) & " fue ganada por:" & Rs(1) & ""
                TimeMarquesina.Enabled = False
                MyMarquesina.Enabled = False
                Me.lblLetrero.Visible = False

End Function


Private Sub hablar(mensaje As String)
Set speaking2 = New SpeechLib.SpVoice
Set speaking2.Voice = speaking2.GetVoices().Item(voz)
    speaking2.Rate = -1
    speaking2.Volume = 100
    
    m = speaking2.Speak(mensaje, SVSFlagsAsync)
End Sub

Private Sub timeTeclas_Timer()
If Me.Visible = False Or Me.tIdPelea = "" Then Exit Sub
If GetAsyncKeyState(33) = -32767 Then 'REPAG
    dobleArena1 = dobleArena1 + 1
    Me.tiempoDobleArena1.Enabled = True
    If dobleArena1 = 2 Then
        dobleArena1 = 0
        
        If Timer1.Enabled = False Then
            If Me.picEspuelas.Visible = False Then
                Call hablar("Inicia cambio de espuelas")
                mmE = confTiempoEspuelas: ssE = 0
                Me.tiempoEspuelas.Caption = Format(mmE, "00") & ":" & Format(ssE, "00")
                Me.tEspuelas.Enabled = True
                Me.picEspuelas.Visible = True
            Else
                Call hablar("Finaliza tiempo para cambio de espuelas")
                Me.picEspuelas.Visible = False
                Me.tEspuelas.Enabled = False
            End If
        Else
            If picArena1.Visible = False Then
                arena1 = confTiempoArena
                Me.tiempoArena1 = confTiempoArena
                Me.picArena1.Visible = True
                Me.TArena1.Enabled = True
            Else
                If picArena2.Visible = True And Me.TArena2.Enabled = False And arena2 = 0 Then
                    FinPelea
                End If
                    Me.picArena1.Visible = False
                    Me.TArena1.Enabled = False
            End If
        End If
    End If
Else
    If GetAsyncKeyState(34) = -32767 Then 'REPAG
        dobleArena2 = dobleArena2 + 1
        Me.tiempoDobleArena2.Enabled = True
        If dobleArena2 = 2 Then
            dobleArena2 = 0
            
             If Timer1.Enabled = False Then
                If Me.picEspuelas.Visible = False Then
                    Call hablar("Inicia cambio de espuelas")
                    mmE = confTiempoEspuelas: ssE = 0
                    Me.tiempoEspuelas.Caption = Format(mmE, "00") & ":" & Format(ssE, "00")
                    Me.tEspuelas.Enabled = True
                    Me.picEspuelas.Visible = True
                Else
                    Call hablar("Finaliza tiempo para cambio de espuelas")
                    Me.picEspuelas.Visible = False
                    Me.tEspuelas.Enabled = False
                    mmE = confTiempoEspuelas: ssE = 0
                End If
            Else
                If picArena2.Visible = False Then
                    arena2 = confTiempoArena
                    Me.tiempoArena2 = confTiempoArena
                    Me.picArena2.Visible = True
                    Me.TArena2.Enabled = True
                Else
                    If picArena1.Visible = True And Me.TArena1.Enabled = False And arena1 = 0 Then
                        FinPelea
                    End If
                    Me.picArena2.Visible = False
                    Me.TArena2.Enabled = False
                End If
            End If
        End If
    Else
        If GetAsyncKeyState(66) = -32767 Then 'B
            dobleB = dobleB + 1
            Me.tiempoDobleB.Enabled = True
            If dobleB = 2 Then
                dobleB = 0
                If Timer1.Enabled = True Then
                    Timer1.Enabled = False
                Else
                    Timer1.Enabled = True
                    HorInicio = Format(Time, "hh:mm:ss")
                End If
            End If
        Else
            If GetAsyncKeyState(116) = -32767 Then 'F5
                If Me.Timer1.Enabled = True Then
                dobleFin = dobleFin + 1
                Me.tiempoDobleFin.Enabled = True
                If dobleFin = 2 Then
                    dobleFin = 0
                    FinPelea
                End If
                End If
            Else
                If GetAsyncKeyState(27) = -32767 Then 'F5
                    'Me.Hide
                End If
            End If
        End If
    End If
End If
End Sub



