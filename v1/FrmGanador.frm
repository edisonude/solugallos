VERSION 5.00
Begin VB.Form FrmGanador 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   14910
   ClientLeft      =   -180
   ClientTop       =   0
   ClientWidth     =   24945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   14910
   ScaleWidth      =   24945
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   -1080
      TabIndex        =   4
      Top             =   4320
      Width           =   240
   End
   Begin VB.PictureBox picWall 
      Height          =   18215
      Left            =   -120
      Picture         =   "FrmGanador.frx":0000
      ScaleHeight     =   18150
      ScaleWidth      =   28155
      TabIndex        =   5
      Top             =   0
      Width           =   28215
      Begin VB.Image Image1 
         Height          =   5010
         Left            =   12735
         Picture         =   "FrmGanador.frx":BB8042
         Top             =   4620
         Width           =   4470
      End
      Begin VB.Image Image3 
         Height          =   4230
         Left            =   12195
         Picture         =   "FrmGanador.frx":C01184
         Top             =   195
         Width           =   5985
      End
      Begin VB.Label tTiempo 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   99.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   2535
         Left            =   6900
         TabIndex        =   15
         Top             =   1680
         Width           =   6015
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   60
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000763EB&
         Height          =   1455
         Index           =   0
         Left            =   2400
         TabIndex        =   14
         Top             =   2145
         Width           =   4575
      End
      Begin VB.Label tIdGallo2 
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label tIdGallo1 
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label tIdPelea 
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape cuerda22 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   5775
         Left            =   19200
         Top             =   3840
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.Shape cuerda2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   4575
         Left            =   19560
         Top             =   4800
         Width           =   5295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "EMPATE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   60
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0048F277&
         Height          =   1815
         Left            =   11880
         TabIndex        =   10
         Top             =   13440
         Width           =   6735
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
         ForeColor       =   &H000763EB&
         Height          =   1455
         Index           =   3
         Left            =   3180
         TabIndex        =   9
         Top             =   345
         Width           =   3015
      End
      Begin VB.Label TPelea 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   99.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   2535
         Left            =   6840
         TabIndex        =   8
         Top             =   -240
         Width           =   2895
      End
      Begin VB.Label TCuerda1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Mandarinos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   69.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   960
         TabIndex        =   7
         Top             =   9720
         Width           =   12480
      End
      Begin VB.Label TCuerda2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Sable"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   69.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1815
         Left            =   16320
         TabIndex        =   6
         Top             =   9720
         Width           =   12255
      End
      Begin VB.Shape cuerda11 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   5775
         Left            =   4800
         Top             =   3840
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.Shape cuerda1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   4575
         Left            =   5280
         Top             =   4800
         Width           =   5295
      End
      Begin VB.Shape empate22 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   1695
         Left            =   11640
         Top             =   12000
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Shape empate33 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   1695
         Left            =   15240
         Top             =   12000
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Shape empate3 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   1215
         Left            =   15240
         Top             =   12240
         Width           =   3375
      End
      Begin VB.Shape empate2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   1215
         Left            =   11880
         Top             =   12240
         Width           =   3375
      End
   End
   Begin VB.Image IEmpate 
      Height          =   1455
      Index           =   1
      Left            =   14415
      Picture         =   "FrmGanador.frx":C53BA6
      Stretch         =   -1  'True
      Top             =   11655
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   2835
      Left            =   15240
      Picture         =   "FrmGanador.frx":C6125C
      Top             =   7320
      Width           =   2475
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuerda 1:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   9600
      TabIndex        =   1
      Top             =   11760
      Width           =   3975
   End
   Begin VB.Label TConsecutivo 
      Height          =   210
      Left            =   -1290
      TabIndex        =   3
      Top             =   1245
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuerda 2:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   19200
      TabIndex        =   2
      Top             =   11760
      Width           =   4500
   End
   Begin VB.Label TOpcion 
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Height          =   6255
      Index           =   1
      Left            =   18000
      TabIndex        =   0
      Top             =   5160
      Width           =   4815
   End
   Begin VB.Image IEmpate 
      Height          =   1080
      Index           =   0
      Left            =   14730
      Picture         =   "FrmGanador.frx":C780CE
      Top             =   11835
      Width           =   3630
   End
End
Attribute VB_Name = "FrmGanador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As Integer
Dim Winn As String
Dim galloGanador As Integer
Dim puntos1 As Integer
Dim puntos2 As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 33 Or KeyCode = 34 Then
    con = con + 1
    Select Case con
        Case 1
            Call ocultarTodo
            Me.cuerda11.Visible = True
            Winn = Me.TCuerda1
            galloGanador = Me.tIdGallo1
             puntos1 = 2
            puntos2 = 0
        Case 2
            Call ocultarTodo
            Me.cuerda22.Visible = True
            Winn = Me.TCuerda2
            galloGanador = Me.tIdGallo2
                        puntos1 = 0
            puntos2 = 2
        Case 3
            Call ocultarTodo
            Me.empate22.Visible = True
            Me.empate33.Visible = True
            Winn = "EMPATE"
            galloGanador = -1
            puntos1 = 1
            puntos2 = 1
            con = 0
    End Select
End If

If KeyCode = 116 Then
    If Winn = "EMPATE" Then
        Duracion = Format(confTiempo, "00") & ":00"
    End If
    
        SQL = "UPDATE Peleas SET jugada='si', tiempo='" & Duracion & "', ganador =" & galloGanador & ", puntos1 =" & puntos1 & ", puntos2 =" & puntos2 & " WHERE idPelea=" & Me.tIdPelea & ""
        'SQL = "UPDATE Peleas SET tiempo='" & Duracion & "', ganador =" & galloGanador & " WHERE idPelea=" & Me.tIdPelea & ""
        Call guardarRDO
        
    FrmGano.TWin = Winn
    FrmGano.Show
    FrmGano.TWin = Winn
    
    If hayRed Then
        frmStart.establecerGanador (Winn)
    End If
    
     VCuerda1 = ""
     VColor1 = ""
     VCuerda2 = ""
     VColor2 = ""
     VValor = ""
     Duracion = ""
     HorInicio = ""
     
     'ganador
     FrmPantalla.tIdPelea = ""
     frmStart.cargarPeleas
     Unload Me
End If
End Sub

Private Sub Form_Load()
con = 0

'Preparación de la ventana
Me.TCuerda1.ForeColor = confColor1
Me.cuerda1.BackColor = confColor1
Me.cuerda11.BackColor = confColor1
Me.empate22.BackColor = confColor1
Me.empate2.BackColor = confColor1

Me.TCuerda2.ForeColor = confColor2
Me.cuerda2.BackColor = confColor2
Me.cuerda22.BackColor = confColor2
Me.empate33.BackColor = confColor2
Me.empate3.BackColor = confColor2

Me.Left = 15015
galloGanador = -1
End Sub

Private Sub TOpcion_Click(Index As Integer)
'Dim qry As New rdoQuery
'sql = "Update Informacion set Ganador='" & Winn & "' where Consecutivo=" & Me.TConsecutivo & ""
'Set qry.ActiveConnection = RDOCONEXION
'    qry.sql = sql
'    qry.Execute
'    qry.Close
'
'Unload Me
'
' VCuerda1 = ""
' VColor1 = ""
' VCuerda2 = ""
' VColor2 = ""
' VValor = ""
' Duracion = ""
' HorInicio = ""
'
'FrmInfo.Show
End Sub

Public Sub ocultarTodo()
Me.cuerda11.Visible = False
Me.cuerda22.Visible = False
Me.empate22.Visible = False
Me.empate33.Visible = False
End Sub

