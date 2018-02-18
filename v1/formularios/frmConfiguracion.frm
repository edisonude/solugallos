VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{ABC786B9-EE6E-43EC-A219-9ECAC7B40E0E}#1.0#0"; "ColorPicker.ocx"
Begin VB.Form frmConfiguracionTablero 
   BackColor       =   &H00555555&
   BorderStyle     =   0  'None
   Caption         =   "frmAjustes"
   ClientHeight    =   8685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmMarco 
      Appearance      =   0  'Flat
      BackColor       =   &H00555555&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8655
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   8535
      Begin VB.CheckBox chkReportesA4 
         Caption         =   "Reportes en A4"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1215
         TabIndex        =   28
         Top             =   7545
         Width           =   2055
      End
      Begin VB.CheckBox chkPrueba 
         Caption         =   "Habilitar modo prueba"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   27
         Top             =   7560
         Width           =   2295
      End
      Begin VB.CommandButton btnSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   26
         Top             =   7920
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00808080&
         Caption         =   "Voz del sistema"
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   960
         TabIndex        =   20
         Top             =   3000
         Width           =   6495
         Begin VB.ComboBox listVoces 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   240
            TabIndex        =   23
            Top             =   840
            Width           =   3855
         End
         Begin VB.CommandButton cmdProbarVoz 
            Caption         =   "Probar"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4440
            TabIndex        =   21
            Top             =   840
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   2
            Left            =   240
            OleObjectBlob   =   "frmConfiguracion.frx":0000
            TabIndex        =   22
            Top             =   480
            Width           =   3495
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00808080&
         Caption         =   "Componentes de la ventana"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   960
         TabIndex        =   13
         Top             =   6120
         Width           =   6495
         Begin VB.CheckBox verValor 
            Caption         =   "Ver valor"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox verColorGallo 
            Caption         =   "Ver color de gallo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   16
            Top             =   480
            Width           =   2055
         End
         Begin VB.CheckBox verAnuncios 
            Caption         =   "Ver anuncios"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            TabIndex        =   15
            Top             =   480
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox verCaidoCorrido 
            Caption         =   "Notificar gallo caido o corrido"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   14
            Top             =   840
            Width           =   3255
         End
      End
      Begin VB.CommandButton btnAplicarAjustes 
         Caption         =   "Aplicar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2085
         TabIndex        =   12
         Top             =   7920
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         Caption         =   "Control de los tiempos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   7815
         Begin VB.ComboBox tArenaIncluido 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "frmConfiguracion.frx":009A
            Left            =   4920
            List            =   "frmConfiguracion.frx":00A4
            TabIndex        =   25
            Text            =   "si"
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox tTiempoArena 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6000
            TabIndex        =   18
            Text            =   "5"
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox tTiempoAjuste 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4920
            TabIndex        =   9
            Text            =   "10"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox tTiempoEspuelasAjuste 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4920
            TabIndex        =   8
            Text            =   "5"
            Top             =   960
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   5
            Left            =   240
            OleObjectBlob   =   "frmConfiguracion.frx":00B0
            TabIndex        =   10
            Top             =   480
            Width           =   4575
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   40
            Left            =   240
            OleObjectBlob   =   "frmConfiguracion.frx":0162
            TabIndex        =   11
            Top             =   960
            Width           =   4575
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   1
            Left            =   240
            OleObjectBlob   =   "frmConfiguracion.frx":020C
            TabIndex        =   19
            Top             =   1440
            Width           =   5655
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   4
            Left            =   240
            OleObjectBlob   =   "frmConfiguracion.frx":02D2
            TabIndex        =   24
            Top             =   1920
            Width           =   4575
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00808080&
         Caption         =   "Color texto de las cuerdas"
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   960
         TabIndex        =   1
         Top             =   4440
         Width           =   6495
         Begin VB.CommandButton btnReestablecerColores 
            Caption         =   "Reestablecer"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4440
            TabIndex        =   2
            Top             =   960
            Width           =   1695
         End
         Begin ClrPckr.ColorPicker tColorCuerda1 
            Height          =   375
            Left            =   2160
            TabIndex        =   3
            Top             =   480
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
         End
         Begin ClrPckr.ColorPicker tColorCuerda2 
            Height          =   375
            Left            =   2160
            TabIndex        =   4
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   3
            Left            =   480
            OleObjectBlob   =   "frmConfiguracion.frx":0374
            TabIndex        =   5
            Top             =   480
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   0
            Left            =   480
            OleObjectBlob   =   "frmConfiguracion.frx":03EA
            TabIndex        =   6
            Top             =   960
            Width           =   1575
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   0
         OleObjectBlob   =   "frmConfiguracion.frx":0460
         Top             =   120
      End
   End
End
Attribute VB_Name = "frmConfiguracionTablero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents speaking2 As SpeechLib.SpVoice
Attribute speaking2.VB_VarHelpID = -1
Private Sub btnAplicarAjustes_Click()
'Verifico datos nulos
If Me.tTiempoAjuste = "" Or Me.tTiempoEspuelasAjuste = "" Then
    MsgBox "Asigne los tiempos para la pelea y el cambio de espuelas", vbCritical
    Exit Sub
End If

'Actualizo los ajustes generales
SQL = "Update AjustesGenerales SET " & _
        "tiempo=" & Me.tTiempoAjuste & ", " & _
        "tiempoEspuelas=" & Me.tTiempoEspuelasAjuste & ", " & _
        "tiempoArena=" & Me.tTiempoArena & ", " & _
        "arenaIncluido='" & Me.tArenaIncluido & "', " & _
        "direccionReloj='ascendente', " & _
        "cuerda1='" & Me.tColorCuerda1.color & "', " & _
        "cuerda2='" & Me.tColorCuerda2.color & "', " & _
        "valor='" & IIf(Me.verValor.Value = 1, "Si", "No") & "', " & _
        "colorGallo='" & IIf(Me.verColorGallo.Value = 1, "Si", "No") & "', " & _
        "anuncios='" & IIf(Me.verAnuncios.Value = 1, "Si", "No") & "', " & _
        "corridoCaido='" & IIf(Me.verCaidoCorrido.Value = 1, "Si", "No") & "', " & _
        "voz=" & Left(Me.listVoces.Text, 1) - 1 & "," & _
        "modoPrueba='" & IIf(Me.chkPrueba.Value = 1, "Si", "No") & "'," & _
        "reportesA4='" & IIf(Me.chkReportesA4.Value = 1, "Si", "No") & "'"

If guardarRDO Then
    menGuardadoExitoso
    Call cargarValoresConfig
Else
    menGuardadoFallo
End If
End Sub

Private Sub btnReestablecerColores_Click()
Me.tColorCuerda1.color = vbWhite
Me.tColorCuerda2.color = vbWhite
End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub cmdProbarVoz_Click()
Dim m As Long
Dim v As Integer
v = Left(Me.listVoces.Text, 1) - 1

Set speaking2 = New SpeechLib.SpVoice
Set speaking2.Voice = speaking2.GetVoices().Item(v)
    speaking2.Rate = -1
    speaking2.Volume = 100
    m = speaking2.Speak("Galleros Colombia", SVSFlagsAsync)
End Sub

Private Sub Form_Load()
Cooper_skin Me

'Carga lista de voces del sistema
Set speaking2 = New SpeechLib.SpVoice
For i = 0 To speaking2.GetVoices().Count - 1
    listVoces.AddItem i + 1 & ". " & speaking2.GetVoices.Item(i).GetDescription
Next

'Cargo los ajustes generales
Call cargarAjustes
End Sub

Private Function cargarAjustes() As Boolean
Dim qry As New rdoQuery
Dim rs As rdoResultset

SQL = "Select * from AjustesGenerales"
Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rs = qry.OpenResultset(rdOpenDynamic)
    'Llenos los datos
        Me.tTiempoAjuste = rs("tiempo")
        Me.tTiempoEspuelasAjuste = rs("tiempoEspuelas")
        Me.tTiempoArena = rs("tiempoArena")
        Me.tArenaIncluido = rs("arenaIncluido")
        'Me.tDireccionReloj = rs("direccionReloj")
        Me.tColorCuerda1.color = rs("cuerda1")
        Me.tColorCuerda2.color = rs("cuerda2")
        Me.verValor.Value = IIf(rs("valor") = "Si", 1, 0)
        Me.verAnuncios.Value = IIf(rs("anuncios") = "Si", 1, 0)
        Me.verCaidoCorrido.Value = IIf(rs("corridoCaido") = "Si", 1, 0)
        Me.verColorGallo.Value = IIf(rs("colorGallo") = "Si", 1, 0)
        Me.chkPrueba.Value = IIf(rs("modoPrueba") = "Si", 1, 0)
        Me.chkReportesA4.Value = IIf(rs("reportesA4") = "Si", 1, 0)
        
        Me.listVoces.Text = Me.listVoces.List(rs("voz"))
qry.Close
End Function

