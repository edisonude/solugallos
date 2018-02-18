VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmPeleaLibre 
   BackColor       =   &H00555555&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00555555&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   240
      ScaleHeight     =   135
      ScaleWidth      =   9015
      TabIndex        =   13
      Top             =   480
      Width           =   9015
   End
   Begin VB.PictureBox picLabel 
      BackColor       =   &H00555555&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   9015
      TabIndex        =   11
      Top             =   120
      Width           =   9015
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gallera Libre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   5535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   4935
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   9015
      Begin VB.CheckBox chkValor 
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   4080
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox tValor 
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
         Left            =   1800
         TabIndex        =   5
         Top             =   4080
         Width           =   3375
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00808080&
         Caption         =   "Cuerda 2"
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
         Left            =   240
         TabIndex        =   18
         Top             =   2400
         Width           =   5535
         Begin VB.CheckBox chkDato22 
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkDato11 
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox tDato22 
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
            Left            =   1560
            TabIndex        =   4
            Top             =   840
            Width           =   3375
         End
         Begin VB.TextBox tDato11 
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
            Left            =   1560
            TabIndex        =   3
            Top             =   360
            Width           =   3375
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   1
            Left            =   480
            OleObjectBlob   =   "frmPeleaLibre.frx":0000
            TabIndex        =   19
            Top             =   870
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   3
            Left            =   480
            OleObjectBlob   =   "frmPeleaLibre.frx":0064
            TabIndex        =   20
            Top             =   390
            Width           =   975
         End
      End
      Begin VB.TextBox tNoPelea 
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
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Asignar"
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
         Left            =   6240
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton btnMen 
         Caption         =   "Mensajes"
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
         Left            =   6240
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdIni 
         Caption         =   "Iniciar/Pausar Reloj "
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
         Left            =   6240
         TabIndex        =   7
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         Caption         =   "Cuerda 1"
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
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   5535
         Begin VB.CheckBox chkDato2 
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkDato1 
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox tDato1 
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
            Left            =   1560
            TabIndex        =   1
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox tDato2 
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
            Left            =   1560
            TabIndex        =   2
            Top             =   840
            Width           =   3375
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   0
            Left            =   480
            OleObjectBlob   =   "frmPeleaLibre.frx":00C8
            TabIndex        =   15
            Top             =   870
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   2
            Left            =   480
            OleObjectBlob   =   "frmPeleaLibre.frx":012C
            TabIndex        =   16
            Top             =   390
            Width           =   975
         End
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
         Left            =   6240
         TabIndex        =   9
         Top             =   3600
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   4
         Left            =   240
         OleObjectBlob   =   "frmPeleaLibre.frx":0190
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   5
         Left            =   720
         OleObjectBlob   =   "frmPeleaLibre.frx":0202
         TabIndex        =   21
         Top             =   4110
         Width           =   975
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   -120
      OleObjectBlob   =   "frmPeleaLibre.frx":0264
      Top             =   480
   End
End
Attribute VB_Name = "frmPeleaLibre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tTotalGallos_DragDrop(Source As Control, X As Single, Y As Single)
End Sub

Private Sub frmDatos_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub btnSalir_Click()
Me.Hide
End Sub

Private Sub cmdAsignar_Click()
Unload FrmPantalla
Unload FrmPantallaLibre
FrmPantallaLibre.Show
FrmPantallaLibre.tIdPelea = 1
FrmPantallaLibre.TCuerda1 = IIf(Me.chkDato2.Value = 1, Me.tDato2, "")
FrmPantallaLibre.TCuerda2 = IIf(Me.chkDato22.Value = 1, Me.tDato22, "")
FrmPantallaLibre.tIdGallo1 = 0
FrmPantallaLibre.tIdGallo2 = 0
FrmPantallaLibre.tCiudad1 = IIf(Me.chkDato1.Value = 1, Me.tDato1, "")
FrmPantallaLibre.tCiudad2 = IIf(Me.chkDato11.Value = 1, Me.tDato11, "")


'Dim peso As String
'If Me.tPeso1 = Me.tPeso2 Then
'    peso = Me.tPeso1
'Else
'    peso = Me.tPeso1 & "/" & Me.tPeso2
'End If

FrmPantallaLibre.TPelea = tNoPelea
FrmPantallaLibre.tPeso = IIf(Me.chkValor.Value = 1, FormatCurrency(Me.tValor, 0), "")
If FrmPantallaLibre.tPeso = "" Then
    FrmPantallaLibre.etiqueta(2).Visible = False
Else
    FrmPantallaLibre.etiqueta(2).Visible = True
End If

FrmPantallaLibre.SetFocus
End Sub

Private Sub cmdIni_Click()
FrmPantallaLibre.ini
FrmPantallaLibre.SetFocus
End Sub

Private Sub Form_Load()
Cooper_skin Me

Me.Top = frmMenu.ubicacion.Top + 350
Me.Left = frmMenu.ubicacion.Left
End Sub

Private Sub tValor_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
End Sub
