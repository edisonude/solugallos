VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmStart 
   Appearance      =   0  'Flat
   BackColor       =   &H00555555&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalir 
      Caption         =   "x"
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
      Left            =   13725
      TabIndex        =   24
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Orden final de las peleas"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   13575
      Begin VB.CommandButton cmdIni 
         Caption         =   "ini"
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
         Left            =   12720
         TabIndex        =   38
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "foco"
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
         Left            =   12720
         TabIndex        =   34
         Top             =   1560
         Width           =   615
      End
      Begin VB.Frame frmMensaje 
         Caption         =   "Frame3"
         Height          =   3975
         Left            =   720
         TabIndex        =   31
         Top             =   840
         Visible         =   0   'False
         Width           =   10335
         Begin VB.CommandButton Command4 
            Caption         =   "Habl"
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
            Left            =   9120
            TabIndex        =   37
            Top             =   2160
            Width           =   615
         End
         Begin VB.CommandButton Command3 
            Caption         =   "x"
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
            Left            =   9120
            TabIndex        =   36
            Top             =   2760
            Width           =   615
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Del"
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
            Left            =   9120
            TabIndex        =   35
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox tmen 
            Height          =   3135
            Left            =   480
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   480
            Width           =   8535
         End
         Begin VB.CommandButton btnEnv 
            Caption         =   "Env"
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
            Left            =   9120
            TabIndex        =   32
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.CommandButton btnMen 
         Caption         =   "Men"
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
         Left            =   12720
         TabIndex        =   30
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton btnVolver 
         Caption         =   "<-"
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
         Left            =   12720
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton btnFin 
         Caption         =   "Fin"
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
         Left            =   12720
         TabIndex        =   18
         Top             =   4680
         Width           =   615
      End
      Begin VB.PictureBox Picture5 
         Height          =   4935
         Index           =   1
         Left            =   12600
         ScaleHeight     =   4935
         ScaleWidth      =   15
         TabIndex        =   15
         Top             =   360
         Width           =   15
      End
      Begin MSComctlLib.ListView listaGallos 
         Height          =   4860
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   8573
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IdPelea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cuerda"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pla. Cue"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Pla. Nac"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Anillo"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "VS"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cuerda 2"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Pla. Cue2"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Pla. Nac2"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Anillo"
            Object.Width           =   1940
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   615
         Index           =   2
         Left            =   12720
         OleObjectBlob   =   "frmStart.frx":0000
         TabIndex        =   16
         Top             =   3360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel tNPeleas 
         Height          =   495
         Left            =   12720
         OleObjectBlob   =   "frmStart.frx":006C
         TabIndex        =   17
         Top             =   4080
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00555555&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   360
      ScaleHeight     =   135
      ScaleWidth      =   12495
      TabIndex        =   2
      Top             =   480
      Width           =   12495
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmStart.frx":00C6
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   2535
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   11535
      Begin VB.CommandButton cmdLes2 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10920
         TabIndex        =   58
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdLes1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         TabIndex        =   57
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox Picture5 
         Height          =   1095
         Index           =   0
         Left            =   3120
         ScaleHeight     =   1095
         ScaleWidth      =   15
         TabIndex        =   56
         Top             =   240
         Width           =   15
      End
      Begin VB.OptionButton opRedAct 
         Caption         =   "Activada"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3240
         TabIndex        =   44
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton opRedDes 
         Caption         =   "Desactivada"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3240
         TabIndex        =   43
         Top             =   840
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton btnSiguiente 
         Caption         =   "Siguiente"
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
         Left            =   9240
         TabIndex        =   26
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton btnAsignar 
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
         Left            =   9240
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.PictureBox Picture4 
         Height          =   15
         Left            =   240
         ScaleHeight     =   15
         ScaleWidth      =   11055
         TabIndex        =   10
         Top             =   1320
         Width           =   11055
      End
      Begin VB.PictureBox Picture2 
         Height          =   975
         Left            =   5400
         ScaleHeight     =   915
         ScaleWidth      =   555
         TabIndex        =   0
         Top             =   1440
         Width           =   615
         Begin VB.Image Image1 
            Height          =   915
            Left            =   -120
            Picture         =   "frmStart.frx":02FA
            Stretch         =   -1  'True
            Top             =   0
            Width           =   750
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   4
         Left            =   240
         OleObjectBlob   =   "frmStart.frx":432C
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tCuerda1 
         Height          =   375
         Left            =   1320
         OleObjectBlob   =   "frmStart.frx":4390
         TabIndex        =   7
         Top             =   1800
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   9
         Left            =   240
         OleObjectBlob   =   "frmStart.frx":43FC
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPeso1 
         Height          =   375
         Left            =   1320
         OleObjectBlob   =   "frmStart.frx":445C
         TabIndex        =   9
         Top             =   2160
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   13
         Left            =   6360
         OleObjectBlob   =   "frmStart.frx":44C4
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tCuerda2 
         Height          =   375
         Left            =   7440
         OleObjectBlob   =   "frmStart.frx":4528
         TabIndex        =   12
         Top             =   1800
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   15
         Left            =   6360
         OleObjectBlob   =   "frmStart.frx":4594
         TabIndex        =   13
         Top             =   2160
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPeso2 
         Height          =   375
         Left            =   7440
         OleObjectBlob   =   "frmStart.frx":45F4
         TabIndex        =   14
         Top             =   2160
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   0
         Left            =   240
         OleObjectBlob   =   "frmStart.frx":465C
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPlaca1 
         Height          =   375
         Left            =   1320
         OleObjectBlob   =   "frmStart.frx":46BE
         TabIndex        =   20
         Top             =   1440
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel Placa 
         Height          =   375
         Index           =   3
         Left            =   6360
         OleObjectBlob   =   "frmStart.frx":4728
         TabIndex        =   21
         Top             =   1440
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPlaca2 
         Height          =   375
         Left            =   7440
         OleObjectBlob   =   "frmStart.frx":478A
         TabIndex        =   22
         Top             =   1440
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   5
         Left            =   240
         OleObjectBlob   =   "frmStart.frx":47F4
         TabIndex        =   45
         Top             =   240
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPeleasRestantes 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmStart.frx":486C
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   1
         Left            =   240
         OleObjectBlob   =   "frmStart.frx":48CC
         TabIndex        =   47
         Top             =   600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel tNumPelea 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmStart.frx":493C
         TabIndex        =   48
         Top             =   600
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tTiempo 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmStart.frx":499C
         TabIndex        =   49
         Top             =   960
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   3
         Left            =   240
         OleObjectBlob   =   "frmStart.frx":49FE
         TabIndex        =   50
         Top             =   960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   6
         Left            =   3240
         OleObjectBlob   =   "frmStart.frx":4A62
         TabIndex        =   51
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   7
         Left            =   4680
         OleObjectBlob   =   "frmStart.frx":4ADA
         TabIndex        =   52
         ToolTipText     =   "dsadasd"
         Top             =   600
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   8
         Left            =   4680
         OleObjectBlob   =   "frmStart.frx":4B36
         TabIndex        =   53
         Top             =   840
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel tIp 
         Height          =   255
         Left            =   5400
         OleObjectBlob   =   "frmStart.frx":4B9C
         TabIndex        =   54
         Top             =   600
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel tCliente 
         Height          =   255
         Left            =   5400
         OleObjectBlob   =   "frmStart.frx":4C1C
         TabIndex        =   55
         Top             =   840
         Width           =   1215
      End
      Begin MSWinsockLib.Winsock Red 
         Left            =   8760
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label tIdPelea 
         Height          =   255
         Left            =   11160
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Label tColor1 
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label tColor2 
      Height          =   375
      Left            =   0
      TabIndex        =   41
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label tCiudad1 
      Height          =   375
      Left            =   0
      TabIndex        =   40
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label tCiudad2 
      Height          =   375
      Left            =   0
      TabIndex        =   39
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label tIdGallo2 
      Height          =   375
      Left            =   0
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label tIdGallo1 
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Incio de las peleas"
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
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim li As ListItem
Dim encabezado As String
Dim numeroGallo As Integer
Dim nPelea As Integer
Dim numPeleas As Integer
Dim restantes As Integer

'Para voz
Dim WithEvents speaking As SpeechLib.SpVoice
Attribute speaking.VB_VarHelpID = -1
Dim WithEvents speaking2 As SpeechLib.SpVoice
Attribute speaking2.VB_VarHelpID = -1
Private speakingvoice As SpeechLib.ISpeechObjectToken
Dim m, n As Long
Dim RecdTime As Boolean


Private Sub btnEnv_Click()
FrmPantalla.lblLetrero.Caption = Me.tmen
FrmPantalla.Mejores
End Sub

Private Sub btnMen_Click()
frmMensaje.Visible = True
End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub btnAsignar_Click()
Dim peso As String
Unload FrmPantallaLibre
Unload FrmPantalla
FrmPantalla.Show
FrmPantalla.tIdPelea = Me.tIdPelea
FrmPantalla.TCuerda1 = Me.TCuerda1
FrmPantalla.TCuerda2 = Me.TCuerda2
FrmPantalla.tIdGallo1 = Me.tIdGallo1
FrmPantalla.tIdGallo2 = Me.tIdGallo2
FrmPantalla.tCiudad1 = Me.tCiudad1
FrmPantalla.tCiudad2 = Me.tCiudad2
FrmPantalla.TColor1 = Me.TColor1
FrmPantalla.TColor2 = Me.TColor2

If Me.tPeso1 = Me.tPeso2 Then
    peso = Me.tPeso1
Else
    peso = Me.tPeso1 & "/" & Me.tPeso2
End If

FrmPantalla.TPelea = Format(nPelea, "00")
FrmPantalla.tPeso = peso

FrmPantalla.SetFocus

If hayRed Then Call Me.actualizarPantalla
End Sub

Private Sub btnFin_Click()

With listaGallos
    For i = 1 To .ListItems.Count
        conPelea = HallaConsecutivo("Select Max(idPelea) As NConsecutivo from Peleas")
                   
        SQL = "UPDATE Peleas SET orden = " & .ListItems(i).SubItems(1) & " WHERE idPelea = " & .ListItems(i) & ""
        Call guardarRDO
    Next
End With
menGuardadoExitoso
Unload Me
End Sub

Private Sub cmdIni_Click()
FrmPantalla.ini
'For iCount = 1 To 1000
'Next
'FrmPantalla.ini
FrmPantalla.SetFocus
End Sub

Private Sub cmdLes1_Click()
FrmPantalla.TCuerda1.FontSize = FrmPantalla.TCuerda1.FontSize - 1
End Sub

Private Sub cmdLes2_Click()
FrmPantalla.TCuerda2.FontSize = FrmPantalla.TCuerda2.FontSize - 1
End Sub

Private Sub Command1_Click()
If FrmPantalla.Visible = True Then
    FrmPantalla.SetFocus
End If
If FrmGanador.Visible = True Then
    FrmGanador.SetFocus
End If
End Sub

Private Sub btnSiguiente_Click()
nPelea = nPelea + 1
Me.tNumPelea = Format(nPelea, "000")
Call cargarPelea(nPelea)
End Sub

Private Sub btnVolver_Click()
If confArenaIncluido = "no" Then
    Dim tiem As Integer
    tiem = tiempoaEntero(Me.tTiempo)
    tiem = tiem + confTiempoArena
    FrmPantalla.LTurno = pasarAHora(tiem)
End If
    
Unload FrmGanador
FrmPantalla.Show
FrmPantalla.SetFocus
End Sub
 
Private Sub Command2_Click()
FrmPantalla.QuitarMejores
End Sub

Private Sub Command3_Click()
frmMensaje.Visible = False
End Sub

Private Sub Command4_Click()
Call hablar(Me.tmen)
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
Cooper_skin Me

Red.LocalPort = 1103
'Establece piuerto ip del pc local
Me.tIp = Red.LocalIP & "/" & Red.LocalPort

Me.Top = frmMenu.ubicacion.Top + 350
Me.Left = frmMenu.ubicacion.Left

Call cargarPeleas

'SQL = "SELECT * FROM resumenPeleas where jugada ='no' order by orden asc"
'
'rs.Open SQL, cnn, adOpenStatic, adLockOptimistic
'If rs.RecordCount <= 0 Then
'    MsgBox "No existen peleas programadas", vbCritical
'    Exit Sub
'End If
'
'rs.MoveFirst
'For i = 1 To rs.RecordCount
'    Set li = listaGallos.ListItems.Add(, , rs("orden"))
'        li.SubItems(1) = rs("idPelea")
'        li.SubItems(2) = rs("Cuerda")
'        li.SubItems(3) = rs("placaCuerda")
'        li.SubItems(4) = rs("placaNacional")
'        li.SubItems(5) = rs("anillo")
'        li.SubItems(6) = "vs"
'        li.SubItems(7) = rs("Cuerda2")
'        li.SubItems(8) = rs("placaCuerda2")
'        li.SubItems(9) = rs("placaNacional2")
'        li.SubItems(10) = rs("anillo2")
'    rs.MoveNext
'Next
'
'numPeleas = rs.RecordCount
restantes = numPeleas
If numPeleas = 0 Then
   ' MsgBox "No hay peleas para ordenar", vbInformation, "Sin peleas"
    Unload Me
    Exit Sub
End If

'nPelea = 1
'Call cargarPelea(nPelea)

Me.tPeleasRestantes = Format(numPeleas, "000")
Me.tNumPelea = Format(nPelea, "000")

'rs.Close
End Sub

Public Function cargarPeleas()
Dim qry As New rdoQuery
Dim rs As rdoResultset

SQL = "SELECT * FROM resumenPeleas where jugada ='no' order by orden asc"
Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rs = qry.OpenResultset(rdOpenDynamic)
        If rs.RowCount <= 0 Then
            MsgBox "No existen peleas programadas", vbCritical
            Exit Function
        End If
        
        rs.MoveFirst
        listaGallos.ListItems.Clear
        For i = 1 To rs.RowCount
        Set li = listaGallos.ListItems.Add(, , rs("orden"))
            li.SubItems(1) = rs("idPelea")
            li.SubItems(2) = rs("Cuerda")
            li.SubItems(3) = IIf(IsNull(rs("placaCuerda")), "", rs("placaCuerda"))
            li.SubItems(4) = IIf(IsNull(rs("placaNacional")), "", rs("placaNacional"))
            li.SubItems(5) = IIf(IsNull(rs("anillo")), "", rs("anillo"))
            li.SubItems(6) = "vs"
            li.SubItems(7) = rs("Cuerda2")
            li.SubItems(8) = IIf(IsNull(rs("placaCuerda2")), "", rs("placaCuerda2"))
            li.SubItems(9) = IIf(IsNull(rs("placaNacional2")), "", rs("placaNacional2"))
            li.SubItems(10) = IIf(IsNull(rs("anillo2")), "", rs("anillo2"))
            rs.MoveNext
        Next
numPeleas = rs.RowCount
qry.Close
End Function

Private Function galloDuplicado(Placa As String) As Boolean
Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Select * from Gallos where placa='" & Placa & "'"
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        Set rst = qry.OpenResultset(rdOpenDynamic)
            If (rst.RowCount >= 1) Then
                galloDuplicado = True
            Else
                galloDuplicado = False
            End If
    qry.Close
End Function

Private Sub buscarPelea()
SQL = "SELECT * " & _
        "FROM Peleas WHERE idPelea = " & tSorteo.Text & " and orden = 0"
rs.Open SQL, cnn, adOpenStatic, adLockOptimistic

If rs.RecordCount >= 1 Then
    Me.tPlaca1.Caption = rs("placa1")
    Me.tPlaca2.Caption = rs("placa2")
    rs.Close
    Call buscarGallo1
    Call buscarGallo2
Else
    Me.tPlaca1 = "Sin placa"
    Me.TCuerda1 = "Sin cuerda"
    Me.tPeso1 = "Sin peso"
    Me.tPlaca2 = "Sin placa"
    Me.TCuerda2 = "Sin cuerda"
    Me.tPeso2 = "Sin peso"
    
    MsgBox "Este número de pelea no existe", vbCritical, "Sin pelea"
    
    'Me.tSorteo.SetFocus
    'seleccionarTexto tSorteo
    rs.Close
End If
End Sub

Private Sub buscarGallo1()
SQL = "SELECT * " & _
        "FROM Gallos WHERE placa = '" & tPlaca1 & "'"
rs.Open SQL, cnn, adOpenStatic, adLockOptimistic

If rs.RecordCount >= 1 Then
    Me.TCuerda1 = nombreCuerda(rs("idCuerda"))
    Me.tPeso1 = rs("peso")
End If
rs.Close
End Sub

Private Sub buscarGallo2()
SQL = "SELECT * " & _
        "FROM Gallos WHERE placa = '" & tPlaca2 & "'"
rs.Open SQL, cnn, adOpenStatic, adLockOptimistic

If rs.RecordCount >= 1 Then
    Me.TCuerda2 = nombreCuerda(rs("idCuerda"))
    Me.tPeso2 = rs("peso")
End If
rs.Close
End Sub

Private Sub tPlaca1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call buscarGallo1
End Sub

Private Sub tPlaca1_LostFocus()
Me.tPlaca1 = UCase(tPlaca1)
Call buscarGallo1
End Sub

Private Sub tPlaca2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call buscarGallo2
End Sub

Private Sub tPlaca2_LostFocus()
Me.tPlaca2 = UCase(tPlaca2)
Call buscarGallo2
End Sub

Private Function pasaGallo(pla As String) As Boolean
With listaGallos
    For i = 1 To .ListItems.Count
        If .ListItems.Item(i).SubItems(2) = pla Or .ListItems.Item(i).SubItems(5) = pla Then
            validarGallo = False
            Exit Function
        End If
    Next
    pasaGallo = True
End With
End Function

Private Sub tPosPelea_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub cargarPelea(numero As Integer)
SQL = "SELECT * FROM resumenPeleas where orden = " & numero & ""

Dim rs2 As New ADODB.Recordset
rs2.Open SQL, cnn, adOpenStatic, adLockOptimistic
If rs2.RecordCount <= 0 Then
    MsgBox "No hay pelea", vbCritical
    Exit Sub
End If
Me.tIdPelea = rs2("idPelea")
Me.tPlaca1 = IIf(IsNull(rs2("anillo")), "", rs2("anillo"))
Me.tPlaca2 = IIf(IsNull(rs2("anillo2")), "", rs2("anillo2"))
Me.TCuerda1 = rs2("Cuerda")
Me.TCuerda2 = rs2("Cuerda2")
Me.tPeso1 = rs2("peso")
Me.tPeso2 = rs2("peso2")
Me.tIdGallo1 = rs2("idGallo")
Me.tIdGallo2 = rs2("idGallo2")

Me.tCiudad1 = IIf(IsNull(rs2("ciudad")), "", rs2("ciudad"))
Me.tCiudad2 = IIf(IsNull(rs2("ciudad2")), "", rs2("ciudad2"))
Me.TColor1 = IIf(IsNull(rs2("colorPluma")), "", rs2("colorPluma"))
Me.TColor2 = IIf(IsNull(rs2("colorPluma2")), "", rs2("colorPluma2"))

rs2.Close

End Sub

Public Sub siguientePelea()
restantes = restantes - 1
If restantes > 0 Then
    nPelea = nPelea + 1
    Call cargarPelea(nPelea)
End If
End Sub

Private Sub hablar(mensaje As String)
Set speaking2 = New SpeechLib.SpVoice
Set speaking2.Voice = speaking2.GetVoices().Item(voz)
    speaking2.Rate = -1
    speaking2.Volume = 100
    
    m = speaking2.Speak(mensaje, SVSFlagsAsync)
End Sub

Private Sub listaGallos_DblClick()
nPelea = listaGallos.SelectedItem
Me.tNumPelea = nPelea
Call cargarPelea(nPelea)
End Sub

Private Sub opRedAct_Click()
If Me.opRedAct.Value = True Then
    On Error GoTo errorSub

    With Red
        .Close
        .Listen
    End With
    Me.tCliente = "Esperando"
    Exit Sub
errorSub:
    MsgBox "Error : " & Err.Description, vbCritical
End If
End Sub

Private Sub Red_Close()
'Finaliza la conexión
Me.tCliente.Caption = "Desconectado"
hayRed = False
Me.opRedDes.Value = True
Red.Close
End Sub

Private Sub Red_ConnectionRequest(ByVal requestID As Long)
tCliente.Caption = "Conectando"
If Red.State <> sckClosed Then
    Red.Close ' close
End If
Red.Accept requestID
    
tCliente.Caption = Red.RemoteHostIP
hayRed = True
End Sub

Private Sub Red_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "*** Error : " & Description & vbCrLf, vbCritical
Red_Close
End Sub

'*****************************************************************************
'MENSAJES ATRAVES DE LA RED
'*****************************************************************************

'Actualiza la pantalla remota
Public Sub actualizarPantalla()
mensajeRed = "1_" & Me.TCuerda1 & "_" & Me.TCuerda2 & "_10_0_60_3_5_si"
Call enviarMensaje(mensajeRed)
End Sub

'Inicia el reloj de la pelea
Public Sub iniciarReloj()
mensajeRed = "2"
Call enviarMensaje(mensajeRed)
End Sub

'Pausa el reloj de la pelea
Public Sub pausarReloj()
mensajeRed = "3"
Call enviarMensaje(mensajeRed)
End Sub

'Executa el reloj de arena derecho
Public Sub execArenaDerecho()
mensajeRed = "4"
Call enviarMensaje(mensajeRed)
End Sub

'Executa el reloj de arena izquierdo
Public Sub execArenaIzquierdo()
mensajeRed = "5"
Call enviarMensaje(mensajeRed)
End Sub

'Executa el reloj de espuelas
Public Sub execEspuelas()
mensajeRed = "6"
Call enviarMensaje(mensajeRed)
End Sub

'Establece el ganador de la pelea
Public Sub establecerGanador(win As String)
mensajeRed = "7_" & win
Call enviarMensaje(mensajeRed)
End Sub

'Establece un mensaje sobre la camara
Public Sub establecerMensaje(titulo As String, mensaje As String, size As String)
mensajeRed = "8_" & titulo & "_" & mensaje & "_" & size
Call enviarMensaje(mensajeRed)
End Sub

'Envia el mensaje a traves de la red
Private Sub enviarMensaje(men As String)
On Error GoTo errorSub
    Red.SendData men
Exit Sub
errorSub:
MsgBox "Error : " & Err.Description
Red_Close
End Sub
