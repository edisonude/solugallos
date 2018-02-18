VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmSorteoPeleaAutomatico 
   BackColor       =   &H00808080&
   Caption         =   "Generar separacion minima"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   15855
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView listCuerdasOriginales 
      Height          =   375
      Left            =   13200
      TabIndex        =   67
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "idCuerda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "idFrente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "idCiudad"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Número de intentos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   5415
      Begin VB.Timer timePrint 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   120
         Top             =   3240
      End
      Begin VB.CommandButton cmdBajarPeso 
         Caption         =   "Probar bajar peso a gallos libres"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   4200
         Width           =   3615
      End
      Begin VB.PictureBox picLine 
         Height          =   15
         Index           =   2
         Left            =   240
         ScaleHeight     =   15
         ScaleWidth      =   4935
         TabIndex        =   25
         Top             =   3720
         Width           =   4935
      End
      Begin VB.CommandButton cmdFinalizar 
         Caption         =   "Finalizar y Guardar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir sorteo previo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   3120
         Width           =   2415
      End
      Begin VB.PictureBox picLine 
         Height          =   15
         Index           =   1
         Left            =   240
         ScaleHeight     =   15
         ScaleWidth      =   4935
         TabIndex        =   24
         Top             =   3000
         Width           =   4935
      End
      Begin VB.TextBox tDetalles 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   1680
         Width           =   4935
      End
      Begin VB.PictureBox picLine 
         Height          =   15
         Index           =   0
         Left            =   240
         ScaleHeight     =   15
         ScaleWidth      =   4935
         TabIndex        =   21
         Top             =   1200
         Width           =   4935
      End
      Begin VB.CommandButton cmdIniciarSorteo 
         Caption         =   "Iniciar sorteo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   0
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox tIntentos 
         Alignment       =   2  'Center
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
         Left            =   720
         TabIndex        =   1
         Text            =   "1"
         Top             =   720
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   3
         Left            =   240
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":0000
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   4
         Left            =   240
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":007C
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   5
         Left            =   240
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":00E6
         TabIndex        =   26
         Top             =   3840
         Width           =   1455
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":0158
      Top             =   0
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   17400
      TabIndex        =   6
      Text            =   "limite"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
      Height          =   495
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":038C
      TabIndex        =   5
      Top             =   360
      Width           =   5295
   End
   Begin MSComctlLib.ListView listGallos 
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cueda"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Gallo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "peso"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ubicado"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "cuenta"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ListView listaPeleas 
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Gallo1"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Gallo2"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ListView listGallosPendientes 
      Height          =   375
      Left            =   11400
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cueda"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Gallo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "peso"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ubicado"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "cuenta"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ListView listFinalPeleas 
      Height          =   5535
      Left            =   5880
      TabIndex        =   11
      Top             =   1080
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9763
      View            =   3
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
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IdG1"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cuerda1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Gallo1"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Peso1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "idG2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cuerda2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Gallo2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Peso2"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView listGallosMejor 
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Gallo1"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Gallo2"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Información"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   5415
      Begin VB.CommandButton cmdUndoFrentes 
         Caption         =   "Restablecer Frentes"
         Height          =   495
         Left            =   1800
         TabIndex        =   69
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdOrganizarCiudades 
         Caption         =   "Organizar Ciudad"
         Height          =   375
         Left            =   240
         TabIndex        =   68
         Top             =   720
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   1
         Left            =   240
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":041A
         TabIndex        =   15
         Top             =   360
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel tGallos 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":04B0
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   2
         Left            =   3960
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":0510
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel tDiv 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":0572
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
      Height          =   255
      Index           =   6
      Left            =   5880
      OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":05D2
      TabIndex        =   27
      Top             =   720
      Width           =   2895
   End
   Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
      Height          =   255
      Index           =   7
      Left            =   5880
      OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":0660
      TabIndex        =   28
      Top             =   6720
      Width           =   2895
   End
   Begin MSComctlLib.ListView listGallosLibres 
      Height          =   3015
      Left            =   5880
      TabIndex        =   10
      Top             =   7080
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cueda"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Gallo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "peso"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ubicado"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "cuenta"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Frame frmComodin 
      BackColor       =   &H00808080&
      Caption         =   "Creación gallo comodín"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5880
      TabIndex        =   43
      Top             =   7080
      Visible         =   0   'False
      Width           =   10455
      Begin MSComctlLib.ListView listaCuerdas 
         Height          =   2055
         Left            =   1560
         TabIndex        =   44
         Top             =   840
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3625
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Id"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cuerda"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Propietario"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox tPeso 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox tPlacaCuerda 
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
         Left            =   1920
         TabIndex        =   55
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox tPlacaNacional 
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
         Left            =   1920
         TabIndex        =   54
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox tAnillo 
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
         Left            =   1920
         TabIndex        =   53
         Top             =   2400
         Width           =   1695
      End
      Begin VB.ComboBox tColorPlumas 
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
         Left            =   5400
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox tColorPico 
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
         Left            =   5400
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   1455
         Width           =   2175
      End
      Begin VB.ComboBox tTipoCresta 
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
         Left            =   5400
         Sorted          =   -1  'True
         TabIndex        =   50
         Top             =   2415
         Width           =   2175
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
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
         Left            =   7920
         TabIndex        =   49
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox tColorPatas 
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
         Left            =   5400
         Sorted          =   -1  'True
         TabIndex        =   48
         Top             =   1935
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancelar"
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
         Left            =   7920
         TabIndex        =   47
         Top             =   1920
         Width           =   1815
      End
      Begin VB.PictureBox picSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   6600
         Picture         =   "frmSorteoPeleaAutomatico.frx":06D2
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   46
         Top             =   480
         Width           =   390
      End
      Begin VB.TextBox tCuerda 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   45
         Top             =   480
         Width           =   4935
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   16
         Left            =   360
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":0DD4
         TabIndex        =   57
         Top             =   1470
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   17
         Left            =   360
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":0E44
         TabIndex        =   58
         Top             =   990
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   19
         Left            =   3840
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":0EA4
         TabIndex        =   59
         Top             =   975
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   21
         Left            =   360
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":0F14
         TabIndex        =   60
         Top             =   1950
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   22
         Left            =   360
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":0F88
         TabIndex        =   61
         Top             =   2430
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   23
         Left            =   3840
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":0FEC
         TabIndex        =   62
         Top             =   1470
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   24
         Left            =   3840
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":1058
         TabIndex        =   63
         Top             =   2430
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   26
         Left            =   3840
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":10C6
         TabIndex        =   64
         Top             =   1950
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   12
         Left            =   360
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":1134
         TabIndex        =   65
         Top             =   480
         Width           =   855
      End
      Begin VB.Label tIdCuerda 
         Height          =   255
         Left            =   7080
         TabIndex        =   66
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame frmOpcionesLibre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Opciones para el gallo libre"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   5880
      TabIndex        =   29
      Top             =   7080
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   40
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton cmdCrearComodín 
         Caption         =   "Crearle un rival comodín"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   39
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CommandButton cmdEliminarLibre 
         Caption         =   "Eliminar este gallo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   38
         Top             =   2040
         Width           =   2775
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   8
         Left            =   720
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":1198
         TabIndex        =   30
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   9
         Left            =   1440
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":1206
         TabIndex        =   31
         Top             =   720
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   10
         Left            =   1440
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":126A
         TabIndex        =   32
         Top             =   1200
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   11
         Left            =   1440
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":12CA
         TabIndex        =   33
         Top             =   960
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel glCuerda 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":132E
         TabIndex        =   34
         Top             =   720
         Width           =   6855
      End
      Begin ACTIVESKINLibCtl.SkinLabel glPeso 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":1392
         TabIndex        =   35
         Top             =   1200
         Width           =   6855
      End
      Begin ACTIVESKINLibCtl.SkinLabel glAnillo 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":13F2
         TabIndex        =   36
         Top             =   960
         Width           =   6855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   15
         Left            =   840
         OleObjectBlob   =   "frmSorteoPeleaAutomatico.frx":1456
         TabIndex        =   37
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label glIndex 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3480
         TabIndex        =   42
         Top             =   360
         Width           =   615
      End
      Begin VB.Label glId 
         Caption         =   "ID"
         Height          =   255
         Left            =   2400
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Lista de gallos pendientes"
      Height          =   195
      Left            =   1920
      TabIndex        =   13
      Top             =   5280
      Width           =   1830
   End
End
Attribute VB_Name = "frmSorteoPeleaAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nGallos As Integer
Dim listaGallos() As String
Dim gallosXpeso() As String

Dim lPesos(1 To 33) As String

Dim noGallosPendientes As Integer
Dim lGallosPendientes() As String

Dim noGallosLibres As Integer
Dim lGallosLibres() As String

Dim gallosContra As Integer
Dim gallosContrincantes() As String

'Total de los gallos por peso que se estan analizando
Dim Total As Integer
Dim pesoAct As String

'Mejor resultado
Dim noGallosMejor As Integer

Private Sub cmdAgregar_Click()
If Me.tPlacaCuerda = "" And Me.tAnillo = "" And Me.tPlacaNacional = "" Then
    MsgBox "Debe digitar por lo menos una identificación del gallo", vbCritical, "Sin id de gallo"
    Me.tPlacaCuerda.SetFocus
    Exit Sub
End If

End Sub

Private Sub cmdBajarPeso_Click()
'Limpio los detalles
Me.tDetalles = ""

If Me.listGallosLibres.ListItems.Count = 0 Then
    MsgBox "No existen gallos libres para rebajarles el peso", vbInformation
    Exit Sub
End If

Dim nInt As Integer
noGallosMejor = 9999
'Iniciamos los intentos
For nInt = 1 To Val(Me.tIntentos)
    'Reseteo todo
    Call reset

    For i = 1 To 33
        pesoAct = lPesos(i)
        Call ordenar2
    Next
    'Verificamos los gallos libres
    If Me.listGallosLibres.ListItems.Count = 0 Then
        noGallosMejor = 0
        Exit For
    Else
        If noGallosMejor > Me.listGallosLibres.ListItems.Count Then
            noGallosMejor = Me.listGallosLibres.ListItems.Count
            'Copiar datos
            Me.listGallosMejor.ListItems.Clear
            For j = 1 To Me.listaPeleas.ListItems.Count
                Set li = Me.listGallosMejor.ListItems.Add(, , Me.listaPeleas.ListItems(j))
                    li.SubItems(1) = Me.listaPeleas.ListItems(j).SubItems(1)
            Next
        End If
    End If
Next

If noGallosMejor > 0 Then
    'Copiar datos
    Me.listaPeleas.ListItems.Clear
    For j = 1 To Me.listGallosMejor.ListItems.Count
        Set li = Me.listaPeleas.ListItems.Add(, , Me.listGallosMejor.ListItems(j))
            li.SubItems(1) = Me.listGallosMejor.ListItems(j).SubItems(1)
    Next
End If

Call verPeleas

'Crea resumen
Me.tDetalles.Text = "Se encontraron " & Me.listFinalPeleas.ListItems.Count & " peleas." & vbCrLf
If Me.listGallosLibres.ListItems.Count > 0 Then
    Me.tDetalles = Me.tDetalles & vbCrLf & "Quedaron " & Me.listGallosLibres.ListItems.Count & " gallos libres, faltantes por rival."
Else
    Me.tDetalles = Me.tDetalles & vbCrLf & "No quedaron gallos libres, se organizarón todas las peleas"
End If
End Sub

Private Sub cmdCancelar_Click()
Me.frmOpcionesLibre.Visible = False
End Sub

Private Sub cmdCrearComodín_Click()
frmComodin.Visible = True
frmComodin.ZOrder 0

Me.tPeso = Me.glPeso

'Ancho de la lista de cuerdas
With listaCuerdas
    .ColumnHeaders(1).Width = .Width * 0
    .ColumnHeaders(2).Width = .Width * 0.55
    .ColumnHeaders(3).Width = .Width * 0.4
    .ZOrder 0
End With

Call cargarcuerdas
'cargar colores
Call cargarColores(Me.tColorPlumas)
Call cargarColores(Me.tColorPico)
Call cargarColores(Me.tColorPatas)
Call cargarCrestas(Me.tTipoCresta)
End Sub

Private Sub cmdEliminarLibre_Click()
If MsgBox("¿Está seguro de eliminar definitivamente este gallo?", vbQuestion + vbYesNo) = vbYes Then
    'Imprimir registro de confirmación
    Call imprimirRegistro(Val(Me.glId))
    
    SQL = "Delete * from Gallos where idGallo=" & Me.glId & ""
    Call guardarRDO
    
    Me.listGallosLibres.ListItems.Remove (Val(glIndex))
    Me.frmOpcionesLibre.Visible = False
End If
End Sub

Private Sub cmdFinalizar_Click()
Dim conPelea As Integer
With listaPeleas
    For i = 1 To .ListItems.Count
        conPelea = HallaConsecutivo("Select Max(idPelea) As NConsecutivo from Peleas")
        
        If (trabajandoCon = "SoluPollos" And conPelea = 1) Then
            conPelea = conPelea + 1000
        End If
        
        SQL = "insert into Peleas values(" & conPelea & "," & .ListItems.Item(i) & _
        ",'" & .ListItems.Item(i).SubItems(1) & "','" & Format(Date, "dd/mm/yyyy") & "',0,'',0,0,0,'no')"
            
        Call guardarRDO
    Next
End With

Call menGuardadoExitoso
End Sub

Private Sub cmdImprimir_Click()
'Borro todas las peleas temporales

SQL = "Delete * from PeleasTmp"
Call guardarRDO
Dim conPelea As Integer
With listaPeleas
    For i = 1 To .ListItems.Count
        conPelea = HallaConsecutivo("Select Max(idPelea) As NConsecutivo from PeleasTmp")
        
        SQL = "insert into PeleasTmp values(" & conPelea & "," & .ListItems.Item(i) & _
        ",'" & .ListItems.Item(i).SubItems(1) & "','" & Format(Date, "dd/mm/yyyy") & "',0,'',0,0,0,'no')"
            
        Call guardarRDO
    Next
End With
Me.timePrint.Enabled = True
End Sub

Private Sub cmdIniciarSorteo_Click()
'Limpio los detalles
Me.tDetalles = ""

Dim nInt As Integer
noGallosMejor = 9999
'Iniciamos los intentos
For nInt = 1 To Val(Me.tIntentos)
    'Reseteo todo
    Call Form_Load

    For i = 1 To 33
        pesoAct = lPesos(i)
        Call ordenar
    Next
    'Verificamos los gallos libres
    If Me.listGallosLibres.ListItems.Count = 0 Then
        noGallosMejor = 0
        Exit For
    Else
        If noGallosMejor > Me.listGallosLibres.ListItems.Count Then
            noGallosMejor = Me.listGallosLibres.ListItems.Count
            'Copiar datos
            Me.listGallosMejor.ListItems.Clear
            For j = 1 To Me.listaPeleas.ListItems.Count
                Set li = Me.listGallosMejor.ListItems.Add(, , Me.listaPeleas.ListItems(j))
                    li.SubItems(1) = Me.listaPeleas.ListItems(j).SubItems(1)
            Next
        End If
    End If
Next

If noGallosMejor > 0 Then
    'Copiar datos
    Me.listaPeleas.ListItems.Clear
    For j = 1 To Me.listGallosMejor.ListItems.Count
        Set li = Me.listaPeleas.ListItems.Add(, , Me.listGallosMejor.ListItems(j))
            li.SubItems(1) = Me.listGallosMejor.ListItems(j).SubItems(1)
    Next
End If

Call verPeleas

'Crea resumen
Me.tDetalles.Text = "Se encontraron " & Me.listFinalPeleas.ListItems.Count & " peleas." & vbCrLf
If Me.listGallosLibres.ListItems.Count > 0 Then
    Me.tDetalles = Me.tDetalles & vbCrLf & "Quedaron " & Me.listGallosLibres.ListItems.Count & " gallos libres, faltantes por rival."
Else
    Me.tDetalles = Me.tDetalles & vbCrLf & "No quedaron gallos libres, se organizarón todas las peleas"
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command7_Click()

End Sub

Private Sub cmdOrganizarCiudades_Click()
Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Select * from Cuerdas where idCuerda <> -1"
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        Set rst = qry.OpenResultset(rdOpenDynamic)
            listCuerdasOriginales.ListItems.Clear
            While rst.EOF = False
                Set li = listCuerdasOriginales.ListItems.Add(, , rst("idCuerda"))
                    li.SubItems(1) = rst("Frente")
                    li.SubItems(2) = rst("Ciudad")
                
                rst.MoveNext
            Wend
        qry.Close
Call asignarFrentexCiudad
End Sub
Private Function UndoFrentes2()

End Function
Private Function UndoFrentes()


    For j = 1 To Me.listCuerdasOriginales.ListItems.Count
  
        SQL = "Update Cuerdas set Frente = " & Me.listCuerdasOriginales.ListItems(j).SubItems(1) & " where idCuerda = " & Me.listCuerdasOriginales.ListItems(j)
        guardarRDO

    Next
End Function
Private Function asignarFrentexCiudad()
Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Update Cuerdas set Frente = Ciudad where idCuerda <> -1"
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        qry.Execute
        qry.Close
        
End Function

Private Sub cmdUndoFrentes_Click()
Call UndoFrentes
End Sub

Private Sub Form_Load()
Cooper_skin Me

'Establecer ancho para las listas
With listFinalPeleas
    .ColumnHeaders(1).Width = .Width * 0
    .ColumnHeaders(2).Width = .Width * 0.3
    .ColumnHeaders(3).Width = .Width * 0.1
    .ColumnHeaders(4).Width = .Width * 0.09
    .ColumnHeaders(5).Width = .Width * 0
    .ColumnHeaders(6).Width = .Width * 0.3
    .ColumnHeaders(7).Width = .Width * 0.1
    .ColumnHeaders(8).Width = .Width * 0.09
End With

With listGallosLibres
    .ColumnHeaders(1).Width = .Width * 0.6
    .ColumnHeaders(2).Width = .Width * 0.39
    .ColumnHeaders(3).Width = .Width * 0
    .ColumnHeaders(4).Width = .Width * 0
    .ColumnHeaders(5).Width = .Width * 0
End With

noGallosPendientes = 0
ReDim gallosContrincantes(3, 1)
ReDim gallosXpeso(0, 5)
ReDim lGallosLibres(0, 5)
ReDim lGallosPendientes(0, 5)

'Cargo numero de gallos
nGallos = totalGallos
Me.tGallos = nGallos
Me.tDiv = nGallos \ 2

'Iniciando pesos
lPesos(1) = "2,10"
lPesos(2) = "2,11"
lPesos(3) = "2,12"
lPesos(4) = "2,13"
lPesos(5) = "2,14"
lPesos(6) = "2,15"
lPesos(7) = "3"
lPesos(8) = "3,1"
lPesos(9) = "3,2"
lPesos(10) = "3,3"
lPesos(11) = "3,4"
lPesos(12) = "3,5"
lPesos(13) = "3,6"
lPesos(14) = "3,7"
lPesos(15) = "3,8"
lPesos(16) = "3,9"
lPesos(17) = "3,10"
lPesos(18) = "3,11"
lPesos(19) = "3,12"
lPesos(20) = "3,13"
lPesos(21) = "3,14"
lPesos(22) = "3,15"
lPesos(23) = "4"
lPesos(24) = "4,1"
lPesos(25) = "4,2"
lPesos(26) = "4,3"
lPesos(27) = "4,4"
lPesos(28) = "4,5"
lPesos(29) = "4,6"
lPesos(30) = "4,7"
lPesos(31) = "4,8"
lPesos(32) = "4,9"
lPesos(33) = "4,10"

Me.listGallos.ListItems.Clear
Me.listaPeleas.ListItems.Clear
Me.listGallosPendientes.ListItems.Clear
Me.listGallosLibres.ListItems.Clear
Me.listFinalPeleas.ListItems.Clear
End Sub

Private Sub reset()
noGallosPendientes = 0
ReDim gallosContrincantes(3, 1)
ReDim gallosXpeso(0, 5)
ReDim lGallosLibres(0, 5)
ReDim lGallosPendientes(0, 5)

'Cargo numero de gallos
nGallos = totalGallos
Me.tGallos = nGallos

'Iniciando pesos
lPesos(1) = "2,10"
lPesos(2) = "2,11"
lPesos(3) = "2,12"
lPesos(4) = "2,13"
lPesos(5) = "2,14"
lPesos(6) = "2,15"
lPesos(7) = "3"
lPesos(8) = "3,1"
lPesos(9) = "3,2"
lPesos(10) = "3,3"
lPesos(11) = "3,4"
lPesos(12) = "3,5"
lPesos(13) = "3,6"
lPesos(14) = "3,7"
lPesos(15) = "3,8"
lPesos(16) = "3,9"
lPesos(17) = "3,10"
lPesos(18) = "3,11"
lPesos(19) = "3,12"
lPesos(20) = "3,13"
lPesos(21) = "3,14"
lPesos(22) = "3,15"
lPesos(23) = "4"
lPesos(24) = "4,1"
lPesos(25) = "4,2"
lPesos(26) = "4,3"
lPesos(27) = "4,4"
lPesos(28) = "4,5"
lPesos(29) = "4,6"
lPesos(30) = "4,7"
lPesos(31) = "4,8"
lPesos(32) = "4,9"
lPesos(33) = "4,10"

'Redimensionando la lista
'ReDim listaGallos(1 To nGallos, 1 To nGallos, 1 To nGallos)



Me.listGallos.ListItems.Clear
Me.listaPeleas.ListItems.Clear
Me.listGallosPendientes.ListItems.Clear
'Me.listGallosLibres.ListItems.Clear
End Sub

Private Function totalGallos() As Integer
Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Select count(idGallo) as nGallos from Gallos where idGallo<>-1 and idGallo not in (select idGallo1 from Peleas) and idGallo not in (select idGallo2 from Peleas)"
    Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    totalGallos = rst("nGallos")
    qry.Close
End Function

Private Function totalGallosxPeso(peso As String) As Integer
Dim qry As New rdoQuery
Dim rst As rdoResultset

'SQL = "Select count(idGallo) as nGallos from Gallos where peso='" & peso & "'"
SQL = "SELECT count(g.idGallo) as nGallos from Gallos g where g.idGallo not in (SELECT Gallos.idGallo " & _
    "FROM Gallos inner JOIN Peleas ON (Gallos.idGallo = Peleas.idGallo2) or " & _
    "(Gallos.idGallo = Peleas.idGallo1)) and g.peso='" & peso & "'"
    
    Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    totalGallosxPeso = rst("nGallos")
    qry.Close
End Function

Private Sub cargarGallosxPeso(peso As String)
Total = totalGallosxPeso(peso)
If Total <> 0 Then
    'Redimensiono la lista para los gallos del peso
    ReDim gallosXpeso(1 To Total, 5)
    'Lleno la lista con los gallos de ese peso
    Dim no As Integer
    Dim frente As Integer
    Dim qry As New rdoQuery
    Dim rst As rdoResultset
    no = 1
    Dim li As ListItem
    'SQL = "Select * from Gallos where peso='" & peso & "' order by idCuerdaPelea"
    SQL = "SELECT g.* from Gallos g where g.idGallo not in (SELECT Gallos.idGallo " & _
    "FROM Gallos inner JOIN Peleas ON (Gallos.idGallo = Peleas.idGallo2) or " & _
    "(Gallos.idGallo = Peleas.idGallo1)) and g.peso='" & peso & "'"
    
    Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    While rst.EOF = False
        'Preguntar si es frente
        frente = cuerdaEsFrente(rst("idCuerdaPelea"))
        If frente = 0 Then
            frente = rst("idCuerdaPelea")
        End If
            
        gallosXpeso(no, 1) = frente
        gallosXpeso(no, 2) = rst("idGallo")
        gallosXpeso(no, 3) = rst("peso")
        gallosXpeso(no, 4) = "no"
        gallosXpeso(no, 5) = "0"
        Call agregarGalloLista(Me.listGallos, gallosXpeso, no)
        no = no + 1
        rst.MoveNext
    Wend
    'MsgBox peso & " : " & rst.RowCount
    qry.Close
    
    Call generarPesos
Else
    ReDim gallosXpeso(0, 5)
    Me.listGallos.ListItems.Clear
End If
End Sub

Private Sub cargarGallosxPeso2(peso As String)
Total = totalGallosxPeso(peso)
If Total <> 0 Then
    'Redimensiono la lista para los gallos del peso
    ReDim gallosXpeso(1 To Total, 5)
    'Lleno la lista con los gallos de ese peso
    Dim no As Integer
    Dim frente As Integer
    Dim qry As New rdoQuery
    Dim rst As rdoResultset
    Dim saltos As Integer
    saltos = 0
    no = 1
    Dim li As ListItem
    SQL = "Select * from Gallos where peso='" & peso & "' order by idCuerdaPelea"
    Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    While rst.EOF = False
        'Validar que no este
        For i = 1 To Me.listaPeleas.ListItems.Count
            If rst("idGallo") = Me.listaPeleas.ListItems(i) Or rst("idGallo") = Me.listaPeleas.ListItems(i).SubItems(1) Then
                
                saltos = saltos + 1
                GoTo sigue
            End If
        Next
    
        'Preguntar si es frente
        frente = cuerdaEsFrente(rst("idCuerdaPelea"))
        If frente = 0 Then
            frente = rst("idCuerdaPelea")
        End If
            
        gallosXpeso(no, 1) = frente
        gallosXpeso(no, 2) = rst("idGallo")
        gallosXpeso(no, 3) = rst("peso")
        gallosXpeso(no, 4) = "no"
        gallosXpeso(no, 5) = "0"
        Call agregarGalloLista(Me.listGallos, gallosXpeso, no)
        no = no + 1
sigue:
        rst.MoveNext
    Wend
    'MsgBox peso & " : " & rst.RowCount
    qry.Close
    

    
    Total = Total - saltos
    If saltos > 0 Then
    gallosXpeso = setFilasMatriz(gallosXpeso, Total, 5)
    End If
    
    Call generarPesos
    
'    If no = 1 And gallosXpeso(1, 1) = "" Then
'    total = 0
'        ReDim gallosXpeso(0, 5)
'    Else
'    Call generarPesos
'    End If
Else
    ReDim gallosXpeso(0, 5)
    Me.listGallos.ListItems.Clear
End If
End Sub



Private Sub ordenar()
'Variables de validacion
Dim cuerdaAct As Integer
Dim ubicadoAct As String
Dim cuerdaPru As Integer
Dim ubicadoPru As String

'Cargo los gallos del peso correspondiente
Call cargarGallosxPeso(pesoAct)

'Pregunto si hay gallos pendientes
If noGallosPendientes > 0 Then
    'Verifico si los gallos se deben liberar
    If Not pesoValido(pesoAct, lGallosPendientes(1, 3)) Then
       Call liberarGallosPendientes
    Else
        'Procesar gallos pendientes
        gallosXpeso = agregarFilasMatrizPreserve(gallosXpeso, noGallosPendientes, 5)
        'Agrego los gallos pendientes
        Dim pos As Integer
        pos = Total + 1
        For i = 1 To noGallosPendientes
            gallosXpeso(pos, 1) = lGallosPendientes(i, 1)
            gallosXpeso(pos, 2) = lGallosPendientes(i, 2)
            gallosXpeso(pos, 3) = lGallosPendientes(i, 3)
            gallosXpeso(pos, 4) = lGallosPendientes(i, 4)
            gallosXpeso(pos, 5) = "999"
            Call agregarGalloLista(Me.listGallos, gallosXpeso, pos)
            pos = pos + 1
        Next
        Call ordenarMatrizDesc(gallosXpeso, 5)
    End If
    GoTo continue
Else
continue:
    Dim j, k As Integer
    For j = 1 To UBound(gallosXpeso) - 1
        cuerdaAct = gallosXpeso(j, 1)
        ubicadoAct = gallosXpeso(j, 4)
        
        'Si no he ubicado el gallo actual
        If ubicadoAct = "no" Then
            gallosContra = 0
            Erase gallosContrincantes
        
            For k = j + 1 To UBound(gallosXpeso)
                cuerdaPru = gallosXpeso(k, 1)
                ubicadoPru = gallosXpeso(k, 4)
                
                'Verificamos si es posible que halla pelea
                If ubicadoPru = "no" And cuerdaPru <> cuerdaAct Then
                    gallosContra = gallosContra + 1
                    ReDim Preserve gallosContrincantes(3, gallosContra)
                    
                    'Creo la lista de contrincantes
                    gallosContrincantes(1, gallosContra) = gallosContra
                    gallosContrincantes(2, gallosContra) = gallosXpeso(k, 2)
                    gallosContrincantes(3, gallosContra) = k
                End If
            Next
            
            'Verifico si existen contrincantes
            If gallosContra > 0 Then
                Dim contra As Integer
                contra = random(gallosContra)
                
                Call crearPelea(Val(gallosXpeso(j, 2)), Val(gallosContrincantes(2, contra)))
                gallosXpeso(j, 4) = "si"
                gallosXpeso(gallosContrincantes(3, contra), 4) = "si"
            End If
        End If
    Next
    
    'Verifico si quedaron gallos pendientes
     Call validarGallosPendientes
    
    
    Call refrescarLista
End If
End Sub

Private Sub ordenar2()
'Variables de validacion
Dim cuerdaAct As Integer
Dim ubicadoAct As String
Dim pesoActu As String
Dim cuerdaPru As Integer
Dim ubicadoPru As String
Dim pesoPruu As String

Dim pos As Integer

Dim li As ListItem
Dim idP, idP2 As Integer

'If pesoAct = "4,1" Then
'    MsgBox "p"
'End If

'Cargo los gallos del peso correspondiente
Call cargarGallosxPeso2(pesoAct)


'Agrego los gallos libres
If pesoAct <> "4,10" Then
For i = 1 To Me.listGallosLibres.ListItems.Count
    If i > Me.listGallosLibres.ListItems.Count Then Exit For
    idP = bucarPosVector(lPesos, pesoAct)
    If lPesos(idP + 1) = Me.listGallosLibres.ListItems(i).SubItems(2) Then
        'Agrego el gallo libre a los gallos actuales del menor peso
        gallosXpeso = agregarFilasMatrizPreserve(gallosXpeso, 1, 5)
        'Agrego los gallos pendientes
        
        pos = Total + 1
        
            gallosXpeso(pos, 1) = Me.listGallosLibres.ListItems(i)
            gallosXpeso(pos, 2) = Me.listGallosLibres.ListItems(i).SubItems(1)
            gallosXpeso(pos, 3) = Me.listGallosLibres.ListItems(i).SubItems(2)
            gallosXpeso(pos, 4) = Me.listGallosLibres.ListItems(i).SubItems(3)
            gallosXpeso(pos, 5) = Me.listGallosLibres.ListItems(i).SubItems(4)
            Call agregarGalloLista(Me.listGallos, gallosXpeso, pos)
            Total = pos
        
        Call ordenarMatrizDesc(gallosXpeso, 5)

        Me.listGallosLibres.ListItems.Remove (i)
        i = i - 1
        If Me.listGallosLibres.ListItems.Count = 0 Then Exit For
    End If
Next
End If

Call refrescarLista

'Pregunto si hay gallos pendientes
If noGallosPendientes > 0 Then
    'Verifico si los gallos se deben liberar
    If Not pesoValido(pesoAct, lGallosPendientes(1, 3)) Then
       Call liberarGallosPendientes
    Else
        'Procesar gallos pendientes
        gallosXpeso = agregarFilasMatrizPreserve(gallosXpeso, noGallosPendientes, 5)
        'Agrego los gallos pendientes
        pos = Total + 1
        For i = 1 To noGallosPendientes
            gallosXpeso(pos, 1) = lGallosPendientes(i, 1)
            gallosXpeso(pos, 2) = lGallosPendientes(i, 2)
            gallosXpeso(pos, 3) = lGallosPendientes(i, 3)
            gallosXpeso(pos, 4) = lGallosPendientes(i, 4)
            gallosXpeso(pos, 5) = "999"
            Call agregarGalloLista(Me.listGallos, gallosXpeso, pos)
            pos = pos + 1
        Next
        Call ordenarMatrizDesc(gallosXpeso, 5)
    End If
    GoTo continue
Else
continue:
    Dim j, k As Integer
    For j = 1 To UBound(gallosXpeso) - 1
        cuerdaAct = gallosXpeso(j, 1)
        ubicadoAct = gallosXpeso(j, 4)
        pesoActu = gallosXpeso(j, 3)
        
        'Si no he ubicado el gallo actual
        If ubicadoAct = "no" Then
            gallosContra = 0
            Erase gallosContrincantes
        
            For k = j + 1 To UBound(gallosXpeso)
                cuerdaPru = gallosXpeso(k, 1)
                ubicadoPru = gallosXpeso(k, 4)
                pesoPruu = gallosXpeso(k, 3)
                
                'Verificamos si es posible que halla pelea
                If ubicadoPru = "no" And cuerdaPru <> cuerdaAct And pesoPelea(pesoActu, pesoPruu) Then
                    gallosContra = gallosContra + 1
                    ReDim Preserve gallosContrincantes(3, gallosContra)
                    
                    'Creo la lista de contrincantes
                    gallosContrincantes(1, gallosContra) = gallosContra
                    gallosContrincantes(2, gallosContra) = gallosXpeso(k, 2)
                    gallosContrincantes(3, gallosContra) = k
                End If
            Next
            
            'Verifico si existen contrincantes
            If gallosContra > 0 Then
                Dim contra As Integer
                contra = random(gallosContra)
                
                Call crearPelea(Val(gallosXpeso(j, 2)), Val(gallosContrincantes(2, contra)))
                gallosXpeso(j, 4) = "si"
                gallosXpeso(gallosContrincantes(3, contra), 4) = "si"
            End If
        End If
    Next
    
    'Verifico si quedaron gallos pendientes
     Call validarGallosPendientes
    
    
    Call refrescarLista
End If
End Sub

Private Sub crearPelea(gallo1 As Integer, gallo2 As Integer)
Set li = listaPeleas.ListItems.Add(, , gallo1)
    li.SubItems(1) = gallo2
End Sub


Private Sub refrescarLista()
Dim i As Integer
Me.listGallos.ListItems.Clear
For i = 1 To UBound(gallosXpeso)
    Call agregarGalloLista(Me.listGallos, gallosXpeso, i)
Next
End Sub

Private Sub recorrerMatriz(matriz)
Dim z As Integer
Dim x As Integer

For z = 1 To UBound(matriz, 2)
    For x = 1 To UBound(matriz)
        MsgBox "(" & x & "," & z & ") = " & matriz(x, z)
    Next
Next
End Sub

Private Sub generarPesos()
Dim f As Integer
Dim f2 As Integer
Dim idC As Integer
Dim peso As Integer

For f = 1 To UBound(gallosXpeso)
    peso = 0
    idC = gallosXpeso(f, 1)
    For f2 = 1 To UBound(gallosXpeso)
        If idC = gallosXpeso(f2, 1) Then
            peso = peso + 1
        End If
    Next
    gallosXpeso(f, 5) = peso
Next
Call ordenarMatrizDesc(gallosXpeso, 5)
Call refrescarLista
End Sub

Public Sub ordenarMatrizDesc(matriz() As String, col As Integer)
Dim vectAux(1, 5) As String
For i = 1 To UBound(matriz) - 1
    For j = i + 1 To UBound(matriz)
        If matriz(i, col) < matriz(j, col) Then
            vectAux(1, 1) = matriz(i, 1)
            vectAux(1, 2) = matriz(i, 2)
            vectAux(1, 3) = matriz(i, 3)
            vectAux(1, 4) = matriz(i, 4)
            vectAux(1, 5) = matriz(i, 5)
            
            matriz(i, 1) = matriz(j, 1)
            matriz(i, 2) = matriz(j, 2)
            matriz(i, 3) = matriz(j, 3)
            matriz(i, 4) = matriz(j, 4)
            matriz(i, 5) = matriz(j, 5)
            
            matriz(j, 1) = vectAux(1, 1)
            matriz(j, 2) = vectAux(1, 2)
            matriz(j, 3) = vectAux(1, 3)
            matriz(j, 4) = vectAux(1, 4)
            matriz(j, 5) = vectAux(1, 5)
        End If
    Next
Next
End Sub

Private Sub validarGallosPendientes()
Dim f As Integer
Dim f2 As Integer
Dim idC As Integer
Dim peso As Integer

noGallosPendientes = 0
Me.listGallosPendientes.ListItems.Clear
For f = 1 To UBound(gallosXpeso)
    If "no" = gallosXpeso(f, 4) Then
        noGallosPendientes = noGallosPendientes + 1
    End If
Next

If noGallosPendientes > 0 Then
    ReDim lGallosPendientes(1 To noGallosPendientes, 5)
    Dim cuenta As Integer
    cuenta = 1
    For f = 1 To UBound(gallosXpeso)
        If "no" = gallosXpeso(f, 4) Then
            lGallosPendientes(cuenta, 1) = gallosXpeso(f, 1)
            lGallosPendientes(cuenta, 2) = gallosXpeso(f, 2)
            lGallosPendientes(cuenta, 3) = gallosXpeso(f, 3)
            lGallosPendientes(cuenta, 4) = gallosXpeso(f, 4)
            lGallosPendientes(cuenta, 5) = gallosXpeso(f, 5)
            
            Call agregarGalloLista(Me.listGallosPendientes, lGallosPendientes, cuenta)
            cuenta = cuenta + 1
        End If
    Next
End If
'noGallosPendientes = noGallosPendientes + Me.listGallosPendientes.ListItems.Count
End Sub

Private Sub liberarGallosPendientes()
Dim f As Integer
Dim f2 As Integer
Dim idC As Integer
Dim peso As Integer

noGallosPendientes = 0
For f = 1 To UBound(lGallosPendientes)
    Call agregarGalloLista(Me.listGallosLibres, lGallosPendientes, f)
Next

Me.listGallosPendientes.ListItems.Clear
Erase lGallosPendientes
End Sub

Public Function pesoValido(pesoMayor As String, pesoMenor As String) As Boolean
Dim idP, idP2 As Integer
idP = bucarPosVector(lPesos, pesoMenor)
idP2 = bucarPosVector(lPesos, pesoMayor)
If idP2 - idP = 1 Then
    pesoValido = True
Else
    pesoValido = False
End If
End Function

Public Function pesoPelea(peso1 As String, peso2 As String) As Boolean
Dim idP, idP2, dif  As Integer
idP = bucarPosVector(lPesos, pesoMenor)
idP2 = bucarPosVector(lPesos, pesoMayor)
dif = idP - idP2
If dif > -2 And dif < 2 Then
    pesoPelea = True
Else
    pesoPelea = False
End If
End Function

Private Sub agregarGalloLista(lista As ListView, matriz() As String, nu As Integer)
Set li = lista.ListItems.Add(, , matriz(nu, 1))
    li.SubItems(1) = matriz(nu, 2)
    li.SubItems(2) = matriz(nu, 3)
    li.SubItems(3) = matriz(nu, 4)
    li.SubItems(4) = matriz(nu, 5)
End Sub

Private Function consultarGallo(idGallo As Integer) As String
Dim qry As New rdoQuery
Dim rst As rdoResultset
Dim cadena As String

SQL = "Select * from consultaGallos where idGallo=" & idGallo & ""
    Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    cadena = idGallo & "|" & rst("Cuerda") & "|" & rst("anillo") & "|" & rst("peso")
    qry.Close
    consultarGallo = cadena
End Function

Private Function cuerdaEsFrente(id As Integer) As Integer
Dim qry As New rdoQuery
Dim rst As rdoResultset
Dim cadena As String

SQL = "Select * from Cuerdas where idCuerda=" & id & ""
    Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    
    cuerdaEsFrente = rst("Frente")
    qry.Close
End Function

'Lista las peleas encontradas
Private Sub verPeleas()
'Recorro la lista de peleas
Dim i As Integer
Dim idGallo1 As Integer
Dim idGallo2 As Integer
Dim respuesta As String
Dim sep() As String

listFinalPeleas.ListItems.Clear

For i = 1 To Me.listaPeleas.ListItems.Count
    idGallo1 = Me.listaPeleas.ListItems.Item(i)
    idGallo2 = Me.listaPeleas.ListItems.Item(i).SubItems(1)
    respuesta = consultarGallo(idGallo1)
    sep = Split(respuesta, "|")
    
    Set li = listFinalPeleas.ListItems.Add(, , sep(0))
    li.SubItems(1) = sep(1)
    li.SubItems(2) = sep(2)
    li.SubItems(3) = sep(3)
    
    respuesta = consultarGallo(idGallo2)
    sep = Split(respuesta, "|")
    
    li.SubItems(4) = sep(0)
    li.SubItems(5) = sep(1)
    li.SubItems(6) = sep(2)
    li.SubItems(7) = sep(3)
Next
End Sub

Private Sub listGallosLibres_DblClick()
Dim res As String
Dim sep() As String

Me.glIndex = listGallosLibres.SelectedItem.Index
res = consultarGallo(listGallosLibres.ListItems(Val(glIndex)).SubItems(1))
sep = Split(res, "|")

Me.glAnillo = sep(2)
Me.glCuerda = sep(1)
Me.glPeso = sep(3)
Me.glId = sep(0)

Me.frmOpcionesLibre.Visible = True
Me.frmOpcionesLibre.ZOrder 0
End Sub

Private Sub imprimirRegistro(id As Integer)
strArchivo = pathBD

Dim oAcces As Access.Application
Set oAcces = New Access.Application

oAcces.OpenCurrentDatabase strArchivo, False, keyBD
oAcces.Visible = False
oAcces.DoCmd.OpenReport "det_gallo_del", acViewPreview, , "idGallo=" & id

oAcces.DoCmd.PrintOut acPrintAll
oAcces.CloseCurrentDatabase
oAcces.Quit
Set oAcces = Nothing
End Sub

Private Sub imprimirLista()
strArchivo = pathBD

Dim oAcces As Access.Application
Set oAcces = New Access.Application

oAcces.OpenCurrentDatabase strArchivo, False, keyBD
oAcces.Visible = False
oAcces.DoCmd.OpenReport "inf_peleas_temporales", acViewPreview

oAcces.DoCmd.PrintOut acPrintAll
oAcces.CloseCurrentDatabase
oAcces.Quit
Set oAcces = Nothing
End Sub


Public Sub cargarcuerdas()
'Cargo las cuerdas
Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Select * from Cuerdas"
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        Set rst = qry.OpenResultset(rdOpenDynamic)
            listaCuerdas.ListItems.Clear
            While rst.EOF = False
                If rst("idCuerda") <> -1 Then
                    Set li = listaCuerdas.ListItems.Add(, , rst("idCuerda"))
                        li.SubItems(1) = rst("Cuerda")
                        li.SubItems(2) = rst("Propietario")
                End If
                rst.MoveNext
            Wend
        qry.Close
End Sub

Private Sub timePrint_Timer()
Call imprimirLista
Me.timePrint.Enabled = False
End Sub
