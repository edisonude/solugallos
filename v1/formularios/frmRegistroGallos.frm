VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegistroGallos 
   BackColor       =   &H00555555&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00555555&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   360
      ScaleHeight     =   135
      ScaleWidth      =   11655
      TabIndex        =   45
      Top             =   360
      Width           =   11655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   8055
      Left            =   360
      TabIndex        =   16
      Top             =   360
      Width           =   11655
      Begin VB.CommandButton cmdPrintPeleas 
         Caption         =   "Imprimir peleas"
         Height          =   255
         Left            =   9600
         TabIndex        =   47
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprimir"
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
         Left            =   4920
         TabIndex        =   7
         Top             =   7440
         Width           =   2175
      End
      Begin VB.PictureBox picSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   6360
         Picture         =   "frmRegistroGallos.frx":0000
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   37
         Top             =   360
         Width           =   390
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
         Left            =   7440
         TabIndex        =   8
         Top             =   7440
         Width           =   2175
      End
      Begin VB.CommandButton btnGuardar 
         Caption         =   "Guardar"
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
         Left            =   2520
         TabIndex        =   6
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Frame frmDatos 
         BackColor       =   &H00808080&
         Caption         =   "Información de los gallos"
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
         Height          =   6255
         Left            =   345
         TabIndex        =   18
         Top             =   840
         Width           =   11055
         Begin VB.TextBox tMes 
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
            Left            =   6600
            TabIndex        =   50
            Text            =   "4"
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox tJaula 
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
            Left            =   4440
            TabIndex        =   4
            Top             =   1800
            Width           =   855
         End
         Begin VB.CheckBox chkComodin 
            Caption         =   "Comodín"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7560
            TabIndex        =   46
            Top             =   360
            Width           =   1095
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
            Left            =   1800
            TabIndex        =   1
            Text            =   "3,"
            Top             =   360
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
            Left            =   1800
            TabIndex        =   3
            Top             =   840
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
            Left            =   1800
            TabIndex        =   9
            Top             =   1320
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
            Left            =   1800
            TabIndex        =   10
            Top             =   1800
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
            Left            =   5280
            Sorted          =   -1  'True
            TabIndex        =   2
            Text            =   "COLORADO"
            Top             =   360
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
            Left            =   5280
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   855
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
            Left            =   5280
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   2295
            Visible         =   0   'False
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
            Left            =   9120
            TabIndex        =   5
            Top             =   840
            Width           =   1815
         End
         Begin VB.PictureBox Picture2 
            Height          =   15
            Left            =   240
            ScaleHeight     =   15
            ScaleWidth      =   10695
            TabIndex        =   21
            Top             =   2280
            Width           =   10695
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
            Left            =   5280
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   1335
            Width           =   2175
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   8880
            ScaleHeight     =   375
            ScaleWidth      =   2055
            TabIndex        =   19
            Top             =   360
            Width           =   2055
            Begin VB.Label lAddObservacion 
               BackStyle       =   0  'Transparent
               Caption         =   "Agregar observación"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000F7FD&
               Height          =   255
               Left            =   120
               MouseIcon       =   "frmRegistroGallos.frx":0702
               MousePointer    =   99  'Custom
               TabIndex        =   20
               Top             =   0
               Width           =   1935
            End
         End
         Begin MSComctlLib.ListView listaGallos 
            Height          =   3420
            Left            =   240
            TabIndex        =   22
            Top             =   2700
            Width           =   10545
            _ExtentX        =   18600
            _ExtentY        =   6033
            SortKey         =   1
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
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "id"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Peso"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Color"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "P cuerda"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "P nacional"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Anillo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Comodín"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Patas"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Cresta"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "observacion"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "pico"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "jaula"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "mes"
               Object.Width           =   2540
            EndProperty
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   5
            Left            =   240
            OleObjectBlob   =   "frmRegistroGallos.frx":0A0C
            TabIndex        =   23
            Top             =   870
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   6
            Left            =   240
            OleObjectBlob   =   "frmRegistroGallos.frx":0A7C
            TabIndex        =   24
            Top             =   390
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   300
            Index           =   7
            Left            =   240
            OleObjectBlob   =   "frmRegistroGallos.frx":0ADC
            TabIndex        =   25
            Top             =   2400
            Width           =   2775
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   1
            Left            =   3720
            OleObjectBlob   =   "frmRegistroGallos.frx":0B6C
            TabIndex        =   26
            Top             =   375
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   2
            Left            =   8160
            OleObjectBlob   =   "frmRegistroGallos.frx":0BDC
            TabIndex        =   27
            Top             =   1440
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel tGallosCuerda 
            Height          =   375
            Left            =   8160
            OleObjectBlob   =   "frmRegistroGallos.frx":0C4E
            TabIndex        =   28
            Top             =   1800
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   4
            Left            =   240
            OleObjectBlob   =   "frmRegistroGallos.frx":0CA8
            TabIndex        =   29
            Top             =   1350
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   8
            Left            =   240
            OleObjectBlob   =   "frmRegistroGallos.frx":0D1C
            TabIndex        =   30
            Top             =   1830
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   3
            Left            =   3720
            OleObjectBlob   =   "frmRegistroGallos.frx":0D80
            TabIndex        =   31
            Top             =   870
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   10
            Left            =   3720
            OleObjectBlob   =   "frmRegistroGallos.frx":0DEC
            TabIndex        =   32
            Top             =   2310
            Visible         =   0   'False
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   11
            Left            =   9600
            OleObjectBlob   =   "frmRegistroGallos.frx":0E5A
            TabIndex        =   33
            Top             =   1440
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel tTotalGallos 
            Height          =   375
            Left            =   9600
            OleObjectBlob   =   "frmRegistroGallos.frx":0ECA
            TabIndex        =   34
            Top             =   1800
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   9
            Left            =   3720
            OleObjectBlob   =   "frmRegistroGallos.frx":0F24
            TabIndex        =   35
            Top             =   1350
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   12
            Left            =   3720
            OleObjectBlob   =   "frmRegistroGallos.frx":0F92
            TabIndex        =   48
            Top             =   1830
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   13
            Left            =   5880
            OleObjectBlob   =   "frmRegistroGallos.frx":0FF4
            TabIndex        =   49
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label tIdGallo 
            Caption         =   "-1"
            Height          =   255
            Left            =   2880
            TabIndex        =   36
            Top             =   360
            Width           =   615
         End
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
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   4935
      End
      Begin VB.CheckBox chkImprimir 
         Caption         =   "Imprimir al guardar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9240
         TabIndex        =   17
         Top             =   7080
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   0
         Left            =   360
         OleObjectBlob   =   "frmRegistroGallos.frx":1052
         TabIndex        =   38
         Top             =   360
         Width           =   855
      End
      Begin MSComctlLib.ListView listaCuerdas 
         Height          =   3375
         Left            =   1320
         TabIndex        =   39
         Top             =   720
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
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
      Begin VB.Frame frmObservacion 
         Caption         =   "Observación al gallo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   2400
         TabIndex        =   40
         Top             =   1560
         Visible         =   0   'False
         Width           =   8895
         Begin VB.TextBox tObservacion 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   43
            Top             =   480
            Width           =   8055
         End
         Begin VB.CommandButton cmdGuardarObservacion 
            Caption         =   "Guardar"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   42
            Top             =   2640
            Width           =   2175
         End
         Begin VB.CommandButton cmdSalirObservacion 
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
            Height          =   375
            Left            =   8400
            TabIndex        =   41
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Label tIdCuerda 
         Height          =   255
         Left            =   6840
         TabIndex        =   44
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox picLabel 
      BackColor       =   &H00555555&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   11655
      TabIndex        =   14
      Top             =   0
      Width           =   11655
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Información de los gallos por cuerda"
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
         TabIndex        =   15
         Top             =   0
         Width           =   5535
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmRegistroGallos.frx":10B6
      Top             =   360
   End
End
Attribute VB_Name = "frmRegistroGallos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Función Api SendMessageLong ( para desplegar la lista en forma automática )
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Dim li As ListItem
Dim encabezado As String
Dim numeroGallo As Integer

Dim gallosCuerda As Integer
Dim totalGallos As Integer


Private Sub btnGuardar_Click()
If tIdCuerda.Caption = "" Or listaGallos.ListItems.Count <= 0 Then
    menFaltanDatos
    Exit Sub
End If

Dim conGallo As Integer
With listaGallos
    For i = 1 To .ListItems.Count
        conGallo = .ListItems.Item(i)
        If conGallo = -1 Then
            conGallo = HallaConsecutivo("Select Max(idGallo) As NConsecutivo from Gallos")
            
            If (trabajandoCon = "SoluPollos" And conGallo = 1) Then
                conGallo = conGallo + 1000
            End If
    
            SQL = "INSERT INTO Gallos values(" & conGallo & "," & tIdCuerda & "," & tIdCuerda & ",'" & .ListItems.Item(i).SubItems(3) & _
                    "','" & .ListItems.Item(i).SubItems(4) & "','" & .ListItems.Item(i).SubItems(5) & _
                    "','" & .ListItems.Item(i).SubItems(1) & "','" & .ListItems.Item(i).SubItems(2) & _
                    "','" & .ListItems.Item(i).SubItems(10) & "','" & .ListItems.Item(i).SubItems(7) & _
                    "','" & .ListItems.Item(i).SubItems(8) & "','" & Format(Date, "dd/mm/yyyy") & _
                    "','" & .ListItems.Item(i).SubItems(9) & "','" & .ListItems.Item(i).SubItems(6) & _
                    "'," & .ListItems.Item(i).SubItems(11) & "," & .ListItems.Item(i).SubItems(12) & ")"
            
            If .ListItems.Item(i).SubItems(2) <> "" Then
                Call addColor(.ListItems.Item(i).SubItems(2))
            End If
            
            If .ListItems.Item(i).SubItems(6) <> "" Then
                Call addColor(.ListItems.Item(i).SubItems(6))
            End If
            
            If .ListItems.Item(i).SubItems(7) <> "" Then
                Call addColor(.ListItems.Item(i).SubItems(7))
            End If
            
            If .ListItems.Item(i).SubItems(8) <> "" Then
                Call addCresta(.ListItems.Item(i).SubItems(8))
            End If
        Else
            If galloExiste(conGallo) Then
                SQL = "UPDATE Gallos set placaCuerda='" & .ListItems.Item(i).SubItems(3) & _
                    "',idCuerdaPelea='" & tIdCuerda & _
                    "',placaNacional='" & .ListItems.Item(i).SubItems(4) & _
                    "',anillo='" & .ListItems.Item(i).SubItems(5) & _
                    "',peso='" & .ListItems.Item(i).SubItems(1) & _
                    "',colorPluma='" & .ListItems.Item(i).SubItems(2) & _
                    "',colorPico='" & .ListItems.Item(i).SubItems(10) & _
                    "',colorPatas='" & .ListItems.Item(i).SubItems(7) & _
                    "',tipoCresta='" & .ListItems.Item(i).SubItems(8) & _
                    "',fecha='" & Format(Date, "dd/mm/yyyy") & _
                    "',observaciones='" & .ListItems.Item(i).SubItems(9) & _
                    "',comodin='" & .ListItems.Item(i).SubItems(6) & _
                    "' where idGallo=" & conGallo & ""
            Else
                SQL = "INSERT INTO Gallos values(" & conGallo & "," & tIdCuerda & "," & tIdCuerda & ",'" & .ListItems.Item(i).SubItems(3) & _
                    "','" & .ListItems.Item(i).SubItems(4) & "','" & .ListItems.Item(i).SubItems(5) & _
                    "','" & .ListItems.Item(i).SubItems(1) & "','" & .ListItems.Item(i).SubItems(2) & _
                    "','" & .ListItems.Item(i).SubItems(10) & "','" & .ListItems.Item(i).SubItems(7) & _
                    "','" & .ListItems.Item(i).SubItems(8) & "','" & Format(Date, "dd/mm/yyyy") & _
                    "','" & .ListItems.Item(i).SubItems(9) & "','" & .ListItems.Item(i).SubItems(6) & "')"
            End If
        End If
        
        If Not guardarRDO Then
            Call menGuardadoFallo
        End If
    Next
    Call menGuardadoExitoso
    If chkImprimir.Value = 1 Then
        Call imprimirRegistro(Me.tIdCuerda)
    End If
    
    Unload Me
    frmRegistroGallos.Show
End With
End Sub

Private Sub imprimirRegistro(id As Integer)
strArchivo = pathBD

Dim oAcces As Access.Application
Set oAcces = New Access.Application

oAcces.OpenCurrentDatabase strArchivo, False, keyBD
oAcces.Visible = False
oAcces.DoCmd.OpenReport "inf_gallos_cuerda", acViewPreview, , "idCuerda=" & id

oAcces.DoCmd.PrintOut acPrintAll
oAcces.CloseCurrentDatabase
oAcces.Quit
Set oAcces = Nothing
End Sub

Private Sub imprimirPeleas(id As Integer)
strArchivo = pathBD

Dim q(2) As String

q(1) = "idCuerda=" & id
q(1) = "idCuerda2=" & id


Dim oAcces As Access.Application
Set oAcces = New Access.Application

oAcces.OpenCurrentDatabase strArchivo, False, keyBD
oAcces.Visible = False
oAcces.DoCmd.OpenReport "consolidadoFinal", acViewPreview, , "idCuerda=" & id & " OR " & "idCuerda2=" & id

oAcces.DoCmd.PrintOut acPrintAll
oAcces.CloseCurrentDatabase
oAcces.Quit
Set oAcces = Nothing
End Sub

Private Sub btnImprimir_Click()
If Me.tIdCuerda = "" Then
    MsgBox "Debe seleccionar la cuerda que quiere reimprimir", vbInformation
    Exit Sub
End If
Call imprimirRegistro(Me.tIdCuerda)
End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub cmdAgregar_Click()
If Me.tPlacaCuerda = "" And Me.tAnillo = "" And Me.tPlacaNacional = "" Then
    MsgBox "Debe digitar por lo menos una identificación del gallo", vbCritical, "Sin id de gallo"
    Me.tPlacaCuerda.SetFocus
    Exit Sub
End If

Set li = listaGallos.ListItems.Add(, , tIdGallo)
    li.SubItems(1) = Me.tPeso
    li.SubItems(2) = Me.tColorPlumas
    li.SubItems(3) = Me.tPlacaCuerda
    li.SubItems(4) = Me.tPlacaNacional
    li.SubItems(5) = Me.tAnillo
    li.SubItems(6) = IIf(chkComodin.Value = 1, "si", "no")
    li.SubItems(7) = Me.tColorPatas
    li.SubItems(8) = Me.tTipoCresta
    li.SubItems(9) = Me.tObservacion
    li.SubItems(10) = Me.tColorPico
    li.SubItems(11) = Me.tJaula
    li.SubItems(12) = Me.tMes
    
   Call limpiarCampos
   
   Me.tPeso.SetFocus
   gallosCuerda = gallosCuerda + 1
   totalGallos = totalGallos + 1
   
   tGallosCuerda = gallosCuerda
   tTotalGallos = totalGallos
   
   Me.tPeso.Text = "3,"
   Me.tColorPlumas = "COLORADO"
   Me.tMes = 3
   'Me.tAnillo.Text = totalGallos + 1 + 100
   
End Sub

Private Sub cmdGuardarObservacion_Click()
frmObservacion.Visible = False
End Sub

Private Sub cmdPrintPeleas_Click()
If Me.tIdCuerda = "" Then
    MsgBox "Debe seleccionar la cuerda que quiere reimprimir", vbInformation
    Exit Sub
End If
Call imprimirPeleas(Me.tIdCuerda)
End Sub

Private Sub cmdSalirObservacion_Click()
If MsgBox("Si sale de esta forma no se guardrá la observación." & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo) = vbYes Then
    tObservacion = ""
    Me.frmObservacion.Visible = False
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Cooper_skin Me

Me.Top = frmMenu.ubicacion.Top + 350
Me.Left = frmMenu.ubicacion.Left

'Ancho de la lista de cuerdas
With listaCuerdas
    .ColumnHeaders(1).Width = .Width * 0
    .ColumnHeaders(2).Width = .Width * 0.55
    .ColumnHeaders(3).Width = .Width * 0.4
    .ZOrder 0
End With

'Ancho de la lista de gallos
With listaGallos
    .ColumnHeaders(1).Width = .Width * 0 'id
    .ColumnHeaders(2).Width = .Width * 0.08 'peso
    .ColumnHeaders(3).Width = .Width * 0.15 'color
    .ColumnHeaders(4).Width = .Width * 0.15 'placa cuerda
    .ColumnHeaders(5).Width = .Width * 0.15 'placa nacional
    .ColumnHeaders(6).Width = .Width * 0 'anillo
    .ColumnHeaders(7).Width = .Width * 0.15  'comodin
    .ColumnHeaders(8).Width = .Width * 0 'patas
    .ColumnHeaders(9).Width = .Width * 0 'cresta
    .ColumnHeaders(10).Width = .Width * 0   'observacion
    .ColumnHeaders(11).Width = .Width * 0   'pico
    .ColumnHeaders(12).Width = .Width * 0.15   'jaula
    .ColumnHeaders(13).Width = .Width * 0.15 'mes
End With

'Total de galllos
totalGallos = HallaConsecutivo("Select Max(idGallo) As NConsecutivo from Gallos") - 1
Me.tTotalGallos = totalGallos

Call cargarcuerdas
'cargar colores
Call cargarColores(Me.tColorPlumas)
Call cargarColores(Me.tColorPico)
Call cargarColores(Me.tColorPatas)
Call cargarCrestas(Me.tTipoCresta)

frmObservacion.ZOrder 0
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
Private Function galloExiste(id As Integer) As Boolean
Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Select * from Gallos where idGallo=" & id & ""
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        Set rst = qry.OpenResultset(rdOpenDynamic)
            If (rst.RowCount >= 1) Then
                galloExiste = True
            Else
                galloExiste = False
            End If
    qry.Close
End Function
Private Function galloPelea(id As Integer) As Boolean
Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Select * from Peleas where idGallo1=" & id & " or idGallo2=" & id & " and jugada='no'"
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        Set rst = qry.OpenResultset(rdOpenDynamic)
            If (rst.RowCount >= 1) Then
                galloPelea = True
            Else
                galloPelea = False
            End If
    qry.Close
End Function

Private Sub limpiarCampos()
Me.tPeso = ""
Me.tColorPlumas = ""
Me.tPlacaCuerda = ""
Me.tPlacaNacional = ""
Me.tAnillo = ""
Me.tColorPico = ""
Me.tColorPatas = ""
Me.tTipoCresta = ""
Me.tObservacion = ""
Me.tIdGallo = "-1"
Me.chkComodin.Value = 0
End Sub

Private Sub Label2_Click()

End Sub

Private Sub lAddObservacion_Click()
On Error Resume Next
frmObservacion.Visible = IIf(frmObservacion.Visible = True, False, True)
tObservacion.SetFocus
End Sub

Private Sub listaCuerdas_DblClick()
With listaCuerdas
    Me.tIdCuerda = .ListItems(.SelectedItem.Index)
    Me.tCuerda = .ListItems(.SelectedItem.Index).SubItems(1)
    Me.listaCuerdas.Visible = False
End With
End Sub

Private Sub listaCuerdas_SelChange()
soloUnaSeleccion listaCuerdas
End Sub

Private Sub listaGallos_DblClick()
Dim ind As Integer
With listaGallos
    ind = .SelectedItem.Index
    
    If galloPelea(.ListItems(ind)) Then
        MsgBox "Este gallo se encuentra en una pelea no jugada y no es posible modificarlo", vbInformation
        Exit Sub
    End If
        
    Me.tIdGallo = .ListItems(ind)
    Me.tPeso = .ListItems(ind).SubItems(1)
    Me.tColorPlumas = .ListItems(ind).SubItems(2)
    Me.tPlacaCuerda = .ListItems(ind).SubItems(3)
    Me.tPlacaNacional = .ListItems(ind).SubItems(4)
    Me.tAnillo = .ListItems(ind).SubItems(5)
    Me.chkComodin.Value = IIf(.ListItems(ind).SubItems(6) = "si", 1, 0)
    Me.tColorPatas = .ListItems(ind).SubItems(7)
    Me.tTipoCresta = .ListItems(ind).SubItems(8)
    Me.tObservacion = .ListItems(ind).SubItems(9)
    Me.tColorPico = .ListItems(ind).SubItems(10)
    Me.tJaula = .ListItems(ind).SubItems(11)
    Me.tMes = .ListItems(ind).SubItems(12)
    .ListItems.Remove (ind)
End With

If tIdGallo <> -1 Then
    SQL = "Delete * from Gallos where idGallo=" & tIdGallo & ""
    Call guardarRDO
End If

gallosCuerda = gallosCuerda - 1
totalGallos = totalGallos - 1

tGallosCuerda = gallosCuerda
tTotalGallos = totalGallos
End Sub

Private Sub listaGallos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    Dim gallo As Integer
    gallo = Me.listaGallos.ListItems(listaGallos.SelectedItem.Index)
    If gallo <> -1 Then
        SQL = "Delete * from Gallos where idGallo=" & gallo & ""
        Call guardarRDO
        numeroGallo = numeroGallo - 1
    End If
    listaGallos.ListItems.Remove (listaGallos.SelectedItem.Index)
End If
End Sub

Private Sub picSearch_Click()
If Me.listaCuerdas.Visible = True Then
    Me.listaCuerdas.Visible = False
Else
    Me.listaCuerdas.Visible = True
End If
End Sub

Private Sub tPlacaCuerda_GotFocus()
Me.tPlacaCuerda.ToolTipText = ""
Me.tPlacaCuerda.BackColor = vbWhite
End Sub

Private Sub tTipoCresta_GotFocus()
SendMessageLong tTipoCresta.hwnd, &H14F, True, 0
End Sub
Private Sub tColorPatas_GotFocus()
SendMessageLong tColorPatas.hwnd, &H14F, True, 0
End Sub

Private Sub tColorPico_GotFocus()
SendMessageLong tColorPico.hwnd, &H14F, True, 0
End Sub

Private Sub tColorPlumas_GotFocus()
SendMessageLong tColorPlumas.hwnd, &H14F, True, 0
End Sub

Private Sub tCuerda_Change()
Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Select * from Cuerdas WHERE Cuerda like '%" & Me.tCuerda & "%' and idCuerda<>-1"
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        Set rst = qry.OpenResultset(rdOpenDynamic)
            listaCuerdas.ListItems.Clear
            While rst.EOF = False
                Set li = listaCuerdas.ListItems.Add(, , rst("idCuerda"))
                    li.SubItems(1) = rst("Cuerda")
                    li.SubItems(2) = rst("Propietario")
                rst.MoveNext
            Wend
        qry.Close
End Sub

Private Sub tCuerda_GotFocus()
Me.listaCuerdas.Visible = True
seleccionarTexto tCuerda
End Sub

Private Sub tIdCuerda_Change()
Dim qry As New rdoQuery
Dim rst As rdoResultset

Call limpiarCampos
Me.frmDatos.Enabled = True
Me.listaGallos.ListItems.Clear

SQL = "Select * from Gallos where idCuerdaPelea=" & tIdCuerda & " and fecha Between #" & Format(DateAdd("d", -1, Date), "mm/dd/yyyy") & "# and #" & Format(DateAdd("d", 1, Date), "mm/dd/yyyy") & "#"
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        Set rst = qry.OpenResultset(rdOpenDynamic)
            listaGallos.ListItems.Clear
            While rst.EOF = False
                Set li = listaGallos.ListItems.Add(, , rst("idGallo"))
                    li.SubItems(1) = IIf(IsNull(rst("peso")), "", rst("peso"))
                    li.SubItems(2) = IIf(IsNull(rst("colorPluma")), "", rst("colorPluma"))
                    li.SubItems(3) = IIf(IsNull(rst("placaCuerda")), "", rst("placaCuerda"))
                    li.SubItems(4) = IIf(IsNull(rst("placaNacional")), "", rst("placaNacional"))
                    li.SubItems(5) = IIf(IsNull(rst("anillo")), "", rst("anillo"))
                    li.SubItems(6) = IIf(IsNull(rst("comodin")), "", rst("comodin"))
                    li.SubItems(7) = IIf(IsNull(rst("colorPatas")), "", rst("colorPatas"))
                    li.SubItems(8) = IIf(IsNull(rst("tipoCresta")), "", rst("tipoCresta"))
                    li.SubItems(9) = IIf(IsNull(rst("observaciones")), "", rst("observaciones"))
                    li.SubItems(10) = IIf(IsNull(rst("colorPico")), "", rst("colorPico"))
                rst.MoveNext
            Wend
        gallosCuerda = rst.RowCount
        Me.tGallosCuerda = gallosCuerda
        qry.Close
        
        
        Me.tPeso.Text = "3,"
        Me.tColorPlumas = "COLORADO"
        'Me.tAnillo.Text = totalGallos + 1 + 100
        Me.tJaula = Me.obtenerProximaJaula
        
        Me.tPeso.SetFocus
        Me.tPeso.SelStart = 2
End Sub

Public Function obtenerProximaJaula() As Integer
obtenerProximaJaula = Me.obtenerJaulaUltimoGalloRegistrado() + 1
End Function

Private Sub tPeso_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
End Sub

Private Sub tPlacaCuerda_LostFocus()
Me.tPlacaCuerda = UCase(tPlacaCuerda)

If Me.tPlacaCuerda = "" Then Exit Sub
Call buscarGallo(tPlacaCuerda, "", Val(tIdCuerda))

If Me.placaPeleaDuplicada(Me.tPlacaCuerda.Text) Then
    Me.tPlacaCuerda.ToolTipText = "Ya existe otro gallo con esta misma placa"
    Me.tPlacaCuerda.BackColor = vbRed
End If
End Sub
Private Sub tPlacaNacional_LostFocus()
If Me.tPlacaNacional = "" Then Exit Sub
Me.tPlacaNacional = UCase(tPlacaNacional)

If placaDuplicada(tPlacaNacional) And tIdGallo = -1 Then
'    If MsgBox("Ya existe otro gallo con este mismo número de placa nacional: " & tPlacaNacional & "." & vbCrLf & "¿Desea cargar los datos?", vbQuestion + vbYesNo) = vbYes Then
'        Call buscarGallo("", tPlacaNacional, -1)
'    End If
    MsgBox "Ya existe otro gallo con esta misma placa nacional", vbCritical, "Gallo duplicado"
    Me.tPlacaNacional = ""
    Me.tPlacaNacional.SetFocus
End If
End Sub
Private Sub tAnillo_LostFocus()
Me.tAnillo = UCase(tAnillo)
End Sub
Private Sub tColorPlumas_LostFocus()
Me.tColorPlumas = UCase(tColorPlumas)
End Sub
Private Sub tColorPico_LostFocus()
Me.tColorPico = UCase(tColorPico)
End Sub
Private Sub tColorPatas_LostFocus()
Me.tColorPatas = UCase(tColorPatas)
End Sub

Private Sub tTipoCresta_LostFocus()
Me.tTipoCresta = UCase(tTipoCresta)
End Sub


Private Function placaDuplicada(Placa As String) As Boolean
Dim qry As New rdoQuery
Dim rst As rdoResultset

placaDuplicada = False

SQL = "Select * from Gallos where placaNacional='" & Placa & "'"
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        Set rst = qry.OpenResultset(rdOpenDynamic)
            If (rst.RowCount >= 1) Then
                placaDuplicada = True
                GoTo salirFuncion
            End If
            
            Dim i As Integer
            For i = 1 To listaGallos.ListItems.Count
                If (Placa = listaGallos.ListItems(i).SubItems(4)) Then
                    placaDuplicada = True
                    GoTo salirFuncion
                End If
            Next
        
salirFuncion:
    qry.Close
End Function

Public Sub buscarGallo(pCuerda As String, pNacional As String, cuerda As Integer)
Dim qry As New rdoQuery
Dim rst As rdoResultset
Set qry.ActiveConnection = RDOCONEXION
    If cuerda = -1 Then
        SQL = "Select * from Gallos where placaNacional='" & pNacional & "'"
    Else
        SQL = "Select * from Gallos where idCuerdaProp=" & cuerda & " and placaCuerda='" & pCuerda & "'"
    End If
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    If rst.RowCount > 0 Then
        Me.tIdGallo = IIf(IsNull(rst("idGallo")), "", rst("idGallo"))
        tPeso = IIf(IsNull(rst("peso")), "", rst("peso"))
        Me.tColorPlumas = IIf(IsNull(rst("colorPluma")), "", rst("colorPluma"))
        Me.tPlacaCuerda = IIf(IsNull(rst("placaCuerda")), "", rst("placaCuerda"))
        Me.tPlacaNacional = IIf(IsNull(rst("placaNacional")), "", rst("placaNacional"))
        Me.tAnillo = IIf(IsNull(rst("anillo")), "", rst("anillo"))
        Me.tColorPico = IIf(IsNull(rst("colorPico")), "", rst("colorPico"))
        Me.tColorPatas = IIf(IsNull(rst("colorPatas")), "", rst("colorPatas"))
        Me.tTipoCresta = IIf(IsNull(rst("tipoCresta")), "", rst("tipoCresta"))
        Me.tObservacion = IIf(IsNull(rst("observaciones")), "", rst("observaciones"))
    End If
    qry.Close
End Sub

'Determina si una placa se encuentra ya registrada por otro gallo.
'Puede no ser un error, pero se usa como advertencia.

Public Function placaPeleaDuplicada(Placa As String) As Boolean
Dim qry As New rdoQuery
Dim rst As rdoResultset
Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = "Select count(idGallo) as cuenta from Gallos where placaCuerda='" & Placa & "'"
    Set rst = qry.OpenResultset(rdOpenDynamic)
    If rst.RowCount > 0 Then
        If (CInt(rst("cuenta")) > 0) Then
           placaPeleaDuplicada = True
        Else
            placaPeleaDuplicada = False
        End If
    End If
    qry.Close
End Function

Public Function obtenerJaulaUltimoGalloRegistrado()
Dim qry As New rdoQuery
Dim rst As rdoResultset
Set qry.ActiveConnection = RDOCONEXION

SQL = "SELECT jaula FROM Gallos where idGallo = (Select Max(idGallo) As NConsecutivo from Gallos)"

    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    If rst.RowCount > 0 Then
        obtenerJaulaUltimoGalloRegistrado = rst("jaula")
        Exit Function
    End If
    obtenerJaulaUltimoGalloRegistrado = 0
    qry.Close
End Function

