VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H00404040&
   Caption         =   "Admin Gallera - Pollos"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   12960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picResize 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   11175
      Picture         =   "frmMenu.frx":0000
      ScaleHeight     =   690
      ScaleWidth      =   3420
      TabIndex        =   11
      Top             =   450
      Width           =   3420
      Begin VB.Label iResHeight 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Height          =   555
         Left            =   2685
         TabIndex        =   15
         Top             =   75
         Width           =   540
      End
      Begin VB.Label iResTop 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Height          =   525
         Left            =   1830
         TabIndex        =   14
         Top             =   75
         Width           =   585
      End
      Begin VB.Label iResWidth 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Height          =   510
         Left            =   870
         TabIndex        =   13
         Top             =   105
         Width           =   780
      End
      Begin VB.Label iResLeft 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Height          =   510
         Left            =   105
         TabIndex        =   12
         Top             =   90
         Width           =   585
      End
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   0
      Picture         =   "frmMenu.frx":7BE2
      ScaleHeight     =   1680
      ScaleWidth      =   1800
      TabIndex        =   10
      Top             =   0
      Width           =   1800
   End
   Begin VB.CommandButton cmdLibre 
      Caption         =   "Libre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   9
      Top             =   6840
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ANUNCIAR"
      Height          =   450
      Left            =   1440
      TabIndex        =   8
      Top             =   9480
      Width           =   2595
   End
   Begin ACTIVESKINLibCtl.SkinLabel lMainTitle 
      Height          =   855
      Left            =   1815
      OleObjectBlob   =   "frmMenu.frx":11B0C
      TabIndex        =   6
      Top             =   450
      Width           =   10140
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   15
      OleObjectBlob   =   "frmMenu.frx":11BAA
      Top             =   285
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00555555&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8805
      Left            =   -45
      Picture         =   "frmMenu.frx":11DDE
      ScaleHeight     =   8805
      ScaleWidth      =   4290
      TabIndex        =   0
      Top             =   1800
      Width           =   4290
      Begin VB.Image Image1 
         Height          =   840
         Left            =   120
         MouseIcon       =   "frmMenu.frx":1E73C
         MousePointer    =   99  'Custom
         Picture         =   "frmMenu.frx":1E88E
         Top             =   7680
         Width           =   870
      End
      Begin VB.Label imMenu 
         BackStyle       =   0  'Transparent
         Height          =   795
         Index           =   3
         Left            =   75
         TabIndex        =   4
         Top             =   3345
         Width           =   4155
      End
      Begin VB.Label imMenu 
         BackStyle       =   0  'Transparent
         Height          =   795
         Index           =   2
         Left            =   60
         TabIndex        =   3
         Top             =   2235
         Width           =   4155
      End
      Begin VB.Label imMenu 
         BackStyle       =   0  'Transparent
         Height          =   795
         Index           =   1
         Left            =   60
         TabIndex        =   2
         Top             =   1140
         Width           =   4155
      End
      Begin VB.Label imMenu 
         BackStyle       =   0  'Transparent
         Height          =   810
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   45
         Width           =   4155
      End
      Begin VB.Image opMenu 
         Height          =   915
         Index           =   3
         Left            =   15
         Picture         =   "frmMenu.frx":20F50
         Top             =   3300
         Width           =   4230
      End
      Begin VB.Image opMenu 
         Height          =   915
         Index           =   2
         Left            =   15
         Picture         =   "frmMenu.frx":2D9A2
         Top             =   2205
         Width           =   4230
      End
      Begin VB.Image opMenu 
         Height          =   915
         Index           =   1
         Left            =   15
         Picture         =   "frmMenu.frx":3A3F4
         Top             =   1110
         Width           =   4230
      End
      Begin VB.Image opMenu 
         Height          =   915
         Index           =   0
         Left            =   0
         Picture         =   "frmMenu.frx":46E46
         Top             =   15
         Width           =   4290
      End
      Begin VB.Label imMenu 
         BackStyle       =   0  'Transparent
         Height          =   795
         Index           =   4
         Left            =   75
         TabIndex        =   5
         Top             =   4455
         Width           =   4155
      End
      Begin VB.Image opMenu 
         Height          =   915
         Index           =   4
         Left            =   15
         Picture         =   "frmMenu.frx":53B74
         Top             =   4425
         Width           =   4215
      End
      Begin VB.Image iMenu 
         Height          =   5340
         Left            =   0
         Picture         =   "frmMenu.frx":604D2
         Top             =   15
         Width           =   4275
      End
      Begin VB.Image opcion2 
         Height          =   870
         Index           =   1
         Left            =   0
         Picture         =   "frmMenu.frx":AAB74
         Top             =   6600
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Image opcion1 
         Height          =   915
         Index           =   1
         Left            =   0
         Picture         =   "frmMenu.frx":B6AEE
         Top             =   6600
         Width           =   4215
      End
      Begin VB.Image opcion2 
         Height          =   915
         Index           =   0
         Left            =   0
         Picture         =   "frmMenu.frx":C344C
         Top             =   5520
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Image opcion1 
         Height          =   870
         Index           =   0
         Left            =   0
         Picture         =   "frmMenu.frx":CFDAA
         Top             =   5520
         Width           =   4215
      End
   End
   Begin VB.Label ubicacion 
      BackColor       =   &H000000FF&
      Height          =   615
      Left            =   4440
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLibre_Click()
frmPeleaLibre.Show , Me
End Sub

Private Sub Command1_Click()
FrmAnunciador.Show
End Sub

Private Sub Form_Load()
Cooper_skin Me

Call ocultarMenus

'Posicione picResize
picResize.Left = Screen.Width - (picResize.Width + 500)

Me.lMainTitle = "Administración Gallera - " & trabajandoCon
Me.Caption = Me.lMainTitle

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_Click()
frmConfiguracionTablero.Show
End Sub

Private Sub iMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ocultarMenus
End Sub

Private Sub imMenu_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0:
        frmRegistroCuerdas.Show , Me
    Case 1:
        frmRegistroGallos.Show , Me
    Case 2:
        frmSorteoPelea.Show , Me
    Case 3:
        frmSorteoOrden.Show , Me
    Case 4:
        frmStart.Show , Me
End Select
End Sub

Private Sub imMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.opMenu(Index).Visible = True
End Sub

Private Sub ocultarMenus()
For i = 0 To 4
    Me.opMenu(i).Visible = False
Next
End Sub

Private Sub iResHeight_Click()
FrmPantalla.Height = FrmPantalla.Height + 50
End Sub

Private Sub iResLeft_Click()
FrmPantalla.Left = FrmPantalla.Left - 50
End Sub

Private Sub iResTop_Click()
FrmPantalla.Top = FrmPantalla.Top - 50
End Sub

Private Sub iResWidth_Click()
FrmPantalla.Width = FrmPantalla.Width + 50
End Sub

Private Sub opcion1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.opcion2(Index).Visible = True
End Sub

Private Sub opcion2_Click(Index As Integer)
Dim reportToUse As String
reportToUse = IIf(reportesA4 = "Si", "consolidadoFinal_A4", "consolidadoFinal")
Select Case Index
    Case 0
        Dim oAcces As Access.Application
        Set oAcces = New Access.Application
        
        oAcces.OpenCurrentDatabase pathBD, False, keyBD
        oAcces.Visible = False
        oAcces.DoCmd.OpenReport reportToUse, acViewPreview
        
        oAcces.DoCmd.PrintOut acPrintAll
        oAcces.CloseCurrentDatabase
        oAcces.Quit
        Set oAcces = Nothing
    Case 1
        frmClasificacion.Show
End Select
End Sub

Private Sub picMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.opcion2(0).Visible = False
Me.opcion2(1).Visible = False
End Sub


