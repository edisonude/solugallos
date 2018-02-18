VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmAnunciador 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   14910
   ClientLeft      =   -180
   ClientTop       =   0
   ClientWidth     =   30000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   14910
   ScaleWidth      =   30000
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tiPausa 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   20160
      Top             =   120
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   -1080
      TabIndex        =   1
      Top             =   4320
      Width           =   240
   End
   Begin VB.PictureBox picWall 
      Height          =   18000
      Left            =   -120
      Picture         =   "FrmAnunciador.frx":0000
      ScaleHeight     =   17940
      ScaleWidth      =   29940
      TabIndex        =   2
      Top             =   0
      Width           =   30000
      Begin VB.Timer tiDicePelea 
         Interval        =   1000
         Left            =   19680
         Top             =   120
      End
      Begin VB.Timer tiPasaPeleas 
         Interval        =   30
         Left            =   19080
         Top             =   120
      End
      Begin MSComctlLib.ListView lista 
         Height          =   12195
         Left            =   1290
         TabIndex        =   4
         Top             =   2640
         Width           =   12780
         _ExtentX        =   22543
         _ExtentY        =   21511
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   4210752
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cuerda1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Gallo1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cuerda2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Gallo2"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgLogo 
         Height          =   4230
         Left            =   12570
         Picture         =   "FrmAnunciador.frx":BB8042
         Top             =   195
         Width           =   5985
      End
      Begin VB.Label etiqueta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pelea"
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
         Left            =   23040
         TabIndex        =   10
         Top             =   120
         Width           =   5490
      End
      Begin VB.Label gallo2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TuGallera"
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
         Height          =   1965
         Left            =   13965
         TabIndex        =   9
         Top             =   11760
         Width           =   15690
      End
      Begin VB.Label cuerda2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TuGallera"
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
         Height          =   1965
         Left            =   13965
         TabIndex        =   8
         Top             =   9720
         Width           =   15690
      End
      Begin VB.Label gallo1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TuGallera"
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
         Height          =   1965
         Left            =   13965
         TabIndex        =   7
         Top             =   7560
         Width           =   15690
      End
      Begin VB.Label cuerda1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TuGallera"
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
         Height          =   1965
         Left            =   13965
         TabIndex        =   6
         Top             =   5520
         Width           =   15690
      End
      Begin VB.Label nPelea 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "888"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   180
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   4395
         Left            =   23040
         TabIndex        =   5
         Top             =   840
         Width           =   5490
      End
      Begin VB.Label etiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listado de peleas"
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
         Left            =   1230
         TabIndex        =   3
         Top             =   1140
         Width           =   8610
      End
   End
   Begin VB.Label TConsecutivo 
      Height          =   210
      Left            =   -1290
      TabIndex        =   0
      Top             =   1245
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "FrmAnunciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim peleas() As String
Dim pos As Integer
Dim nPeleas As Integer
Dim ord As Integer

'Constantes par SendMessage
Const WM_VSCROLL = &H115
Const SB_BOTTOM = 7
Const SB_TOP = 6
  
'Api SendMessage
Private Declare Function SendMessage _
    Lib "user32" _
    Alias "SendMessageA" _
        (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
'Preparación de la ventana
Me.cuerda1.ForeColor = confColor1
Me.gallo1.ForeColor = confColor1
Me.cuerda2.ForeColor = confColor2
Me.gallo2.ForeColor = confColor2


con = 0
Me.Top = -50
Me.Width = 30000
Me.Height = 18000

Me.Left = 15015
Me.picWall.Height = Me.Height
Me.picWall.Width = Me.Width

pos = 1
ord = 1

Me.nPelea = ord

'Cargar peleas
Dim rs2 As New ADODB.Recordset
SQL = "SELECT * FROM resumenPeleas ORDER BY Orden Asc"
rs2.Open SQL, cnn, adOpenStatic, adLockOptimistic
nPeleas = rs2.RecordCount

ReDim peleas(nPeleas, 5)
rs2.MoveFirst

Dim i As Integer
For i = 1 To rs2.RecordCount
    peleas(i, 1) = rs2("orden")
    peleas(i, 2) = rs2("Cuerda")
    peleas(i, 3) = rs2("anillo")
    peleas(i, 4) = rs2("Cuerda2")
    peleas(i, 5) = rs2("anillo2")
    rs2.MoveNext
Next
rs2.Close

With lista
    .ColumnHeaders(1).Width = .Width * 0.1
    .ColumnHeaders(2).Width = .Width * 0.3
    .ColumnHeaders(3).Width = .Width * 0.15
    .ColumnHeaders(4).Width = .Width * 0.3
    .ColumnHeaders(5).Width = .Width * 0.15
End With
End Sub

Private Sub tiDicePelea_Timer()
Me.tiPasaPeleas.Enabled = False
Call cargarPelea(ord)

Dim li As ListItem
Set li = lista.ListItems.Add(, , peleas(ord, 1))
    li.SubItems(1) = peleas(ord, 2)
    li.SubItems(2) = peleas(ord, 3)
    li.SubItems(3) = peleas(ord, 4)
    li.SubItems(4) = peleas(ord, 5)
    
    Desplazar lista, SB_BOTTOM
ord = ord + 1
If ord > nPeleas Then
    Me.tiDicePelea.Enabled = False
    Exit Sub
End If
Me.nPelea = ord
tiDicePelea.Enabled = False
tiPausa.Enabled = True
End Sub

Private Sub tiPasaPeleas_Timer()
If pos = nPeleas Then pos = 1
Call cargarPelea(pos)
pos = pos + 1
End Sub

Private Sub cargarPelea(pos As Integer)
Me.cuerda1 = peleas(pos, 2)
Me.gallo1 = "Gallo: " & peleas(pos, 3)
Me.cuerda2 = peleas(pos, 4)
Me.gallo2 = "Gallo: " & peleas(pos, 5)
End Sub

Private Sub Desplazar(Control As Control, Accion As Long)
    'Ejecutamos SendMessage pasandole el control y el mensaje
    ret = SendMessage(Control.hwnd, WM_VSCROLL, Accion, 0)
End Sub

Private Sub tiPausa_Timer()
tiDicePelea.Enabled = True
tiPasaPeleas.Enabled = True
tiPausa.Enabled = False
End Sub
