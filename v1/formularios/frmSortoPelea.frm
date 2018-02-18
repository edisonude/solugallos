VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmSorteoPelea 
   BackColor       =   &H00555555&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   13035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAutomatico 
      Caption         =   "Sorteo Automático"
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
      Left            =   10200
      TabIndex        =   46
      Top             =   240
      Width           =   2175
   End
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
      Left            =   12600
      TabIndex        =   34
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Lista final de peleas"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   12375
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
         Left            =   11640
         TabIndex        =   33
         Top             =   3600
         Width           =   615
      End
      Begin VB.PictureBox Picture5 
         Height          =   3735
         Index           =   1
         Left            =   11400
         ScaleHeight     =   3735
         ScaleWidth      =   15
         TabIndex        =   30
         Top             =   360
         Width           =   15
      End
      Begin MSComctlLib.ListView listaGallos 
         Height          =   3660
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   6456
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IdGallo1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IdGallo2"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PlacaCuerda 1"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Anillo 1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "VS"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "PlacaCuerda 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Anillo 2"
            Object.Width           =   2540
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   615
         Index           =   2
         Left            =   11640
         OleObjectBlob   =   "frmSortoPelea.frx":0000
         TabIndex        =   31
         Top             =   2280
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel tNPeleas 
         Height          =   495
         Left            =   11640
         OleObjectBlob   =   "frmSortoPelea.frx":006C
         TabIndex        =   32
         Top             =   3000
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
      ScaleWidth      =   12375
      TabIndex        =   9
      Top             =   600
      Width           =   12375
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmSortoPelea.frx":00C6
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   2775
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   12375
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
         Index           =   1
         Left            =   6480
         TabIndex        =   3
         Top             =   960
         Width           =   1215
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
         Index           =   1
         Left            =   8160
         TabIndex        =   4
         Top             =   960
         Width           =   1215
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
         Index           =   1
         Left            =   9720
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.PictureBox pLabel 
         Height          =   15
         Index           =   1
         Left            =   120
         ScaleHeight     =   15
         ScaleWidth      =   5295
         TabIndex        =   37
         Top             =   1440
         Width           =   5295
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
         Index           =   0
         Left            =   3600
         TabIndex        =   2
         Top             =   960
         Width           =   1215
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
         Index           =   0
         Left            =   2040
         TabIndex        =   1
         Top             =   960
         Width           =   1215
      End
      Begin VB.PictureBox Picture5 
         Height          =   1815
         Index           =   0
         Left            =   11400
         ScaleHeight     =   1815
         ScaleWidth      =   15
         TabIndex        =   29
         Top             =   240
         Width           =   15
      End
      Begin VB.PictureBox picAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   11640
         Picture         =   "frmSortoPelea.frx":02FA
         ScaleHeight     =   675
         ScaleWidth      =   660
         TabIndex        =   6
         Top             =   1320
         Width           =   660
      End
      Begin VB.PictureBox Picture4 
         Height          =   15
         Left            =   6120
         ScaleHeight     =   15
         ScaleWidth      =   5295
         TabIndex        =   20
         Top             =   540
         Width           =   5295
      End
      Begin VB.PictureBox pLabel 
         Height          =   15
         Index           =   0
         Left            =   120
         ScaleHeight     =   15
         ScaleWidth      =   5295
         TabIndex        =   15
         Top             =   540
         Width           =   5295
      End
      Begin VB.PictureBox Picture2 
         Height          =   1455
         Left            =   5400
         ScaleHeight     =   1395
         ScaleWidth      =   675
         TabIndex        =   7
         Top             =   360
         Width           =   735
         Begin VB.Image Image1 
            Height          =   1395
            Left            =   -120
            Picture         =   "frmSortoPelea.frx":1A70
            Top             =   0
            Width           =   870
         End
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
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   0
         Left            =   240
         OleObjectBlob   =   "frmSortoPelea.frx":5AA2
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   3
         Left            =   240
         OleObjectBlob   =   "frmSortoPelea.frx":5B08
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   4
         Left            =   360
         OleObjectBlob   =   "frmSortoPelea.frx":5B7C
         TabIndex        =   16
         Top             =   1560
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tCuerda 
         Height          =   375
         Index           =   0
         Left            =   1440
         OleObjectBlob   =   "frmSortoPelea.frx":5BE0
         TabIndex        =   17
         Top             =   1560
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   9
         Left            =   360
         OleObjectBlob   =   "frmSortoPelea.frx":5C44
         TabIndex        =   18
         Top             =   1920
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPeso 
         Height          =   375
         Index           =   0
         Left            =   1440
         OleObjectBlob   =   "frmSortoPelea.frx":5CA4
         TabIndex        =   19
         Top             =   1920
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   13
         Left            =   6360
         OleObjectBlob   =   "frmSortoPelea.frx":5D08
         TabIndex        =   21
         Top             =   1560
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tCuerda 
         Height          =   375
         Index           =   1
         Left            =   7440
         OleObjectBlob   =   "frmSortoPelea.frx":5D6C
         TabIndex        =   22
         Top             =   1560
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   15
         Left            =   6360
         OleObjectBlob   =   "frmSortoPelea.frx":5DD0
         TabIndex        =   23
         Top             =   1920
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPeso 
         Height          =   375
         Index           =   1
         Left            =   7440
         OleObjectBlob   =   "frmSortoPelea.frx":5E30
         TabIndex        =   24
         Top             =   1920
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   1
         Left            =   11640
         OleObjectBlob   =   "frmSortoPelea.frx":5E94
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPelea 
         Height          =   495
         Left            =   11640
         OleObjectBlob   =   "frmSortoPelea.frx":5EF6
         TabIndex        =   28
         Top             =   720
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   5
         Left            =   1920
         OleObjectBlob   =   "frmSortoPelea.frx":5F50
         TabIndex        =   35
         Top             =   600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   6
         Left            =   3480
         OleObjectBlob   =   "frmSortoPelea.frx":5FC0
         TabIndex        =   36
         Top             =   600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   7
         Left            =   360
         OleObjectBlob   =   "frmSortoPelea.frx":6024
         TabIndex        =   38
         Top             =   2280
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tColor 
         Height          =   375
         Index           =   0
         Left            =   1440
         OleObjectBlob   =   "frmSortoPelea.frx":6086
         TabIndex        =   39
         Top             =   2280
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   8
         Left            =   6360
         OleObjectBlob   =   "frmSortoPelea.frx":60EA
         TabIndex        =   40
         Top             =   2280
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tColor 
         Height          =   375
         Index           =   1
         Left            =   7440
         OleObjectBlob   =   "frmSortoPelea.frx":614C
         TabIndex        =   41
         Top             =   2280
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   10
         Left            =   6360
         OleObjectBlob   =   "frmSortoPelea.frx":61B0
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   11
         Left            =   6360
         OleObjectBlob   =   "frmSortoPelea.frx":6216
         TabIndex        =   43
         Top             =   600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   12
         Left            =   8040
         OleObjectBlob   =   "frmSortoPelea.frx":628A
         TabIndex        =   44
         Top             =   600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   255
         Index           =   14
         Left            =   9600
         OleObjectBlob   =   "frmSortoPelea.frx":62FA
         TabIndex        =   45
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label tIdGallo 
         Height          =   255
         Index           =   1
         Left            =   7560
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.Label tIdGallo 
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sorteo de las peleas"
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
      TabIndex        =   11
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmSorteoPelea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim li As ListItem
Dim encabezado As String
Dim numeroGallo As Integer
Dim nPelea As Integer
Dim numPeleas As Integer

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub btnFin_Click()
Dim conPelea As Integer

With listaGallos
    For i = 1 To .ListItems.Count
        conPelea = HallaConsecutivo("Select Max(idPelea) As NConsecutivo from Peleas")
        
        SQL = "insert into Peleas values(" & conPelea & "," & .ListItems.Item(i) & _
        ",'" & .ListItems.Item(i).SubItems(1) & "','" & Format(Date, "dd/mm/yyyy") & "',0,'',0,0,0,'no')"
            
        Call guardarRDO
    Next
End With
menGuardadoExitoso
End Sub


Private Sub cmdAutomatico_Click()
frmSorteoPeleaAutomatico.Show
End Sub

Private Sub Form_Load()
Cooper_skin Me

Me.Top = frmMenu.ubicacion.Top + 350
Me.Left = frmMenu.ubicacion.Left

numPeleas = 0

nPelea = HallaConsecutivo("Select Max(idPelea) As NConsecutivo from Peleas")
Me.TPelea = nPelea

'Ancho de la lista
Me.listaGallos.ColumnHeaders(1).Width = 0
Me.listaGallos.ColumnHeaders(2).Width = 0
Me.listaGallos.ColumnHeaders(3).Width = Me.listaGallos.Width * 0.24
Me.listaGallos.ColumnHeaders(4).Width = Me.listaGallos.Width * 0.24
Me.listaGallos.ColumnHeaders(5).Width = Me.listaGallos.Width * 0.05
Me.listaGallos.ColumnHeaders(6).Width = Me.listaGallos.Width * 0.23
Me.listaGallos.ColumnHeaders(7).Width = Me.listaGallos.Width * 0.23
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

Private Sub picAdd_Click()
If Me.tIdGallo(0).Caption = "" Or Me.tIdGallo(1).Caption = "" Then
    MsgBox "Debe digitar los dos gallos a pelear", vbCritical, "Sin gallo"
    Me.tPlacaCuerda(0).SetFocus
    Exit Sub
End If

If Not pasaGallo(Val(Me.tIdGallo(0))) Then
    MsgBox "Este gallo " & tPlacaCuerda(0) & " ya se encuentra en lista de pelea", vbExclamation, "Gallo ya pelea"
    Me.tPlacaCuerda(0).SetFocus
    Me.tPlacaCuerda(0) = ""
    Me.tPlacaNacional(0) = ""
    Me.tAnillo(0) = ""
    Exit Sub
End If

If Not pasaGallo(Val(Me.tIdGallo(1))) Then
    MsgBox "Este gallo " & tPlacaCuerda(1) & " ya se encuentra en lista de pelea", vbExclamation, "Gallo ya pelea"
    Me.tPlacaCuerda(1).SetFocus
    Me.tPlacaCuerda(1) = ""
    Me.tPlacaNacional(1) = ""
    Me.tAnillo(1) = ""
    Exit Sub
End If

Set li = listaGallos.ListItems.Add(, , Me.tIdGallo(0))
    li.SubItems(1) = Me.tIdGallo(1)
    li.SubItems(2) = Me.tPlacaCuerda(0)
    li.SubItems(3) = Me.tAnillo(0)
    li.SubItems(4) = "vs"
    li.SubItems(5) = Me.tPlacaCuerda(1)
    li.SubItems(6) = Me.tAnillo(1)
    
    Me.tIdGallo(0) = ""
    Me.tPlacaCuerda(0) = ""
    Me.tPlacaNacional(0) = ""
    Me.tAnillo(0) = ""
    Me.tPeso(0) = "Sin peso"
    Me.tColor(0) = "Sin color"
    Me.tCuerda(0) = "Sin cuerda"
    
    Me.tIdGallo(1) = ""
    Me.tPlacaCuerda(1) = ""
    Me.tPlacaNacional(1) = ""
    Me.tAnillo(1) = ""
    Me.tPeso(1) = "Sin peso"
    Me.tColor(1) = "Sin color"
    Me.tCuerda(1) = "Sin cuerda"
    
    nPelea = nPelea + 1
    numPeleas = numPeleas + 1
    Me.tNPeleas = numPeleas
End Sub

Private Sub buscarGallo(placaNacional As String, placaCuerda As String, anillo As String, ind As Integer)
SQL = "SELECT g.*,c.Cuerda as cuerda FROM Gallos g INNER JOIN Cuerdas c ON g.idCuerdaPelea = c.idCuerda " & _
    "WHERE placaNacional='" & placaNacional & "' or placaCuerda='" & placaCuerda & "' or anillo='" & anillo & "'"

Dim qry As New rdoQuery
Dim rs As rdoResultset

Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rs = qry.OpenResultset(rdOpenDynamic)

If rs.RowCount >= 1 Then
    Me.tIdGallo(ind) = IIf(IsNull(rs("idGallo")), "", rs("idGallo"))
    Me.tPlacaCuerda(ind) = IIf(IsNull(rs("placaCuerda")), "", rs("placaCuerda"))
    Me.tPlacaNacional(ind) = IIf(IsNull(rs("placaNacional")), "", rs("placaNacional"))
    Me.tAnillo(ind) = IIf(IsNull(rs("anillo")), "", rs("anillo"))
    Me.tCuerda(ind) = IIf(IsNull(rs("cuerda")), "", rs("cuerda"))
    Me.tPeso(ind) = IIf(IsNull(rs("peso")), "", rs("peso"))
    Me.tPlacaCuerda(ind) = IIf(IsNull(rs("placaCuerda")), "", rs("placaCuerda"))
    Me.tColor(ind) = IIf(IsNull(rs("colorPluma")), "", rs("colorPluma"))
Else
    Me.tIdGallo(ind) = ""
    Me.tPlacaCuerda(ind) = ""
    Me.tPlacaNacional(ind) = ""
    Me.tAnillo(ind) = ""
    Me.tCuerda(ind) = "Sin cuerda"
    Me.tPeso(ind) = "Sin peso"
    Me.tColor(ind) = "Sin color"
End If
rs.Close
End Sub

Private Function pasaGallo(pla As Integer) As Boolean
If pla = 0 Then pasaGallo = True: Exit Function
With listaGallos
    For i = 1 To .ListItems.Count
        If .ListItems.Item(i) = pla Or .ListItems.Item(i).SubItems(1) = pla Then
            validarGallo = False
            Exit Function
        End If
    Next
    pasaGallo = True
End With
End Function

Private Sub tAnillo_LostFocus(Index As Integer)
If Me.tAnillo(Index) = "" Then Exit Sub
tAnillo(Index) = UCase(tAnillo(Index))
Call buscarGallo("0", "0", tAnillo(Index), Index)
End Sub

Private Sub tPlacaCuerda_LostFocus(Index As Integer)
If Me.tPlacaCuerda(Index) = "" Then Exit Sub
tPlacaCuerda(Index) = UCase(tPlacaCuerda(Index))
Call buscarGallo("0", tPlacaCuerda(Index), "0", Index)
End Sub

Private Sub tPlacaNacional_LostFocus(Index As Integer)
If Me.tPlacaNacional(Index) = "" Then Exit Sub
tPlacaNacional(Index) = UCase(tPlacaNacional(Index))
Call buscarGallo(tPlacaNacional(Index), "0", "0", Index)
End Sub
