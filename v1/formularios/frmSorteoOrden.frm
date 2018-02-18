VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmSorteoOrden 
   Appearance      =   0  'Flat
   BackColor       =   &H00555555&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   13035
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
      Left            =   12600
      TabIndex        =   32
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
         TabIndex        =   18
         Top             =   4680
         Width           =   615
      End
      Begin VB.PictureBox Picture5 
         Height          =   4935
         Index           =   1
         Left            =   11400
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
         Width           =   11025
         _ExtentX        =   19447
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IdPelea"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Orden"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Placa Gallo 1"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "VS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Placa Gallo 1"
            Object.Width           =   8820
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   615
         Index           =   2
         Left            =   11640
         OleObjectBlob   =   "frmSorteoOrden.frx":0000
         TabIndex        =   16
         Top             =   3360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel tNPeleas 
         Height          =   495
         Left            =   11640
         OleObjectBlob   =   "frmSorteoOrden.frx":006C
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
      ScaleWidth      =   12375
      TabIndex        =   2
      Top             =   480
      Width           =   12375
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmSorteoOrden.frx":00C6
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   2535
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   11535
      Begin VB.CommandButton btnAutomatico 
         Caption         =   "Automático"
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
         Left            =   7680
         TabIndex        =   33
         Top             =   720
         Width           =   1935
      End
      Begin VB.PictureBox Picture5 
         Height          =   1095
         Index           =   0
         Left            =   3720
         ScaleHeight     =   1095
         ScaleWidth      =   15
         TabIndex        =   31
         Top             =   240
         Width           =   15
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
         Left            =   5520
         TabIndex        =   28
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox tSorteo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.PictureBox Picture5 
         Height          =   1095
         Index           =   2
         Left            =   2040
         ScaleHeight     =   1095
         ScaleWidth      =   15
         TabIndex        =   25
         Top             =   240
         Width           =   15
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
            Picture         =   "frmSorteoOrden.frx":02FA
            Stretch         =   -1  'True
            Top             =   0
            Width           =   750
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   4
         Left            =   240
         OleObjectBlob   =   "frmSorteoOrden.frx":432C
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tCuerda1 
         Height          =   375
         Left            =   1320
         OleObjectBlob   =   "frmSorteoOrden.frx":4390
         TabIndex        =   7
         Top             =   1800
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   9
         Left            =   240
         OleObjectBlob   =   "frmSorteoOrden.frx":43FC
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPeso1 
         Height          =   375
         Left            =   1320
         OleObjectBlob   =   "frmSorteoOrden.frx":445C
         TabIndex        =   9
         Top             =   2160
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   13
         Left            =   6360
         OleObjectBlob   =   "frmSorteoOrden.frx":44C4
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tCuerda2 
         Height          =   375
         Left            =   7440
         OleObjectBlob   =   "frmSorteoOrden.frx":4528
         TabIndex        =   12
         Top             =   1800
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   15
         Left            =   6360
         OleObjectBlob   =   "frmSorteoOrden.frx":4594
         TabIndex        =   13
         Top             =   2160
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPeso2 
         Height          =   375
         Left            =   7440
         OleObjectBlob   =   "frmSorteoOrden.frx":45F4
         TabIndex        =   14
         Top             =   2160
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   0
         Left            =   240
         OleObjectBlob   =   "frmSorteoOrden.frx":465C
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPlaca1 
         Height          =   375
         Left            =   1320
         OleObjectBlob   =   "frmSorteoOrden.frx":46BE
         TabIndex        =   20
         Top             =   1440
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel Placa 
         Height          =   375
         Index           =   3
         Left            =   6360
         OleObjectBlob   =   "frmSorteoOrden.frx":4728
         TabIndex        =   21
         Top             =   1440
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPlaca2 
         Height          =   375
         Left            =   7440
         OleObjectBlob   =   "frmSorteoOrden.frx":478A
         TabIndex        =   22
         Top             =   1440
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   5
         Left            =   240
         OleObjectBlob   =   "frmSorteoOrden.frx":47F4
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel tPeleasRestantes 
         Height          =   615
         Left            =   240
         OleObjectBlob   =   "frmSorteoOrden.frx":486C
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   6
         Left            =   3840
         OleObjectBlob   =   "frmSorteoOrden.frx":48CC
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   1
         Left            =   2160
         OleObjectBlob   =   "frmSorteoOrden.frx":493E
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel tNumPelea 
         Height          =   615
         Left            =   2160
         OleObjectBlob   =   "frmSorteoOrden.frx":49AE
         TabIndex        =   30
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio de peleas"
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
Attribute VB_Name = "frmSorteoOrden"
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

Private Sub btnAutomatico_Click()
frmOrdenAutomatico.Show
End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub btnAsignar_Click()
If Me.tPlaca1 = "Sin placa" Or Me.tPlaca2 = "Sin placa" Then
    MsgBox "Debe digitar el número de la pelea", vbCritical, "Sin pelea"
    Me.tSorteo.SetFocus
    seleccionarTexto tSorteo
    Exit Sub
End If

Set li = listaGallos.ListItems.Add(, , Me.tSorteo)
    li.SubItems(1) = nPelea
    li.SubItems(2) = Me.tPlaca1
    li.SubItems(3) = "vs"
    li.SubItems(4) = Me.tPlaca2

    nPelea = nPelea + 1
    restantes = restantes - 1
    
    Me.tNPeleas = listaGallos.ListItems.Count
    Me.tNumPelea = Format(nPelea, "000")
    Me.tPeleasRestantes = Format(restantes, "000")
    
    Me.tPlaca2 = "Sin placa"
    Me.tPeso2 = "Sin peso"
    Me.TCuerda2 = "Sin cuerda"
    Me.tPlaca1 = "Sin placa"
    Me.tPeso1 = "Sin peso"
    Me.TCuerda1 = "Sin cuerda"
    
    If restantes = 0 Then
        Me.btnAsignar.Enabled = False
        MsgBox "Finalizo el sorteo de orden de las peleas", vbInformation, "Fin"
        tSorteo.Enabled = False
    Else
        Me.tSorteo = ""
        Me.tSorteo.SetFocus
    End If
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

Private Sub Form_Load()
Cooper_skin Me

Me.Top = frmMenu.ubicacion.Top + 350
Me.Left = frmMenu.ubicacion.Left

numPeleas = peleasPendientes

If numPeleas = 0 Then
    MsgBox "No hay peleas para ordenar", vbInformation, "Sin peleas"
    Unload Me
    Exit Sub
End If

restantes = numPeleas

'nPelea = 1

'Obtengo el numero de la pelea inicial
nPelea = HallaConsecutivo("Select Max(orden) As NConsecutivo from Peleas")

Me.tPeleasRestantes = Format(numPeleas, "000")
Me.tNumPelea = Format(nPelea, "000")
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

Private Sub buscarPelea()
Dim rs2 As New ADODB.Recordset
SQL = "SELECT * " & _
        "FROM Peleas WHERE idPelea = " & tSorteo.Text & " and orden = 0"
rs2.Open SQL, cnn, adOpenStatic, adLockOptimistic

If rs2.RecordCount >= 1 Then
    Me.tPlaca1.Caption = rs2("idGallo1")
    Me.tPlaca2.Caption = rs2("idGallo2")
    rs2.Close
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
    
    Me.tSorteo.SetFocus
    seleccionarTexto tSorteo
    rs2.Close
End If

End Sub

Private Sub buscarGallo1()
SQL = "SELECT * " & _
        "FROM Gallos WHERE idGallo = " & tPlaca1 & ""
rs.Open SQL, cnn, adOpenStatic, adLockOptimistic

If rs.RecordCount >= 1 Then
    Me.TCuerda1 = nombreCuerda(rs("idCuerdaPelea"))
    Me.tPeso1 = IIf(IsNull(rs("peso")), "", rs("peso"))
End If
rs.Close
End Sub

Private Sub buscarGallo2()
SQL = "SELECT * " & _
        "FROM Gallos WHERE idGallo = " & tPlaca2 & ""
rs.Open SQL, cnn, adOpenStatic, adLockOptimistic

If rs.RecordCount >= 1 Then
    Me.TCuerda2 = nombreCuerda(rs("idCuerdaPelea"))
    Me.tPeso2 = IIf(IsNull(rs("peso")), "", rs("peso"))
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

Private Sub tSorteo_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
End Sub

Private Sub tSorteo_LostFocus()
If Me.tSorteo = "" Then Exit Sub
Call buscarPelea
End Sub
