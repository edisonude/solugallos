VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmOrdenAutomatico 
   BackColor       =   &H00808080&
   Caption         =   "Generar separacion minima"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   13980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPropietario 
      Caption         =   "Ordenar por propietario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   30
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton btnRellenar 
      Caption         =   "Rellenar"
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
      Left            =   11760
      TabIndex        =   29
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton btnFinalizar 
      Caption         =   "Finalizar"
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
      TabIndex        =   28
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton btnAnunciar 
      Caption         =   "Anunciar"
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
      Height          =   495
      Left            =   2040
      TabIndex        =   27
      Top             =   7680
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   15
      Index           =   1
      Left            =   240
      ScaleHeight     =   15
      ScaleWidth      =   3375
      TabIndex        =   26
      Top             =   7440
      Width           =   3375
   End
   Begin VB.CommandButton btnMejor 
      Caption         =   "Mejor resul"
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
      Left            =   8880
      TabIndex        =   25
      Top             =   480
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel tNint 
      Height          =   495
      Left            =   240
      OleObjectBlob   =   "frmOrdenAutomatico.frx":0000
      TabIndex        =   23
      Top             =   6840
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   15
      Index           =   0
      Left            =   240
      ScaleHeight     =   15
      ScaleWidth      =   3375
      TabIndex        =   22
      Top             =   6480
      Width           =   3375
   End
   Begin VB.CommandButton btnOrdenar 
      Caption         =   "Ordenar"
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
      Left            =   7440
      TabIndex        =   20
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton btnRepintar 
      Caption         =   "Repintar"
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
      Left            =   10320
      TabIndex        =   19
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Resetear"
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
      Left            =   2040
      TabIndex        =   18
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton btnSortear 
      Caption         =   "Sortear"
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
      TabIndex        =   17
      Top             =   5760
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
      Height          =   615
      Index           =   0
      Left            =   240
      OleObjectBlob   =   "frmOrdenAutomatico.frx":005C
      TabIndex        =   3
      Top             =   120
      Width           =   6975
   End
   Begin VB.Frame marco 
      Caption         =   "Configuracion sorteo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   3375
      Begin VB.TextBox tMaxRep 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   16
         Text            =   "1000"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox tMaxInt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   15
         Text            =   "100"
         Top             =   480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   5
         Left            =   240
         OleObjectBlob   =   "frmOrdenAutomatico.frx":0100
         TabIndex        =   13
         Top             =   480
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   6
         Left            =   240
         OleObjectBlob   =   "frmOrdenAutomatico.frx":0170
         TabIndex        =   14
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.Frame marco 
      Caption         =   "Información de peleas"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Index           =   1
      Left            =   4200
      TabIndex        =   1
      Top             =   960
      Width           =   8895
      Begin MSComctlLib.ListView lista 
         Height          =   7095
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   12515
         SortKey         =   3
         View            =   3
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "idPelea"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cuerda1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cuerda2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Orden"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame marco 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Información de peleas"
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   3375
      Begin VB.TextBox tSepMin 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   12
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton btnSepMin 
         Caption         =   "Generar separacion minima"
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
         Left            =   360
         TabIndex        =   11
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox tNoDias 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   10
         Text            =   "1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox tGallosxCuerda 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   9
         Text            =   "4"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox tPeleasOrdenar 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   1
         Left            =   240
         OleObjectBlob   =   "frmOrdenAutomatico.frx":01E8
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   2
         Left            =   240
         OleObjectBlob   =   "frmOrdenAutomatico.frx":0268
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   3
         Left            =   240
         OleObjectBlob   =   "frmOrdenAutomatico.frx":02EA
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   4
         Left            =   240
         OleObjectBlob   =   "frmOrdenAutomatico.frx":0352
         TabIndex        =   7
         Top             =   2280
         Width           =   2175
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmOrdenAutomatico.frx":03BE
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
      Height          =   375
      Index           =   7
      Left            =   240
      OleObjectBlob   =   "frmOrdenAutomatico.frx":05F2
      TabIndex        =   21
      Top             =   6600
      Width           =   1935
   End
End
Attribute VB_Name = "frmOrdenAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DEclaración de la Función Api SendMessage
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Dim nPeleas As Integer
Dim nGallosxCuerda As Integer
Dim nDias As Integer
Dim sepMin As Integer

Dim peleas() As String
Dim peleasBase() As String
Dim peleasMejor() As String
Dim posAct As Integer
Dim nInt  As Integer
Dim peleasOrdenadas As Integer

Dim maxRep As Double

Private Sub btnAnunciar_Click()
FrmAnunciador.Show
End Sub

Private Sub btnFinalizar_Click()
'If faltanPeleas Then
'    MsgBox "Aun faltan feleas ordenadas. " & vbCrLf & "Nose puede finalizar si faltan peleas", vbCritical
'    Exit Sub
'End If
            

Dim i As Integer
For i = 1 To nPeleas
    SQL = "UPDATE Peleas SET Orden=" & peleasMejor(i, 4) & " WHERE idPelea=" & peleasMejor(i, 1) & ""
    Call guardarRDO
Next
Call menGuardadoExitoso
Me.btnAnunciar.Enabled = True
End Sub

Private Sub btnMejor_Click()
Dim li As ListItem
Dim i As Integer
lista.ListItems.Clear
For i = 1 To nPeleas
    Set li = lista.ListItems.Add(, , peleasMejor(i, 1))
        li.SubItems(1) = peleasMejor(i, 2)
        li.SubItems(2) = peleasMejor(i, 3)
        li.SubItems(3) = peleasMejor(i, 4)
Next
End Sub

Private Sub btnOrdenar_Click()
' Ordena alfabéticamente la columna con números _
  ( es la columna que tiene en el tag el valor NUMBER )
  
  Dim Formato As String
Dim strData() As String

Dim Columna As Long

Call SendMessage(Me.hwnd, WM_SETREDRAW, 0&, 0&)


Columna = 3

Formato = String(30, "0") & "." & String(30, "0")
      
With lista.ListItems
    If (Columna > 0) Then
        For i = 1 To .Count
            With .Item(i).ListSubItems(Columna)
                .Tag = .Text & Chr$(0) & .Tag
                If IsNumeric(.Text) Then
                    If CDbl(.Text) >= 0 Then
                        .Text = Format(CDbl(.Text), _
                            Formato)
                    Else
                        .Text = "&" & InvNumber( _
                            Format(0 - CDbl(.Text), _
                            Formato))
                    End If
                Else
                    .Text = ""
                End If
            End With
        Next i
    Else
        For i = 1 To .Count
            With .Item(i)
                .Tag = .Text & Chr$(0) & .Tag
                If IsNumeric(.Text) Then
                    If CDbl(.Text) >= 0 Then
                        .Text = Format(CDbl(.Text), _
                            Formato)
                    Else
                        .Text = "&" & InvNumber( _
                            Format(0 - CDbl(.Text), _
                            Formato))
                    End If
                Else
                    .Text = ""
                End If
            End With
        Next i
    End If
End With
  
' Ordena alfabéticamente
  
lista.SortOrder = (lista.SortOrder + 1) Mod 2
lista.SortKey = 3
lista.Sorted = True
  
With lista.ListItems
    If (Columna > 0) Then
        For i = 1 To .Count
            With .Item(i).ListSubItems(Columna)
                strData = Split(.Tag, Chr$(0))
                .Text = strData(0)
                .Tag = strData(1)
            End With
        Next i
    Else
        For i = 1 To .Count
            With .Item(i)
                strData = Split(.Tag, Chr$(0))
                .Text = strData(0)
                .Tag = strData(1)
            End With
        Next i
    End If
End With
End Sub

Private Sub btnRellenar_Click()
If MsgBox("¿Está seguro de rellenar el orden de las peleas?" & vbCrLf & "Las peleas no organizadas se pondran de ultimas en la lista", vbQuestion + vbYesNo) = vbYes Then
    For i = 1 To nPeleas
        If peleasMejor(i, 4) = 0 Then
            peleasOrdenadas = peleasOrdenadas + 1
            peleasMejor(i, 4) = peleasOrdenadas
        End If
    Next
End If

Call btnMejor_Click
End Sub

Private Sub btnRepintar_Click()
Dim li As ListItem
Dim i As Integer
lista.ListItems.Clear
For i = 1 To nPeleas
    Set li = lista.ListItems.Add(, , peleas(i, 1))
        li.SubItems(1) = peleas(i, 2)
        li.SubItems(2) = peleas(i, 3)
        li.SubItems(3) = peleas(i, 4)
Next
End Sub

Private Sub btnReset_Click()
Call reset

nInt = Val(tMaxInt)
tNint = nInt

peleasOrdenadas = 0
peleasMejor = peleasBase

sepMin = Val(Me.tSepMin)


Call btnRepintar_Click
End Sub

Private Sub btnSepMin_Click()
sepMin = (nPeleas / (nGallosxCuerda / nDias)) \ 2
Me.tSepMin = sepMin
End Sub

Private Sub Command3_Click()

End Sub

Private Sub btnSortear_Click()
Dim ind As Integer
Dim ultOrd  As Integer
Dim paso As Integer

While nInt >= 0
    While faltanPeleas And maxRep >= 0
        ind = random(nPeleas)
        
        While peleaOrdenada(ind)
            ind = random(nPeleas)
        Wend
        
        ultOrd = ordUlPelea(peleas(ind, 2), peleas(ind, 3))
        paso = posAct - ultOrd
        
        If paso > sepMin Then
            peleas(ind, 4) = posAct
            posAct = posAct + 1
        End If
        maxRep = maxRep - 1
    Wend
    If maxRep < 0 Then
        nInt = nInt - 1
        tNint.Caption = nInt
        DoEvents
        'Verificar si es mejor resultado
        Call esMejorResultado
        Call reset
    Else
        peleasMejor = peleas
        nInt = -2
    End If
Wend
Call btnRepintar_Click
If nInt = -2 Then
    MsgBox "Las peleas se ordenaron correctamente", vbInformation
Else
    MsgBox "No se logro organizar las peleas, resetee e intente de nuevo", vbCritical
End If
End Sub

Private Sub Command1_Click()
Dim li As ListItem
Dim i As Integer
lista.ListItems.Clear
For i = 1 To nPeleas
    Set li = lista.ListItems.Add(, , peleas(i, 1))
        li.SubItems(1) = peleas(i, 2)
        li.SubItems(2) = peleas(i, 3)
        li.SubItems(3) = peleas(i, 4)
Next
End Sub

Private Sub chkPropietario_Click()
Call Form_Load
Me.chkPropietario.Visible = True
Me.chkPropietario.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next

Cooper_skin Me

peleasOrdenadas = 0

'Cargo los datos para el sorteo
nPeleas = peleasxOrdenar
nGallosxCuerda = Val(Me.tGallosxCuerda)
nDias = Val(Me.tNoDias)
Me.tPeleasOrdenar = nPeleas
maxRep = Val(tMaxRep)
nInt = Val(tMaxInt)
tNint = nInt

'Posicion inicial para ordenar
posAct = 1

'Cargar peleas
ReDim peleas(nPeleas, 4)

Dim rs2 As New ADODB.Recordset
SQL = "SELECT * FROM resumenPeleas"
rs2.Open SQL, cnn, adOpenStatic, adLockOptimistic
rs2.MoveFirst

Dim i As Integer
For i = 1 To rs2.RecordCount
    peleas(i, 1) = rs2("idPelea")


If Me.chkPropietario.Value = 1 Then
    '   VERSION PRUEBA ORDENANDO POR PROPIETARIOS
    peleas(i, 2) = rs2("Propietario")
    peleas(i, 3) = rs2("Propietario2")
Else
'   VERSION ORIGINAL ORDENANDO POR FRENTES
    If rs2("Frente") = 0 Then
        peleas(i, 2) = rs2("idCuerda")
    Else
        peleas(i, 2) = rs2("Frente")
    End If

    If rs2("Frente2") = 0 Then
        peleas(i, 3) = rs2("idCuerda2")
    Else
        peleas(i, 3) = rs2("Frente2")
    End If
End If

    
    peleas(i, 4) = rs2("orden")
    rs2.MoveNext
Next
rs2.Close

'Matriz de peleas base
peleasBase = peleas

Call btnSepMin_Click
End Sub

Private Function ordUlPelea(c1, c2) As Integer
Dim i As Integer
Dim ord As Integer
ord = -1
For i = 1 To nPeleas
    If (peleas(i, 2) = c1 Or peleas(i, 3) = c1 Or peleas(i, 2) = c2 Or peleas(i, 3) = c2) Then
        If ord < peleas(i, 4) Then
            ord = peleas(i, 4)
        End If
    End If
Next
If ord = 0 Then ord = sepMin * -1
ordUlPelea = ord
End Function

Private Function peleaOrdenada(ind As Integer) As Boolean
If peleas(ind, 4) = 0 Then
    peleaOrdenada = False
Else
    peleaOrdenada = True
End If
End Function

Private Function faltanPeleas() As Boolean
For i = 1 To nPeleas
    If (peleas(i, 4) = 0) Then
        faltanPeleas = True
        Exit Function
    End If
Next
faltanPeleas = False
End Function


Private Sub reset()
'Cargo los datos para el sorteo
nPeleas = peleasxOrdenar
Me.tPeleasOrdenar = nPeleas
maxRep = Val(tMaxRep)

'Posicion inicial para ordenar
posAct = 1

'Cargar peleas
peleas = peleasBase
End Sub

Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function

Private Sub esMejorResultado()
Dim i As Integer
Dim nfal As Integer
For i = 1 To nPeleas
    If (peleas(i, 4) = 0) Then
        nfal = nfal + 1
    End If
Next
If (nPeleas - nfal) > peleasOrdenadas Then
    peleasOrdenadas = (nPeleas - nfal)
    peleasMejor = peleas
End If
End Sub

