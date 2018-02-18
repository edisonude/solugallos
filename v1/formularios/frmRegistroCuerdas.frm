VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegistroCuerdas 
   BackColor       =   &H00555555&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Información de la cuerda y los invitados"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   360
      TabIndex        =   13
      Top             =   240
      Width           =   8775
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
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox tPropietario 
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
         Left            =   2280
         TabIndex        =   1
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox tDocumento 
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
         Left            =   2280
         TabIndex        =   3
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox tMail 
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
         Left            =   2280
         TabIndex        =   2
         Top             =   1440
         Width           =   4935
      End
      Begin VB.TextBox tCelular 
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
         Left            =   2280
         TabIndex        =   4
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         Caption         =   "Invitados"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   360
         TabIndex        =   14
         Top             =   3360
         Width           =   7935
         Begin VB.TextBox tDocInvitado 
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
            Left            =   240
            TabIndex        =   9
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox tNombreInvitado 
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
            Left            =   2640
            TabIndex        =   10
            Top             =   840
            Width           =   4215
         End
         Begin VB.CommandButton picAdd 
            Caption         =   "+"
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
            Left            =   7080
            TabIndex        =   11
            Top             =   720
            Width           =   615
         End
         Begin MSComctlLib.ListView listaInvitados 
            Height          =   1860
            Left            =   240
            TabIndex        =   12
            Top             =   1800
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   3281
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Documento"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre"
               Object.Width           =   8819
            EndProperty
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   5
            Left            =   240
            OleObjectBlob   =   "frmRegistroCuerdas.frx":0000
            TabIndex        =   15
            Top             =   360
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   6
            Left            =   2640
            OleObjectBlob   =   "frmRegistroCuerdas.frx":006A
            TabIndex        =   16
            Top             =   360
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
            Height          =   375
            Index           =   7
            Left            =   240
            OleObjectBlob   =   "frmRegistroCuerdas.frx":00CE
            TabIndex        =   17
            Top             =   1320
            Width           =   1935
         End
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
         Left            =   1920
         TabIndex        =   7
         Top             =   7320
         Width           =   2175
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
         Left            =   4200
         TabIndex        =   8
         Top             =   7320
         Width           =   2175
      End
      Begin VB.ComboBox tPais 
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
         Left            =   2280
         TabIndex        =   5
         Top             =   2880
         Width           =   2295
      End
      Begin VB.ComboBox tCiudad 
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
         Left            =   5760
         TabIndex        =   6
         Top             =   2880
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   0
         Left            =   360
         OleObjectBlob   =   "frmRegistroCuerdas.frx":014A
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   1
         Left            =   360
         OleObjectBlob   =   "frmRegistroCuerdas.frx":01B2
         TabIndex        =   19
         Top             =   960
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   2
         Left            =   360
         OleObjectBlob   =   "frmRegistroCuerdas.frx":0224
         TabIndex        =   20
         Top             =   1920
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   3
         Left            =   360
         OleObjectBlob   =   "frmRegistroCuerdas.frx":028E
         TabIndex        =   21
         Top             =   1440
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   4
         Left            =   360
         OleObjectBlob   =   "frmRegistroCuerdas.frx":02F2
         TabIndex        =   22
         Top             =   2400
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   8
         Left            =   360
         OleObjectBlob   =   "frmRegistroCuerdas.frx":0358
         TabIndex        =   23
         Top             =   2880
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel etiqueta 
         Height          =   375
         Index           =   9
         Left            =   4800
         OleObjectBlob   =   "frmRegistroCuerdas.frx":03B8
         TabIndex        =   24
         Top             =   2880
         Width           =   855
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmRegistroCuerdas.frx":041C
      Top             =   0
   End
End
Attribute VB_Name = "frmRegistroCuerdas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Dim li As ListItem
Dim frente As Integer

Private Sub btnGuardar_Click()
If tPropietario.Text = "" Or tCuerda.Text = "" Then
    menFaltanDatos
    Exit Sub
End If

If Me.tPais = "" Or Me.tCiudad = "" Then
    MsgBox "Debe seleccionar el pais y la ciudad"
    Exit Sub
End If

If cuerdaDuplicada Then
    seleccionarTexto tCuerda
    menCuerdaDuplicada
    Exit Sub
End If

Dim consecutivo As Integer
Dim conIvitado As Integer
consecutivo = HallaConsecutivo("Select Max(idCuerda) As NConsecutivo from Cuerdas")

    Call actualizarPais
    Call actualizarCiudad

Dim ciudad As Integer
ciudad = dimeCiudad(Me.tPais, Me.tCiudad)

If ciudad = -1 Then
    MsgBox "Existe un problema con la ciudad, verifique los datos", vbCritical
    Exit Sub
End If

SQL = "insert into Cuerdas values(" & consecutivo & ",'" & UCase(tCuerda.Text) & "','" & Me.tDocumento & "','" & UCase(Me.tPropietario.Text) & _
        "','" & Me.tMail.Text & "','" & Me.tCelular.Text & "','" & Format(Date, "dd/mm/yyyy") & "'," & ciudad & "," & frente & ")"
            
If guardarRDO Then
    With listaInvitados
        For i = 1 To .ListItems.Count
            conIvitado = HallaConsecutivo("Select Max(id) As NConsecutivo from Invitados")
        
            SQL = "insert into Invitados values(" & conIvitado & ",'" & .ListItems.Item(i) & "','" & .ListItems.Item(i).SubItems(1) & "')"
            
            If guardarRDO Then
                SQL = "insert into InvitadosxCuerda values(" & consecutivo & "," & Me.tDocumento & "," & conIvitado & ")"
                Call guardarRDO
            End If
        Next
    End With
    
    menGuardadoExitoso
    frente = 0
    Unload Me
    frmRegistroCuerdas.Show
Else
    Call menGuardadoFallo
End If
End Sub

Private Sub actualizarCiudad()
On Error Resume Next
    Dim consecutivo As Integer
    consecutivo = HallaConsecutivo("Select Max(idCiudad) As NConsecutivo from ciudad")
    
    SQL = "insert into ciudad values(" & consecutivo & "," & dimePais(Me.tPais) & ",'" & UCase(tCiudad) & "')"
    Call guardarRDO
End Sub
Private Sub actualizarPais()
On Error Resume Next
    Dim consecutivo As Integer
    consecutivo = HallaConsecutivo("Select Max(id) As NConsecutivo from pais")

    SQL = "insert into Pais values(" & consecutivo & ",'" & UCase(tPais) & "')"
    Call guardarRDO
End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Cooper_skin Me

Me.Top = frmMenu.ubicacion.Top + 350
Me.Left = frmMenu.ubicacion.Left
Call cargarPaises
End Sub

Private Sub cargarPaises()
Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Select * from pais"
    Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    While rst.EOF = False
        Me.tPais.AddItem rst("pais")
        rst.MoveNext
    Wend
    qry.Close
End Sub

Private Function cuerdaDuplicada() As Boolean
Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Select * from Cuerdas where Cuerda='" & Me.tCuerda.Text & "'"
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        Set rst = qry.OpenResultset(rdOpenDynamic)
            If (rst.RowCount >= 1) Then
                cuerdaDuplicada = True
            Else
                cuerdaDuplicada = False
            End If
    qry.Close
End Function

Private Sub limpiarCampos()
frente = 0
Me.tCuerda = ""
Me.tPropietario = ""
Me.tDocumento = ""
Me.tMail = ""
Me.tCelular = ""
Me.tDocInvitado = ""
Me.tNombreInvitado = ""
Me.tPais = ""
Me.tCiudad = ""
Me.listaInvitados.ListItems.Clear

Call cargarPaises
End Sub

Private Sub listaInvitados_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.listaInvitados.ListItems.Count <= 0 Then Exit Sub
If KeyCode = 46 Then
    listaInvitados.ListItems.Remove (listaInvitados.SelectedItem.Index)
End If
End Sub

Private Sub picAdd_Click()
If Me.tNombreInvitado = "" Then
    MsgBox "Debe digitar minimo el nombre del invitado", vbCritical, "Sin invitado"
    Me.tNombreInvitado.SetFocus
    Exit Sub
End If

 Set li = listaInvitados.ListItems.Add(, , Me.tDocInvitado)
    li.SubItems(1) = Me.tNombreInvitado

Me.tNombreInvitado = ""
Me.tDocInvitado = ""
End Sub


Private Sub tCelular_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
End Sub

Private Sub tCiudad_GotFocus()
SendMessageLong tCiudad.hwnd, &H14F, True, 0
End Sub

Private Sub tCuerda_LostFocus()
Me.tCuerda = UCase(tCuerda)

Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Select * from Cuerdas where Cuerda='" & Me.tCuerda & "'"
    Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    If rst.RowCount >= 1 Then
        If MsgBox("Ya existe otra cuerda con este mismo nombre." & vbCrLf & "¿Desea crearse como un frente de esta?", vbQuestion + vbYesNo, "Crear frente") = vbYes Then
            frente = rst("idCuerda")
            
            Me.tCuerda = Me.tCuerda & " " & contadorCuerdas()
            Me.tPropietario = rst("Propietario")
            Me.tMail = IIf(IsNull(rst("Email")), "", rst("Email"))
            Me.tDocumento = IIf(IsNull(rst("Documento")), "", rst("Documento"))
            Me.tCelular = IIf(IsNull(rst("Celular")), "", rst("Celular"))
            
            Call llenarPaisCiudad(Me.tCiudad, Me.tPais, rst("Ciudad"))
            
            'Me.tCiudad = dimeNombreCiudad(CInt(rst("Ciudad")))
        Else
            frente = 0
            Me.tCuerda = ""
        End If
    End If
    qry.Close

End Sub

Private Function contadorCuerdas()
Dim qry As New rdoQuery
Dim rst As rdoResultset

SQL = "Select * from Cuerdas where Cuerda like '" & UCase(Me.tCuerda) & "%'"
    Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
       contadorCuerdas = rst.RowCount + 1
    qry.Close
End Function

Private Sub tDocumento_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
End Sub

Private Sub tPais_Click()
Dim qry As New rdoQuery
Dim rst As rdoResultset

    SQL = "SELECT p.pais, c.idCiudad, c.ciudad as ciud FROM pais p INNER JOIN ciudad c ON p.id = c.idPais WHERE p.pais ='" & tPais & "'"

    Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    Me.tCiudad.Clear
    While rst.EOF = False
        Me.tCiudad.AddItem rst("ciud")
        rst.MoveNext
    Wend
    qry.Close
End Sub

Private Sub tPais_GotFocus()
SendMessageLong tPais.hwnd, &H14F, True, 0
End Sub


