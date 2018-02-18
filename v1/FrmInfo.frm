VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   14385
   ClientLeft      =   885
   ClientTop       =   540
   ClientWidth     =   25200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14385
   ScaleWidth      =   25200
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   45
      Top             =   2625
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   30
      OleObjectBlob   =   "FrmInfo.frx":0000
      Top             =   210
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   11460
      Left            =   480
      ScaleHeight     =   11430
      ScaleWidth      =   24540
      TabIndex        =   2
      Top             =   315
      Width           =   24570
      Begin VB.PictureBox PicIngrese 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   2520
         ScaleHeight     =   840
         ScaleWidth      =   4575
         TabIndex        =   5
         Top             =   4860
         Width           =   4575
      End
      Begin VB.ComboBox CmbColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1245
         ItemData        =   "FrmInfo.frx":0234
         Left            =   495
         List            =   "FrmInfo.frx":0236
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "Seleccione color..."
         Top             =   1515
         Visible         =   0   'False
         Width           =   9180
      End
      Begin VB.ComboBox CmbCuerdas 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1245
         ItemData        =   "FrmInfo.frx":0238
         Left            =   480
         List            =   "FrmInfo.frx":023A
         Sorted          =   -1  'True
         TabIndex        =   0
         Text            =   "Seleccione cuerda..."
         Top             =   1515
         Width           =   9660
      End
      Begin VB.TextBox TValor2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000763EB&
         Height          =   1350
         Left            =   480
         TabIndex        =   6
         Top             =   1455
         Visible         =   0   'False
         Width           =   9660
      End
      Begin VB.TextBox TValor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000763EB&
         Height          =   1350
         Left            =   480
         TabIndex        =   3
         Top             =   1455
         Visible         =   0   'False
         Width           =   9660
      End
      Begin VB.Label LTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUERDA #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   60
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1800
         Left            =   -15
         TabIndex        =   4
         Top             =   -195
         Width           =   8805
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   5445
         Left            =   11520
         Picture         =   "FrmInfo.frx":023C
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   5520
      End
   End
   Begin VB.Label TAcoolor2 
      Height          =   870
      Left            =   75
      TabIndex        =   7
      Top             =   1485
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "FrmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim Tmp As Long
'Dim val, val1 As Integer
'Dim out As Boolean
'
'Private Sub CmbColor_GotFocus()
'If Me.TAcoolor2 <> "" Then Exit Sub
'Tmp = SendMessage(CmbColor.hwnd, &H14F, 1, ByVal 0&)
''Me.Text2.Enabled = True
'End Sub
'
'Private Sub CmbColor_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 27 Then End
'If KeyCode = 33 Then KeyCode = 38
'If KeyCode = 34 Then KeyCode = 40
'
'
'If KeyCode = 66 Or KeyCode = 98 Then
'    KeyCode = 0
'    If LTitulo.Caption = "COLOR #1" Then
'        Call cuerda1
'    Else
'        Call cuerda2
'    End If
'End If
'
'If KeyCode = 116 Or KeyCode = 13 Then
'    KeyCode = 0
'    If Me.CmbColor = "Seleccione color..." Or Me.CmbCuerdas = "" Then
'        frmerrores.lberror = "SELECCIONE EL COLOR"
'        CmbColor.SetFocus
'        frmerrores.Show
'        Exit Sub
'    End If
'
'    If LTitulo.Caption = "COLOR #1" Then
'        VColor1 = Me.CmbColor
'        KeyCode = 13
'        Call cuerda2
'    Else
'        VColor2 = Me.CmbColor
'        KeyCode = 13
'        Call AValor
'    End If
'End If
'End Sub
'
'Private Sub CmbColor_KeyPress(KeyAscii As Integer)
'If KeyAscii = 66 Or KeyAscii = 98 Then KeyAscii = 0
'End Sub
'
'Private Sub CmbCuerdas_GotFocus()
'If Me.TAcoolor2 <> "" Then Exit Sub
'Tmp = SendMessage(CmbCuerdas.hwnd, &H14F, 10, ByVal 0&)
'End Sub
'
'Private Sub CmbCuerdas_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 33 Then KeyCode = 38
'If KeyCode = 34 Then KeyCode = 40
'
'If KeyCode = 66 Or KeyCode = 98 Then
'    KeyCode = 0
'    If LTitulo.Caption = "CUERDA #2" Then
'        Call Color1
'    End If
'End If
'
'If KeyCode = 116 Or KeyCode = 13 Then
'    KeyCode = 0
'    If Me.CmbCuerdas = "Seleccione cuerda..." Or Me.CmbCuerdas = "" Then
'        frmerrores.lberror = "SELECCIONE LA CUERDA"
'        CmbCuerdas.SetFocus
'        frmerrores.Show
'        Exit Sub
'    End If
'
'    If LTitulo.Caption = "CUERDA #1" Then
'        VCuerda1 = Me.CmbCuerdas
'        KeyCode = 13
'        Call Color1
'    Else
'        VCuerda2 = Me.CmbCuerdas
'        KeyCode = 13
'        Call Color2
'    End If
'End If
'End Sub
'
'Private Sub CmbCuerdas_KeyPress(KeyAscii As Integer)
'If KeyAscii = 66 Or KeyAscii = 98 Then KeyAscii = 0
'End Sub
'
'Private Sub Command1_Click()
'FrmEspuelas.Show
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'If Me.CmbColor.Visible = True And KeyAscii = 100 Then
'    Tmp = SendMessage(CmbColor.hwnd, &H14F, 1, ByVal 0&)
'End If
'If KeyAscii = 27 Then
'    out = True
'    Unload Me
'End If
'End Sub
'
'Private Sub Form_Load()
'White_skin Me
'out = False
'
'Me.Left = 15015
''val = 0
''val1 =
'
'
''Carga los colores de los gallos
'Dim qryv As New rdoQuery
'Dim rsv As rdoResultset
'Dim SQL As String
'
'    SQL = "Select * from Colores"
'    Set qryv.ActiveConnection = RDOCONEXION
'        qryv.SQL = SQL
'        Set rsv = qryv.OpenResultset(rdOpenDynamic)
'            rsv.MoveLast
'            rsv.MoveFirst
'
'            While rsv.EOF = False
'                Me.CmbColor.AddItem (rsv("Color"))
'                rsv.MoveNext
'            Wend
'
''Carga los colores de los gallos
'Dim qryv2 As New rdoQuery
''
'
'    SQL = "Select * from Cuerdas"
'    Set qryv2.ActiveConnection = RDOCONEXION
'        qryv2.SQL = SQL
'        Set rsv = qryv2.OpenResultset(rdOpenDynamic)
'            rsv.MoveLast
'            rsv.MoveFirst
'
'            While rsv.EOF = False
'                Me.CmbCuerdas.AddItem (rsv("Cuerda"))
'                rsv.MoveNext
'            Wend
'
'TValor = 0
'End Sub
'
'
''Private Sub TCuerda_Change()
''If TCuerda = "" Then Exit Sub
''PicIngrese.Visible = False
''If Len(TCuerda) >= 19 Then
''    Me.TCuerda.Height = 2325
''Else
''    Me.TCuerda.Height = 1155
''End If
''End Sub
'
'Private Sub TCuerda_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'If out = True Then Unload FrmBar
'End Sub
'
'Private Sub TAcoolor2_Change()
'If TAcoolor2 <> "" Then
'    Call AValor
'Timer1.Enabled = True
'End If
'End Sub
'
'Private Sub Text1_GotFocus()
''Dim Con As Integer
''Con = ConEspacios(Me.TCuerda)
''Con = Len(Me.TCuerda) - Con
''If Con < 3 Then
''    frmerrores.lberror = "Debe ingresar el nombre completo de la cuerda" & Chr(13) & _
''    "para poder continuar."
''    TCuerda.SetFocus
''    PicIngrese.Visible = True
''    frmerrores.Show
''    Exit Sub
''End If
''
''If Cuerda1 = "" Then
''    Cuerda1 = Me.TCuerda
''    Me.LTitulo = "COLOR #1"
''Else
''    Cuerda2 = Me.TCuerda
''    Me.LTitulo = "COLOR #2"
''End If
''
''    Me.TCuerda = ""
''    Me.TCuerda.Visible = False
''    Me.Text1.Enabled = False
''    Me.CmbColor.Visible = True
''    Me.CmbColor.SetFocus
'End Sub
'
'Private Sub Text2_GotFocus()
''If CmbColor = "Seleccione color..." Then
''    frmerrores.lberror = "Debe seleccionar el color del gallo" & Chr(13) & _
''    "para poder continuar."
''    CmbColor.SetFocus
''    frmerrores.Show
''    Exit Sub
''End If
''
''If Color1 = "" Then
''    Color1 = Me.CmbColor
''    Me.LTitulo = "CUERDA #2"
''    Me.TCuerda.Visible = True
''    PicIngrese.Visible = True
''    Me.Text1.Enabled = True
''    Me.TCuerda.SetFocus
''Else
''    Color2 = Me.CmbColor
''    Me.LTitulo = "VALOR"
''    FrmValores.Show
''    FrmValores.Left = TValor.Left
''    FrmValores.Top = TValor.Top + 200
''    TValor.Visible = True
''    TValor2.Visible = True
''    Text1.Enabled = False
''    Text2.Enabled = False
''    Text3.Enabled = True
''    Me.TValor.SetFocus
''End If
''
''    'Comprueba si el color es nuevo
''    Dim i As Integer
''    i = 0
''    While i < Me.CmbColor.ListCount
''        If Me.CmbColor = Me.CmbColor.List(i) Then
''            i = Me.CmbColor.ListCount + 2
''        Else: i = i + 1
''        End If
''    Wend
''
''    If i <> Me.CmbColor.ListCount + 2 Then
''        Dim sql As String
''        Dim qry As New rdoQuery
''
''        sql = "Insert into Colores values ('" & Me.CmbColor & "')"
''        Set qry.ActiveConnection = RDOCONEXION
''            qry.sql = sql
''            qry.Execute
''            qry.Close
''            Me.CmbColor = "Seleccione color..."
''        'Se refresca la lista
''
''            Dim qryv As New rdoQuery
''            Dim rsv As rdoResultset
''            Me.CmbColor.Clear
''            sql = "Select * from Colores"
''            Set qryv.ActiveConnection = RDOCONEXION
''                qryv.sql = sql
''                Set rsv = qryv.OpenResultset(rdOpenDynamic)
''                    rsv.MoveLast
''                    rsv.MoveFirst
''                    While rsv.EOF = False
''                        Me.CmbColor.AddItem (rsv("Color"))
''                        rsv.MoveNext
''                    Wend
''    End If
''
''    Me.CmbColor.Visible = False
''    Me.Text2.Enabled = False
'End Sub
'
'
'Private Sub Timer1_Timer()
'Call AValor
'TAcoolor2 = ""
'Timer1.Enabled = False
'End Sub
'
'
'Private Sub TValor_Change()
'If TValor = "" Then
'    TValor2 = FormatCurrency(0, 0, True, True, True)
'End If
'    TValor2 = FormatCurrency(val, 0, True, True, True)
'End Sub
'
'Private Sub TValor_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 27 Or KeyCode = 116 Then Unload Me
'
'If KeyCode = 66 Or KeyCode = 98 Then
'    KeyCode = 0
'        Call Color2
'End If
'
'If KeyCode = 33 Or KeyCode = 40 Then
'    val = val + 50000
'End If
'
'If KeyCode = 34 Then
'    val = val - 10000
'End If
'
'Me.TValor = val
'VValor = val
'If KeyCode = 116 Or KeyCode = 13 Then
'    If Me.TValor = "0" Or TValor = "" Then
'        frmerrores.lberror = "SELECCIONE EL VALOR"
'        'TValor.SetFocus
'        frmerrores.Show
'        Exit Sub
'    End If
'
'    VValor = val
'    val = 0
'    KeyCode = 13
'    Call APantalla
'End If
'
'
'End Sub
'
'Private Sub TValor_KeyPress(KeyAscii As Integer)
'KeyAscii = SoloNumeros(KeyAscii)
'End Sub
'
'Private Sub TValor2_GotFocus()
'Me.TValor.SetFocus
'End Sub
'
'
'Private Function Color1()
'    Me.LTitulo = "COLOR #1"
'    Me.CmbCuerdas.Visible = False
'    Me.CmbColor.Visible = True
'    Me.TValor.Visible = False
'    Me.TValor2.Visible = False
'    Me.CmbColor.SetFocus
'End Function
'
'Private Function Color2()
'    Me.LTitulo = "COLOR #2"
'    Me.CmbCuerdas.Visible = False
'    Me.CmbColor.Visible = True
'    Me.TValor.Visible = False
'    Me.TValor2.Visible = False
'    Me.CmbColor.SetFocus
'End Function
'
'Private Function cuerda2()
'    Me.LTitulo = "CUERDA #2"
'    Me.CmbCuerdas.Visible = True
'    Me.CmbColor.Visible = False
'    Me.CmbCuerdas.SetFocus
'End Function
'
'Private Function cuerda1()
'    Me.LTitulo = "CUERDA #1"
'    Me.CmbCuerdas.Visible = True
'    Me.CmbColor.Visible = False
'    Me.CmbCuerdas.SetFocus
'End Function
'
'Private Function AValor()
'    Me.LTitulo = "VALOR"
'    Me.CmbCuerdas.Visible = False
'    Me.CmbColor.Visible = False
'    Me.TValor2.Visible = True
'    Me.TValor.Visible = True
'    Me.TValor.Enabled = True
'    Me.TValor.SetFocus
'End Function
'
'Private Function APantalla()
'
'FrmPantalla.tCuerda1 = VCuerda1
'FrmPantalla.TColor1 = VColor1
'FrmPantalla.tCuerda2 = VCuerda2
'FrmPantalla.TColor2 = VColor2
'FrmPantalla.TValor = FormatCurrency(VValor, 0, True, True, True)
''FrmPantalla.TCuerda1 = Cuerda1
'FrmPantalla.Show
'Me.Hide
'End Function
'
