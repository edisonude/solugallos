VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClasificacion 
   Caption         =   "Clasificación"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin VB.CommandButton cmdTotal 
         Caption         =   "Total"
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
         Left            =   4200
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton btnGuardar 
         Caption         =   "Refrescar"
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
         Left            =   8760
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdImprimir 
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
         Height          =   375
         Left            =   6480
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   615
         Left            =   360
         OleObjectBlob   =   "frmClasificacion.frx":0000
         TabIndex        =   3
         Top             =   480
         Width           =   4095
      End
      Begin MSComctlLib.ListView lista 
         Height          =   7380
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   10545
         _ExtentX        =   18600
         _ExtentY        =   13018
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuerda"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Puntos"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "PG"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "PE"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "PP"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Tiempo"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   -120
      OleObjectBlob   =   "frmClasificacion.frx":0090
      Top             =   120
   End
End
Attribute VB_Name = "frmClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pg As Integer
Dim pe As Integer
Dim pp As Integer
Dim np As Integer
Dim Total As Boolean

Private Sub btnGuardar_Click()
Unload Me
frmClasificacion.Show
End Sub

Private Sub cmdImprimir_Click()
    Dim reportToUse As String
    reportToUse = IIf(reportesA4 = "Si", "ConFinClasificacion_A4", "ConFinClasificacion")

    Dim oAcces As Access.Application
    Set oAcces = New Access.Application
    
    oAcces.OpenCurrentDatabase pathBD, False, keyBD
    oAcces.Visible = False

    If Total Then
        oAcces.DoCmd.OpenReport "ConFinClasificacionTotal", acViewPreview
    Else
        oAcces.DoCmd.OpenReport reportToUse, acViewPreview
    End If

    oAcces.DoCmd.PrintOut acPrintAll
    oAcces.CloseCurrentDatabase
    oAcces.Quit
    Set oAcces = Nothing
End Sub

Private Sub cmdTotal_Click()
Total = True

Dim puntosT As Integer
Dim pgT As Integer
Dim ppT As Integer
Dim peT As Integer
Dim tiempo1T As Integer
Dim tiempoC As Integer
Dim tiempo2T As String
Dim nPeleasT As Integer

Dim p As Integer
Dim t As Integer
Dim tp As Integer
Dim dif As Integer
Dim t2 As String

Dim idC As Integer

Dim ptt As Integer

Dim SQL2 As String


'Elimina la clasificacion
SQL = "Delete * from ClasificacionT"
Call guardarRDO

Dim rst As New ADODB.Recordset
SQL = "Select * from ClasificacionA"
rst.Open SQL, cnn, adOpenStatic, adLockOptimistic

Dim rst2 As New ADODB.Recordset
SQL2 = "Select * from Clasificacion"
rst2.Open SQL2, cnn, adOpenStatic, adLockOptimistic
        
    While rst.EOF = False
        idC = Val(rst("idCuerda"))
        If idC = -1 Then rst.MoveNext
        idC = Val(rst("idCuerda"))
        
'        SQL = "Select * from Clasificacion"
'        Set qry2.ActiveConnection = RDOCONEXION
'            qry2.SQL = SQL
'            Set rst2 = qry2.OpenResultset(rdOpenDynamic)
'            rst2.MoveFirst
'            While idC <> Val(rst2("idCuerda"))
'                rst2.MoveNext
'            Wend


        
         While idC <> Val(rst2("idCuerda"))
                rst2.MoveNext
            Wend


                puntosT = Val(rst("puntos")) + Val(rst2("puntos"))
                tiempo1T = Val(rst("tiempo1")) + Val(rst2("tiempo1"))
                nPeleasT = Val(rst("nPeleas")) + Val(rst2("nPeleas"))
                pgT = Val(rst("pg")) + Val(rst2("pg"))
                peT = Val(rst("pe")) + Val(rst2("pe"))
                ppT = Val(rst("pp")) + Val(rst2("pp"))
                tiempoC = tiempo1T
                tiempo2T = pasarAHora(tiempoC)
                
                rst2.MoveFirst
                
                SQL = "Insert into ClasificacionT values (" & idC & "," & puntosT & "," & tiempo1T & ",'" & tiempo2T & "'," & nPeleasT & "," & pgT & "," & peT & "," & ppT & ")"
                Call guardarRDO
                
'                qry2.Close
        rst.MoveNext
    Wend

Call cargarListaTotal
End Sub

Private Function pasarAHora(Tiempo As Integer) As String
Dim resultado As String
Dim hora As Integer
Dim minuto As Integer
Dim segundo As Integer

hora = Tiempo \ 3600
Tiempo = Tiempo - (hora * 3600)
minuto = Tiempo \ 60
Tiempo = Tiempo - (minuto * 60)
segundo = Tiempo

resultado = Format(hora, "00") & ":" & Format(minuto, "00") & ":" & Format(segundo, "00")
pasarAHora = resultado
End Function


Private Sub Form_Load()
Cooper_skin Me

Total = False

Dim ancho As Integer
ancho = Me.lista.Width

Me.lista.ColumnHeaders(1).Width = ancho * 0.45
Me.lista.ColumnHeaders(2).Width = ancho * 0.1
Me.lista.ColumnHeaders(3).Width = ancho * 0.1
Me.lista.ColumnHeaders(4).Width = ancho * 0.1
Me.lista.ColumnHeaders(5).Width = ancho * 0.1
Me.lista.ColumnHeaders(6).Width = ancho * 0.12


Dim p As Integer
Dim t As Integer
Dim tp As Integer
Dim dif As Integer
Dim t2 As String

Dim idC As Integer

Dim ptt As Integer

Dim qry As New rdoQuery
Dim rst As rdoResultset

'Elimina la clasificacion
SQL = "Delete * from Clasificacion"
Call guardarRDO

SQL = "Select * from Cuerdas"
Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    rst.MoveFirst
    While rst.EOF = False
        idC = Val(rst("idCuerda"))
        If idC = -1 Then rst.MoveNext
        idC = Val(rst("idCuerda"))
        Call obtenerPuntajes(idC)
        
        ptt = pg * 2 + pe
        
        t = obtenerTiempo(idC)
        tp = t \ 60
        dif = t - (tp * 60)
        t2 = tp & ":" & Format(dif, "00")
        
        SQL = "Insert into Clasificacion values (" & idC & "," & ptt & "," & t & ",'" & t2 & "'," & np & "," & pg & "," & pe & "," & pp & ")"
        Call guardarRDO
        rst.MoveNext
    Wend

Call cargarLista

End Sub

Private Sub cargarLista()
Dim qry As New rdoQuery
Dim rst As rdoResultset
lista.ListItems.Clear
SQL = "Select * from ConFinClasificacion"
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        Set rst = qry.OpenResultset(rdOpenDynamic)
            While rst.EOF = False
                Set li = lista.ListItems.Add(, , rst(0))
                    li.SubItems(1) = rst(1)
                    li.SubItems(2) = rst(2)
                    li.SubItems(3) = rst(3)
                    li.SubItems(4) = rst(4)
                    li.SubItems(5) = rst(5)
                rst.MoveNext
            Wend
    qry.Close
End Sub

Private Sub cargarListaTotal()
Dim qry As New rdoQuery
Dim rst As rdoResultset
lista.ListItems.Clear
SQL = "Select * from ConFinClasificacionTotal"
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        Set rst = qry.OpenResultset(rdOpenDynamic)
            While rst.EOF = False
                Set li = lista.ListItems.Add(, , rst(0))
                    li.SubItems(1) = rst(1)
                    li.SubItems(2) = rst(2)
                    li.SubItems(3) = rst(3)
                    li.SubItems(4) = rst(4)
                    li.SubItems(5) = rst(5)
                rst.MoveNext
            Wend
    qry.Close
End Sub

Private Sub obtenerPuntajes(c As Integer)
Dim qry As New rdoQuery
Dim rst As rdoResultset
Dim puntosCuerda As Integer
Dim tiempoCuerda As Integer
Dim comodin As Boolean

np = 0
pg = 0
pe = 0
pp = 0

SQL = "Select * from resumenPeleas where jugada='si' and (idCuerda=" & c & " or idCuerda2=" & c & ")"
Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    If rst.RowCount > 0 Then
        rst.MoveFirst
        While rst.EOF = False
            If rst("idCuerda") = c Then
                If esComodin(rst("idGallo")) Then
                    GoTo sigue
                End If
            End If
            
            If rst("idCuerda2") = c Then
                If esComodin(rst("idGallo2")) Then
                    GoTo sigue
                End If
            End If
        
            If rst("tiempo") <> "" Then
                np = np + 1
                If rst("idCuerda") = c And rst("puntos1") = 2 Then
                    pg = pg + 1
                Else
                    If rst("idCuerda2") = c And rst("puntos2") = 2 Then
                        pg = pg + 1
                    Else
                        If rst("puntos1") = 1 Then
                            pe = pe + 1
                        Else
                            pp = pp + 1
                        End If
                    End If
                End If
            End If
sigue:
            rst.MoveNext
        Wend
    End If
    qry.Close
End Sub

Private Function obtenerTiempo(c As Integer) As Integer
Dim qry As New rdoQuery
Dim rst As rdoResultset
Dim puntosCuerda As Integer
Dim tiempoCuerda As Integer

SQL = "Select * from resumenPeleas"
Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    If rst.RowCount <= 0 Then
       ' MsgBox "No existen peleas"
        Exit Function
    End If
        rst.MoveFirst

puntosCuerda = 0
tiempoCuerda = 0
While rst.EOF = False
    If rst("idCuerda") = c Then
    
    If esComodin(rst("idGallo")) Then
        GoTo sigue
    End If
                
    If Not IsNull(rst("puntos1")) Then
        If Val(rst("puntos1")) > 0 Then
            tiempoCuerda = tiempoCuerda + tiempoaEntero(IIf(IsNull(rst("tiempo")), 0, rst("tiempo")))
            'puntosCuerda = puntosCuerda + Val(rst("puntos1"))
        End If
    End If
    End If
    
    If rst("idCuerda2") = c Then
    
    If esComodin(rst("idGallo2")) Then
        GoTo sigue
    End If
                
    If Not IsNull(rst("puntos2")) Then
        If Val(rst("puntos2")) > 0 Then
            tiempoCuerda = tiempoCuerda + tiempoaEntero(IIf(IsNull(rst("tiempo")), 0, rst("tiempo")))
            'puntosCuerda = puntosCuerda + Val(rst("puntos2"))
        End If
    End If
    End If
sigue:
    rst.MoveNext
Wend
qry.Close
obtenerTiempo = tiempoCuerda
End Function

Private Function tiempoaEntero(Tiempo As String) As Integer
If Tiempo = "" Then
    tiempoaEntero = 0
    Exit Function
End If

Dim segundos As Integer
Dim minutos As Integer

minutos = Val(Left(Tiempo, 2))
segundos = Val(Right(Tiempo, 2))

segundos = segundos + (minutos * 60)
tiempoaEntero = segundos
End Function

Private Function esComodin(idG As Integer) As Boolean
Dim rst As rdoResultset
Dim qry As New rdoQuery
SQL = "Select * from gallos where idGallo=" & idG & " and comodin='si'"
Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rst = qry.OpenResultset(rdOpenDynamic)
    If rst.RowCount > 0 Then
        esComodin = True
    Else
        esComodin = False
    End If
End Function

