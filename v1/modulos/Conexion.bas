Attribute VB_Name = "Conexion"
Option Explicit

'Variables de aplicacion
Public sinProducto  As String
Public sinCliente  As String

Public idUsuActual As Integer


'Variables para conexion
Public strDSN As String
Public strUSER As String
Public strPASS As String
Public strCON As String
Public RDOCONEXION As rdoConnection
Public RDOAMBIENTE As rdoEnvironment

'Conexion ADOB
Public cnn As New ADODB.Connection
Public rs As New ADODB.Recordset
Sub main()
    Dim O As DSN
    Set O = New DSN
    
BuscaDSN:
    Dim Lista_Dsn() As String
    O.ObtenerDSN Lista_Dsn()
    
    Dim i As Integer
    For i = LBound(Lista_Dsn) To UBound(Lista_Dsn) - 1
        If Lista_Dsn(i) = "Urano" Then
            GoTo CreaConexion
            Exit For
        End If
    Next
    
CreaDSN:

    O.ODBC_DSN_TIPO = Usuario
    With O
        .ODBC_DSN_NAME = "Urano"
        .ODBC_DRIVER_NAME = "Microsoft Access Driver (*.mdb)"
        .ODBC_DATA_SOURCE = App.Path & "\Datos.mdb"
    Call O.Crear_Dsn
    End With

CreaConexion:

strDSN = "Urano"
strUSER = "Urano"
strPASS = "c741852963c"
strCON = "DSN=" & strDSN & "; VID= " & strUSER & ";Pwd=" & strPASS & ";"
Set RDOAMBIENTE = rdoCreateEnvironment(strDSN, strUSER, strCON)
With RDOAMBIENTE
    .LoginTimeout = 10
    .CursorDriver = rdUseOdbc
Set RDOCONEXION = .OpenConnection(strDSN, rdDriverNoPrompt, False, strCON)
End With
Set O = Nothing

'CONEXION CON ADOB
With cnn
    .CursorLocation = adUseClient
    .Open "Provider=Microsoft.Jet.OLEDB.4.0; " & _
            "Data Source=" & App.Path & "\Datos.mdb" & ";" & _
            "Jet OLEDB:Database Password=c741852963c"
End With

sinProducto = App.Path & "\Productos\sinProducto.gif"
sinCliente = App.Path & "\Clientes\sinCliente.gif"
idUsuActual = 1

'frmMenu.Show
frmSplash.Show
End Sub

Public Sub Skin_skin(ByVal Formulario As Form)
Formulario.Skin1.LoadSkin App.Path & "\Skins\Copper.skn"
Formulario.Skin1.ApplySkin Formulario.hwnd
End Sub
