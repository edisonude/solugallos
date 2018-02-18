Attribute VB_Name = "ModulePrimero"
Option Explicit
'Declaración del Api SetLayeredWindowAttributes que establece la transparencia al form
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
(ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    
'Recupera el estilo de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long) As Long
    
'Declaración del Api SetWindowLong necesaria para aplicar un estilo _
 al form antes de usar el Api SetLayeredWindowAttributes
  
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Public Const keyBD = "g900421553"

'Variables de aplicacion
Public SQL As String
Public pathBD As String
Public trabajandoCon As String
Public rstTrans As rdoResultset

'Variables de configuracion
Public voz As Integer
Public confTiempoArena As Integer
Public confTiempo As Integer
Public confTiempoEspuelas As Integer
Public confArenaIncluido As String
Public confReloj As String
Public confColor1 As String
Public confColor2 As String
Public confVerValor As String
Public confVerColorGallo As String
Public confVerAnuncios As String
Public confModoPrueba As String
Public reportesA4 As String
'variable para manejar formulario de ganador en peleas libres
Public fin As Integer

Public VCuerda1 As String
Public VColor1 As String
Public VCuerda2 As String
Public VColor2 As String
Public VValor As String
Public Duracion As String
Public HorInicio As String
Public FechaTrabajo As Date

'Variables para la red
Public mensajeRed As String
Public hayRed As Boolean

'Variables para conexion
Public strDSN As String
Public strUSER As String
Public strPASS As String
Public strCON As String
Public RDOCONEXION As rdoConnection
Public RDOAMBIENTE As rdoEnvironment

'hash for code
Public hash As String

'Conexion ADOB
Public cnn As New ADODB.Connection
Public rs As New ADODB.Recordset
Sub Main()

'Determina en que ambiente se va a trabajar, con esto tambien se cambia la bd
'con la cual se va a trabajar
trabajandoCon = "SoluPollos"
trabajandoCon = "SoluGallos"

'Encontrar ruta de BD
Dim fileConfBD  As String
If trabajandoCon = "SoluPollos" Then
    fileConfBD = App.Path & "\pathBD2"
Else
    fileConfBD = App.Path & "\pathBD"
End If

'Encontrar ruta de BD
Open fileConfBD For Input As #1
Dim Linea As String, Total As String
Do Until EOF(1)
    Line Input #1, pathBD
Loop
Close #1

    Dim O As DSN
    Set O = New DSN
    
BuscaDSN:
    Dim Lista_Dsn() As String
    O.ObtenerDSN Lista_Dsn()
    
    Dim i As Integer
    For i = LBound(Lista_Dsn) To UBound(Lista_Dsn) - 1
        If Lista_Dsn(i) = trabajandoCon Then
            GoTo CreaConexion
            Exit For
        End If
    Next
    
CreaDSN:

    O.ODBC_DSN_TIPO = Usuario
    With O
        .ODBC_DSN_NAME = trabajandoCon
        .ODBC_DRIVER_NAME = "Microsoft Access Driver (*.mdb)"
        .ODBC_DATA_SOURCE = pathBD
    Call O.Crear_Dsn
    End With

CreaConexion:

strDSN = trabajandoCon
strUSER = trabajandoCon
strPASS = "g900421553"
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
            "Data Source=" & pathBD & ";" & _
            "Jet OLEDB:Database Password=g900421553"
End With

Call cargarValoresConfig

'FrmGanador.Show
'frmMenu.Show
'FrmPantalla.Show


FrmInicio.Show
'frmBloqueo.Show

'FrmInfo.Show
'FrmValores.Show
'Form1.Show
'FrmPantalla.Show
'FrmNoCorrido.Show
End Sub


Public Sub White_skin(ByVal Formulario As Form)
Formulario.Skin1.LoadSkin App.Path & "\Skins\princ"
Formulario.Skin1.ApplySkin Formulario.hwnd
End Sub

Public Sub Uno_skin(ByVal Formulario As Form)
Formulario.Skin1.LoadSkin App.Path & "\Skins\tres"
Formulario.Skin1.ApplySkin Formulario.hwnd
End Sub
Public Sub Error_skin(ByVal Formulario As Form)
Formulario.Skin1.LoadSkin App.Path & "\Skins\error"
Formulario.Skin1.ApplySkin Formulario.hwnd
End Sub
Public Sub Cooper_skin(ByVal Formulario As Form)
Formulario.Skin1.LoadSkin App.Path & "\Skins\Copper.skn"
Formulario.Skin1.ApplySkin Formulario.hwnd
End Sub

Public Function ConEspacios(texto) As Integer
Dim Nespacios, i, j As Integer

i = 0
Nespacios = 0
j = 1
While i < j
    If InStr(texto, " ") Then
        Nespacios = Nespacios + 1
        texto = Mid(texto, InStr(texto, " ") + 1)
    Else
        i = j + 2
    End If
    i = i + 1
    j = j + 1
Wend
ConEspacios = Nespacios
End Function

Public Function Aplicar_Transparencia(ByVal hwnd As Long, valor As Integer) As Long
  
Dim Msg As Long
On Error Resume Next
If valor < 0 Or valor > 255 Then
   Aplicar_Transparencia = 1
Else
   Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
   Msg = Msg Or WS_EX_LAYERED
   SetWindowLong hwnd, GWL_EXSTYLE, Msg
   SetLayeredWindowAttributes hwnd, 0, valor, LWA_ALPHA
   Aplicar_Transparencia = 0
End If
If Err Then Aplicar_Transparencia = 2
End Function

Public Function cargarValoresConfig() As Boolean
Dim qry As New rdoQuery
Dim rs As rdoResultset
SQL = "Select * from AjustesGenerales"
Set qry.ActiveConnection = RDOCONEXION
    qry.SQL = SQL
    Set rs = qry.OpenResultset(rdOpenDynamic)
        confTiempo = rs("tiempo")
        confTiempoEspuelas = rs("tiempoEspuelas")
        confTiempoArena = rs("tiempoArena")
        confArenaIncluido = rs("arenaIncluido")
        confReloj = rs("direccionReloj")
        confColor1 = rs("cuerda1")
        confColor2 = rs("cuerda2")
        confVerValor = rs("valor")
        confVerColorGallo = rs("colorGallo")
        confVerAnuncios = rs("anuncios")
        voz = rs("voz")
        confModoPrueba = rs("modoPrueba")
        reportesA4 = rs("reportesA4")
qry.Close
End Function

Public Function tiempoaEntero(Tiempo As String) As Integer
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

Public Function pasarAHora(Tiempo As Integer) As String
Dim resultado As String
Dim hora As Integer
Dim minuto As Integer
Dim segundo As Integer

hora = Tiempo \ 3600
Tiempo = Tiempo - (hora * 3600)
minuto = Tiempo \ 60
Tiempo = Tiempo - (minuto * 60)
segundo = Tiempo

resultado = Format(minuto, "00") & ":" & Format(segundo, "00")
pasarAHora = resultado
End Function

Public Function Get_Numero_Serie(ByVal s_Drive As String) As Long
        
    Dim o_Fso As Scripting.FileSystemObject
    Dim o_Drive As Drive
      
    ' Creamos un nuevo objeto de tipo Scripting FileSystemObject
    Set o_Fso = New Scripting.FileSystemObject
      
    ' Si el Drive no es un vbnullstring
    If s_Drive <> "" Then
        ' Recuperamos el Drive para poder acceder _
         en las siguientes lineas
        Set o_Drive = o_Fso.GetDrive(s_Drive)
    End If
      
    With o_Drive
          
        ' Si está disponible
        If .IsReady Then
            Get_Numero_Serie = Not .SerialNumber
        Else
            MsgBox " No se puede acceder a la unidad ", vbCritical
            Get_Numero_Serie = -1
        End If
    End With
      
    ' Eliminamos los objetos instanciados
    Set o_Drive = Nothing
    Set o_Fso = Nothing
End Function
