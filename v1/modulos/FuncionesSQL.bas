Attribute VB_Name = "FuncionesSQL"
'Halla un consecutivo
Public Function HallaConsecutivo(Csql) As Integer
    Dim qry As New rdoQuery
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = Csql
    Set rst = qry.OpenResultset(rdOpenDynamic)
        If IsNull(rst("NConsecutivo")) Then
            HallaConsecutivo = 1
        Else
            HallaConsecutivo = rst("NConsecutivo") + 1
            HallaConsecutivo = IIf(HallaConsecutivo = 0, 1, HallaConsecutivo)
        End If
        qry.Close
End Function

Public Function guardarRDO() As Boolean
On Error GoTo Control
    Dim qry As New rdoQuery
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = SQL
        qry.Execute
            guardarRDO = True
        qry.Close
        Exit Function
Control:
    'MsgBox Err.Description
    guardarRDO = False
End Function

Public Sub addColor(color As String)
On Error Resume Next
    Dim qry As New rdoQuery
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = "INSERT INTO Colores values('" & color & "')"
        qry.Execute
        qry.Close
End Sub
Public Sub addCresta(cresta As String)
On Error Resume Next
    Dim qry As New rdoQuery
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = "INSERT INTO tipoCresta values('" & cresta & "')"
        qry.Execute
        qry.Close
End Sub

Public Function nombreCuerda(id As Integer) As String
Dim rs2 As New ADODB.Recordset
SQL = "SELECT Cuerda " & _
        "FROM Cuerdas WHERE idCuerda = " & id & ""
rs2.Open SQL, cnn, adOpenStatic, adLockOptimistic

If rs2.RecordCount >= 1 Then
    nombreCuerda = rs2("Cuerda")
End If
rs2.Close
End Function

Public Function peleasPendientes() As Integer
Dim rs2 As New ADODB.Recordset
SQL = "SELECT count(idPelea) As Numero " & _
        "FROM Peleas WHERE orden = 0"
rs2.Open SQL, cnn, adOpenStatic, adLockOptimistic
peleasPendientes = rs2("Numero")
rs2.Close
End Function

'Determina el id de una ciudad
Public Function dimeCiudad(pais As String, ciudad As String) As Integer
    Dim qry As New rdoQuery
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = "SELECT c.* FROM ciudad c INNER JOIN pais p ON c.idPais = p.id WHERE c.ciudad = '" & ciudad & "' and p.pais = '" & pais & "'"
    Set rst = qry.OpenResultset(rdOpenDynamic)
        If rst.RowCount > 0 Then
            dimeCiudad = rst("idCiudad")
        Else
            dimeCiudad = -1
        End If
        qry.Close
End Function
'Determina el id de una ciudad
Public Function dimePais(pais As String) As Integer
    Dim qry As New rdoQuery
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = "SELECT * FROM pais WHERE pais = '" & pais & "'"
    Set rst = qry.OpenResultset(rdOpenDynamic)
        If rst.RowCount > 0 Then
            dimePais = rst("id")
        Else
            dimePais = -1
        End If
        qry.Close
End Function
'Determina el id de una ciudad
Public Sub llenarPaisCiudad(ciudad As ComboBox, pais As ComboBox, ciu)
    Dim qry As New rdoQuery
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = "SELECT c.ciudad as ciudad,p.pais as pais FROM ciudad c INNER JOIN pais p ON c.idPais = p.id WHERE c.idCiudad = " & ciu & ""
    Set rst = qry.OpenResultset(rdOpenDynamic)
        If rst.RowCount > 0 Then
            ciudad.Text = rst("ciudad")
            pais.Text = rst("pais")
        End If
        qry.Close
End Sub
'Cargar colores
Public Sub cargarColores(lista As ComboBox)
    Dim qry As New rdoQuery
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = "Select * from Colores"
    Set rst = qry.OpenResultset(rdOpenDynamic)
        While rst.EOF = False
            lista.AddItem rst("Color")
            rst.MoveNext
        Wend
        qry.Close
End Sub

'Cargar crestas
Public Sub cargarCrestas(lista As ComboBox)
    Dim qry As New rdoQuery
    Set qry.ActiveConnection = RDOCONEXION
        qry.SQL = "Select * from tipoCresta"
    Set rst = qry.OpenResultset(rdOpenDynamic)
        While rst.EOF = False
            lista.AddItem rst("tipoCresta")
            rst.MoveNext
        Wend
        qry.Close
End Sub

Public Function peleasxOrdenar() As Integer
Dim rs2 As New ADODB.Recordset
SQL = "SELECT count(idPelea) As Numero " & _
        "FROM Peleas WHERE orden = 0"
rs2.Open SQL, cnn, adOpenStatic, adLockOptimistic
peleasxOrdenar = rs2("Numero")
rs2.Close
End Function

