Attribute VB_Name = "Objetos"
Public Sub seleccionarTexto(texto As TextBox)
texto.SetFocus
texto.SelStart = 0
texto.SelLength = Len(texto)
End Sub

Public Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
'Permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If InStr("0123456789,", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function

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

Public Function IsFormLoaded(FormToCheck As Form) As Integer
Dim y As Integer

For y = 0 To Forms.Count - 1
If Forms(y) Is FormToCheck Then
    IsFormLoaded = True
    Exit Function
    End If
Next
IsFormLoaded = False
End Function

Public Sub baseFlexGrid(flex As MSHFlexGrid, encabezado As String)
With flex
    .AllowUserResizing = flexResizeColumns      ' -- Permitir redimensionar por columnas
    .FixedCols = 0                              ' -- Tipo de filas fijas - NO
    .FixedRows = 1                              ' -- Tipo de columnas fijas - SI
   ' .ForeColorFixed = vbHighlight               ' -- Color de los encabezados
    '.BackColorFixed = vbWhite                   ' -- BackGround de los encabezados
    '.GridLinesFixed = flexGridDots              ' -- Estilo de la linea del Grid de los headers
    .RowHeight(0) = 450                         ' -- Alto de la fila de encabezado
    .GridColor = RGB(190, 190, 190)             ' -- Color de las líneas del Grid
    .SelectionMode = flexSelectionByRow         ' -- Seleccionar fila completa
    .FormatString = encabezado
    
    Dim i As Integer
    Dim ancho As Double
    ancho = (.Width / (.Cols - 1)) - 50
    .ColWidth(0) = 0
    For i = 1 To .Cols - 1
        .ColWidth(i) = ancho
    Next
    
    .Row = 0
    For i = 0 To .Cols - 1
        .col = i
        .CellFontBold = True
    Next
    
    .Row = 1: .col = 0: .ColSel = .Cols - 1
    .Refresh
End With
End Sub

Public Sub quitarFilaBlanca(flex As MSHFlexGrid)
With flex
    If .Rows = 3 And .TextMatrix(1, 0) = "" Then
        .RemoveItem 1
    End If
End With
End Sub
Public Sub busquedaFlex(flex As MSHFlexGrid, encabezado As String)
With flex
    .FormatString = encabezado
    
    Dim i As Integer
    Dim ancho As Double
    ancho = (.Width / (.Cols - 1)) - 100
    .ColWidth(0) = 0
    For i = 1 To .Cols - 1
        .ColWidth(i) = ancho
    Next
    
    .Row = 0
    For i = 0 To .Cols - 1
        .col = i
        .CellFontBold = True
    Next
    
    If .Rows = 1 Then
       .RowSel = 0
       .ColSel = 0
    Else
        .RowSel = 1: .ColSel = .Cols - 1
    End If
    .Refresh
End With
End Sub

Public Sub soloUnaSeleccion(flex As MSHFlexGrid)
With flex
    .col = 0
    .RowSel = .Row
    .ColSel = .Cols - 1
End With
End Sub

Public Sub columnaMoneda(flex As MSHFlexGrid, col As Integer)
Dim fil As Integer
With flex
    .col = col
    .ColAlignment = 7
    For fil = 1 To .Rows - 1
        .TextMatrix(fil, col) = Format(.TextMatrix(fil, col), "$ #,##0.0")
    Next
End With
End Sub

Public Sub repintarFlex(flex As MSHFlexGrid, SQL As String, encabezado As String)
rs.Open SQL, cnn, adOpenStatic, adLockOptimistic

Set flex.DataSource = rs
    busquedaFlex flex, encabezado
    rs.Close
End Sub

Public Sub baseFlex(flex As MSHFlexGrid, SQL As String, encabezado As String)
rs.Open SQL, cnn, adOpenStatic, adLockOptimistic

Set flex.DataSource = rs
    baseFlexGrid flex, encabezado
    rs.Close
End Sub

'Funcion que convierte un valor numerico a formato moneda
Public Function convertirAMoneda(Valor As Double) As String
Dim resultado As String
resultado = Format(Valor, "$ #,##0.00")
convertirAMoneda = resultado
End Function

'Funcion que convierte un formato moneda a un valor numerico
Public Function convertirAValor(Valor As String) As Double
Dim resultado As String
resultado = Format(Valor, "#,##0.00")
convertirAValor = CDbl(resultado)
End Function

Public Sub hablar(mensaje As String)

End Sub

Public Function random(limite As Integer) As Integer
random = Int(limite * Rnd + 1)
End Function
