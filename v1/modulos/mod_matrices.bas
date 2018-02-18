Attribute VB_Name = "mod_matrices"
'Agrega un dato a un listview de una matriz de datos en una posicion especifica
Public Sub agregarDatoMatriz(lista As ListView, matriz() As String, nu As Integer)
Set li = lista.ListItems.Add(, , matriz(nu, 1))
    li.SubItems(1) = matriz(nu, 2)
    li.SubItems(2) = matriz(nu, 3)
    li.SubItems(3) = matriz(nu, 4)
    li.SubItems(4) = matriz(nu, 5)
End Sub

'Busca la poscion de un dato en un vector
Public Function bucarPosVector(vector, valor) As Integer
Dim f, id As Integer
Dim f2 As Integer
Dim idC As Integer
Dim peso As Integer

For f = 1 To UBound(vector)
    If vector(f) = valor Then
        id = f
        bucarPosVector = id
        Exit Function
    End If
Next
End Function

Public Function agregarFilasMatrizPreserve(matrizOrg, nFilas As Integer, nCols As Integer)
Dim matrizNew() As String
Dim filN, colN As Integer
filN = UBound(matrizOrg) + nFilas
colN = nCols

ReDim matrizNew(filN, colN)

'Copio los datos
Dim f, c As Integer
For f = 1 To filN - nFilas
    For c = 1 To colN
        matrizNew(f, c) = matrizOrg(f, c)
    Next
Next
agregarFilasMatrizPreserve = matrizNew
End Function

Public Function setFilasMatriz(matrizOrg, nFilas As Integer, nCols As Integer)
Dim matrizNew() As String
Dim filN, colN As Integer
filN = nFilas
colN = nCols

ReDim matrizNew(filN, colN)

'Copio los datos
Dim f, c As Integer
For f = 1 To filN
    For c = 1 To colN
        matrizNew(f, c) = matrizOrg(f, c)
    Next
Next
setFilasMatriz = matrizNew
End Function

