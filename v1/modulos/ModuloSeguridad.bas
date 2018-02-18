Attribute VB_Name = "ModuloSeguridad"
Public Function decodificar(strCadena As String, strSemilla As String) As String
Dim decoded As String
Dim c As Double

c = CDbl(strCadena)
c = (c / 63) - 1990

decoded = c
decoded = Right(decoded, Len(decoded) - 1)
strCadena = ""
For i = 1 To Len(s) Step 3
    strCadena = strCadena & Chr(Val(Mid(decoded, i, 3)))
Next
End Function

Public Function codificar(strCadena As String, strSemilla As String) As String

Dim s As String
Dim coded As String

s = UCase(strCadena)
For i = 1 To Len(s)
    coded = coded & "0" & Asc(Mid(s, i, 1))
Next

coded = 8 & coded
coded = (CDbl(coded) + 1990) * 63

codificar = coded
End Function

