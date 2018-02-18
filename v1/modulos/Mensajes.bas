Attribute VB_Name = "Mensajes"
Public Sub menFaltanDatos()
MsgBox "Faltan datos requeridos para continuar el proceso", vbCritical, "Faltan datos"
End Sub

Public Sub menPersonaDuplicada()
MsgBox "Ya existe otra persona con el mismo documento", vbExclamation, "Imposible guardar"
End Sub

Public Sub menCuerdaDuplicada()
MsgBox "Ya existe otra cuerda con el mismo nombre", vbExclamation, "Imposible guardar"
End Sub

Public Sub menGuardadoExitoso()
MsgBox "El registro se guardo correctamente", vbInformation, "Guardado exitoso"
End Sub
Public Sub menGuardadoFallo()
MsgBox "El registro no se pudo guardar correctamente, intente más tarde", vbCritical, "Fallo guardado"
End Sub
