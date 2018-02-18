Attribute VB_Name = "Archivos"
Public controlar As New FileSystemObject

'CREA CARPETAS SINO EXISTEN
Public Function CreaCarpetas(Carpeta)
    If controlar.FolderExists(Carpeta) = False Then
        controlar.CreateFolder (Carpeta)
    End If
End Function

'PREGUNTA SI EXISTE EL DIRECTORIO
Public Function DirectorioExiste(Ruta) As Boolean
    If controlar.FileExists(Ruta) = False Then
        DirectorioExiste = False
    Else
        DirectorioExiste = True
    End If
End Function
