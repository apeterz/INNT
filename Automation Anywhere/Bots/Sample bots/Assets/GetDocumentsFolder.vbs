'Ojo es una funcion!
Function GetDocumentsFolder()
Dim objShell
Set objShell = CreateObject("WScript.Shell")
GetDocumentsFolder = objShell.SpecialFolders("MyDocuments")
End Function