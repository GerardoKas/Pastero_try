Attribute VB_Name = "modOpenySaveFilename"
Option Explicit


Function getSaveFilename() As String
'dialog.Flags = cdlOFNNoValidate + cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist
Dim dialog As CommonDialog
Set dialog = Principal.cmdlg
setDialogProperties Principal.cmdlg
sExtension = getExtension(sFileName)
If sExtension = "" Then
    dialog.DefaultExt = "txt"
End If
If sFilePath <> "" Then
    dialog.InitDir = sFilePath
End If
If sFileName <> "" Then
    dialog.filename = dialog.InitDir & sFileName
    MsgBox dialog.filename
End If
On Error GoTo ErrName
dialog.ShowSave
getSaveFilename = dialog.filename

sFileName = dialog.filename

Exit Function
ErrName:
    getSaveFilename = ""
    MsgBox " No lo Grabaste. JO!" & vbCrLf & vbCrLf & "Problema " & Err.Description & vbCrLf & Err.Number & vbCrLf & Err.Source
End Function

Function getOpenFilename() As String
Dim file As String
Dim dialog As CommonDialog
Set dialog = Principal.cmdlg
setDialogProperties Principal.cmdlg

On Error GoTo ErrName
dialog.ShowOpen

file = dialog.filename
sExtension = modFuncFileIO.getExtension(file)
sFilePath = modFuncFileIO.getBasePath(file)
sFileName = file
getOpenFilename = file
Exit Function
ErrName:
If Err.Number = 32755 Then
    MsgBox "No has Abierto Nada", vbOKOnly, "On The Open"
    getOpenFilename = ""
End If
End Function


Sub setDialogProperties(dlg As CommonDialog)
dlg.Flags = cdlOFNNoValidate
dlg.CancelError = False 'genera error si se puelsa cancelar
dlg.Filter = "Texto plano *.txt|*.txt|WordPad *.rtf|*.rtf|Pagina Web *.html|*.html|Todos los Archivos *.*|*.*"
dlg.FilterIndex = 4
End Sub
