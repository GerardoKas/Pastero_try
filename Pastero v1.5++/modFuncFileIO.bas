Attribute VB_Name = "modFuncFileIO"
Option Explicit

Function getTempName(extension As String)
If Strings.Left$(extension, 1) = "." Then extension = Strings.Right$(extension, Len(extension) - 1)
getTempName = "c:\TEMP\[" & DateTime.Date$ & "]." & extension
'aca se define el nombre temporal
End Function

Function grabarFichero(fname As String)
Dim f As Integer
f = FreeFile
If fname = "" Then Exit Function
sExtension = getExtension(fname)

Open fname For Binary As #f
If sExtension = "rtf" Then
Put #f, 1, Principal.TheText.TextRTF
Else
Put #f, 1, Principal.TheText.Text
End If
Close #f
grabado = True
End Function

Public Sub AbrirYLeerAlInicio(fichero As String)
Principal.Enabled = False
Principal.MousePointer = vbHourglass
If FicheroExiste(fichero) Then

    Principal.TheText.Text = modFuncFileIO.openFile(fichero)
    tellSaved 3
Else
    MsgBox fichero & vbCrLf & "No es un fichero"
    tellSaved 2
End If
Principal.Enabled = True
Principal.MousePointer = vbDefault
Principal.TheText.SelStart = 0
'principal.SetFocus
End Sub

Function FicheroExiste(filespec As String) As Boolean
'    If FileSystem.Dir(filespec, vbNormal + vbArchive) = "" Then
'        FicheroExiste = True
'
'    End If
'    Exit Function
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If (fso.FileExists(filespec)) Then
    FicheroExiste = True: Exit Function
  Else
    FicheroExiste = False: Exit Function
  End If
End Function

Function funcAntesdeBorrar()
'Dim resp As VbMsgBoxResult
'Dim nf As String  ' ''  algun tipo de fileanem
'
'''no se que es antesdeborrar . sera al cerrar
'If grabado = True Then funcAntesdeBorrar = True: Exit Function
'
'resp = MsgBox("No se han guardado Los cambios, deseas Guardarlo?", vbYesNoCancel + vbApplicationModal + vbQuestion + vbMsgBoxSetForeground, "Grabar?")
'If resp = vbYes Then
'    nf = modOpenySaveFilename.getSaveFilename()
'    If nf = "" Then
'        funcAntesdeBorrar = True
'        Exit Function
'    End If
'    grabarFichero (nf)
'    funcAntesdeBorrar = True
'ElseIf resp = vbCancel Then
'    funcAntesdeBorrar = False
'Else
'    funcAntesdeBorrar = True
'End If
End Function

Public Function openFile(fichero As String) As String
Dim contenido As String
Dim tamano As Long
Dim f As Integer
Dim pos2 As Integer
f = FreeFile
'savedFilename = fichero
'fileWasOpened = True
fichero = ponerComillasOk(fichero)
sFilePath = getBasePath(fichero)
sExtension = getExtension(fichero)
sFileName = basename(fichero)

 
Debug.Print Err.Description & Err.Number
If Err.Number = 75 Then MsgBox "No es un arhivo": Exit Function
Principal.TheText.LoadFile (sFileName)
MsgBox "abierto2"
'Open sFileName For Binary As #f
'tamano = FileLen(fichero)
'contenido = Input(tamano, #f)
'Close #f
'contenido = Replace(contenido, Chr(0), Chr(12))
'contenido = replaceVblf(contenido)
openFile = contenido
End Function

Function ponerComillasOk(fname) As String
    If Strings.Left$(fname, 1) <> """" Then
        fname = """" & fname & """"
    End If
    ponerComillasOk = fname
End Function

Public Function getBasePath(file)
Dim pos As Integer
    pos = InStrRev(file, "\")
    getBasePath = Left(file, pos)
End Function

Public Function basename(file)
    Dim pos As Integer
    pos = InStrRev(file, "\")
    basename = Right(file, Len(file) - pos)
End Function

Public Function getExtension(fname As String) As String
    Dim et As String, n As Long
    
    If fname = "" Then getExtension = "": Exit Function
     n = InStrRev(fname, ".")
    If n > 0 Then
        et = Right(fname, Len(fname) - n)
    Else
        et = ""
    End If
    getExtension = et
    Debug.Print getExtension
End Function

