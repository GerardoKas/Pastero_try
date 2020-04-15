Attribute VB_Name = "modTitulacion"
Option Explicit
Public Function titular() As String
With Principal.TheText
    .SelStart = 0
    .SelLength = 255
    titular = getTituloTexto(.SelText)
    .SelLength = 0
End With
Principal.txtTitulo = titular
End Function

Private Function getTituloTexto(texto As String)
Dim tit As String
'comprobar puntuacion no posible
''Exit Function
''If principal.TheText.Tag = "Nodo" Then Exit Function
''If fileWasOpened = True Then
''    crearTitulo = basename(savedFilename)
''    Exit Function
tit = changeBadChars(texto)
tit = Strings.Left$(tit, 31)
tit = Trim(tit)
sFileName = tit & ""
getTituloTexto = sFileName
End Function

''sacar enteres al principio y al final. por si el texto viene despues

Function changeBadChars(texto As String)
texto = Replace(texto, ":", "-")
texto = Replace(texto, "?", "!")
texto = Replace(texto, """", "'")
texto = Replace(texto, "/", "-")
texto = Replace(texto, "\", "-")
texto = Replace(texto, "|", "-")
texto = Replace(texto, "<", "[")
texto = Replace(texto, ">", "]")
texto = Replace(texto, "*", "·")
texto = Replace(texto, vbNullChar, "-")
texto = Replace(texto, ".", "-")
changeBadChars = texto
End Function


Function EngadirExtension(file As String, ext As String)
    Dim r As String
    r = Strings.Right$(file, 3)
End Function
