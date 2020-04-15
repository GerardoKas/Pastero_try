Attribute VB_Name = "modPastero"
Public Const maxW = 640
Public Const maxH = 480
Public Const maxArea = 1600
Public Const maxTitle = 128
Public Const minw = 320
Public Const minh = 240

Public availW As Long
Public availH As Long

Public esMaxi As Boolean
Public sFilePath As String
Public sExtension As String
Public sFileName As String
Public savedFilename As String
Public fileWasOpened As Boolean
Public esTemporal As Boolean

Public Const bar = vbCrLf & "(...)" & vbCrLf


Public cEstado As Estado
Public grabado As Boolean
Public Enum Estado
    Nuevo = 0
    Abierto = 1
    Guardado = 2
    Cambiado = 3
End Enum
    
Public Const MyApp = "Pastero_Drop_n_Save"
Public Const MySection = "Config"

'''''''''''''''''''''''''''
Public MyDataObject As RichTextLib.DataObject

Sub tellSaved(opt As Estado)
'no me cargaun dim del type

If Principal.TheText.Text <> "" Then grabado = True ': opt = 0
'Select Case opt
'Case Estado.Changed ' unsaved
'    'cambiando y sin grabar
'    grabado = False
'    Principal.Caption = "Paste:" & sFileName
'Case Estado.saved 'saved
'    'recien grabado
'    grabado = True
'    Principal.Caption = "Saved. " & sFileName
'    txtTitulo.Text = basename(sFileName)
'Case Estado.Nuevo 'new
'    'sin nada escrito, nuevo
'    grabado = True
'    Principal.Caption = "New"
'Case Estado.Abierto 'opened
'    'abierto un archivo y no cambiado
'    grabado = True
'    Principal.Caption = "Open:" & filename
'    'statitulo = basename(filename)
'End Select
End Sub

Function ponerArriba(frmHwnd As Long, Optional x = 200, Optional y = 0)
'w = Int(Screen.Width / Screen.TwipsPerPixelX)
ok = SetWindowPos(frmHwnd, -1, x, y, minw, minh, 0)
ponerArriba = ok
End Function

Public Sub posicionVentana()
Dim x As Long, y As Long
x = GetSetting(MyApp, MySection, "Left", 300)
y = GetSetting(MyApp, MySection, "Top", 0)
If x > Screen.Width Or x < 0 Then x = 0
If y > Screen.Height Or y < 0 Then y = 0
x = x / Screen.TwipsPerPixelX
y = y / Screen.TwipsPerPixelY
SetWindowPos Principal.hwnd, -1, x, y, minw, minh, 0

End Sub
Function sizeMini()
If Principal.WindowState = vbMaximized Then
 Principal.WindowState = vbNormal
End If
Principal.ScaleMode = vbPixels
'X = Principal.ScaleLeft
'Y = Principal.ScaleWidth
Principal.Width = minw * Screen.TwipsPerPixelX
Principal.Height = minh * Screen.TwipsPerPixelY
End Function

Function sizeMaxi()
'X = Principal.ScaleLeft
'Y = Principal.ScaleTop
On Error Resume Next
Principal.Width = maxW * Screen.TwipsPerPixelX
Principal.Height = maxH * Screen.TwipsPerPixelY
End Function


Public Function replaceVblf(texto) As String
Dim reg As RegExp
Set reg = New RegExp
reg.Global = True
reg.MultiLine = True
reg.Pattern = "\n"
If reg.Test(texto) Then

replaceVblf = reg.Replace(texto, "<BR>")
End If
'TheText.Text = replaceVblf
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' z = SetWindowPos(frmFormatoSave.hwnd, -1, 320, 240, 313 + 30, 54 + 30, 0)
''''''''''''''''''''''''''''''''''''''''''''''''''
