VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Principal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pastero (for Paste)"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4200
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   240
   ScaleMode       =   0  'User
   ScaleWidth      =   320
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2010
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPaste 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "Paste"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   480
         MaskColor       =   &H00004080&
         TabIndex        =   5
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdMinMax 
         BackColor       =   &H80000007&
         Caption         =   "Maxi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.TextBox txtTitulo 
      Height          =   240
      Left            =   2040
      TabIndex        =   2
      Text            =   "Nombre Del Archivo"
      Top             =   0
      Width           =   1920
   End
   Begin RichTextLib.RichTextBox TheText 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   1
      TextRTF         =   $"frmPrincipal.frx":038A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   3600
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox richAlternativo 
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      _Version        =   393217
      TextRTF         =   $"frmPrincipal.frx":040C
   End
   Begin VB.Menu mnuCabeceraMenu 
      Caption         =   "Enlazar"
      Visible         =   0   'False
      Begin VB.Menu mnuPegarRTF 
         Caption         =   "Pegar RTF"
      End
      Begin VB.Menu mnuPegarTXT 
         Caption         =   "Pegar TXT"
      End
      Begin VB.Menu mnuContenidoPegar 
         Caption         =   "Pegar/Drop Texto en Porciones (...)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContenidoPegarImagen 
         Caption         =   "Pegar Imagen/es"
      End
      Begin VB.Menu mnuZz 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilelist 
         Caption         =   "Listado de Ficheros"
      End
      Begin VB.Menu mnuLinkWeb 
         Caption         =   "Listado de ficheros en HTML"
      End
      Begin VB.Menu mnuZ2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContenidoSinMas 
         Caption         =   "Contenidos Sin mas"
      End
      Begin VB.Menu mnuContenidoHTML 
         Caption         =   "Contenidos (HTML-code)"
      End
      Begin VB.Menu mnuContenidoPseudocode 
         Caption         =   "Contenidos Pseudocode (TXT)"
      End
      Begin VB.Menu mnuZ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImgsHTML 
         Caption         =   "Listado de Imagenes HTML"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuNuevo 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "Abrir"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu mnuSaveOver 
         Caption         =   "Guardar Encima"
         Enabled         =   0   'False
         Shortcut        =   ^G
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDirs 
         Caption         =   "Guardar En..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDir 
         Caption         =   "Cambiar Directorio"
      End
      Begin VB.Menu mnuAbrirNotepad 
         Caption         =   "Abrir en Notepad"
      End
      Begin VB.Menu mnuAbrirNavegador 
         Caption         =   "Abrir en el Navegador"
      End
      Begin VB.Menu mnuAbrirCarpetaContenedora 
         Caption         =   "AbrirCarpetaContenedora"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Imprimir (¿PDF?)"
      End
      Begin VB.Menu mnNone 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesordenar 
         Caption         =   "Desordenado"
      End
      Begin VB.Menu mnuSortLines 
         Caption         =   "Ordenar Lineas"
      End
      Begin VB.Menu mnuDelDuplicate 
         Caption         =   "Eliminar Duplicados Juntos"
      End
      Begin VB.Menu mnuEngadirCode 
         Caption         =   "AñadirCodigo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCodeAtBegin 
         Caption         =   "Poner Codigo Al Inicio"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuReplaceVBcrlf 
         Caption         =   "Poner <BR> en \n"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "Añadir Formato"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuVoid 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindReplace 
         Caption         =   "Buscar o Reemplazar"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuInsetarfecha 
         Caption         =   "Insertar Fecha y Hora"
      End
      Begin VB.Menu mnuColorYFuente 
         Caption         =   "Color y Fuente"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'INICIO

'un simple paste del sistemaa
Private Sub cmdPaste_Click()
SendMessage Principal.TheText.hwnd, WM_PASTE, 0, 0
Exit Sub
End Sub

Private Sub cmdMinMax_Click()
If (cmdMinMax.Caption = "Mini") Then
    sizeMini
    cmdMinMax.Caption = "Maxi"
Else
    sizeMaxi
    cmdMinMax.Caption = "Mini"
End If
End Sub

Private Sub cmdAbout_Click()
MsgBox "Creado Por Gerardo Castro Mtz. " & vbCrLf & "Desde 2006, en Vigo (hasta el dia de hoy Abril_2020" & _
vbCrLf & "Ubicado en : " & vbCrLf & App.Path, vbOKOnly, "* * *" & App.CompanyName & "* * *"
End Sub

Private Sub Form_Load()
Dim linea As String
Dim f As String
f = Interaction.Command$()

sFilePath = GetSetting(MyApp, MySection, "LastDir", "")
If sFilePath = "" Then
sFilePath = App.Path
End If

On Error Resume Next
If FicheroExiste(f) = True Then
    Principal.TheText.LoadFile f
    'tellSaved (1)
    sFilePath = modFuncFileIO.getBasePath(f)
    posicionVentana
End If

'txtTitulo =
'= "Arrastra y suelta texto o rtf para copiar ... Dropea un conjunto de archivos encima del forumario"
'Len(TheText.Text) \ 1024 & "Kbs" & "..." & sFileName & "..." ' & FileDateTime(filename)
'TheText.SetFocus
'TheText.Span ' para buscar fin de cadenas

End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting MyApp, MySection, "Left", Principal.Left
SaveSetting MyApp, MySection, "Top", Principal.Top
SaveSetting MyApp, MySection, "LastDir", sFilePath
End Sub

''Private Sub Form_Unload(Cancel As Integer)
''Dim comamdoIcono
''Dim comandoExec
''Dim shortpath As String
''Dim longitud As String
''Dim bufferlen As String
''Dim longpath As String
''
''If funcAntesdeBorrar = False Then Cancel = 1: Exit Sub ' cancel=1 quiere decir continuar. si e0 es cerrar
''
''
''SaveSetting MyApp, MySection, "Left", principal.Left
''SaveSetting MyApp, MySection, "Top", principal.Top
''
''SaveSetting MyApp, MySection, "LastDir", sFilePath
''
'''If valor = 0 Then MsgBox "No he podido capturar la ruta corta de " & App.EXEName
''
''shortpath = Strings.String$(255, vbNullChar)
''bufferlen = Strings.Len(shortpath)
''longpath = App.Path & "\" & vbNullChar
''longitud = GetShortPathName(longpath, shortpath, bufferlen)
''shortpath = Left(shortpath, (longitud - 1))
''Dim icono As String, exec As String
''icono = shortpath & ", 1"
''exec = shortpath & " %%*"
''
''
''''''''regicono = "cmd /C reg add HKCR\.nfo\DefaultIcon /ve /d """ & icono & """ /f" 'icono
'''deletear -- -- -- comandoprevio = "cmd /C reg add HKCR\.nfo\Shell\Open\Command"
''''''''''regexec = "cmd /C reg add HKCR\.nfo\Shell\Open\Command /ve  /d """ & exec & """ /f" 'programa
''
'''comandoicono = "cmd /C reg add HKCR\.txt\DefaultIcon /ve /d """ & icono & """ /f" 'icono
'''comandoprevio = "cmd /C reg add HKCR\.txt\Shell\Open\Command"
'''comandoExec = "cmd /C reg add HKCR\.txt\Shell\Open\Command /ve  /d """ & exec & """ /f" 'programa
'''''''''''''''''Debug.Print regicono & vbCrLf & regexec
'''Shell regicono
'''Shell regexec
'''MsgBox "APP: " & vbCrLf & App.Path
'''''''''''Debug.Print (icono & exec)
'''''''''''Debug.Print vbCrLf
''Debug.Print comandoExec
'''MsgBox comandoicono & vbCrLf & comandoExec
''
''End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

If Data.GetFormat(vbCFFiles) = True Then
    Dim DataTexto As String
    Dim i As Integer
    Dim hora As String, dia As String
    Dim r As RichTextBox
    
    
    'TheText = ""
        For i = 1 To Data.Files.count
            
           sFilename = Data.Files(i)
            hora = Strings.Format$(FileDateTime(sFilename), "Long Time")
            'fecha = Format(fecha, "dddd d, mmm de yyyy a las hh:mm:ss")
            dia = Strings.Format$(FileDateTime(sFilename), "Long Date")
            Set r = TheText
            
            r.LoadFile sFilename
            Principal.TheText.SelText = "((" & basename(sFilename) & "))" & vbCrLf & "<:" & dia & " A las " & hora & ":>" & vbCrLf
            Principal.TheText.SelText = r.Text
            Principal.TheText.SelText = vbCrLf & "<<.FIN.>>" & vbCrLf & vbCrLf
            'TheText.Text = principal.thetext.Text & vbCrLf & "(...)" & vbCrLf & "Titulo: " &sFileName& " " & vbCrLf & vbCrLf & principal.thetext.LoadFile(filename)
        Next
   Set r = Nothing
   
   
Else
    Principal.SetFocus
    MsgBox "Nothing Happens"
End If
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
TheText.SetFocus
If Principal.WindowState = vbMinimized Then Exit Sub
'Debug.Print "Window State " & principal.WindowState
With TheText
.Width = Principal.ScaleWidth
.Height = Principal.ScaleHeight - .Top
End With
availW = -Me.ScaleWidth - Principal.ScaleLeft + Screen.Width
availH = Screen.Height - Principal.ScaleWidth - Principal.ScaleHeight
If Principal.ScaleWidth >= 30 Then
    cmdMinMax.Caption = "Mini"
ElseIf Principal.ScaleWidth < 320 Then
    cmdMinMax.Caption = "Maxi"
End If
If Principal.ScaleHeight >= 240 Then
    cmdMinMax.Caption = "Mini"
ElseIf Principal.ScaleHeight < 240 Then
    cmdMinMax.Caption = "Maxi"
End If
txtTitulo.Width = Principal.ScaleWidth - txtTitulo.Left
End Sub


Private Sub mnuAbrir_Click()
Dim file As String
'cmdlg.InitDir = sFilePath
file = getOpenFilename()
If FicheroExiste(file) Then
     TheText.LoadFile file
    'txtTitulo = Len(TheText) \ 1024 & "Kb's"
    txtTitulo = modTitulacion.titular
Else
MsgBox "NO has abierto archivo", vbOKOnly
End If


End Sub

Private Sub mnuAbrirCarpetaContenedora_Click()
'abrar lacarpa
Shell "explorer """ & sFilePath & """", vbNormalFocus

End Sub

Private Sub mnuAbrirNavegador_Click()
Dim t As String

t = getTempName("html")
grabarFichero (t)
ShellExecute Principal.hwnd, "open", t, 0, 0, SW_SHOWNORMAL
End Sub

Private Sub mnuAbrirNotepad_Click()
Dim t As String
t = getTempName("txt")

grabarFichero (t)
Shell "Notepad.exe """ & t & """", vbNormalFocus
End Sub

Private Sub mnuCodeAtBegin_Click()
'CODIGO DE COMENTARIO A INICI
End Sub

Private Sub mnuColorYFuente_Click()
frmOptions.Show 0, Me
End Sub

Private Sub mnuContenidoPegar_Click()
Dim nt As String
If Clipboard.GetFormat(vbCFText) Then
    Principal.TheText.SelText = (Clipboard.GetText(vbCFText)) & bar
ElseIf Clipboard.GetFormat(vbCFRTF) Then
    Principal.TheText.SelRTF = (Clipboard.GetText(vbCFRTF)) & bar
ElseIf Clipboard.GetFormat(vbCFDIB) Then
    'nt = Clipboard.GetData(vbCFDIB)
    SendMessage Principal.TheText.hwnd, WM_PASTE, 0, 0
Else
    MsgBox "Clipboard Files " & vbCrLf & "No puedo capturar" & vbCrLf & "Prueba el comando paste con el boton derecho encima del programa"
End If
'TheText.Text = principal.thetext.Text & Trim(nt) & bar
'TheText.SelStart = Len(TheText.Text)



End Sub

Private Sub mnuContenidoHTML_Click()
If MyDataObject.Files.count <> 0 Then
    Dim DataTexto As String
    Dim i As Integer
    Dim dia As String, hora As String
        For i = 1 To MyDataObject.Files.count
           sFilename = MyDataObject.Files(i)
            If FileLen(sFilename) <> 0 Then
            richAlternativo.LoadFile sFilename, 1
            
            'Replace richAlternative.Text, vbCrLf, vbCrLf & "<br>", 1, -1
            hora = Format(FileDateTime(sFilename), "Long Time")
            dia = Format(FileDateTime(sFilename), "Long Date")
            Principal.TheText.SelText = "<a href=""" & sFilename & """>" & basename(sFilename) & "</a><br>" & vbCrLf & _
            "<pre>" & dia & " A las " & hora & "</pre>" & vbCrLf
            Principal.TheText.SelText = richAlternativo.Text & vbCrLf & "<HR>" & vbCrLf
            
            End If
        Next

Else
    Principal.SetFocus
    MsgBox "Nothing Happens"
End If
End Sub

Private Sub mnuContenidoPegarImagen_Click()
Dim Item As Variant
On Error GoTo MsgErrorNoPic

For Each Item In MyDataObject.Files
Clipboard.SetData LoadPicture(Item), vbCFDIB
SendMessage Principal.TheText.hwnd, WM_PASTE, 0, 0
Next
sExtension = "rtf"
Exit Sub
MsgErrorNoPic:
MsgBox "ESTO NO ERARN iMAGENES" & vbCrLf & Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source


End Sub

Private Sub mnuContenidoPseudoCode_Click()
'Sltar un monte de archivos txt y pnerlos en formato de html todo elcontenido
Dim hora As String, dia As String
If MyDataObject.GetFormat(vbCFFiles) = False Then Principal.SetFocus: Exit Sub
     Dim i As Integer
    
        For i = 1 To MyDataObject.Files.count
        
           sFilename = MyDataObject.Files(i)
            If FileLen(sFilename) <> 0 Then
            richAlternativo.LoadFile sFilename
            
            hora = Format(FileDateTime(sFilename), "Long Time")
            dia = Format(FileDateTime(sFilename), "Long Date")
            
            Principal.TheText.SelText = "((->" & basename(sFilename) & "<-))" & vbCrLf & "<!: " & dia & " A las " & hora & " :!>" _
            & vbCrLf & richAlternativo.Text & vbCrLf _
            & "<<.FIN.>>" & vbCrLf & vbCrLf
            End If
            
            'TheText.Text = principal.thetext.Text & vbCrLf & "(...)" & vbCrLf & "Titulo: " &sFileName& " " & vbCrLf & vbCrLf & principal.thetext.LoadFile(filename)
        Next
    Principal.SetFocus
End Sub

Private Sub mnuContenidoSinMas_Click()
Dim hora As String, dia As String
If MyDataObject.GetFormat(vbCFFiles) = True Then
    Dim DataTexto As String
    Dim i As Integer
    Dim r As Object
        For i = 1 To MyDataObject.Files.count
        '''''''''''''''''''''''''''''''''''''''
        'crear un rtf virtual donde cargar ls archivos
        '''''''''''''''''''''''''''''''''''''''''
            sFilename = MyDataObject.Files(i)
            If FileLen(sFilename) <> 0 Then
            richAlternativo.LoadFile sFilename, 1
            Principal.TheText.SelText = richAlternativo.Text
            Principal.TheText.SelText = bar
            End If
            Next
Else
    MsgBox "Tienes que soltar uno o varios archivos"
    
End If
End Sub

'THIISOK
Private Sub mnuDelDuplicate_Click()
Dim aLines() As String, linedel() As Boolean, i As Integer
If Left(TheText, 1) = "" Then Exit Sub
Me.MousePointer = vbHourglass

aLines = Split(TheText.Text, vbCrLf)
ReDim linedel(UBound(aLines))
For i = 0 To UBound(aLines) - 1
    If StrComp(aLines(i), aLines(i + 1), vbTextCompare) = 0 Then
        linedel(i + 1) = True
    End If
Next
Dim t As String
For i = 0 To UBound(aLines)
    'si no hay quer eliminarla
    If Not linedel(i) = True Then
        t = t & aLines(i) & vbCrLf
    End If
Next
TheText.Text = t
Me.MousePointer = vbArrow
End Sub

Private Sub mnuDesordenar_Click()
Dim t As String
Dim i As Integer
Dim ps, num, temp
funcAntesdeBorrar
t = Principal.TheText.Text
t = Replace(t, vbCrLf, " ")
ps = Split(t, " ")

For i = 0 To UBound(ps)
    num = Int(Rnd(1) * UBound(ps))
    temp = ps(i)
    ps(i) = ps(num)
    ps(num) = temp
Next
TheText.Text = Join(ps, " ")
End Sub

Private Sub mnuDir_Click()
BrowseForFolder (sFilePath)
End Sub


Private Sub mnuFilelist_Click()
Dim i As Integer
Dim TxtEnlaces As String
On Error GoTo FINITO
'CROREREGIR FOR =1
    For i = 1 To MyDataObject.Files.count
    TxtEnlaces = TxtEnlaces & vbCrLf & Chr(34) & MyDataObject.Files(i) & Chr(34)
    
    Next
    Principal.TheText.SelText = TxtEnlaces
FINITO:
MsgBox Err.Description

End Sub

Private Sub mnuFindReplace_Click()
frmSearch.Show
End Sub

Private Sub mnuGuardar_Click()
Dim f As String

f = modOpenySaveFilename.getSaveFilename()

If f <> "" Then
    grabarFichero f
Else
    MsgBox "No hubo nombre de archivo. No se ha guardado" & vbCrLf & Err.Description, vbOKOnly
End If

End Sub

Private Sub mnuImgsHTML_Click()
Dim TxtEnlaces As String
Dim i As Integer
'MsgBox "MEBU CLIKED IMGS"
On Error GoTo FINITO
    For i = 1 To MyDataObject.Files.count
    TxtEnlaces = TxtEnlaces & vbCrLf & "<a href=""" & MyDataObject.Files(i) & """><IMG SRC=""" _
    & MyDataObject.Files(i) & """ WIDTH=320></a><br> " & MyDataObject.Files(i) & "<hr>" & vbCrLf
    Next
    Principal.TheText.SelText = TxtEnlaces
FINITO:
    
End Sub

Private Sub mnuInsetarfecha_Click()
'If filename = "" Then
'filename = "C:\temp\TEMP-.TXT"
'grabarFichero (filename)
TheText.SelRTF = Format(Date, "Long Date") & " a las "
TheText.SelRTF = Format(Time, "Long Time")

            
'TheText.SelRTF = Format(Date$, "dddd d mmm yyyy", vbSunday) & vbCrLf
'TheText.SelRTF = Format(Time$, "hh:mm:ss") & vbCrLf

'TheText.SelRTF = "(" & Date$ & ")"
End Sub

Private Sub mnuLinkWeb_Click()
''PEGAR TxtEnlaces WEB DE DATOBJECTFILES
Dim TxtEnlaces As String
Dim hora As String, dia As String
Dim filename As String
Dim i As Integer
Dim r As RichTextBox

On Error GoTo FINITO

    For i = 1 To MyDataObject.Files.count
            filename = MyDataObject.Files(i)
            'richAlternativo.LoadFile filename
            hora = Format(FileDateTime(filename), "Long Time")
            dia = Format(FileDateTime(filename), "Long Date")
    TxtEnlaces = "<a href=""" & filename & """>" & filename & "</a><br>" & vbCrLf & _
        "<i>" & "Fecha : " & dia & ", a las " & hora & "</i>" & vbCrLf & "<br>Tamaño: " & FileLen(filename) & " Kb." & _
        vbCrLf & "<hr>" & vbCrLf
    Principal.TheText.SelText = TxtEnlaces
        
    Next
    
Exit Sub
FINITO:
MsgBox " ERROR" & vbCrLf & "ERNUMBR:" & Err.Number & vbCrLf & Err.Description
End Sub

'Private Sub mnuNotepad_Click()
'Dim tempName As String
'If grabado = True And filename <> "" Then
'tempName = filename
'Else
'tempName = "c:\TEMP\Temp_" & DateTime.Date$ & ".txt"
'End If
'grabarFichero tempName
'Shell "Notepad.exe " & tempName, vbNormalFocus
'End Sub

Private Sub mnuNuevo_Click()
'''''''''''''''''
funcAntesdeBorrar
'''''''''''''''''
sFilename = ""

TheText.Text = ""
'''''''''''
tellSaved 2
'''''''''''
End Sub

Private Sub mnuPegarRTF_Click()

Dim d As String

If MyDataObject.GetFormat(vbCFRTF) = True Then
   d = MyDataObject.GetData(vbCFRTF)
   'añade al final del texto, o donde este el puntero
   Principal.TheText.SelRTF = d
    
    Debug.Print (d)
    Principal.TheText.Refresh
ElseIf MyDataObject.GetFormat(vbCFText) = True Then
    d = MyDataObject.GetData(vbCFText)
    Principal.TheText.SelText = d
    
    Else
        MsgBox "Eso no era RTF. Ni text Prueba con : PegarTXT"
    End If
End Sub

Private Sub mnuPegarTXT_Click()
Dim t As String
Dim d
If MyDataObject.GetFormat(vbCFText) = True Then
   d = MyDataObject.GetData(vbCFText)
   Principal.TheText.SelText = d & bar
    
Else
    MsgBox ("No funciona pegr TxT")
End If
End Sub

Private Sub mnuPrint_Click()
On Error GoTo 0
On Error GoTo NoPrint
cmdlg.Flags = cdlPDPrintSetup
cmdlg.ShowPrinter

'MsgBox Printer.DriverName
'TheText.SelPrint hDC
Debug.Print (vbCrLf & Printer.DeviceName & vbCrLf & Printer.ColorMode _
& vbCrLf & Printer.Duplex)
TheText.SelPrint Printer.hDC
'MsgBox Err.Number
Exit Sub
NoPrint:
'If Err.Number Then MsgBox ("No se imprimira nada") : MsgBox("ERROR " & Err.Description)

End Sub

Private Sub mnuReplaceVBcrlf_Click()
TheText = replaceVblf(TheText.Text)
End Sub

'Private Sub mnuSaveAs_Click()
''save
'filename = getSaveFilename(cmdlg)
'If filename = "" Then MsgBox "No Se Puede Grabar sin nombre": Exit Sub
'
'grabarFichero filename
'txtTitulo = filename
'sFilePath = basepath(filename)
'
'tellSaved 1
'End Sub

Private Sub mnuSaveOver_Click()
If sFilePath = "" Then
'mnuDir_Click
End If

grabarFichero sFilePath & "\" & txtTitulo
tellSaved 1
End Sub

Private Sub mnuSortLines_Click()
Dim aLines() As String, temp As String
Dim i As Integer, j As Integer
aLines = Split(TheText.Text, vbCrLf)
For i = 0 To UBound(aLines)
    For j = i + 1 To UBound(aLines)
        'si i es mayor que j
        If StrComp(aLines(i), aLines(j), vbTextCompare) > 0 Then
            temp = aLines(j)
            aLines(j) = aLines(i)
            aLines(i) = temp
        End If
    Next
Next
'cambiaporcompletoel_texto_por eso no usar selTExT
TheText.Text = Join(aLines, vbCrLf)
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''AQUI''''''''''''''''''''''''
Private Sub TheText_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
txtTitulo = Len(TheText) / 1024 & " Kbs"

Set MyDataObject = Data
'PREAPARAR PARA COTINUAR
If MyDataObject.GetFormat(vbCFFiles) = True Then
    mnuContenidoPegarImagen.Enabled = True
    mnuContenidoPegar = False
    mnuContenidoSinMas = True
    mnuContenidoHTML.Enabled = True
    mnuContenidoPseudocode.Enabled = True
    mnuFilelist.Enabled = True
    mnuImgsHTML.Enabled = True
    mnuLinkWeb.Enabled = True
ElseIf MyDataObject.GetFormat(vbCFDIB) = True Then
    mnuContenidoPegarImagen.Enabled = True
    mnuContenidoSinMas = False
    mnuContenidoHTML.Enabled = False
    mnuContenidoPseudocode.Enabled = False
    mnuFilelist.Enabled = False
    mnuImgsHTML.Enabled = False
    mnuLinkWeb.Enabled = False
Else
    mnuPegarRTF.Enabled = True
    mnuPegarTXT.Enabled = True
    mnuContenidoSinMas = False
    mnuContenidoPegarImagen.Enabled = True
    'mnuContenidoPegar = True
    mnuContenidoHTML.Enabled = False
    mnuContenidoPseudocode.Enabled = False
    mnuFilelist.Enabled = False
    mnuImgsHTML.Enabled = False
    mnuLinkWeb.Enabled = False
    
        
    
End If


PopupMenu mnuCabeceraMenu

End Sub
''''''''''''''''SIN___________USAR''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Private Sub TimerStatus_Timer()
'Dim stat As String
'"1234567890---ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'On Error Resume Next
'txtTitulo = Right(txtTitulo, Len(txtTitulo) - 1) & Left(txtTitulo, 1) & " - "

'End Sub

'
'Private Sub TheText_Validate(Cancel As Boolean)
'TheText.SetFocus
'
'End Sub


