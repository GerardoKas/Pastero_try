Attribute VB_Name = "modApis"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PASTE = &H302


Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
(ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'Call GetShortPathName(Path, shortpath, 255)

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Const SW_SHOWNORMAL = 1

Public Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0





'Set objFld = obj.BrowseForFolder(hwnd, "Titulo Ventana", Bif_Flags, CarpetaOrigen)

Public Function BrowseForFolder(folder As String)
Dim obj, objFld, objPath
Dim objOrigen
'Flags of Style for the Popup
Const BIF_NEWDIALOGSTYLE = &H40
Const BIF_RETURNONLYFSDIRS = &H1     ' For finding a folder to start document searching
Const BIF_DONTGOBELOWDOMAIN = &H2    ' For starting the Find Computer
Const BIF_STATUSTEXT = &H4
Const BIF_RETURNFSANCESTORS = &H8
Const BIF_EDITBOX = &H10
Const BIF_VALIDATE = &H20             ' insist on valid result (or CANCEL)
Const BIF_BROWSEFORCOMPUTER = &H1000   ' Browsing for Computers.
Const BIF_BROWSEFORPRINTER = &H2000    ' Browsing for Printers
Const BIF_BROWSEINCLUDEFILES = &H4000  ' Browsing for Everything
'Common Folders For Open
Const ssfALTSTARTUP = &H1D
Const ssfAPPDATA = &H1A
Const ssfBITBUCKET = &HA
Const ssfCOMMONALTSTARTUP = &H1E
Const ssfCOMMONAPPDATA = &H23
Const ssfCOMMONDESKTOPDIR = &H19
Const ssfCOMMONFAVORITES = &H1F
Const ssfCOMMONPROGRAMS = &H17
Const ssfCOMMONSTARTMENU = &H16
Const ssfCOMMONSTARTUP = &H18
Const ssfCONTROLS = &H3
Const ssfCOOKIES = &H21
Const ssfDESKTOP = &H0
Const ssfDESKTOPDIRECTORY = &H10
Const ssfDRIVES = &H11
Const ssfFAVORITES = &H6
Const ssfFONTS = &H14
Const ssfHISTORY = &H22
Const ssfINTERNETCACHE = &H20
Const ssfLOCALAPPDATA = &H1C
Const ssfMYPICTURES = &H27
Const ssfNETHOOD = &H13
Const ssfNETWORK = &H12
Const ssfPERSONAL = &H5
Const ssfPRINTERS = &H4
Const ssfPRINTHOOD = &H1B
Const ssfPROFILE = &H28
Const ssfPROGRAMFILES = &H26
Const ssfPROGRAMS = &H2
Const ssfRECENT = &H8
Const ssfSENDTO = &H9
Const ssfSTARTMENU = &HB
Const ssfSTARTUP = &H7
Const ssfSYSTEM = &H25
Const ssfTEMPLATES = &H15
Const ssfWINDOWS = &H24

'On Error GoTo noSel
Set obj = CreateObject("Shell.Application")
Set objOrigen = obj.NameSpace(folder)
Set objFld = obj.BrowseForFolder(Principal.hwnd, "Buscando Carpeta :", BIF_NEWDIALOGSTYLE Or BIF_RETURNONLYFSDIRS Or BIF_EDITBOX, "")
If objFld Is Nothing Then
BrowseForFolder = ""
Else
Set objPath = objFld.Items.Item
BrowseForFolder = objPath.Path
sFilePath = objPath.Path
End If

End Function
