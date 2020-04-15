VERSION 5.00
Begin VB.Form frmFormatoSave 
   Caption         =   "FORMATO AL GRABAR"
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ButonFormato 
      Caption         =   "HTML"
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton ButonFormato 
      Caption         =   "TXT"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton ButonFormato 
      Caption         =   "RTF"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "SIN FORMATO"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "CON FORMATO"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmFormatoSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ButonFormato_Click(Index As Integer)
Select Case Index
Case 0
    sExtension = "rtf"
Case 1
    sExtension = "txt"
Case 2
    sExtension = "html"
End Select
Unload Me

End Sub

Private Sub Command1_Click(Index As Integer)

End Sub

Private Sub Form_Resize()
Dim z
 

End Sub

Private Sub Form_Unload(Cancel As Integer)
If sExtension = "" Then MsgBox "No haselegido elformato.. RTF(colores), TXT(solo tto) y HTML (con etiquetas o tags)"


End Sub
