VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formImports 
   Caption         =   "0%"
   ClientHeight    =   2088
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5052
   OleObjectBlob   =   "formImports.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "formImports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Public route As String

Private Sub UserForm_Activate()

  Dim msg As MsoButtonState

  '''''''''''''''''''''''''''''''''''''''''''''''''
  '''        APERTURA DEL LIBRO ARCHIVO         '''
  '''''''''''''''''''''''''''''''''''''''''''''''''
  Set origin = Workbooks.Open(route)

  msg = MsgBox("Advertencia fueron verificadas las cabeceras de las tablas del archivo que se encuentra en:" + _
  vbNewLine + vbNewLine + "Nota:" + vbNewLine + "Las CABECERAS no pueden estar vacias" + vbNewLine + vbNewLine + " ruta del archivo:" + vbNewLine + CStr(route) & "", vbExclamation + vbYesNo, "Cabeceras Vacias")

  If msg = vbYes Then
    ''' SE LLAMA A LA FUNCION PARA EXTRAER LA INFORMACION '''
    Call extraerdatos
  Else
    Unload Me
    Windows(origin.Name).Activate
  End If

End Sub

Private Sub UserForm_Initialize()

  Me.ProgressBarGeneral = 0
  Me.ProgressBarOneforOne = 0
  route = Application.getOpenFilename

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Unload Me
End Sub

