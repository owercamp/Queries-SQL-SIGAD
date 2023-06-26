VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formClear 
   Caption         =   "Limpiando..."
   ClientHeight    =   732
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4284
   OleObjectBlob   =   "formClear.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "formClear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub UserForm_Activate()
  Call ClearNonAlphaNumeric
End Sub
