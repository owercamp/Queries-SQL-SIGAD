VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formMix 
   Caption         =   "Forms"
   ClientHeight    =   1872
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   4332
   OleObjectBlob   =   "formMix.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "formMix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub btn_confirm_Click()
  If (Me.lblMsg.Caption = "Ingrese la cantidad de ENFASIS") Then
    number_emphasis = CInt(Me.txt_cantidad)
    Me.Hide
    Me.lblMsg.Caption = Empty
    Me.txt_cantidad = Empty
  ElseIf (Me.lblMsg.Caption = "Ingrese la cantidad de DIAGNOSTICOS") Then
    number_diag = CInt(Me.txt_cantidad)
    Me.Hide
    Me.lblMsg.Caption = Empty
    Me.txt_cantidad = Empty
  ElseIf (Me.lblMsg.Caption = "Por favor ingrese el numero ID correspondiente a la orden en SIGAD") Then
    idOrden = Me.txt_cantidad
    Me.Hide
    Me.lblMsg.Caption = Empty
    Me.Caption = Empty
    Me.txt_cantidad = Empty
  ElseIf (Me.lblMsg.Caption = "Ingrese el n" & Chr(250) & "mero de orden SIGAD") Then
    sigad = Me.txt_cantidad
    Me.Hide
    Me.lblMsg.Caption = Empty
    Me.txt_cantidad = Empty
  End If
End Sub
