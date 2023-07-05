Attribute VB_Name = "rename"
Option Explicit

Public Sub ImprimirNombresTablas()
  Dim tabla As ListObject
  Dim ws As Worksheet
  Dim newName As String

  For Each ws In ThisWorkbook.Worksheets
    If ws.Name = "BASE P2" Then
      For Each tabla In ws.ListObjects
'        newName = tabla.Name & "1"
'        tabla.Name = newName
        Debug.Print tabla.Name ' Imprimir el nombre de la tabla en la ventana de "Immediate"
        ' Puedes utilizar el siguiente código para imprimir en la hoja de Excel:
        ' ws.Range("A1").Value = tabla.Name
      Next tabla
    End If
  Next ws
End Sub

Public Sub renameTables()
  Dim rng As range, item As Variant
  
  Set rng = ThisWorkbook.Worksheets("BASE P2").range("BW2", ThisWorkbook.Worksheets("BASE P2").range("BW2").End(xlDown))
  With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With
  
  For Each item In rng
    Cells.Replace What:=CStr(item.value), Replacement:=CStr(item.Offset(, 1).value), LookAt:= _
    xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
    ReplaceFormat:=False
    DoEvents
  Next item
  
  With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
  End With
  
End Sub
