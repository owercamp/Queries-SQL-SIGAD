Attribute VB_Name = "rename"
Option Explicit

Public Sub ImprimirNombresTablas()
  Dim tabla As ListObject
  Dim ws As Worksheet
  Dim newName As String

  For Each ws In ThisWorkbook.Worksheets
    For Each tabla In ws.ListObjects
      ' newName = tabla.Name & "1"
      ' tabla.Name = newName
      Debug.Print tabla.Name ' Imprimir el nombre de la tabla en la ventana de "Immediate"
      ' Puedes utilizar el siguiente código para imprimir en la hoja de Excel:
      ' ws.Range("A1").Value = tabla.Name
    Next tabla
  Next ws
End Sub

Public Sub renameTables()

  Do While Not IsEmpty(ActiveCell)
    Cells.Replace What:=ActiveCell.Offset(, 2).value, Replacement:=ActiveCell.value, LookAt:= _
    xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
    ReplaceFormat:=False
    ActiveCell.Offset(1, 0).Select
  Loop
  
End Sub
