Attribute VB_Name = "emptyVerified"
'namespace=vba-files\Module\validations
Option Explicit

Public Sub verifiedEmpty()

  Select Case origin.ActiveSheet.Name
   Case "DIAGNOSTICOS"
    range("tbl_diagnosticos[[#Headers],[IDENTIFICACION]]").End(xlDown).Offset(1, 0).Select
   Case "ENFASIS"
    range("tbl_enfasis[[#Headers],[IDENTIFICACION]]").End(xlDown).Offset(1, 0).Select
   Case "TRABAJADORES"
    range("tbl_trabajadores[[#Headers],[estado]]").End(xlDown).Offset(1, 0).Select
   Case "EMO"
    range("tbl_emo[[#Headers],[NRO IDENFICACION]]").End(xlDown).Offset(1, 0).Select
   Case "AUDIO"
    range("tbl_audio[[#Headers],[NROAIDENFICACION]]").End(xlDown).Offset(1, 0).Select
   Case "OPTO"
    range("tbl_opto[[#Headers],[NRO IDENFICACION]]").End(xlDown).Offset(1, 0).Select
   Case "VISIO"
    range("tbl_visio[[#Headers],[NRO IDENFICACION]]").End(xlDown).Offset(1, 0).Select
   Case "ESPIRO"
    range("tbl_espiro_info[[#Headers],[NRO IDENFICACION]]").End(xlDown).Offset(1, 0).Select
   Case "OSTEO"
    range("tbl_osteo[[#Headers],[NRO IDENFICACION]]").End(xlDown).Offset(1, 0).Select
   Case "COMPLEMENTARIOS"
    range("tbl_complementarios[[#Headers],[NRO IDENFICACION]]").End(xlDown).Offset(1, 0).Select
   Case "PSICOSENSOMETRICA"
    range("tbl_psicosensometrica[[#Headers],[NRO IDENFICACION]]").End(xlDown).Offset(1, 0).Select
   Case "PSICOTECNICA"
    range("tbl_psicotecnica[[#Headers],[NRO IDENFICACION]]").End(xlDown).Offset(1, 0).Select
  End Select

  selection.EntireRow.Select
  range(selection, selection.Cells(Rows.Count)).Select
  selection.EntireRow.Delete shift:=xlUp

End Sub

' This subroutine corrects the antiquity of the active cell value.
Public Sub correctionAntiquity()

  Dim valor As Integer

  With Application
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual  
  End With

  ' Loop until the cell to the left of the active cell is empty
  Do Until IsEmpty(ActiveCell.Offset(, -2).Value)
    valor = Len(ActiveCell.Value)
    ' If the length of the active cell value is greater than 5, truncate and remove commas
    If valor > 5 Then
      ActiveCell.Value = VBA.Mid$(ActiveCell.Value, 1, 2)
      ActiveCell.Value = VBA.Replace(ActiveCell.Value, ",", "")
    End If
    ActiveCell.Offset(1, 0).Select
  Loop

  With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
  End With

End Sub

Public Sub Size()
  ' This subroutine scales the active cell value by 100 and formats it to two decimal places if the value does not contain a comma.
  Do Until IsEmpty(ActiveCell.Offset(0, -2))
    If VBA.InStr(ActiveCell.value, ",") = 0 Then
      ActiveCell = ActiveCell.value / 100
      ActiveCell.NumberFormat = "0.00"
    End If
    ActiveCell.Offset(1, 0).Select
  Loop
End Sub

Public Sub incapacity()
  ' This subroutine processes the incapacity data in the active cell.
  Dim incapacity As String
  Dim number As String

  With Application
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
  End With

  Do Until IsEmpty(ActiveCell.Offset(, -7))
    ' Extract the incapacity and number values from the active cell and the cell 1 column to the right.
    incapacity = ActiveCell.Value
    number = ActiveCell.Offset(, 1).Value
    ' Check if the active cell is not numeric and not empty, then update the active cell and the cell 1 column to the right.
    If Not IsNumeric(ActiveCell) And Not IsEmpty(ActiveCell) Then
      ActiveCell.Value = number
      ActiveCell.Offset(, 1).Value = incapacity
    ' If the active cell and the cell 1 column to the right are not numeric, clear the active cell.
    ElseIf Not IsNumeric(ActiveCell) And Not IsNumeric(ActiveCell.Offset(, 1)) Then
      ActiveCell.Value = ""
    End If
    ActiveCell.Offset(1, 0).Select
  Loop

  With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
  End With
End Sub