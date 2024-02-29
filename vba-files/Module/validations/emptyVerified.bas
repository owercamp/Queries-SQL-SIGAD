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

  ' Loop Until the cell To the left of the active cell is empty
  Do Until IsEmpty(ActiveCell.Offset(, -2).value)
    valor = Len(ActiveCell.value)
    ' If the length of the active cell value is greater than 5, truncate And remove commas
    If valor > 5 Then
      If VBA.InStr(1, ActiveCell.value, "0", vbTextCompare) = 1 Then
        Dim pos As Integer
        pos = VBA.InStr(1, ActiveCell.value, ",", vbTextCompare) + 2
        ActiveCell.value = VBA.Left$(ActiveCell.value, pos)
      Else
        ActiveCell.value = VBA.Mid$(ActiveCell.value, 1, 2)
        ActiveCell.value = VBA.Replace(ActiveCell.value, ",", "")
      End If
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
  ' This subroutine scales the active cell value by 100 And formats it To two decimal places If the value does Not contain a comma.
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
    ' Extract the incapacity And number values from the active cell And the cell 1 column To the right.
    incapacity = ActiveCell.Value
    number = ActiveCell.Offset(, 1).Value
    ' Check If the active cell is Not numeric And Not empty, Then update the active cell And the cell 1 column To the right.
    If Not IsNumeric(ActiveCell) And Not IsEmpty(ActiveCell) Then
      ActiveCell.Value = number
      ActiveCell.Offset(, 1).Value = incapacity
      ' If the active cell And the cell 1 column To the right are Not numeric, clear the active cell.
    Elseif Not IsNumeric(ActiveCell) And Not IsNumeric(ActiveCell.Offset(, 1)) Then
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