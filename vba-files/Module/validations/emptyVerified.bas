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

Public Sub correctionAntiquity()

  Dim valor As Integer

  With Application
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual  
  End With

  Do Until IsEmpty(ActiveCell.Offset(, -2).value)
    valor = Len(ActiveCell.value)
    If valor > 3 Then
      ActiveCell = VBA.Mid$(ActiveCell.value, 1, 2)
      ActiveCell = VBA.Replace(ActiveCell, ",", "")
    End If
    ActiveCell.Offset(1, 0).Select
  Loop

  With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
  End With

End Sub
