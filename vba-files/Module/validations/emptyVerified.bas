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
