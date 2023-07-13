Attribute VB_Name = "config_psico"
'namespace=vba-files\Module\configs
Option Explicit

Public Sub configPsico()
  '
  ' configPsicotecnica: realiza la configuracion inicial de las cabeceras asi como el nombre de la tabla
  '
  '
  range("A1") = "NRO IDENFICACION"
  range("B1") = "PACIENTE"
  range("C1") = "PRUEBA PSICOTECNICA"
  range("D1") = "DIAGNOSTICO PPAL (CUMPLE, NO CUMPLE)"
  range("E1") = "DIAGNOSTICO OBS"
  range("F1") = "emo_id(orden_lista_trabajadoresid)"
  range("G1") = "ID_PSICOTECNICA"
  range("H1") = "SCRIPT psicotecnica"
  range("I1") = "LLAVE"
  ActiveSheet.ListObjects.Add(xlSrcRange, range("$A$1:$I$2"), , xlYes).Name = _
  "tbl_psicotecnica"
  ActiveSheet.ListObjects("tbl_psicotecnica").TableStyle = "TableStyleLight9"

  range("tbl_psicotecnica[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Style = "Neutral"
  range("tbl_psicotecnica[[#Data],[emo_id(orden_lista_trabajadoresid)]]").Style = "Notas"
  range("tbl_psicotecnica[[#Headers],[LLAVE]]").Style = "Neutral"
  range("tbl_psicotecnica[[#Data],[LLAVE]]").Style = "Notas"
  range("tbl_psicotecnica[[#Headers],[SCRIPT psicotecnica]]").Style = "Celda de comprobaci" & ChrW(243) & "n"
  range("tbl_psicotecnica[[#Data],[SCRIPT psicotecnica]]").Style = "Salida"
  
  Call formatTable("tbl_psicotecnica")

End Sub

