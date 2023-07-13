Attribute VB_Name = "config_comple"
'namespace=vba-files\Module\configs
Option Explicit

Public Sub configComple()
  '
  ' configComplementarios: realiza la configuracion inicial de las cabeceras asi como el nombre de la tabla
  '
  '
  range("A3") = "NRO IDENFICACION"
  range("B3") = "PROCEDIMIENTO"
  range("C3") = "DIAG_ PPAL"
  range("D3") = "DIAG_ PPAL OBS"
  range("E3") = "DIAG_ REL/1"
  range("F3") = "DIAG_ REL/2"
  range("G3") = "DIAG_ REL/3"
  range("H3") = "HALLAZGOS"
  range("I3") = "emo_id(orden_lista_trabajadoresid)"
  range("J3") = "ID_COMPLEMENTARIOS"
  range("K3") = "SCRIPT complementarios"
  range("L3") = "SCRIPT complementarios_diagnos_observaciones_pivot"
  range("M3") = "LLAVE"
  ActiveSheet.ListObjects.Add(xlSrcRange, range("$A$3:$M$4"), , xlYes).Name = _
  "tbl_complementarios"
  ActiveSheet.ListObjects("tbl_complementarios").TableStyle = "TableStyleLight9"

  range("tbl_complementarios[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Style = "Neutral"
  range("tbl_complementarios[[#Data],[emo_id(orden_lista_trabajadoresid)]]").Style = "Notas"
  range("tbl_complementarios[[#Headers],[LLAVE]]").Style = "Neutral"
  range("tbl_complementarios[[#Data],[LLAVE]]").Style = "Notas"
  range("tbl_complementarios[[#Headers],[SCRIPT complementarios]:[SCRIPT complementarios_diagnos_observaciones_pivot]]").Style = "Celda de comprobaci" & ChrW(243) & "n"
  range("tbl_complementarios[[#Data],[SCRIPT complementarios]:[SCRIPT complementarios_diagnos_observaciones_pivot]]").Style = "Salida"
  
  Call formatTable("tbl_complementarios")
  
End Sub

