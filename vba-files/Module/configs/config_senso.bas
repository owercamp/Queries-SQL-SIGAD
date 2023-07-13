Attribute VB_Name = "config_senso"
'namespace=vba-files\Module\configs
Option Explicit

Public Sub configSenso()
  '
  ' configSensometrica: realiza la configuracion inicial de las cabeceras asi como el nombre de la tabla
  '
  '
  range("A2") = "NRO IDENFICACION"
  range("B2") = "PACIENTE"
  range("C2") = "PRUEBA PSICOSENSOMETRICA"
  range("D2") = "DIAGNOSTICO PPAL"
  range("E2") = "DIAGNOSTICO OBS"
  range("F2") = "DIAGNOSTICO REL/1"
  range("G2") = "DIAGNOSTICO REL/2"
  range("H2") = "DIAGNOSTICO REL/3"
  range("I2") = "CONTROLES MENSUALES"
  range("J2") = "CONTROLES BIMENSUAL"
  range("K2") = "CONTROLES TRIMESTRALES"
  range("L2") = "CONTROLES 6 MESES"
  range("M2") = "CONTROLES 1 A" & ChrW(209) & "O"
  range("N2") = "CONTROLES CONFIRMATORIA"
  range("O2") = "id_diagnostico_ppal"
  range("P2") = "emo_id(orden_lista_trabajadoresid)"
  range("Q2") = "ID_PSICOSENSOMETRICA"
  range("R2") = "SCRIPT psicosensometrica"
  range("S2") = "SCRIPT psicosenso_diagnos_observaciones_pivot"
  range("T2") = "SCRIPT psicosensometricas_recomendaciones_pivot"
  ActiveSheet.ListObjects.Add(xlSrcRange, range("$A$2:$T$3"), , xlYes).Name = _
  "tbl_psicosensometrica"
  ActiveSheet.ListObjects("tbl_psicosensometrica").TableStyle = "TableStyleLight9"

  range("tbl_psicosensometrica[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Style = "Neutral"
  range("tbl_psicosensometrica[[#Data],[emo_id(orden_lista_trabajadoresid)]]").Style = "Notas"
  range("tbl_psicosensometrica[[#Headers],[SCRIPT psicosensometrica]:[SCRIPT psicosensometricas_recomendaciones_pivot]]").Style = "Celda de comprobaci" & ChrW(243) & "n"
  range("tbl_psicosensometrica[[#Data],[SCRIPT psicosensometrica]:[SCRIPT psicosensometricas_recomendaciones_pivot]]").Style = "Salida"
  
  Call formatTable("tbl_psicosensometrica")
  
End Sub

