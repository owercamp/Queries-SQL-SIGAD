Attribute VB_Name = "config_visio"
'namespace=vba-files\Module\configs
Option Explicit

Public Sub configVisio()
  '
  ' configVisio: realiza la configuracion inicial de las cabeceras asi como el nombre de la tabla
  '
  '
  range("A3") = "NRO IDENFICACION"
  range("B3") = "VISIO/ANT_ LABORAL ILUMINACION INADECUADA"
  range("C3") = "VISIO/ANT_ LABORALVISIO RADIACIONES UV"
  range("D3") = "VISIO/ANT_ LABORAL MALA VENTILACION"
  range("E3") = "VISIO/ANT_ LABORAL GASES TOXICOS"
  range("F3") = "SINTOMAS FOTOFOBIA"
  range("G3") = "SINTOMAS OJO ROJO"
  range("H3") = "SINTOMAS LAGRIMEO"
  range("I3") = "SINTOMAS VISION BORROSA"
  range("J3") = "SINTOMAS ARDOR"
  range("K3") = "SINTOMAS VISION DOBLE"
  range("L3") = "SINTOMAS CANSANCIO"
  range("M3") = "SINTOMAS MALA VISION CERCANA"
  range("N3") = "SINTOMAS DOLOR"
  range("O3") = "SINTOMAS MALA VISON LEJANA"
  range("P3") = "SINTOMAS SECRECION"
  range("Q3") = "SINTOMAS CEFALEA"
  range("R3") = "OTROS SINTOMAS"
  range("S3") = "CABEZA PARPADOS"
  range("T3") = "CABEZA PARPADOS OBS"
  range("U3") = "CABEZA CONJUNTIVAS"
  range("V3") = "CABEZA OBS CONJUNTIVAS"
  range("W3") = "CABEZA ESCLERAS"
  range("X3") = "CABEZA OBS ESCLERAS"
  range("Y3") = "CABEZA PUPILAS"
  range("Z3") = "CABEZA PUPILAS OBS"
  range("AA3") = "Imp/DIAG VL0OD NORMAL"
  range("AB3") = "Imp/DIAG VL0OI NORMAL"
  range("AC3") = "Imp/DIAG VP0OD NORMAL"
  range("AD3") = "Imp/DIAG VP0OI NORMAL"
  range("AE3") = "Imp/DIAG VL0OD DISMINUIDO"
  range("AF3") = "Imp/DIAG VL0OI DISMINUIDO"
  range("AG3") = "Imp/DIAG VP0OD DISMINUIDO"
  range("AH3") = "Imp/DIAG VP0OI DISMINUIDO"
  range("AI3") = "Imp/DIAG VL0OD NORMAL RX"
  range("AJ3") = "Imp/DIAG VL0OI NORMAL RX"
  range("AK3") = "Imp/DIAG VP0OD NORMAL RX"
  range("AL3") = "Imp/DIAG VP0OI NORMAL RX"
  range("AM3") = "Imp/DIAG VL0OD DISMINUIDO RX"
  range("AN3") = "Imp/DIAG VL0OI DISMINUIDO RX"
  range("AO3") = "Imp/DIAG VP0OD DISMINUIDO RX"
  range("AP3") = "Imp/DIAG VP0OI DISMINUIDO RX"
  range("AQ3") = "0"
  range("AR3") = "RESULTADO VISIO"
  range("AS3") = "Imp/DIAG OBS"
  range("AT3") = "REC CORRECCION VISUAL PARA TRABAJAR"
  range("AU3") = "REC USO RX PARA VISION PROX"
  range("AV3") = "REC USO AR VIDEO TRMINAL"
  range("AW3") = "REC USO RX DESCANSO"
  range("AX3") = "REC USO LENTES PROT_ SOLAR"
  range("AY3") = "REC USO PERMANENTE RX OPTICA"
  range("AZ3") = "REC USO EPP VISUAL"
  range("BA3") = "REC PYP"
  range("BB3") = "REC PAUSAS ACTIVAS"
  range("BC3") = "REC LUBRICANTE OCULAR"
  range("BD3") = "RECOMENDACIONES OBS"
  range("BE3") = "REM_ VALORACION OFTALM_"
  range("BF3") = "REM_ VALORACION OPTO_ COMPLETA"
  range("BG3") = "REM_ TOPOGRAFIA CORNEAL"
  range("BH3") = "REM_ TRATAM_ ORTOPTICA"
  range("BI3") = "REM_ TEST FARNSWORTH"
  range("BJ3") = "REALIZAR PRUEBA AMBULATORIA"
  range("BK3") = "OTRAS REMISIONES"
  range("BL3") = "CONTROL MENSUAL"
  range("BM3") = "CONTROLES_BIMESTRALES"
  range("BN3") = "CONTROL TRIMESTRAL"
  range("BO3") = "CONTROL 6 MESES"
  range("BP3") = "CONTROL 1 A" & ChrW(209) & "O"
  range("BQ3") = "CONTROL CONFIRMATORIA"
  range("BR3") = "emo_id(orden_lista_trabajadoresid)"
  range("BS3") = "ID_VISIOMETRIA"
  range("BT3") = "SCRIPT vi_visiometria"
  range("BU3") = "SCRIPT vi_visiometria_antecedentes"
  range("BV3") = "SCRIPT vi_visiometria_sintomas"
  range("BW3") = "SCRIPT vi_vl"
  range("BX3") = "SCRIPT vi_vp"
  range("BY3") = "SCRIPT vi_visiometria_recomendaciones"
  range("BZ3") = "SCRIPT vi_visiometria_remisiones"
  range("CA3") = "LLAVE"
  ActiveSheet.ListObjects.Add(xlSrcRange, range("$A$3:$CA$4"), , xlYes).Name = _
  "tbl_visio"
  ActiveSheet.ListObjects("tbl_visio").TableStyle = "TableStyleLight9"

  range("tbl_visio[[#Headers],[0]:[RESULTADO VISIO]]").Style = "Neutral"
  range("tbl_visio[[#Data],[0]:[RESULTADO VISIO]]").Style = "Notas"
  range("tbl_visio[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Style = "Neutral"
  range("tbl_visio[[#Data],[emo_id(orden_lista_trabajadoresid)]]").Style = "Notas"
  range("tbl_visio[[#Headers],[LLAVE]]").Style = "Neutral"
  range("tbl_visio[[#Data],[LLAVE]]").Style = "Notas"
  range("tbl_visio[[#Headers],[SCRIPT vi_visiometria]:[SCRIPT vi_visiometria_remisiones]]").Style = "Celda de comprobaci" & ChrW(243) & "n"
  range("tbl_visio[[#Data],[SCRIPT vi_visiometria]:[SCRIPT vi_visiometria_remisiones]]").Style = "Salida"
  
  Call formatTable("tbl_visio")

End Sub
