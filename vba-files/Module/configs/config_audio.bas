Attribute VB_Name = "config_audio"
'namespace=vba-files\Module\configs
Option Explicit

Public Sub configAudio()
  '
  ' configAudio: realiza la configuracion inicial de las cabeceras asi como el nombre de la tabla
  '
  '
  range("A3") = "NROAIDENFICACION"
  range("B3") = "EPP ESPECIFICO / AUDITIVO"
  range("C3") = "EPP ESPECIFICO / AUDITIVO COPA"
  range("D3") = "EPP ESPECIFICO / AUDITIVO INSERCION"
  range("E3") = "EPP ESPECIFICO / AUDITIVO DOBLE"
  range("F3") = "PABELLON AURIC_ OIDO DER_"
  range("G3") = "PABELLON AURIC_ OIDO DER_ OBS"
  range("H3") = "PABELLON AURIC_ OIDO IZQ_"
  range("I3") = "PABELLON AURIC_ OIDO IZQ_ OBS"
  range("J3") = "CONDUCTO AUDIT_ OIDO DER_"
  range("K3") = "CONDUCTO AUDIT_ OIDO DER_ OBS"
  range("L3") = "CONDUCTO AUDIT_ OIDO IZQ_"
  range("M3") = "CONDUCTO AUDIT_ OIDO IZQ_ OBS"
  range("N3") = "MEMBRANA TIMP_ OIDO DER"
  range("O3") = "MEMBRANA TIMP_ OIDO DER_ OBS"
  range("P3") = "MEMBRANA TIMP_ OIDO IZQ"
  range("Q3") = "MEMBRANA TIMP_ OIDO IZQ_ OBS"
  range("R3") = "TIPO DE EXAMEN"
  range("S3") = "OD 500"
  range("T3") = "OD 1000"
  range("U3") = "OD 2000"
  range("V3") = "OD 3000"
  range("W3") = "OD 4000"
  range("X3") = "OD 6000"
  range("Y3") = "OD 8000"
  range("Z3") = "PTA OD"
  range("AA3") = "OI 500"
  range("AB3") = "OI 1000"
  range("AC3") = "OI 2000"
  range("AD3") = "OI 3000"
  range("AE3") = "OI 4000"
  range("AF3") = "OI 6000"
  range("AG3") = "OI 8000"
  range("AH3") = "PTA OI"
  range("AI3") = "CONTROL SEGUN PVE"
  range("AJ3") = "CONFIRMATORIA"
  range("AK3") = "REMISION ORL"
  range("AL3") = "PRUEBAS COMPLEMENTARIAS"
  range("AM3") = "LIMPIEZA DE OIDO"
  range("AN3") = "LIMPIEZA OD"
  range("AO3") = "LIMPIEZA OI"
  range("AP3") = "REPOSO AUDITIVO EXTRALAB"
  range("AQ3") = "ROTAR DIADEMA TELEFONICA"
  range("AR3") = "CONDUCIR CON VENTANAS CERRADAS"
  range("AS3") = "USO DE EPP AUDITIVO"
  range("AT3") = "CONTROLES MENSUALES"
  range("AU3") = "CONTROLES_BIMESTRALES"
  range("AV3") = "CONTROLES TRIMESTRALES"
  range("AW3") = "CONTROLES 6 MESES"
  range("AX3") = "CONTROLES 1 A" & ChrW(209) & "O"
  range("AY3") = "DIAG PPAL"
  range("AZ3") = "DIAG INTERNO"
  range("BA3") = "DIAG GATI SO"
  range("BB3") = "tipo1 au_oido"
  range("BC3") = "tipo2 au_oido"
  range("BD3") = "frecuencia OD"
  range("BE3") = "frecuencia OI"
  range("BF3") = "emo_id(orden_lista_trabajadoresid)"
  range("BG3") = "ID_AUDIOMETRIA"
  range("BH3") = "SCRIPT au_audiometria"
  range("BI3") = "SCRIPT au_audiometria_recomendacion"
  range("BJ3") = "SCRIPT au_oido"
  range("BK3") = "LLAVE"
  ActiveSheet.ListObjects.Add(xlSrcRange, range("$A$3:$BK$4"), , xlYes).Name = _
  "tbl_audio"
  ActiveSheet.ListObjects("tbl_audio").TableStyle = "TableStyleLight9"

  range("tbl_audio[[#Headers],[PTA OD]]").Style = "Neutral"
  range("tbl_audio[[#Data],[PTA OD]]").Style = "Notas"
  range("tbl_audio[[#Headers],[PTA OI]]").Style = "Neutral"
  range("tbl_audio[[#Data],[PTA OI]]").Style = "Notas"
  range("tbl_audio[[#Headers],[LLAVE]]").Style = "Neutral"
  range("tbl_audio[[#Data],[LLAVE]]").Style = "Notas"
  range("tbl_audio[[#Headers],[tipo1 au_oido]:[emo_id(orden_lista_trabajadoresid)]]").Style = "Neutral"
  range("tbl_audio[[#Data],[tipo1 au_oido]:[emo_id(orden_lista_trabajadoresid)]]").Style = "Notas"
  range("tbl_audio[[#Headers],[SCRIPT au_audiometria]:[SCRIPT au_oido]]").Style = "Celda de comprobaci" & ChrW(243) & "n"
  range("tbl_audio[[#Data],[SCRIPT au_audiometria]:[SCRIPT au_oido]]").Style = "Salida"
  
  Call formatTable("tbl_audio")
  
End Sub
