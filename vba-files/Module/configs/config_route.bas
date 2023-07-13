Attribute VB_Name = "config_route"
'namespace=vba-files\Module\configs
Option Explicit

Public Sub configRoute()
  '
  ' configRutas: realiza la configuracion inicial de las cabeceras asi como el nombre de la tabla
  '
  '
  range("B3") = "nombre"
  range("C3") = "ruta"
  range("B4") = "INFO"
  range("C4") = "C:\Users\SOANDES-DSOFT\Documents\MACRO\ARCHIVO"
  range("B5") = "CONSOLIDADO"
  range("C5") = _
  "C:\Users\SOANDES-DSOFT\Documents\Ower Campos\Solicitud DRA Diana Casta" & ChrW(241) & "eda\Cargue Reporte Empresas"
  range("B6") = "SCRIPT"
  range("C6") = "C:\Users\SOANDES-DSOFT\Documents\Ower Campos\Script"
  range("B7") = "CARGOS"
  range("C7") = _
  "C:\Users\SOANDES-DSOFT\Documents\Soandes Procesos\Plantillas\ultima_actualizada\Cargos - Empresas"
  range("B8") = "BACKUP"
  range("C8") = _
  "C:\Users\SOANDES-DSOFT\Documents\Ower Campos\Backup Libro"
  range("B9") = "SQL"
  range("C9") = "C:\Users\SOANDES-DSOFT\Documents\MACRO\"
  ActiveSheet.ListObjects.Add(xlSrcRange, range("$B$3:$C$9"), , xlYes).Name = _
  "tbl_rutas"

  range("E3") = "tabla"
  range("F3") = "auto incremental"
  range("E4") = "idOrdenListaTrabajadores"
  range("F4") = "0"
  range("E5") = "idEmo"
  range("F5") = "0"
  range("E6") = "idAudiometria"
  range("F6") = "0"
  range("E7") = "idOptometria"
  range("F7") = "0"
  range("E8") = "idDiagnostico"
  range("F8") = "0"
  range("E9") = "idVisiometria"
  range("F9") = "0"
  range("E10") = "idEspirometria"
  range("F10") = "0"
  range("E11") = "idOsteomuscular"
  range("F11") = "0"
  range("E12") = "idComplementarios"
  range("F12") = "0"
  range("E13") = "idPsicotecnica"
  range("F13") = "0"
  range("E14") = "idPsicosensomentrica"
  range("F14") = "0"
  ActiveSheet.ListObjects.Add(xlSrcRange, range("$E$3:$F$14"), , xlYes).Name = _
  "tbl_ids"
  Cells.EntireColumn.AutoFit

End Sub

