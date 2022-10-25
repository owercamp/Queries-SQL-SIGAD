Attribute VB_Name = "DataEspiro"
Option Explicit

Sub EspiroData()
  Dim espiro_destiny_dictionary As Scripting.Dictionary
  Dim espiro_origin_dictionary As Scripting.Dictionary
  Dim espiro_destiny_header, espiro_origin_header, espiro_origin_value As Object
  Dim ItemEspiroDestiny, ItemEspiroOrigin, ItemData As Variant

  Set espiro_origin = origin.Worksheets("ESPIRO") '' ESPIRO DEL LIBRO ORIGEN ''
  espiro_destiny.Select
  ActiveSheet.Range("A5").Select
  Set espiro_destiny_header = espiro_destiny.Range("A3", espiro_destiny.Range("A3").End(xlToRight))
  Set espiro_origin_header = espiro_origin.Range("A1", espiro_origin.Range("A1").End(xlToRight))
  Set espiro_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set espiro_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (espiro_origin.Range("A2") <> Empty And espiro_origin.Range("A3") <> Empty) Then
    Set espiro_origin_value = espiro_origin.Range("A2", espiro_origin.Range("A2").End(xlDown))
  ElseIf (espiro_origin.Range("A2") <> Empty And espiro_origin.Range("A3") = Empty) Then
    Set espiro_origin_value = espiro_origin.Range("A2")
  End If

  '/***
  '   En los diccionarios de "espiro_destiny_dictionary" y  "espiro_origin_dictionary"
  '   se almacena los numeros de la columnas.
  '*/

  ' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO
  For Each ItemEspiroDestiny In espiro_destiny_header
    On Error Goto espiroError
    espiro_destiny_dictionary.Add espiro_headers(ItemEspiroDestiny), (ItemEspiroDestiny.Column - 1)
  Next ItemEspiroDestiny

  ' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN
  For Each ItemEspiroOrigin In espiro_origin_header
    On Error Goto espiroError
    espiro_origin_dictionary.Add espiro_headers(ItemEspiroOrigin), (ItemEspiroOrigin.Column - 1)
  Next ItemEspiroOrigin

  numbers = 1
  porcentaje = 0
  counts = espiro_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  For Each ItemData In espiro_origin_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & espiro_destiny.Name
      porcentaje = porcentaje + vals
      porcentajeGeneral = porcentajeGeneral + valsGeneral
      formImports.ProgressBarOneforOne.Width = oneForOne
      formImports.ProgressBarGeneral.Width = generalAll
      formImports.porcentageGeneral.Caption = CStr(VBA.Round(porcentajeGeneral * 100, 1)) & "%"
      formImports.porcentageOneoforOne.Caption = CStr(VBA.Round(porcentaje * 100, 1)) & "%"
      formImports.Caption = CStr(nameCompany)
      If formImports.ProgressBarGeneral.Width > (formImports.content_ProgressBarGeneral.Width / 2) Then: formImports.porcentageGeneral.ForeColor = RGB(255, 255, 255)
        If formImports.ProgressBarGeneral.Width < (formImports.content_ProgressBarGeneral.Width / 2) Then: formImports.porcentageGeneral.ForeColor = RGB(0, 0, 0)
          If formImports.ProgressBarOneforOne.Width > (formImports.content_ProgressBarOneforOne.Width / 2) Then: formImports.porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
            If formImports.ProgressBarOneforOne.Width < (formImports.content_ProgressBarOneforOne.Width / 2) Then: formImports.porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
              ActiveCell.offset(, espiro_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.offset(, espiro_origin_dictionary( "NRO IDENFICACION")))
              ActiveCell.offset(, espiro_destiny_dictionary("ALERGIAS")) = charters(ItemData.offset(, espiro_origin_dictionary( "ALERGIAS")))
              ActiveCell.offset(, espiro_destiny_dictionary("ALERGIAS OBS")) = charters(ItemData.offset(, espiro_origin_dictionary( "ALERGIAS OBS")))
              ActiveCell.offset(, espiro_destiny_dictionary("TUBERCULOSIS")) = charters(ItemData.offset(, espiro_origin_dictionary( "TUBERCULOSIS")))
              ActiveCell.offset(, espiro_destiny_dictionary("TOS CRONICA")) = charters(ItemData.offset(, espiro_origin_dictionary( "TOS CRONICA")))
              ActiveCell.offset(, espiro_destiny_dictionary("GRIPAS FRECUENTES")) = charters(ItemData.offset(, espiro_origin_dictionary( "GRIPAS FRECUENTES")))
              ActiveCell.offset(, espiro_destiny_dictionary("FARINGITIS")) = charters(ItemData.offset(, espiro_origin_dictionary( "FARINGITIS")))
              ActiveCell.offset(, espiro_destiny_dictionary("FARINGOAMIGDALITIS")) = charters(ItemData.offset(, espiro_origin_dictionary( "FARINGOAMIGDALITIS")))
              ActiveCell.offset(, espiro_destiny_dictionary("RINITIS")) = charters(ItemData.offset(, espiro_origin_dictionary( "RINITIS")))
              ActiveCell.offset(, espiro_destiny_dictionary("SINUSITIS")) = charters(ItemData.offset(, espiro_origin_dictionary( "SINUSITIS")))
              ActiveCell.offset(, espiro_destiny_dictionary("CX TORAX")) = charters(ItemData.offset(, espiro_origin_dictionary( "CX TORAX")))
              ActiveCell.offset(, espiro_destiny_dictionary("CX TORAX OBS")) = charters(ItemData.offset(, espiro_origin_dictionary( "CX TORAX OBS")))
              ActiveCell.offset(, espiro_destiny_dictionary("ASMA BRONQUIAL")) = charters(ItemData.offset(, espiro_origin_dictionary( "ASMA BRONQUIAL")))
              ActiveCell.offset(, espiro_destiny_dictionary("BRONQUITIS")) = charters(ItemData.offset(, espiro_origin_dictionary( "BRONQUITIS")))
              ActiveCell.offset(, espiro_destiny_dictionary("NEUMONIA")) = charters(ItemData.offset(, espiro_origin_dictionary( "NEUMONIA")))
              ActiveCell.offset(, espiro_destiny_dictionary("TRAUMA COSTAL")) = charters(ItemData.offset(, espiro_origin_dictionary( "TRAUMA COSTAL")))
              ActiveCell.offset(, espiro_destiny_dictionary("CANCER")) = charters(ItemData.offset(, espiro_origin_dictionary( "CANCER")))
              ActiveCell.offset(, espiro_destiny_dictionary("CANCER OBS")) = charters(ItemData.offset(, espiro_origin_dictionary( "CANCER OBS")))
              ActiveCell.offset(, espiro_destiny_dictionary("OTROS RESPIRATORIOS")) = charters(ItemData.offset(, espiro_origin_dictionary( "OTROS RESPIRATORIOS")))
              ActiveCell.offset(, espiro_destiny_dictionary("RIESGO QUIMICO / POLVOS")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "RIESGO QUIMICO / POLVOS")))
              ActiveCell.offset(, espiro_destiny_dictionary("RIESGO QUIMICO / FIBRAS")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "RIESGO QUIMICO / FIBRAS")))
              ActiveCell.offset(, espiro_destiny_dictionary("RIESGO QUIMICO / LIQUIDOS")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "RIESGO QUIMICO / LIQUIDOS")))
              ActiveCell.offset(, espiro_destiny_dictionary("RIESGO QUIMICO /GASES")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "RIESGO QUIMICO /GASES")))
              ActiveCell.offset(, espiro_destiny_dictionary("RIESGO QUIMICO / VAPORES")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "RIESGO QUIMICO / VAPORES")))
              ActiveCell.offset(, espiro_destiny_dictionary("RIESGO QUIMICO / HUMOS")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "RIESGO QUIMICO / HUMOS")))
              ActiveCell.offset(, espiro_destiny_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "RIESGO QUIMICO /MATERIAL PARTICULADO")))
              ActiveCell.offset(, espiro_destiny_dictionary("OTROS RIESGOS QUIMICOS")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "OTROS RIESGOS QUIMICOS")))
              ActiveCell.offset(, espiro_destiny_dictionary("EPP ESPECIFICO / TAPABOCA")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "EPP ESPECIFICO / TAPABOCA")))
              ActiveCell.offset(, espiro_destiny_dictionary("EPP ESPECIFICO / RESPIRADOR")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "EPP ESPECIFICO / RESPIRADOR")))
              ActiveCell.offset(, espiro_destiny_dictionary("ACT_ FISICA")) = charters(ItemData.offset(, espiro_origin_dictionary( "ACT_ FISICA")))
              ActiveCell.offset(, espiro_destiny_dictionary("FUMA")) = charters(ItemData.offset(, espiro_origin_dictionary( "FUMA")))
              ActiveCell.offset(, espiro_destiny_dictionary("CIGARRILLOS DIA")) = charters(ItemData.offset(, espiro_origin_dictionary( "CIGARRILLOS DIA")))
              ActiveCell.offset(, espiro_destiny_dictionary("FRECUENCIA")) = charters(ItemData.offset(, espiro_origin_dictionary( "FRECUENCIA")))
              ActiveCell.offset(, espiro_destiny_dictionary("TIEMPO EN ANOS")) = charters(ItemData.offset(, espiro_origin_dictionary( "TIEMPO EN ANOS")))
              ActiveCell.offset(, espiro_destiny_dictionary("INTERPRETACION")) = charters(ItemData.offset(, espiro_origin_dictionary( "INTERPRETACION")))
              ActiveCell.offset(, espiro_destiny_dictionary("PESO")) = charters(ItemData.offset(, espiro_origin_dictionary( "PESO")))
              ActiveCell.offset(, espiro_destiny_dictionary("TALLA")) = charters(ItemData.offset(, espiro_origin_dictionary( "TALLA")))
              ActiveCell.offset(, espiro_destiny_dictionary("FVC PRED DIAG_")) = charters(ItemData.offset(, espiro_origin_dictionary( "FVC PRED DIAG_")))
              ActiveCell.offset(, espiro_destiny_dictionary("FVC %TEOR DIAG_")) = charters(ItemData.offset(, espiro_origin_dictionary( "FVC %TEOR DIAG_")))
              ActiveCell.offset(, espiro_destiny_dictionary("FEV1 PRED DIAG_")) = charters(ItemData.offset(, espiro_origin_dictionary( "FEV1 PRED DIAG_")))
              ActiveCell.offset(, espiro_destiny_dictionary("FEV1 %TEOR DIAG_")) = charters(ItemData.offset(, espiro_origin_dictionary( "FEV1 %TEOR DIAG_")))
              ActiveCell.offset(, espiro_destiny_dictionary("FEV1/FVC PRED DIAG_")) = charters(ItemData.offset(, espiro_origin_dictionary( "FEV1/FVC PRED DIAG_")))
              ActiveCell.offset(, espiro_destiny_dictionary("FEV1/FVC %TEOR DIAG_")) = charters(ItemData.offset(, espiro_origin_dictionary( "FEV1/FVC %TEOR DIAG_")))
              ActiveCell.offset(, espiro_destiny_dictionary("PEF PRED DIAG_")) = charters(ItemData.offset(, espiro_origin_dictionary( "PEF PRED DIAG_")))
              ActiveCell.offset(, espiro_destiny_dictionary("PEF %TEOR DIAG_")) = charters(ItemData.offset(, espiro_origin_dictionary( "PEF %TEOR DIAG_")))
              ActiveCell.offset(, espiro_destiny_dictionary("FEF 25-75 PRED DIAG_")) = charters(ItemData.offset(, espiro_origin_dictionary( "FEF 25-75 PRED DIAG_")))
              ActiveCell.offset(, espiro_destiny_dictionary("FEF 25-75 %TEOR DIAG_")) = charters(ItemData.offset(, espiro_origin_dictionary( "FEF 25-75 %TEOR DIAG_")))
              ActiveCell.offset(, espiro_destiny_dictionary("DIAG_ PPAL")) = charters(ItemData.offset(, espiro_origin_dictionary( "DIAG_ PPAL")))
              ActiveCell.offset(, espiro_destiny_dictionary("DIAG_ OBS")) = charters(ItemData.offset(, espiro_origin_dictionary( "DIAG_ OBS")))
              ActiveCell.offset(, espiro_destiny_dictionary("DIAG_ REL/1")) = charters(ItemData.offset(, espiro_origin_dictionary( "DIAG_ REL/1")))
              ActiveCell.offset(, espiro_destiny_dictionary("DIAG_ REL/2")) = charters(ItemData.offset(, espiro_origin_dictionary( "DIAG_ REL/2")))
              ActiveCell.offset(, espiro_destiny_dictionary("DIAG_ REL/3")) = charters(ItemData.offset(, espiro_origin_dictionary( "DIAG_ REL/3")))
              ActiveCell.offset(, espiro_destiny_dictionary("TIPO_INTERPRETACION")) = charters(ItemData.offset(, espiro_origin_dictionary( "TIPO_INTERPRETACION")))
              ActiveCell.offset(, espiro_destiny_dictionary("TIPO_GRADO")) = charters(ItemData.offset(, espiro_origin_dictionary( "TIPO_GRADO")))
              ActiveCell.offset(, espiro_destiny_dictionary("RESULTADO_ESPIROMETRIA")) = charters(ItemData.offset(, espiro_origin_dictionary( "RESULTADO_ESPIROMETRIA")))
              ActiveCell.offset(, espiro_destiny_dictionary("REC/GRALES DEJAR DE FUMAR")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "REC/GRALES DEJAR DE FUMAR")))
              ActiveCell.offset(, espiro_destiny_dictionary("REC/GRALES CONTINUAR CONTROLES EPS")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "REC/GRALES CONTINUAR CONTROLES EPS")))
              ActiveCell.offset(, espiro_destiny_dictionary("REC/GRALES BAJAR DE PESO")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "REC/GRALES BAJAR DE PESO")))
              ActiveCell.offset(, espiro_destiny_dictionary("REC/GRALES TOMAR RAYOS X TORAX")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "REC/GRALES TOMAR RAYOS X TORAX")))
              ActiveCell.offset(, espiro_destiny_dictionary("REC/GRALES REALIZAR EJERC_ 3X SEMANA")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "REC/GRALES REALIZAR EJERC_ 3X SEMANA")))
              ActiveCell.offset(, espiro_destiny_dictionary("REC/GRALES VALORAC_ EPS X NEUMOLOGIA")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "REC/GRALES VALORAC_ EPS X NEUMOLOGIA")))
              ActiveCell.offset(, espiro_destiny_dictionary("REC/LAB UTILIZAR EPR")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "REC/LAB UTILIZAR EPR")))
              ActiveCell.offset(, espiro_destiny_dictionary("REC/LAB INGRESAR SVE")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "REC/LAB INGRESAR SVE")))
              ActiveCell.offset(, espiro_destiny_dictionary("CONTROLES MENSUAL")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "CONTROLES MENSUAL")))
              ActiveCell.offset(, espiro_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "CONTROLES_BIMESTRALES")))
              ActiveCell.offset(, espiro_destiny_dictionary("CONTROLES TRIMESTRAL")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "CONTROLES TRIMESTRAL")))
              ActiveCell.offset(, espiro_destiny_dictionary("CONTROLES SEMESTRAL")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "CONTROLES SEMESTRAL")))
              ActiveCell.offset(, espiro_destiny_dictionary("CONTROLES ANUAL")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "CONTROLES ANUAL")))
              ActiveCell.offset(, espiro_destiny_dictionary("CONTROLES CONFIRMATORIA")) = charters_empty(ItemData.offset(, espiro_origin_dictionary( "CONTROLES CONFIRMATORIA")))
              ActiveCell.offset(, espiro_destiny_dictionary("TECNICA ACEPTABLE")) = charters(ItemData.offset(, espiro_origin_dictionary( "TECNICA ACEPTABLE")))
              ActiveCell.offset(, espiro_destiny_dictionary("ID_ESPIROMETRIA")) = ActiveCell.offset(-1, espiro_destiny_dictionary("ID_ESPIROMETRIA")) + 1
              ActiveCell.offset(1, 0).Select
              numbers = numbers + 1
              numbersGeneral = numbersGeneral + 1
              DoEvents
            Next ItemData

            Range("$A4").Select
            Call dataDuplicate
            Range("$A4", Range("$A4").End(xlDown)).Select
            Call formatter

            Set espiro_origin_value = Nothing
            Set espiro_destiny_header = Nothing
            Set espiro_origin_header = Nothing
            espiro_destiny_dictionary.RemoveAll
            espiro_origin_dictionary.RemoveAll

 espiroError:
            resume next
End Sub