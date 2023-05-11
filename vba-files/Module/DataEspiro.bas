Attribute VB_Name = "DataEspiro"
Option Explicit

Public Sub EspiroData()
  Dim espiro_destiny_dictionary As Scripting.Dictionary
  Dim espiro_origin_dictionary As Scripting.Dictionary
  Dim espiro_destiny_header As Object, espiro_origin_header As Object, espiro_origin_value As Object
  Dim ItemEspiroDestiny As Variant, ItemEspiroOrigin As Variant, ItemData As Variant

  Set espiro_origin = origin.Worksheets("ESPIRO") '' ESPIRO DEL LIBRO ORIGEN ''
  espiro_destiny.Select
  ActiveSheet.Range("A4").Select
  Set espiro_destiny_header = espiro_destiny.Range("A3", espiro_destiny.Range("A3").End(xlToRight))
  Set espiro_origin_header = espiro_origin.Range("A1", espiro_origin.Range("A1").End(xlToRight))
  Set espiro_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set espiro_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (espiro_origin.Range("A2") <> Empty And espiro_origin.Range("A3") <> Empty) Then
    Set espiro_origin_value = espiro_origin.Range("A2", espiro_origin.Range("A2").End(xlDown))
  ElseIf (espiro_origin.Range("A2") <> Empty And espiro_origin.Range("A3") = Empty) Then
    Set espiro_origin_value = espiro_origin.Range("A2")
  End If

  ''   En los diccionarios de "espiro_destiny_dictionary" y  "espiro_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemEspiroDestiny In espiro_destiny_header
    On Error GoTo espiroError
    espiro_destiny_dictionary.Add espiro_headers(ItemEspiroDestiny), (ItemEspiroDestiny.Column - 1)
  Next ItemEspiroDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemEspiroOrigin In espiro_origin_header
    On Error GoTo espiroError
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
      If formImports.ProgressBarGeneral.Width > (formImports.content_ProgressBarGeneral.Width / 2) Then
        formImports.porcentageGeneral.ForeColor = RGB(255, 255, 255)
      End If
      If formImports.ProgressBarGeneral.Width < (formImports.content_ProgressBarGeneral.Width / 2) Then
        formImports.porcentageGeneral.ForeColor = RGB(0, 0, 0)
      End If
      If formImports.ProgressBarOneforOne.Width > (formImports.content_ProgressBarOneforOne.Width / 2) Then
        formImports.porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
      End If
      If formImports.ProgressBarOneforOne.Width < (formImports.content_ProgressBarOneforOne.Width / 2) Then
        formImports.porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
      End If
      If (typeExams(charters(ItemData.Offset(, espiro_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        With ActiveCell
          .Offset(, espiro_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, espiro_origin_dictionary("NRO IDENFICACION")))
          .Offset(, espiro_destiny_dictionary("ALERGIAS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("ALERGIAS")))
          .Offset(, espiro_destiny_dictionary("ALERGIAS OBS")) = charters(ItemData.Offset(, espiro_origin_dictionary("ALERGIAS OBS")))
          .Offset(, espiro_destiny_dictionary("TUBERCULOSIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TUBERCULOSIS")))
          .Offset(, espiro_destiny_dictionary("TOS CRONICA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TOS CRONICA")))
          .Offset(, espiro_destiny_dictionary("GRIPAS FRECUENTES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("GRIPAS FRECUENTES")))
          .Offset(, espiro_destiny_dictionary("FARINGITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("FARINGITIS")))
          .Offset(, espiro_destiny_dictionary("FARINGOAMIGDALITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("FARINGOAMIGDALITIS")))
          .Offset(, espiro_destiny_dictionary("RINITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RINITIS")))
          .Offset(, espiro_destiny_dictionary("SINUSITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("SINUSITIS")))
          .Offset(, espiro_destiny_dictionary("CX TORAX")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CX TORAX")))
          .Offset(, espiro_destiny_dictionary("CX TORAX OBS")) = charters(ItemData.Offset(, espiro_origin_dictionary("CX TORAX OBS")))
          .Offset(, espiro_destiny_dictionary("ASMA BRONQUIAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("ASMA BRONQUIAL")))
          .Offset(, espiro_destiny_dictionary("BRONQUITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("BRONQUITIS")))
          .Offset(, espiro_destiny_dictionary("NEUMONIA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("NEUMONIA")))
          .Offset(, espiro_destiny_dictionary("TRAUMA COSTAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TRAUMA COSTAL")))
          .Offset(, espiro_destiny_dictionary("CANCER")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CANCER")))
          .Offset(, espiro_destiny_dictionary("CANCER OBS")) = charters(ItemData.Offset(, espiro_origin_dictionary("CANCER OBS")))
          .Offset(, espiro_destiny_dictionary("OTROS RESPIRATORIOS")) = charters(ItemData.Offset(, espiro_origin_dictionary("OTROS RESPIRATORIOS")))
          .Offset(, espiro_destiny_dictionary("RIESGO QUIMICO / POLVOS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / POLVOS")))
          .Offset(, espiro_destiny_dictionary("RIESGO QUIMICO / FIBRAS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / FIBRAS")))
          .Offset(, espiro_destiny_dictionary("RIESGO QUIMICO / LIQUIDOS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / LIQUIDOS")))
          .Offset(, espiro_destiny_dictionary("RIESGO QUIMICO /GASES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO /GASES")))
          .Offset(, espiro_destiny_dictionary("RIESGO QUIMICO / VAPORES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / VAPORES")))
          .Offset(, espiro_destiny_dictionary("RIESGO QUIMICO / HUMOS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / HUMOS")))
          .Offset(, espiro_destiny_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")))
          .Offset(, espiro_destiny_dictionary("OTROS RIESGOS QUIMICOS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("OTROS RIESGOS QUIMICOS")))
          .Offset(, espiro_destiny_dictionary("EPP ESPECIFICO / TAPABOCA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("EPP ESPECIFICO / TAPABOCA")))
          .Offset(, espiro_destiny_dictionary("EPP ESPECIFICO / RESPIRADOR")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("EPP ESPECIFICO / RESPIRADOR")))
          .Offset(, espiro_destiny_dictionary("ACT_ FISICA")) = typeActivity(charters(ItemData.Offset(, espiro_origin_dictionary("ACT_ FISICA"))))
          .Offset(, espiro_destiny_dictionary("FUMA")) = typeSmoke(charters(ItemData.Offset(, espiro_origin_dictionary("FUMA"))))
          .Offset(, espiro_destiny_dictionary("CIGARRILLOS DIA")) = charters(ItemData.Offset(, espiro_origin_dictionary("CIGARRILLOS DIA")))
          .Offset(, espiro_destiny_dictionary("FRECUENCIA")) = charters(ItemData.Offset(, espiro_origin_dictionary("FRECUENCIA")))
          .Offset(, espiro_destiny_dictionary("TIEMPO EN ANOS")) = charters(ItemData.Offset(, espiro_origin_dictionary("TIEMPO EN ANOS")))
          .Offset(, espiro_destiny_dictionary("INTERPRETACION")) = charters(ItemData.Offset(, espiro_origin_dictionary("INTERPRETACION")))
          .Offset(, espiro_destiny_dictionary("PESO")) = charters(ItemData.Offset(, espiro_origin_dictionary("PESO")))
          .Offset(, espiro_destiny_dictionary("TALLA")) = charters(ItemData.Offset(, espiro_origin_dictionary("TALLA")))
          .Offset(, espiro_destiny_dictionary("FVC PRED DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FVC PRED DIAG_")))
          .Offset(, espiro_destiny_dictionary("FVC %TEOR DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FVC %TEOR DIAG_")))
          .Offset(, espiro_destiny_dictionary("FEV1 PRED DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1 PRED DIAG_")))
          .Offset(, espiro_destiny_dictionary("FEV1 %TEOR DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1 %TEOR DIAG_")))
          .Offset(, espiro_destiny_dictionary("FEV1/FVC PRED DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1/FVC PRED DIAG_")))
          .Offset(, espiro_destiny_dictionary("FEV1/FVC %TEOR DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1/FVC %TEOR DIAG_")))
          .Offset(, espiro_destiny_dictionary("PEF PRED DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("PEF PRED DIAG_")))
          .Offset(, espiro_destiny_dictionary("PEF %TEOR DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("PEF %TEOR DIAG_")))
          .Offset(, espiro_destiny_dictionary("FEF 25-75 PRED DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEF 25-75 PRED DIAG_")))
          .Offset(, espiro_destiny_dictionary("FEF 25-75 %TEOR DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEF 25-75 %TEOR DIAG_")))
          .Offset(, espiro_destiny_dictionary("DIAG_ PPAL")) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ PPAL")))
          .Offset(, espiro_destiny_dictionary("DIAG_ OBS")) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ OBS")))
          .Offset(, espiro_destiny_dictionary("DIAG_ REL/1")) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/1")))
          .Offset(, espiro_destiny_dictionary("DIAG_ REL/2")) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/2")))
          .Offset(, espiro_destiny_dictionary("DIAG_ REL/3")) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/3")))
          .Offset(, espiro_destiny_dictionary("TIPO_INTERPRETACION")) = charters(ItemData.Offset(, espiro_origin_dictionary("TIPO_INTERPRETACION")))
          .Offset(, espiro_destiny_dictionary("TIPO_GRADO")) = charters(ItemData.Offset(, espiro_origin_dictionary("TIPO_GRADO")))
          .Offset(, espiro_destiny_dictionary("REC/GRALES DEJAR DE FUMAR")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES DEJAR DE FUMAR")))
          .Offset(, espiro_destiny_dictionary("REC/GRALES CONTINUAR CONTROLES EPS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES CONTINUAR CONTROLES EPS")))
          .Offset(, espiro_destiny_dictionary("REC/GRALES BAJAR DE PESO")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES BAJAR DE PESO")))
          .Offset(, espiro_destiny_dictionary("REC/GRALES TOMAR RAYOS X TORAX")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES TOMAR RAYOS X TORAX")))
          .Offset(, espiro_destiny_dictionary("REC/GRALES REALIZAR EJERC_ 3X SEMANA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES REALIZAR EJERC_ 3X SEMANA")))
          .Offset(, espiro_destiny_dictionary("REC/GRALES VALORAC_ EPS X NEUMOLOGIA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES VALORAC_ EPS X NEUMOLOGIA")))
          .Offset(, espiro_destiny_dictionary("REC/LAB UTILIZAR EPR")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/LAB UTILIZAR EPR")))
          .Offset(, espiro_destiny_dictionary("REC/LAB INGRESAR SVE")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/LAB INGRESAR SVE")))
          .Offset(, espiro_destiny_dictionary("CONTROLES MENSUAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES MENSUAL")))
          .Offset(, espiro_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES_BIMESTRALES")))
          .Offset(, espiro_destiny_dictionary("CONTROLES TRIMESTRAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES TRIMESTRAL")))
          .Offset(, espiro_destiny_dictionary("CONTROLES SEMESTRAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES SEMESTRAL")))
          .Offset(, espiro_destiny_dictionary("CONTROLES ANUAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES ANUAL")))
          .Offset(, espiro_destiny_dictionary("CONTROLES CONFIRMATORIA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES CONFIRMATORIA")))
          .Offset(, espiro_destiny_dictionary("TECNICA ACEPTABLE")) = charters(ItemData.Offset(, espiro_origin_dictionary("TECNICA ACEPTABLE")))
          If (.Row = 4) Then
            .Offset(, espiro_destiny_dictionary("ID_ESPIROMETRIA")) = Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$10").value)
          Else
            .Offset(, espiro_destiny_dictionary("ID_ESPIROMETRIA")) = .Offset(-1, espiro_destiny_dictionary("ID_ESPIROMETRIA")) + 1
          End If
          .Offset(1, 0).Select
        End With
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData

    Range("$A4").Select
    Call dataDuplicate
    Range("$A4", Range("$A4").End(xlDown)).Select
    Call formatter
    Range("$BN4:$BS4").Select
    Call greaterThanOne
    Range("$BN4:$BS4").Select
    Call iqualCero

    Set espiro_origin_value = Nothing
    Set espiro_destiny_header = Nothing
    Set espiro_origin_header = Nothing
    espiro_destiny_dictionary.RemoveAll
    espiro_origin_dictionary.RemoveAll

 espiroError:
    Resume Next
End Sub
