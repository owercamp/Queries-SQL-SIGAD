Attribute VB_Name = "DataEspiro"
Option Explicit

'TODO: EspiroData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - espiro_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - espiro_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - espiro_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - espiro_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - espiro_origin_value: Una variable de objeto para almacenar el rango de los datos de diagnosticos de la hoja de origen.
'* - ItemData: Una variable de objeto para almacenar el rango de los datos de diagnosticos de la hoja de origen.
'* - ItemEspiroDestiny: Una variable de objeto para almacenar el rango de los datos de diagnosticos de la hoja de origen.
'* - ItemEspiroOrigin: Una variable de objeto para almacenar el rango de los datos de diagnosticos de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Public Sub EspiroData()
  Dim espiro_destiny_dictionary As Scripting.Dictionary
  Dim espiro_origin_dictionary As Scripting.Dictionary
  Dim espiro_destiny_header As Object, espiro_origin_header As Object, espiro_origin_value As Object
  Dim ItemEspiroDestiny As Variant, ItemEspiroOrigin As Variant, ItemData As Variant
  Dim currenCell As range, aumentFromRow As LongPtr, aumentFromID As LongPtr
  
  Set espiro_origin = origin.Worksheets("ESPIRO") '' ESPIRO DEL LIBRO ORIGEN ''
  espiro_destiny.Select
  ActiveSheet.range("A4").Select
  Set currenCell = ActiveCell
  Set espiro_destiny_header = espiro_destiny.range("A3", espiro_destiny.range("A3").End(xlToRight))
  Set espiro_origin_header = espiro_origin.range("A1", espiro_origin.range("A1").End(xlToRight))
  Set espiro_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set espiro_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (espiro_origin.range("A2") <> Empty And espiro_origin.range("A3") <> Empty) Then
    Set espiro_origin_value = espiro_origin.range("A2", espiro_origin.range("A2").End(xlDown))
  ElseIf (espiro_origin.range("A2") <> Empty And espiro_origin.range("A3") = Empty) Then
    Set espiro_origin_value = espiro_origin.range("A2")
  End If

  ''   En los diccionarios de "espiro_destiny_dictionary" y  "espiro_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemEspiroDestiny In espiro_destiny_header
    On Error Resume Next
    espiro_destiny_dictionary.Add espiro_headers(ItemEspiroDestiny), (ItemEspiroDestiny.Column - 1)
    On Error GoTo 0
  Next ItemEspiroDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemEspiroOrigin In espiro_origin_header
    On Error Resume Next
    espiro_origin_dictionary.Add espiro_headers(ItemEspiroOrigin), (ItemEspiroOrigin.Column - 1)
    On Error GoTo 0
  Next ItemEspiroOrigin

  numbers = 1
  porcentaje = 0
  aumentFromRow = 0
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$10").value
  counts = espiro_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each ItemData In espiro_origin_value
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & espiro_destiny.Name
      porcentaje = porcentaje + vals
      porcentajeGeneral = porcentajeGeneral + valsGeneral
      .ProgressBarOneforOne.Width = oneForOne
      .ProgressBarGeneral.Width = generalAll
      .porcentageGeneral.Caption = CStr(VBA.Round(porcentajeGeneral * 100, 1)) & "%"
      .porcentageOneoforOne.Caption = CStr(VBA.Round(porcentaje * 100, 1)) & "%"
      
      If .ProgressBarGeneral.Width > (.content_ProgressBarGeneral.Width / 2) Then
        .porcentageGeneral.ForeColor = RGB(255, 255, 255)
      ElseIf .ProgressBarGeneral.Width < (.content_ProgressBarGeneral.Width / 2) Then
        .porcentageGeneral.ForeColor = RGB(0, 0, 0)
      End If
      
      If .ProgressBarOneforOne.Width > (.content_ProgressBarOneforOne.Width / 2) Then
        .porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
      ElseIf .ProgressBarOneforOne.Width < (.content_ProgressBarOneforOne.Width / 2) Then
        .porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
      End If
      
      .Caption = CStr(nameCompany)

      If (typeExams(charters(ItemData.Offset(, espiro_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, espiro_origin_dictionary("NRO IDENFICACION")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("ALERGIAS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("ALERGIAS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("ALERGIAS OBS")) = charters(ItemData.Offset(, espiro_origin_dictionary("ALERGIAS OBS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("TUBERCULOSIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TUBERCULOSIS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("TOS CRONICA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TOS CRONICA")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("GRIPAS FRECUENTES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("GRIPAS FRECUENTES")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FARINGITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("FARINGITIS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FARINGOAMIGDALITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("FARINGOAMIGDALITIS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("RINITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RINITIS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("SINUSITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("SINUSITIS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("CX TORAX")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CX TORAX")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("CX TORAX OBS")) = charters(ItemData.Offset(, espiro_origin_dictionary("CX TORAX OBS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("ASMA BRONQUIAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("ASMA BRONQUIAL")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("BRONQUITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("BRONQUITIS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("NEUMONIA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("NEUMONIA")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("TRAUMA COSTAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TRAUMA COSTAL")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("CANCER")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CANCER")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("CANCER OBS")) = charters(ItemData.Offset(, espiro_origin_dictionary("CANCER OBS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("OTROS RESPIRATORIOS")) = charters(ItemData.Offset(, espiro_origin_dictionary("OTROS RESPIRATORIOS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("RIESGO QUIMICO / POLVOS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / POLVOS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("RIESGO QUIMICO / FIBRAS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / FIBRAS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("RIESGO QUIMICO / LIQUIDOS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / LIQUIDOS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("RIESGO QUIMICO /GASES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO /GASES")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("RIESGO QUIMICO / VAPORES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / VAPORES")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("RIESGO QUIMICO / HUMOS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / HUMOS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("OTROS RIESGOS QUIMICOS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("OTROS RIESGOS QUIMICOS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("EPP ESPECIFICO / TAPABOCA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("EPP ESPECIFICO / TAPABOCA")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("EPP ESPECIFICO / RESPIRADOR")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("EPP ESPECIFICO / RESPIRADOR")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("ACT_ FISICA")) = typeActivity(charters(ItemData.Offset(, espiro_origin_dictionary("ACT_ FISICA"))))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FUMA")) = typeSmoke(charters(ItemData.Offset(, espiro_origin_dictionary("FUMA"))))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("CIGARRILLOS DIA")) = charters(ItemData.Offset(, espiro_origin_dictionary("CIGARRILLOS DIA")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FRECUENCIA")) = charters(ItemData.Offset(, espiro_origin_dictionary("FRECUENCIA")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("TIEMPO EN ANOS")) = charters(ItemData.Offset(, espiro_origin_dictionary("TIEMPO EN ANOS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("INTERPRETACION")) = charters(ItemData.Offset(, espiro_origin_dictionary("INTERPRETACION")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("PESO")) = charters(ItemData.Offset(, espiro_origin_dictionary("PESO")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("TALLA")) = charters(ItemData.Offset(, espiro_origin_dictionary("TALLA")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FVC PRED DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FVC PRED DIAG_")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FVC %TEOR DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FVC %TEOR DIAG_")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FEV1 PRED DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1 PRED DIAG_")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FEV1 %TEOR DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1 %TEOR DIAG_")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FEV1/FVC PRED DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1/FVC PRED DIAG_")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FEV1/FVC %TEOR DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1/FVC %TEOR DIAG_")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("PEF PRED DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("PEF PRED DIAG_")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("PEF %TEOR DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("PEF %TEOR DIAG_")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FEF 25-75 PRED DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEF 25-75 PRED DIAG_")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("FEF 25-75 %TEOR DIAG_")) = charters(ItemData.Offset(, espiro_origin_dictionary("FEF 25-75 %TEOR DIAG_")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("DIAG_ PPAL")) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ PPAL")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("DIAG_ OBS")) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ OBS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("DIAG_ REL/1")) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/1")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("DIAG_ REL/2")) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/2")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("DIAG_ REL/3")) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/3")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("TIPO_INTERPRETACION")) = charters(ItemData.Offset(, espiro_origin_dictionary("TIPO_INTERPRETACION")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("TIPO_GRADO")) = charters(ItemData.Offset(, espiro_origin_dictionary("TIPO_GRADO")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("REC/GRALES DEJAR DE FUMAR")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES DEJAR DE FUMAR")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("REC/GRALES CONTINUAR CONTROLES EPS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES CONTINUAR CONTROLES EPS")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("REC/GRALES BAJAR DE PESO")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES BAJAR DE PESO")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("REC/GRALES TOMAR RAYOS X TORAX")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES TOMAR RAYOS X TORAX")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("REC/GRALES REALIZAR EJERC_ 3X SEMANA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES REALIZAR EJERC_ 3X SEMANA")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("REC/GRALES VALORAC_ EPS X NEUMOLOGIA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES VALORAC_ EPS X NEUMOLOGIA")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("REC/LAB UTILIZAR EPR")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/LAB UTILIZAR EPR")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("REC/LAB INGRESAR SVE")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/LAB INGRESAR SVE")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("CONTROLES MENSUAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES MENSUAL")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES_BIMESTRALES")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("CONTROLES TRIMESTRAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES TRIMESTRAL")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("CONTROLES SEMESTRAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES SEMESTRAL")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("CONTROLES ANUAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES ANUAL")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("CONTROLES CONFIRMATORIA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES CONFIRMATORIA")))
        currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("TECNICA ACEPTABLE")) = charters(ItemData.Offset(, espiro_origin_dictionary("TECNICA ACEPTABLE")))
        If (currenCell.Offset(aumentFromRow, 0).row = 4) Then
          currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("ID_ESPIROMETRIA")) = Trim(aumentFromID)
        Else
          aumentFromID = aumentFromID + 1
          currenCell.Offset(aumentFromRow, espiro_destiny_dictionary("ID_ESPIROMETRIA")) = Trim(aumentFromID)
        End If
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      aumentFromRow = aumentFromRow + 1
      DoEvents
    Next ItemData
  End With

  range("$A4").Select
  Call dataDuplicate
  range("$A4", range("$A4").End(xlDown)).Select
  Call formatter
  range("$BN4:$BS4").Select
  Call greaterThanOne
  range("$BN4:$BS4").Select
  Call iqualCero

  Set espiro_origin_value = Nothing
  Set espiro_destiny_header = Nothing
  Set espiro_origin_header = Nothing
  espiro_destiny_dictionary.RemoveAll
  espiro_origin_dictionary.RemoveAll

End Sub
