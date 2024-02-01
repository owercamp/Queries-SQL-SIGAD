Attribute VB_Name = "DataEspiro"
'namespace=vba-files\Module\informations
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
Dim aumentFromID As LongPtr
Public Sub EspiroData(ByVal name_sheet As String)
  Dim espiro_destiny_dictionary As Scripting.Dictionary
  Dim espiro_origin_dictionary As Scripting.Dictionary
  Dim espiro_destiny_header As Object, espiro_origin_header As Object, espiro_origin_value As Object
  Dim ItemEspiroDestiny As Object, ItemEspiroOrigin As Object, ItemData As Object, espiro_origin As Object

  Set espiro_origin = origin.Worksheets(name_sheet) '' ESPIRO DEL LIBRO ORIGEN ''
  espiro_destiny.Select
  espiro_destiny.Range("$A4").Select
  Set espiro_destiny_header = espiro_destiny.Range("$A3", espiro_destiny.Range("$A3").End(xlToRight))
  Set espiro_origin_header = espiro_origin.Range("$A1", espiro_origin.Range("$A1").End(xlToRight))
  Set espiro_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set espiro_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (espiro_origin.Range("$A2") <> Empty And espiro_origin.Range("$A3") <> Empty) Then
    Set espiro_origin_value = espiro_origin.Range("$A2", espiro_origin.Range("$A2").End(xlDown))
  ElseIf (espiro_origin.Range("$A2") <> Empty And espiro_origin.Range("$A3") = Empty) Then
    Set espiro_origin_value = espiro_origin.Range("$A2")
  End If

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  Dim value_data As String
  For Each ItemEspiroDestiny In espiro_destiny_header
    value_data = espiro_headers(ItemEspiroDestiny)
    If espiro_destiny_dictionary.Exists(value_data) = False And value_data <> Empty Then
      espiro_destiny_dictionary.Add value_data, (ItemEspiroDestiny.Column - 1)
    End If
  Next ItemEspiroDestiny
  
  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemEspiroOrigin In espiro_origin_header
    value_data = espiro_headers(ItemEspiroOrigin)
    If espiro_origin_dictionary.Exists(value_data) = False And value_data <> Empty Then
      espiro_origin_dictionary.Add value_data, (ItemEspiroOrigin.Column - 1)
    End If
  Next ItemEspiroOrigin

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$10").value
  counts = espiro_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  Dim type_exam As String
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

      type_exam = typeExams(Trim(ItemData.Offset(, espiro_origin_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        ActiveCell.Offset(, espiro_destiny_dictionary("NRO IDENFICACION")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("NRO IDENFICACION"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("ALERGIAS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("ALERGIAS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("ALERGIAS OBS")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("ALERGIAS OBS"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("TUBERCULOSIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TUBERCULOSIS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("TOS CRONICA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TOS CRONICA")))
        ActiveCell.Offset(, espiro_destiny_dictionary("GRIPAS FRECUENTES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("GRIPAS FRECUENTES")))
        ActiveCell.Offset(, espiro_destiny_dictionary("FARINGITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("FARINGITIS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("FARINGOAMIGDALITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("FARINGOAMIGDALITIS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("RINITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RINITIS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("SINUSITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("SINUSITIS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("CX TORAX")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CX TORAX")))
        ActiveCell.Offset(, espiro_destiny_dictionary("CX TORAX OBS")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("CX TORAX OBS"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("ASMA BRONQUIAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("ASMA BRONQUIAL")))
        ActiveCell.Offset(, espiro_destiny_dictionary("BRONQUITIS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("BRONQUITIS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("NEUMONIA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("NEUMONIA")))
        ActiveCell.Offset(, espiro_destiny_dictionary("TRAUMA COSTAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TRAUMA COSTAL")))
        ActiveCell.Offset(, espiro_destiny_dictionary("CANCER")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CANCER")))
        ActiveCell.Offset(, espiro_destiny_dictionary("CANCER OBS")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("CANCER OBS"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("OTROS RESPIRATORIOS")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("OTROS RESPIRATORIOS"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("RIESGO QUIMICO / POLVOS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / POLVOS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("RIESGO QUIMICO / FIBRAS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / FIBRAS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("RIESGO QUIMICO / LIQUIDOS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / LIQUIDOS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("RIESGO QUIMICO /GASES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO /GASES")))
        ActiveCell.Offset(, espiro_destiny_dictionary("RIESGO QUIMICO / VAPORES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / VAPORES")))
        ActiveCell.Offset(, espiro_destiny_dictionary("RIESGO QUIMICO / HUMOS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / HUMOS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")))
        ActiveCell.Offset(, espiro_destiny_dictionary("OTROS RIESGOS QUIMICOS")) = VBA.Trim$(ItemData.Offset(, espiro_origin_dictionary("OTROS RIESGOS QUIMICOS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("EPP ESPECIFICO / TAPABOCA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("EPP ESPECIFICO / TAPABOCA")))
        ActiveCell.Offset(, espiro_destiny_dictionary("EPP ESPECIFICO / RESPIRADOR")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("EPP ESPECIFICO / RESPIRADOR")))
        ActiveCell.Offset(, espiro_destiny_dictionary("ACT_ FISICA")) = typeActivity(Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("ACT_ FISICA")))))
        ActiveCell.Offset(, espiro_destiny_dictionary("FUMA")) = typeSmoke(Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("FUMA")))))
        ActiveCell.Offset(, espiro_destiny_dictionary("CIGARRILLOS DIA")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("CIGARRILLOS DIA"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("FRECUENCIA")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("FRECUENCIA"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("TIEMPO EN ANOS")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("TIEMPO EN ANOS"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("INTERPRETACION")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("INTERPRETACION"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("PESO")) = Trim(ItemData.Offset(, espiro_origin_dictionary("PESO")))
        ActiveCell.Offset(, espiro_destiny_dictionary("TALLA")) = Trim(ItemData.Offset(, espiro_origin_dictionary("TALLA")))
        ActiveCell.Offset(, espiro_destiny_dictionary("FVC PRED DIAG_")) = Trim(ItemData.Offset(, espiro_origin_dictionary("FVC PRED DIAG_")))
        ActiveCell.Offset(, espiro_destiny_dictionary("FVC %TEOR DIAG_")) = Trim(ItemData.Offset(, espiro_origin_dictionary("FVC %TEOR DIAG_")))
        ActiveCell.Offset(, espiro_destiny_dictionary("FEV1 PRED DIAG_")) = Trim(ItemData.Offset(, espiro_origin_dictionary("FEV1 PRED DIAG_")))
        ActiveCell.Offset(, espiro_destiny_dictionary("FEV1 %TEOR DIAG_")) = Trim(ItemData.Offset(, espiro_origin_dictionary("FEV1 %TEOR DIAG_")))
        ActiveCell.Offset(, espiro_destiny_dictionary("FEV1/FVC PRED DIAG_")) = Trim(ItemData.Offset(, espiro_origin_dictionary("FEV1/FVC PRED DIAG_")))
        ActiveCell.Offset(, espiro_destiny_dictionary("FEV1/FVC %TEOR DIAG_")) = Trim(ItemData.Offset(, espiro_origin_dictionary("FEV1/FVC %TEOR DIAG_")))
        ActiveCell.Offset(, espiro_destiny_dictionary("PEF PRED DIAG_")) = Trim(ItemData.Offset(, espiro_origin_dictionary("PEF PRED DIAG_")))
        ActiveCell.Offset(, espiro_destiny_dictionary("PEF %TEOR DIAG_")) = Trim(ItemData.Offset(, espiro_origin_dictionary("PEF %TEOR DIAG_")))
        ActiveCell.Offset(, espiro_destiny_dictionary("FEF 25-75 PRED DIAG_")) = Trim(ItemData.Offset(, espiro_origin_dictionary("FEF 25-75 PRED DIAG_")))
        ActiveCell.Offset(, espiro_destiny_dictionary("FEF 25-75 %TEOR DIAG_")) = Trim(ItemData.Offset(, espiro_origin_dictionary("FEF 25-75 %TEOR DIAG_")))
        ActiveCell.Offset(, espiro_destiny_dictionary("DIAG_ PPAL")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("DIAG_ PPAL"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("DIAG_ OBS")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("DIAG_ OBS"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("DIAG_ REL/1")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/1"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("DIAG_ REL/2")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/2"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("DIAG_ REL/3")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/3"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("TIPO_INTERPRETACION")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("TIPO_INTERPRETACION"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("TIPO_GRADO")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("TIPO_GRADO"))))
        ActiveCell.Offset(, espiro_destiny_dictionary("REC/GRALES DEJAR DE FUMAR")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES DEJAR DE FUMAR")))
        ActiveCell.Offset(, espiro_destiny_dictionary("REC/GRALES CONTINUAR CONTROLES EPS")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES CONTINUAR CONTROLES EPS")))
        ActiveCell.Offset(, espiro_destiny_dictionary("REC/GRALES BAJAR DE PESO")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES BAJAR DE PESO")))
        ActiveCell.Offset(, espiro_destiny_dictionary("REC/GRALES TOMAR RAYOS X TORAX")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES TOMAR RAYOS X TORAX")))
        ActiveCell.Offset(, espiro_destiny_dictionary("REC/GRALES REALIZAR EJERC_ 3X SEMANA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES REALIZAR EJERC_ 3X SEMANA")))
        ActiveCell.Offset(, espiro_destiny_dictionary("REC/GRALES VALORAC_ EPS X NEUMOLOGIA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES VALORAC_ EPS X NEUMOLOGIA")))
        ActiveCell.Offset(, espiro_destiny_dictionary("REC/LAB UTILIZAR EPR")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/LAB UTILIZAR EPR")))
        ActiveCell.Offset(, espiro_destiny_dictionary("REC/LAB INGRESAR SVE")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/LAB INGRESAR SVE")))
        ActiveCell.Offset(, espiro_destiny_dictionary("CONTROLES MENSUAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES MENSUAL")))
        ActiveCell.Offset(, espiro_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES_BIMESTRALES")))
        ActiveCell.Offset(, espiro_destiny_dictionary("CONTROLES TRIMESTRAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES TRIMESTRAL")))
        ActiveCell.Offset(, espiro_destiny_dictionary("CONTROLES SEMESTRAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES SEMESTRAL")))
        ActiveCell.Offset(, espiro_destiny_dictionary("CONTROLES ANUAL")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES ANUAL")))
        ActiveCell.Offset(, espiro_destiny_dictionary("CONTROLES CONFIRMATORIA")) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES CONFIRMATORIA")))
        ActiveCell.Offset(, espiro_destiny_dictionary("TECNICA ACEPTABLE")) = Trim(UCase(ItemData.Offset(, espiro_origin_dictionary("TECNICA ACEPTABLE"))))
        If (ActiveCell.Row <> 4) Then
          aumentFromID = aumentFromID + 1
        End If
        ActiveCell.Offset(, espiro_destiny_dictionary("ID_ESPIROMETRIA")) = aumentFromID
        ActiveCell.Offset(1, 0).Select
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents
      End If
    Next ItemData
  End With

  Call dataDuplicate(espiro_destiny.Range("tbl_espiro_info[[#Data],[NRO IDENFICACION]]"))
  Call formatter(espiro_destiny.Range("tbl_espiro_info[[#Data],[NRO IDENFICACION]]"))
  Call greaterThanOne(espiro_destiny.Range("tbl_espiro_info[[CONTROLES MENSUAL]:[CONTROLES CONFIRMATORIA]]"), "ESPIRO")
  Call iqualCero(espiro_destiny.Range("tbl_espiro_info[[CONTROLES MENSUAL]:[CONTROLES CONFIRMATORIA]]"), "ESPIRO")

  Set espiro_origin_value = Nothing
  Set espiro_destiny_header = Nothing
  Set espiro_origin_header = Nothing
  espiro_destiny_dictionary.RemoveAll
  espiro_origin_dictionary.RemoveAll

End Sub