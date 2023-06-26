Attribute VB_Name = "DataWorkersEmo"
Option Explicit

'TODO: DataEmoWorkers - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - emo_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - emo_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - emo_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - emo_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - emo_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
'* - ItemEmoDestiny: Una variable de objeto para almacenar los valores de la columna de la hoja de destino.
'* - ItemEmoOrigin: Una variable de objeto para almacenar los valores de la columna de la hoja de origen.
'* - ItemData: Una variable de objeto para almacenar los valores de la hoja de origen.
'* ------------------------------------------------------------------------------------------------------------------
Public Sub DataEmoWorkers()

  Dim emo_destiny_dictionary As Scripting.Dictionary
  Dim emo_origin_dictionary As Scripting.Dictionary
  Dim emo_destiny_header As Object, emo_origin_header As Object, emo_origin_value As Object
  Dim ItemEmoDestiny As Variant, ItemEmoOrigin As Variant, ItemData As Variant
  Dim currenCell As range, aumentFromRow As LongPtr, aumentFromID As LongPtr
  
  Set emo_origin = origin.Worksheets("EMO") '' EMO DEL LIBRO ORIGEN ''
  emo_destiny.Select
  ActiveSheet.range("A5").Select
  Set currenCell = ActiveCell
  Set emo_destiny_header = emo_destiny.range("A4", emo_destiny.range("A4").End(xlToRight))
  Set emo_origin_header = emo_origin.range("A1", emo_origin.range("A1").End(xlToRight))
  Set emo_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (emo_origin.range("A2") <> Empty And emo_origin.range("A3") <> Empty) Then
    Set emo_origin_value = emo_origin.range("A2", emo_origin.range("A2").End(xlDown))
  ElseIf (emo_origin.range("A2") <> Empty And emo_origin.range("A3") = Empty) Then
    Set emo_origin_value = emo_origin.range("A2")
  End If

  ''   En los diccionarios de "emo_destiny_dictionary" y  "emo_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemEmoDestiny In emo_destiny_header
    On Error Resume Next
    emo_destiny_dictionary.Add emo_headers(ItemEmoDestiny), (ItemEmoDestiny.Column - 1)
    On Error GoTo 0
  Next ItemEmoDestiny

    '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemEmoOrigin In emo_origin_header
    On Error Resume Next
    emo_origin_dictionary.Add emo_headers(ItemEmoOrigin), (ItemEmoOrigin.Column - 1)
    On Error GoTo 0
  Next ItemEmoOrigin

  numbers = 1
  oneForOne = 0
  porcentaje = 0
  aumentFromRow = 0
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$5").value
  counts = emo_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each ItemData In emo_origin_value
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & emo_destiny.Name
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

      If (typeExams(charters(ItemData.Offset(, emo_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, emo_origin_dictionary("NRO IDENFICACION")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO FISICO / RUIDO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RUIDO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO FISICO / ILUMINACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / ILUMINACION")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO FISICO / VIBRACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / VIBRACION")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO FISICO / TEMP EXTREMAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / TEMP EXTREMAS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO FISICO / PRES ATMOSFERICA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / PRES ATMOSFERICA")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO FISICO / RAD IONIZANTES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RAD IONIZANTES")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO FISICO / RAD NO IONIZANTES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RAD NO IONIZANTES")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO DE OTROS FACTORES FISICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO DE OTROS FACTORES FISICOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO BIOLOGICO / VIRUS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / VIRUS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO BIOLOGICO / BACTERIAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / BACTERIAS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO BIOLOGICO / HONGOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / HONGOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO BIOLOGICO / RICKETSIAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / RICKETSIAS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO BIOLOGICO / PARASITOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / PARASITOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO BIOLOGICO / FLUIDOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / FLUIDOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO BIOLOGICO / PICADURAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / PICADURAS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO BIOLOGICO / MORDEDURAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / MORDEDURAS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("OTROS RIESGOS BIOLOGICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS BIOLOGICOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO QUIMICO / POLVOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / POLVOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO QUIMICO / FIBRAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / FIBRAS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO QUIMICO / LIQUIDOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / LIQUIDOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO QUIMICO /GASES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO /GASES")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO QUIMICO / VAPORES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / VAPORES")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO QUIMICO / HUMOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / HUMOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("OTROS RIESGOS QUIMICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS QUIMICOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO PSICO / CARACT DEL GRUPO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CARACT DEL GRUPO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO PSICO / INTERFACES TAREA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / INTERFACES TAREA")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO PSICO / CARACT ORGANIZACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CARACT ORGANIZACION")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO PSICO / CONDICIONES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CONDICIONES")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO PSICO / JORNADA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / JORNADA")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("OTROS PSICO LABORAL")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS PSICO LABORAL")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO_BIOMECANICO_POSTURA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_POSTURA")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO_BIOMECANICO_ESFUERZO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_ESFUERZO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("OTROS RIESGOS BIOMECANICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS BIOMECANICOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("FENOMENOS NATURALES / SISMO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / SISMO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("FENOMENOS NATURALES / TERREMOTO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / TERREMOTO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("FENOMENOS NATURALES / VENDAVAL")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / VENDAVAL")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("FENOMENOS NATURALES / INUNDACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / INUNDACION")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("FENOMENOS NATURALES / DERRUMBE")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / DERRUMBE")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("FENOMENOS NATURALES / PRECIPITACIONES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / PRECIPITACIONES")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("FENOMENOS NATURALES / OTROS NATURALES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / OTROS NATURALES")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("FECHA ACCIDENTE")) = charters(ItemData.Offset(, emo_origin_dictionary("FECHA ACCIDENTE")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("ACCIDENTE_PASO_EN_EMPRESA")) = charters(ItemData.Offset(, emo_origin_dictionary("ACCIDENTE_PASO_EN_EMPRESA")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("TIPO ACCIDENTE")) = charters(ItemData.Offset(, emo_origin_dictionary("TIPO ACCIDENTE")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("NATURALEZA LESION")) = charters(ItemData.Offset(, emo_origin_dictionary("NATURALEZA LESION")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("PARTE AFECTADA")) = charters(ItemData.Offset(, emo_origin_dictionary("PARTE AFECTADA")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("INCAPACIDAD")) = charters(ItemData.Offset(, emo_origin_dictionary("INCAPACIDAD")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("SECUELAS")) = charters(ItemData.Offset(, emo_origin_dictionary("SECUELAS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("NOMBRE ENFERMEDAD")) = charters(ItemData.Offset(, emo_origin_dictionary("NOMBRE ENFERMEDAD")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("ETAPA")) = charters(ItemData.Offset(, emo_origin_dictionary("ETAPA")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("OBSERVACIONES DE ENFERMEDAD")) = charters(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES DE ENFERMEDAD")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("ACT_ FISICA")) = typeActivity(charters(ItemData.Offset(, emo_origin_dictionary("ACT_ FISICA"))))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("FUMA")) = typeSmoke(charters(ItemData.Offset(, emo_origin_dictionary("FUMA"))))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONSUMO DE ALCOHOL")) = charters(ItemData.Offset(, emo_origin_dictionary("CONSUMO DE ALCOHOL")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("PESO")) = charters(ItemData.Offset(, emo_origin_dictionary("PESO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("TALLA")) = charters(ItemData.Offset(, emo_origin_dictionary("TALLA")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("TENSION ARTERIAL")) = charters(ItemData.Offset(, emo_origin_dictionary("TENSION ARTERIAL")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("FREC_ CARDIACA")) = charters(ItemData.Offset(, emo_origin_dictionary("FREC_ CARDIACA")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("FREC_ RESPIRATORIA")) = charters(ItemData.Offset(, emo_origin_dictionary("FREC_ RESPIRATORIA")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("PERIMETRO ABDOMINAL")) = charters(ItemData.Offset(, emo_origin_dictionary("PERIMETRO ABDOMINAL")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("LATERALIDAD")) = charters(ItemData.Offset(, emo_origin_dictionary("LATERALIDAD")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("OBS DIAGS")) = charters(ItemData.Offset(, emo_origin_dictionary("OBS DIAGS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONCEPTO DE EVALUACION")) = validateConcepts(charters(ItemData.Offset(, emo_origin_dictionary("CONCEPTO DE EVALUACION"))))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("OBSERVACIONES DEL CONCEPTO")) = charters(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES DEL CONCEPTO")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RECOMENDACIONES ESPECIFICAS")) = charters(ItemData.Offset(, emo_origin_dictionary("RECOMENDACIONES ESPECIFICAS")))
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("REMISION EPS")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CONTROL PERIODICO OCUPACIONAL")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("UTILIZACION EPP ACORDE AL CARGO")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("REALIZACION DE PRUEBAS COMPLEMENTARIAS")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("HABITOS NUTRICIONALES")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("EJERCICIO REGULAR 3 VECES POR SEMANA")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("DEJAR DE FUMAR")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("REDUCIR CONSUMO ALCOHOL")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("OBSERVACIONES")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("OSTEOMUSCULAR")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("VISUAL")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("ALTURAS")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("BIOLOGICO")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("MANIPULACION DE ALIMENTOS")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("QUIMICO")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("CUIDADO DE LA VOZ")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("TEMPERATURAS EXTREMAS")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("ESPACIOS CONFINADOS")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("PIEL")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("RESPIRATORIA")) = "0"
        currenCell.Offset(aumentFromRow, emo_destiny_dictionary("AUDITIVO")) = "0"
        If (currenCell.Offset(aumentFromRow, 0).row = 5) Then
          currenCell.Offset(aumentFromRow, emo_destiny_dictionary("ID_EMO")) = Trim(aumentFromID)
        Else
          aumentFromID = aumentFromID + 1
          currenCell.Offset(aumentFromRow, emo_destiny_dictionary("ID_EMO")) = Trim(aumentFromID)
        End If
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      aumentFromRow = aumentFromRow + 1
      DoEvents
    Next ItemData
  End With

  range("$BH5").Select
  Call thisText
  range("$EK5").Select
  Call dataDuplicate
  range("$EL5").Select
  Call dataDuplicate
  range("$A5").Select
  Call dataDuplicate
  range("$EO5").Select
  Call Risk
  Call riskPre_ingreso
  range("$A5", range("$A5").End(xlDown)).Select
  Call formatter

  Set emo_origin_value = Nothing
  Set emo_destiny_header = Nothing
  Set emo_origin_header = Nothing
  emo_destiny_dictionary.RemoveAll
  emo_origin_dictionary.RemoveAll

End Sub
