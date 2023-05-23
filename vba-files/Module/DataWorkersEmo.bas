Attribute VB_Name = "DataWorkersEmo"
Option Explicit

' DataEmoWorkers - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'------------------------------------------------------------------------------------------------------------------
' Variables:
' - emo_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de destino.
' - emo_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de origen.
' - emo_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
' - emo_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
' - emo_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
' - ItemEmoDestiny: Una variable de objeto para almacenar los valores de la columna de la hoja de destino.
' - ItemEmoOrigin: Una variable de objeto para almacenar los valores de la columna de la hoja de origen.
' - ItemData: Una variable de objeto para almacenar los valores de la hoja de origen.
' ------------------------------------------------------------------------------------------------------------------
Public Sub DataEmoWorkers()

  Dim emo_destiny_dictionary As Scripting.Dictionary
  Dim emo_origin_dictionary As Scripting.Dictionary
  Dim emo_destiny_header As Object, emo_origin_header As Object, emo_origin_value As Object
  Dim ItemEmoDestiny As Variant, ItemEmoOrigin As Variant, ItemData As Variant

  Call deleteFormatConditions
  Set emo_origin = origin.Worksheets("EMO") '' EMO DEL LIBRO ORIGEN ''
  emo_destiny.Select
  ActiveSheet.Range("A5").Select
  Set emo_destiny_header = emo_destiny.Range("A4", emo_destiny.Range("A4").End(xlToRight))
  Set emo_origin_header = emo_origin.Range("A1", emo_origin.Range("A1").End(xlToRight))
  Set emo_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (emo_origin.Range("A2") <> Empty And emo_origin.Range("A3") <> Empty) Then
    Set emo_origin_value = emo_origin.Range("A2", emo_origin.Range("A2").End(xlDown))
  ElseIf (emo_origin.Range("A2") <> Empty And emo_origin.Range("A3") = Empty) Then
    Set emo_origin_value = emo_origin.Range("A2")
  End If

  ''   En los diccionarios de "emo_destiny_dictionary" y  "emo_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemEmoDestiny In emo_destiny_header
    On Error GoTo emoError
    emo_destiny_dictionary.Add emo_headers(ItemEmoDestiny), (ItemEmoDestiny.Column - 1)
  Next ItemEmoDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemEmoOrigin In emo_origin_header
    On Error GoTo emoError
    emo_origin_dictionary.Add emo_headers(ItemEmoOrigin), (ItemEmoOrigin.Column - 1)
  Next ItemEmoOrigin

  numbers = 1
  porcentaje = 0
  counts = emo_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  For Each ItemData In emo_origin_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & emo_destiny.Name
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
      If (typeExams(charters(ItemData.Offset(, emo_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        With ActiveCell
          .Offset(, emo_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, emo_origin_dictionary("NRO IDENFICACION")))
          .Offset(, emo_destiny_dictionary("RIESGO FISICO / RUIDO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RUIDO")))
          .Offset(, emo_destiny_dictionary("RIESGO FISICO / ILUMINACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / ILUMINACION")))
          .Offset(, emo_destiny_dictionary("RIESGO FISICO / VIBRACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / VIBRACION")))
          .Offset(, emo_destiny_dictionary("RIESGO FISICO / TEMP EXTREMAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / TEMP EXTREMAS")))
          .Offset(, emo_destiny_dictionary("RIESGO FISICO / PRES ATMOSFERICA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / PRES ATMOSFERICA")))
          .Offset(, emo_destiny_dictionary("RIESGO FISICO / RAD IONIZANTES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RAD IONIZANTES")))
          .Offset(, emo_destiny_dictionary("RIESGO FISICO / RAD NO IONIZANTES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RAD NO IONIZANTES")))
          .Offset(, emo_destiny_dictionary("RIESGO DE OTROS FACTORES FISICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO DE OTROS FACTORES FISICOS")))
          .Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / VIRUS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / VIRUS")))
          .Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / BACTERIAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / BACTERIAS")))
          .Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / HONGOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / HONGOS")))
          .Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / RICKETSIAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / RICKETSIAS")))
          .Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / PARASITOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / PARASITOS")))
          .Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / FLUIDOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / FLUIDOS")))
          .Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / PICADURAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / PICADURAS")))
          .Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / MORDEDURAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / MORDEDURAS")))
          .Offset(, emo_destiny_dictionary("OTROS RIESGOS BIOLOGICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS BIOLOGICOS")))
          .Offset(, emo_destiny_dictionary("RIESGO QUIMICO / POLVOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / POLVOS")))
          .Offset(, emo_destiny_dictionary("RIESGO QUIMICO / FIBRAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / FIBRAS")))
          .Offset(, emo_destiny_dictionary("RIESGO QUIMICO / LIQUIDOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / LIQUIDOS")))
          .Offset(, emo_destiny_dictionary("RIESGO QUIMICO /GASES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO /GASES")))
          .Offset(, emo_destiny_dictionary("RIESGO QUIMICO / VAPORES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / VAPORES")))
          .Offset(, emo_destiny_dictionary("RIESGO QUIMICO / HUMOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / HUMOS")))
          .Offset(, emo_destiny_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")))
          .Offset(, emo_destiny_dictionary("OTROS RIESGOS QUIMICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS QUIMICOS")))
          .Offset(, emo_destiny_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL")))
          .Offset(, emo_destiny_dictionary("RIESGO PSICO / CARACT DEL GRUPO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CARACT DEL GRUPO")))
          .Offset(, emo_destiny_dictionary("RIESGO PSICO / INTERFACES TAREA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / INTERFACES TAREA")))
          .Offset(, emo_destiny_dictionary("RIESGO PSICO / CARACT ORGANIZACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CARACT ORGANIZACION")))
          .Offset(, emo_destiny_dictionary("RIESGO PSICO / CONDICIONES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CONDICIONES")))
          .Offset(, emo_destiny_dictionary("RIESGO PSICO / JORNADA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / JORNADA")))
          .Offset(, emo_destiny_dictionary("OTROS PSICO LABORAL")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS PSICO LABORAL")))
          .Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_POSTURA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_POSTURA")))
          .Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_ESFUERZO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_ESFUERZO")))
          .Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO")))
          .Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA")))
          .Offset(, emo_destiny_dictionary("OTROS RIESGOS BIOMECANICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS BIOMECANICOS")))
          .Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS")))
          .Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS")))
          .Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO")))
          .Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO")))
          .Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO")))
          .Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS")))
          .Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS")))
          .Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS")))
          .Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD")))
          .Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / SISMO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / SISMO")))
          .Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / TERREMOTO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / TERREMOTO")))
          .Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / VENDAVAL")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / VENDAVAL")))
          .Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / INUNDACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / INUNDACION")))
          .Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / DERRUMBE")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / DERRUMBE")))
          .Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / PRECIPITACIONES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / PRECIPITACIONES")))
          .Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / OTROS NATURALES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / OTROS NATURALES")))
          .Offset(, emo_destiny_dictionary("FECHA ACCIDENTE")) = charters(ItemData.Offset(, emo_origin_dictionary("FECHA ACCIDENTE")))
          .Offset(, emo_destiny_dictionary("ACCIDENTE_PASO_EN_EMPRESA")) = charters(ItemData.Offset(, emo_origin_dictionary("ACCIDENTE_PASO_EN_EMPRESA")))
          .Offset(, emo_destiny_dictionary("TIPO ACCIDENTE")) = charters(ItemData.Offset(, emo_origin_dictionary("TIPO ACCIDENTE")))
          .Offset(, emo_destiny_dictionary("NATURALEZA LESION")) = charters(ItemData.Offset(, emo_origin_dictionary("NATURALEZA LESION")))
          .Offset(, emo_destiny_dictionary("PARTE AFECTADA")) = charters(ItemData.Offset(, emo_origin_dictionary("PARTE AFECTADA")))
          .Offset(, emo_destiny_dictionary("INCAPACIDAD")) = charters(ItemData.Offset(, emo_origin_dictionary("INCAPACIDAD")))
          .Offset(, emo_destiny_dictionary("SECUELAS")) = charters(ItemData.Offset(, emo_origin_dictionary("SECUELAS")))
          .Offset(, emo_destiny_dictionary("NOMBRE ENFERMEDAD")) = charters(ItemData.Offset(, emo_origin_dictionary("NOMBRE ENFERMEDAD")))
          .Offset(, emo_destiny_dictionary("ETAPA")) = charters(ItemData.Offset(, emo_origin_dictionary("ETAPA")))
          .Offset(, emo_destiny_dictionary("OBSERVACIONES DE ENFERMEDAD")) = charters(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES DE ENFERMEDAD")))
          .Offset(, emo_destiny_dictionary("ACT_ FISICA")) = typeActivity(charters(ItemData.Offset(, emo_origin_dictionary("ACT_ FISICA"))))
          .Offset(, emo_destiny_dictionary("FUMA")) = typeSmoke(charters(ItemData.Offset(, emo_origin_dictionary("FUMA"))))
          .Offset(, emo_destiny_dictionary("CONSUMO DE ALCOHOL")) = charters(ItemData.Offset(, emo_origin_dictionary("CONSUMO DE ALCOHOL")))
          .Offset(, emo_destiny_dictionary("PESO")) = charters(ItemData.Offset(, emo_origin_dictionary("PESO")))
          .Offset(, emo_destiny_dictionary("TALLA")) = charters(ItemData.Offset(, emo_origin_dictionary("TALLA")))
          .Offset(, emo_destiny_dictionary("TENSION ARTERIAL")) = charters(ItemData.Offset(, emo_origin_dictionary("TENSION ARTERIAL")))
          .Offset(, emo_destiny_dictionary("FREC_ CARDIACA")) = charters(ItemData.Offset(, emo_origin_dictionary("FREC_ CARDIACA")))
          .Offset(, emo_destiny_dictionary("FREC_ RESPIRATORIA")) = charters(ItemData.Offset(, emo_origin_dictionary("FREC_ RESPIRATORIA")))
          .Offset(, emo_destiny_dictionary("PERIMETRO ABDOMINAL")) = charters(ItemData.Offset(, emo_origin_dictionary("PERIMETRO ABDOMINAL")))
          .Offset(, emo_destiny_dictionary("LATERALIDAD")) = charters(ItemData.Offset(, emo_origin_dictionary("LATERALIDAD")))
          .Offset(, emo_destiny_dictionary("OBS DIAGS")) = charters(ReplaceNonAlphaNumeric(ItemData.Offset(, emo_origin_dictionary("OBS DIAGS"))))
          .Offset(, emo_destiny_dictionary("CONCEPTO DE EVALUACION")) = charters(validateConcepts(ReplaceNonAlphaNumeric(ItemData.Offset(, emo_origin_dictionary("CONCEPTO DE EVALUACION")))))
          .Offset(, emo_destiny_dictionary("OBSERVACIONES DEL CONCEPTO")) = charters(ReplaceNonAlphaNumeric(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES DEL CONCEPTO"))))
          .Offset(, emo_destiny_dictionary("RECOMENDACIONES ESPECIFICAS")) = charters(ReplaceNonAlphaNumeric(ItemData.Offset(, emo_origin_dictionary("RECOMENDACIONES ESPECIFICAS"))))
          .Offset(, emo_destiny_dictionary("REMISION EPS")) = "0"
          .Offset(, emo_destiny_dictionary("CONTROL PERIODICO OCUPACIONAL")) = "0"
          .Offset(, emo_destiny_dictionary("UTILIZACION EPP ACORDE AL CARGO")) = "0"
          .Offset(, emo_destiny_dictionary("REALIZACION DE PRUEBAS COMPLEMENTARIAS")) = "0"
          .Offset(, emo_destiny_dictionary("HABITOS NUTRICIONALES")) = "0"
          .Offset(, emo_destiny_dictionary("EJERCICIO REGULAR 3 VECES POR SEMANA")) = "0"
          .Offset(, emo_destiny_dictionary("DEJAR DE FUMAR")) = "0"
          .Offset(, emo_destiny_dictionary("REDUCIR CONSUMO ALCOHOL")) = "0"
          .Offset(, emo_destiny_dictionary("OBSERVACIONES")) = "0"
          .Offset(, emo_destiny_dictionary("OSTEOMUSCULAR")) = "0"
          .Offset(, emo_destiny_dictionary("VISUAL")) = "0"
          .Offset(, emo_destiny_dictionary("ALTURAS")) = "0"
          .Offset(, emo_destiny_dictionary("BIOLOGICO")) = "0"
          .Offset(, emo_destiny_dictionary("MANIPULACION DE ALIMENTOS")) = "0"
          .Offset(, emo_destiny_dictionary("QUIMICO")) = "0"
          .Offset(, emo_destiny_dictionary("CUIDADO DE LA VOZ")) = "0"
          .Offset(, emo_destiny_dictionary("TEMPERATURAS EXTREMAS")) = "0"
          .Offset(, emo_destiny_dictionary("ESPACIOS CONFINADOS")) = "0"
          .Offset(, emo_destiny_dictionary("PIEL")) = "0"
          .Offset(, emo_destiny_dictionary("RESPIRATORIA")) = "0"
          .Offset(, emo_destiny_dictionary("AUDITIVO")) = "0"
          If (.Row = 5) Then
            .Offset(, emo_destiny_dictionary("ID_EMO")) = Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$5").value)
          Else
            .Offset(, emo_destiny_dictionary("ID_EMO")) = .Offset(-1, emo_destiny_dictionary("ID_EMO")) + 1
          End If
          .Offset(1, 0).Select
        End With
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData

    Call thisText("$BH5")
    Call dataDuplicate("$EK5")
    Call dataDuplicate("$EL5")
    Call dataDuplicate("$A5")
    Call Risk("$EO5")
    Call riskPre_ingreso("$EO5")
    Call formatter("$A5ks")
    Range("$BC5").Select
    Call date_accident

    Set emo_origin_value = Nothing
    Set emo_destiny_header = Nothing
    Set emo_origin_header = Nothing
    emo_destiny_dictionary.RemoveAll
    emo_origin_dictionary.RemoveAll

 emoError:
    Resume Next
End Sub
