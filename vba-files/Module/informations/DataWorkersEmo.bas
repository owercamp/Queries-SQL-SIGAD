Attribute VB_Name = "DataWorkersEmo"
'namespace=vba-files\Module\informations
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
Dim aumentFromID As LongPtr
Public Sub DataEmoWorkers(ByVal name_sheet As String)
  Dim emo_destiny_dictionary As Scripting.Dictionary
  Dim emo_origin_dictionary As Scripting.Dictionary
  Dim emo_destiny_header As Object, emo_origin_header As Object, emo_origin_value As Object
  Dim ItemEmoDestiny As Object, ItemEmoOrigin As Object, ItemData As Object, emo_origin As Object

  Set emo_origin = origin.Worksheets(name_sheet) '' EMO DEL LIBRO ORIGEN ''
  emo_destiny.Select
  emo_destiny.Range("$A5").Select
  Set emo_destiny_header = emo_destiny.Range("$A4", emo_destiny.Range("$A4").End(xlToRight))
  Set emo_origin_header = emo_origin.Range("$A1", emo_origin.Range("$A1").End(xlToRight))
  Set emo_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (emo_origin.Range("$A2") <> Empty And emo_origin.Range("$A3") <> Empty) Then
    Set emo_origin_value = emo_origin.Range("$A2", emo_origin.Range("$A2").End(xlDown))
  ElseIf (emo_origin.Range("$A2") <> Empty And emo_origin.Range("$A3") = Empty) Then
    Set emo_origin_value = emo_origin.Range("$A2")
  End If

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  Dim value_data As String
  For Each ItemEmoDestiny In emo_destiny_header
    value_data = emo_headers(ItemEmoDestiny)
    If emo_destiny_dictionary.Exists(value_data) = False And value_data <> Empty Then
      emo_destiny_dictionary.Add value_data, (ItemEmoDestiny.Column - 1)
    End If
  Next ItemEmoDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemEmoOrigin In emo_origin_header
    value_data = emo_headers(ItemEmoOrigin)
    If emo_origin_dictionary.Exists(value_data) = False And value_data <> Empty Then
      emo_origin_dictionary.Add value_data, (ItemEmoOrigin.Column - 1)
    End If
  Next ItemEmoOrigin

  numbers = 1
  oneForOne = 0
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$5").value
  counts = emo_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  Dim type_exam As String
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
      Elseif .ProgressBarGeneral.Width < (.content_ProgressBarGeneral.Width / 2) Then
        .porcentageGeneral.ForeColor = RGB(0, 0, 0)
      End If

      If .ProgressBarOneforOne.Width > (.content_ProgressBarOneforOne.Width / 2) Then
        .porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
      Elseif .ProgressBarOneforOne.Width < (.content_ProgressBarOneforOne.Width / 2) Then
        .porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
      End If

      .Caption = CStr(nameCompany)

      type_exam = typeExams(Trim(ItemData.Offset(, emo_origin_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        ActiveCell.Offset(, emo_destiny_dictionary("NRO IDENFICACION")) = Trim(ItemData.Offset(, emo_origin_dictionary("NRO IDENFICACION")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO FISICO / RUIDO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RUIDO")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO FISICO / ILUMINACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / ILUMINACION")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO FISICO / VIBRACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / VIBRACION")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO FISICO / TEMP EXTREMAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / TEMP EXTREMAS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO FISICO / PRES ATMOSFERICA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / PRES ATMOSFERICA")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO FISICO / RAD IONIZANTES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RAD IONIZANTES")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO FISICO / RAD NO IONIZANTES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RAD NO IONIZANTES")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO DE OTROS FACTORES FISICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO DE OTROS FACTORES FISICOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / VIRUS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / VIRUS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / BACTERIAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / BACTERIAS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / HONGOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / HONGOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / RICKETSIAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / RICKETSIAS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / PARASITOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / PARASITOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / FLUIDOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / FLUIDOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / PICADURAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / PICADURAS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / MORDEDURAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / MORDEDURAS")))
        ActiveCell.Offset(, emo_destiny_dictionary("OTROS RIESGOS BIOLOGICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS BIOLOGICOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / POLVOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / POLVOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / FIBRAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / FIBRAS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / LIQUIDOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / LIQUIDOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO QUIMICO /GASES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO /GASES")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / VAPORES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / VAPORES")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / HUMOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / HUMOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")))
        ActiveCell.Offset(, emo_destiny_dictionary("OTROS RIESGOS QUIMICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS QUIMICOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO PSICO / CARACT DEL GRUPO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CARACT DEL GRUPO")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO PSICO / INTERFACES TAREA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / INTERFACES TAREA")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO PSICO / CARACT ORGANIZACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CARACT ORGANIZACION")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO PSICO / CONDICIONES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CONDICIONES")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO PSICO / JORNADA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / JORNADA")))
        ActiveCell.Offset(, emo_destiny_dictionary("OTROS PSICO LABORAL")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS PSICO LABORAL")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_POSTURA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_POSTURA")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_ESFUERZO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_ESFUERZO")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO")))
        ActiveCell.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA")))
        ActiveCell.Offset(, emo_destiny_dictionary("OTROS RIESGOS BIOMECANICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS BIOMECANICOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO")))
        ActiveCell.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO")))
        ActiveCell.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO")))
        ActiveCell.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS")))
        ActiveCell.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS")))
        ActiveCell.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD")))
        ActiveCell.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / SISMO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / SISMO")))
        ActiveCell.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / TERREMOTO")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / TERREMOTO")))
        ActiveCell.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / VENDAVAL")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / VENDAVAL")))
        ActiveCell.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / INUNDACION")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / INUNDACION")))
        ActiveCell.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / DERRUMBE")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / DERRUMBE")))
        ActiveCell.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / PRECIPITACIONES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / PRECIPITACIONES")))
        ActiveCell.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / OTROS NATURALES")) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / OTROS NATURALES")))
        ActiveCell.Offset(, emo_destiny_dictionary("FECHA ACCIDENTE")) = Trim(ItemData.Offset(, emo_origin_dictionary("FECHA ACCIDENTE")))
        ActiveCell.Offset(, emo_destiny_dictionary("ACCIDENTE_PASO_EN_EMPRESA")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("ACCIDENTE_PASO_EN_EMPRESA"))))
        ActiveCell.Offset(, emo_destiny_dictionary("TIPO ACCIDENTE")) = Trim(ItemData.Offset(, emo_origin_dictionary("TIPO ACCIDENTE")))
        ActiveCell.Offset(, emo_destiny_dictionary("NATURALEZA LESION")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("NATURALEZA LESION"))))
        ActiveCell.Offset(, emo_destiny_dictionary("PARTE AFECTADA")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("PARTE AFECTADA"))))
        ActiveCell.Offset(, emo_destiny_dictionary("INCAPACIDAD")) = Trim(ItemData.Offset(, emo_origin_dictionary("INCAPACIDAD")))
        ActiveCell.Offset(, emo_destiny_dictionary("SECUELAS")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("SECUELAS"))))
        ActiveCell.Offset(, emo_destiny_dictionary("NOMBRE ENFERMEDAD")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("NOMBRE ENFERMEDAD"))))
        ActiveCell.Offset(, emo_destiny_dictionary("ETAPA")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("ETAPA"))))
        ActiveCell.Offset(, emo_destiny_dictionary("OBSERVACIONES DE ENFERMEDAD")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES DE ENFERMEDAD"))))
        ActiveCell.Offset(, emo_destiny_dictionary("ACT_ FISICA")) = typeActivity(Trim(UCase(ItemData.Offset(, emo_origin_dictionary("ACT_ FISICA")))))
        ActiveCell.Offset(, emo_destiny_dictionary("FUMA")) = typeSmoke(Trim(ItemData.Offset(, emo_origin_dictionary("FUMA"))))
        ActiveCell.Offset(, emo_destiny_dictionary("CONSUMO DE ALCOHOL")) = Trim(ItemData.Offset(, emo_origin_dictionary("CONSUMO DE ALCOHOL")))
        ActiveCell.Offset(, emo_destiny_dictionary("PESO")) = Trim(ItemData.Offset(, emo_origin_dictionary("PESO")))
        ActiveCell.Offset(, emo_destiny_dictionary("TALLA")) = Trim(ItemData.Offset(, emo_origin_dictionary("TALLA")))
        ActiveCell.Offset(, emo_destiny_dictionary("TENSION ARTERIAL")) = Trim(ItemData.Offset(, emo_origin_dictionary("TENSION ARTERIAL")))
        ActiveCell.Offset(, emo_destiny_dictionary("FREC_ CARDIACA")) = Trim(ItemData.Offset(, emo_origin_dictionary("FREC_ CARDIACA")))
        ActiveCell.Offset(, emo_destiny_dictionary("FREC_ RESPIRATORIA")) = Trim(ItemData.Offset(, emo_origin_dictionary("FREC_ RESPIRATORIA")))
        ActiveCell.Offset(, emo_destiny_dictionary("PERIMETRO ABDOMINAL")) = Trim(ItemData.Offset(, emo_origin_dictionary("PERIMETRO ABDOMINAL")))
        ActiveCell.Offset(, emo_destiny_dictionary("LATERALIDAD")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("LATERALIDAD"))))
        ActiveCell.Offset(, emo_destiny_dictionary("OBS DIAGS")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("OBS DIAGS"))))
        ActiveCell.Offset(, emo_destiny_dictionary("CONCEPTO DE EVALUACION")) = validateConcepts(Trim(UCase(ItemData.Offset(, emo_origin_dictionary("CONCEPTO DE EVALUACION")))))
        ActiveCell.Offset(, emo_destiny_dictionary("OBSERVACIONES DEL CONCEPTO")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES DEL CONCEPTO"))))
        ActiveCell.Offset(, emo_destiny_dictionary("RECOMENDACIONES ESPECIFICAS")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("RECOMENDACIONES ESPECIFICAS"))))
        ActiveCell.Offset(, emo_destiny_dictionary("REMISION EPS")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("CONTROL PERIODICO OCUPACIONAL")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("UTILIZACION EPP ACORDE AL CARGO")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("REALIZACION DE PRUEBAS COMPLEMENTARIAS")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("HABITOS NUTRICIONALES")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("EJERCICIO REGULAR 3 VECES POR SEMANA")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("DEJAR DE FUMAR")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("REDUCIR CONSUMO ALCOHOL")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("OBSERVACIONES")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("OSTEOMUSCULAR")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("VISUAL")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("ALTURAS")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("BIOLOGICO")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("MANIPULACION DE ALIMENTOS")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("QUIMICO")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("CUIDADO DE LA VOZ")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("TEMPERATURAS EXTREMAS")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("ESPACIOS CONFINADOS")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("PIEL")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("RESPIRATORIA")) = "0"
        ActiveCell.Offset(, emo_destiny_dictionary("AUDITIVO")) = "0"
        If (Activecell.Row <> 5) Then
          aumentFromID = aumentFromID + 1
        End If
        ActiveCell.Offset(, emo_destiny_dictionary("ID_EMO")) = aumentFromID
        ActiveCell.Offset(1, 0).Select
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents
      End If
    Next ItemData
  End With

  Call thisText(emo_destiny.Range("tbl_emo[[#Data],[INCAPACIDAD]]"))
  Call dataDuplicate(emo_destiny.Range("tbl_emo[[#Data],[orden_lista_trabajadoresid]]"))
  Call dataDuplicate(emo_destiny.Range("tbl_emo[[#Data],[id_emo]]"))
  Call dataDuplicate(emo_destiny.Range("tbl_emo[[#Data],[NRO IDENTIFICACION]]"))
  Call Risk(emo_destiny.Range("tbl_emo[[#Data],[SCRIPT ics_emo_riesgos]]"))
  Call riskPre_ingreso(emo_destiny.Range("tbl_emo[[#Data],[SCRIPT ics_emo_riesgos]]"))
  Call formatter(emo_destiny.Range("tbl_emo[[#Data],[NRO IDENTIFICACION]]"))

  Set emo_origin_value = Nothing
  Set emo_destiny_header = Nothing
  Set emo_origin_header = Nothing
  emo_destiny_dictionary.RemoveAll
  emo_origin_dictionary.RemoveAll

End Sub