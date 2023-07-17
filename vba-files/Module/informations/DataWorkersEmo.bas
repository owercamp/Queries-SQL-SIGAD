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
Dim emo_origin_dictionary As Scripting.Dictionary
Dim aumentFromID As LongPtr
Public Sub DataEmoWorkers()
  Dim tbl_emo As Object, emo_origin_header As Object, emo_origin_value As Object
  Dim ItemEmoOrigin As Variant, ItemData As Variant

  Set emo_origin = origin.Worksheets("EMO") '' EMO DEL LIBRO ORIGEN ''
  emo_destiny.Select
  Set tbl_emo = ActiveSheet.ListObjects("tbl_emo")
  Set emo_origin_header = emo_origin.range("A1", emo_origin.range("A1").End(xlToRight))
  Set emo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (emo_origin.range("A2") <> Empty And emo_origin.range("A3") <> Empty) Then
    Set emo_origin_value = emo_origin.range("A2", emo_origin.range("A2").End(xlDown))
  Elseif (emo_origin.range("A2") <> Empty And emo_origin.range("A3") = Empty) Then
    Set emo_origin_value = emo_origin.range("A2")
  End If

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemEmoOrigin In emo_origin_header
    On Error Resume Next
    emo_origin_dictionary.Add emo_headers(ItemEmoOrigin), (ItemEmoOrigin.Column - 1)
    On Error Goto 0
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

      If (typeExams(charters(ItemData.Offset(, emo_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        Select Case numbers
          Case 1
            Call addNewRegister(tbl_emo.ListRows(1), aumentFromID, ItemData)
          Case Else
            aumentFromID = aumentFromID + 1
            Call addNewRegister(tbl_emo.ListRows.Add, aumentFromID, ItemData)
        End Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
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
  Set emo_origin_header = Nothing
  emo_origin_dictionary.RemoveAll

End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrement As LongPtr, ByVal ItemData As Variant)

  With table
    .Range(1) = charters(ItemData.Offset(, emo_origin_dictionary("NRO IDENFICACION")))
    .Range(2) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RUIDO")))
    .Range(3) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / ILUMINACION")))
    .Range(4) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / VIBRACION")))
    .Range(5) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / TEMP EXTREMAS")))
    .Range(6) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / PRES ATMOSFERICA")))
    .Range(7) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RAD IONIZANTES")))
    .Range(8) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RAD NO IONIZANTES")))
    .Range(9) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO DE OTROS FACTORES FISICOS")))
    .Range(10) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / VIRUS")))
    .Range(11) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / BACTERIAS")))
    .Range(12) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / HONGOS")))
    .Range(13) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / RICKETSIAS")))
    .Range(14) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / PARASITOS")))
    .Range(15) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / FLUIDOS")))
    .Range(16) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / PICADURAS")))
    .Range(17) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / MORDEDURAS")))
    .Range(18) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS BIOLOGICOS")))
    .Range(19) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / POLVOS")))
    .Range(20) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / FIBRAS")))
    .Range(21) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / LIQUIDOS")))
    .Range(22) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO /GASES")))
    .Range(23) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / VAPORES")))
    .Range(24) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / HUMOS")))
    .Range(25) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")))
    .Range(26) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS QUIMICOS")))
    .Range(27) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL")))
    .Range(28) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CARACT DEL GRUPO")))
    .Range(29) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / INTERFACES TAREA")))
    .Range(30) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CARACT ORGANIZACION")))
    .Range(31) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CONDICIONES")))
    .Range(32) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / JORNADA")))
    .Range(33) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS PSICO LABORAL")))
    .Range(34) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_POSTURA")))
    .Range(35) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_ESFUERZO")))
    .Range(36) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO")))
    .Range(37) = charters_empty(ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA")))
    .Range(38) = charters_empty(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS BIOMECANICOS")))
    .Range(39) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS")))
    .Range(40) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS")))
    .Range(41) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO")))
    .Range(42) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO")))
    .Range(43) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO")))
    .Range(44) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS")))
    .Range(45) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS")))
    .Range(46) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS")))
    .Range(47) = charters_empty(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD")))
    .Range(48) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / SISMO")))
    .Range(49) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / TERREMOTO")))
    .Range(50) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / VENDAVAL")))
    .Range(51) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / INUNDACION")))
    .Range(52) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / DERRUMBE")))
    .Range(53) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / PRECIPITACIONES")))
    .Range(54) = charters_empty(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / OTROS NATURALES")))
    .Range(55) = charters(ItemData.Offset(, emo_origin_dictionary("FECHA ACCIDENTE")))
    .Range(56) = charters(ItemData.Offset(, emo_origin_dictionary("ACCIDENTE_PASO_EN_EMPRESA")))
    .Range(57) = charters(ItemData.Offset(, emo_origin_dictionary("TIPO ACCIDENTE")))
    .Range(58) = charters(ItemData.Offset(, emo_origin_dictionary("NATURALEZA LESION")))
    .Range(59) = charters(ItemData.Offset(, emo_origin_dictionary("PARTE AFECTADA")))
    .Range(60) = charters(ItemData.Offset(, emo_origin_dictionary("INCAPACIDAD")))
    .Range(61) = charters(ItemData.Offset(, emo_origin_dictionary("SECUELAS")))
    .Range(62) = charters(ItemData.Offset(, emo_origin_dictionary("NOMBRE ENFERMEDAD")))
    .Range(63) = charters(ItemData.Offset(, emo_origin_dictionary("ETAPA")))
    .Range(64) = charters(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES DE ENFERMEDAD")))
    .Range(65) = typeActivity(charters(ItemData.Offset(, emo_origin_dictionary("ACT_ FISICA"))))
    .Range(66) = typeSmoke(charters(ItemData.Offset(, emo_origin_dictionary("FUMA"))))
    .Range(67) = charters(ItemData.Offset(, emo_origin_dictionary("CONSUMO DE ALCOHOL")))
    .Range(68) = charters(ItemData.Offset(, emo_origin_dictionary("PESO")))
    .Range(69) = charters(ItemData.Offset(, emo_origin_dictionary("TALLA")))
    .Range(72) = charters(ItemData.Offset(, emo_origin_dictionary("TENSION ARTERIAL")))
    .Range(73) = charters(ItemData.Offset(, emo_origin_dictionary("FREC_ CARDIACA")))
    .Range(74) = charters(ItemData.Offset(, emo_origin_dictionary("FREC_ RESPIRATORIA")))
    .Range(75) = charters(ItemData.Offset(, emo_origin_dictionary("PERIMETRO ABDOMINAL")))
    .Range(76) = charters(ItemData.Offset(, emo_origin_dictionary("LATERALIDAD")))
    .Range(97) = charters(ItemData.Offset(, emo_origin_dictionary("OBS DIAGS")))
    .Range(98) = validateConcepts(charters(ItemData.Offset(, emo_origin_dictionary("CONCEPTO DE EVALUACION"))))
    .Range(99) = charters(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES DEL CONCEPTO")))
    .Range(133) = charters(ItemData.Offset(, emo_origin_dictionary("RECOMENDACIONES ESPECIFICAS")))
    .Range(112) = "0"
    .Range(113) = "0"
    .Range(114) = "0"
    .Range(115) = "0"
    .Range(116) = "0"
    .Range(117) = "0"
    .Range(118) = "0"
    .Range(119) = "0"
    .Range(120) = "0"
    .Range(121) = "0"
    .Range(122) = "0"
    .Range(123) = "0"
    .Range(124) = "0"
    .Range(125) = "0"
    .Range(126) = "0"
    .Range(127) = "0"
    .Range(128) = "0"
    .Range(129) = "0"
    .Range(130) = "0"
    .Range(131) = "0"
    .Range(132) = "0"
    .Range(142) = Trim(aumentFromID)
    DoEvents
  End With

End Sub