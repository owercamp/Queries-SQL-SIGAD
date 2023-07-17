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
Dim espiro_origin_dictionary As Scripting.Dictionary
Dim aumentFromID As LongPtr
Public Sub EspiroData()
  Dim tbl_espiro As Object, espiro_origin_header As Object, espiro_origin_value As Object
  Dim ItemEspiroOrigin As Variant, ItemData As Variant
  
  Set espiro_origin = origin.Worksheets("ESPIRO") '' ESPIRO DEL LIBRO ORIGEN ''
  espiro_destiny.Select
  Set tbl_espiro = ActiveSheet.ListObjects("tbl_espiro_info")
  Set espiro_origin_header = espiro_origin.range("A1", espiro_origin.range("A1").End(xlToRight))
  Set espiro_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (espiro_origin.range("A2") <> Empty And espiro_origin.range("A3") <> Empty) Then
    Set espiro_origin_value = espiro_origin.range("A2", espiro_origin.range("A2").End(xlDown))
  ElseIf (espiro_origin.range("A2") <> Empty And espiro_origin.range("A3") = Empty) Then
    Set espiro_origin_value = espiro_origin.range("A2")
  End If

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemEspiroOrigin In espiro_origin_header
    On Error Resume Next
    espiro_origin_dictionary.Add espiro_headers(ItemEspiroOrigin), (ItemEspiroOrigin.Column - 1)
    On Error GoTo 0
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
        Select Case numbers
          Case 1
            Call addNewRegister(tbl_espiro.ListRows(1), aumentFromID, ItemData)
          Case Else
            aumentFromID = aumentFromID + 1
            Call addNewRegister(tbl_espiro.ListRows.Add, aumentFromID, ItemData)
        End Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
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
  Set espiro_origin_header = Nothing
  espiro_origin_dictionary.RemoveAll

End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrement As LongPtr, ByVal ItemData As Variant)

  With table
    .Range(1) = charters(ItemData.Offset(, espiro_origin_dictionary("NRO IDENFICACION")))
    .Range(2) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("ALERGIAS")))
    .Range(3) = charters(ItemData.Offset(, espiro_origin_dictionary("ALERGIAS OBS")))
    .Range(4) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TUBERCULOSIS")))
    .Range(5) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TOS CRONICA")))
    .Range(6) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("GRIPAS FRECUENTES")))
    .Range(7) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("FARINGITIS")))
    .Range(8) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("FARINGOAMIGDALITIS")))
    .Range(9) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RINITIS")))
    .Range(10) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("SINUSITIS")))
    .Range(11) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CX TORAX")))
    .Range(12) = charters(ItemData.Offset(, espiro_origin_dictionary("CX TORAX OBS")))
    .Range(13) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("ASMA BRONQUIAL")))
    .Range(14) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("BRONQUITIS")))
    .Range(15) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("NEUMONIA")))
    .Range(16) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("TRAUMA COSTAL")))
    .Range(17) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CANCER")))
    .Range(18) = charters(ItemData.Offset(, espiro_origin_dictionary("CANCER OBS")))
    .Range(19) = charters(ItemData.Offset(, espiro_origin_dictionary("OTROS RESPIRATORIOS")))
    .Range(20) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / POLVOS")))
    .Range(21) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / FIBRAS")))
    .Range(22) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / LIQUIDOS")))
    .Range(23) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO /GASES")))
    .Range(24) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / VAPORES")))
    .Range(25) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO / HUMOS")))
    .Range(26) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")))
    .Range(27) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("OTROS RIESGOS QUIMICOS")))
    .Range(28) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("EPP ESPECIFICO / TAPABOCA")))
    .Range(29) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("EPP ESPECIFICO / RESPIRADOR")))
    .Range(30) = typeActivity(charters(ItemData.Offset(, espiro_origin_dictionary("ACT_ FISICA"))))
    .Range(31) = typeSmoke(charters(ItemData.Offset(, espiro_origin_dictionary("FUMA"))))
    .Range(32) = charters(ItemData.Offset(, espiro_origin_dictionary("CIGARRILLOS DIA")))
    .Range(33) = charters(ItemData.Offset(, espiro_origin_dictionary("FRECUENCIA")))
    .Range(34) = charters(ItemData.Offset(, espiro_origin_dictionary("TIEMPO EN ANOS")))
    .Range(35) = charters(ItemData.Offset(, espiro_origin_dictionary("INTERPRETACION")))
    .Range(36) = charters(ItemData.Offset(, espiro_origin_dictionary("PESO")))
    .Range(37) = charters(ItemData.Offset(, espiro_origin_dictionary("TALLA")))
    .Range(40) = charters(ItemData.Offset(, espiro_origin_dictionary("FVC PRED DIAG_")))
    .Range(41) = charters(ItemData.Offset(, espiro_origin_dictionary("FVC %TEOR DIAG_")))
    .Range(42) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1 PRED DIAG_")))
    .Range(43) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1 %TEOR DIAG_")))
    .Range(44) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1/FVC PRED DIAG_")))
    .Range(45) = charters(ItemData.Offset(, espiro_origin_dictionary("FEV1/FVC %TEOR DIAG_")))
    .Range(46) = charters(ItemData.Offset(, espiro_origin_dictionary("PEF PRED DIAG_")))
    .Range(47) = charters(ItemData.Offset(, espiro_origin_dictionary("PEF %TEOR DIAG_")))
    .Range(48) = charters(ItemData.Offset(, espiro_origin_dictionary("FEF 25-75 PRED DIAG_")))
    .Range(49) = charters(ItemData.Offset(, espiro_origin_dictionary("FEF 25-75 %TEOR DIAG_")))
    .Range(50) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ PPAL")))
    .Range(51) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ OBS")))
    .Range(52) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/1")))
    .Range(53) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/2")))
    .Range(54) = charters(ItemData.Offset(, espiro_origin_dictionary("DIAG_ REL/3")))
    .Range(55) = charters(ItemData.Offset(, espiro_origin_dictionary("TIPO_INTERPRETACION")))
    .Range(56) = charters(ItemData.Offset(, espiro_origin_dictionary("TIPO_GRADO")))
    .Range(58) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES DEJAR DE FUMAR")))
    .Range(59) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES CONTINUAR CONTROLES EPS")))
    .Range(60) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES BAJAR DE PESO")))
    .Range(61) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES TOMAR RAYOS X TORAX")))
    .Range(62) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES REALIZAR EJERC_ 3X SEMANA")))
    .Range(63) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/GRALES VALORAC_ EPS X NEUMOLOGIA")))
    .Range(64) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/LAB UTILIZAR EPR")))
    .Range(65) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("REC/LAB INGRESAR SVE")))
    .Range(66) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES MENSUAL")))
    .Range(67) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES_BIMESTRALES")))
    .Range(68) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES TRIMESTRAL")))
    .Range(69) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES SEMESTRAL")))
    .Range(70) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES ANUAL")))
    .Range(71) = charters_empty(ItemData.Offset(, espiro_origin_dictionary("CONTROLES CONFIRMATORIA")))
    .Range(72) = charters(ItemData.Offset(, espiro_origin_dictionary("TECNICA ACEPTABLE")))
    .Range(78) = autoIncrement
  End With

End Sub