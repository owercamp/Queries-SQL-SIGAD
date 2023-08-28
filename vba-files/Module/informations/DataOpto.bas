Attribute VB_Name = "DataOpto"
'namespace=vba-files\Module\informations
Option Explicit

'TODO: OptoData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - opto_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - opto_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - opto_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - opto_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - opto_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Dim opto_origin_dictionary As Scripting.Dictionary
Dim aumentFromIDOpto As LongPtr, aumentFromIDDiagnostic As LongPtr
Public Sub OptoData()
  Dim tbl_opto As Object, opto_origin_header As Object, opto_origin_value As Object
  Dim ItemOptoOrigin As Variant, ItemData As Variant
  
  Set opto_origin = origin.Worksheets("OPTO") '' OPTO DEL LIBRO ORIGEN ''
  opto_destiny.Select
  Set tbl_opto = ActiveSheet.ListObjects("tbl_opto")
  Set opto_origin_header = opto_origin.range("A1", opto_origin.range("A1").End(xlToRight))
  Set opto_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (opto_origin.range("A2") <> Empty And opto_origin.range("A3") <> Empty) Then
    Set opto_origin_value = opto_origin.range("A2", opto_origin.range("A2").End(xlDown))
  ElseIf (opto_origin.range("A2") <> Empty And opto_origin.range("A3") = Empty) Then
    Set opto_origin_value = opto_origin.range("A2")
  End If

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemOptoOrigin In opto_origin_header
    On Error Resume Next
    opto_origin_dictionary.Add opto_headers(ItemOptoOrigin), (ItemOptoOrigin.Column - 1)
    On Error GoTo 0
  Next ItemOptoOrigin

  numbers = 1
  porcentaje = 0
  
  aumentFromIDOpto = destiny.Worksheets("RUTAS").range("$F$7").value
  aumentFromIDDiagnostic = destiny.Worksheets("RUTAS").range("$F$8").value
  counts = opto_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each ItemData In opto_origin_value
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & opto_destiny.Name
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
      
      If (typeExams(charters(ItemData.Offset(, opto_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        Select Case numbers
          Case 1
            Call addNewRegister(tbl_opto.ListRows(1), aumentFromIDOpto, aumentFromIDDiagnostic, ItemData)
          Case Else
            aumentFromIDOpto = aumentFromIDOpto + 1
            aumentFromIDDiagnostic = aumentFromIDDiagnostic + 1
            Call addNewRegister(tbl_opto.ListRows.Add, aumentFromIDOpto, aumentFromIDDiagnostic, ItemData)
        End Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData
  End With

  range("$A4").Select
  Call dataDuplicate
  range("$BD4:$BI4").Select
  Call greaterThanOne
  range("$BD4:$BI4").Select
  Call iqualCero
  range("$BK4").Select
  Call dataDuplicate
  range("$BL4").Select
  Call dataDuplicate
  range("$BM4").Select
  Call dataDuplicate
  range("$A4", range("$A4").End(xlDown)).Select
  Call formatter

  Set opto_origin_value = Nothing
  Set opto_origin_header = Nothing
  opto_origin_dictionary.RemoveAll

End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrementOpto As LongPtr, ByVal autoIncrementDiagnostic As LongPtr, ByVal ItemData As Variant)

  With table
    .Range(1) = charters(ItemData.Offset(, opto_origin_dictionary("IDENTIFICACION")))
    .Range(2) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")))
    .Range(3) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL USUARIO COMPUTADOR")))
    .Range(4) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")))
    .Range(5) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL CAMBIOS TEMPREATURA")))
    .Range(6) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")))
    .Range(7) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")))
    .Range(8) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS FOTOFOBIA")))
    .Range(9) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS OJO ROJO")))
    .Range(10) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS LAGRIMEO")))
    .Range(11) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS VISION BORROSA")))
    .Range(12) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS ARDOR")))
    .Range(13) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS VISION DOBLE")))
    .Range(14) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS CANSANCIO")))
    .Range(15) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS MALA VISION CERCANA")))
    .Range(16) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS DOLOR")))
    .Range(17) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS MALA VISON LEJANA")))
    .Range(18) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS SECRECION")))
    .Range(19) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS CEFALEA")))
    .Range(20) = charters(ItemData.Offset(, opto_origin_dictionary("OTROS SINTOMAS")))
    .Range(21) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - PARPADOS")))
    .Range(22) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - PARPADOS OBS")))
    .Range(23) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - CONJUNTIVAS")))
    .Range(24) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - OBS CONJUNTIVAS")))
    .Range(25) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - ESCLERAS")))
    .Range(26) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - OBS ESCLERAS")))
    .Range(27) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - PUPILAS")))
    .Range(28) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - PUPILAS OBS")))
    .Range(29) = charters(ItemData.Offset(, opto_origin_dictionary("MOT/OCUL COVERT TEST LEJOS")))
    .Range(30) = charters(ItemData.Offset(, opto_origin_dictionary("MOT/OCUL COVERT TEST CERCA")))
    .Range(31) = charters(ItemData.Offset(, opto_origin_dictionary("ESTADO DE CORRECCION")))
    .Range(32) = charters(ItemData.Offset(, opto_origin_dictionary("PATOLOGIA OCULAR")))
    .Range(33) = charters(ItemData.Offset(, opto_origin_dictionary("DIAG PPAL")))
    .Range(35) = charters(ItemData.Offset(, opto_origin_dictionary("DIAG OBS")))
    .Range(36) = charters(ItemData.Offset(, opto_origin_dictionary("DIAG REL/1")))
    .Range(37) = charters(ItemData.Offset(, opto_origin_dictionary("DIAG REL/2")))
    .Range(38) = charters(ItemData.Offset(, opto_origin_dictionary("DIAG REL/3")))
    .Range(39) = charters(ItemData.Offset(, opto_origin_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")))
    .Range(40) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO AR VIDEO TRMINAL")))
    .Range(41) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO DE LENTES DE PROTECCION SOLAR")))
    .Range(42) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO EPP VISUAL")))
    .Range(43) = charters(ItemData.Offset(, opto_origin_dictionary("REC PAUSAS ACTIVAS")))
    .Range(44) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO RX VISION PROXIMA")))
    .Range(45) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO RX DESCANSO")))
    .Range(46) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO PERMANENTE RX OPTICA")))
    .Range(47) = charters(ItemData.Offset(, opto_origin_dictionary("REC PYP")))
    .Range(48) = charters(ItemData.Offset(, opto_origin_dictionary("REC LUBRICANTE OCULAR")))
    .Range(49) = charters(ItemData.Offset(, opto_origin_dictionary("RECOMENDACIONES OBS")))
    .Range(50) = charters(ItemData.Offset(, opto_origin_dictionary("REM_ VALORACION OFTALM_")))
    .Range(51) = charters(ItemData.Offset(, opto_origin_dictionary("REM_ TOPOGRAFIA CORNEAL")))
    .Range(52) = charters(ItemData.Offset(, opto_origin_dictionary("REM_ TRATAM_ ORTOPTICA")))
    .Range(53) = charters(ItemData.Offset(, opto_origin_dictionary("REM_ TEST FARNSWORTH")))
    .Range(54) = charters(ItemData.Offset(, opto_origin_dictionary("REALIZAR PRUEBA AMBULATORIA")))
    .Range(55) = charters(ItemData.Offset(, opto_origin_dictionary("REMISIONES OBS")))
    .Range(56) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES MENSUAL")))
    .Range(57) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES_BIMESTRALES")))
    .Range(58) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES TRIMESTRAL")))
    .Range(59) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES 6 MESES")))
    .Range(60) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES 1 ANO")))
    .Range(61) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES CONFIRMATORIA")))
    .Range(64) = autoIncrementOpto
    .Range(65) = autoIncrementDiagnostic
    DoEvents
  End With

End Sub