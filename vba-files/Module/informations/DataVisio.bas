Attribute VB_Name = "DataVisio"
'namespace=vba-files\Module\informations
Option Explicit

'TODO: VisioData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - visio_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - visio_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - visio_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - visio_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - visio_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Dim visio_origin_dictionary As Scripting.Dictionary
Dim aumentFromID As LongPtr
Public Sub VisioData()
  Dim tbl_visio As Object, visio_origin As Object
  
  Set visio_origin = origin.Worksheets("VISIO").Range("A1") '' VISIO DEL LIBRO ORIGEN ''
  
  Set tbl_visio = visio_destiny.ListObjects("tbl_visio")
  Set visio_origin_dictionary = CreateObject("Scripting.Dictionary")
  
  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each item In Range(visio_origin, visio_origin.End(xlToRight))
    If visio_origin_dictionary.Exists(visio_headers(item)) = False Then
      visio_origin_dictionary.Add visio_headers(item), item.Column
    End If
  Next item

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$9").value
  counts = Ubound(origin.Worksheets("VISIO").Range("A1").CurrentRegion.Value, 1) - 1
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each item In Range(visio_origin.offset(1, 0), visio_origin.offset(1, 0).End(xlDown))
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & visio_destiny.Name
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

      If (typeExams(charters(item.Offset(, visio_origin_dictionary("TIPO EXAMEN") - 1))) <> "EGRESO") Then
        If item.value <> "" And item.Row = 2 Then
          Call addNewRegister(tbl_visio.ListRows(1), aumentFromID, item)
          DoEvents
        ElseIf item.value <> "" And item.Row > 2 Then
          aumentFromID = aumentFromID + 1
          Call addNewRegister(tbl_visio.ListRows.Add, aumentFromID, item)
          DoEvents
        ElseIf item.value = "" Or item.value = VbNullString Then
          Exit For
        End If
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
      End If
    Next item
  End With

  range("$A4").Select
  Call dataDuplicate
  range("$BL4:$BQ4").Select
  Call greaterThanOne
  range("$BL4:$BQ4").Select
  Call iqualCero
  range("$BR4").Select
  Call dataDuplicate
  range("$BS4").Select
  Call dataDuplicate
  range("$A4", range("$A4").End(xlDown)).Select
  Call formatter

  Set visio_origin = Nothing
  visio_origin_dictionary.RemoveAll

End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrement As LongPtr, ByVal information As Object)

  With table
    .Range(1) = charters(information(, visio_origin_dictionary("NRO IDENFICACION")))
    .Range(2) = charters_empty(information(, visio_origin_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")))
    .Range(3) = charters_empty(information(, visio_origin_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")))
    .Range(4) = charters_empty(information(, visio_origin_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")))
    .Range(5) = charters_empty(information(, visio_origin_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")))
    .Range(6) = charters_empty(information(, visio_origin_dictionary("SINTOMAS FOTOFOBIA")))
    .Range(7) = charters_empty(information(, visio_origin_dictionary("SINTOMAS OJO ROJO")))
    .Range(8) = charters_empty(information(, visio_origin_dictionary("SINTOMAS LAGRIMEO")))
    .Range(9) = charters_empty(information(, visio_origin_dictionary("SINTOMAS VISION BORROSA")))
    .Range(10) = charters_empty(information(, visio_origin_dictionary("SINTOMAS ARDOR")))
    .Range(11) = charters_empty(information(, visio_origin_dictionary("SINTOMAS VISION DOBLE")))
    .Range(12) = charters_empty(information(, visio_origin_dictionary("SINTOMAS CANSANCIO")))
    .Range(13) = charters_empty(information(, visio_origin_dictionary("SINTOMAS MALA VISION CERCANA")))
    .Range(14) = charters_empty(information(, visio_origin_dictionary("SINTOMAS DOLOR")))
    .Range(15) = charters_empty(information(, visio_origin_dictionary("SINTOMAS MALA VISON LEJANA")))
    .Range(16) = charters_empty(information(, visio_origin_dictionary("SINTOMAS SECRECION")))
    .Range(17) = charters_empty(information(, visio_origin_dictionary("SINTOMAS CEFALEA")))
    .Range(18) = charters(information(, visio_origin_dictionary("OTROS SINTOMAS")))
    .Range(19) = charters(information(, visio_origin_dictionary("CABEZA - PARPADOS")))
    .Range(20) = charters(information(, visio_origin_dictionary("CABEZA - PARPADOS OBS")))
    .Range(21) = charters(information(, visio_origin_dictionary("CABEZA - CONJUNTIVAS")))
    .Range(22) = charters(information(, visio_origin_dictionary("CABEZA - OBS CONJUNTIVAS")))
    .Range(23) = charters(information(, visio_origin_dictionary("CABEZA - ESCLERAS")))
    .Range(24) = charters(information(, visio_origin_dictionary("CABEZA - OBS ESCLERAS")))
    .Range(25) = charters(information(, visio_origin_dictionary("CABEZA - PUPILAS")))
    .Range(26) = charters(information(, visio_origin_dictionary("CABEZA - PUPILAS OBS")))
    .Range(27) = charters(information(, visio_origin_dictionary("IMP/DIAG VL0OD NORMAL")))
    .Range(28) = charters(information(, visio_origin_dictionary("IMP/DIAG VL0OI NORMAL")))
    .Range(29) = charters(information(, visio_origin_dictionary("IMP/DIAG VP0OD NORMAL")))
    .Range(30) = charters(information(, visio_origin_dictionary("IMP/DIAG VP0OI NORMAL")))
    .Range(31) = charters(information(, visio_origin_dictionary("IMP/DIAG VL0OD DISMINUIDO")))
    .Range(32) = charters(information(, visio_origin_dictionary("IMP/DIAG VL0OI DISMINUIDO")))
    .Range(33) = charters(information(, visio_origin_dictionary("IMP/DIAG VP0OD DISMINUIDO")))
    .Range(34) = charters(information(, visio_origin_dictionary("IMP/DIAG VP0OI DISMINUIDO")))
    .Range(35) = charters(information(, visio_origin_dictionary("IMP/DIAG VL0OD NORMAL RX")))
    .Range(36) = charters(information(, visio_origin_dictionary("IMP/DIAG VL0OI NORMAL RX")))
    .Range(37) = charters(information(, visio_origin_dictionary("IMP/DIAG VP0OD NORMAL RX")))
    .Range(38) = charters(information(, visio_origin_dictionary("IMP/DIAG VP0OI NORMAL RX")))
    .Range(39) = charters(information(, visio_origin_dictionary("IMP/DIAG VL0OD DISMINUIDO RX")))
    .Range(40) = charters(information(, visio_origin_dictionary("IMP/DIAG VL0OI DISMINUIDO RX")))
    .Range(41) = charters(information(, visio_origin_dictionary("IMP/DIAG VP0OD DISMINUIDO RX")))
    .Range(42) = charters(information(, visio_origin_dictionary("IMP/DIAG VP0OI DISMINUIDO RX")))
    .Range(45) = charters(information(, visio_origin_dictionary("IMP/DIAG OBS")))
    .Range(46) = charters(information(, visio_origin_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")))
    .Range(47) = charters(information(, visio_origin_dictionary("REC USO RX PARA VISION PROX")))
    .Range(48) = charters(information(, visio_origin_dictionary("REC USO AR VIDEO TRMINAL")))
    .Range(49) = charters(information(, visio_origin_dictionary("REC USO RX DESCANSO")))
    .Range(50) = charters(information(, visio_origin_dictionary("REC USO LENTES PROT_ SOLAR")))
    .Range(51) = charters(information(, visio_origin_dictionary("REC USO PERMANENTE RX OPTICA")))
    .Range(52) = charters(information(, visio_origin_dictionary("REC USO EPP VISUAL")))
    .Range(53) = charters(information(, visio_origin_dictionary("REC PYP")))
    .Range(54) = charters(information(, visio_origin_dictionary("REC PAUSAS ACTIVAS")))
    .Range(55) = charters(information(, visio_origin_dictionary("REC LUBRICANTE OCULAR")))
    .Range(56) = charters(information(, visio_origin_dictionary("RECOMENDACIONES OBS")))
    .Range(57) = charters(information(, visio_origin_dictionary("REM_ VALORACION OFTALM_")))
    .Range(58) = charters(information(, visio_origin_dictionary("REM_ VALORACION OPTO_ COMPLETA")))
    .Range(59) = charters(information(, visio_origin_dictionary("REM_ TOPOGRAFIA CORNEAL")))
    .Range(60) = charters(information(, visio_origin_dictionary("REM_ TRATAM_ ORTOPTICA")))
    .Range(61) = charters(information(, visio_origin_dictionary("REM_ TEST FARNSWORTH")))
    .Range(62) = charters(information(, visio_origin_dictionary("REALIZAR PRUEBA AMBULATORIA")))
    .Range(63) = charters(information(, visio_origin_dictionary("OTRAS REMISIONES")))
    .Range(64) = charters(information(, visio_origin_dictionary("CONTROL MENSUAL")))
    .Range(65) = charters(information(, visio_origin_dictionary("CONTROLES_BIMESTRALES")))
    .Range(66) = charters(information(, visio_origin_dictionary("CONTROL TRIMESTRAL")))
    .Range(67) = charters(information(, visio_origin_dictionary("CONTROL 6 MESES")))
    .Range(68) = charters(information(, visio_origin_dictionary("CONTROL 1 ANO")))
    .Range(69) = charters(information(, visio_origin_dictionary("CONTROL CONFIRMATORIA")))
    .Range(71) = autoIncrement
  End With

End Sub