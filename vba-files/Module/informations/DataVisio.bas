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
Public Sub VisioData()

  Dim visio_destiny_dictionary As Scripting.Dictionary
  Dim visio_origin_dictionary As Scripting.Dictionary
  Dim visio_destiny_header As Object, visio_origin_header As Object, visio_origin_value As Object
  Dim ItemVisioDestiny As Variant, ItemVisioOrigin As Variant, ItemData As Variant
  Dim currenCell As range, aumentFromRow As LongPtr, aumentFromID As LongPtr
  
  Set visio_origin = origin.Worksheets("VISIO") '' VISIO DEL LIBRO ORIGEN ''
  visio_destiny.Select
  ActiveSheet.range("A4").Select
  Set currenCell = ActiveCell
  Set visio_destiny_header = visio_destiny.range("A3", visio_destiny.range("A3").End(xlToRight))
  Set visio_origin_header = visio_origin.range("A1", visio_origin.range("A1").End(xlToRight))
  Set visio_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set visio_origin_dictionary = CreateObject("Scripting.Dictionary")
  
  If (visio_origin.range("A2") <> Empty And visio_origin.range("A3") <> Empty) Then
    Set visio_origin_value = visio_origin.range("A2", visio_origin.range("A2").End(xlDown))
  ElseIf (visio_origin.range("A2") <> Empty And visio_origin.range("A3") = Empty) Then
    Set visio_origin_value = visio_origin.range("A2")
  End If

  ''   En los diccionarios de "visio_destiny_dictionary" y  "visio_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemVisioDestiny In visio_destiny_header
    On Error Resume Next
    visio_destiny_dictionary.Add visio_headers(ItemVisioDestiny), (ItemVisioDestiny.Column - 1)
    On Error GoTo 0
  Next ItemVisioDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemVisioOrigin In visio_origin_header
    On Error Resume Next
    visio_origin_dictionary.Add visio_headers(ItemVisioOrigin), (ItemVisioOrigin.Column - 1)
    On Error GoTo 0
  Next ItemVisioOrigin

  numbers = 1
  porcentaje = 0
  aumentFromRow = 0
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$9").value
  counts = visio_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each ItemData In visio_origin_value
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
      
      If (typeExams(charters(ItemData.Offset(, visio_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, visio_origin_dictionary("NRO IDENFICACION")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS FOTOFOBIA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS FOTOFOBIA")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS OJO ROJO")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS OJO ROJO")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS LAGRIMEO")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS LAGRIMEO")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS VISION BORROSA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS VISION BORROSA")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS ARDOR")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS ARDOR")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS VISION DOBLE")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS VISION DOBLE")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS CANSANCIO")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS CANSANCIO")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS MALA VISION CERCANA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS MALA VISION CERCANA")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS DOLOR")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS DOLOR")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS MALA VISON LEJANA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS MALA VISON LEJANA")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS SECRECION")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS SECRECION")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("SINTOMAS CEFALEA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS CEFALEA")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("OTROS SINTOMAS")) = charters(ItemData.Offset(, visio_origin_dictionary("OTROS SINTOMAS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CABEZA - PARPADOS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - PARPADOS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CABEZA - PARPADOS OBS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - PARPADOS OBS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CABEZA - CONJUNTIVAS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - CONJUNTIVAS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CABEZA - OBS CONJUNTIVAS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - OBS CONJUNTIVAS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CABEZA - ESCLERAS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - ESCLERAS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CABEZA - OBS ESCLERAS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - OBS ESCLERAS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CABEZA - PUPILAS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - PUPILAS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CABEZA - PUPILAS OBS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - PUPILAS OBS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VL0OD NORMAL")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD NORMAL")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VL0OI NORMAL")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI NORMAL")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VP0OD NORMAL")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD NORMAL")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VP0OI NORMAL")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI NORMAL")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VL0OD DISMINUIDO")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD DISMINUIDO")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VL0OI DISMINUIDO")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI DISMINUIDO")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VP0OD DISMINUIDO")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD DISMINUIDO")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VP0OI DISMINUIDO")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI DISMINUIDO")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VL0OD NORMAL RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD NORMAL RX")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VL0OI NORMAL RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI NORMAL RX")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VP0OD NORMAL RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD NORMAL RX")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VP0OI NORMAL RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI NORMAL RX")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VL0OD DISMINUIDO RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD DISMINUIDO RX")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VL0OI DISMINUIDO RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI DISMINUIDO RX")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VP0OD DISMINUIDO RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD DISMINUIDO RX")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG VP0OI DISMINUIDO RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI DISMINUIDO RX")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("IMP/DIAG OBS")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG OBS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")) = charters(ItemData.Offset(, visio_origin_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REC USO RX PARA VISION PROX")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO RX PARA VISION PROX")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REC USO AR VIDEO TRMINAL")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO AR VIDEO TRMINAL")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REC USO RX DESCANSO")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO RX DESCANSO")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REC USO LENTES PROT_ SOLAR")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO LENTES PROT_ SOLAR")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REC USO PERMANENTE RX OPTICA")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO PERMANENTE RX OPTICA")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REC USO EPP VISUAL")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO EPP VISUAL")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REC PYP")) = charters(ItemData.Offset(, visio_origin_dictionary("REC PYP")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REC PAUSAS ACTIVAS")) = charters(ItemData.Offset(, visio_origin_dictionary("REC PAUSAS ACTIVAS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REC LUBRICANTE OCULAR")) = charters(ItemData.Offset(, visio_origin_dictionary("REC LUBRICANTE OCULAR")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("RECOMENDACIONES OBS")) = charters(ItemData.Offset(, visio_origin_dictionary("RECOMENDACIONES OBS")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REM_ VALORACION OFTALM_")) = charters(ItemData.Offset(, visio_origin_dictionary("REM_ VALORACION OFTALM_")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REM_ VALORACION OPTO_ COMPLETA")) = charters(ItemData.Offset(, visio_origin_dictionary("REM_ VALORACION OPTO_ COMPLETA")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REM_ TOPOGRAFIA CORNEAL")) = charters(ItemData.Offset(, visio_origin_dictionary("REM_ TOPOGRAFIA CORNEAL")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REM_ TRATAM_ ORTOPTICA")) = charters(ItemData.Offset(, visio_origin_dictionary("REM_ TRATAM_ ORTOPTICA")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REM_ TEST FARNSWORTH")) = charters(ItemData.Offset(, visio_origin_dictionary("REM_ TEST FARNSWORTH")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("REALIZAR PRUEBA AMBULATORIA")) = charters(ItemData.Offset(, visio_origin_dictionary("REALIZAR PRUEBA AMBULATORIA")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("OTRAS REMISIONES")) = charters(ItemData.Offset(, visio_origin_dictionary("OTRAS REMISIONES")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CONTROL MENSUAL")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROL MENSUAL")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROLES_BIMESTRALES")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CONTROL TRIMESTRAL")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROL TRIMESTRAL")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CONTROL 6 MESES")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROL 6 MESES")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CONTROL 1 ANO")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROL 1 ANO")))
        currenCell.Offset(aumentFromRow, visio_destiny_dictionary("CONTROL CONFIRMATORIA")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROL CONFIRMATORIA")))
        If (currenCell.Offset(aumentFromRow, 0).row = 4) Then
          currenCell.Offset(aumentFromRow, visio_destiny_dictionary("ID_VISIOMETRIA")) = Trim(aumentFromID)
        Else
          aumentFromID = aumentFromID + 1
          currenCell.Offset(aumentFromRow, visio_destiny_dictionary("ID_VISIOMETRIA")) = Trim(aumentFromID)
        End If
        aumentFromRow = aumentFromRow + 1
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData
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

  Set visio_origin_value = Nothing
  Set visio_destiny_header = Nothing
  Set visio_origin_header = Nothing
  visio_destiny_dictionary.RemoveAll
  visio_origin_dictionary.RemoveAll

End Sub
