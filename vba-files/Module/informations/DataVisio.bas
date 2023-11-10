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
Dim aumentFromID As LongPtr
Public Sub VisioData(ByVal name_sheet As String)
  Dim visio_destiny_dictionary As Scripting.Dictionary
  Dim visio_origin_dictionary As Scripting.Dictionary
  Dim visio_destiny_header As Object, visio_origin_header As Object, visio_origin_value As Object
  Dim ItemVisioDestiny As Object, ItemVisioOrigin As Object, ItemData As Object, visio_origin As Object

  Set visio_origin = origin.Worksheets(name_sheet) '' VISIO DEL LIBRO ORIGEN ''
  visio_destiny.Select
  visio_destiny.Range("$A4").Select
  Set visio_destiny_header = visio_destiny.Range("$A3", visio_destiny.Range("$A3").End(xlToRight))
  Set visio_origin_header = visio_origin.Range("$A1", visio_origin.Range("$A1").End(xlToRight))
  Set visio_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set visio_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (visio_origin.Range("$A2") <> Empty And visio_origin.Range("$A3") <> Empty) Then
    Set visio_origin_value = visio_origin.Range("$A2", visio_origin.Range("$A2").End(xlDown))
  ElseIf (visio_origin.Range("$A2") <> Empty And visio_origin.Range("$A3") = Empty) Then
    Set visio_origin_value = visio_origin.Range("$A2")
  End If

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  Dim value_data As String
  For Each ItemVisioDestiny In visio_destiny_header
    value_data = visio_headers(ItemVisioDestiny)
    If visio_destiny_dictionary.Exists(value_data) = False And value_data <> Empty Then
      visio_destiny_dictionary.Add value_data, (ItemVisioDestiny.Column - 1)
    End If
  Next ItemVisioDestiny
  
  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemVisioOrigin In visio_origin_header
    value_data = visio_headers(ItemVisioOrigin)
    If visio_origin_dictionary.Exists(value_data) = False And value_data <> Empty Then
      visio_origin_dictionary.Add value_data, (ItemVisioOrigin.Column - 1)
    End If
  Next ItemVisioOrigin

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$9").value
  counts = visio_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  Dim type_exam As String
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

      type_exam = typeExams(Trim(ItemData.Offset(, visio_origin_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        ActiveCell.Offset(, visio_destiny_dictionary("NRO IDENFICACION")) = Trim(ItemData.Offset(, visio_origin_dictionary("NRO IDENFICACION")))
        ActiveCell.Offset(, visio_destiny_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")))
        ActiveCell.Offset(, visio_destiny_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")))
        ActiveCell.Offset(, visio_destiny_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")))
        ActiveCell.Offset(, visio_destiny_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS FOTOFOBIA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS FOTOFOBIA")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS OJO ROJO")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS OJO ROJO")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS LAGRIMEO")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS LAGRIMEO")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS VISION BORROSA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS VISION BORROSA")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS ARDOR")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS ARDOR")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS VISION DOBLE")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS VISION DOBLE")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS CANSANCIO")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS CANSANCIO")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS MALA VISION CERCANA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS MALA VISION CERCANA")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS DOLOR")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS DOLOR")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS MALA VISON LEJANA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS MALA VISON LEJANA")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS SECRECION")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS SECRECION")))
        ActiveCell.Offset(, visio_destiny_dictionary("SINTOMAS CEFALEA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS CEFALEA")))
        ActiveCell.Offset(, visio_destiny_dictionary("OTROS SINTOMAS")) = Trim(UCase(ItemData.Offset(, visio_origin_dictionary("OTROS SINTOMAS"))))
        ActiveCell.Offset(, visio_destiny_dictionary("CABEZA - PARPADOS")) = Trim(UCase(ItemData.Offset(, visio_origin_dictionary("CABEZA - PARPADOS"))))
        ActiveCell.Offset(, visio_destiny_dictionary("CABEZA - PARPADOS OBS")) = Trim(UCase(ItemData.Offset(, visio_origin_dictionary("CABEZA - PARPADOS OBS"))))
        ActiveCell.Offset(, visio_destiny_dictionary("CABEZA - CONJUNTIVAS")) = Trim(UCase(ItemData.Offset(, visio_origin_dictionary("CABEZA - CONJUNTIVAS"))))
        ActiveCell.Offset(, visio_destiny_dictionary("CABEZA - OBS CONJUNTIVAS")) = Trim(UCase(ItemData.Offset(, visio_origin_dictionary("CABEZA - OBS CONJUNTIVAS"))))
        ActiveCell.Offset(, visio_destiny_dictionary("CABEZA - ESCLERAS")) = Trim(UCase(ItemData.Offset(, visio_origin_dictionary("CABEZA - ESCLERAS"))))
        ActiveCell.Offset(, visio_destiny_dictionary("CABEZA - OBS ESCLERAS")) = Trim(UCase(ItemData.Offset(, visio_origin_dictionary("CABEZA - OBS ESCLERAS"))))
        ActiveCell.Offset(, visio_destiny_dictionary("CABEZA - PUPILAS")) = Trim(UCase(ItemData.Offset(, visio_origin_dictionary("CABEZA - PUPILAS"))))
        ActiveCell.Offset(, visio_destiny_dictionary("CABEZA - PUPILAS OBS")) = Trim(UCase(ItemData.Offset(, visio_origin_dictionary("CABEZA - PUPILAS OBS"))))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VL0OD NORMAL")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD NORMAL")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VL0OI NORMAL")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI NORMAL")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VP0OD NORMAL")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD NORMAL")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VP0OI NORMAL")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI NORMAL")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VL0OD DISMINUIDO")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD DISMINUIDO")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VL0OI DISMINUIDO")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI DISMINUIDO")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VP0OD DISMINUIDO")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD DISMINUIDO")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VP0OI DISMINUIDO")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI DISMINUIDO")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VL0OD NORMAL RX")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD NORMAL RX")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VL0OI NORMAL RX")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI NORMAL RX")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VP0OD NORMAL RX")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD NORMAL RX")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VP0OI NORMAL RX")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI NORMAL RX")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VL0OD DISMINUIDO RX")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD DISMINUIDO RX")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VL0OI DISMINUIDO RX")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI DISMINUIDO RX")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VP0OD DISMINUIDO RX")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD DISMINUIDO RX")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG VP0OI DISMINUIDO RX")) = Trim(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI DISMINUIDO RX")))
        ActiveCell.Offset(, visio_destiny_dictionary("IMP/DIAG OBS")) = Trim(UCase(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG OBS"))))
        ActiveCell.Offset(, visio_destiny_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")) = Trim(ItemData.Offset(, visio_origin_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")))
        ActiveCell.Offset(, visio_destiny_dictionary("REC USO RX PARA VISION PROX")) = Trim(ItemData.Offset(, visio_origin_dictionary("REC USO RX PARA VISION PROX")))
        ActiveCell.Offset(, visio_destiny_dictionary("REC USO AR VIDEO TRMINAL")) = Trim(ItemData.Offset(, visio_origin_dictionary("REC USO AR VIDEO TRMINAL")))
        ActiveCell.Offset(, visio_destiny_dictionary("REC USO RX DESCANSO")) = Trim(ItemData.Offset(, visio_origin_dictionary("REC USO RX DESCANSO")))
        ActiveCell.Offset(, visio_destiny_dictionary("REC USO LENTES PROT_ SOLAR")) = Trim(ItemData.Offset(, visio_origin_dictionary("REC USO LENTES PROT_ SOLAR")))
        ActiveCell.Offset(, visio_destiny_dictionary("REC USO PERMANENTE RX OPTICA")) = Trim(ItemData.Offset(, visio_origin_dictionary("REC USO PERMANENTE RX OPTICA")))
        ActiveCell.Offset(, visio_destiny_dictionary("REC USO EPP VISUAL")) = Trim(ItemData.Offset(, visio_origin_dictionary("REC USO EPP VISUAL")))
        ActiveCell.Offset(, visio_destiny_dictionary("REC PYP")) = Trim(ItemData.Offset(, visio_origin_dictionary("REC PYP")))
        ActiveCell.Offset(, visio_destiny_dictionary("REC PAUSAS ACTIVAS")) = Trim(ItemData.Offset(, visio_origin_dictionary("REC PAUSAS ACTIVAS")))
        ActiveCell.Offset(, visio_destiny_dictionary("REC LUBRICANTE OCULAR")) = Trim(ItemData.Offset(, visio_origin_dictionary("REC LUBRICANTE OCULAR")))
        ActiveCell.Offset(, visio_destiny_dictionary("RECOMENDACIONES OBS")) = Trim(ItemData.Offset(, visio_origin_dictionary("RECOMENDACIONES OBS")))
        ActiveCell.Offset(, visio_destiny_dictionary("REM_ VALORACION OFTALM_")) = Trim(ItemData.Offset(, visio_origin_dictionary("REM_ VALORACION OFTALM_")))
        ActiveCell.Offset(, visio_destiny_dictionary("REM_ VALORACION OPTO_ COMPLETA")) = Trim(ItemData.Offset(, visio_origin_dictionary("REM_ VALORACION OPTO_ COMPLETA")))
        ActiveCell.Offset(, visio_destiny_dictionary("REM_ TOPOGRAFIA CORNEAL")) = Trim(ItemData.Offset(, visio_origin_dictionary("REM_ TOPOGRAFIA CORNEAL")))
        ActiveCell.Offset(, visio_destiny_dictionary("REM_ TRATAM_ ORTOPTICA")) = Trim(ItemData.Offset(, visio_origin_dictionary("REM_ TRATAM_ ORTOPTICA")))
        ActiveCell.Offset(, visio_destiny_dictionary("REM_ TEST FARNSWORTH")) = Trim(ItemData.Offset(, visio_origin_dictionary("REM_ TEST FARNSWORTH")))
        ActiveCell.Offset(, visio_destiny_dictionary("REALIZAR PRUEBA AMBULATORIA")) = Trim(ItemData.Offset(, visio_origin_dictionary("REALIZAR PRUEBA AMBULATORIA")))
        ActiveCell.Offset(, visio_destiny_dictionary("OTRAS REMISIONES")) = Trim(ItemData.Offset(, visio_origin_dictionary("OTRAS REMISIONES")))
        ActiveCell.Offset(, visio_destiny_dictionary("CONTROL MENSUAL")) = Trim(ItemData.Offset(, visio_origin_dictionary("CONTROL MENSUAL")))
        ActiveCell.Offset(, visio_destiny_dictionary("CONTROLES_BIMESTRALES")) = Trim(ItemData.Offset(, visio_origin_dictionary("CONTROLES_BIMESTRALES")))
        ActiveCell.Offset(, visio_destiny_dictionary("CONTROL TRIMESTRAL")) = Trim(ItemData.Offset(, visio_origin_dictionary("CONTROL TRIMESTRAL")))
        ActiveCell.Offset(, visio_destiny_dictionary("CONTROL 6 MESES")) = Trim(ItemData.Offset(, visio_origin_dictionary("CONTROL 6 MESES")))
        ActiveCell.Offset(, visio_destiny_dictionary("CONTROL 1 ANO")) = Trim(ItemData.Offset(, visio_origin_dictionary("CONTROL 1 ANO")))
        ActiveCell.Offset(, visio_destiny_dictionary("CONTROL CONFIRMATORIA")) = Trim(ItemData.Offset(, visio_origin_dictionary("CONTROL CONFIRMATORIA")))
        If (ActiveCell.Row <> 4) Then
          aumentFromID = aumentFromID + 1
        End If
        ActiveCell.Offset(, visio_destiny_dictionary("ID_VISIOMETRIA")) = aumentFromID
        ActiveCell.Offset(1, 0).Select
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents
      End If
    Next ItemData
  End With

  Call dataDuplicate(visio_destiny.Range("tbl_visio[[#Data],[NRO IDENFICACION]]"))
  Call greaterThanOne(visio_destiny.Range("tbl_visio[[CONTROL MENSUAL]:[CONTROL CONFIRMATORIA]]"), "VISIO")
  Call iqualCero(visio_destiny.Range("tbl_visio[[CONTROL MENSUAL]:[CONTROL CONFIRMATORIA]]"), "VISIO")
  Call formatter(visio_destiny.Range("tbl_visio[[#Data],[NRO IDENFICACION]]"))

  Set visio_origin_value = Nothing
  Set visio_destiny_header = Nothing
  Set visio_origin_header = Nothing
  visio_destiny_dictionary.RemoveAll
  visio_origin_dictionary.RemoveAll

End Sub