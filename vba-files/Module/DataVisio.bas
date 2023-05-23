Attribute VB_Name = "DataVisio"
Option Explicit

' VisioData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'------------------------------------------------------------------------------------------------------------------
' Variables:
' - visio_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de destino.
' - visio_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de origen.
' - visio_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
' - visio_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
' - visio_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
' - numbers: Una variable numerica para hacer un seguimiento del número de elementos de datos importados.
' - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
' - counts: Una variable numerica para almacenar el número total de elementos de datos de audio.
' - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
' - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
' - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'------------------------------------------------------------------------------------------------------------------
Public Sub VisioData()

  Dim visio_destiny_dictionary As Scripting.Dictionary
  Dim visio_origin_dictionary As Scripting.Dictionary
  Dim visio_destiny_header As Object, visio_origin_header As Object, visio_origin_value As Object
  Dim ItemVisioDestiny As Variant, ItemVisioOrigin As Variant, ItemData As Variant

  Call deleteFormatConditions
  Set visio_origin = origin.Worksheets("VISIO") '' VISIO DEL LIBRO ORIGEN ''
  visio_destiny.Select
  ActiveSheet.Range("A4").Select
  Set visio_destiny_header = visio_destiny.Range("A3", visio_destiny.Range("A3").End(xlToRight))
  Set visio_origin_header = visio_origin.Range("A1", visio_origin.Range("A1").End(xlToRight))
  Set visio_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set visio_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (visio_origin.Range("A2") <> Empty And visio_origin.Range("A3") <> Empty) Then
    Set visio_origin_value = visio_origin.Range("A2", visio_origin.Range("A2").End(xlDown))
  ElseIf (visio_origin.Range("A2") <> Empty And visio_origin.Range("A3") = Empty) Then
    Set visio_origin_value = visio_origin.Range("A2")
  End If

  ''   En los diccionarios de "visio_destiny_dictionary" y  "visio_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemVisioDestiny In visio_destiny_header
    On Error GoTo visioError
    visio_destiny_dictionary.Add visio_headers(ItemVisioDestiny), (ItemVisioDestiny.Column - 1)
  Next ItemVisioDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemVisioOrigin In visio_origin_header
    On Error GoTo visioError
    visio_origin_dictionary.Add visio_headers(ItemVisioOrigin), (ItemVisioOrigin.Column - 1)
  Next ItemVisioOrigin

  numbers = 1
  porcentaje = 0
  counts = visio_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  For Each ItemData In visio_origin_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & visio_destiny.Name
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
      If (typeExams(charters(ItemData.Offset(, visio_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        With ActiveCell
          .Offset(, visio_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, visio_origin_dictionary("NRO IDENFICACION")))
          .Offset(, visio_destiny_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")))
          .Offset(, visio_destiny_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")))
          .Offset(, visio_destiny_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")))
          .Offset(, visio_destiny_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")))
          .Offset(, visio_destiny_dictionary("SINTOMAS FOTOFOBIA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS FOTOFOBIA")))
          .Offset(, visio_destiny_dictionary("SINTOMAS OJO ROJO")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS OJO ROJO")))
          .Offset(, visio_destiny_dictionary("SINTOMAS LAGRIMEO")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS LAGRIMEO")))
          .Offset(, visio_destiny_dictionary("SINTOMAS VISION BORROSA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS VISION BORROSA")))
          .Offset(, visio_destiny_dictionary("SINTOMAS ARDOR")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS ARDOR")))
          .Offset(, visio_destiny_dictionary("SINTOMAS VISION DOBLE")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS VISION DOBLE")))
          .Offset(, visio_destiny_dictionary("SINTOMAS CANSANCIO")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS CANSANCIO")))
          .Offset(, visio_destiny_dictionary("SINTOMAS MALA VISION CERCANA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS MALA VISION CERCANA")))
          .Offset(, visio_destiny_dictionary("SINTOMAS DOLOR")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS DOLOR")))
          .Offset(, visio_destiny_dictionary("SINTOMAS MALA VISON LEJANA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS MALA VISON LEJANA")))
          .Offset(, visio_destiny_dictionary("SINTOMAS SECRECION")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS SECRECION")))
          .Offset(, visio_destiny_dictionary("SINTOMAS CEFALEA")) = charters_empty(ItemData.Offset(, visio_origin_dictionary("SINTOMAS CEFALEA")))
          .Offset(, visio_destiny_dictionary("OTROS SINTOMAS")) = charters(ItemData.Offset(, visio_origin_dictionary("OTROS SINTOMAS")))
          .Offset(, visio_destiny_dictionary("CABEZA - PARPADOS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - PARPADOS")))
          .Offset(, visio_destiny_dictionary("CABEZA - PARPADOS OBS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - PARPADOS OBS")))
          .Offset(, visio_destiny_dictionary("CABEZA - CONJUNTIVAS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - CONJUNTIVAS")))
          .Offset(, visio_destiny_dictionary("CABEZA - OBS CONJUNTIVAS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - OBS CONJUNTIVAS")))
          .Offset(, visio_destiny_dictionary("CABEZA - ESCLERAS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - ESCLERAS")))
          .Offset(, visio_destiny_dictionary("CABEZA - OBS ESCLERAS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - OBS ESCLERAS")))
          .Offset(, visio_destiny_dictionary("CABEZA - PUPILAS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - PUPILAS")))
          .Offset(, visio_destiny_dictionary("CABEZA - PUPILAS OBS")) = charters(ItemData.Offset(, visio_origin_dictionary("CABEZA - PUPILAS OBS")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VL0OD NORMAL")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD NORMAL")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VL0OI NORMAL")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI NORMAL")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VP0OD NORMAL")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD NORMAL")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VP0OI NORMAL")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI NORMAL")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VL0OD DISMINUIDO")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD DISMINUIDO")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VL0OI DISMINUIDO")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI DISMINUIDO")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VP0OD DISMINUIDO")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD DISMINUIDO")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VP0OI DISMINUIDO")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI DISMINUIDO")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VL0OD NORMAL RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD NORMAL RX")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VL0OI NORMAL RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI NORMAL RX")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VP0OD NORMAL RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD NORMAL RX")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VP0OI NORMAL RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI NORMAL RX")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VL0OD DISMINUIDO RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OD DISMINUIDO RX")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VL0OI DISMINUIDO RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VL0OI DISMINUIDO RX")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VP0OD DISMINUIDO RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OD DISMINUIDO RX")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG VP0OI DISMINUIDO RX")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG VP0OI DISMINUIDO RX")))
          .Offset(, visio_destiny_dictionary("IMP/DIAG OBS")) = charters(ItemData.Offset(, visio_origin_dictionary("IMP/DIAG OBS")))
          .Offset(, visio_destiny_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")) = charters(ItemData.Offset(, visio_origin_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")))
          .Offset(, visio_destiny_dictionary("REC USO RX PARA VISION PROX")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO RX PARA VISION PROX")))
          .Offset(, visio_destiny_dictionary("REC USO AR VIDEO TRMINAL")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO AR VIDEO TRMINAL")))
          .Offset(, visio_destiny_dictionary("REC USO RX DESCANSO")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO RX DESCANSO")))
          .Offset(, visio_destiny_dictionary("REC USO LENTES PROT_ SOLAR")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO LENTES PROT_ SOLAR")))
          .Offset(, visio_destiny_dictionary("REC USO PERMANENTE RX OPTICA")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO PERMANENTE RX OPTICA")))
          .Offset(, visio_destiny_dictionary("REC USO EPP VISUAL")) = charters(ItemData.Offset(, visio_origin_dictionary("REC USO EPP VISUAL")))
          .Offset(, visio_destiny_dictionary("REC PYP")) = charters(ItemData.Offset(, visio_origin_dictionary("REC PYP")))
          .Offset(, visio_destiny_dictionary("REC PAUSAS ACTIVAS")) = charters(ItemData.Offset(, visio_origin_dictionary("REC PAUSAS ACTIVAS")))
          .Offset(, visio_destiny_dictionary("REC LUBRICANTE OCULAR")) = charters(ItemData.Offset(, visio_origin_dictionary("REC LUBRICANTE OCULAR")))
          .Offset(, visio_destiny_dictionary("RECOMENDACIONES OBS")) = charters(ItemData.Offset(, visio_origin_dictionary("RECOMENDACIONES OBS")))
          .Offset(, visio_destiny_dictionary("REM_ VALORACION OFTALM_")) = charters(ItemData.Offset(, visio_origin_dictionary("REM_ VALORACION OFTALM_")))
          .Offset(, visio_destiny_dictionary("REM_ VALORACION OPTO_ COMPLETA")) = charters(ItemData.Offset(, visio_origin_dictionary("REM_ VALORACION OPTO_ COMPLETA")))
          .Offset(, visio_destiny_dictionary("REM_ TOPOGRAFIA CORNEAL")) = charters(ItemData.Offset(, visio_origin_dictionary("REM_ TOPOGRAFIA CORNEAL")))
          .Offset(, visio_destiny_dictionary("REM_ TRATAM_ ORTOPTICA")) = charters(ItemData.Offset(, visio_origin_dictionary("REM_ TRATAM_ ORTOPTICA")))
          .Offset(, visio_destiny_dictionary("REM_ TEST FARNSWORTH")) = charters(ItemData.Offset(, visio_origin_dictionary("REM_ TEST FARNSWORTH")))
          .Offset(, visio_destiny_dictionary("REALIZAR PRUEBA AMBULATORIA")) = charters(ItemData.Offset(, visio_origin_dictionary("REALIZAR PRUEBA AMBULATORIA")))
          .Offset(, visio_destiny_dictionary("OTRAS REMISIONES")) = charters(ItemData.Offset(, visio_origin_dictionary("OTRAS REMISIONES")))
          .Offset(, visio_destiny_dictionary("CONTROL MENSUAL")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROL MENSUAL")))
          .Offset(, visio_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROLES_BIMESTRALES")))
          .Offset(, visio_destiny_dictionary("CONTROL TRIMESTRAL")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROL TRIMESTRAL")))
          .Offset(, visio_destiny_dictionary("CONTROL 6 MESES")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROL 6 MESES")))
          .Offset(, visio_destiny_dictionary("CONTROL 1 ANO")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROL 1 ANO")))
          .Offset(, visio_destiny_dictionary("CONTROL CONFIRMATORIA")) = charters(ItemData.Offset(, visio_origin_dictionary("CONTROL CONFIRMATORIA")))
          If (.Row = 4) Then
            .Offset(, visio_destiny_dictionary("ID_VISIOMETRIA")) = Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$9").value)
          Else
            .Offset(, visio_destiny_dictionary("ID_VISIOMETRIA")) = .Offset(-1, visio_destiny_dictionary("ID_VISIOMETRIA")) + 1
          End If
          .Offset(1, 0).Select
        End With
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData

    Call dataDuplicate("$A4")
    Call greaterThanOne("$BL4:$BQ4")
    Call iqualCero("$BL4:$BQ4")
    Call dataDuplicate("$BR4")
    Call dataDuplicate("$BS4")
    Call formatter("$A4")

    Set visio_origin_value = Nothing
    Set visio_destiny_header = Nothing
    Set visio_origin_header = Nothing
    visio_destiny_dictionary.RemoveAll
    visio_origin_dictionary.RemoveAll

 visioError:
    Resume Next
End Sub
