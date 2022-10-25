Attribute VB_Name = "DataVisio"
Option Explicit

Sub VisioData()

  Dim visio_destiny_dictionary As Scripting.Dictionary
  Dim visio_origin_dictionary As Scripting.Dictionary
  Dim visio_destiny_header, visio_origin_header, visio_origin_value As Object
  Dim ItemVisioDestiny, ItemVisioOrigin, ItemData As Variant

  Set visio_origin = origin.Worksheets("VISIO") '' VISIO DEL LIBRO ORIGEN ''
  visio_destiny.Select
  ActiveSheet.Range("A5").Select
  Set visio_destiny_header = visio_destiny.Range("A3", visio_destiny.Range("A3").End(xlToRight))
  Set visio_origin_header = visio_origin.Range("A1", visio_origin.Range("A1").End(xlToRight))
  Set visio_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set visio_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (visio_origin.Range("A2") <> Empty And visio_origin.Range("A3") <> Empty) Then
    Set visio_origin_value = visio_origin.Range("A2", visio_origin.Range("A2").End(xlDown))
  ElseIf (visio_origin.Range("A2") <> Empty And visio_origin.Range("A3") = Empty) Then
    Set visio_origin_value = visio_origin.Range("A2")
  End If

  '/***
  '   En los diccionarios de "visio_destiny_dictionary" y  "visio_origin_dictionary"
  '   se almacena los numeros de la columnas.
  '*/

  ' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO
  For Each ItemVisioDestiny In visio_destiny_header
    On Error Goto visioError
    visio_destiny_dictionary.Add visio_headers(ItemVisioDestiny), (ItemVisioDestiny.Column - 1)
  Next ItemVisioDestiny

  ' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN
  For Each ItemVisioOrigin In visio_origin_header
    On Error Goto visioError
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
      If formImports.ProgressBarGeneral.Width > (formImports.content_ProgressBarGeneral.Width / 2) Then: formImports.porcentageGeneral.ForeColor = RGB(255, 255, 255)
        If formImports.ProgressBarGeneral.Width < (formImports.content_ProgressBarGeneral.Width / 2) Then: formImports.porcentageGeneral.ForeColor = RGB(0, 0, 0)
          If formImports.ProgressBarOneforOne.Width > (formImports.content_ProgressBarOneforOne.Width / 2) Then: formImports.porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
            If formImports.ProgressBarOneforOne.Width < (formImports.content_ProgressBarOneforOne.Width / 2) Then: formImports.porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
              ActiveCell.offset(, visio_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.offset(, visio_origin_dictionary( "NRO IDENFICACION")))
              ActiveCell.offset(, visio_destiny_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "VISIO/ANT_ LABORAL ILUMINACION INADECUADA")))
              ActiveCell.offset(, visio_destiny_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "VISIO/ANT_ LABORALVISIO RADIACIONES UV")))
              ActiveCell.offset(, visio_destiny_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "VISIO/ANT_ LABORAL MALA VENTILACION")))
              ActiveCell.offset(, visio_destiny_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "VISIO/ANT_ LABORAL GASES TOXICOS")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS FOTOFOBIA")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS FOTOFOBIA")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS OJO ROJO")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS OJO ROJO")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS LAGRIMEO")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS LAGRIMEO")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS VISION BORROSA")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS VISION BORROSA")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS ARDOR")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS ARDOR")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS VISION DOBLE")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS VISION DOBLE")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS CANSANCIO")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS CANSANCIO")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS MALA VISION CERCANA")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS MALA VISION CERCANA")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS DOLOR")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS DOLOR")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS MALA VISON LEJANA")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS MALA VISON LEJANA")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS SECRECION")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS SECRECION")))
              ActiveCell.offset(, visio_destiny_dictionary("SINTOMAS CEFALEA")) = charters_empty(ItemData.offset(, visio_origin_dictionary( "SINTOMAS CEFALEA")))
              ActiveCell.offset(, visio_destiny_dictionary("OTROS SINTOMAS")) = charters(ItemData.offset(, visio_origin_dictionary( "OTROS SINTOMAS")))
              ActiveCell.offset(, visio_destiny_dictionary("CABEZA - PARPADOS")) = charters(ItemData.offset(, visio_origin_dictionary( "CABEZA - PARPADOS")))
              ActiveCell.offset(, visio_destiny_dictionary("CABEZA - PARPADOS OBS")) = charters(ItemData.offset(, visio_origin_dictionary( "CABEZA - PARPADOS OBS")))
              ActiveCell.offset(, visio_destiny_dictionary("CABEZA - CONJUNTIVAS")) = charters(ItemData.offset(, visio_origin_dictionary( "CABEZA - CONJUNTIVAS")))
              ActiveCell.offset(, visio_destiny_dictionary("CABEZA - OBS CONJUNTIVAS")) = charters(ItemData.offset(, visio_origin_dictionary( "CABEZA - OBS CONJUNTIVAS")))
              ActiveCell.offset(, visio_destiny_dictionary("CABEZA - ESCLERAS")) = charters(ItemData.offset(, visio_origin_dictionary( "CABEZA - ESCLERAS")))
              ActiveCell.offset(, visio_destiny_dictionary("CABEZA - OBS ESCLERAS")) = charters(ItemData.offset(, visio_origin_dictionary( "CABEZA - OBS ESCLERAS")))
              ActiveCell.offset(, visio_destiny_dictionary("CABEZA - PUPILAS")) = charters(ItemData.offset(, visio_origin_dictionary( "CABEZA - PUPILAS")))
              ActiveCell.offset(, visio_destiny_dictionary("CABEZA - PUPILAS OBS")) = charters(ItemData.offset(, visio_origin_dictionary( "CABEZA - PUPILAS OBS")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VL0OD NORMAL")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VL0OD NORMAL")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VL0OI NORMAL")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VL0OI NORMAL")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VP0OD NORMAL")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VP0OD NORMAL")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VP0OI NORMAL")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VP0OI NORMAL")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VL0OD DISMINUIDO")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VL0OD DISMINUIDO")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VL0OI DISMINUIDO")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VL0OI DISMINUIDO")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VP0OD DISMINUIDO")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VP0OD DISMINUIDO")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VP0OI DISMINUIDO")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VP0OI DISMINUIDO")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VL0OD NORMAL RX")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VL0OD NORMAL RX")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VL0OI NORMAL RX")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VL0OI NORMAL RX")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VP0OD NORMAL RX")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VP0OD NORMAL RX")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VP0OI NORMAL RX")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VP0OI NORMAL RX")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VL0OD DISMINUIDO RX")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VL0OD DISMINUIDO RX")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VL0OI DISMINUIDO RX")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VL0OI DISMINUIDO RX")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VP0OD DISMINUIDO RX")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VP0OD DISMINUIDO RX")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG VP0OI DISMINUIDO RX")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG VP0OI DISMINUIDO RX")))
              ActiveCell.offset(, visio_destiny_dictionary("RESULTADO VISIO")) = charters(ItemData.offset(, visio_origin_dictionary( "RESULTADO VISIO")))
              ActiveCell.offset(, visio_destiny_dictionary("IMP/DIAG OBS")) = charters(ItemData.offset(, visio_origin_dictionary( "IMP/DIAG OBS")))
              ActiveCell.offset(, visio_destiny_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")) = charters(ItemData.offset(, visio_origin_dictionary( "REC CORRECCION VISUAL PARA TRABAJAR")))
              ActiveCell.offset(, visio_destiny_dictionary("REC USO RX PARA VISION PROX")) = charters(ItemData.offset(, visio_origin_dictionary( "REC USO RX PARA VISION PROX")))
              ActiveCell.offset(, visio_destiny_dictionary("REC USO AR VIDEO TRMINAL")) = charters(ItemData.offset(, visio_origin_dictionary( "REC USO AR VIDEO TRMINAL")))
              ActiveCell.offset(, visio_destiny_dictionary("REC USO RX DESCANSO")) = charters(ItemData.offset(, visio_origin_dictionary( "REC USO RX DESCANSO")))
              ActiveCell.offset(, visio_destiny_dictionary("REC USO LENTES PROT_ SOLAR")) = charters(ItemData.offset(, visio_origin_dictionary( "REC USO LENTES PROT_ SOLAR")))
              ActiveCell.offset(, visio_destiny_dictionary("REC USO PERMANENTE RX OPTICA")) = charters(ItemData.offset(, visio_origin_dictionary( "REC USO PERMANENTE RX OPTICA")))
              ActiveCell.offset(, visio_destiny_dictionary("REC USO EPP VISUAL")) = charters(ItemData.offset(, visio_origin_dictionary( "REC USO EPP VISUAL")))
              ActiveCell.offset(, visio_destiny_dictionary("REC PYP")) = charters(ItemData.offset(, visio_origin_dictionary( "REC PYP")))
              ActiveCell.offset(, visio_destiny_dictionary("REC PAUSAS ACTIVAS")) = charters(ItemData.offset(, visio_origin_dictionary( "REC PAUSAS ACTIVAS")))
              ActiveCell.offset(, visio_destiny_dictionary("REC LUBRICANTE OCULAR")) = charters(ItemData.offset(, visio_origin_dictionary( "REC LUBRICANTE OCULAR")))
              ActiveCell.offset(, visio_destiny_dictionary("RECOMENDACIONES OBS")) = charters(ItemData.offset(, visio_origin_dictionary( "RECOMENDACIONES OBS")))
              ActiveCell.offset(, visio_destiny_dictionary("REM_ VALORACION OFTALM_")) = charters(ItemData.offset(, visio_origin_dictionary( "REM_ VALORACION OFTALM_")))
              ActiveCell.offset(, visio_destiny_dictionary("REM_ VALORACION OPTO_ COMPLETA")) = charters(ItemData.offset(, visio_origin_dictionary( "REM_ VALORACION OPTO_ COMPLETA")))
              ActiveCell.offset(, visio_destiny_dictionary("REM_ TOPOGRAFIA CORNEAL")) = charters(ItemData.offset(, visio_origin_dictionary( "REM_ TOPOGRAFIA CORNEAL")))
              ActiveCell.offset(, visio_destiny_dictionary("REM_ TRATAM_ ORTOPTICA")) = charters(ItemData.offset(, visio_origin_dictionary( "REM_ TRATAM_ ORTOPTICA")))
              ActiveCell.offset(, visio_destiny_dictionary("REM_ TEST FARNSWORTH")) = charters(ItemData.offset(, visio_origin_dictionary( "REM_ TEST FARNSWORTH")))
              ActiveCell.offset(, visio_destiny_dictionary("REALIZAR PRUEBA AMBULATORIA")) = charters(ItemData.offset(, visio_origin_dictionary( "REALIZAR PRUEBA AMBULATORIA")))
              ActiveCell.offset(, visio_destiny_dictionary("OTRAS REMISIONES")) = charters(ItemData.offset(, visio_origin_dictionary( "OTRAS REMISIONES")))
              ActiveCell.offset(, visio_destiny_dictionary("CONTROL MENSUAL")) = charters(ItemData.offset(, visio_origin_dictionary( "CONTROL MENSUAL")))
              ActiveCell.offset(, visio_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters(ItemData.offset(, visio_origin_dictionary( "CONTROLES_BIMESTRALES")))
              ActiveCell.offset(, visio_destiny_dictionary("CONTROL TRIMESTRAL")) = charters(ItemData.offset(, visio_origin_dictionary( "CONTROL TRIMESTRAL")))
              ActiveCell.offset(, visio_destiny_dictionary("CONTROL 6 MESES")) = charters(ItemData.offset(, visio_origin_dictionary( "CONTROL 6 MESES")))
              ActiveCell.offset(, visio_destiny_dictionary("CONTROL 1 ANO")) = charters(ItemData.offset(, visio_origin_dictionary( "CONTROL 1 ANO")))
              ActiveCell.offset(, visio_destiny_dictionary("CONTROL CONFIRMATORIA")) = charters(ItemData.offset(, visio_origin_dictionary( "CONTROL CONFIRMATORIA")))
              ActiveCell.offset(, visio_destiny_dictionary("ID_VISIOMETRIA")) = ActiveCell.offset(-1, visio_destiny_dictionary("ID_VISIOMETRIA")) + 1
              ActiveCell.offset(1, 0).Select
              numbers = numbers + 1
              numbersGeneral = numbersGeneral + 1
              DoEvents
            Next ItemData

            Range("$A4").Select
            Call dataDuplicate
            Range("$BL4:$BQ4").Select
            Call greaterThanOne
            Range("$BL4:$BQ4").Select
            Call iqualCero
            Range("$BR4").Select
            Call dataDuplicate
            Range("$BS4").Select
            Call dataDuplicate
            Range("$A4", Range("$A4").End(xlDown)).Select
            Call formatter

            Set visio_origin_value = Nothing
            Set visio_destiny_header = Nothing
            Set visio_origin_header = Nothing
            visio_destiny_dictionary.RemoveAll
            visio_origin_dictionary.RemoveAll

 visioError:
            resume next
End Sub