Attribute VB_Name = "DataOsteo"
Option Explicit

Sub OsteoData()
  Dim osteo_destiny_dictionary As Scripting.Dictionary
  Dim osteo_origin_dictionary As Scripting.Dictionary
  Dim osteo_destiny_header, osteo_origin_header, osteo_origin_value As Object
  Dim ItemOsteoDestiny, ItemOsteoOrigin, ItemData As Variant

  Set osteo_origin = origin.Worksheets("OSTEO") '' OSTEO DEL LIBRO ORIGEN ''
  osteo_destiny.Select
  ActiveSheet.Range("A5").Select
  Set osteo_destiny_header = osteo_destiny.Range("A3", osteo_destiny.Range("A3").End(xlToRight))
  Set osteo_origin_header = osteo_origin.Range("A1", osteo_origin.Range("A1").End(xlToRight))
  Set osteo_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set osteo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (osteo_origin.Range("A2") <> Empty And osteo_origin.Range("A3") <> Empty) Then
    Set osteo_origin_value = osteo_origin.Range("A2", osteo_origin.Range("A2").End(xlDown))
  ElseIf (osteo_origin.Range("A2") <> Empty And osteo_origin.Range("A3") = Empty) Then
    Set osteo_origin_value = osteo_origin.Range("A2")
  End If

  '/***
  '   En los diccionarios de "osteo_destiny_dictionary" y  "osteo_origin_dictionary"
  '   se almacena los numeros de la columnas.
  '*/

  ' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO
  For Each ItemOsteoDestiny In osteo_destiny_header
    On Error Goto osteoError
    osteo_destiny_dictionary.Add osteo_headers(ItemOsteoDestiny), (ItemOsteoDestiny.Column - 1)
  Next ItemOsteoDestiny

  ' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN
  For Each ItemOsteoOrigin In osteo_origin_header
    On Error Goto osteoError
    osteo_origin_dictionary.Add osteo_headers(ItemOsteoOrigin), (ItemOsteoOrigin.Column - 1)
  Next ItemOsteoOrigin

  numbers = 1
  porcentaje = 0
  counts = osteo_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  For Each ItemData In osteo_origin_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & osteo_destiny.Name
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
              ActiveCell.offset(, osteo_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.offset(, osteo_origin_dictionary( "NRO IDENFICACION")))
              ActiveCell.offset(, osteo_destiny_dictionary("CERVICALGIA")) = charters(ItemData.offset(, osteo_origin_dictionary( "CERVICALGIA")))
              ActiveCell.offset(, osteo_destiny_dictionary("CERVICALGIA OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "CERVICALGIA OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("EPICONDILITIS")) = charters(ItemData.offset(, osteo_origin_dictionary( "EPICONDILITIS")))
              ActiveCell.offset(, osteo_destiny_dictionary("EPICONDILITIS OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "EPICONDILITIS OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("LUMBALGIA")) = charters(ItemData.offset(, osteo_origin_dictionary( "LUMBALGIA")))
              ActiveCell.offset(, osteo_destiny_dictionary("LUMBALGIA OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "LUMBALGIA OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("S_ TUNEL CARPO")) = charters(ItemData.offset(, osteo_origin_dictionary( "S_ TUNEL CARPO")))
              ActiveCell.offset(, osteo_destiny_dictionary("S_ TUNEL CARPO OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "S_ TUNEL CARPO OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("FRACTURAS")) = charters(ItemData.offset(, osteo_origin_dictionary( "FRACTURAS")))
              ActiveCell.offset(, osteo_destiny_dictionary("FRACTURAS OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "FRACTURAS OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("TENDINITIS")) = charters(ItemData.offset(, osteo_origin_dictionary( "TENDINITIS")))
              ActiveCell.offset(, osteo_destiny_dictionary("TENDINITIS OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "TENDINITIS OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("LESION EN MENISCOS OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "LESION EN MENISCOS OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("LESION EN MENISCOS")) = charters(ItemData.offset(, osteo_origin_dictionary( "LESION EN MENISCOS")))
              ActiveCell.offset(, osteo_destiny_dictionary("ESGUINCES")) = charters(ItemData.offset(, osteo_origin_dictionary( "ESGUINCES")))
              ActiveCell.offset(, osteo_destiny_dictionary("ESGUINCES OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "ESGUINCES OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("HOMBRO DOLOROSO")) = charters(ItemData.offset(, osteo_origin_dictionary( "HOMBRO DOLOROSO")))
              ActiveCell.offset(, osteo_destiny_dictionary("HOMBRO DOLOROSO OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "HOMBRO DOLOROSO OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("RADICULOPATIA")) = charters(ItemData.offset(, osteo_origin_dictionary( "RADICULOPATIA")))
              ActiveCell.offset(, osteo_destiny_dictionary("RADICULOPATIA OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "RADICULOPATIA OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("BURSITIS")) = charters(ItemData.offset(, osteo_origin_dictionary( "BURSITIS")))
              ActiveCell.offset(, osteo_destiny_dictionary("BURSITIS OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "BURSITIS OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("ARTROSIS")) = charters(ItemData.offset(, osteo_origin_dictionary( "ARTROSIS")))
              ActiveCell.offset(, osteo_destiny_dictionary("ARTROSIS OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "ARTROSIS OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("ESCOLIOSIS")) = charters(ItemData.offset(, osteo_origin_dictionary( "ESCOLIOSIS")))
              ActiveCell.offset(, osteo_destiny_dictionary("ESCOLIOSIS OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "ESCOLIOSIS OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("RETRACCIONES MUSCULARES")) = charters(ItemData.offset(, osteo_origin_dictionary( "RETRACCIONES MUSCULARES")))
              ActiveCell.offset(, osteo_destiny_dictionary("RETRACCIONES MUSCULARES OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "RETRACCIONES MUSCULARES OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("MALFORMACIONES")) = charters(ItemData.offset(, osteo_origin_dictionary( "MALFORMACIONES")))
              ActiveCell.offset(, osteo_destiny_dictionary("MALFORMACIONES OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "MALFORMACIONES OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("DISCOPATIAS")) = charters(ItemData.offset(, osteo_origin_dictionary( "DISCOPATIAS")))
              ActiveCell.offset(, osteo_destiny_dictionary("DISCOPATIAS OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "DISCOPATIAS OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("FIBROMALGIA")) = charters(ItemData.offset(, osteo_origin_dictionary( "FIBROMALGIA")))
              ActiveCell.offset(, osteo_destiny_dictionary("FIBROMALGIA OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "FIBROMALGIA OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("OTROS ANT_ OSTEOMUSCULARES")) = charters(ItemData.offset(, osteo_origin_dictionary( "OTROS ANT_ OSTEOMUSCULARES")))
              ActiveCell.offset(, osteo_destiny_dictionary("PESO")) = charters(ItemData.offset(, osteo_origin_dictionary( "PESO")))
              ActiveCell.offset(, osteo_destiny_dictionary("TALLA")) = charters(ItemData.offset(, osteo_origin_dictionary( "TALLA")))
              ActiveCell.offset(, osteo_destiny_dictionary("DIAG_ PPAL")) = charters(ItemData.offset(, osteo_origin_dictionary( "DIAG_ PPAL")))
              ActiveCell.offset(, osteo_destiny_dictionary("DIAG_ PPAL OBS")) = charters(ItemData.offset(, osteo_origin_dictionary( "DIAG_ PPAL OBS")))
              ActiveCell.offset(, osteo_destiny_dictionary("DIAG_ REL 1")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "DIAG_ REL 1")))
              ActiveCell.offset(, osteo_destiny_dictionary("DIAG_ REL 2")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "DIAG_ REL 2")))
              ActiveCell.offset(, osteo_destiny_dictionary("DIAG_ REL 3")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "DIAG_ REL 3")))
              ActiveCell.offset(, osteo_destiny_dictionary("REC/PERS ACT_ FISICA CARDIO 3X/SEMANA")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "REC/PERS ACT_ FISICA CARDIO 3X/SEMANA")))
              ActiveCell.offset(, osteo_destiny_dictionary("REC/PERS FORT_ 15 REPETICIONES/3 SERIES")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "REC/PERS FORT_ 15 REPETICIONES/3 SERIES")))
              ActiveCell.offset(, osteo_destiny_dictionary("REC/PERS EJERC_ ESTIRAMIENTO 20 SEG")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "REC/PERS EJERC_ ESTIRAMIENTO 20 SEG")))
              ActiveCell.offset(, osteo_destiny_dictionary("REC/PERS AUTOCUIDADO")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "REC/PERS AUTOCUIDADO")))
              ActiveCell.offset(, osteo_destiny_dictionary("REC/PERS SEGUIMIENTO MEDICO")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "REC/PERS SEGUIMIENTO MEDICO")))
              ActiveCell.offset(, osteo_destiny_dictionary("REC/OCUP SVE PREVENSION LESIONES")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "REC/OCUP SVE PREVENSION LESIONES")))
              ActiveCell.offset(, osteo_destiny_dictionary("REC/OCUP MANIPULACION DE CARGA")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "REC/OCUP MANIPULACION DE CARGA")))
              ActiveCell.offset(, osteo_destiny_dictionary("REC/OCUP PAUSAS ACTIVAS")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "REC/OCUP PAUSAS ACTIVAS")))
              ActiveCell.offset(, osteo_destiny_dictionary("REC/OCUP ANALISIS ERGONOMICOS")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "REC/OCUP ANALISIS ERGONOMICOS")))
              ActiveCell.offset(, osteo_destiny_dictionary("REC/OCUP EVITAR POSTURAS FORZADAS")) = charters_empty(ItemData.offset(, osteo_origin_dictionary( "REC/OCUP EVITAR POSTURAS FORZADAS")))
              ActiveCell.offset(, osteo_destiny_dictionary("RECOM_ OCUPACIONALES")) = charters(ItemData.offset(, osteo_origin_dictionary( "RECOM_ OCUPACIONALES")))
              ActiveCell.offset(, osteo_destiny_dictionary("RECOM_ G/RALES")) = charters(ItemData.offset(, osteo_origin_dictionary( "RECOM_ G/RALES")))
              ActiveCell.offset(, osteo_destiny_dictionary("ID_OSTEOMUSCULAR")) = ActiveCell.offset(-1, osteo_destiny_dictionary("ID_OSTEOMUSCULAR")) + 1
              ActiveCell.offset(1, 0).Select
              numbers = numbers + 1
              numbersGeneral = numbersGeneral + 1
              DoEvents
            Next ItemData

            Range("$A4").Select
            Call dataDuplicate
            Range("$A4", Range("$A4").End(xlDown)).Select
            Call formatter
            
            Set osteo_origin_value = Nothing
            Set osteo_destiny_header = Nothing
            Set osteo_origin_header = Nothing
            osteo_destiny_dictionary.RemoveAll
            osteo_origin_dictionary.RemoveAll

 osteoError:
            resume next
End Sub