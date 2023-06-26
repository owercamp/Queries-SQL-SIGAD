Attribute VB_Name = "DataOsteo"
Option Explicit

'TODO: OsteoData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - osteo_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - osteo_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - osteo_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - osteo_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - osteo_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Public Sub OsteoData()
  Dim osteo_destiny_dictionary As Scripting.Dictionary
  Dim osteo_origin_dictionary As Scripting.Dictionary
  Dim osteo_destiny_header As Object, osteo_origin_header As Object, osteo_origin_value As Object
  Dim ItemOsteoDestiny As Variant, ItemOsteoOrigin As Variant, ItemData As Variant
  Dim currenCell As range, aumentFromRow As LongPtr, aumentFromID As LongPtr
  
  Set osteo_origin = origin.Worksheets("OSTEO") '' OSTEO DEL LIBRO ORIGEN ''
  osteo_destiny.Select
  ActiveSheet.range("A4").Select
  Set currenCell = ActiveCell
  Set osteo_destiny_header = osteo_destiny.range("A3", osteo_destiny.range("A3").End(xlToRight))
  Set osteo_origin_header = osteo_origin.range("A1", osteo_origin.range("A1").End(xlToRight))
  Set osteo_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set osteo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (osteo_origin.range("A2") <> Empty And osteo_origin.range("A3") <> Empty) Then
    Set osteo_origin_value = osteo_origin.range("A2", osteo_origin.range("A2").End(xlDown))
  ElseIf (osteo_origin.range("A2") <> Empty And osteo_origin.range("A3") = Empty) Then
    Set osteo_origin_value = osteo_origin.range("A2")
  End If

  ''   En los diccionarios de "osteo_destiny_dictionary" y  "osteo_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemOsteoDestiny In osteo_destiny_header
    On Error Resume Next
    osteo_destiny_dictionary.Add osteo_headers(ItemOsteoDestiny), (ItemOsteoDestiny.Column - 1)
    On Error GoTo 0
  Next ItemOsteoDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemOsteoOrigin In osteo_origin_header
    On Error Resume Next
    osteo_origin_dictionary.Add osteo_headers(ItemOsteoOrigin), (ItemOsteoOrigin.Column - 1)
    On Error GoTo 0
  Next ItemOsteoOrigin

  numbers = 1
  porcentaje = 0
  aumentFromRow = 0
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$11").value
  counts = osteo_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each ItemData In osteo_origin_value
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & osteo_destiny.Name
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

      If (typeExams(charters(ItemData.Offset(, osteo_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, osteo_origin_dictionary("NRO IDENFICACION")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("CERVICALGIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("CERVICALGIA")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("CERVICALGIA OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("CERVICALGIA OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("EPICONDILITIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("EPICONDILITIS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("EPICONDILITIS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("EPICONDILITIS OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("LUMBALGIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("LUMBALGIA")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("LUMBALGIA OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("LUMBALGIA OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("S_ TUNEL CARPO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("S_ TUNEL CARPO")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("S_ TUNEL CARPO OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("S_ TUNEL CARPO OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("FRACTURAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("FRACTURAS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("FRACTURAS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("FRACTURAS OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("TENDINITIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("TENDINITIS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("TENDINITIS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("TENDINITIS OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("LESION EN MENISCOS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("LESION EN MENISCOS OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("LESION EN MENISCOS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("LESION EN MENISCOS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("ESGUINCES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("ESGUINCES")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("ESGUINCES OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("ESGUINCES OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("HOMBRO DOLOROSO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("HOMBRO DOLOROSO")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("HOMBRO DOLOROSO OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("HOMBRO DOLOROSO OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("RADICULOPATIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("RADICULOPATIA")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("RADICULOPATIA OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("RADICULOPATIA OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("BURSITIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("BURSITIS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("BURSITIS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("BURSITIS OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("ARTROSIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("ARTROSIS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("ARTROSIS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("ARTROSIS OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("ESCOLIOSIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("ESCOLIOSIS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("ESCOLIOSIS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("ESCOLIOSIS OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("RETRACCIONES MUSCULARES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("RETRACCIONES MUSCULARES")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("RETRACCIONES MUSCULARES OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("RETRACCIONES MUSCULARES OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("MALFORMACIONES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("MALFORMACIONES")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("MALFORMACIONES OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("MALFORMACIONES OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("DISCOPATIAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DISCOPATIAS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("DISCOPATIAS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("DISCOPATIAS OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("FIBROMALGIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("FIBROMALGIA")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("FIBROMALGIA OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("FIBROMALGIA OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("OTROS ANT_ OSTEOMUSCULARES")) = charters(ItemData.Offset(, osteo_origin_dictionary("OTROS ANT_ OSTEOMUSCULARES")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("PESO")) = charters(ItemData.Offset(, osteo_origin_dictionary("PESO")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("TALLA")) = charters(ItemData.Offset(, osteo_origin_dictionary("TALLA")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("DIAG_ PPAL")) = charters(ItemData.Offset(, osteo_origin_dictionary("DIAG_ PPAL")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("DIAG_ PPAL OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("DIAG_ PPAL OBS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("DIAG_ REL 1")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DIAG_ REL 1")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("DIAG_ REL 2")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DIAG_ REL 2")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("DIAG_ REL 3")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DIAG_ REL 3")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("REC/PERS ACT_ FISICA CARDIO 3X/SEMANA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS ACT_ FISICA CARDIO 3X/SEMANA")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("REC/PERS FORT_ 15 REPETICIONES/3 SERIES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS FORT_ 15 REPETICIONES/3 SERIES")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("REC/PERS EJERC_ ESTIRAMIENTO 20 SEG")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS EJERC_ ESTIRAMIENTO 20 SEG")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("REC/PERS AUTOCUIDADO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS AUTOCUIDADO")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("REC/PERS SEGUIMIENTO MEDICO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS SEGUIMIENTO MEDICO")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("REC/OCUP SVE PREVENSION LESIONES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP SVE PREVENSION LESIONES")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("REC/OCUP MANIPULACION DE CARGA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP MANIPULACION DE CARGA")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("REC/OCUP PAUSAS ACTIVAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP PAUSAS ACTIVAS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("REC/OCUP ANALISIS ERGONOMICOS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP ANALISIS ERGONOMICOS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("REC/OCUP EVITAR POSTURAS FORZADAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP EVITAR POSTURAS FORZADAS")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("RECOM_ OCUPACIONALES")) = charters(ItemData.Offset(, osteo_origin_dictionary("RECOM_ OCUPACIONALES")))
        currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("RECOM_ G/RALES")) = charters(ItemData.Offset(, osteo_origin_dictionary("RECOM_ G/RALES")))
        If (currenCell.Offset(aumentFromRow, 0).row = 4) Then
          currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("ID_OSTEOMUSCULAR")) = Trim(aumentFromID)
        Else
          aumentFromID = aumentFromID + 1
          currenCell.Offset(aumentFromRow, osteo_destiny_dictionary("ID_OSTEOMUSCULAR")) = Trim(aumentFromID)
        End If
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      aumentFromRow = aumentFromRow + 1
      DoEvents
    Next ItemData
  End With

  range("$A4").Select
  Call dataDuplicate
  range("$A4", range("$A4").End(xlDown)).Select
  Call formatter

  Set osteo_origin_value = Nothing
  Set osteo_destiny_header = Nothing
  Set osteo_origin_header = Nothing
  osteo_destiny_dictionary.RemoveAll
  osteo_origin_dictionary.RemoveAll

End Sub
