Attribute VB_Name = "DataOsteo"
'namespace=vba-files\Module\informations
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
Dim aumentFromID As LongPtr
Public Sub OsteoData(ByVal name_sheet As String)
  Dim osteo_destiny_dictionary As Scripting.Dictionary
  Dim osteo_origin_dictionary As Scripting.Dictionary
  Dim osteo_destiny_header As Object, osteo_origin_header As Object, osteo_origin_value As Object
  Dim ItemOsteoDestiny As Object, ItemOsteoOrigin As Object, ItemData As Object, osteo_origin As Object

  Set osteo_origin = origin.Worksheets(name_sheet) '' OSTEO DEL LIBRO ORIGEN ''
  osteo_destiny.Select
  osteo_destiny.Range("$A4").Select
  Set osteo_destiny_header = osteo_destiny.Range("$A3", osteo_destiny.Range("$A3").End(xlToRight))
  Set osteo_origin_header = osteo_origin.Range("$A1", osteo_origin.Range("$A1").End(xlToRight))
  Set osteo_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set osteo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (osteo_origin.Range("$A2") <> Empty And osteo_origin.Range("$A3") <> Empty) Then
    Set osteo_origin_value = osteo_origin.Range("$A2", osteo_origin.Range("$A2").End(xlDown))
  ElseIf (osteo_origin.Range("$A2") <> Empty And osteo_origin.Range("$A3") = Empty) Then
    Set osteo_origin_value = osteo_origin.Range("$A2")
  End If

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  Dim value_data As String
  For Each ItemOsteoDestiny In osteo_destiny_header
    value_data = osteo_headers(ItemOsteoDestiny)
    If osteo_destiny_dictionary.Exists(value_data) = False And value_data <> Empty Then
      osteo_destiny_dictionary.Add value_data, (ItemOsteoDestiny.Column - 1)
    End If
  Next ItemOsteoDestiny
  
  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemOsteoOrigin In osteo_origin_header
    value_data = osteo_headers(ItemOsteoOrigin)
    If osteo_origin_dictionary.Exists(value_data) = False And value_data <> Empty Then
      osteo_origin_dictionary.Add value_data, (ItemOsteoOrigin.Column - 1)
    End If
  Next ItemOsteoOrigin

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$11").value
  counts = osteo_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  Dim type_exam As String
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

      type_exam = typeExams(Trim(ItemData.Offset(, osteo_origin_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        ActiveCell.Offset(, osteo_destiny_dictionary("NRO IDENFICACION")) = Trim(ItemData.Offset(, osteo_origin_dictionary("NRO IDENFICACION")))
        ActiveCell.Offset(, osteo_destiny_dictionary("CERVICALGIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("CERVICALGIA")))
        ActiveCell.Offset(, osteo_destiny_dictionary("CERVICALGIA OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("CERVICALGIA OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("EPICONDILITIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("EPICONDILITIS")))
        ActiveCell.Offset(, osteo_destiny_dictionary("EPICONDILITIS OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("EPICONDILITIS OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("LUMBALGIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("LUMBALGIA")))
        ActiveCell.Offset(, osteo_destiny_dictionary("LUMBALGIA OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("LUMBALGIA OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("S_ TUNEL CARPO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("S_ TUNEL CARPO")))
        ActiveCell.Offset(, osteo_destiny_dictionary("S_ TUNEL CARPO OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("S_ TUNEL CARPO OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("FRACTURAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("FRACTURAS")))
        ActiveCell.Offset(, osteo_destiny_dictionary("FRACTURAS OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("FRACTURAS OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("TENDINITIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("TENDINITIS")))
        ActiveCell.Offset(, osteo_destiny_dictionary("TENDINITIS OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("TENDINITIS OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("LESION EN MENISCOS OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("LESION EN MENISCOS OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("LESION EN MENISCOS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("LESION EN MENISCOS")))
        ActiveCell.Offset(, osteo_destiny_dictionary("ESGUINCES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("ESGUINCES")))
        ActiveCell.Offset(, osteo_destiny_dictionary("ESGUINCES OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("ESGUINCES OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("HOMBRO DOLOROSO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("HOMBRO DOLOROSO")))
        ActiveCell.Offset(, osteo_destiny_dictionary("HOMBRO DOLOROSO OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("HOMBRO DOLOROSO OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("RADICULOPATIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("RADICULOPATIA")))
        ActiveCell.Offset(, osteo_destiny_dictionary("RADICULOPATIA OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("RADICULOPATIA OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("BURSITIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("BURSITIS")))
        ActiveCell.Offset(, osteo_destiny_dictionary("BURSITIS OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("BURSITIS OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("ARTROSIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("ARTROSIS")))
        ActiveCell.Offset(, osteo_destiny_dictionary("ARTROSIS OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("ARTROSIS OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("ESCOLIOSIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("ESCOLIOSIS")))
        ActiveCell.Offset(, osteo_destiny_dictionary("ESCOLIOSIS OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("ESCOLIOSIS OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("RETRACCIONES MUSCULARES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("RETRACCIONES MUSCULARES")))
        ActiveCell.Offset(, osteo_destiny_dictionary("RETRACCIONES MUSCULARES OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("RETRACCIONES MUSCULARES OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("MALFORMACIONES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("MALFORMACIONES")))
        ActiveCell.Offset(, osteo_destiny_dictionary("MALFORMACIONES OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("MALFORMACIONES OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("DISCOPATIAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DISCOPATIAS")))
        ActiveCell.Offset(, osteo_destiny_dictionary("DISCOPATIAS OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("DISCOPATIAS OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("FIBROMALGIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("FIBROMALGIA")))
        ActiveCell.Offset(, osteo_destiny_dictionary("FIBROMALGIA OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("FIBROMALGIA OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("OTROS ANT_ OSTEOMUSCULARES")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("OTROS ANT_ OSTEOMUSCULARES"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("PESO")) = Trim(ItemData.Offset(, osteo_origin_dictionary("PESO")))
        ActiveCell.Offset(, osteo_destiny_dictionary("TALLA")) = Trim(ItemData.Offset(, osteo_origin_dictionary("TALLA")))
        ActiveCell.Offset(, osteo_destiny_dictionary("DIAG_ PPAL")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("DIAG_ PPAL"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("DIAG_ PPAL OBS")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("DIAG_ PPAL OBS"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("DIAG_ REL 1")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DIAG_ REL 1")))
        ActiveCell.Offset(, osteo_destiny_dictionary("DIAG_ REL 2")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DIAG_ REL 2")))
        ActiveCell.Offset(, osteo_destiny_dictionary("DIAG_ REL 3")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DIAG_ REL 3")))
        ActiveCell.Offset(, osteo_destiny_dictionary("REC/PERS ACT_ FISICA CARDIO 3X/SEMANA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS ACT_ FISICA CARDIO 3X/SEMANA")))
        ActiveCell.Offset(, osteo_destiny_dictionary("REC/PERS FORT_ 15 REPETICIONES/3 SERIES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS FORT_ 15 REPETICIONES/3 SERIES")))
        ActiveCell.Offset(, osteo_destiny_dictionary("REC/PERS EJERC_ ESTIRAMIENTO 20 SEG")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS EJERC_ ESTIRAMIENTO 20 SEG")))
        ActiveCell.Offset(, osteo_destiny_dictionary("REC/PERS AUTOCUIDADO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS AUTOCUIDADO")))
        ActiveCell.Offset(, osteo_destiny_dictionary("REC/PERS SEGUIMIENTO MEDICO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS SEGUIMIENTO MEDICO")))
        ActiveCell.Offset(, osteo_destiny_dictionary("REC/OCUP SVE PREVENSION LESIONES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP SVE PREVENSION LESIONES")))
        ActiveCell.Offset(, osteo_destiny_dictionary("REC/OCUP MANIPULACION DE CARGA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP MANIPULACION DE CARGA")))
        ActiveCell.Offset(, osteo_destiny_dictionary("REC/OCUP PAUSAS ACTIVAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP PAUSAS ACTIVAS")))
        ActiveCell.Offset(, osteo_destiny_dictionary("REC/OCUP ANALISIS ERGONOMICOS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP ANALISIS ERGONOMICOS")))
        ActiveCell.Offset(, osteo_destiny_dictionary("REC/OCUP EVITAR POSTURAS FORZADAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP EVITAR POSTURAS FORZADAS")))
        ActiveCell.Offset(, osteo_destiny_dictionary("RECOM_ OCUPACIONALES")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("RECOM_ OCUPACIONALES"))))
        ActiveCell.Offset(, osteo_destiny_dictionary("RECOM_ G/RALES")) = Trim(UCase(ItemData.Offset(, osteo_origin_dictionary("RECOM_ G/RALES"))))
        If (ActiveCell.Row <> 4) Then
          aumentFromID = aumentFromID + 1
        End If
        ActiveCell.Offset(, osteo_destiny_dictionary("ID_OSTEOMUSCULAR")) = aumentFromID
        ActiveCell.Offset(1, 0).Select
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents
      End If
    Next ItemData
  End With

  Call dataDuplicate(osteo_destiny.Range("tbl_osteo[[#Data],[NRO IDENFICACION]]"))
  Call formatter(osteo_destiny.Range("tbl_osteo[[#Data],[NRO IDENFICACION]]"))

  Set osteo_origin_value = Nothing
  Set osteo_destiny_header = Nothing
  Set osteo_origin_header = Nothing
  osteo_destiny_dictionary.RemoveAll
  osteo_origin_dictionary.RemoveAll

End Sub