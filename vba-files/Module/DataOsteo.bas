Attribute VB_Name = "DataOsteo"
Option Explicit

' OsteoData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'------------------------------------------------------------------------------------------------------------------
' Variables:
' - osteo_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de destino.
' - osteo_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de origen.
' - osteo_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
' - osteo_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
' - osteo_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
' - numbers: Una variable numerica para hacer un seguimiento del número de elementos de datos importados.
' - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
' - counts: Una variable numerica para almacenar el número total de elementos de datos de audio.
' - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
' - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
' - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'------------------------------------------------------------------------------------------------------------------
Public Sub OsteoData()
  Dim osteo_destiny_dictionary As Scripting.Dictionary
  Dim osteo_origin_dictionary As Scripting.Dictionary
  Dim osteo_destiny_header As Object, osteo_origin_header As Object, osteo_origin_value As Object
  Dim ItemOsteoDestiny As Variant, ItemOsteoOrigin As Variant, ItemData As Variant

  Set osteo_origin = origin.Worksheets("OSTEO") '' OSTEO DEL LIBRO ORIGEN ''
  osteo_destiny.Select
  ActiveSheet.Range("A4").Select
  Set osteo_destiny_header = osteo_destiny.Range("A3", osteo_destiny.Range("A3").End(xlToRight))
  Set osteo_origin_header = osteo_origin.Range("A1", osteo_origin.Range("A1").End(xlToRight))
  Set osteo_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set osteo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (osteo_origin.Range("A2") <> Empty And osteo_origin.Range("A3") <> Empty) Then
    Set osteo_origin_value = osteo_origin.Range("A2", osteo_origin.Range("A2").End(xlDown))
  ElseIf (osteo_origin.Range("A2") <> Empty And osteo_origin.Range("A3") = Empty) Then
    Set osteo_origin_value = osteo_origin.Range("A2")
  End If

  ''   En los diccionarios de "osteo_destiny_dictionary" y  "osteo_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemOsteoDestiny In osteo_destiny_header
    On Error GoTo osteoError
    osteo_destiny_dictionary.Add osteo_headers(ItemOsteoDestiny), (ItemOsteoDestiny.Column - 1)
  Next ItemOsteoDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemOsteoOrigin In osteo_origin_header
    On Error GoTo osteoError
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
      If (typeExams(charters(ItemData.Offset(, osteo_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        With ActiveCell
          .Offset(, osteo_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, osteo_origin_dictionary("NRO IDENFICACION")))
          .Offset(, osteo_destiny_dictionary("CERVICALGIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("CERVICALGIA")))
          .Offset(, osteo_destiny_dictionary("CERVICALGIA OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("CERVICALGIA OBS")))
          .Offset(, osteo_destiny_dictionary("EPICONDILITIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("EPICONDILITIS")))
          .Offset(, osteo_destiny_dictionary("EPICONDILITIS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("EPICONDILITIS OBS")))
          .Offset(, osteo_destiny_dictionary("LUMBALGIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("LUMBALGIA")))
          .Offset(, osteo_destiny_dictionary("LUMBALGIA OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("LUMBALGIA OBS")))
          .Offset(, osteo_destiny_dictionary("S_ TUNEL CARPO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("S_ TUNEL CARPO")))
          .Offset(, osteo_destiny_dictionary("S_ TUNEL CARPO OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("S_ TUNEL CARPO OBS")))
          .Offset(, osteo_destiny_dictionary("FRACTURAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("FRACTURAS")))
          .Offset(, osteo_destiny_dictionary("FRACTURAS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("FRACTURAS OBS")))
          .Offset(, osteo_destiny_dictionary("TENDINITIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("TENDINITIS")))
          .Offset(, osteo_destiny_dictionary("TENDINITIS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("TENDINITIS OBS")))
          .Offset(, osteo_destiny_dictionary("LESION EN MENISCOS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("LESION EN MENISCOS OBS")))
          .Offset(, osteo_destiny_dictionary("LESION EN MENISCOS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("LESION EN MENISCOS")))
          .Offset(, osteo_destiny_dictionary("ESGUINCES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("ESGUINCES")))
          .Offset(, osteo_destiny_dictionary("ESGUINCES OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("ESGUINCES OBS")))
          .Offset(, osteo_destiny_dictionary("HOMBRO DOLOROSO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("HOMBRO DOLOROSO")))
          .Offset(, osteo_destiny_dictionary("HOMBRO DOLOROSO OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("HOMBRO DOLOROSO OBS")))
          .Offset(, osteo_destiny_dictionary("RADICULOPATIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("RADICULOPATIA")))
          .Offset(, osteo_destiny_dictionary("RADICULOPATIA OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("RADICULOPATIA OBS")))
          .Offset(, osteo_destiny_dictionary("BURSITIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("BURSITIS")))
          .Offset(, osteo_destiny_dictionary("BURSITIS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("BURSITIS OBS")))
          .Offset(, osteo_destiny_dictionary("ARTROSIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("ARTROSIS")))
          .Offset(, osteo_destiny_dictionary("ARTROSIS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("ARTROSIS OBS")))
          .Offset(, osteo_destiny_dictionary("ESCOLIOSIS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("ESCOLIOSIS")))
          .Offset(, osteo_destiny_dictionary("ESCOLIOSIS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("ESCOLIOSIS OBS")))
          .Offset(, osteo_destiny_dictionary("RETRACCIONES MUSCULARES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("RETRACCIONES MUSCULARES")))
          .Offset(, osteo_destiny_dictionary("RETRACCIONES MUSCULARES OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("RETRACCIONES MUSCULARES OBS")))
          .Offset(, osteo_destiny_dictionary("MALFORMACIONES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("MALFORMACIONES")))
          .Offset(, osteo_destiny_dictionary("MALFORMACIONES OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("MALFORMACIONES OBS")))
          .Offset(, osteo_destiny_dictionary("DISCOPATIAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DISCOPATIAS")))
          .Offset(, osteo_destiny_dictionary("DISCOPATIAS OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("DISCOPATIAS OBS")))
          .Offset(, osteo_destiny_dictionary("FIBROMALGIA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("FIBROMALGIA")))
          .Offset(, osteo_destiny_dictionary("FIBROMALGIA OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("FIBROMALGIA OBS")))
          .Offset(, osteo_destiny_dictionary("OTROS ANT_ OSTEOMUSCULARES")) = charters(ItemData.Offset(, osteo_origin_dictionary("OTROS ANT_ OSTEOMUSCULARES")))
          .Offset(, osteo_destiny_dictionary("PESO")) = charters(ItemData.Offset(, osteo_origin_dictionary("PESO")))
          .Offset(, osteo_destiny_dictionary("TALLA")) = charters(ItemData.Offset(, osteo_origin_dictionary("TALLA")))
          .Offset(, osteo_destiny_dictionary("DIAG_ PPAL")) = charters(ItemData.Offset(, osteo_origin_dictionary("DIAG_ PPAL")))
          .Offset(, osteo_destiny_dictionary("DIAG_ PPAL OBS")) = charters(ItemData.Offset(, osteo_origin_dictionary("DIAG_ PPAL OBS")))
          .Offset(, osteo_destiny_dictionary("DIAG_ REL 1")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DIAG_ REL 1")))
          .Offset(, osteo_destiny_dictionary("DIAG_ REL 2")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DIAG_ REL 2")))
          .Offset(, osteo_destiny_dictionary("DIAG_ REL 3")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("DIAG_ REL 3")))
          .Offset(, osteo_destiny_dictionary("REC/PERS ACT_ FISICA CARDIO 3X/SEMANA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS ACT_ FISICA CARDIO 3X/SEMANA")))
          .Offset(, osteo_destiny_dictionary("REC/PERS FORT_ 15 REPETICIONES/3 SERIES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS FORT_ 15 REPETICIONES/3 SERIES")))
          .Offset(, osteo_destiny_dictionary("REC/PERS EJERC_ ESTIRAMIENTO 20 SEG")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS EJERC_ ESTIRAMIENTO 20 SEG")))
          .Offset(, osteo_destiny_dictionary("REC/PERS AUTOCUIDADO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS AUTOCUIDADO")))
          .Offset(, osteo_destiny_dictionary("REC/PERS SEGUIMIENTO MEDICO")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/PERS SEGUIMIENTO MEDICO")))
          .Offset(, osteo_destiny_dictionary("REC/OCUP SVE PREVENSION LESIONES")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP SVE PREVENSION LESIONES")))
          .Offset(, osteo_destiny_dictionary("REC/OCUP MANIPULACION DE CARGA")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP MANIPULACION DE CARGA")))
          .Offset(, osteo_destiny_dictionary("REC/OCUP PAUSAS ACTIVAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP PAUSAS ACTIVAS")))
          .Offset(, osteo_destiny_dictionary("REC/OCUP ANALISIS ERGONOMICOS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP ANALISIS ERGONOMICOS")))
          .Offset(, osteo_destiny_dictionary("REC/OCUP EVITAR POSTURAS FORZADAS")) = charters_empty(ItemData.Offset(, osteo_origin_dictionary("REC/OCUP EVITAR POSTURAS FORZADAS")))
          .Offset(, osteo_destiny_dictionary("RECOM_ OCUPACIONALES")) = charters(ItemData.Offset(, osteo_origin_dictionary("RECOM_ OCUPACIONALES")))
          .Offset(, osteo_destiny_dictionary("RECOM_ G/RALES")) = charters(ItemData.Offset(, osteo_origin_dictionary("RECOM_ G/RALES")))
          If (.Row = 4) Then
            .Offset(, osteo_destiny_dictionary("ID_OSTEOMUSCULAR")) = Trim$(ThisWorkbook.Worksheets("RUTAS").Range("$F$11").value)
          Else
            .Offset(, osteo_destiny_dictionary("ID_OSTEOMUSCULAR")) = .Offset(-1, osteo_destiny_dictionary("ID_OSTEOMUSCULAR")) + 1
          End If
          .Offset(1, 0).Select
        End With
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData

    Call dataDuplicate("$A4")
    Call formatter("$A4")

    Set osteo_origin_value = Nothing
    Set osteo_destiny_header = Nothing
    Set osteo_origin_header = Nothing
    osteo_destiny_dictionary.RemoveAll
    osteo_origin_dictionary.RemoveAll

 osteoError:
    Resume Next
End Sub
