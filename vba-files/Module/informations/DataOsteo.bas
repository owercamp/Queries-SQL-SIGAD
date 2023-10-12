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
Dim osteo_origin_dictionary As Scripting.Dictionary
Dim aumentFromID As LongPtr
Public Sub OsteoData()
  Dim tbl_osteo As Object, osteo_origin As Object
  
  Set osteo_origin = origin.Worksheets("OSTEO").Range("A1") '' OSTEO DEL LIBRO ORIGEN ''
  
  Set tbl_osteo = osteo_destiny.ListObjects("tbl_osteo")
  Set osteo_origin_dictionary = CreateObject("Scripting.Dictionary")

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each item In Range(osteo_origin, osteo_origin.End(xlToRight))
    If osteo_origin_dictionary.Exists(osteo_headers(item)) = False Then
      osteo_origin_dictionary.Add osteo_headers(item), item.Column
    End If
  Next item

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$11").value
  counts = Ubound(origin.Worksheets("OSTEO").Range("A1").CurrentRegion.Value, 1) - 1
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each item In Range(osteo_origin.offset(1, 0), osteo_origin.offset(1, 0).End(xlDown))
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

      If (typeExams(charters(item.Offset(, osteo_origin_dictionary("TIPO EXAMEN") - 1))) <> "EGRESO") Then
        If item.value <> "" And item.Row = 2 Then
          Call addNewRegister(tbl_osteo.ListRows(1), aumentFromID, item)
          DoEvents
        ElseIf item.value <> "" And item.Row > 2 Then
          aumentFromID = aumentFromID + 1
          Call addNewRegister(tbl_osteo.ListRows.Add, aumentFromID, item)
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
  range("$A4", range("$A4").End(xlDown)).Select
  Call formatter

  Set osteo_origin = Nothing
  osteo_origin_dictionary.RemoveAll

End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrement As LongPtr, ByVal information As Object)

  With table
    .Range(1) = charters(information(, osteo_origin_dictionary("NRO IDENFICACION")))
    .Range(2) = charters_empty(information(, osteo_origin_dictionary("CERVICALGIA")))
    .Range(3) = charters(information(, osteo_origin_dictionary("CERVICALGIA OBS")))
    .Range(4) = charters_empty(information(, osteo_origin_dictionary("EPICONDILITIS")))
    .Range(5) = charters(information(, osteo_origin_dictionary("EPICONDILITIS OBS")))
    .Range(6) = charters_empty(information(, osteo_origin_dictionary("LUMBALGIA")))
    .Range(7) = charters(information(, osteo_origin_dictionary("LUMBALGIA OBS")))
    .Range(8) = charters_empty(information(, osteo_origin_dictionary("S_ TUNEL CARPO")))
    .Range(9) = charters(information(, osteo_origin_dictionary("S_ TUNEL CARPO OBS")))
    .Range(10) = charters_empty(information(, osteo_origin_dictionary("FRACTURAS")))
    .Range(11) = charters(information(, osteo_origin_dictionary("FRACTURAS OBS")))
    .Range(12) = charters_empty(information(, osteo_origin_dictionary("TENDINITIS")))
    .Range(13) = charters(information(, osteo_origin_dictionary("TENDINITIS OBS")))
    .Range(14) = charters(information(, osteo_origin_dictionary("LESION EN MENISCOS OBS")))
    .Range(15) = charters_empty(information(, osteo_origin_dictionary("LESION EN MENISCOS")))
    .Range(16) = charters_empty(information(, osteo_origin_dictionary("ESGUINCES")))
    .Range(17) = charters(information(, osteo_origin_dictionary("ESGUINCES OBS")))
    .Range(18) = charters_empty(information(, osteo_origin_dictionary("HOMBRO DOLOROSO")))
    .Range(19) = charters(information(, osteo_origin_dictionary("HOMBRO DOLOROSO OBS")))
    .Range(20) = charters_empty(information(, osteo_origin_dictionary("RADICULOPATIA")))
    .Range(21) = charters(information(, osteo_origin_dictionary("RADICULOPATIA OBS")))
    .Range(22) = charters_empty(information(, osteo_origin_dictionary("BURSITIS")))
    .Range(23) = charters(information(, osteo_origin_dictionary("BURSITIS OBS")))
    .Range(24) = charters_empty(information(, osteo_origin_dictionary("ARTROSIS")))
    .Range(25) = charters(information(, osteo_origin_dictionary("ARTROSIS OBS")))
    .Range(26) = charters_empty(information(, osteo_origin_dictionary("ESCOLIOSIS")))
    .Range(27) = charters(information(, osteo_origin_dictionary("ESCOLIOSIS OBS")))
    .Range(28) = charters_empty(information(, osteo_origin_dictionary("RETRACCIONES MUSCULARES")))
    .Range(29) = charters(information(, osteo_origin_dictionary("RETRACCIONES MUSCULARES OBS")))
    .Range(30) = charters_empty(information(, osteo_origin_dictionary("MALFORMACIONES")))
    .Range(31) = charters(information(, osteo_origin_dictionary("MALFORMACIONES OBS")))
    .Range(32) = charters_empty(information(, osteo_origin_dictionary("DISCOPATIAS")))
    .Range(33) = charters(information(, osteo_origin_dictionary("DISCOPATIAS OBS")))
    .Range(34) = charters_empty(information(, osteo_origin_dictionary("FIBROMALGIA")))
    .Range(35) = charters(information(, osteo_origin_dictionary("FIBROMALGIA OBS")))
    .Range(36) = charters(information(, osteo_origin_dictionary("OTROS ANT_ OSTEOMUSCULARES")))
    .Range(37) = charters(information(, osteo_origin_dictionary("PESO")))
    .Range(38) = charters(information(, osteo_origin_dictionary("TALLA")))
    .Range(41) = charters(information(, osteo_origin_dictionary("DIAG_ PPAL")))
    .Range(42) = charters(information(, osteo_origin_dictionary("DIAG_ PPAL OBS")))
    .Range(43) = charters_empty(information(, osteo_origin_dictionary("DIAG_ REL 1")))
    .Range(44) = charters_empty(information(, osteo_origin_dictionary("DIAG_ REL 2")))
    .Range(45) = charters_empty(information(, osteo_origin_dictionary("DIAG_ REL 3")))
    .Range(46) = charters_empty(information(, osteo_origin_dictionary("REC/PERS ACT_ FISICA CARDIO 3X/SEMANA")))
    .Range(47) = charters_empty(information(, osteo_origin_dictionary("REC/PERS FORT_ 15 REPETICIONES/3 SERIES")))
    .Range(48) = charters_empty(information(, osteo_origin_dictionary("REC/PERS EJERC_ ESTIRAMIENTO 20 SEG")))
    .Range(49) = charters_empty(information(, osteo_origin_dictionary("REC/PERS AUTOCUIDADO")))
    .Range(50) = charters_empty(information(, osteo_origin_dictionary("REC/PERS SEGUIMIENTO MEDICO")))
    .Range(51) = charters_empty(information(, osteo_origin_dictionary("REC/OCUP SVE PREVENSION LESIONES")))
    .Range(52) = charters_empty(information(, osteo_origin_dictionary("REC/OCUP MANIPULACION DE CARGA")))
    .Range(53) = charters_empty(information(, osteo_origin_dictionary("REC/OCUP PAUSAS ACTIVAS")))
    .Range(54) = charters_empty(information(, osteo_origin_dictionary("REC/OCUP ANALISIS ERGONOMICOS")))
    .Range(55) = charters_empty(information(, osteo_origin_dictionary("REC/OCUP EVITAR POSTURAS FORZADAS")))
    .Range(56) = charters(information(, osteo_origin_dictionary("RECOM_ OCUPACIONALES")))
    .Range(57) = charters(information(, osteo_origin_dictionary("RECOM_ G/RALES")))
    .Range(59) = autoIncrement
  End With

End Sub