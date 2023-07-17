Attribute VB_Name = "DataAudio"
'namespace=vba-files\Module\informations
Option Explicit

'TODO AudioData - Esta subrutina importa datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - audio_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - audio_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - audio_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - audio_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - audio_origin_value: Una variable de objeto para almacenar el rango de los datos de audio de la hoja de origen.
'* - ItemAudioDestiny: Una variable variante para iterar a traves del rango del encabezado de la hoja de destino.
'* - ItemAudioOrigin: Una variable variante para iterar a traves del rango del encabezado de la hoja de origen.
'* - ItemData: Una variable variante para iterar a traves del rango de datos de audio de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Dim audio_origin_dictionary As Scripting.Dictionary
Dim aumentFromID As LongPtr
Public Sub AudioData()
  Dim tbl_audio As Object, audio_origin_header As Object, audio_origin_value As Object
  Dim ItemAudioOrigin As Variant, ItemData As Variant

  Set audio_origin = origin.Worksheets("AUDIO") '' AUDIO DEL LIBRO ORIGEN ''
  audio_destiny.Select
  Set tbl_audio = ActiveSheet.ListObjects("tbl_audio")
  Set audio_origin_header = audio_origin.range("A1", audio_origin.range("A1").End(xlToRight))
  Set audio_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (audio_origin.range("A2") <> Empty And audio_origin.range("A3") <> Empty) Then
    Set audio_origin_value = audio_origin.range("A2", audio_origin.range("A2").End(xlDown))
  ElseIf (audio_origin.range("A2") <> Empty And audio_origin.range("A3") = Empty) Then
    Set audio_origin_value = audio_origin.range("A2")
  End If

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemAudioOrigin In audio_origin_header
    On Error Resume Next
    audio_origin_dictionary.Add audio_headers(ItemAudioOrigin), (ItemAudioOrigin.Column - 1)
    On Error GoTo 0
  Next ItemAudioOrigin

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$6").value
  counts = audio_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each ItemData In audio_origin_value
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & audio_destiny.Name
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

      If (typeExams(charters(ItemData.Offset(, audio_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        Select Case numbers
          Case 1
            Call addNewRegister(tbl_audio.ListRows(1), aumentFromID, ItemData)
          Case Else
            aumentFromID = aumentFromID + 1
            Call addNewRegister(tbl_audio.ListRows.Add, aumentFromID, ItemData)
        End Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData
  End With

  range("$A4").Select
  Call dataDuplicate
  range("$AT4:$AX4").Select
  Call greaterThanOne
  range("$AT4:$AX4").Select
  Call iqualCero
  range("$BF4").Select
  Call dataDuplicate
  range("$BG4").Select
  Call dataDuplicate
  range("$A4", range("$A4").End(xlDown)).Select
  Call formatter

  Set audio_origin_value = Nothing
  Set audio_origin_header = Nothing
  audio_origin_dictionary.RemoveAll

End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrement As LongPtr, ByVal ItemData As Variant)

  With table
    .Range(1) = charters(ItemData.Offset(, audio_origin_dictionary("NROAIDENFICACION")))
    .Range(2) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO")))
    .Range(3) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO COPA")))
    .Range(4) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO INSERCION")))
    .Range(5) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO DOBLE")))
    .Range(6) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO DER_")))
    .Range(7) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO DER_ OBS")))
    .Range(8) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO IZQ_")))
    .Range(9) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO IZQ_ OBS")))
    .Range(10) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO DER_")))
    .Range(11) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO DER_ OBS")))
    .Range(12) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO IZQ_")))
    .Range(13) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO IZQ_ OBS")))
    .Range(14) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO DER")))
    .Range(15) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO DER_ OBS")))
    .Range(16) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO IZQ")))
    .Range(17) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO IZQ_ OBS")))
    .Range(18) = charters(ItemData.Offset(, audio_origin_dictionary("TIPO DE EXAMEN")))
    .Range(19) = charters(ItemData.Offset(, audio_origin_dictionary("OD 500")))
    .Range(20) = charters(ItemData.Offset(, audio_origin_dictionary("OD 1000")))
    .Range(21) = charters(ItemData.Offset(, audio_origin_dictionary("OD 2000")))
    .Range(22) = charters(ItemData.Offset(, audio_origin_dictionary("OD 3000")))
    .Range(23) = charters(ItemData.Offset(, audio_origin_dictionary("OD 4000")))
    .Range(24) = charters(ItemData.Offset(, audio_origin_dictionary("OD 6000")))
    .Range(25) = charters(ItemData.Offset(, audio_origin_dictionary("OD 8000")))
    .Range(27) = charters(ItemData.Offset(, audio_origin_dictionary("OI 500")))
    .Range(28) = charters(ItemData.Offset(, audio_origin_dictionary("OI 1000")))
    .Range(29) = charters(ItemData.Offset(, audio_origin_dictionary("OI 2000")))
    .Range(30) = charters(ItemData.Offset(, audio_origin_dictionary("OI 3000")))
    .Range(31) = charters(ItemData.Offset(, audio_origin_dictionary("OI 4000")))
    .Range(32) = charters(ItemData.Offset(, audio_origin_dictionary("OI 6000")))
    .Range(33) = charters(ItemData.Offset(, audio_origin_dictionary("OI 8000")))
    .Range(35) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONTROL SEGUN PVE")))
    .Range(36) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONFIRMATORIA")))
    .Range(37) = charters_empty(ItemData.Offset(, audio_origin_dictionary("REMISION ORL")))
    .Range(38) = charters_empty(ItemData.Offset(, audio_origin_dictionary("PRUEBAS COMPLEMENTARIAS")))
    .Range(39) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA DE OIDO")))
    .Range(40) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA OD")))
    .Range(41) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA OI")))
    .Range(42) = charters_empty(ItemData.Offset(, audio_origin_dictionary("REPOSO AUDITIVO EXTRALAB")))
    .Range(43) = charters_empty(ItemData.Offset(, audio_origin_dictionary("ROTAR DIADEMA TELEFONICA")))
    .Range(44) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONDUCIR CON VENTANAS CERRADAS")))
    .Range(45) = charters_empty(ItemData.Offset(, audio_origin_dictionary("USO DE EPP AUDITIVO")))
    .Range(46) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES MENSUALES")))
    .Range(47) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES_BIMESTRALES")))
    .Range(48) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES TRIMESTRALES")))
    .Range(49) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES 6 MESES")))
    .Range(50) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES 1 ANO")))

    Select Case charters(ItemData.Offset(, audio_origin_dictionary("DIAG PPAL")))
      Case "NO REFIERE"
        .Range(51) = "#N/A"
      Case Else
        .Range(51) = charters(ItemData.Offset(, audio_origin_dictionary("DIAG PPAL")))
    End Select
    Select Case charters(ItemData.Offset(, audio_origin_dictionary("DIAG INTERNO")))
      Case "NO REFIERE"
        .Range(52) = "#N/A"
      Case Else
        .Range(52) = charters(ItemData.Offset(, audio_origin_dictionary("DIAG INTERNO")))
    End Select
    Select Case charters(ItemData.Offset(, audio_origin_dictionary("DIAG GATI-SO")))
      Case "NO REFIERE"
        .Range(53) = "#N/A"
      Case Else
        .Range(53) = charters(ItemData.Offset(, audio_origin_dictionary("DIAG GATI-SO")))
    End Select
    .Range(59) = autoIncrement
    DoEvents
  End With

End Sub