Attribute VB_Name = "DataAudio"
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
Public Sub AudioData()

  Dim audio_destiny_dictionary As Scripting.Dictionary
  Dim audio_origin_dictionary As Scripting.Dictionary
  Dim audio_destiny_header As Object, audio_origin_header As Object, audio_origin_value As Object
  Dim ItemAudioDestiny As Variant, ItemAudioOrigin As Variant, ItemData As Variant, currenCell As range
  Dim aumentFromRow As LongPtr, aumentFromID As LongPtr

  Set audio_origin = origin.Worksheets("AUDIO") '' AUDIO DEL LIBRO ORIGEN ''
  audio_destiny.Select
  ActiveSheet.range("A4").Select
  Set currenCell = ActiveCell
  Set audio_destiny_header = audio_destiny.range("A3", audio_destiny.range("A3").End(xlToRight))
  Set audio_origin_header = audio_origin.range("A1", audio_origin.range("A1").End(xlToRight))
  Set audio_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set audio_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (audio_origin.range("A2") <> Empty And audio_origin.range("A3") <> Empty) Then
    Set audio_origin_value = audio_origin.range("A2", audio_origin.range("A2").End(xlDown))
  ElseIf (audio_origin.range("A2") <> Empty And audio_origin.range("A3") = Empty) Then
    Set audio_origin_value = audio_origin.range("A2")
  End If

  ''   En los diccionarios de "audio_destiny_dictionary" y  "audio_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemAudioDestiny In audio_destiny_header
    On Error Resume Next
    audio_destiny_dictionary.Add audio_headers(ItemAudioDestiny), (ItemAudioDestiny.Column - 1)
    On Error GoTo 0
  Next ItemAudioDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemAudioOrigin In audio_origin_header
    On Error Resume Next
    audio_origin_dictionary.Add audio_headers(ItemAudioOrigin), (ItemAudioOrigin.Column - 1)
    On Error GoTo 0
  Next ItemAudioOrigin

  numbers = 1
  porcentaje = 0
  aumentFromRow = 0
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
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("NROAIDENFICACION")) = charters(ItemData.Offset(, audio_origin_dictionary("NROAIDENFICACION")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO COPA")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO COPA")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO INSERCION")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO INSERCION")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO DOBLE")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO DOBLE")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("PABELLON AURIC_ OIDO DER_")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO DER_")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("PABELLON AURIC_ OIDO DER_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO DER_ OBS")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("PABELLON AURIC_ OIDO IZQ_")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO IZQ_")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("PABELLON AURIC_ OIDO IZQ_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO IZQ_ OBS")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO DER_")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO DER_")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO DER_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO DER_ OBS")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO IZQ_")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO IZQ_")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO IZQ_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO IZQ_ OBS")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO DER")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO DER")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO DER_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO DER_ OBS")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO IZQ")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO IZQ")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO IZQ_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO IZQ_ OBS")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("TIPO DE EXAMEN")) = charters(ItemData.Offset(, audio_origin_dictionary("TIPO DE EXAMEN")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OD 500")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 500")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OD 1000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 1000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OD 2000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 2000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OD 3000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 3000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OD 4000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 4000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OD 6000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 6000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OD 8000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 8000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OI 500")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 500")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OI 1000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 1000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OI 2000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 2000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OI 3000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 3000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OI 4000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 4000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OI 6000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 6000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("OI 8000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 8000")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONTROL SEGUN PVE")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONTROL SEGUN PVE")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONFIRMATORIA")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONFIRMATORIA")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("REMISION ORL")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("REMISION ORL")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("PRUEBAS COMPLEMENTARIAS")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("PRUEBAS COMPLEMENTARIAS")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("LIMPIEZA DE OIDO")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA DE OIDO")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("LIMPIEZA OD")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA OD")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("LIMPIEZA OI")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA OI")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("REPOSO AUDITIVO EXTRALAB")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("REPOSO AUDITIVO EXTRALAB")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("ROTAR DIADEMA TELEFONICA")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("ROTAR DIADEMA TELEFONICA")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONDUCIR CON VENTANAS CERRADAS")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONDUCIR CON VENTANAS CERRADAS")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("USO DE EPP AUDITIVO")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("USO DE EPP AUDITIVO")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONTROLES MENSUALES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES MENSUALES")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES_BIMESTRALES")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONTROLES TRIMESTRALES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES TRIMESTRALES")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONTROLES 6 MESES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES 6 MESES")))
        currenCell.Offset(aumentFromRow, audio_destiny_dictionary("CONTROLES 1 ANO")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES 1 ANO")))
        If (charters(ItemData.Offset(, audio_origin_dictionary("DIAG PPAL"))) = "NO REFIERE") Then
          currenCell.Offset(aumentFromRow, audio_destiny_dictionary("DIAG PPAL")) = "#N/A"
        Else
          currenCell.Offset(aumentFromRow, audio_destiny_dictionary("DIAG PPAL")) = charters(ItemData.Offset(, audio_origin_dictionary("DIAG PPAL")))
        End If
        If (charters(ItemData.Offset(, audio_origin_dictionary("DIAG INTERNO"))) = "NO REFIERE") Then
          currenCell.Offset(aumentFromRow, audio_destiny_dictionary("DIAG INTERNO")) = "#N/A"
        Else
          currenCell.Offset(aumentFromRow, audio_destiny_dictionary("DIAG INTERNO")) = charters(ItemData.Offset(, audio_origin_dictionary("DIAG INTERNO")))
        End If
        If (charters(ItemData.Offset(, audio_origin_dictionary("DIAG GATI-SO"))) = "NO REFIERE") Then
          currenCell.Offset(aumentFromRow, audio_destiny_dictionary("DIAG GATI-SO")) = "#N/A"
        Else
          currenCell.Offset(aumentFromRow, audio_destiny_dictionary("DIAG GATI-SO")) = charters(ItemData.Offset(, audio_origin_dictionary("DIAG GATI-SO")))
        End If
        If (currenCell.Offset(aumentFromRow, 0).row = 4) Then
          currenCell.Offset(aumentFromRow, audio_destiny_dictionary("ID_AUDIOMETRIA")) = Trim(aumentFromID)
        Else
          aumentFromID = aumentFromID + 1
          currenCell.Offset(aumentFromRow, audio_destiny_dictionary("ID_AUDIOMETRIA")) = Trim(aumentFromID)
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
  Set audio_destiny_header = Nothing
  Set audio_origin_header = Nothing
  audio_destiny_dictionary.RemoveAll
  audio_origin_dictionary.RemoveAll

End Sub
