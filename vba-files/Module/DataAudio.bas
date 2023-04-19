Attribute VB_Name = "DataAudio"
Option Explicit

Sub AudioData()

  Dim audio_destiny_dictionary As Scripting.Dictionary
  Dim audio_origin_dictionary As Scripting.Dictionary
  Dim audio_destiny_header As Object, audio_origin_header As Object, audio_origin_value As Object
  Dim ItemAudioDestiny As Variant, ItemAudioOrigin As Variant, ItemData As Variant

  Set audio_origin = origin.Worksheets("AUDIO") '' AUDIO DEL LIBRO ORIGEN ''
  audio_destiny.Select
  ActiveSheet.Range("A4").Select
  Set audio_destiny_header = audio_destiny.Range("A3", audio_destiny.Range("A3").End(xlToRight))
  Set audio_origin_header = audio_origin.Range("A1", audio_origin.Range("A1").End(xlToRight))
  Set audio_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set audio_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (audio_origin.Range("A2") <> Empty And audio_origin.Range("A3") <> Empty) Then
    Set audio_origin_value = audio_origin.Range("A2", audio_origin.Range("A2").End(xlDown))
  ElseIf (audio_origin.Range("A2") <> Empty And audio_origin.Range("A3") = Empty) Then
    Set audio_origin_value = audio_origin.Range("A2")
  End If

  ''   En los diccionarios de "audio_destiny_dictionary" y  "audio_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemAudioDestiny In audio_destiny_header
    On Error GoTo audioError
    audio_destiny_dictionary.Add audio_headers(ItemAudioDestiny), (ItemAudioDestiny.Column - 1)
  Next ItemAudioDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemAudioOrigin In audio_origin_header
    On Error GoTo audioError
    audio_origin_dictionary.Add audio_headers(ItemAudioOrigin), (ItemAudioOrigin.Column - 1)
  Next ItemAudioOrigin

  numbers = 1
  porcentaje = 0
  counts = audio_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  For Each ItemData In audio_origin_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & audio_destiny.Name
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
      If (typeExams(charters(ItemData.Offset(, audio_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        ActiveCell.Offset(, audio_destiny_dictionary("NROAIDENFICACION")) = charters(ItemData.Offset(, audio_origin_dictionary("NROAIDENFICACION")))
        ActiveCell.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO")))
        ActiveCell.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO COPA")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO COPA")))
        ActiveCell.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO INSERCION")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO INSERCION")))
        ActiveCell.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO DOBLE")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO DOBLE")))
        ActiveCell.Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO DER_")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO DER_")))
        ActiveCell.Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO DER_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO DER_ OBS")))
        ActiveCell.Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO IZQ_")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO IZQ_")))
        ActiveCell.Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO IZQ_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO IZQ_ OBS")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO DER_")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO DER_")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO DER_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO DER_ OBS")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO IZQ_")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO IZQ_")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO IZQ_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO IZQ_ OBS")))
        ActiveCell.Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO DER")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO DER")))
        ActiveCell.Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO DER_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO DER_ OBS")))
        ActiveCell.Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO IZQ")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO IZQ")))
        ActiveCell.Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO IZQ_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO IZQ_ OBS")))
        ActiveCell.Offset(, audio_destiny_dictionary("TIPO DE EXAMEN")) = charters(ItemData.Offset(, audio_origin_dictionary("TIPO DE EXAMEN")))
        ActiveCell.Offset(, audio_destiny_dictionary("OD 500")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 500")))
        ActiveCell.Offset(, audio_destiny_dictionary("OD 1000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 1000")))
        ActiveCell.Offset(, audio_destiny_dictionary("OD 2000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 2000")))
        ActiveCell.Offset(, audio_destiny_dictionary("OD 3000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 3000")))
        ActiveCell.Offset(, audio_destiny_dictionary("OD 4000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 4000")))
        ActiveCell.Offset(, audio_destiny_dictionary("OD 6000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 6000")))
        ActiveCell.Offset(, audio_destiny_dictionary("OD 8000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 8000")))
        ActiveCell.Offset(, audio_destiny_dictionary("OI 500")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 500")))
        ActiveCell.Offset(, audio_destiny_dictionary("OI 1000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 1000")))
        ActiveCell.Offset(, audio_destiny_dictionary("OI 2000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 2000")))
        ActiveCell.Offset(, audio_destiny_dictionary("OI 3000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 3000")))
        ActiveCell.Offset(, audio_destiny_dictionary("OI 4000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 4000")))
        ActiveCell.Offset(, audio_destiny_dictionary("OI 6000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 6000")))
        ActiveCell.Offset(, audio_destiny_dictionary("OI 8000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 8000")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONTROL SEGUN PVE")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONTROL SEGUN PVE")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONFIRMATORIA")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONFIRMATORIA")))
        ActiveCell.Offset(, audio_destiny_dictionary("REMISION ORL")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("REMISION ORL")))
        ActiveCell.Offset(, audio_destiny_dictionary("PRUEBAS COMPLEMENTARIAS")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("PRUEBAS COMPLEMENTARIAS")))
        ActiveCell.Offset(, audio_destiny_dictionary("LIMPIEZA DE OIDO")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA DE OIDO")))
        ActiveCell.Offset(, audio_destiny_dictionary("LIMPIEZA OD")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA OD")))
        ActiveCell.Offset(, audio_destiny_dictionary("LIMPIEZA OI")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA OI")))
        ActiveCell.Offset(, audio_destiny_dictionary("REPOSO AUDITIVO EXTRALAB")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("REPOSO AUDITIVO EXTRALAB")))
        ActiveCell.Offset(, audio_destiny_dictionary("ROTAR DIADEMA TELEFONICA")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("ROTAR DIADEMA TELEFONICA")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONDUCIR CON VENTANAS CERRADAS")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONDUCIR CON VENTANAS CERRADAS")))
        ActiveCell.Offset(, audio_destiny_dictionary("USO DE EPP AUDITIVO")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("USO DE EPP AUDITIVO")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONTROLES MENSUALES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES MENSUALES")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES_BIMESTRALES")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONTROLES TRIMESTRALES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES TRIMESTRALES")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONTROLES 6 MESES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES 6 MESES")))
        ActiveCell.Offset(, audio_destiny_dictionary("CONTROLES 1 ANO")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES 1 ANO")))
        If (charters(ItemData.Offset(, audio_origin_dictionary("DIAG PPAL"))) = "NO REFIERE") Then
          ActiveCell.Offset(, audio_destiny_dictionary("DIAG PPAL")) = "#N/A"
        Else
          ActiveCell.Offset(, audio_destiny_dictionary("DIAG PPAL")) = charters(ItemData.Offset(, audio_origin_dictionary("DIAG PPAL")))
        End If
        If (charters(ItemData.Offset(, audio_origin_dictionary("DIAG INTERNO"))) = "NO REFIERE") Then
          ActiveCell.Offset(, audio_destiny_dictionary("DIAG INTERNO")) = "#N/A"
        Else
          ActiveCell.Offset(, audio_destiny_dictionary("DIAG INTERNO")) = charters(ItemData.Offset(, audio_origin_dictionary("DIAG INTERNO")))
        End If
        If (charters(ItemData.Offset(, audio_origin_dictionary("DIAG GATI-SO"))) = "NO REFIERE") Then
          ActiveCell.Offset(, audio_destiny_dictionary("DIAG GATI-SO")) ="#N/A"
        Else
          ActiveCell.Offset(, audio_destiny_dictionary("DIAG GATI-SO")) = charters(ItemData.Offset(, audio_origin_dictionary("DIAG GATI-SO")))
        End If
        If (ActiveCell.Row = 4) Then
          ActiveCell.Offset(, audio_destiny_dictionary("ID_AUDIOMETRIA")) = Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$6").value)
        Else
          ActiveCell.Offset(, audio_destiny_dictionary("ID_AUDIOMETRIA")) = ActiveCell.Offset(-1, audio_destiny_dictionary("ID_AUDIOMETRIA")) + 1
        End If
        ActiveCell.Offset(1, 0).Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData

    Range("$A4").Select
    Call dataDuplicate
    Range("$AT4:$AX4").Select
    Call greaterThanOne
    Range("$AT4:$AX4").Select
    Call iqualCero
    Range("$BF4").Select
    Call dataDuplicate
    Range("$BG4").Select
    Call dataDuplicate
    Range("$A4", Range("$A4").End(xlDown)).Select
    Call formatter

    Set audio_origin_value = Nothing
    Set audio_destiny_header = Nothing
    Set audio_origin_header = Nothing
    audio_destiny_dictionary.RemoveAll
    audio_origin_dictionary.RemoveAll

 audioError:
    Resume Next
End Sub
