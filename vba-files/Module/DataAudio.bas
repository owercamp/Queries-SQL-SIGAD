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
        With ActiveCell
          .Offset(, audio_destiny_dictionary("NROAIDENFICACION")) = charters(ItemData.Offset(, audio_origin_dictionary("NROAIDENFICACION")))
          .Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO")))
          .Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO COPA")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO COPA")))
          .Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO INSERCION")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO INSERCION")))
          .Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO DOBLE")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO DOBLE")))
          .Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO DER_")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO DER_")))
          .Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO DER_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO DER_ OBS")))
          .Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO IZQ_")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO IZQ_")))
          .Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO IZQ_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO IZQ_ OBS")))
          .Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO DER_")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO DER_")))
          .Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO DER_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO DER_ OBS")))
          .Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO IZQ_")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO IZQ_")))
          .Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO IZQ_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO IZQ_ OBS")))
          .Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO DER")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO DER")))
          .Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO DER_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO DER_ OBS")))
          .Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO IZQ")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO IZQ")))
          .Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO IZQ_ OBS")) = charters(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO IZQ_ OBS")))
          .Offset(, audio_destiny_dictionary("TIPO DE EXAMEN")) = charters(ItemData.Offset(, audio_origin_dictionary("TIPO DE EXAMEN")))
          .Offset(, audio_destiny_dictionary("OD 500")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 500")))
          .Offset(, audio_destiny_dictionary("OD 1000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 1000")))
          .Offset(, audio_destiny_dictionary("OD 2000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 2000")))
          .Offset(, audio_destiny_dictionary("OD 3000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 3000")))
          .Offset(, audio_destiny_dictionary("OD 4000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 4000")))
          .Offset(, audio_destiny_dictionary("OD 6000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 6000")))
          .Offset(, audio_destiny_dictionary("OD 8000")) = charters(ItemData.Offset(, audio_origin_dictionary("OD 8000")))
          .Offset(, audio_destiny_dictionary("OI 500")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 500")))
          .Offset(, audio_destiny_dictionary("OI 1000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 1000")))
          .Offset(, audio_destiny_dictionary("OI 2000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 2000")))
          .Offset(, audio_destiny_dictionary("OI 3000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 3000")))
          .Offset(, audio_destiny_dictionary("OI 4000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 4000")))
          .Offset(, audio_destiny_dictionary("OI 6000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 6000")))
          .Offset(, audio_destiny_dictionary("OI 8000")) = charters(ItemData.Offset(, audio_origin_dictionary("OI 8000")))
          .Offset(, audio_destiny_dictionary("CONTROL SEGUN PVE")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONTROL SEGUN PVE")))
          .Offset(, audio_destiny_dictionary("CONFIRMATORIA")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONFIRMATORIA")))
          .Offset(, audio_destiny_dictionary("REMISION ORL")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("REMISION ORL")))
          .Offset(, audio_destiny_dictionary("PRUEBAS COMPLEMENTARIAS")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("PRUEBAS COMPLEMENTARIAS")))
          .Offset(, audio_destiny_dictionary("LIMPIEZA DE OIDO")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA DE OIDO")))
          .Offset(, audio_destiny_dictionary("LIMPIEZA OD")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA OD")))
          .Offset(, audio_destiny_dictionary("LIMPIEZA OI")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("LIMPIEZA OI")))
          .Offset(, audio_destiny_dictionary("REPOSO AUDITIVO EXTRALAB")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("REPOSO AUDITIVO EXTRALAB")))
          .Offset(, audio_destiny_dictionary("ROTAR DIADEMA TELEFONICA")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("ROTAR DIADEMA TELEFONICA")))
          .Offset(, audio_destiny_dictionary("CONDUCIR CON VENTANAS CERRADAS")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("CONDUCIR CON VENTANAS CERRADAS")))
          .Offset(, audio_destiny_dictionary("USO DE EPP AUDITIVO")) = charters_empty(ItemData.Offset(, audio_origin_dictionary("USO DE EPP AUDITIVO")))
          .Offset(, audio_destiny_dictionary("CONTROLES MENSUALES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES MENSUALES")))
          .Offset(, audio_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES_BIMESTRALES")))
          .Offset(, audio_destiny_dictionary("CONTROLES TRIMESTRALES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES TRIMESTRALES")))
          .Offset(, audio_destiny_dictionary("CONTROLES 6 MESES")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES 6 MESES")))
          .Offset(, audio_destiny_dictionary("CONTROLES 1 ANO")) = charters(ItemData.Offset(, audio_origin_dictionary("CONTROLES 1 ANO")))
          If (charters(ItemData.Offset(, audio_origin_dictionary("DIAG PPAL"))) = "NO REFIERE") Then
            .Offset(, audio_destiny_dictionary("DIAG PPAL")) = "#N/A"
          Else
            .Offset(, audio_destiny_dictionary("DIAG PPAL")) = charters(ReplaceNonAlphaNumeric(ItemData.Offset(, audio_origin_dictionary("DIAG PPAL"))))
          End If
          If (charters(ItemData.Offset(, audio_origin_dictionary("DIAG INTERNO"))) = "NO REFIERE") Then
            .Offset(, audio_destiny_dictionary("DIAG INTERNO")) = "#N/A"
          Else
            .Offset(, audio_destiny_dictionary("DIAG INTERNO")) = charters(ReplaceNonAlphaNumeric(ItemData.Offset(, audio_origin_dictionary("DIAG INTERNO"))))
          End If
          If (charters(ItemData.Offset(, audio_origin_dictionary("DIAG GATI-SO"))) = "NO REFIERE") Then
            .Offset(, audio_destiny_dictionary("DIAG GATI-SO")) ="#N/A"
          Else
            .Offset(, audio_destiny_dictionary("DIAG GATI-SO")) = charters(ReplaceNonAlphaNumeric(ItemData.Offset(, audio_origin_dictionary("DIAG GATI-SO"))))
          End If
          If (.Row = 4) Then
            .Offset(, audio_destiny_dictionary("ID_AUDIOMETRIA")) = Trim$(ThisWorkbook.Worksheets("RUTAS").Range("$F$6").value)
          Else
            .Offset(, audio_destiny_dictionary("ID_AUDIOMETRIA")) = .Offset(-1, audio_destiny_dictionary("ID_AUDIOMETRIA")) + 1
          End If
          .Offset(1, 0).Select
        End With
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData

    Call dataDuplicate("$A4")
    Call greaterThanOne("$AT4:$AX4")
    Call iqualCero("$AT4:$AX4")
    Call dataDuplicate("$BF4")
    Call dataDuplicate("$BG4")
    Call formatter("$A4")

    Set audio_origin_value = Nothing
    Set audio_destiny_header = Nothing
    Set audio_origin_header = Nothing
    audio_destiny_dictionary.RemoveAll
    audio_origin_dictionary.RemoveAll

    Exit Sub

 audioError:
    Resume Next
End Sub
