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
Dim aumentFromID As LongPtr
Public Sub AudioData(ByVal name_sheet As String)
  Dim audio_destiny_dictionary As Scripting.Dictionary
  Dim audio_origin_dictionary As Scripting.Dictionary
  Dim audio_destiny_header As Object, audio_origin_header As Object, audio_origin_value As Object
  Dim ItemAudioDestiny As Object, ItemAudioOrigin As Object, ItemData As Object, audio_origin As Object, cell_active as Range

  Set audio_origin = origin.Worksheets(name_sheet) '' AUDIO DEL LIBRO ORIGEN ''
  audio_destiny.Select
  audio_destiny.Range("$A4").Select
  Set cell_active = ActiveCell
  Set audio_destiny_header = audio_destiny.Range("$A3", audio_destiny.Range("$A3").End(xlToRight))
  Set audio_origin_header = audio_origin.Range("$A1", audio_origin.Range("$A1").End(xlToRight))
  Set audio_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set audio_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (audio_origin.Range("$A2") <> Empty And audio_origin.Range("$A3") <> Empty) Then
    Set audio_origin_value = audio_origin.Range("$A2", audio_origin.Range("$A2").End(xlDown))
  ElseIf (audio_origin.Range("$A2") <> Empty And audio_origin.Range("$A3") = Empty) Then
    Set audio_origin_value = audio_origin.Range("$A2")
  End If

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  Dim value_data As String
  For Each ItemAudioDestiny In audio_destiny_header
    value_data = audio_headers(ItemAudioDestiny)
    If audio_destiny_dictionary.Exists(value_data) = False And value_data <> Empty Then
      audio_destiny_dictionary.Add value_data, (ItemAudioDestiny.Column - 1)
    End If
  Next ItemAudioDestiny
  
  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemAudioOrigin In audio_origin_header
    value_data = audio_headers(ItemAudioOrigin)
    If audio_origin_dictionary.Exists(value_data) = False And value_data <> Empty Then
      audio_origin_dictionary.Add value_data, (ItemAudioOrigin.Column - 1)
    End If
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

  Dim type_exam As String
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

      type_exam = typeExams(Trim(ItemData.Offset(, audio_origin_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        cell_active.Offset(, audio_destiny_dictionary("NROAIDENFICACION")) = Trim(ItemData.Offset(, audio_origin_dictionary("NROAIDENFICACION")))

        search = ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO COPA"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO COPA")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO COPA")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO COPA")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO INSERCION"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO INSERCION")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO INSERCION")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO INSERCION")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("EPP ESPECIFICO / AUDITIVO DOBLE"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO DOBLE")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO DOBLE")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("EPP ESPECIFICO / AUDITIVO DOBLE")) = Trim$(search)
        End If

        cell_active.Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO DER_")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO DER_"))))
        cell_active.Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO DER_ OBS")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO DER_ OBS"))))
        cell_active.Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO IZQ_")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO IZQ_"))))
        cell_active.Offset(, audio_destiny_dictionary("PABELLON AURIC_ OIDO IZQ_ OBS")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("PABELLON AURIC_ OIDO IZQ_ OBS"))))
        cell_active.Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO DER_")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO DER_"))))
        cell_active.Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO DER_ OBS")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO DER_ OBS"))))
        cell_active.Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO IZQ_")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO IZQ_"))))
        cell_active.Offset(, audio_destiny_dictionary("CONDUCTO AUDIT_ OIDO IZQ_ OBS")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("CONDUCTO AUDIT_ OIDO IZQ_ OBS"))))
        cell_active.Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO DER")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO DER"))))
        cell_active.Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO DER_ OBS")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO DER_ OBS"))))
        cell_active.Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO IZQ")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO IZQ"))))
        cell_active.Offset(, audio_destiny_dictionary("MEMBRANA TIMP_ OIDO IZQ_ OBS")) = Trim(ItemData.Offset(, audio_origin_dictionary("MEMBRANA TIMP_ OIDO IZQ_ OBS")))
        cell_active.Offset(, audio_destiny_dictionary("TIPO DE EXAMEN")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("TIPO DE EXAMEN"))))
        cell_active.Offset(, audio_destiny_dictionary("OD 500")) = Trim(ItemData.Offset(, audio_origin_dictionary("OD 500")))
        cell_active.Offset(, audio_destiny_dictionary("OD 1000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OD 1000")))
        cell_active.Offset(, audio_destiny_dictionary("OD 2000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OD 2000")))
        cell_active.Offset(, audio_destiny_dictionary("OD 3000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OD 3000")))
        cell_active.Offset(, audio_destiny_dictionary("OD 4000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OD 4000")))
        cell_active.Offset(, audio_destiny_dictionary("OD 6000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OD 6000")))
        cell_active.Offset(, audio_destiny_dictionary("OD 8000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OD 8000")))
        cell_active.Offset(, audio_destiny_dictionary("OI 500")) = Trim(ItemData.Offset(, audio_origin_dictionary("OI 500")))
        cell_active.Offset(, audio_destiny_dictionary("OI 1000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OI 1000")))
        cell_active.Offset(, audio_destiny_dictionary("OI 2000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OI 2000")))
        cell_active.Offset(, audio_destiny_dictionary("OI 3000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OI 3000")))
        cell_active.Offset(, audio_destiny_dictionary("OI 4000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OI 4000")))
        cell_active.Offset(, audio_destiny_dictionary("OI 6000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OI 6000")))
        cell_active.Offset(, audio_destiny_dictionary("OI 8000")) = Trim(ItemData.Offset(, audio_origin_dictionary("OI 8000")))

        search = ItemData.Offset(, audio_origin_dictionary("CONTROL SEGUN PVE"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("CONTROL SEGUN PVE")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("CONTROL SEGUN PVE")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("CONTROL SEGUN PVE")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("CONFIRMATORIA"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("CONFIRMATORIA")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("CONFIRMATORIA")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("CONFIRMATORIA")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("REMISION ORL"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("REMISION ORL")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("REMISION ORL")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("REMISION ORL")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("PRUEBAS COMPLEMENTARIAS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("PRUEBAS COMPLEMENTARIAS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("PRUEBAS COMPLEMENTARIAS")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("PRUEBAS COMPLEMENTARIAS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("LIMPIEZA DE OIDO"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("LIMPIEZA DE OIDO")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("LIMPIEZA DE OIDO")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("LIMPIEZA DE OIDO")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("LIMPIEZA OD"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("LIMPIEZA OD")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("LIMPIEZA OD")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("LIMPIEZA OD")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("LIMPIEZA OI"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("LIMPIEZA OI")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("LIMPIEZA OI")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("LIMPIEZA OI")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("REPOSO AUDITIVO EXTRALAB"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("REPOSO AUDITIVO EXTRALAB")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("REPOSO AUDITIVO EXTRALAB")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("REPOSO AUDITIVO EXTRALAB")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("ROTAR DIADEMA TELEFONICA"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("ROTAR DIADEMA TELEFONICA")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("ROTAR DIADEMA TELEFONICA")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("ROTAR DIADEMA TELEFONICA")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("CONDUCIR CON VENTANAS CERRADAS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("CONDUCIR CON VENTANAS CERRADAS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("CONDUCIR CON VENTANAS CERRADAS")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("CONDUCIR CON VENTANAS CERRADAS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, audio_origin_dictionary("USO DE EPP AUDITIVO"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("USO DE EPP AUDITIVO")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, audio_destiny_dictionary("USO DE EPP AUDITIVO")) = 0
        Else
          cell_active.Offset(, audio_destiny_dictionary("USO DE EPP AUDITIVO")) = Trim$(search)
        End If

        cell_active.Offset(, audio_destiny_dictionary("CONTROLES MENSUALES")) = Trim(ItemData.Offset(, audio_origin_dictionary("CONTROLES MENSUALES")))
        cell_active.Offset(, audio_destiny_dictionary("CONTROLES_BIMESTRALES")) = Trim(ItemData.Offset(, audio_origin_dictionary("CONTROLES_BIMESTRALES")))
        cell_active.Offset(, audio_destiny_dictionary("CONTROLES TRIMESTRALES")) = Trim(ItemData.Offset(, audio_origin_dictionary("CONTROLES TRIMESTRALES")))
        cell_active.Offset(, audio_destiny_dictionary("CONTROLES 6 MESES")) = Trim(ItemData.Offset(, audio_origin_dictionary("CONTROLES 6 MESES")))
        cell_active.Offset(, audio_destiny_dictionary("CONTROLES 1 ANO")) = Trim(ItemData.Offset(, audio_origin_dictionary("CONTROLES 1 ANO")))
        If (Trim(ItemData.Offset(, audio_origin_dictionary("DIAG PPAL"))) = "NO REFIERE") Then
          cell_active.Offset(, audio_destiny_dictionary("DIAG PPAL")) = "#N/A"
        Else
          cell_active.Offset(, audio_destiny_dictionary("DIAG PPAL")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("DIAG PPAL"))))
        End If
        If (Trim(ItemData.Offset(, audio_origin_dictionary("DIAG INTERNO"))) = "NO REFIERE") Then
          cell_active.Offset(, audio_destiny_dictionary("DIAG INTERNO")) = "#N/A"
        Else
          cell_active.Offset(, audio_destiny_dictionary("DIAG INTERNO")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("DIAG INTERNO"))))
        End If
        If (Trim(ItemData.Offset(, audio_origin_dictionary("DIAG GATI-SO"))) = "NO REFIERE") Then
          cell_active.Offset(, audio_destiny_dictionary("DIAG GATI-SO")) ="#N/A"
        Else
          cell_active.Offset(, audio_destiny_dictionary("DIAG GATI-SO")) = Trim(UCase(ItemData.Offset(, audio_origin_dictionary("DIAG GATI-SO"))))
        End If
        If (cell_active.Row <> 4) Then
          aumentFromID = aumentFromID + 1
        End If
        cell_active.Offset(, audio_destiny_dictionary("ID_AUDIOMETRIA")) = aumentFromID
        Set cell_active = cell_active.Offset(1, 0)
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents     
      End If
    Next ItemData
  End With

  Call dataDuplicate(audio_destiny.Range("tbl_audio[[#Data],[NRO IDENTIFICACION]]"))
  Call greaterThanOne(audio_destiny.Range("tbl_audio[[CONTROLES MENSUALES]:[CONTROLES 1 A" & ChrW(209) & "O]]"),"AUDIO")
  Call iqualCero(audio_destiny.Range("tbl_audio[[CONTROLES MENSUALES]:[CONTROLES 1 A" & ChrW(209) & "O]]"), "AUDIO")
  Call formatter(audio_destiny.Range("tbl_audio[[#Data],[NRO IDENTIFICACION]]"))
  Call internalDiagnosis(audio_destiny.range("tbl_audio[[#Data],[DIAG INTERNO]]"))

  Set audio_origin_value = Nothing
  Set audio_destiny_header = Nothing
  Set audio_origin_header = Nothing
  audio_destiny_dictionary.RemoveAll
  audio_origin_dictionary.RemoveAll

End Sub