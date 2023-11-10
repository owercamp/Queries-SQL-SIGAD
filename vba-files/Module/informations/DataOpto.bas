Attribute VB_Name = "DataOpto"
'namespace=vba-files\Module\informations
Option Explicit

'TODO: OptoData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - opto_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - opto_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - opto_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - opto_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - opto_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Dim aumentFromIDOpto As LongPtr, aumentFromIDDiagnostic As LongPtr
Public Sub OptoData(ByVal name_sheet As String)
  Dim opto_destiny_dictionary As Scripting.Dictionary
  Dim opto_origin_dictionary As Scripting.Dictionary
  Dim opto_destiny_header As Object, opto_origin_header As Object, opto_origin_value As Object
  Dim ItemOptoDestiny As Object, ItemOptoOrigin As Object, ItemData As Object, opto_origin As Object

  Set opto_origin = origin.Worksheets(name_sheet) '' OPTO DEL LIBRO ORIGEN ''
  opto_destiny.Select
  opto_destiny.Range("$A4").Select
  Set opto_destiny_header = opto_destiny.Range("$A3", opto_destiny.Range("$A3").End(xlToRight))
  Set opto_origin_header = opto_origin.Range("$A1", opto_origin.Range("$A1").End(xlToRight))
  Set opto_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set opto_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (opto_origin.Range("$A2") <> Empty And opto_origin.Range("$A3") <> Empty) Then
    Set opto_origin_value = opto_origin.Range("$A2", opto_origin.Range("$A2").End(xlDown))
  ElseIf (opto_origin.Range("$A2") <> Empty And opto_origin.Range("$A3") = Empty) Then
    Set opto_origin_value = opto_origin.Range("$A2")
  End If

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  Dim value_data As String
  For Each ItemOptoDestiny In opto_destiny_header
    value_data = opto_headers(ItemOptoDestiny)
    If opto_destiny_dictionary.Exists(value_data) = False And value_data <> Empty Then
      opto_destiny_dictionary.Add value_data, (ItemOptoDestiny.Column - 1)
    End If
  Next ItemOptoDestiny
  
  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemOptoOrigin In opto_origin_header
    value_data = opto_headers(ItemOptoOrigin)
    If opto_origin_dictionary.Exists(value_data) = False And value_data <> Empty Then
      opto_origin_dictionary.Add value_data, (ItemOptoOrigin.Column - 1)
    End If
  Next ItemOptoOrigin

  numbers = 1
  porcentaje = 0
  
  aumentFromIDOpto = destiny.Worksheets("RUTAS").range("$F$7").value
  aumentFromIDDiagnostic = destiny.Worksheets("RUTAS").range("$F$8").value
  counts = opto_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  Dim type_exam As String
  With formImports
    For Each ItemData In opto_origin_value
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & opto_destiny.Name
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
      
      type_exam = typeExams(Trim(ItemData.Offset(, opto_origin_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        ActiveCell.Offset(, opto_destiny_dictionary("IDENTIFICACION")) = Trim(ItemData.Offset(, opto_origin_dictionary("IDENTIFICACION")))
        ActiveCell.Offset(, opto_destiny_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")))
        ActiveCell.Offset(, opto_destiny_dictionary("VISIO/ANT_ LABORAL USUARIO COMPUTADOR")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL USUARIO COMPUTADOR")))
        ActiveCell.Offset(, opto_destiny_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")))
        ActiveCell.Offset(, opto_destiny_dictionary("VISIO/ANT_ LABORAL CAMBIOS TEMPREATURA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL CAMBIOS TEMPREATURA")))
        ActiveCell.Offset(, opto_destiny_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")))
        ActiveCell.Offset(, opto_destiny_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS FOTOFOBIA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS FOTOFOBIA")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS OJO ROJO")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS OJO ROJO")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS LAGRIMEO")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS LAGRIMEO")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS VISION BORROSA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS VISION BORROSA")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS ARDOR")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS ARDOR")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS VISION DOBLE")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS VISION DOBLE")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS CANSANCIO")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS CANSANCIO")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS MALA VISION CERCANA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS MALA VISION CERCANA")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS DOLOR")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS DOLOR")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS MALA VISON LEJANA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS MALA VISON LEJANA")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS SECRECION")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS SECRECION")))
        ActiveCell.Offset(, opto_destiny_dictionary("SINTOMAS CEFALEA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS CEFALEA")))
        ActiveCell.Offset(, opto_destiny_dictionary("OTROS SINTOMAS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("OTROS SINTOMAS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("CABEZA - PARPADOS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("CABEZA - PARPADOS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("CABEZA - PARPADOS OBS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("CABEZA - PARPADOS OBS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("CABEZA - CONJUNTIVAS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("CABEZA - CONJUNTIVAS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("CABEZA - OBS CONJUNTIVAS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("CABEZA - OBS CONJUNTIVAS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("CABEZA - ESCLERAS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("CABEZA - ESCLERAS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("CABEZA - OBS ESCLERAS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("CABEZA - OBS ESCLERAS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("CABEZA - PUPILAS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("CABEZA - PUPILAS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("CABEZA - PUPILAS OBS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("CABEZA - PUPILAS OBS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("MOT/OCUL COVERT TEST LEJOS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("MOT/OCUL COVERT TEST LEJOS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("MOT/OCUL COVERT TEST CERCA")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("MOT/OCUL COVERT TEST CERCA"))))
        ActiveCell.Offset(, opto_destiny_dictionary("ESTADO DE CORRECCION")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("ESTADO DE CORRECCION"))))
        ActiveCell.Offset(, opto_destiny_dictionary("PATOLOGIA OCULAR")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("PATOLOGIA OCULAR"))))
        ActiveCell.Offset(, opto_destiny_dictionary("DIAG PPAL")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("DIAG PPAL"))))
        ActiveCell.Offset(, opto_destiny_dictionary("DIAG OBS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("DIAG OBS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("DIAG REL/1")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("DIAG REL/1"))))
        ActiveCell.Offset(, opto_destiny_dictionary("DIAG REL/2")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("DIAG REL/2"))))
        ActiveCell.Offset(, opto_destiny_dictionary("DIAG REL/3")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("DIAG REL/3"))))
        ActiveCell.Offset(, opto_destiny_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("REC CORRECCION VISUAL PARA TRABAJAR"))))
        ActiveCell.Offset(, opto_destiny_dictionary("REC USO AR VIDEO TRMINAL")) = Trim(ItemData.Offset(, opto_origin_dictionary("REC USO AR VIDEO TRMINAL")))
        ActiveCell.Offset(, opto_destiny_dictionary("REC USO DE LENTES DE PROTECCION SOLAR")) = Trim(ItemData.Offset(, opto_origin_dictionary("REC USO DE LENTES DE PROTECCION SOLAR")))
        ActiveCell.Offset(, opto_destiny_dictionary("REC USO EPP VISUAL")) = Trim(ItemData.Offset(, opto_origin_dictionary("REC USO EPP VISUAL")))
        ActiveCell.Offset(, opto_destiny_dictionary("REC PAUSAS ACTIVAS")) = Trim(ItemData.Offset(, opto_origin_dictionary("REC PAUSAS ACTIVAS")))
        ActiveCell.Offset(, opto_destiny_dictionary("REC USO RX VISION PROXIMA")) = Trim(ItemData.Offset(, opto_origin_dictionary("REC USO RX VISION PROXIMA")))
        ActiveCell.Offset(, opto_destiny_dictionary("REC USO RX DESCANSO")) = Trim(ItemData.Offset(, opto_origin_dictionary("REC USO RX DESCANSO")))
        ActiveCell.Offset(, opto_destiny_dictionary("REC USO PERMANENTE RX OPTICA")) = Trim(ItemData.Offset(, opto_origin_dictionary("REC USO PERMANENTE RX OPTICA")))
        ActiveCell.Offset(, opto_destiny_dictionary("REC PYP")) = Trim(ItemData.Offset(, opto_origin_dictionary("REC PYP")))
        ActiveCell.Offset(, opto_destiny_dictionary("REC LUBRICANTE OCULAR")) = Trim(ItemData.Offset(, opto_origin_dictionary("REC LUBRICANTE OCULAR")))
        ActiveCell.Offset(, opto_destiny_dictionary("RECOMENDACIONES OBS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("RECOMENDACIONES OBS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("REM_ VALORACION OFTALM_")) = Trim(ItemData.Offset(, opto_origin_dictionary("REM_ VALORACION OFTALM_")))
        ActiveCell.Offset(, opto_destiny_dictionary("REM_ TOPOGRAFIA CORNEAL")) = Trim(ItemData.Offset(, opto_origin_dictionary("REM_ TOPOGRAFIA CORNEAL")))
        ActiveCell.Offset(, opto_destiny_dictionary("REM_ TRATAM_ ORTOPTICA")) = Trim(ItemData.Offset(, opto_origin_dictionary("REM_ TRATAM_ ORTOPTICA")))
        ActiveCell.Offset(, opto_destiny_dictionary("REM_ TEST FARNSWORTH")) = Trim(ItemData.Offset(, opto_origin_dictionary("REM_ TEST FARNSWORTH")))
        ActiveCell.Offset(, opto_destiny_dictionary("REALIZAR PRUEBA AMBULATORIA")) = Trim(ItemData.Offset(, opto_origin_dictionary("REALIZAR PRUEBA AMBULATORIA")))
        ActiveCell.Offset(, opto_destiny_dictionary("REMISIONES OBS")) = Trim(UCase(ItemData.Offset(, opto_origin_dictionary("REMISIONES OBS"))))
        ActiveCell.Offset(, opto_destiny_dictionary("CONTROLES MENSUAL")) = Trim(ItemData.Offset(, opto_origin_dictionary("CONTROLES MENSUAL")))
        ActiveCell.Offset(, opto_destiny_dictionary("CONTROLES_BIMESTRALES")) = Trim(ItemData.Offset(, opto_origin_dictionary("CONTROLES_BIMESTRALES")))
        ActiveCell.Offset(, opto_destiny_dictionary("CONTROLES TRIMESTRAL")) = Trim(ItemData.Offset(, opto_origin_dictionary("CONTROLES TRIMESTRAL")))
        ActiveCell.Offset(, opto_destiny_dictionary("CONTROLES 6 MESES")) = Trim(ItemData.Offset(, opto_origin_dictionary("CONTROLES 6 MESES")))
        ActiveCell.Offset(, opto_destiny_dictionary("CONTROLES 1 ANO")) = Trim(ItemData.Offset(, opto_origin_dictionary("CONTROLES 1 ANO")))
        ActiveCell.Offset(, opto_destiny_dictionary("CONTROLES CONFIRMATORIA")) = Trim(ItemData.Offset(, opto_origin_dictionary("CONTROLES CONFIRMATORIA")))
        If (ActiveCell.Row <> 4) Then
          aumentFromIDOpto = aumentFromIDOpto + 1
          aumentFromIDDiagnostic = aumentFromIDDiagnostic + 1
        End If
        ActiveCell.Offset(, opto_destiny_dictionary("ID_OPTOMETRIA")) = aumentFromIDOpto
        ActiveCell.Offset(, opto_destiny_dictionary("OP_DIAGNOSTICO")) = aumentFromIDDiagnostic
        ActiveCell.Offset(1, 0).Select
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents
      End If
    Next ItemData
  End With

  Call dataDuplicate(opto_destiny.Range("tbl_opto[[#Data],[NRO IDENFICACION]]"))
  Call greaterThanOne(opto_destiny.Range("tbl_opto[[CONTROLES MENSUAL]:[CONTROLES CONFIRMATORIA]]"), "OPTO")
  Call iqualCero(opto_destiny.Range("tbl_opto[[CONTROLES MENSUAL]:[CONTROLES CONFIRMATORIA]]"), "OPTO")
  Call formatter(opto_destiny.Range("tbl_opto[[#Data],[NRO IDENFICACION]]"))

  Set opto_origin_value = Nothing
  Set opto_destiny_header = Nothing
  Set opto_origin_header = Nothing
  opto_destiny_dictionary.RemoveAll
  opto_origin_dictionary.RemoveAll

End Sub