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
Public Sub OptoData()

  Dim opto_destiny_dictionary As Scripting.Dictionary
  Dim opto_origin_dictionary As Scripting.Dictionary
  Dim opto_destiny_header As Object, opto_origin_header As Object, opto_origin_value As Object
  Dim ItemOptoDestiny As Variant, ItemOptoOrigin As Variant, ItemData As Variant
  Dim currenCell As range, aumentFromRow As LongPtr, aumentFromIDOpto As LongPtr, aumentFromIDDiagnostic As LongPtr
  
  Set opto_origin = origin.Worksheets("OPTO") '' OPTO DEL LIBRO ORIGEN ''
  opto_destiny.Select
  ActiveSheet.range("A4").Select
  Set currenCell = ActiveCell
  Set opto_destiny_header = opto_destiny.range("A3", opto_destiny.range("A3").End(xlToRight))
  Set opto_origin_header = opto_origin.range("A1", opto_origin.range("A1").End(xlToRight))
  Set opto_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set opto_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (opto_origin.range("A2") <> Empty And opto_origin.range("A3") <> Empty) Then
    Set opto_origin_value = opto_origin.range("A2", opto_origin.range("A2").End(xlDown))
  ElseIf (opto_origin.range("A2") <> Empty And opto_origin.range("A3") = Empty) Then
    Set opto_origin_value = opto_origin.range("A2")
  End If

  ''   En los diccionarios de "opto_destiny_dictionary" y  "opto_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemOptoDestiny In opto_destiny_header
    On Error Resume Next
    opto_destiny_dictionary.Add opto_headers(ItemOptoDestiny), (ItemOptoDestiny.Column - 1)
    On Error GoTo 0
  Next ItemOptoDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemOptoOrigin In opto_origin_header
    On Error Resume Next
    opto_origin_dictionary.Add opto_headers(ItemOptoOrigin), (ItemOptoOrigin.Column - 1)
    On Error GoTo 0
  Next ItemOptoOrigin

  numbers = 1
  porcentaje = 0
  aumentFromRow = 0
  aumentFromIDOpto = destiny.Worksheets("RUTAS").range("$F$7").value
  aumentFromIDDiagnostic = destiny.Worksheets("RUTAS").range("$F$8").value
  counts = opto_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

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
      
      If (typeExams(charters(ItemData.Offset(, opto_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("IDENTIFICACION")) = charters(ItemData.Offset(, opto_origin_dictionary("IDENTIFICACION")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("VISIO/ANT_ LABORAL USUARIO COMPUTADOR")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL USUARIO COMPUTADOR")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("VISIO/ANT_ LABORAL CAMBIOS TEMPREATURA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL CAMBIOS TEMPREATURA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS FOTOFOBIA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS FOTOFOBIA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS OJO ROJO")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS OJO ROJO")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS LAGRIMEO")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS LAGRIMEO")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS VISION BORROSA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS VISION BORROSA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS ARDOR")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS ARDOR")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS VISION DOBLE")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS VISION DOBLE")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS CANSANCIO")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS CANSANCIO")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS MALA VISION CERCANA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS MALA VISION CERCANA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS DOLOR")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS DOLOR")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS MALA VISON LEJANA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS MALA VISON LEJANA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS SECRECION")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS SECRECION")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("SINTOMAS CEFALEA")) = charters_empty(ItemData.Offset(, opto_origin_dictionary("SINTOMAS CEFALEA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("OTROS SINTOMAS")) = charters(ItemData.Offset(, opto_origin_dictionary("OTROS SINTOMAS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CABEZA - PARPADOS")) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - PARPADOS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CABEZA - PARPADOS OBS")) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - PARPADOS OBS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CABEZA - CONJUNTIVAS")) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - CONJUNTIVAS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CABEZA - OBS CONJUNTIVAS")) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - OBS CONJUNTIVAS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CABEZA - ESCLERAS")) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - ESCLERAS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CABEZA - OBS ESCLERAS")) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - OBS ESCLERAS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CABEZA - PUPILAS")) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - PUPILAS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CABEZA - PUPILAS OBS")) = charters(ItemData.Offset(, opto_origin_dictionary("CABEZA - PUPILAS OBS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("MOT/OCUL COVERT TEST LEJOS")) = charters(ItemData.Offset(, opto_origin_dictionary("MOT/OCUL COVERT TEST LEJOS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("MOT/OCUL COVERT TEST CERCA")) = charters(ItemData.Offset(, opto_origin_dictionary("MOT/OCUL COVERT TEST CERCA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("ESTADO DE CORRECCION")) = charters(ItemData.Offset(, opto_origin_dictionary("ESTADO DE CORRECCION")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("PATOLOGIA OCULAR")) = charters(ItemData.Offset(, opto_origin_dictionary("PATOLOGIA OCULAR")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("DIAG PPAL")) = charters(ItemData.Offset(, opto_origin_dictionary("DIAG PPAL")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("DIAG OBS")) = charters(ItemData.Offset(, opto_origin_dictionary("DIAG OBS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("DIAG REL/1")) = charters(ItemData.Offset(, opto_origin_dictionary("DIAG REL/1")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("DIAG REL/2")) = charters(ItemData.Offset(, opto_origin_dictionary("DIAG REL/2")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("DIAG REL/3")) = charters(ItemData.Offset(, opto_origin_dictionary("DIAG REL/3")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")) = charters(ItemData.Offset(, opto_origin_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REC USO AR VIDEO TRMINAL")) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO AR VIDEO TRMINAL")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REC USO DE LENTES DE PROTECCION SOLAR")) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO DE LENTES DE PROTECCION SOLAR")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REC USO EPP VISUAL")) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO EPP VISUAL")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REC PAUSAS ACTIVAS")) = charters(ItemData.Offset(, opto_origin_dictionary("REC PAUSAS ACTIVAS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REC USO RX VISION PROXIMA")) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO RX VISION PROXIMA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REC USO RX DESCANSO")) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO RX DESCANSO")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REC USO PERMANENTE RX OPTICA")) = charters(ItemData.Offset(, opto_origin_dictionary("REC USO PERMANENTE RX OPTICA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REC PYP")) = charters(ItemData.Offset(, opto_origin_dictionary("REC PYP")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REC LUBRICANTE OCULAR")) = charters(ItemData.Offset(, opto_origin_dictionary("REC LUBRICANTE OCULAR")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("RECOMENDACIONES OBS")) = charters(ItemData.Offset(, opto_origin_dictionary("RECOMENDACIONES OBS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REM_ VALORACION OFTALM_")) = charters(ItemData.Offset(, opto_origin_dictionary("REM_ VALORACION OFTALM_")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REM_ TOPOGRAFIA CORNEAL")) = charters(ItemData.Offset(, opto_origin_dictionary("REM_ TOPOGRAFIA CORNEAL")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REM_ TRATAM_ ORTOPTICA")) = charters(ItemData.Offset(, opto_origin_dictionary("REM_ TRATAM_ ORTOPTICA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REM_ TEST FARNSWORTH")) = charters(ItemData.Offset(, opto_origin_dictionary("REM_ TEST FARNSWORTH")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REALIZAR PRUEBA AMBULATORIA")) = charters(ItemData.Offset(, opto_origin_dictionary("REALIZAR PRUEBA AMBULATORIA")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("REMISIONES OBS")) = charters(ItemData.Offset(, opto_origin_dictionary("REMISIONES OBS")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CONTROLES MENSUAL")) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES MENSUAL")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES_BIMESTRALES")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CONTROLES TRIMESTRAL")) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES TRIMESTRAL")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CONTROLES 6 MESES")) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES 6 MESES")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CONTROLES 1 ANO")) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES 1 ANO")))
        currenCell.Offset(aumentFromRow, opto_destiny_dictionary("CONTROLES CONFIRMATORIA")) = charters(ItemData.Offset(, opto_origin_dictionary("CONTROLES CONFIRMATORIA")))
        If (currenCell.Offset(aumentFromRow, 0).row = 4) Then
          currenCell.Offset(aumentFromRow, opto_destiny_dictionary("ID_OPTOMETRIA")) = Trim(aumentFromIDOpto)
          currenCell.Offset(aumentFromRow, opto_destiny_dictionary("OP_DIAGNOSTICO")) = Trim(aumentFromIDDiagnostic)
        Else
          aumentFromIDOpto = aumentFromIDOpto + 1
          aumentFromIDDiagnostic = aumentFromIDDiagnostic + 1
          currenCell.Offset(aumentFromRow, opto_destiny_dictionary("ID_OPTOMETRIA")) = Trim(aumentFromIDOpto)
          currenCell.Offset(aumentFromRow, opto_destiny_dictionary("OP_DIAGNOSTICO")) = Trim(aumentFromIDDiagnostic)
        End If
        aumentFromRow = aumentFromRow + 1
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData
  End With

  range("$A4").Select
  Call dataDuplicate
  range("$BD4:$BI4").Select
  Call greaterThanOne
  range("$BD4:$BI4").Select
  Call iqualCero
  range("$BK4").Select
  Call dataDuplicate
  range("$BL4").Select
  Call dataDuplicate
  range("$BM4").Select
  Call dataDuplicate
  range("$A4", range("$A4").End(xlDown)).Select
  Call formatter

  Set opto_origin_value = Nothing
  Set opto_destiny_header = Nothing
  Set opto_origin_header = Nothing
  opto_destiny_dictionary.RemoveAll
  opto_origin_dictionary.RemoveAll

End Sub
