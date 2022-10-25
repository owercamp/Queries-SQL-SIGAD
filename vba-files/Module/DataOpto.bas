Attribute VB_Name = "DataOpto"
Option Explicit

Sub OptoData()

  Dim opto_destiny_dictionary As Scripting.Dictionary
  Dim opto_origin_dictionary As Scripting.Dictionary
  Dim opto_destiny_header, opto_origin_header, opto_origin_value As Object
  Dim ItemOptoDestiny, ItemOptoOrigin, ItemData As Variant

  Set opto_origin = origin.Worksheets("OPTO") '' OPTO DEL LIBRO ORIGEN ''
  opto_destiny.Select
  ActiveSheet.Range("A5").Select
  Set opto_destiny_header = opto_destiny.Range("A3", opto_destiny.Range("A3").End(xlToRight))
  Set opto_origin_header = opto_origin.Range("A1", opto_origin.Range("A1").End(xlToRight))
  Set opto_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set opto_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (opto_origin.Range("A2") <> Empty And opto_origin.Range("A3") <> Empty) Then
    Set opto_origin_value = opto_origin.Range("A2", opto_origin.Range("A2").End(xlDown))
  ElseIf (opto_origin.Range("A2") <> Empty And opto_origin.Range("A3") = Empty) Then
    Set opto_origin_value = opto_origin.Range("A2")
  End If

  '/***
  '   En los diccionarios de "opto_destiny_dictionary" y  "opto_origin_dictionary"
  '   se almacena los numeros de la columnas.
  '*/

  ' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO
  For Each ItemOptoDestiny In opto_destiny_header
    On Error Goto optoError
    opto_destiny_dictionary.Add opto_headers(ItemOptoDestiny), (ItemOptoDestiny.Column - 1)
  Next ItemOptoDestiny

  ' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN
  For Each ItemOptoOrigin In opto_origin_header
    On Error Goto optoError
    opto_origin_dictionary.Add opto_headers(ItemOptoOrigin), (ItemOptoOrigin.Column - 1)
  Next ItemOptoOrigin

  numbers = 1
  porcentaje = 0
  counts = opto_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  For Each ItemData In opto_origin_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & opto_destiny.Name
        porcentaje = porcentaje + vals
        porcentajeGeneral = porcentajeGeneral + valsGeneral
        formImports.ProgressBarOneforOne.Width = oneForOne
        formImports.ProgressBarGeneral.Width = generalAll
        formImports.porcentageGeneral.Caption = CStr(VBA.Round(porcentajeGeneral * 100, 1)) & "%"
        formImports.porcentageOneoforOne.Caption = CStr(VBA.Round(porcentaje * 100, 1)) & "%"
        formImports.Caption = CStr(nameCompany)
        If formImports.ProgressBarGeneral.Width > (formImports.content_ProgressBarGeneral.Width / 2) Then: formImports.porcentageGeneral.ForeColor = RGB(255, 255, 255)
          If formImports.ProgressBarGeneral.Width < (formImports.content_ProgressBarGeneral.Width / 2) Then: formImports.porcentageGeneral.ForeColor = RGB(0, 0, 0)
            If formImports.ProgressBarOneforOne.Width > (formImports.content_ProgressBarOneforOne.Width / 2) Then: formImports.porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
              If formImports.ProgressBarOneforOne.Width < (formImports.content_ProgressBarOneforOne.Width / 2) Then: formImports.porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
                ActiveCell.offset(, opto_destiny_dictionary("IDENTIFICACION")) = charters(ItemData.offset(, opto_origin_dictionary( "IDENTIFICACION")))
                ActiveCell.offset(, opto_destiny_dictionary("VISIO/ANT_ LABORAL ILUMINACION INADECUADA")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "VISIO/ANT_ LABORAL ILUMINACION INADECUADA")))
                ActiveCell.offset(, opto_destiny_dictionary("VISIO/ANT_ LABORAL USUARIO COMPUTADOR")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "VISIO/ANT_ LABORAL USUARIO COMPUTADOR")))
                ActiveCell.offset(, opto_destiny_dictionary("VISIO/ANT_ LABORALVISIO RADIACIONES UV")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "VISIO/ANT_ LABORALVISIO RADIACIONES UV")))
                ActiveCell.offset(, opto_destiny_dictionary("VISIO/ANT_ LABORAL CAMBIOS TEMPREATURA")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "VISIO/ANT_ LABORAL CAMBIOS TEMPREATURA")))
                ActiveCell.offset(, opto_destiny_dictionary("VISIO/ANT_ LABORAL MALA VENTILACION")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "VISIO/ANT_ LABORAL MALA VENTILACION")))
                ActiveCell.offset(, opto_destiny_dictionary("VISIO/ANT_ LABORAL GASES TOXICOS")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "VISIO/ANT_ LABORAL GASES TOXICOS")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS FOTOFOBIA")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS FOTOFOBIA")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS OJO ROJO")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS OJO ROJO")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS LAGRIMEO")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS LAGRIMEO")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS VISION BORROSA")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS VISION BORROSA")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS ARDOR")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS ARDOR")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS VISION DOBLE")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS VISION DOBLE")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS CANSANCIO")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS CANSANCIO")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS MALA VISION CERCANA")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS MALA VISION CERCANA")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS DOLOR")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS DOLOR")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS MALA VISON LEJANA")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS MALA VISON LEJANA")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS SECRECION")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS SECRECION")))
                ActiveCell.offset(, opto_destiny_dictionary("SINTOMAS CEFALEA")) = charters_empty(ItemData.offset(, opto_origin_dictionary( "SINTOMAS CEFALEA")))
                ActiveCell.offset(, opto_destiny_dictionary("OTROS SINTOMAS")) = charters(ItemData.offset(, opto_origin_dictionary( "OTROS SINTOMAS")))
                ActiveCell.offset(, opto_destiny_dictionary("CABEZA - PARPADOS")) = charters(ItemData.offset(, opto_origin_dictionary( "CABEZA - PARPADOS")))
                ActiveCell.offset(, opto_destiny_dictionary("CABEZA - PARPADOS OBS")) = charters(ItemData.offset(, opto_origin_dictionary( "CABEZA - PARPADOS OBS")))
                ActiveCell.offset(, opto_destiny_dictionary("CABEZA - CONJUNTIVAS")) = charters(ItemData.offset(, opto_origin_dictionary( "CABEZA - CONJUNTIVAS")))
                ActiveCell.offset(, opto_destiny_dictionary("CABEZA - OBS CONJUNTIVAS")) = charters(ItemData.offset(, opto_origin_dictionary( "CABEZA - OBS CONJUNTIVAS")))
                ActiveCell.offset(, opto_destiny_dictionary("CABEZA - ESCLERAS")) = charters(ItemData.offset(, opto_origin_dictionary( "CABEZA - ESCLERAS")))
                ActiveCell.offset(, opto_destiny_dictionary("CABEZA - OBS ESCLERAS")) = charters(ItemData.offset(, opto_origin_dictionary( "CABEZA - OBS ESCLERAS")))
                ActiveCell.offset(, opto_destiny_dictionary("CABEZA - PUPILAS")) = charters(ItemData.offset(, opto_origin_dictionary( "CABEZA - PUPILAS")))
                ActiveCell.offset(, opto_destiny_dictionary("CABEZA - PUPILAS OBS")) = charters(ItemData.offset(, opto_origin_dictionary( "CABEZA - PUPILAS OBS")))
                ActiveCell.offset(, opto_destiny_dictionary("MOT/OCUL COVERT TEST LEJOS")) = charters(ItemData.offset(, opto_origin_dictionary( "MOT/OCUL COVERT TEST LEJOS")))
                ActiveCell.offset(, opto_destiny_dictionary("MOT/OCUL COVERT TEST CERCA")) = charters(ItemData.offset(, opto_origin_dictionary( "MOT/OCUL COVERT TEST CERCA")))
                ActiveCell.offset(, opto_destiny_dictionary("ESTADO DE CORRECCION")) = charters(ItemData.offset(, opto_origin_dictionary( "ESTADO DE CORRECCION")))
                ActiveCell.offset(, opto_destiny_dictionary("PATOLOGIA OCULAR")) = charters(ItemData.offset(, opto_origin_dictionary( "PATOLOGIA OCULAR")))
                ActiveCell.offset(, opto_destiny_dictionary("DIAG PPAL")) = charters(ItemData.offset(, opto_origin_dictionary( "DIAG PPAL")))
                ActiveCell.offset(, opto_destiny_dictionary("DIAG OBS")) = charters(ItemData.offset(, opto_origin_dictionary( "DIAG OBS")))
                ActiveCell.offset(, opto_destiny_dictionary("DIAG REL/1")) = charters(ItemData.offset(, opto_origin_dictionary( "DIAG REL/1")))
                ActiveCell.offset(, opto_destiny_dictionary("DIAG REL/2")) = charters(ItemData.offset(, opto_origin_dictionary( "DIAG REL/2")))
                ActiveCell.offset(, opto_destiny_dictionary("DIAG REL/3")) = charters(ItemData.offset(, opto_origin_dictionary( "DIAG REL/3")))
                ActiveCell.offset(, opto_destiny_dictionary("REC CORRECCION VISUAL PARA TRABAJAR")) = charters(ItemData.offset(, opto_origin_dictionary( "REC CORRECCION VISUAL PARA TRABAJAR")))
                ActiveCell.offset(, opto_destiny_dictionary("REC USO AR VIDEO TRMINAL")) = charters(ItemData.offset(, opto_origin_dictionary( "REC USO AR VIDEO TRMINAL")))
                ActiveCell.offset(, opto_destiny_dictionary("REC USO DE LENTES DE PROTECCION SOLAR")) = charters(ItemData.offset(, opto_origin_dictionary( "REC USO DE LENTES DE PROTECCION SOLAR")))
                ActiveCell.offset(, opto_destiny_dictionary("REC USO EPP VISUAL")) = charters(ItemData.offset(, opto_origin_dictionary( "REC USO EPP VISUAL")))
                ActiveCell.offset(, opto_destiny_dictionary("REC PAUSAS ACTIVAS")) = charters(ItemData.offset(, opto_origin_dictionary( "REC PAUSAS ACTIVAS")))
                ActiveCell.offset(, opto_destiny_dictionary("REC USO RX VISION PROXIMA")) = charters(ItemData.offset(, opto_origin_dictionary( "REC USO RX VISION PROXIMA")))
                ActiveCell.offset(, opto_destiny_dictionary("REC USO RX DESCANSO")) = charters(ItemData.offset(, opto_origin_dictionary( "REC USO RX DESCANSO")))
                ActiveCell.offset(, opto_destiny_dictionary("REC USO PERMANENTE RX OPTICA")) = charters(ItemData.offset(, opto_origin_dictionary( "REC USO PERMANENTE RX OPTICA")))
                ActiveCell.offset(, opto_destiny_dictionary("REC PYP")) = charters(ItemData.offset(, opto_origin_dictionary( "REC PYP")))
                ActiveCell.offset(, opto_destiny_dictionary("REC LUBRICANTE OCULAR")) = charters(ItemData.offset(, opto_origin_dictionary( "REC LUBRICANTE OCULAR")))
                ActiveCell.offset(, opto_destiny_dictionary("RECOMENDACIONES OBS")) = charters(ItemData.offset(, opto_origin_dictionary( "RECOMENDACIONES OBS")))
                ActiveCell.offset(, opto_destiny_dictionary("REM_ VALORACION OFTALM_")) = charters(ItemData.offset(, opto_origin_dictionary( "REM_ VALORACION OFTALM_")))
                ActiveCell.offset(, opto_destiny_dictionary("REM_ TOPOGRAFIA CORNEAL")) = charters(ItemData.offset(, opto_origin_dictionary( "REM_ TOPOGRAFIA CORNEAL")))
                ActiveCell.offset(, opto_destiny_dictionary("REM_ TRATAM_ ORTOPTICA")) = charters(ItemData.offset(, opto_origin_dictionary( "REM_ TRATAM_ ORTOPTICA")))
                ActiveCell.offset(, opto_destiny_dictionary("REM_ TEST FARNSWORTH")) = charters(ItemData.offset(, opto_origin_dictionary( "REM_ TEST FARNSWORTH")))
                ActiveCell.offset(, opto_destiny_dictionary("REALIZAR PRUEBA AMBULATORIA")) = charters(ItemData.offset(, opto_origin_dictionary( "REALIZAR PRUEBA AMBULATORIA")))
                ActiveCell.offset(, opto_destiny_dictionary("REMISIONES OBS")) = charters(ItemData.offset(, opto_origin_dictionary( "REMISIONES OBS")))
                ActiveCell.offset(, opto_destiny_dictionary("CONTROLES MENSUAL")) = charters(ItemData.offset(, opto_origin_dictionary( "CONTROLES MENSUAL")))
                ActiveCell.offset(, opto_destiny_dictionary("CONTROLES_BIMESTRALES")) = charters(ItemData.offset(, opto_origin_dictionary( "CONTROLES_BIMESTRALES")))
                ActiveCell.offset(, opto_destiny_dictionary("CONTROLES TRIMESTRAL")) = charters(ItemData.offset(, opto_origin_dictionary( "CONTROLES TRIMESTRAL")))
                ActiveCell.offset(, opto_destiny_dictionary("CONTROLES 6 MESES")) = charters(ItemData.offset(, opto_origin_dictionary( "CONTROLES 6 MESES")))
                ActiveCell.offset(, opto_destiny_dictionary("CONTROLES 1 ANO")) = charters(ItemData.offset(, opto_origin_dictionary( "CONTROLES 1 ANO")))
                ActiveCell.offset(, opto_destiny_dictionary("CONTROLES CONFIRMATORIA")) = charters(ItemData.offset(, opto_origin_dictionary( "CONTROLES CONFIRMATORIA")))
                ActiveCell.offset(, opto_destiny_dictionary("ID_OPTOMETRIA")) = ActiveCell.offset(-1, opto_destiny_dictionary("ID_OPTOMETRIA")) + 1
                ActiveCell.offset(, opto_destiny_dictionary("OP_DIAGNOSTICO")) = ActiveCell.offset(-1, opto_destiny_dictionary("OP_DIAGNOSTICO")) + 1
                ActiveCell.offset(1, 0).Select
                numbers = numbers + 1
                numbersGeneral = numbersGeneral + 1
                DoEvents
              Next ItemData

              Range("$A4").Select
              Call dataDuplicate
              Range("$BD4:$BI4").Select
              Call greaterThanOne
              Range("$BD4:$BI4").Select
              Call iqualCero
              Range("$BK4").Select
              Call dataDuplicate
              Range("$BL4").Select
              Call dataDuplicate
              Range("$BM4").Select
              Call dataDuplicate
              Range("$A4", Range("$A4").End(xlDown)).Select
              Call formatter

              Set opto_origin_value = Nothing
              Set opto_destiny_header = Nothing
              Set opto_origin_header = Nothing
              opto_destiny_dictionary.RemoveAll
              opto_origin_dictionary.RemoveAll
 optoError:
              resume next
End Sub