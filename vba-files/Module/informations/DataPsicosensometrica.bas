Attribute VB_Name = "DataPsicosensometrica"
'namespace=vba-files\Module\informations
Option Explicit

'TODO: PsicosensometricaData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - psicosensometrica_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - psicosensometrica_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - psicosensometrica_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - psicosensometrica_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - psicosensometrica_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Dim aumentFromID As LongPtr
Public Sub PsicosensometricaData(ByVal name_sheet As String)
  Dim psicosensometrica_destiny_dictionary As Scripting.Dictionary
  Dim psicosensometrica_origin_dictionary As Scripting.Dictionary
  Dim psicosensometrica_destiny_header As Object, psicosensometrica_origin_header As Object, psicosensometrica_origin_value As Object
  Dim ItemPsicosensometricaDestiny As Object, ItemPsicosensometricaOrigin As Object, ItemData As Object, senso_origin As Object

  Set senso_origin = origin.Worksheets(name_sheet) '' PSICOSENSOMETRICA DEL LIBRO ORIGEN ''

  senso_destiny.Select
  senso_destiny.Range("$A3").Select
  Set psicosensometrica_destiny_header = senso_destiny.Range("$A2", senso_destiny.Range("$A2").End(xlToRight))
  Set psicosensometrica_origin_header = senso_origin.Range("$A1", senso_origin.Range("$A1").End(xlToRight))
  Set psicosensometrica_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set psicosensometrica_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (senso_origin.Range("$A2") <> Empty And senso_origin.Range("$A3") <> Empty) Then
    Set psicosensometrica_origin_value = senso_origin.Range("$A2", senso_origin.Range("$A2").End(xlDown))
  ElseIf (senso_origin.Range("$A2") <> Empty And senso_origin.Range("$A3") = Empty) Then
    Set psicosensometrica_origin_value = senso_origin.Range("$A2")
  End If

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  Dim value_data As String
  For Each ItemPsicosensometricaDestiny In psicosensometrica_destiny_header
    value_data = psicosensometrica_headers(ItemPsicosensometricaDestiny)
    If psicosensometrica_destiny_dictionary.Exists(value_data) = False And value_data <> Empty Then
      psicosensometrica_destiny_dictionary.Add value_data, (ItemPsicosensometricaDestiny.Column - 1)
    End If
  Next ItemPsicosensometricaDestiny
  
  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemPsicosensometricaOrigin In psicosensometrica_origin_header
    value_data = psicosensometrica_headers(ItemPsicosensometricaOrigin)
    If psicosensometrica_origin_dictionary.Exists(value_data) = False And value_data <> Empty Then
      psicosensometrica_origin_dictionary.Add value_data, (ItemPsicosensometricaOrigin.Column - 1)
    End If
  Next ItemPsicosensometricaOrigin

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$14").value
  counts = psicosensometrica_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  Dim type_exam As String
  With formImports
    For Each ItemData In psicosensometrica_origin_value
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & senso_destiny.Name
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
      
      type_exam = typeExams(Trim(ItemData.Offset(, psicosensometrica_origin_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("NRO IDENFICACION")) = Trim(ItemData.Offset(, psicosensometrica_origin_dictionary("NRO IDENFICACION")))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("PACIENTE")) = Trim(UCase(ItemData.Offset(, psicosensometrica_origin_dictionary("PACIENTE"))))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("PRUEBA PSICOSENSOMETRICA")) = Trim(UCase(ItemData.Offset(, psicosensometrica_origin_dictionary("PRUEBA PSICOSENSOMETRICA"))))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO PPAL")) = Trim(UCase(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO PPAL"))))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO OBS")) = Trim(UCase(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO OBS"))))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/1")) = Trim(UCase(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/1"))))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/2")) = Trim(UCase(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/2"))))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/3")) = Trim(UCase(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/3"))))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("CONTROLES MENSUALES")) = Trim(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES MENSUALES")))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("CONTROLES BIMENSUAL")) = Trim(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES BIMENSUAL")))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("CONTROLES TRIMESTRALES")) = Trim(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES TRIMESTRALES")))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("CONTROLES 6 MESES")) = Trim(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES 6 MESES")))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("CONTROLES 1 ANO")) = Trim(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES 1 ANO")))
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("CONTROLES CONFIRMATORIA")) = Trim(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES CONFIRMATORIA")))
        If (ActiveCell.Row <> 3) Then
          aumentFromID = aumentFromID + 1
        End If
        ActiveCell.Offset(, psicosensometrica_destiny_dictionary("ID_PSICOSENSOMETRICA")) = aumentFromID
        ActiveCell.Offset(1, 0).Select
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents
      End If
    Next ItemData
  End With

  Call greaterThanOne(senso_destiny.Range("tbl_psicosensometrica[[CONTROLES MENSUALES]:[CONTROLES CONFIRMATORIA]]"), "PSICOSENSOMETRICA")
  Call iqualCero(senso_destiny.Range("tbl_psicosensometrica[[CONTROLES MENSUALES]:[CONTROLES CONFIRMATORIA]]"), "PSICOSENSOMETRICA")
  Call formatter(senso_destiny.Range("tbl_psicosensometrica[[#Data],[NRO IDENFICACION]]"))

  Set psicosensometrica_origin_value = Nothing
  Set psicosensometrica_destiny_header = Nothing
  Set psicosensometrica_origin_header = Nothing
  psicosensometrica_destiny_dictionary.RemoveAll
  psicosensometrica_origin_dictionary.RemoveAll

End Sub