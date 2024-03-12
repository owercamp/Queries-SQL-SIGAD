Attribute VB_Name = "DataPsicotecnica"
'namespace=vba-files\Module\informations
Option Explicit

'TODO: PsicotecnicaData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - psicotecnica_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - psicotecnica_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - psicotecnica_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - psicotecnica_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - psicotecnica_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Dim aumentFromID As LongPtr
Public Sub PsicotecnicaData(ByVal name_sheet As String)
  Dim psicotecnica_destiny_dictionary As Scripting.Dictionary
  Dim psicotecnica_origin_dictionary As Scripting.Dictionary
  Dim psicotecnica_destiny_header As Object, psicotecnica_origin_header As Object, psicotecnica_origin_value As Object
  Dim ItemPsicotecnicaDestiny As Object, ItemPsicotecnicaOrigin As Object, ItemData As Object, psico_origin As Object

  Set psico_origin = origin.Worksheets(name_sheet) '' PSICOTECNICA DEL LIBRO ORIGEN ''

  psico_destiny.Select
  psico_destiny.Range("$A2").Select
  Set psicotecnica_destiny_header = psico_destiny.Range("$A1", psico_destiny.Range("$A1").End(xlToRight))
  Set psicotecnica_origin_header = psico_origin.Range("$A1", psico_origin.Range("$A1").End(xlToRight))
  Set psicotecnica_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set psicotecnica_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (psico_origin.Range("$A2") <> Empty And psico_origin.Range("$A3") <> Empty) Then
    Set psicotecnica_origin_value = psico_origin.Range("$A2", psico_origin.Range("$A2").End(xlDown))
  ElseIf (psico_origin.Range("$A2") <> Empty And psico_origin.Range("$A3") = Empty) Then
    Set psicotecnica_origin_value = psico_origin.Range("$A2")
  End If

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  Dim value_data As String
  For Each ItemPsicotecnicaDestiny In psicotecnica_destiny_header
    value_data = psicotecnica_headers(ItemPsicotecnicaDestiny)
    If psicotecnica_destiny_dictionary.Exists(value_data) = False And value_data <> Empty Then
      psicotecnica_destiny_dictionary.Add value_data, (ItemPsicotecnicaDestiny.Column - 1)
    End If
  Next ItemPsicotecnicaDestiny
  
  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemPsicotecnicaOrigin In psicotecnica_origin_header
    value_data = psicotecnica_headers(ItemPsicotecnicaOrigin)
    If psicotecnica_origin_dictionary.Exists(value_data) = False And value_data <> Empty Then
      psicotecnica_origin_dictionary.Add value_data, (ItemPsicotecnicaOrigin.Column - 1)
    End If
  Next ItemPsicotecnicaOrigin

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$13").value
  counts = psicotecnica_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  Dim type_exam As String
  With formImports
    For Each ItemData In psicotecnica_origin_value
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & psico_destiny.Name
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
      
      type_exam = typeExams(Trim(ItemData.Offset(, psicotecnica_origin_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        ActiveCell.Offset(, psicotecnica_destiny_dictionary("NRO IDENFICACION")) = Trim(ItemData.Offset(, psicotecnica_origin_dictionary("NRO IDENFICACION")))
        ActiveCell.Offset(, psicotecnica_destiny_dictionary("PACIENTE")) = Trim(UCase(ItemData.Offset(, psicotecnica_origin_dictionary("PACIENTE"))))
        ActiveCell.Offset(, psicotecnica_destiny_dictionary("PRUEBA PSICOTECNICA")) = Trim(UCase(ItemData.Offset(, psicotecnica_origin_dictionary("PRUEBA PSICOTECNICA"))))
        ActiveCell.Offset(, psicotecnica_destiny_dictionary("DIAGNOSTICO PPAL (CUMPLE, NO CUMPLE)")) = Trim(UCase(ItemData.Offset(, psicotecnica_origin_dictionary("DIAGNOSTICO PPAL (CUMPLE, NO CUMPLE)"))))
        ActiveCell.Offset(, psicotecnica_destiny_dictionary("DIAGNOSTICO OBS")) = Trim(UCase(ItemData.Offset(, psicotecnica_origin_dictionary("DIAGNOSTICO OBS"))))
        If (ActiveCell.Row <> 2) Then
          aumentFromID = aumentFromID + 1
        End If
        ActiveCell.Offset(, psicotecnica_destiny_dictionary("ID_PSICOTECNICA")) = aumentFromID
        ActiveCell.Offset(1, 0).Select
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents
      End If
    Next ItemData
  End With

  Call meetsfails(psico_destiny.Range("tbl_psicotecnica[[#Data],[DIAGNOSTICO PPAL (CUMPLE, NO CUMPLE)]]"))
  Call formatter(psico_destiny.Range("tbl_psicotecnica[[#Data],[NRO IDENFICACION]]"))

  Set psicotecnica_origin_value = Nothing
  Set psicotecnica_destiny_header = Nothing
  Set psicotecnica_origin_header = Nothing
  psicotecnica_destiny_dictionary.RemoveAll
  psicotecnica_origin_dictionary.RemoveAll

End Sub