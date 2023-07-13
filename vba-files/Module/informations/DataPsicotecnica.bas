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
Public Sub PsicotecnicaData()
  Dim psicotecnica_destiny_dictionary As Scripting.Dictionary
  Dim psicotecnica_origin_dictionary As Scripting.Dictionary
  Dim psicotecnica_destiny_header As Object, psicotecnica_origin_header As Object, psicotecnica_origin_value As Object
  Dim ItemPsicotecnicaDestiny As Variant, ItemPsicotecnicaOrigin As Variant, ItemData As Variant
  Dim currenCell As range, aumentFromRow As LongPtr, aumentFromID As LongPtr
  
  On Error GoTo tecnica:
  Set psico_origin = origin.Worksheets("PSICOTECNICA") '' PSICOTECNICA DEL LIBRO ORIGEN ''
  On Error GoTo 0
  
  psico_destiny.Select
  ActiveSheet.range("A2").Select
  Set currenCell = ActiveCell
  Set psicotecnica_destiny_header = psico_destiny.range("A1", psico_destiny.range("A1").End(xlToRight))
  Set psicotecnica_origin_header = psico_origin.range("A1", psico_origin.range("A1").End(xlToRight))
  Set psicotecnica_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set psicotecnica_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (psico_origin.range("A2") <> Empty And psico_origin.range("A3") <> Empty) Then
    Set psicotecnica_origin_value = psico_origin.range("A2", psico_origin.range("A2").End(xlDown))
  ElseIf (psico_origin.range("A2") <> Empty And psico_origin.range("A3") = Empty) Then
    Set psicotecnica_origin_value = psico_origin.range("A2")
  End If

  ''   En los diccionarios de "psicotecnica_destiny_dictionary" y  "psicotecnica_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemPsicotecnicaDestiny In psicotecnica_destiny_header
    On Error Resume Next
    psicotecnica_destiny_dictionary.Add psicotecnica_headers(ItemPsicotecnicaDestiny), (ItemPsicotecnicaDestiny.Column - 1)
    On Error GoTo 0
  Next ItemPsicotecnicaDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemPsicotecnicaOrigin In psicotecnica_origin_header
    On Error Resume Next
    psicotecnica_origin_dictionary.Add psicotecnica_headers(ItemPsicotecnicaOrigin), (ItemPsicotecnicaOrigin.Column - 1)
    On Error GoTo 0
  Next ItemPsicotecnicaOrigin

  numbers = 1
  porcentaje = 0
  aumentFromRow = 0
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$13").value
  counts = psicotecnica_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

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
      
      If (typeExams(charters(ItemData.Offset(, psicotecnica_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        currenCell.Offset(aumentFromRow, psicotecnica_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, psicotecnica_origin_dictionary("NRO IDENFICACION")))
        currenCell.Offset(aumentFromRow, psicotecnica_destiny_dictionary("PACIENTE")) = charters(ItemData.Offset(, psicotecnica_origin_dictionary("PACIENTE")))
        currenCell.Offset(aumentFromRow, psicotecnica_destiny_dictionary("PRUEBA PSICOTECNICA")) = charters(ItemData.Offset(, psicotecnica_origin_dictionary("PRUEBA PSICOTECNICA")))
        currenCell.Offset(aumentFromRow, psicotecnica_destiny_dictionary("DIAGNOSTICO PPAL (CUMPLE, NO CUMPLE)")) = charters(ItemData.Offset(, psicotecnica_origin_dictionary("DIAGNOSTICO PPAL (CUMPLE, NO CUMPLE)")))
        currenCell.Offset(aumentFromRow, psicotecnica_destiny_dictionary("DIAGNOSTICO OBS")) = charters(ItemData.Offset(, psicotecnica_origin_dictionary("DIAGNOSTICO OBS")))
        If (currenCell.Offset(aumentFromRow, 0).row = 2) Then
          currenCell.Offset(aumentFromRow, psicotecnica_destiny_dictionary("ID_PSICOTECNICA")) = Trim(aumentFromID)
        Else
          aumentFromID = aumentFromID + 1
          currenCell.Offset(aumentFromRow, psicotecnica_destiny_dictionary("ID_PSICOTECNICA")) = Trim(aumentFromID)
        End If
        aumentFromRow = aumentFromRow + 1
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData
  End With

  range("D2").Select
  Call meetsfails
  range("$A2", range("$A2").End(xlDown)).Select
  Call formatter

  Set psicotecnica_origin_value = Nothing
  Set psicotecnica_destiny_header = Nothing
  Set psicotecnica_origin_header = Nothing
  psicotecnica_destiny_dictionary.RemoveAll
  psicotecnica_origin_dictionary.RemoveAll

  Exit Sub

tecnica:
  Set psico_origin = origin.Worksheets("PSICOLOGIA")
  Resume Next
End Sub
