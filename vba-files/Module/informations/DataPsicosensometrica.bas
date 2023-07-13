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
Public Sub PsicosensometricaData()
  Dim psicosensometrica_destiny_dictionary As Scripting.Dictionary
  Dim psicosensometrica_origin_dictionary As Scripting.Dictionary
  Dim psicosensometrica_destiny_header As Object, psicosensometrica_origin_header As Object, psicosensometrica_origin_value As Object
  Dim ItemPsicosensometricaDestiny As Variant, ItemPsicosensometricaOrigin As Variant, ItemData As Variant
  Dim currenCell As range, aumentFromRow As LongPtr, aumentFromID As LongPtr
  
  On Error GoTo metrica:
  Set senso_origin = origin.Worksheets("PSICOSENSOMETRICA") '' PSICOSENSOMETRICA DEL LIBRO ORIGEN ''
  On Error GoTo 0
  
  senso_destiny.Select
  ActiveSheet.range("A3").Select
  Set currenCell = ActiveCell
  Set psicosensometrica_destiny_header = senso_destiny.range("A2", senso_destiny.range("A2").End(xlToRight))
  Set psicosensometrica_origin_header = senso_origin.range("A1", senso_origin.range("A1").End(xlToRight))
  Set psicosensometrica_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set psicosensometrica_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (senso_origin.range("A2") <> Empty And senso_origin.range("A3") <> Empty) Then
    Set psicosensometrica_origin_value = senso_origin.range("A2", senso_origin.range("A2").End(xlDown))
  ElseIf (senso_origin.range("A2") <> Empty And senso_origin.range("A3") = Empty) Then
    Set psicosensometrica_origin_value = senso_origin.range("A2")
  End If

  ''   En los diccionarios de "psicosensometrica_destiny_dictionary" y  "psicosensometrica_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemPsicosensometricaDestiny In psicosensometrica_destiny_header
    On Error Resume Next
    psicosensometrica_destiny_dictionary.Add psicosensometrica_headers(ItemPsicosensometricaDestiny), (ItemPsicosensometricaDestiny.Column - 1)
    On Error GoTo 0
  Next ItemPsicosensometricaDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemPsicosensometricaOrigin In psicosensometrica_origin_header
    On Error Resume Next
    psicosensometrica_origin_dictionary.Add psicosensometrica_headers(ItemPsicosensometricaOrigin), (ItemPsicosensometricaOrigin.Column - 1)
    On Error GoTo 0
  Next ItemPsicosensometricaOrigin

  numbers = 1
  porcentaje = 0
  aumentFromRow = 0
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$14").value
  counts = psicosensometrica_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

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
      
      If (typeExams(charters(ItemData.Offset(, psicosensometrica_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("NRO IDENFICACION")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("PACIENTE")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("PACIENTE")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("PRUEBA PSICOSENSOMETRICA")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("PRUEBA PSICOSENSOMETRICA")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("DIAGNOSTICO PPAL")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO PPAL")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("DIAGNOSTICO OBS")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO OBS")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/1")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/1")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/2")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/2")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/3")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/3")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("CONTROLES MENSUALES")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES MENSUALES")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("CONTROLES BIMENSUAL")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES BIMENSUAL")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("CONTROLES TRIMESTRALES")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES TRIMESTRALES")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("CONTROLES 6 MESES")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES 6 MESES")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("CONTROLES 1 ANO")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES 1 ANO")))
        currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("CONTROLES CONFIRMATORIA")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES CONFIRMATORIA")))
        If (currenCell.Offset(aumentFromRow, 0).row = 3) Then
          currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("ID_PSICOSENSOMETRICA")) = Trim(aumentFromID)
        Else
          aumentFromID = aumentFromID + 1
          currenCell.Offset(aumentFromRow, psicosensometrica_destiny_dictionary("ID_PSICOSENSOMETRICA")) = Trim(aumentFromID)
        End If
        aumentFromRow = aumentFromRow + 1
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData
  End With

  range("$I3:$N3").Select
  Call greaterThanOne
  range("$I3:$N3").Select
  Call iqualCero
  range("$A3", range("$A3").End(xlDown)).Select
  Call formatter

  Set psicosensometrica_origin_value = Nothing
  Set psicosensometrica_destiny_header = Nothing
  Set psicosensometrica_origin_header = Nothing
  psicosensometrica_destiny_dictionary.RemoveAll
  psicosensometrica_origin_dictionary.RemoveAll

  Exit Sub

metrica:
  Set senso_origin = origin.Worksheets("PSICOMOTRIZ")
  Resume Next
End Sub
