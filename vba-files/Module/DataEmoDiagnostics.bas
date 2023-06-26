Attribute VB_Name = "DataEmoDiagnostics"
Option Explicit

'TODO: DataDiagnosticsEmo - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - diagnostics_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - emo_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - diagnostics_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - emo_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - emo_origin_value: Una variable de objeto para almacenar el rango de los datos de diagnosticos de la hoja de origen.
'* - ItemDiagnosticsDestiny: Una variable variante para iterar a traves del rango del encabezado de la hoja de destino.
'* - ItemEmoOrigin: Una variable variante para iterar a traves del rango del encabezado de la hoja de origen.
'* - ItemData: Una variable variante para iterar a traves del rango de los datos de diagnosticos de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Public Sub DataDiagnosticsEmo()

  Dim diagnostics_destiny_dictionary As Scripting.Dictionary
  Dim emo_origin_dictionary As Scripting.Dictionary
  Dim diagnostics_destiny_header As Object, emo_origin_header As Object, emo_origin_value As Object
  Dim ItemDiagnosticsDestiny As Variant, ItemEmoOrigin As Variant, ItemData As Variant
  Dim currenCell As range, aumentFromRow As LongPtr
  
  Set emo_origin = origin.Worksheets("EMO") '' EMO DEL LIBRO ORIGEN ''
  diagnostics_destiny.Select
  ActiveSheet.range("A5").Select
  
  Set currenCell = ActiveCell
  Set diagnostics_destiny_header = diagnostics_destiny.range("A4", diagnostics_destiny.range("A4").End(xlToRight))
  Set emo_origin_header = emo_origin.range("A1", emo_origin.range("A1").End(xlToRight))
  Set diagnostics_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (emo_origin.range("A2") <> Empty And emo_origin.range("A3") <> Empty) Then
    Set emo_origin_value = emo_origin.range("A2", emo_origin.range("A2").End(xlDown))
  ElseIf (emo_origin.range("A2") <> Empty And emo_origin.range("A3") = Empty) Then
    Set emo_origin_value = emo_origin.range("A2")
  End If

  ''   En los diccionarios de "diagnostics_destiny_dictionary" y  "emo_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  x = 1
  For Each ItemDiagnosticsDestiny In diagnostics_destiny_header
    On Error Resume Next
    diagnostics_destiny_dictionary.Add diagnostics_header(ItemDiagnosticsDestiny), (ItemDiagnosticsDestiny.Column - 1)
    On Error GoTo 0
  Next ItemDiagnosticsDestiny

  x = 1
  For Each ItemEmoOrigin In emo_origin_header
    On Error Resume Next
    emo_origin_dictionary.Add diagnostics_header(ItemEmoOrigin), (ItemEmoOrigin.Column - 1)
    On Error GoTo 0
  Next ItemEmoOrigin

  numbers = 1
  porcentaje = 0
  aumentFromRow = 0
  counts = emo_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each ItemData In emo_origin_value
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & diagnostics_destiny.Name
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

      If (typeExams(charters(ItemData.Offset(, emo_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        currenCell.Offset(aumentFromRow, diagnostics_destiny_dictionary("IDENTIFICACION")) = charters(ItemData.Offset(, emo_origin_dictionary("IDENTIFICACION")))
        currenCell.Offset(aumentFromRow, diagnostics_destiny_dictionary("CODIGO DIAG PPAL")) = charters(ItemData.Offset(, emo_origin_dictionary("CODIGO DIAG PPAL")))
        currenCell.Offset(aumentFromRow, diagnostics_destiny_dictionary("DIAG PPAL")) = charters(ItemData.Offset(, emo_origin_dictionary("DIAG PPAL")))
        For i = 1 To ((emo_origin_dictionary.Count - 5) / 2)
          currenCell.Offset(aumentFromRow, diagnostics_destiny_dictionary("CODIGO DIAG REL" & i)) = charters(ItemData.Offset(, emo_origin_dictionary("CODIGO DIAG REL" & i)))
          currenCell.Offset(aumentFromRow, diagnostics_destiny_dictionary("DIAG REL " & i)) = charters(ItemData.Offset(, emo_origin_dictionary("DIAG REL " & i)))
        Next i
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      aumentFromRow = aumentFromRow + 1
      DoEvents
    Next ItemData
  End With

  range("$A5").Select
  Call dataDuplicate
  range("$A5", range("$A5").End(xlDown)).Select
  Call formatter

  Set diagnostics_destiny_header = Nothing
  Set emo_origin_header = Nothing
  Set emo_origin_value = Nothing
  diagnostics_destiny_dictionary.RemoveAll
  emo_origin_dictionary.RemoveAll
  
End Sub
