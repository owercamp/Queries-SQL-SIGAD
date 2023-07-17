Attribute VB_Name = "DataEmoDiagnostics"
'namespace=vba-files\Module\informations
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
Dim emo_origin_dictionary As Scripting.Dictionary
Public Sub DataDiagnosticsEmo()
  Dim tbl_diagnostics As Object, emo_origin_header As Object, emo_origin_value As Object
  Dim ItemEmoOrigin As Variant, ItemData As Variant, counter As LongPtr
  
  Set emo_origin = origin.Worksheets("EMO") '' EMO DEL LIBRO ORIGEN ''
  diagnostics_destiny.Select
  Set tbl_diagnostics = ActiveSheet.ListObjects("tbl_diagnosticos")
  Set emo_origin_header = emo_origin.range("A1", emo_origin.range("A1").End(xlToRight))
  Set emo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (emo_origin.range("A2") <> Empty And emo_origin.range("A3") <> Empty) Then
    Set emo_origin_value = emo_origin.range("A2", emo_origin.range("A2").End(xlDown))
  ElseIf (emo_origin.range("A2") <> Empty And emo_origin.range("A3") = Empty) Then
    Set emo_origin_value = emo_origin.range("A2")
  End If

  x = 1
  For Each ItemEmoOrigin In emo_origin_header
    On Error Resume Next
    emo_origin_dictionary.Add diagnostics_header(ItemEmoOrigin), (ItemEmoOrigin.Column - 1)
    On Error GoTo 0
  Next ItemEmoOrigin

  numbers = 1
  porcentaje = 0
  
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

      counter = (emo_origin_dictionary.Count - 5) / 2
      If (typeExams(charters(ItemData.Offset(, emo_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        Select Case numbers
          Case 1
            Call addNewRegister(tbl_diagnostics.ListRows(1), counter, ItemData)
          Case Else
            Call addNewRegister(tbl_diagnostics.ListRows.Add, counter, ItemData)
        End Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData
  End With

  range("$A5").Select
  Call dataDuplicate
  range("$A5", range("$A5").End(xlDown)).Select
  Call formatter

  Set emo_origin_header = Nothing
  Set emo_origin_value = Nothing
  emo_origin_dictionary.RemoveAll
  
End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal numberMaxEmphasis As LongPtr, ByVal ItemData As Variant)

  Dim numberEmphasis As LongPtr
  numberEmphasis = 1
  With table
    .Range(1) = charters(ItemData.Offset(, emo_origin_dictionary("IDENTIFICACION")))
    .Range(4) = charters(ItemData.Offset(, emo_origin_dictionary("CODIGO DIAG PPAL")))
    .Range(5) = charters(ItemData.Offset(, emo_origin_dictionary("DIAG PPAL")))
    For i = 6 to 71 Step 2
      Select Case numberMaxEmphasis
        Case 0
          Exit For
        Case Is > 0
          .Range(i) = charters(ItemData.Offset(, emo_origin_dictionary("CODIGO DIAG REL" & numberEmphasis)))
          .Range(i + 1) = charters(ItemData.Offset(, emo_origin_dictionary("DIAG REL " & numberEmphasis)))
          numberEmphasis = numberEmphasis + 1
          numberMaxEmphasis = numberMaxEmphasis - 1
      End Select
    Next i
    DoEvents
  End With

End Sub