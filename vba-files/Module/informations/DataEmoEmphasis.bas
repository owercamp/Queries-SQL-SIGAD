Attribute VB_Name = "DataEmoEmphasis"
'namespace=vba-files\Module\informations
Option Explicit

'TODO: DataEmphasisEmo - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - emphasis_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - emo_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - emphasis_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - emo_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - emo_origin_value: Una variable de objeto para almacenar el rango de los datos de diagnosticos de la hoja de origen.
'* - ItemEmphasisDestiny: Una variable variante para iterar a traves del rango del encabezado de la hoja de destino.
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
Public Sub DataEmphasisEmo()
  Dim tbl_emphasis As Object, emo_origin_header As Object, emo_origin_value As Object
  Dim ItemEmoOrigin As Variant, ItemData As Variant, counter As LongPtr
  
  Set emo_origin = origin.Worksheets("EMO") '' EMO DEL LIBRO ORIGEN ''
  emphasis_destiny.Select
  Set tbl_emphasis = ActiveSheet.ListObjects("tbl_enfasis")
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
    emo_origin_dictionary.Add emphasis_headers(ItemEmoOrigin), (ItemEmoOrigin.Column - 1)
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
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & emphasis_destiny.Name
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

      counter = (emo_origin_dictionary.Count - 2) / 3
      If (typeExams(charters(ItemData.Offset(, emo_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        Select Case numbers
          Case 1
            Call addNewRegister(tbl_emphasis.ListRows(1), counter, ItemData)
          Case Else
            Call addNewRegister(tbl_emphasis.ListRows.Add, counter, ItemData)
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
    For i = 3 to 71 Step 4
      Select Case numberMaxEmphasis
        Case 0
          Exit For
        Case Is > 0
          .Range(i) = charters(ItemData.Offset(, emo_origin_dictionary("ENFASIS_" & numberEmphasis)))
          .Range(i + 1) = emphasisConcepts(charters(ItemData.Offset(, emo_origin_dictionary("CONCEPTO AL ENFASIS_" & numberEmphasis))), charters(ItemData.Offset(, emo_origin_dictionary("ENFASIS_" & numberEmphasis))))
          .Range(i + 2) = charters(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES_AL_ENFASIS_" & numberEmphasis)))
          numberEmphasis = numberEmphasis + 1
          numberMaxEmphasis = numberMaxEmphasis - 1
      End Select
    Next i
  End With

End Sub