Attribute VB_Name = "DataEmoEmphasis"
Option Explicit

' DataEmphasisEmo - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'------------------------------------------------------------------------------------------------------------------
' Variables:
' - emphasis_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de destino.
' - emo_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de origen.
' - emphasis_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
' - emo_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
' - emo_origin_value: Una variable de objeto para almacenar el rango de los datos de diagnosticos de la hoja de origen.
' - ItemEmphasisDestiny: Una variable variante para iterar a traves del rango del encabezado de la hoja de destino.
' - ItemEmoOrigin: Una variable variante para iterar a traves del rango del encabezado de la hoja de origen.
' - ItemData: Una variable variante para iterar a traves del rango de los datos de diagnosticos de la hoja de origen.
' - numbers: Una variable numerica para hacer un seguimiento del número de elementos de datos importados.
' - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
' - counts: Una variable numerica para almacenar el número total de elementos de datos de audio.
' - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
' - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
' - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'------------------------------------------------------------------------------------------------------------------
Public Sub DataEmphasisEmo()

  Dim emphasis_destiny_dictionary As Scripting.Dictionary
  Dim emo_origin_dictionary As Scripting.Dictionary
  Dim emphasis_destiny_header As Object, emo_origin_header As Object, emo_origin_value As Object
  Dim ItemEmphasisDestiny As Variant, ItemEmoOrigin As Variant, ItemData As Variant

  Call deleteFormatConditions
  Set emo_origin = origin.Worksheets("EMO") '' EMO DEL LIBRO ORIGEN ''
  emphasis_destiny.Select
  ActiveSheet.Range("A5").Select
  Set emphasis_destiny_header = emphasis_destiny.Range("A4", emphasis_destiny.Range("A4").End(xlToRight))
  Set emo_origin_header = emo_origin.Range("A1", emo_origin.Range("A1").End(xlToRight))
  Set emphasis_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (emo_origin.Range("A2") <> Empty And emo_origin.Range("A3") <> Empty) Then
    Set emo_origin_value = emo_origin.Range("A2", emo_origin.Range("A2").End(xlDown))
  ElseIf (emo_origin.Range("A2") <> Empty And emo_origin.Range("A3") = Empty) Then
    Set emo_origin_value = emo_origin.Range("A2")
  End If

  ''   En los diccionarios de "emphasis_destiny_dictionary" y  "emo_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  x = 1
  For Each ItemEmphasisDestiny In emphasis_destiny_header
    On Error GoTo emphasisError
    emphasis_destiny_dictionary.Add emphasis_headers(ItemEmphasisDestiny), (ItemEmphasisDestiny.Column - 1)
  Next ItemEmphasisDestiny

  x = 1
  For Each ItemEmoOrigin In emo_origin_header
    On Error GoTo emphasisError
    emo_origin_dictionary.Add emphasis_headers(ItemEmoOrigin), (ItemEmoOrigin.Column - 1)
  Next ItemEmoOrigin

  numbers = 1
  porcentaje = 0
  counts = emo_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  For Each ItemData In emo_origin_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & emphasis_destiny.Name
      porcentaje = porcentaje + vals
      porcentajeGeneral = porcentajeGeneral + valsGeneral
      formImports.ProgressBarOneforOne.Width = oneForOne
      formImports.ProgressBarGeneral.Width = generalAll
      formImports.porcentageGeneral.Caption = CStr(VBA.Round(porcentajeGeneral * 100, 1)) & "%"
      formImports.porcentageOneoforOne.Caption = CStr(VBA.Round(porcentaje * 100, 1)) & "%"
      formImports.Caption = CStr(nameCompany)
      If formImports.ProgressBarGeneral.Width > (formImports.content_ProgressBarGeneral.Width / 2) Then
        formImports.porcentageGeneral.ForeColor = RGB(255, 255, 255)
      End If
      If formImports.ProgressBarGeneral.Width < (formImports.content_ProgressBarGeneral.Width / 2) Then
        formImports.porcentageGeneral.ForeColor = RGB(0, 0, 0)
      End If
      If formImports.ProgressBarOneforOne.Width > (formImports.content_ProgressBarOneforOne.Width / 2) Then
        formImports.porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
      End If
      If formImports.ProgressBarOneforOne.Width < (formImports.content_ProgressBarOneforOne.Width / 2) Then
        formImports.porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
      End If
      If (typeExams(charters(ItemData.Offset(, emo_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        With ActiveCell
          .Offset(, emphasis_destiny_dictionary("IDENTIFICACION")) = charters(ItemData.Offset(, emo_origin_dictionary("IDENTIFICACION")))
          For i = 1 To ((emo_origin_dictionary.Count - 2) / 3)
            .Offset(, emphasis_destiny_dictionary("ENFASIS_" & i)) = charters(ReplaceNonAlphaNumeric(ItemData.Offset(, emo_origin_dictionary("ENFASIS_" & i))))
            .Offset(, emphasis_destiny_dictionary("CONCEPTO AL ENFASIS_" & i)) = emphasisConcepts(charters(ReplaceNonAlphaNumeric(ItemData.Offset(, emo_origin_dictionary("CONCEPTO AL ENFASIS_" & i)))), charters(ReplaceNonAlphaNumeric(ItemData.Offset(, emo_origin_dictionary("ENFASIS_" & i)))))
            .Offset(, emphasis_destiny_dictionary("OBSERVACIONES_AL_ENFASIS_" & i)) = charters(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES_AL_ENFASIS_" & i)))
          Next i
          .Offset(1, 0).Select
        End With
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData

    Call dataDuplicate("$A5")
    Call formatter("$A5")

    Set emphasis_destiny_header = Nothing
    Set emo_origin_header = Nothing
    Set emo_origin_value = Nothing
    emphasis_destiny_dictionary.RemoveAll
    emo_origin_dictionary.RemoveAll

 emphasisError:
    Resume Next
End Sub
