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
Public Sub DataDiagnosticsEmo(ByVal name_sheet As String)
  Dim diagnostics_destiny_dictionary As Scripting.Dictionary
  Dim emo_origin_dictionary As Scripting.Dictionary
  Dim diagnostics_destiny_header As Object, emo_origin_header As Object, emo_origin_value As Object
  Dim ItemDiagnosticsDestiny As Object, ItemEmoOrigin As Object, ItemData As Object, x As Long, emo_origin As Object, cell_active as Range

  Set emo_origin = origin.Worksheets(name_sheet) '' EMO DEL LIBRO ORIGEN ''
  diagnostics_destiny.Select
  diagnostics_destiny.Range("$A5").Select
  Set cell_active = ActiveCell

  Set diagnostics_destiny_header = diagnostics_destiny.Range("$A4", diagnostics_destiny.Range("$A4").End(xlToRight))
  Set emo_origin_header = emo_origin.Range("$A1", emo_origin.Range("$A1").End(xlToRight))
  Set diagnostics_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (emo_origin.Range("$A2") <> Empty And emo_origin.Range("$A3") <> Empty) Then
    Set emo_origin_value = emo_origin.Range("$A2", emo_origin.Range("$A2").End(xlDown))
  ElseIf (emo_origin.Range("$A2") <> Empty And emo_origin.Range("$A3") = Empty) Then
    Set emo_origin_value = emo_origin.Range("$A2")
  End If

  Dim value_data As String
  x = 1
  For Each ItemDiagnosticsDestiny In diagnostics_destiny_header
    value_data = diagnostics_header(ItemDiagnosticsDestiny, x)
    If diagnostics_destiny_dictionary.Exists(value_data) = False And value_data <> Empty Then
      diagnostics_destiny_dictionary.Add value_data, (ItemDiagnosticsDestiny.Column - 1)
      If diagnostics_destiny_dictionary.Exists("DIAG REL " & x) = True Then
        x = x + 1
      End If
    End If
  Next ItemDiagnosticsDestiny
  
  x = 1
  For Each ItemEmoOrigin In emo_origin_header
    value_data = diagnostics_header(ItemEmoOrigin, x)
    If emo_origin_dictionary.Exists(value_data) = False And value_data <> Empty Then
      emo_origin_dictionary.Add value_data, (ItemEmoOrigin.Column - 1)
      If emo_origin_dictionary.Exists("DIAG REL " & x) = True Then
        x = x + 1
      End If
    End If
  Next ItemEmoOrigin

  numbers = 1
  porcentaje = 0
  
  counts = emo_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  Dim type_exam As String
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

      type_exam = typeExams(Trim(ItemData.Offset(, emo_origin_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        cell_active.Offset(, diagnostics_destiny_dictionary("IDENTIFICACION")) = Trim(ItemData.Offset(, emo_origin_dictionary("IDENTIFICACION")))
        cell_active.Offset(, diagnostics_destiny_dictionary("CODIGO DIAG PPAL")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("CODIGO DIAG PPAL"))))
        cell_active.Offset(, diagnostics_destiny_dictionary("DIAG PPAL")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("DIAG PPAL"))))
        For i = 1 To ((emo_origin_dictionary.Count - 5) / 2)
          cell_active.Offset(, diagnostics_destiny_dictionary("CODIGO DIAG REL" & i)) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("CODIGO DIAG REL" & i))))
          cell_active.Offset(, diagnostics_destiny_dictionary("DIAG REL " & i)) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("DIAG REL " & i))))
        Next i
        Set cell_active = cell_active.Offset(1, 0)
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents
      End If
    Next ItemData
  End With

  Call dataDuplicate(diagnostics_destiny.Range("tbl_diagnosticos[[#Data],[IDENTIFICACION]]"))
  Call formatter(diagnostics_destiny.Range("tbl_diagnosticos[[#Data],[IDENTIFICACION]]"))

  Set diagnostics_destiny_header = Nothing
  Set emo_origin_header = Nothing
  Set emo_origin_value = Nothing
  diagnostics_destiny_dictionary.RemoveAll
  emo_origin_dictionary.RemoveAll
  
End Sub