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
  Dim tbl_emphasis As Object, emo_origin As Object, counter As LongPtr
  
  Set emo_origin = origin.Worksheets("EMO").Range("A1") '' EMO DEL LIBRO ORIGEN ''
  
  Set tbl_emphasis = emphasis_destiny.ListObjects("tbl_enfasis")
  Set emo_origin_dictionary = CreateObject("Scripting.Dictionary")

  x = 1
  For Each item In Range(emo_origin, emo_origin.End(xlToRight))
    If emo_origin_dictionary.Exists(emphasis_headers(item)) = False Then
      emo_origin_dictionary.Add emphasis_headers(item), item.Column
      If emo_origin_dictionary.Exists("OBSERVACIONES_AL_ENFASIS_" & x) = True Then
        x = x + 1
      End If
    End If
  Next item

  numbers = 1
  porcentaje = 0
  
  counts = Ubound(origin.Worksheets("EMO").Range("A1").CurrentRegion.Value, 1) - 1
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each item In Range(emo_origin.offset(1, 0), emo_origin.offset(1, 0).End(xlDown))
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
      If (typeExams(charters(item.Offset(, emo_origin_dictionary("TIPO EXAMEN") - 1))) <> "EGRESO") Then
        If item.Value <> "" And item.Row = 2 Then
          Call addNewRegister(tbl_emphasis.ListRows(1), counter - 1, item)
          DoEvents
        ElseIf item.Value <> "" And item.Row > 2 Then
          Call addNewRegister(tbl_emphasis.ListRows.Add, counter - 1, item)
          DoEvents
        ElseIf item.Value = "" Or item.Value = VbNullString Then
          Exit For
        End If
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
      End If
    Next item
  End With

  Call dataDuplicate(emphasis_destiny.Range("tbl_enfasis[[#Data],[IDENTIFICACION]]"))
  Call formatter(emphasis_destiny.Range("tbl_enfasis[[#Data],[IDENTIFICACION]]"))

  Set emo_origin = Nothing
  emo_origin_dictionary.RemoveAll

End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal numberMaxEmphasis As LongPtr, ByVal information As Object)

  Dim numberEmphasis As LongPtr
  numberEmphasis = 1
  With table
    .Range(1) = charters(information(, emo_origin_dictionary("IDENTIFICACION")))
    For i = 3 to 71 Step 4
      Select Case numberMaxEmphasis
        Case 0
          Exit For
        Case Is > 0
          .Range(i) = charters(information(, emo_origin_dictionary("ENFASIS_" & numberEmphasis)))
          .Range(i + 1) = emphasisConcepts(charters(information(, emo_origin_dictionary("CONCEPTO AL ENFASIS_" & numberEmphasis))), charters(information(, emo_origin_dictionary("ENFASIS_" & numberEmphasis))))
          .Range(i + 2) = charters(information(, emo_origin_dictionary("OBSERVACIONES_AL_ENFASIS_" & numberEmphasis)))
          numberEmphasis = numberEmphasis + 1
          numberMaxEmphasis = numberMaxEmphasis - 1
      End Select
    Next i
  End With

End Sub