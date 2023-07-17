Attribute VB_Name = "DataComplementario"
'namespace=vba-files\Module\informations
Option Explicit

'TODO: ComplementarioData - Esta subrutina importa datos de complementario desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - comple_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - comple_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - comple_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - comple_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - comple_origin_value: Una variable de objeto para almacenar el rango de los datos de complementario de la hoja de origen.
'* - ItemCompleDestiny: Una variable variante para iterar a traves del rango del encabezado de la hoja de destino.
'* - ItemCompleOrigin: Una variable variante para iterar a traves del rango del encabezado de la hoja de origen.
'* - ItemData: Una variable variante para iterar a traves del rango de los datos de complementario de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Dim comple_origin_dictionary As Scripting.Dictionary
Dim  aumentFromID As LongPtr
Public Sub ComplementarioData()
  Dim tbl_complementarios As Object, comple_origin_header As Object, comple_origin_value As Object
  Dim ItemCompleOrigin As Variant, ItemData As Variant
  
  On Error GoTo com:
  Set comple_origin = origin.Worksheets("COMPLEMENTARIOS") '' COMPLEMENTARIOS DEL LIBRO ORIGEN ''
  On Error GoTo 0
  
  comple_destiny.Select
  Set tbl_complementarios = ActiveSheet.ListObjects("tbl_complementarios")
  Set comple_origin_header = comple_origin.range("A1", comple_origin.range("A1").End(xlToRight))
  Set comple_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (comple_origin.range("A2") <> Empty And comple_origin.range("A3") <> Empty) Then
    Set comple_origin_value = comple_origin.range("A2", comple_origin.range("A2").End(xlDown))
  ElseIf (comple_origin.range("A2") <> Empty And comple_origin.range("A3") = Empty) Then
    Set comple_origin_value = comple_origin.range("A2")
  End If

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemCompleOrigin In comple_origin_header
    On Error Resume Next
    comple_origin_dictionary.Add comple_headers(ItemCompleOrigin), (ItemCompleOrigin.Column - 1)
    On Error GoTo 0
  Next ItemCompleOrigin

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$12").value
  counts = comple_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each ItemData In comple_origin_value
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & comple_destiny.Name
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

      If (typeExams(charters(ItemData.Offset(, comple_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        Select Case numbers
          Case 1
            Call addNewRegister(tbl_complementarios.ListRows(1), aumentFromID, ItemData)
          Case Else
            aumentFromID = aumentFromID + 1
            Call addNewRegister(tbl_complementarios.ListRows.Add, aumentFromID, ItemData)
        End Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData
  End With

  range("$A4").Select
  Call dataDuplicate
  range("$A4", range("$A4").End(xlDown)).Select
  Call formatter

  Set comple_origin_value = Nothing
  Set comple_origin_header = Nothing
  comple_origin_dictionary.RemoveAll

  Exit Sub

com:
  Set comple_origin = origin.Worksheets("COMPLEMENTARIO")
  Resume Next
End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrement As LongPtr, ByVal ItemData As Variant)

  With table
    .Range(1) = charters(ItemData.Offset(, comple_origin_dictionary("NRO IDENFICACION")))
    .Range(2) = typeComplements(charters(ItemData.Offset(, comple_origin_dictionary("PROCEDIMIENTO"))))
    .Range(3) = charters(ItemData.Offset(, comple_origin_dictionary("DIAG_ PPAL")))
    .Range(4) = charters(ItemData.Offset(, comple_origin_dictionary("DIAG_ PPAL OBS")))
    .Range(5) = charters(ItemData.Offset(, comple_origin_dictionary("DIAG_ REL/1")))
    .Range(6) = charters(ItemData.Offset(, comple_origin_dictionary("DIAG_ REL/2")))
    .Range(7) = charters(ItemData.Offset(, comple_origin_dictionary("DIAG_ REL/3")))
    .Range(8) = charters(ItemData.Offset(, comple_origin_dictionary("HALLAZGOS")))
    .Range(10) = autoIncrement
    DoEvents
  End With

End Sub