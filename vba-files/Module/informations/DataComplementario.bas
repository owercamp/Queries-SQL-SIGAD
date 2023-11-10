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
Dim  aumentFromID As LongPtr
Public Sub ComplementarioData(ByVal name_sheet As String)
  Dim comple_destiny_dictionary As Scripting.Dictionary
  Dim comple_origin_dictionary As Scripting.Dictionary
  Dim comple_destiny_header As Object, comple_origin_header As Object, comple_origin_value As Object
  Dim ItemCompleDestiny As Object, ItemCompleOrigin As Object, ItemData As Object, comple_origin As Object

  Set comple_origin = origin.Worksheets(name_sheet) '' COMPLEMENTARIOS DEL LIBRO ORIGEN ''

  comple_destiny.Select
  comple_destiny.Range("$A4").Select
  Set comple_destiny_header = comple_destiny.Range("$A3", comple_destiny.Range("$A3").End(xlToRight))
  Set comple_origin_header = comple_origin.Range("$A1", comple_origin.Range("$A1").End(xlToRight))
  Set comple_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set comple_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (comple_origin.Range("$A2") <> Empty And comple_origin.Range("$A3") <> Empty) Then
    Set comple_origin_value = comple_origin.Range("$A2", comple_origin.Range("$A2").End(xlDown))
  ElseIf (comple_origin.Range("$A2") <> Empty And comple_origin.Range("$A3") = Empty) Then
    Set comple_origin_value = comple_origin.Range("$A2")
  End If

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  Dim value_data As String
  For Each ItemCompleDestiny In comple_destiny_header
    value_data = comple_headers(ItemCompleDestiny)
    If comple_destiny_dictionary.Exists(value_data) = False And value_data <> Empty Then
      comple_destiny_dictionary.Add value_data, (ItemCompleDestiny.Column - 1)
    End If
  Next ItemCompleDestiny
  
  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemCompleOrigin In comple_origin_header
    value_data = comple_headers(ItemCompleOrigin)
    If comple_origin_dictionary.Exists(value_data) = False And value_data <> Empty Then
      comple_origin_dictionary.Add value_data, (ItemCompleOrigin.Column - 1)
    End If
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

  Dim type_exam As String
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

      type_exam = typeExams(Trim(ItemData.Offset(, comple_origin_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        ActiveCell.Offset(, comple_destiny_dictionary("NRO IDENFICACION")) = Trim(ItemData.Offset(, comple_origin_dictionary("NRO IDENFICACION")))
        ActiveCell.Offset(, comple_destiny_dictionary("PROCEDIMIENTO")) = typeComplements(Trim(UCase(ItemData.Offset(, comple_origin_dictionary("PROCEDIMIENTO")))))
        ActiveCell.Offset(, comple_destiny_dictionary("DIAG_ PPAL")) = Trim(UCase(ItemData.Offset(, comple_origin_dictionary("DIAG_ PPAL"))))
        ActiveCell.Offset(, comple_destiny_dictionary("DIAG_ PPAL OBS")) = Trim(UCase(ItemData.Offset(, comple_origin_dictionary("DIAG_ PPAL OBS"))))
        ActiveCell.Offset(, comple_destiny_dictionary("DIAG_ REL/1")) = Trim(ItemData.Offset(, comple_origin_dictionary("DIAG_ REL/1")))
        ActiveCell.Offset(, comple_destiny_dictionary("DIAG_ REL/2")) = Trim(ItemData.Offset(, comple_origin_dictionary("DIAG_ REL/2")))
        ActiveCell.Offset(, comple_destiny_dictionary("DIAG_ REL/3")) = Trim(ItemData.Offset(, comple_origin_dictionary("DIAG_ REL/3")))
        ActiveCell.Offset(, comple_destiny_dictionary("HALLAZGOS")) = Trim(ItemData.Offset(, comple_origin_dictionary("HALLAZGOS")))
        If (ActiveCell.Row <> 4) Then
          aumentFromID = aumentFromID + 1
        End If
        ActiveCell.Offset(, comple_destiny_dictionary("ID_COMPLEMENTARIOS")) = aumentFromID
        ActiveCell.Offset(1, 0).Select
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents
      End If
    Next ItemData
  End With

  Call dataDuplicate(comple_destiny.Range("tbl_complementarios[[#Data],[NRO IDENFICACION]]"))
  Call formatter(comple_destiny.Range("tbl_complementarios[[#Data],[NRO IDENFICACION]]"))

  Set comple_origin_value = Nothing
  Set comple_destiny_header = Nothing
  Set comple_origin_header = Nothing
  comple_destiny_dictionary.RemoveAll
  comple_origin_dictionary.RemoveAll

End Sub