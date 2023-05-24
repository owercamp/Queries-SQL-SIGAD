Attribute VB_Name = "DataComplementario"
Option Explicit

' ComplementarioData - Esta subrutina importa datos de complementario desde una hoja de origen a una hoja de destino.
'------------------------------------------------------------------------------------------------------------------
' Variables:
' - comple_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de destino.
' - comple_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de origen.
' - comple_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
' - comple_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
' - comple_origin_value: Una variable de objeto para almacenar el rango de los datos de complementario de la hoja de origen.
' - ItemCompleDestiny: Una variable variante para iterar a través del rango del encabezado de la hoja de destino.
' - ItemCompleOrigin: Una variable variante para iterar a través del rango del encabezado de la hoja de origen.
' - ItemData: Una variable variante para iterar a través del rango de los datos de complementario de la hoja de origen.
' - numbers: Una variable numerica para hacer un seguimiento del número de elementos de datos importados.
' - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
' - counts: Una variable numerica para almacenar el número total de elementos de datos de audio.
' - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
' - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
' - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'------------------------------------------------------------------------------------------------------------------
Public Sub ComplementarioData()
  Dim comple_destiny_dictionary As Scripting.Dictionary
  Dim comple_origin_dictionary As Scripting.Dictionary
  Dim comple_destiny_header As Object, comple_origin_header As Object, comple_origin_value As Object
  Dim ItemCompleDestiny As Variant, ItemCompleOrigin As Variant, ItemData As Variant

  On Error GoTo com:
  Set comple_origin = origin.Worksheets("COMPLEMENTARIOS") '' COMPLEMENTARIOS DEL LIBRO ORIGEN ''

  comple_destiny.Select
  ActiveSheet.Range("A4").Select
  Set comple_destiny_header = comple_destiny.Range("A3", comple_destiny.Range("A3").End(xlToRight))
  Set comple_origin_header = comple_origin.Range("A1", comple_origin.Range("A1").End(xlToRight))
  Set comple_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set comple_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (comple_origin.Range("A2") <> Empty And comple_origin.Range("A3") <> Empty) Then
    Set comple_origin_value = comple_origin.Range("A2", comple_origin.Range("A2").End(xlDown))
  ElseIf (comple_origin.Range("A2") <> Empty And comple_origin.Range("A3") = Empty) Then
    Set comple_origin_value = comple_origin.Range("A2")
  End If

  ''   En los diccionarios de "comple_destiny_dictionary" y  "comple_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''


  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemCompleDestiny In comple_destiny_header
    On Error GoTo compleError
    comple_destiny_dictionary.Add comple_headers(ItemCompleDestiny), (ItemCompleDestiny.Column - 1)
  Next ItemCompleDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemCompleOrigin In comple_origin_header
    On Error GoTo compleError
    comple_origin_dictionary.Add comple_headers(ItemCompleOrigin), (ItemCompleOrigin.Column - 1)
  Next ItemCompleOrigin

  numbers = 1
  porcentaje = 0
  counts = comple_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  For Each ItemData In comple_origin_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & comple_destiny.Name
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
      If (typeExams(charters(ItemData.Offset(, comple_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        With ActiveCell
          .Offset(, comple_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, comple_origin_dictionary("NRO IDENFICACION")))
          .Offset(, comple_destiny_dictionary("PROCEDIMIENTO")) = typeComplements(charters(ReplaceNonAlphaNumeric(ItemData.Offset(, comple_origin_dictionary("PROCEDIMIENTO")))))
          .Offset(, comple_destiny_dictionary("DIAG_ PPAL")) = charters(ItemData.Offset(, comple_origin_dictionary("DIAG_ PPAL")))
          .Offset(, comple_destiny_dictionary("DIAG_ PPAL OBS")) = charters(ItemData.Offset(, comple_origin_dictionary("DIAG_ PPAL OBS")))
          .Offset(, comple_destiny_dictionary("DIAG_ REL/1")) = charters(ItemData.Offset(, comple_origin_dictionary("DIAG_ REL/1")))
          .Offset(, comple_destiny_dictionary("DIAG_ REL/2")) = charters(ItemData.Offset(, comple_origin_dictionary("DIAG_ REL/2")))
          .Offset(, comple_destiny_dictionary("DIAG_ REL/3")) = charters(ItemData.Offset(, comple_origin_dictionary("DIAG_ REL/3")))
          .Offset(, comple_destiny_dictionary("HALLAZGOS")) = charters(ItemData.Offset(, comple_origin_dictionary("HALLAZGOS")))
          If (.Row = 4) Then
            .Offset(, comple_destiny_dictionary("ID_COMPLEMENTARIOS")) = Trim$(ThisWorkbook.Worksheets("RUTAS").Range("$F$12").value)
          Else
            .Offset(, comple_destiny_dictionary("ID_COMPLEMENTARIOS")) = .Offset(-1, comple_destiny_dictionary("ID_COMPLEMENTARIOS")) + 1
          End If
          .Offset(1, 0).Select
        End With
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData

    Call dataDuplicate("$A4")
    Call formatter("$A4")

    Set comple_origin_value = Nothing
    Set comple_destiny_header = Nothing
    Set comple_origin_header = Nothing
    comple_destiny_dictionary.RemoveAll
    comple_origin_dictionary.RemoveAll

 compleError:
    Resume Next
 com:
    Set comple_origin = origin.Worksheets("COMPLEMENTARIO")
    Resume Next
End Sub
