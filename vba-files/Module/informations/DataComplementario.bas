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
  Dim tbl_complementarios As Object, xNumber As Long, comple_origin As Variant
  
  On Error GoTo com:
  comple_origin = origin.Worksheets("COMPLEMENTARIOS").Range("A1").CurrentRegion.value '' COMPLEMENTARIOS DEL LIBRO ORIGEN ''
  On Error GoTo 0
  
  comple_destiny.Select
  Set tbl_complementarios = ActiveSheet.ListObjects("tbl_complementarios")
  Set comple_origin_dictionary = CreateObject("Scripting.Dictionary")

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For xNumber = 1 To Ubound(comple_origin, 2)
    On Error Resume Next
    comple_origin_dictionary.Add comple_headers(comple_origin(1, xNumber)), xNumber
    On Error GoTo 0
  Next xNumber

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$12").value
  counts = Ubound(comple_origin, 1) - 1
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For xNumber = 2 To Ubound(comple_origin, 1)
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

      If (typeExams(charters(comple_origin(xNumber, comple_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        Select Case numbers
          Case 1
            Call addNewRegister(tbl_complementarios.ListRows(1), aumentFromID, comple_origin, xNumber)
            DoEvents
          Case Else
            aumentFromID = aumentFromID + 1
            Call addNewRegister(tbl_complementarios.ListRows.Add, aumentFromID, comple_origin, xNumber)
            DoEvents
        End Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
    Next xNumber
  End With

  range("$A4").Select
  Call dataDuplicate
  range("$A4", range("$A4").End(xlDown)).Select
  Call formatter

  Set comple_origin = Nothing
  comple_origin_dictionary.RemoveAll

  Exit Sub

com:
  comple_origin = origin.Worksheets("COMPLEMENTARIO").Range("A1").CurrentRegion.value
  Resume Next
End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrement As LongPtr, ByVal information As Variant, ByVal x As Long)

  With table
    .Range(1) = charters(information(x, comple_origin_dictionary("NRO IDENFICACION")))
    .Range(2) = typeComplements(charters(information(x, comple_origin_dictionary("PROCEDIMIENTO"))))
    .Range(3) = charters(information(x, comple_origin_dictionary("DIAG_ PPAL")))
    .Range(4) = charters(information(x, comple_origin_dictionary("DIAG_ PPAL OBS")))
    .Range(5) = charters(information(x, comple_origin_dictionary("DIAG_ REL/1")))
    .Range(6) = charters(information(x, comple_origin_dictionary("DIAG_ REL/2")))
    .Range(7) = charters(information(x, comple_origin_dictionary("DIAG_ REL/3")))
    .Range(8) = charters(information(x, comple_origin_dictionary("HALLAZGOS")))
    .Range(10) = autoIncrement
  End With

End Sub