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
  Dim tbl_complementarios As Object, comple_origin As Object, SheetName As String
  
  SheetName = "COMPLEMENTARIOS"
  On Error GoTo com:
  Set comple_origin = origin.Worksheets(SheetName).Range("A1") '' COMPLEMENTARIOS DEL LIBRO ORIGEN ''
  On Error GoTo 0
  
  Set tbl_complementarios = comple_destiny.ListObjects("tbl_complementarios")
  Set comple_origin_dictionary = CreateObject("Scripting.Dictionary")

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each item In Range(comple_origin, comple_origin.End(xlToRight))
    If comple_origin_dictionary.Exists(comple_headers(item)) = False Then
      comple_origin_dictionary.Add comple_headers(item), item.Column
    End If
  Next item

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$12").value
  counts = Ubound(origin.Worksheets(SheetName).Range("A1").CurrentRegion.Value, 1) - 1
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For Each item In Range(comple_origin.offset(1, 0), comple_origin.offset(1, 0).End(xlDown))
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

      If (typeExams(charters(item.Offset(, comple_origin_dictionary("TIPO EXAMEN") - 1))) <> "EGRESO") Then
        If item.value <> "" And item.Row = 2 Then
          Call addNewRegister(tbl_complementarios.ListRows(1), aumentFromID, item)
          DoEvents
        ElseIf item.value <> "" And item.Row > 2 Then
          aumentFromID = aumentFromID + 1
          Call addNewRegister(tbl_complementarios.ListRows.Add, aumentFromID, item)
          DoEvents
        ElseIf item.value = "" Or item.value = VbNullString Then
          Exit For
        End If
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
      End If
    Next item
  End With

  Call dataDuplicate(comple_destiny.Range("tbl_complementarios[[#Data],[NRO IDENFICACION]]"))
  Call formatter(comple_destiny.Range("tbl_complementarios[[#Data],[NRO IDENFICACION]]"))

  Set comple_origin = Nothing
  comple_origin_dictionary.RemoveAll

  Exit Sub

com:
  SheetName = "COMPLEMENTARIO"
  comple_origin = origin.Worksheets(SheetName).Range("A1")
  Resume Next
End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrement As LongPtr, ByVal information As Object)

  With table
    .Range(1) = charters(information(, comple_origin_dictionary("NRO IDENFICACION")))
    .Range(2) = typeComplements(charters(information(, comple_origin_dictionary("PROCEDIMIENTO"))))
    .Range(3) = charters(information(, comple_origin_dictionary("DIAG_ PPAL")))
    .Range(4) = charters(information(, comple_origin_dictionary("DIAG_ PPAL OBS")))
    .Range(5) = charters(information(, comple_origin_dictionary("DIAG_ REL/1")))
    .Range(6) = charters(information(, comple_origin_dictionary("DIAG_ REL/2")))
    .Range(7) = charters(information(, comple_origin_dictionary("DIAG_ REL/3")))
    .Range(8) = charters(information(, comple_origin_dictionary("HALLAZGOS")))
    .Range(10) = autoIncrement
  End With

End Sub