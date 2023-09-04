Attribute VB_Name = "DataPsicotecnica"
'namespace=vba-files\Module\informations
Option Explicit

'TODO: PsicotecnicaData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - psicotecnica_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - psicotecnica_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - psicotecnica_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - psicotecnica_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - psicotecnica_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Dim psicotecnica_origin_dictionary As Scripting.Dictionary
Dim aumentFromID As LongPtr
Public Sub PsicotecnicaData()
  Dim tbl_psicotecnica As Object, xNumber As Long, psico_origin As Variant
  
  On Error GoTo tecnica:
  psico_origin = origin.Worksheets("PSICOTECNICA").Range("A1").CurrentRegion.value '' PSICOTECNICA DEL LIBRO ORIGEN ''
  On Error GoTo 0
  
  psico_destiny.Select
  Set tbl_psicotecnica = ActiveSheet.ListObjects("tbl_psicotecnica")
  Set psicotecnica_origin_dictionary = CreateObject("Scripting.Dictionary")

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For xNumber = 1 To Ubound(psico_origin, 2)
    On Error Resume Next
    psicotecnica_origin_dictionary.Add psicotecnica_headers(psico_origin(1, xNumber)), xNumber
    On Error GoTo 0    
  Next xNumber

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$13").value
  counts = Ubound(psico_origin, 1) - 1
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For xNumber = 2 To Ubound(psico_origin, 1)
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & psico_destiny.Name
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
      
      If (typeExams(charters(psico_origin(xNumber, psicotecnica_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        Select Case numbers
          Case 1
            Call addNewRegister(tbl_psicotecnica.ListRows(1), aumentFromID, psico_origin, xNumber)
          Case Else
            aumentFromID = aumentFromID + 1
            Call addNewRegister(tbl_psicotecnica.ListRows.Add, aumentFromID, psico_origin, xNumber)
        End Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      Call addTimer
    Next xNumber
  End With

  range("D2").Select
  Call meetsfails
  range("$A2", range("$A2").End(xlDown)).Select
  Call formatter

  Set psico_origin = Nothing
  psicotecnica_origin_dictionary.RemoveAll

  Exit Sub

tecnica:
  psico_origin = origin.Worksheets("PSICOLOGIA").Range("A1").CurrentRegion.value
  Resume Next
End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrement As LongPtr, ByVal information As Variant, ByVal x As Long)

  With table
    .Range(1) = charters(information(x, psicotecnica_origin_dictionary("NRO IDENFICACION")))
    .Range(2) = charters(information(x, psicotecnica_origin_dictionary("PACIENTE")))
    .Range(3) = charters(information(x, psicotecnica_origin_dictionary("PRUEBA PSICOTECNICA")))
    .Range(4) = charters(information(x, psicotecnica_origin_dictionary("DIAGNOSTICO PPAL (CUMPLE, NO CUMPLE)")))
    .Range(5) = charters(information(x, psicotecnica_origin_dictionary("DIAGNOSTICO OBS")))
    .Range(7) = autoIncrement
  End With

End Sub