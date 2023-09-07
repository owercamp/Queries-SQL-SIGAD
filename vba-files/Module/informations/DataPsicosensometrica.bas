Attribute VB_Name = "DataPsicosensometrica"
'namespace=vba-files\Module\informations
Option Explicit

'TODO: PsicosensometricaData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - psicosensometrica_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - psicosensometrica_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - psicosensometrica_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - psicosensometrica_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - psicosensometrica_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
'* - numbers: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
'* - counts: Una variable numerica para almacenar el numero total de elementos de datos de audio.
'* - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
'* - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
'* - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'* ------------------------------------------------------------------------------------------------------------------
Dim psicosensometrica_origin_dictionary As Scripting.Dictionary
Dim aumentFromID As LongPtr
Public Sub PsicosensometricaData()
  Dim tbl_psicosensometrica As Object, xNumber As Long, senso_origin As Variant
  
  On Error GoTo metrica:
  senso_origin = origin.Worksheets("PSICOSENSOMETRICA").Range("A1").CurrentRegion.value '' PSICOSENSOMETRICA DEL LIBRO ORIGEN ''
  On Error GoTo 0
  
  senso_destiny.Select
  Set tbl_psicosensometrica = ActiveSheet.ListObjects("tbl_psicosensometrica")
  Set psicosensometrica_origin_dictionary = CreateObject("Scripting.Dictionary")

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For xNumber = 1 To Ubound(senso_origin, 2)
    On Error Resume Next
    psicosensometrica_origin_dictionary.Add psicosensometrica_headers(senso_origin(1, xNumber)), xNumber
    On Error GoTo 0    
  Next xNumber

  numbers = 1
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$14").value
  counts = Ubound(senso_origin, 1) - 1
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  With formImports
    For xNumber = 2 To Ubound(senso_origin, 1)
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & senso_destiny.Name
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
      
      If (typeExams(charters(senso_origin(xNumber, psicosensometrica_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        Select Case numbers
          Case 1
            Call addNewRegister(tbl_psicosensometrica.ListRows(1), aumentFromID, senso_origin, xNumber)
            DoEvents
          Case Else
            aumentFromID = aumentFromID + 1
            Call addNewRegister(tbl_psicosensometrica.ListRows.Add, aumentFromID, senso_origin, xNumber)
            DoEvents
        End Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
    Next xNumber
  End With

  range("$I3:$N3").Select
  Call greaterThanOne
  range("$I3:$N3").Select
  Call iqualCero
  range("$A3", range("$A3").End(xlDown)).Select
  Call formatter

  Set senso_origin = Nothing
  psicosensometrica_origin_dictionary.RemoveAll

  Exit Sub

metrica:
  senso_origin = origin.Worksheets("PSICOMOTRIZ").Range("A1").CurrentRegion.value
  Resume Next
End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrement As LongPtr, ByVal information As Variant, ByVal x As Long)

  With table
    .Range(1) = charters(information(x, psicosensometrica_origin_dictionary("NRO IDENFICACION")))
    .Range(2) = charters(information(x, psicosensometrica_origin_dictionary("PACIENTE")))
    .Range(3) = charters(information(x, psicosensometrica_origin_dictionary("PRUEBA PSICOSENSOMETRICA")))
    .Range(4) = charters(information(x, psicosensometrica_origin_dictionary("DIAGNOSTICO PPAL")))
    .Range(5) = charters(information(x, psicosensometrica_origin_dictionary("DIAGNOSTICO OBS")))
    .Range(6) = charters(information(x, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/1")))
    .Range(7) = charters(information(x, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/2")))
    .Range(8) = charters(information(x, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/3")))
    .Range(9) = charters(information(x, psicosensometrica_origin_dictionary("CONTROLES MENSUALES")))
    .Range(10) = charters(information(x, psicosensometrica_origin_dictionary("CONTROLES BIMENSUAL")))
    .Range(11) = charters(information(x, psicosensometrica_origin_dictionary("CONTROLES TRIMESTRALES")))
    .Range(12) = charters(information(x, psicosensometrica_origin_dictionary("CONTROLES 6 MESES")))
    .Range(13) = charters(information(x, psicosensometrica_origin_dictionary("CONTROLES 1 ANO")))
    .Range(14) = charters(information(x, psicosensometrica_origin_dictionary("CONTROLES CONFIRMATORIA")))
    .Range(17) = autoIncrement
  End With

End Sub