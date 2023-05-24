Attribute VB_Name = "DataPsicotecnica"
Option Explicit

' PsicotecnicaData - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'------------------------------------------------------------------------------------------------------------------
' Variables:
' - psicotecnica_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de destino.
' - psicotecnica_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los números de columna de la hoja de origen.
' - psicotecnica_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
' - psicotecnica_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
' - psicotecnica_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
' - numbers: Una variable numerica para hacer un seguimiento del número de elementos de datos importados.
' - porcentaje: Una variable numerica para calcular el porcentaje de elementos de datos importados.
' - counts: Una variable numerica para almacenar el número total de elementos de datos de audio.
' - vals: Una variable numerica para calcular el valor de incremento de la barra de progreso.
' - oneForOne: Una variable numerica para hacer un seguimiento del progreso de la barra de progreso para cada elemento de datos.
' - widthOneforOne: Una variable numerica para calcular el ancho de la barra de progreso para cada elemento de datos.
'------------------------------------------------------------------------------------------------------------------
Public Sub PsicotecnicaData()
  Dim psicotecnica_destiny_dictionary As Scripting.Dictionary
  Dim psicotecnica_origin_dictionary As Scripting.Dictionary
  Dim psicotecnica_destiny_header As Object, psicotecnica_origin_header As Object, psicotecnica_origin_value As Object
  Dim ItemPsicotecnicaDestiny As Variant, ItemPsicotecnicaOrigin As Variant, ItemData As Variant

  On Error GoTo tecnica:
  Set psico_origin = origin.Worksheets("PSICOTECNICA") '' PSICOTECNICA DEL LIBRO ORIGEN ''

  psico_destiny.Select
  ActiveSheet.Range("A2").Select
  Set psicotecnica_destiny_header = psico_destiny.Range("A1", psico_destiny.Range("A1").End(xlToRight))
  Set psicotecnica_origin_header = psico_origin.Range("A1", psico_origin.Range("A1").End(xlToRight))
  Set psicotecnica_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set psicotecnica_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (psico_origin.Range("A2") <> Empty And psico_origin.Range("A3") <> Empty) Then
    Set psicotecnica_origin_value = psico_origin.Range("A2", psico_origin.Range("A2").End(xlDown))
  ElseIf (psico_origin.Range("A2") <> Empty And psico_origin.Range("A3") = Empty) Then
    Set psicotecnica_origin_value = psico_origin.Range("A2")
  End If

  ''   En los diccionarios de "psicotecnica_destiny_dictionary" y  "psicotecnica_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemPsicotecnicaDestiny In psicotecnica_destiny_header
    On Error GoTo psicotecnicaError
    psicotecnica_destiny_dictionary.Add psicotecnica_headers(ItemPsicotecnicaDestiny), (ItemPsicotecnicaDestiny.Column - 1)
  Next ItemPsicotecnicaDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemPsicotecnicaOrigin In psicotecnica_origin_header
    On Error GoTo psicotecnicaError
    psicotecnica_origin_dictionary.Add psicotecnica_headers(ItemPsicotecnicaOrigin), (ItemPsicotecnicaOrigin.Column - 1)
  Next ItemPsicotecnicaOrigin

  numbers = 1
  porcentaje = 0
  counts = psicotecnica_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  For Each ItemData In psicotecnica_origin_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & psico_destiny.Name
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
      If (typeExams(charters(ItemData.Offset(, psicotecnica_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        With ActiveCell
          .Offset(, psicotecnica_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, psicotecnica_origin_dictionary("NRO IDENFICACION")))
          .Offset(, psicotecnica_destiny_dictionary("PACIENTE")) = charters(ItemData.Offset(, psicotecnica_origin_dictionary("PACIENTE")))
          .Offset(, psicotecnica_destiny_dictionary("PRUEBA PSICOTECNICA")) = charters(ItemData.Offset(, psicotecnica_origin_dictionary("PRUEBA PSICOTECNICA")))
          .Offset(, psicotecnica_destiny_dictionary("DIAGNOSTICO PPAL (CUMPLE, NO CUMPLE)")) = charters(ItemData.Offset(, psicotecnica_origin_dictionary("DIAGNOSTICO PPAL (CUMPLE, NO CUMPLE)")))
          .Offset(, psicotecnica_destiny_dictionary("DIAGNOSTICO OBS")) = charters(ItemData.Offset(, psicotecnica_origin_dictionary("DIAGNOSTICO OBS")))
          If (.Row = 2) Then
            .Offset(, psicotecnica_destiny_dictionary("ID_PSICOTECNICA")) = Trim$(ThisWorkbook.Worksheets("RUTAS").Range("$F$13").value)
          Else
            .Offset(, psicotecnica_destiny_dictionary("ID_PSICOTECNICA")) = .Offset(-1, psicotecnica_destiny_dictionary("ID_PSICOTECNICA")) + 1
          End If
          .Offset(1, 0).Select
        End With
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData

    Call meetsfails("$D2")
    Call formatter("$A2")

    Set psicotecnica_origin_value = Nothing
    Set psicotecnica_destiny_header = Nothing
    Set psicotecnica_origin_header = Nothing
    psicotecnica_destiny_dictionary.RemoveAll
    psicotecnica_origin_dictionary.RemoveAll

    Exit Sub

 psicotecnicaError:
    Resume Next

 tecnica:
    Set psico_origin = origin.Worksheets("PSICOLOGIA")
    Resume Next
End Sub
