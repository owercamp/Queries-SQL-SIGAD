Attribute VB_Name = "DataPsicosensometrica"
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
Public Sub PsicosensometricaData()
  Dim psicosensometrica_destiny_dictionary As Scripting.Dictionary
  Dim psicosensometrica_origin_dictionary As Scripting.Dictionary
  Dim psicosensometrica_destiny_header As Object, psicosensometrica_origin_header As Object, psicosensometrica_origin_value As Object
  Dim ItemPsicosensometricaDestiny As Variant, ItemPsicosensometricaOrigin As Variant, ItemData As Variant

  On Error GoTo metrica:
  Set senso_origin = origin.Worksheets("PSICOSENSOMETRICA") '' PSICOSENSOMETRICA DEL LIBRO ORIGEN ''

  senso_destiny.Select
  ActiveSheet.Range("A3").Select
  Set psicosensometrica_destiny_header = senso_destiny.Range("A2", senso_destiny.Range("A2").End(xlToRight))
  Set psicosensometrica_origin_header = senso_origin.Range("A1", senso_origin.Range("A1").End(xlToRight))
  Set psicosensometrica_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set psicosensometrica_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (senso_origin.Range("A2") <> Empty And senso_origin.Range("A3") <> Empty) Then
    Set psicosensometrica_origin_value = senso_origin.Range("A2", senso_origin.Range("A2").End(xlDown))
  ElseIf (senso_origin.Range("A2") <> Empty And senso_origin.Range("A3") = Empty) Then
    Set psicosensometrica_origin_value = senso_origin.Range("A2")
  End If

  ''   En los diccionarios de "psicosensometrica_destiny_dictionary" y  "psicosensometrica_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  For Each ItemPsicosensometricaDestiny In psicosensometrica_destiny_header
    On Error GoTo psicotecnicaError
    psicosensometrica_destiny_dictionary.Add psicosensometrica_headers(ItemPsicosensometricaDestiny), (ItemPsicosensometricaDestiny.Column - 1)
  Next ItemPsicosensometricaDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemPsicosensometricaOrigin In psicosensometrica_origin_header
    On Error GoTo psicotecnicaError
    psicosensometrica_origin_dictionary.Add psicosensometrica_headers(ItemPsicosensometricaOrigin), (ItemPsicosensometricaOrigin.Column - 1)
  Next ItemPsicosensometricaOrigin

  numbers = 1
  porcentaje = 0
  counts = psicosensometrica_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  For Each ItemData In psicosensometrica_origin_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & senso_destiny.Name
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
      If (typeExams(charters(ItemData.Offset(, psicosensometrica_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        With ActiveCell
          .Offset(, psicosensometrica_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("NRO IDENFICACION")))
          .Offset(, psicosensometrica_destiny_dictionary("PACIENTE")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("PACIENTE")))
          .Offset(, psicosensometrica_destiny_dictionary("PRUEBA PSICOSENSOMETRICA")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("PRUEBA PSICOSENSOMETRICA")))
          .Offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO PPAL")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO PPAL")))
          .Offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO OBS")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO OBS")))
          .Offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/1")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/1")))
          .Offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/2")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/2")))
          .Offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/3")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("DIAGNOSTICO REL/3")))
          .Offset(, psicosensometrica_destiny_dictionary("CONTROLES MENSUALES")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES MENSUALES")))
          .Offset(, psicosensometrica_destiny_dictionary("CONTROLES BIMENSUAL")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES BIMENSUAL")))
          .Offset(, psicosensometrica_destiny_dictionary("CONTROLES TRIMESTRALES")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES TRIMESTRALES")))
          .Offset(, psicosensometrica_destiny_dictionary("CONTROLES 6 MESES")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES 6 MESES")))
          .Offset(, psicosensometrica_destiny_dictionary("CONTROLES 1 ANO")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES 1 ANO")))
          .Offset(, psicosensometrica_destiny_dictionary("CONTROLES CONFIRMATORIA")) = charters(ItemData.Offset(, psicosensometrica_origin_dictionary("CONTROLES CONFIRMATORIA")))
          If (.Row = 3) Then
            .Offset(, psicosensometrica_destiny_dictionary("ID_PSICOSENSOMETRICA")) = Trim$(ThisWorkbook.Worksheets("RUTAS").Range("$F$14").value)
          Else
            .Offset(, psicosensometrica_destiny_dictionary("ID_PSICOSENSOMETRICA")) = .Offset(-1, psicosensometrica_destiny_dictionary("ID_PSICOSENSOMETRICA")) + 1
          End If
          .Offset(1, 0).Select
        End With
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData

    Call greaterThanOne("$I3:$N3")
    Call iqualCero("$I3:$N3")
    Call formatter("$A3")

    Set psicosensometrica_origin_value = Nothing
    Set psicosensometrica_destiny_header = Nothing
    Set psicosensometrica_origin_header = Nothing
    psicosensometrica_destiny_dictionary.RemoveAll
    psicosensometrica_origin_dictionary.RemoveAll

    Exit Sub

 psicotecnicaError:
    Resume Next
 metrica:
    Set senso_origin = origin.Worksheets("PSICOMOTRIZ")
    Resume Next
End Sub
