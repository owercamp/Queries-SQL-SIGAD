Attribute VB_Name = "DataPsicosensometrica"
Option Explicit

Sub PsicosensometricaData()
  Dim psicosensometrica_destiny_dictionary As Scripting.Dictionary
  Dim psicosensometrica_origin_dictionary As Scripting.Dictionary
  Dim psicosensometrica_destiny_header, psicosensometrica_origin_header, psicosensometrica_origin_value As Object
  Dim ItemPsicosensometricaDestiny, ItemPsicosensometricaOrigin, ItemData As Variant

  On Error Goto metrica:
  Set senso_origin = origin.Worksheets("PSICOSENSOMETRICA") '' PSICOSENSOMETRICA DEL LIBRO ORIGEN ''

  senso_destiny.Select
  ActiveSheet.Range("A4").Select
  Set psicosensometrica_destiny_header = senso_destiny.Range("A2", senso_destiny.Range("A2").End(xlToRight))
  Set psicosensometrica_origin_header = senso_origin.Range("A1", senso_origin.Range("A1").End(xlToRight))
  Set psicosensometrica_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set psicosensometrica_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (senso_origin.Range("A2") <> Empty And senso_origin.Range("A3") <> Empty) Then
    Set psicosensometrica_origin_value = senso_origin.Range("A2", senso_origin.Range("A2").End(xlDown))
  ElseIf (senso_origin.Range("A2") <> Empty And senso_origin.Range("A3") = Empty) Then
    Set psicosensometrica_origin_value = senso_origin.Range("A2")
  End If

  '/***
  '   En los diccionarios de "psicosensometrica_destiny_dictionary" y  "psicosensometrica_origin_dictionary"
  '   se almacena los numeros de la columnas.
  '*/

  ' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO
  For Each ItemPsicosensometricaDestiny In psicosensometrica_destiny_header
    On Error Goto psicotecnicaError
    psicosensometrica_destiny_dictionary.Add psicosensometrica_headers(ItemPsicosensometricaDestiny), (ItemPsicosensometricaDestiny.Column - 1)
  Next ItemPsicosensometricaDestiny

  ' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN
  For Each ItemPsicosensometricaOrigin In psicosensometrica_origin_header
    On Error Goto psicotecnicaError
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
      If formImports.ProgressBarGeneral.Width > (formImports.content_ProgressBarGeneral.Width / 2) Then: formImports.porcentageGeneral.ForeColor = RGB(255, 255, 255)
        If formImports.ProgressBarGeneral.Width < (formImports.content_ProgressBarGeneral.Width / 2) Then: formImports.porcentageGeneral.ForeColor = RGB(0, 0, 0)
          If formImports.ProgressBarOneforOne.Width > (formImports.content_ProgressBarOneforOne.Width / 2) Then: formImports.porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
            If formImports.ProgressBarOneforOne.Width < (formImports.content_ProgressBarOneforOne.Width / 2) Then: formImports.porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("NRO IDENFICACION")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "NRO IDENFICACION")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("PACIENTE")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "PACIENTE")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("PRUEBA PSICOSENSOMETRICA")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "PRUEBA PSICOSENSOMETRICA")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO PPAL")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "DIAGNOSTICO PPAL")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO OBS")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "DIAGNOSTICO OBS")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/1")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "DIAGNOSTICO REL/1")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/2")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "DIAGNOSTICO REL/2")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("DIAGNOSTICO REL/3")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "DIAGNOSTICO REL/3")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("CONTROLES MENSUALES")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "CONTROLES MENSUALES")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("CONTROLES BIMENSUAL")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "CONTROLES BIMENSUAL")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("CONTROLES TRIMESTRALES")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "CONTROLES TRIMESTRALES")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("CONTROLES 6 MESES")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "CONTROLES 6 MESES")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("CONTROLES 1 ANO")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "CONTROLES 1 ANO")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("CONTROLES CONFIRMATORIA")) = charters(ItemData.offset(, psicosensometrica_origin_dictionary( "CONTROLES CONFIRMATORIA")))
              ActiveCell.offset(, psicosensometrica_destiny_dictionary("ID_PSICOSENSOMETRICA")) = ActiveCell.offset(-1, psicosensometrica_destiny_dictionary("ID_PSICOSENSOMETRICA")) + 1
              ActiveCell.offset(1, 0).Select
              numbers = numbers + 1
              numbersGeneral = numbersGeneral + 1
              DoEvents
            Next ItemData

            Range("$I3:$N3").Select
            Call greaterThanOne
            Range("$I3:$N3").Select
            Call iqualCero
            Range("$A3", Range("$A3").End(xlDown)).Select
            Call formatter

            Set psicosensometrica_origin_value = Nothing
            Set psicosensometrica_destiny_header = Nothing
            Set psicosensometrica_origin_header = Nothing
            psicosensometrica_destiny_dictionary.RemoveAll
            psicosensometrica_origin_dictionary.RemoveAll

 psicotecnicaError:
            resume next
 metrica:
            Set senso_origin = origin.Worksheets("PSICOMOTRIZ")
            resume next
End Sub