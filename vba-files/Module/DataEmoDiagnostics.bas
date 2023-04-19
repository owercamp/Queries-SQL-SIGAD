Attribute VB_Name = "DataEmoDiagnostics"
Option Explicit

Sub DataDiagnosticsEmo()

  Dim diagnostics_destiny_dictionary As Scripting.Dictionary
  Dim emo_origin_dictionary As Scripting.Dictionary
  Dim diagnostics_destiny_header As Object, emo_origin_header As Object, emo_origin_value As Object
  Dim ItemDiagnosticsDestiny As Variant, ItemEmoOrigin As Variant, ItemData As Variant

  Set emo_origin = origin.Worksheets("EMO") '' EMO DEL LIBRO ORIGEN ''
  diagnostics_destiny.Select
  ActiveSheet.Range("A5").Select

  Set diagnostics_destiny_header = diagnostics_destiny.Range("A4", diagnostics_destiny.Range("A4").End(xlToRight))
  Set emo_origin_header = emo_origin.Range("A1", emo_origin.Range("A1").End(xlToRight))
  Set diagnostics_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (emo_origin.Range("A2") <> Empty And emo_origin.Range("A3") <> Empty) Then
    Set emo_origin_value = emo_origin.Range("A2", emo_origin.Range("A2").End(xlDown))
  ElseIf (emo_origin.Range("A2") <> Empty And emo_origin.Range("A3") = Empty) Then
    Set emo_origin_value = emo_origin.Range("A2")
  End If

  ''   En los diccionarios de "diagnostics_destiny_dictionary" y  "emo_origin_dictionary" ''
  ''   se almacena los numeros de la columnas. ''

  x = 1
  For Each ItemDiagnosticsDestiny In diagnostics_destiny_header
    On Error GoTo diagnosticsError
    diagnostics_destiny_dictionary.Add diagnostics_header(ItemDiagnosticsDestiny), (ItemDiagnosticsDestiny.Column - 1)
  Next ItemDiagnosticsDestiny

  x = 1
  For Each ItemEmoOrigin In emo_origin_header
    On Error GoTo diagnosticsError
    emo_origin_dictionary.Add diagnostics_header(ItemEmoOrigin), (ItemEmoOrigin.Column - 1)
  Next ItemEmoOrigin

  numbers = 1
  porcentaje = 0
  counts = emo_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  oneForOne = 0
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  For Each ItemData In emo_origin_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & diagnostics_destiny.Name
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
      If (typeExams(charters(ItemData.Offset(, emo_origin_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        ActiveCell.Offset(, diagnostics_destiny_dictionary("IDENTIFICACION")) = charters(ItemData.Offset(, emo_origin_dictionary("IDENTIFICACION")))
        ActiveCell.Offset(, diagnostics_destiny_dictionary("CODIGO DIAG PPAL")) = charters(ItemData.Offset(, emo_origin_dictionary("CODIGO DIAG PPAL")))
        ActiveCell.Offset(, diagnostics_destiny_dictionary("DIAG PPAL")) = charters(ItemData.Offset(, emo_origin_dictionary("DIAG PPAL")))
        For i = 1 To ((emo_origin_dictionary.Count - 5) / 2)
          ActiveCell.Offset(, diagnostics_destiny_dictionary("CODIGO DIAG REL" & i)) = charters(ItemData.Offset(, emo_origin_dictionary("CODIGO DIAG REL" & i)))
          ActiveCell.Offset(, diagnostics_destiny_dictionary("DIAG REL " & i)) = charters(ItemData.Offset(, emo_origin_dictionary("DIAG REL " & i)))
        Next i
        ActiveCell.Offset(1, 0).Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      DoEvents
    Next ItemData

    Range("$A5").Select
    Call dataDuplicate
    Range("$A5", Range("$A5").End(xlDown)).Select
    Call formatter

    Set diagnostics_destiny_header = Nothing
    Set emo_origin_header = Nothing
    Set emo_origin_value = Nothing
    diagnostics_destiny_dictionary.RemoveAll
    emo_origin_dictionary.RemoveAll
 diagnosticsError:
    Resume Next
End Sub
