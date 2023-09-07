Attribute VB_Name = "DataWorkers"
'namespace=vba-files\Module\informations
Option Explicit

'TODO: Workers - En esta subrutina se importan datos de trabajadores desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - workers_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - workers_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - range_active: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* ------------------------------------------------------------------------------------------------------------------
Dim emo_dictionary As Scripting.Dictionary
Dim aumentFromID As LongPtr
Public Sub Workers()
  Dim tbl_workers As Object, company_name As String, xNumber As Long, emo_origin As Variant

  emo_origin = origin.Worksheets("EMO").Range("A1").CurrentRegion.value '' EMO DEL LIBRO ORIGEN ''
  Windows(destiny.Name).Activate
  worker_destiny.Select
  Set tbl_workers = ActiveSheet.ListObjects("tbl_trabajadores")
  Set emo_dictionary = CreateObject("Scripting.Dictionary")

  formMix.lblMsg.Caption = "Por favor ingrese el numero ID correspondiente a la orden en SIGAD"
  formMix.Caption = "N" & Chr(250) & "mero de Orden"
  formMix.Show
  formMix.txt_cantidad.SetFocus

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For xNumber = 1 To Ubound(emo_origin, 2)
    On Error Resume Next
    emo_dictionary.Add header_worker(emo_origin(1, xNumber)), xNumber
    On Error GoTo 0
  Next xNumber

  oneForOne = 0
  generalAll = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$4").value
  counts = Ubound(emo_origin, 1) - 1
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  widthGeneral = formImports.content_ProgressBarOneforOne.Width / totalData
  vals = 1 / counts
  valsGeneral = 1 / totalData
  company_name = Application.InputBox(prompt:="ingrese el Nombre del contrato", Title:="Nombre del contrato", Default:="", Type:=2)

  With formImports
    For xNumber = 2 To Ubound(emo_origin, 1)
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbers) & " de " & CStr(totalData) & "(" & CStr(totalData - numbers) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & worker_destiny.Name
      .ProgressBarOneforOne.Width = oneForOne
      .ProgressBarGeneral.Width = generalAll
      porcentaje = porcentaje + vals
      porcentajeGeneral = porcentajeGeneral + valsGeneral
      .porcentageGeneral.Caption = CStr(VBA.Round(porcentajeGeneral * 100, 1)) & "%"
      .porcentageOneoforOne.Caption = CStr(VBA.Round(porcentaje * 100, 1)) & "%"

      If .ProgressBarGeneral.Width > (.content_ProgressBarGeneral.Width / 2) Then
        .porcentageGeneral.ForeColor = RGB(255, 255, 255)
      Elseif .ProgressBarGeneral.Width < (.content_ProgressBarGeneral.Width / 2) Then
        .porcentageGeneral.ForeColor = RGB(0, 0, 0)
      End If

      If .ProgressBarOneforOne.Width > (.content_ProgressBarOneforOne.Width / 2) Then
        .porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
      Elseif .ProgressBarOneforOne.Width < (.content_ProgressBarOneforOne.Width / 2) Then
        .porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
      End If

      .Caption = CStr(nameCompany)

      If (typeExams(charters(emo_origin(xNumber, emo_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        Select Case numbers
          Case 1
            Call addNewRegister(tbl_workers.ListRows(1), aumentFromID, emo_origin, company_name, xNumber)
            DoEvents
          Case Else
            aumentFromID = aumentFromID + 1
            Call addNewRegister(tbl_workers.ListRows.Add, aumentFromID, emo_origin, company_name, xNumber)
            DoEvents
        End Select
      End If
      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
    Next xNumber
  End With

  range("$F5").Select
  Call dataDuplicate
  range("$J5").Select
  Call dataDuplicate
  range("$I5").Select
  Call dataDuplicate
  range("$T5").Select
  Call dataDuplicate
  range("$AW5").Select
  Call dataDuplicate
  range("$J5", range("$J5").End(xlDown)).Select
  Call formatter
  range("$H5").Select
  range(selection, selection.End(xlDown)).Select
  selection.TextToColumns Destination:=range("$H5"), DataType:=xlFixedWidth, _
  FieldInfo:=Array(Array(0, 4), Array(15, 9)), TrailingMinusNumbers:=True
  range("tbl_trabajadores[[#Headers],[FECHA INGRESO]]").Select

  Set emo_origin = Nothing
  emo_dictionary.RemoveAll
End Sub

Private Sub addNewRegister(ByVal table As Object, ByVal autoIncrement As LongPtr, ByVal information As Variant, ByVal company_name As String, ByVal x As Long)

  With table
    .Range(1) = "8"
    .Range(2) = charters(company_name)
    .Range(4) = charters(information(x, emo_dictionary("DESTINO")))
    .Range(5) = city(charters(information(x, emo_dictionary("CIUDAD"))))
    .Range(6) = charters(information(x, emo_dictionary("INGRESO REGISTRO")))
    .Range(7) = typeExams(charters(information(x, emo_dictionary("TIPO EXAMEN"))))
    .Range(8) = charters(information(x, emo_dictionary("FECHA INGRESO")))
    .Range(9) = charters(information(x, emo_dictionary("PACIENTE")))
    .Range(10) = charters(information(x, emo_dictionary("NRO IDENFICACION")))
    .Range(11) = charters(information(x, emo_dictionary("EDAD")))
    .Range(13) = charters(information(x, emo_dictionary("ESTRATO")))
    .Range(14) = charters(information(x, emo_dictionary("GENERO")))
    .Range(15) = charters(information(x, emo_dictionary("NRO HIJOS")))
    .Range(17) = typeSex(charters(information(x, emo_dictionary("RAZA"))))
    .Range(18) = typeCivil(charters(information(x, emo_dictionary("ESTADO CIVIL"))))
    .Range(19) = school(charters(information(x, emo_dictionary("ESCOLARIDAD"))))
    .Range(20) = charters(information(x, emo_dictionary("CARGO USUARIO")))
    .Range(22) = IIf(charters(information(x, emo_dictionary("LAB DURACION EN ANOS"))) = "SIN DATO", "", charters(information(x, emo_dictionary("LAB DURACION EN ANOS"))))
    .Range(24) = "ARMYWEB"
    .Range(25) = 1
    .Range(49) = autoIncrement            
    .Range(50) = idOrden
  End With

End Sub