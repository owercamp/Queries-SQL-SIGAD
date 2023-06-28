Attribute VB_Name = "DataWorkers"
Option Explicit

'TODO: Workers - En esta subrutina se importan datos de trabajadores desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - workers_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - workers_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - workers_value: Una variable de objeto para almacenar el rango de los datos de trabajadores.
'* - ItemData: Una variable de objeto para almacenar el rango de los datos de trabajadores.
'* - range_active: Una variable numerica para hacer un seguimiento del numero de elementos de datos importados.
'* ------------------------------------------------------------------------------------------------------------------
Public Sub Workers()
  Dim workers_dictionary As Scripting.Dictionary
  Dim emo_dictionary As Scripting.Dictionary
  Dim workers_header As Object, emo_header As Object, workers_value As Object
  Dim ItemWorks As Variant, ItemEmo As Variant, ItemData As Variant
  Dim range_active As Integer, company_name As String
  Dim currenCell As range, aumentFromRow As LongPtr, aumentFromID As LongPtr
  
  Set emo_origin = origin.Worksheets("EMO") '' EMO DEL LIBRO ORIGEN ''
  Windows(destiny.Name).Activate
  worker_destiny.Select
  ActiveSheet.range("A5").Select
  Set currenCell = ActiveCell
  Set workers_header = worker_destiny.range("A4", worker_destiny.range("A4").End(xlToRight))
  Set emo_header = emo_origin.range("A1", emo_origin.range("A1").End(xlToRight))
  Set workers_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_dictionary = CreateObject("Scripting.Dictionary")
  
  If (emo_origin.range("A2") <> Empty And emo_origin.range("A3") <> Empty) Then
    Set workers_value = emo_origin.range("A2", emo_origin.range("A2").End(xlDown))
    formMix.lblMsg.Caption = "Por favor ingrese el numero ID correspondiente a la orden en SIGAD"
    formMix.Caption = "N" & Chr(250) & "mero de Orden"
    formMix.Show
    formMix.txt_cantidad.SetFocus
  ElseIf (emo_origin.range("A2") <> Empty And emo_origin.range("A3") = Empty) Then
    Set workers_value = emo_origin.range("A2")
    formMix.lblMsg.Caption = "Por favor ingrese el numero ID correspondiente a la orden en SIGAD"
    formMix.Caption = "N" & Chr(250) & "mero de Orden"
    formMix.Show
    formMix.txt_cantidad.SetFocus
  End If
  
  '' CABECERAS DE LA HOJA TRABAJADORES DEL LIBRO DESTINO ''
  For Each ItemWorks In workers_header
    On Error Resume Next
    workers_dictionary.Add header_worker(ItemWorks), (ItemWorks.Column - 1)
    On Error GoTo 0
  Next ItemWorks
  
  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemEmo In emo_header
    On Error Resume Next
    emo_dictionary.Add header_worker(ItemEmo), (ItemEmo.Column - 1)
    On Error GoTo 0
  Next ItemEmo
  
  oneForOne = 0
  generalAll = 0
  aumentFromRow = 0
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$4").value
  counts = workers_value.Count
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  widthGeneral = formImports.content_ProgressBarOneforOne.Width / totalData
  vals = 1 / counts
  valsGeneral = 1 / totalData
  company_name = Application.InputBox(prompt:="ingrese el Nombre del contrato", Title:="Nombre del contrato", Default:="", Type:=2)
  
  With formImports
    For Each ItemData In workers_value
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
      ElseIf .ProgressBarGeneral.Width < (.content_ProgressBarGeneral.Width / 2) Then
          .porcentageGeneral.ForeColor = RGB(0, 0, 0)
      End If
      
      If .ProgressBarOneforOne.Width > (.content_ProgressBarOneforOne.Width / 2) Then
          .porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
      ElseIf .ProgressBarOneforOne.Width < (.content_ProgressBarOneforOne.Width / 2) Then
          .porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
      End If
      
      .Caption = CStr(nameCompany)
      
      If (typeExams(charters(ItemData.Offset(, emo_dictionary("TIPO EXAMEN")))) <> "EGRESO") Then
        currenCell.Offset(aumentFromRow, 0) = "8"
        currenCell.Offset(aumentFromRow, workers_dictionary("NOMBRE CONTRATO")) = charters(company_name)
        currenCell.Offset(aumentFromRow, workers_dictionary("DESTINO")) = charters(ItemData.Offset(, emo_dictionary("DESTINO")))
        currenCell.Offset(aumentFromRow, workers_dictionary("CIUDAD")) = city(charters(ItemData.Offset(, emo_dictionary("CIUDAD"))))
        currenCell.Offset(aumentFromRow, workers_dictionary("INGRESO REGISTRO")) = charters(ItemData.Offset(, emo_dictionary("INGRESO REGISTRO")))
        currenCell.Offset(aumentFromRow, workers_dictionary("TIPO EXAMEN")) = typeExams(charters(ItemData.Offset(, emo_dictionary("TIPO EXAMEN"))))
        currenCell.Offset(aumentFromRow, workers_dictionary("FECHA INGRESO")) = charters(ItemData.Offset(, emo_dictionary("FECHA INGRESO")))
        currenCell.Offset(aumentFromRow, workers_dictionary("PACIENTE")) = charters(ItemData.Offset(, emo_dictionary("PACIENTE")))
        currenCell.Offset(aumentFromRow, workers_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, emo_dictionary("NRO IDENFICACION")))
        currenCell.Offset(aumentFromRow, workers_dictionary("EDAD")) = charters(ItemData.Offset(, emo_dictionary("EDAD")))
        currenCell.Offset(aumentFromRow, workers_dictionary("ESTRATO")) = charters(ItemData.Offset(, emo_dictionary("ESTRATO")))
        currenCell.Offset(aumentFromRow, workers_dictionary("GENERO")) = charters(ItemData.Offset(, emo_dictionary("GENERO")))
        currenCell.Offset(aumentFromRow, workers_dictionary("NRO HIJOS")) = charters(ItemData.Offset(, emo_dictionary("NRO HIJOS")))
        currenCell.Offset(aumentFromRow, workers_dictionary("RAZA")) = typeSex(charters(ItemData.Offset(, emo_dictionary("RAZA"))))
        currenCell.Offset(aumentFromRow, workers_dictionary("ESTADO CIVIL")) = typeCivil(charters(ItemData.Offset(, emo_dictionary("ESTADO CIVIL"))))
        currenCell.Offset(aumentFromRow, workers_dictionary("ESCOLARIDAD")) = school(charters(ItemData.Offset(, emo_dictionary("ESCOLARIDAD"))))
        currenCell.Offset(aumentFromRow, workers_dictionary("CARGO USUARIO")) = charters(ItemData.Offset(, emo_dictionary("CARGO USUARIO")))
        currenCell.Offset(aumentFromRow, workers_dictionary("LAB DURACION EN ANOS")) = charters(ItemData.Offset(, emo_dictionary("LAB DURACION EN ANOS")))
        currenCell.Offset(aumentFromRow, workers_dictionary("FUENTE")) = charters("ARMYWEB")
        currenCell.Offset(aumentFromRow, workers_dictionary("TIPO ACTIVIDAD")) = charters("1")
        If (currenCell.Offset(aumentFromRow, 0).row = 5) Then
          currenCell.Offset(aumentFromRow, workers_dictionary("idOrdenListaTrabajadores")) = Trim(aumentFromID)
        Else
          aumentFromID = aumentFromID + 1
          currenCell.Offset(aumentFromRow, workers_dictionary("idOrdenListaTrabajadores")) = Trim(aumentFromID)
        End If
        currenCell.Offset(aumentFromRow, workers_dictionary("idOrden")) = idOrden
      End If

      numbers = numbers + 1
      numbersGeneral = numbersGeneral + 1
      aumentFromRow = aumentFromRow + 1
      DoEvents
    Next ItemData

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
    range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=range("$H5"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 4), Array(15, 9)), TrailingMinusNumbers:=True
    range("tbl_trabajadores[[#Headers],[FECHA INGRESO]]").Select

    Set workers_header = Nothing
    Set emo_header = Nothing
    Set emo_origin = Nothing
    workers_dictionary.RemoveAll
    emo_dictionary.RemoveAll
  End With
End Sub

