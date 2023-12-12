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
Dim aumentFromID As LongPtr
Public Sub Workers(ByVal name_sheet As String)
  Dim workers_dictionary As Scripting.Dictionary
  Dim emo_dictionary As Scripting.Dictionary
  Dim workers_header As Object, emo_header As Object, workers_value As Object
  Dim ItemWorks As Object, ItemEmo As Object, ItemData As Object
  Dim company_name As String, emo_origin As Object

  Set emo_origin = origin.Worksheets(name_sheet) '' EMO DEL LIBRO ORIGEN ''
  Windows(destiny.Name).Activate
  worker_destiny.Select
  worker_destiny.Range("$A5").Select
  Set workers_header = worker_destiny.Range("$A4", worker_destiny.Range("$A4").End(xlToRight))
  Set emo_header = emo_origin.Range("$A1", emo_origin.Range("$A1").End(xlToRight))
  Set workers_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_dictionary = CreateObject("Scripting.Dictionary")

  formMix.lblMsg.Caption = "Por favor ingrese el numero ID correspondiente a la orden en SIGAD"
  formMix.Caption = "N" & ChrW(250) & "mero de Orden"
  formMix.Show
  formMix.txt_cantidad.SetFocus

  If (emo_origin.Range("$A2") <> Empty And emo_origin.Range("$A3") <> Empty) Then
    Set workers_value = emo_origin.Range("$A2", emo_origin.Range("$A2").End(xlDown))
  ElseIf (emo_origin.Range("$A2") <> Empty And emo_origin.Range("$A3") = Empty) Then
    Set workers_value = emo_origin.Range("$A2")
  End If

  '' CABECERAS DE LA HOJA TRABAJADORES DEL LIBRO DESTINO ''
  Dim value_data As String
  For Each ItemWorks In workers_header
    value_data = header_worker(ItemWorks)
    If workers_dictionary.Exists(value_data) = False And value_data <> Empty Then
      workers_dictionary.Add value_data, (ItemWorks.Column - 1)
    End If
  Next ItemWorks

  '' CABECERAS DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemEmo In emo_header
    value_data = header_worker(ItemEmo)
    If emo_dictionary.Exists(value_data) = False And value_data <> Empty Then
      emo_dictionary.Add value_data, (ItemEmo.Column - 1)
    End If
  Next ItemEmo

  oneForOne = 0
  generalAll = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$4").value
  counts = workers_value.Count
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  widthGeneral = formImports.content_ProgressBarOneforOne.Width / totalData
  vals = 1 / counts
  valsGeneral = 1 / totalData
  company_name = Application.InputBox(prompt:="ingrese el Nombre del contrato", Title:="Nombre del contrato", Default:="", Type:=2)

  Dim type_exam As String
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
      Elseif .ProgressBarGeneral.Width < (.content_ProgressBarGeneral.Width / 2) Then
        .porcentageGeneral.ForeColor = RGB(0, 0, 0)
      End If

      If .ProgressBarOneforOne.Width > (.content_ProgressBarOneforOne.Width / 2) Then
        .porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
      Elseif .ProgressBarOneforOne.Width < (.content_ProgressBarOneforOne.Width / 2) Then
        .porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
      End If

      .Caption = CStr(nameCompany)

      type_exam = typeExams(Trim(ItemData.Offset(, emo_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        ActiveCell = "8"
        ActiveCell.Offset(, workers_dictionary("NOMBRE CONTRATO")) = Trim(company_name)
        ActiveCell.Offset(, workers_dictionary("DESTINO")) = Trim(ItemData.Offset(, emo_dictionary("DESTINO")))
        ActiveCell.Offset(, workers_dictionary("CIUDAD")) = city(Trim(UCase(ItemData.Offset(, emo_dictionary("CIUDAD")))))
        ActiveCell.Offset(, workers_dictionary("INGRESO REGISTRO")) = Trim(ItemData.Offset(, emo_dictionary("INGRESO REGISTRO")))
        ActiveCell.Offset(, workers_dictionary("TIPO EXAMEN")) = UCase(type_exam)
        ActiveCell.Offset(, workers_dictionary("FECHA INGRESO")) = Trim(ItemData.Offset(, emo_dictionary("FECHA INGRESO")))
        ActiveCell.Offset(, workers_dictionary("PACIENTE")) = Trim(UCase(ItemData.Offset(, emo_dictionary("PACIENTE"))))
        ActiveCell.Offset(, workers_dictionary("NRO IDENFICACION")) = Trim(ItemData.Offset(, emo_dictionary("NRO IDENFICACION")))
        ActiveCell.Offset(, workers_dictionary("EDAD")) = Trim(ItemData.Offset(, emo_dictionary("EDAD")))
        ActiveCell.Offset(, workers_dictionary("ESTRATO")) = Trim(UCase(ItemData.Offset(, emo_dictionary("ESTRATO"))))
        ActiveCell.Offset(, workers_dictionary("GENERO")) = Trim(UCase(ItemData.Offset(, emo_dictionary("GENERO"))))
        ActiveCell.Offset(, workers_dictionary("NRO HIJOS")) = Trim(ItemData.Offset(, emo_dictionary("NRO HIJOS")))
        ActiveCell.Offset(, workers_dictionary("RAZA")) = typeSex(Trim(UCase(ItemData.Offset(, emo_dictionary("RAZA")))))
        ActiveCell.Offset(, workers_dictionary("ESTADO CIVIL")) = typeCivil(Trim(UCase(ItemData.Offset(, emo_dictionary("ESTADO CIVIL")))))
        ActiveCell.Offset(, workers_dictionary("ESCOLARIDAD")) = school(Trim(UCase(ItemData.Offset(, emo_dictionary("ESCOLARIDAD")))))
        ActiveCell.Offset(, workers_dictionary("CARGO USUARIO")) = Trim(UCase(ItemData.Offset(, emo_dictionary("CARGO USUARIO"))))
        ActiveCell.Offset(, workers_dictionary("LAB DURACION EN ANOS")) = IIf(Trim(ItemData.Offset(, emo_dictionary("LAB DURACION EN ANOS"))) = "SIN DATO", "", Trim(ItemData.Offset(, emo_dictionary("LAB DURACION EN ANOS"))))
        ActiveCell.Offset(, workers_dictionary("FUENTE")) = "ARMYWEB"
        ActiveCell.Offset(, workers_dictionary("TIPO ACTIVIDAD")) = "1"
        If (ActiveCell.Row <> 5) Then
          aumentFromID = aumentFromID + 1
        End If
        ActiveCell.Offset(, workers_dictionary("idOrdenListaTrabajadores")) = aumentFromID
        ActiveCell.Offset(, workers_dictionary("idOrden")) = idOrden
        ActiveCell.Offset(1, 0).Select
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents
      End If
    Next ItemData
  End With

  Call dataDuplicate(worker_destiny.Range("tbl_trabajadores[[#Data],[INGRESO]]"))
  Call dataDuplicate(worker_destiny.Range("tbl_trabajadores[[#Data],[NRO IDENFICACION]]"))
  Call dataDuplicate(worker_destiny.Range("tbl_trabajadores[[#Data],[PACIENTE]]"))
  Call dataDuplicate(worker_destiny.Range("tbl_trabajadores[[#Data],[CARGO USUARIO]]"))
  Call dataDuplicate(worker_destiny.Range("tbl_trabajadores[[#Data],[idOrdenListaTrabajadores]]"))
  Call formatter(worker_destiny.Range("tbl_trabajadores[[#Data],[NRO IDENFICACION]]"))

  Set workers_header = Nothing
  Set emo_header = Nothing
  Set emo_origin = Nothing
  workers_dictionary.RemoveAll
  emo_dictionary.RemoveAll
End Sub