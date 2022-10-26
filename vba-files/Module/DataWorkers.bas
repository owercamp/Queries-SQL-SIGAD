Attribute VB_Name = "DataWorkers"
Option Explicit

Sub Workers()

  Dim workers_dictionary As Scripting.Dictionary
  Dim emo_dictionary As Scripting.Dictionary
  Dim workers_header, emo_header, workers_value As Object
  Dim ItemWorks, ItemEmo, ItemData As Variant
  Dim range_active As Integer

  Set emo_origin = origin.Worksheets("EMO") '' EMO DEL LIBRO ORIGEN ''
  Windows(destiny.Name).Activate
  worker_destiny.Select
  ActiveSheet.Range("A6").Select
  Set workers_header = worker_destiny.Range("A4", worker_destiny.Range("A4").End(xlToRight))
  Set emo_header = emo_origin.Range("A1", emo_origin.Range("A1").End(xlToRight))
  Set workers_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_dictionary = CreateObject("Scripting.Dictionary")

  If (emo_origin.Range("A2") <> Empty And emo_origin.Range("A3") <> Empty) Then
    Set workers_value = emo_origin.Range("A2", emo_origin.Range("A2").End(xlDown))
    formMix.lblMsg.Caption = "Por favor ingrese el numero ID correspondiente a la orden en SIGAD"
    formMix.Caption = "N" & Chr(250) & "mero de Orden"
    formMix.Show
    formMix.txt_cantidad.SetFocus
  ElseIf (emo_origin.Range("A2") <> Empty And emo_origin.Range("A3") = Empty) Then
    Set workers_value = emo_origin.Range("A2")
    formMix.lblMsg.Caption = "Por favor ingrese el numero ID correspondiente a la orden en SIGAD"
    formMix.Caption = "N" & Chr(250) & "mero de Orden"
    formMix.Show
    formMix.txt_cantidad.SetFocus
  End If

  formMix.Caption = "Forms"
  formMix.lblMsg.Caption = "Ingrese la cantidad de ENFASIS"
  formMix.Show
  formMix.txt_cantidad.SetFocus
  formMix.lblMsg.Caption = "Ingrese la cantidad de DIAGNOSTICOS"
  formMix.Show
  formMix.txt_cantidad.SetFocus

  '/***
  '   En los diccionarios de "workers_dictionary" y  "emo_dictionary"
  '   se almacena los numeros de la columnas.
  '*/

  ' CABECERAS DE LA HOJA TRABAJADORES DEL LIBRO DESTINO
  For Each ItemWorks In workers_header
    On Error GoTo workersError
    workers_dictionary.Add header_worker(ItemWorks), (ItemWorks.Column - 1)
  Next ItemWorks

  ' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN
  For Each ItemEmo In emo_header
    On Error GoTo workersError
    emo_dictionary.Add header_worker(ItemEmo), (ItemEmo.Column - 1)
  Next ItemEmo

  oneForOne = 0
  generalAll = 0
  counts = workers_value.Count
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts
  widthGeneral = formImports.content_ProgressBarOneforOne.Width / totalData
  vals = 1 / counts
  valsGeneral = 1 / totalData
  For Each ItemData In workers_value
    oneForOne = oneForOne + widthOneforOne
    generalAll = generalAll + widthGeneral
    formImports.lblGeneral.Caption = "importando " & CStr(numbers) & " de " & CStr(totalData) & "(" & CStr(totalData - numbers) & ") REGISTROS"
      formImports.lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & worker_destiny.Name
      formImports.ProgressBarOneforOne.Width = oneForOne
      formImports.ProgressBarGeneral.Width = generalAll
      porcentaje = porcentaje + vals
      porcentajeGeneral = porcentajeGeneral + valsGeneral
      formImports.porcentageGeneral.Caption = CStr(VBA.Round(porcentajeGeneral * 100, 1)) & "%"
      formImports.porcentageOneoforOne.Caption = CStr(VBA.Round(porcentaje * 100, 1)) & "%"
      If formImports.ProgressBarGeneral.Width > (formImports.content_ProgressBarGeneral.Width / 2) Then: formImports.porcentageGeneral.ForeColor = RGB(255, 255, 255)
        If formImports.ProgressBarGeneral.Width < (formImports.content_ProgressBarGeneral.Width / 2) Then: formImports.porcentageGeneral.ForeColor = RGB(0, 0, 0)
          If formImports.ProgressBarOneforOne.Width > (formImports.content_ProgressBarOneforOne.Width / 2) Then: formImports.porcentageOneoforOne.ForeColor = RGB(255, 255, 255)
            If formImports.ProgressBarOneforOne.Width < (formImports.content_ProgressBarOneforOne.Width / 2) Then: formImports.porcentageOneoforOne.ForeColor = RGB(0, 0, 0)
              formImports.Caption = CStr(nameCompany)
              ActiveCell = "8"
              ActiveCell.Offset(, workers_dictionary("NOMBRE CONTRATO")) = charters(ItemData.Offset(, emo_dictionary("NOMBRE CONTRATO")))
              ActiveCell.Offset(, workers_dictionary("DESTINO")) = charters(ItemData.Offset(, emo_dictionary("DESTINO")))
              ActiveCell.Offset(, workers_dictionary("CIUDAD")) = city(charters(ItemData.Offset(, emo_dictionary("CIUDAD"))))
              ActiveCell.Offset(, workers_dictionary("INGRESO REGISTRO")) = charters(ItemData.Offset(, emo_dictionary("INGRESO REGISTRO")))
              ActiveCell.Offset(, workers_dictionary("TIPO EXAMEN")) = typeExams(charters(ItemData.Offset(, emo_dictionary("TIPO EXAMEN"))))
              ActiveCell.Offset(, workers_dictionary("FECHA INGRESO")) = charters(ItemData.Offset(, emo_dictionary("FECHA INGRESO")))
              ActiveCell.Offset(, workers_dictionary("PACIENTE")) = charters(ItemData.Offset(, emo_dictionary("PACIENTE")))
              ActiveCell.Offset(, workers_dictionary("NRO IDENFICACION")) = charters(ItemData.Offset(, emo_dictionary("NRO IDENFICACION")))
              ActiveCell.Offset(, workers_dictionary("EDAD")) = charters(ItemData.Offset(, emo_dictionary("EDAD")))
              ActiveCell.Offset(, workers_dictionary("ESTRATO")) = charters(ItemData.Offset(, emo_dictionary("ESTRATO")))
              ActiveCell.Offset(, workers_dictionary("GENERO")) = charters(ItemData.Offset(, emo_dictionary("GENERO")))
              ActiveCell.Offset(, workers_dictionary("NRO HIJOS")) = charters(ItemData.Offset(, emo_dictionary("NRO HIJOS")))
              ActiveCell.Offset(, workers_dictionary("RAZA")) = typeSex(charters(ItemData.Offset(, emo_dictionary("RAZA"))))
              ActiveCell.Offset(, workers_dictionary("ESTADO CIVIL")) = typeCivil(charters(ItemData.Offset(, emo_dictionary("ESTADO CIVIL"))))
              ActiveCell.Offset(, workers_dictionary("ESCOLARIDAD")) = school(charters(ItemData.Offset(, emo_dictionary("ESCOLARIDAD"))))
              ActiveCell.Offset(, workers_dictionary("CARGO USUARIO")) = charters(ItemData.Offset(, emo_dictionary("CARGO USUARIO")))
              ActiveCell.Offset(, workers_dictionary("LAB DURACION EN ANOS")) = charters(ItemData.Offset(, emo_dictionary("LAB DURACION EN ANOS")))
              ActiveCell.Offset(, workers_dictionary("FUENTE")) = charters("ARMYWEB")
              ActiveCell.Offset(, workers_dictionary("TIPO ACTIVIDAD")) = charters("1")
              ActiveCell.Offset(, workers_dictionary("idOrdenListaTrabajadores")) = ActiveCell.Offset(-1, workers_dictionary("idOrdenListaTrabajadores")) + 1
              ActiveCell.Offset(, workers_dictionary("idOrden")) = idOrden
              ActiveCell.Offset(1, 0).Select
              numbers = numbers + 1
              numbersGeneral = numbersGeneral + 1
              DoEvents
            Next ItemData

            Range("$F5").Select
            Call dataDuplicate
            Range("$J5").Select
            Call dataDuplicate
            Range("$I5").Select
            Call dataDuplicate
            Range("$T5").Select
            Call dataDuplicate
            Range("$AW5").Select
            Call dataDuplicate
            Range("$J5", Range("$J5").End(xlDown)).Select
            Call formatter

            Set workers_header = Nothing
            Set emo_header = Nothing
            Set emo_origin = Nothing
            workers_dictionary.RemoveAll
            emo_dictionary.RemoveAll
workersError:
            Resume Next
End Sub
