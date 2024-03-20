Attribute VB_Name = "DataWorkersEmo"
'namespace=vba-files\Module\informations
Option Explicit

'TODO: DataEmoWorkers - En esta subrutina se importan datos de audio desde una hoja de origen a una hoja de destino.
'* ------------------------------------------------------------------------------------------------------------------
'* Variables:
'* - emo_destiny_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de destino.
'* - emo_origin_dictionary: Un objeto Scripting.Dictionary para almacenar los numeros de columna de la hoja de origen.
'* - emo_destiny_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de destino.
'* - emo_origin_header: Una variable de objeto para almacenar el rango del encabezado de la hoja de origen.
'* - emo_origin_value: Una variable de objeto para almacenar los valores de la hoja de origen.
'* - ItemEmoDestiny: Una variable de objeto para almacenar los valores de la columna de la hoja de destino.
'* - ItemEmoOrigin: Una variable de objeto para almacenar los valores de la columna de la hoja de origen.
'* - ItemData: Una variable de objeto para almacenar los valores de la hoja de origen.
'* ------------------------------------------------------------------------------------------------------------------
Dim aumentFromID As LongPtr
Public Sub DataEmoWorkers(ByVal name_sheet As String)
  Dim emo_destiny_dictionary As Scripting.Dictionary
  Dim emo_origin_dictionary As Scripting.Dictionary
  Dim emo_destiny_header As Object, emo_origin_header As Object, emo_origin_value As Object
  Dim ItemEmoDestiny As Object, ItemEmoOrigin As Object, ItemData As Object, emo_origin As Object, cell_active as Range

  Set emo_origin = origin.Worksheets(name_sheet) '' EMO DEL LIBRO ORIGEN ''
  emo_destiny.Select
  emo_destiny.Range("$A5").Select
  Set cell_active = ActiveCell
  Set emo_destiny_header = emo_destiny.Range("$A4", emo_destiny.Range("$A4").End(xlToRight))
  Set emo_origin_header = emo_origin.Range("$A1", emo_origin.Range("$A1").End(xlToRight))
  Set emo_destiny_dictionary = CreateObject("Scripting.Dictionary")
  Set emo_origin_dictionary = CreateObject("Scripting.Dictionary")

  If (emo_origin.Range("$A2") <> Empty And emo_origin.Range("$A3") <> Empty) Then
    Set emo_origin_value = emo_origin.Range("$A2", emo_origin.Range("$A2").End(xlDown))
  ElseIf (emo_origin.Range("$A2") <> Empty And emo_origin.Range("$A3") = Empty) Then
    Set emo_origin_value = emo_origin.Range("$A2")
  End If

  '' CABECERAS DE LA HOJA EMO DEL LIBRO DESTINO ''
  Dim value_data As String
  For Each ItemEmoDestiny In emo_destiny_header
    value_data = emo_headers(ItemEmoDestiny)
    If emo_destiny_dictionary.Exists(value_data) = False And value_data <> Empty Then
      emo_destiny_dictionary.Add value_data, (ItemEmoDestiny.Column - 1)
    End If
  Next ItemEmoDestiny

  '' CABECERA DE LA HOJA EMO DEL LIBRO ORIGEN ''
  For Each ItemEmoOrigin In emo_origin_header
    value_data = emo_headers(ItemEmoOrigin)
    If emo_origin_dictionary.Exists(value_data) = False And value_data <> Empty Then
      emo_origin_dictionary.Add value_data, (ItemEmoOrigin.Column - 1)
    End If
  Next ItemEmoOrigin

  numbers = 1
  oneForOne = 0
  porcentaje = 0
  
  aumentFromID = destiny.Worksheets("RUTAS").range("$F$5").value
  counts = emo_origin_value.Count
  formImports.ProgressBarOneforOne.Width = 0
  formImports.porcentageOneoforOne = "0%"
  vals = 1 / counts
  widthOneforOne = formImports.content_ProgressBarOneforOne.Width / counts

  Dim type_exam As String
  With formImports
    For Each ItemData In emo_origin_value
      oneForOne = oneForOne + widthOneforOne
      generalAll = generalAll + widthGeneral
      .lblGeneral.Caption = "importando " & CStr(numbersGeneral) & " de " & CStr(totalData) & "(" & CStr(totalData - numbersGeneral) & ") REGISTROS"
      .lblDescription.Caption = "importando " & CStr(numbers) & " de " & CStr(counts) & "(" & CStr(counts - numbers) & ") " & emo_destiny.Name
      porcentaje = porcentaje + vals
      porcentajeGeneral = porcentajeGeneral + valsGeneral
      .ProgressBarOneforOne.Width = oneForOne
      .ProgressBarGeneral.Width = generalAll
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

      type_exam = typeExams(Trim(ItemData.Offset(, emo_origin_dictionary("TIPO EXAMEN"))))
      If (type_exam <> "EGRESO") Then
        cell_active.Offset(, emo_destiny_dictionary("NRO IDENFICACION")) = Trim(ItemData.Offset(, emo_origin_dictionary("NRO IDENFICACION")))

        search = ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RUIDO"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / RUIDO")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / RUIDO")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / RUIDO")) = Trim$(ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RUIDO")))
        End If

        search = ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / ILUMINACION"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / ILUMINACION")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / ILUMINACION")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / ILUMINACION")) = Trim$(search)
        End If

        search = ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / VIBRACION"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / VIBRACION")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / VIBRACION")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / VIBRACION")) = Trim$(search)
        End If

        search = ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / TEMP EXTREMAS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / TEMP EXTREMAS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / TEMP EXTREMAS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / TEMP EXTREMAS")) = Trim$(search)
        End If

        search = ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / PRES ATMOSFERICA"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / PRES ATMOSFERICA")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / PRES ATMOSFERICA")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / PRES ATMOSFERICA")) = Trim$(search)
        End If

        search = ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RAD IONIZANTES"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / RAD IONIZANTES")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / RAD IONIZANTES")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / RAD IONIZANTES")) = Trim$(search)
        End If

        search = ItemData.Offset(, emo_origin_dictionary("RIESGO FISICO / RAD NO IONIZANTES"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / RAD NO IONIZANTES")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / RAD NO IONIZANTES")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO FISICO / RAD NO IONIZANTES")) = Trim$(search)
        End If
        
        cell_active.Offset(, emo_destiny_dictionary("RIESGO DE OTROS FACTORES FISICOS")) = VBA.Trim$(ItemData.Offset(, emo_origin_dictionary("RIESGO DE OTROS FACTORES FISICOS")))

        search = ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / VIRUS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / VIRUS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / VIRUS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / VIRUS")) = Trim$(search)
        End If

        search = ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / BACTERIAS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / BACTERIAS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / BACTERIAS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / BACTERIAS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / HONGOS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / HONGOS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / HONGOS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / HONGOS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / RICKETSIAS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / RICKETSIAS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / RICKETSIAS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / RICKETSIAS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / PARASITOS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / PARASITOS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / PARASITOS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / PARASITOS")) = Trim$(search)
        End If

        search = ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / FLUIDOS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / FLUIDOS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / FLUIDOS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / FLUIDOS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / PICADURAS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / PICADURAS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / PICADURAS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / PICADURAS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO BIOLOGICO / MORDEDURAS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / MORDEDURAS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / MORDEDURAS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO BIOLOGICO / MORDEDURAS")) = Trim$(search)
        End If

        cell_active.Offset(, emo_destiny_dictionary("OTROS RIESGOS BIOLOGICOS")) = VBA.Trim$(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS BIOLOGICOS")))

        search = ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / POLVOS"))
        If (withIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / POLVOS")) = 1
        ElseIf (withoutIncidence.Exists(Trim(search))) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / POLVOS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / POLVOS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / FIBRAS"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / FIBRAS")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / FIBRAS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / FIBRAS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / LIQUIDOS"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / LIQUIDOS")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / LIQUIDOS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / LIQUIDOS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO /GASES"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO /GASES")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO /GASES")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO /GASES")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / VAPORES"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / VAPORES")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / VAPORES")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / VAPORES")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO / HUMOS"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / HUMOS")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / HUMOS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO / HUMOS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO QUIMICO /MATERIAL PARTICULADO")) = Trim$(search)
        End If

        cell_active.Offset(, emo_destiny_dictionary("OTROS RIESGOS QUIMICOS")) = VBA.Trim$(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS QUIMICOS")))
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / GESTION ORGANIZACIONAL")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CARACT DEL GRUPO"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / CARACT DEL GRUPO")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / CARACT DEL GRUPO")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / CARACT DEL GRUPO")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / INTERFACES TAREA"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / INTERFACES TAREA")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / INTERFACES TAREA")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / INTERFACES TAREA")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CARACT ORGANIZACION"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / CARACT ORGANIZACION")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / CARACT ORGANIZACION")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / CARACT ORGANIZACION")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / CONDICIONES"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / CONDICIONES")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / CONDICIONES")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / CONDICIONES")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO PSICO / JORNADA"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / JORNADA")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / JORNADA")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO PSICO / JORNADA")) = Trim$(search)
        End If

        cell_active.Offset(, emo_destiny_dictionary("OTROS PSICO LABORAL")) = VBA.Trim$(ItemData.Offset(, emo_origin_dictionary("OTROS PSICO LABORAL")))
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_POSTURA"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_POSTURA")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_POSTURA")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_POSTURA")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_ESFUERZO"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_ESFUERZO")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_ESFUERZO")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_ESFUERZO")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_MOVREPETITIVO")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("RIESGO_BIOMECANICO_MANIPULACION_CARGA")) = Trim$(search)
        End If

        cell_active.Offset(, emo_destiny_dictionary("OTROS RIESGOS BIOMECANICOS")) = VBA.Trim$(ItemData.Offset(, emo_origin_dictionary("OTROS RIESGOS BIOMECANICOS")))

        search = ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / MECANICOS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ELECTRICOS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / LOCATIVO")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TECNOLOGICO")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ACC DE TRANSITO")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / PUBLICOS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / TRABAJO EN ALTURAS")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / ESPACIOS CONFINADOS")) = Trim$(search)
        End If

        cell_active.Offset(, emo_destiny_dictionary("CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD")) = VBA.Trim$(ItemData.Offset(, emo_origin_dictionary("CONDICIONES DE SEGURIDAD / OTROS DE SEGURIDAD")))

        search = ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / SISMO"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / SISMO")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / SISMO")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / SISMO")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / TERREMOTO"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / TERREMOTO")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / TERREMOTO")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / TERREMOTO")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / VENDAVAL"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / VENDAVAL")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / VENDAVAL")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / VENDAVAL")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / INUNDACION"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / INUNDACION")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / INUNDACION")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / INUNDACION")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / DERRUMBE"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / DERRUMBE")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / DERRUMBE")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / DERRUMBE")) = Trim$(search)
        End If
        
        search = ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / PRECIPITACIONES"))
        If (withIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / PRECIPITACIONES")) = 1
        ElseIf (withoutIncidence.Exists(search)) Then
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / PRECIPITACIONES")) = 0
        Else
          cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / PRECIPITACIONES")) = Trim$(search)
        End If

        cell_active.Offset(, emo_destiny_dictionary("FENOMENOS NATURALES / OTROS NATURALES")) = VBA.Trim$(ItemData.Offset(, emo_origin_dictionary("FENOMENOS NATURALES / OTROS NATURALES")))
        cell_active.Offset(, emo_destiny_dictionary("FECHA ACCIDENTE")) = Trim(ItemData.Offset(, emo_origin_dictionary("FECHA ACCIDENTE")))
        cell_active.Offset(, emo_destiny_dictionary("ACCIDENTE_PASO_EN_EMPRESA")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("ACCIDENTE_PASO_EN_EMPRESA"))))
        cell_active.Offset(, emo_destiny_dictionary("TIPO ACCIDENTE")) = Trim(ItemData.Offset(, emo_origin_dictionary("TIPO ACCIDENTE")))
        cell_active.Offset(, emo_destiny_dictionary("NATURALEZA LESION")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("NATURALEZA LESION"))))
        cell_active.Offset(, emo_destiny_dictionary("PARTE AFECTADA")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("PARTE AFECTADA"))))
        cell_active.Offset(, emo_destiny_dictionary("INCAPACIDAD")) = Trim(ItemData.Offset(, emo_origin_dictionary("INCAPACIDAD")))
        cell_active.Offset(, emo_destiny_dictionary("SECUELAS")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("SECUELAS"))))
        cell_active.Offset(, emo_destiny_dictionary("NOMBRE ENFERMEDAD")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("NOMBRE ENFERMEDAD"))))
        cell_active.Offset(, emo_destiny_dictionary("ETAPA")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("ETAPA"))))
        cell_active.Offset(, emo_destiny_dictionary("OBSERVACIONES DE ENFERMEDAD")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES DE ENFERMEDAD"))))
        cell_active.Offset(, emo_destiny_dictionary("ACT_ FISICA")) = typeActivity(Trim(UCase(ItemData.Offset(, emo_origin_dictionary("ACT_ FISICA")))))
        cell_active.Offset(, emo_destiny_dictionary("FUMA")) = typeSmoke(Trim(ItemData.Offset(, emo_origin_dictionary("FUMA"))))
        cell_active.Offset(, emo_destiny_dictionary("CONSUMO DE ALCOHOL")) = Trim(ItemData.Offset(, emo_origin_dictionary("CONSUMO DE ALCOHOL")))
        cell_active.Offset(, emo_destiny_dictionary("PESO")) = Trim(ItemData.Offset(, emo_origin_dictionary("PESO")))
        cell_active.Offset(, emo_destiny_dictionary("TALLA")) = Trim(ItemData.Offset(, emo_origin_dictionary("TALLA")))
        cell_active.Offset(, emo_destiny_dictionary("TENSION ARTERIAL")) = Trim(ItemData.Offset(, emo_origin_dictionary("TENSION ARTERIAL")))
        cell_active.Offset(, emo_destiny_dictionary("FREC_ CARDIACA")) = Trim(ItemData.Offset(, emo_origin_dictionary("FREC_ CARDIACA")))
        cell_active.Offset(, emo_destiny_dictionary("FREC_ RESPIRATORIA")) = Trim(ItemData.Offset(, emo_origin_dictionary("FREC_ RESPIRATORIA")))
        cell_active.Offset(, emo_destiny_dictionary("PERIMETRO ABDOMINAL")) = Trim(ItemData.Offset(, emo_origin_dictionary("PERIMETRO ABDOMINAL")))
        cell_active.Offset(, emo_destiny_dictionary("LATERALIDAD")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("LATERALIDAD"))))
        cell_active.Offset(, emo_destiny_dictionary("OBS DIAGS")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("OBS DIAGS"))))
        cell_active.Offset(, emo_destiny_dictionary("CONCEPTO DE EVALUACION")) = validateConcepts(Trim(UCase(ItemData.Offset(, emo_origin_dictionary("CONCEPTO DE EVALUACION")))))
        cell_active.Offset(, emo_destiny_dictionary("OBSERVACIONES DEL CONCEPTO")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("OBSERVACIONES DEL CONCEPTO"))))
        cell_active.Offset(, emo_destiny_dictionary("RECOMENDACIONES ESPECIFICAS")) = Trim(UCase(ItemData.Offset(, emo_origin_dictionary("RECOMENDACIONES ESPECIFICAS"))))
        cell_active.Offset(, emo_destiny_dictionary("REMISION EPS")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("CONTROL PERIODICO OCUPACIONAL")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("UTILIZACION EPP ACORDE AL CARGO")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("REALIZACION DE PRUEBAS COMPLEMENTARIAS")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("HABITOS NUTRICIONALES")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("EJERCICIO REGULAR 3 VECES POR SEMANA")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("DEJAR DE FUMAR")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("REDUCIR CONSUMO ALCOHOL")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("OBSERVACIONES")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("OSTEOMUSCULAR")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("VISUAL")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("ALTURAS")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("BIOLOGICO")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("MANIPULACION DE ALIMENTOS")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("QUIMICO")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("CUIDADO DE LA VOZ")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("TEMPERATURAS EXTREMAS")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("ESPACIOS CONFINADOS")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("PIEL")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("RESPIRATORIA")) = "0"
        cell_active.Offset(, emo_destiny_dictionary("AUDITIVO")) = "0"
        If (cell_active.Row <> 5) Then
          aumentFromID = aumentFromID + 1
        End If
        cell_active.Offset(, emo_destiny_dictionary("ID_EMO")) = aumentFromID
        Set cell_active = cell_active.Offset(1, 0)
        numbers = numbers + 1
        numbersGeneral = numbersGeneral + 1
        DoEvents
      End If
    Next ItemData
  End With

  Call thisText(emo_destiny.Range("tbl_emo[[#Data],[INCAPACIDAD]]"))
  Call dataDuplicate(emo_destiny.Range("tbl_emo[[#Data],[orden_lista_trabajadoresid]]"))
  Call dataDuplicate(emo_destiny.Range("tbl_emo[[#Data],[id_emo]]"))
  Call dataDuplicate(emo_destiny.Range("tbl_emo[[#Data],[NRO IDENTIFICACION]]"))
  Call Risk(emo_destiny.Range("tbl_emo[[#Data],[SCRIPT ics_emo_riesgos]]"))
  Call riskPre_ingreso(emo_destiny.Range("tbl_emo[[#Data],[SCRIPT ics_emo_riesgos]]"))
  Call formatter(emo_destiny.Range("tbl_emo[[#Data],[NRO IDENTIFICACION]]"))

  Set emo_origin_value = Nothing
  Set emo_destiny_header = Nothing
  Set emo_origin_header = Nothing
  emo_destiny_dictionary.RemoveAll
  emo_origin_dictionary.RemoveAll

End Sub