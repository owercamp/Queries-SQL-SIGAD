Attribute VB_Name = "FunctionBook"
Option Explicit
Public sigad As Variant
Public trabajadores, emo, audio, visio, opto, espiro, osteo, complementarios, psicotecnica, psicosensometrica, enfasis, diag As Worksheet

Sub cargos()
  Attribute cargos.VB_ProcData.VB_Invoke_Func = "k\n14"
  Workbooks.Open (ThisWorkbook.Worksheets("RUTAS").Range("C7").value)
End Sub

Sub folder(route, folderName, workbookActive, YearNow, MonthNow)
  Dim splitRoute As String
  splitRoute = Application.PathSeparator

  If Dir(route, vbDirectory) = Empty Then: MkDir route
    If Dir(route & splitRoute & YearNow, vbDirectory) = Empty Then: MkDir (route & splitRoute & YearNow)
      If Dir(route & splitRoute & YearNow & splitRoute & MonthNow, vbDirectory) = Empty Then: MkDir (route & splitRoute & YearNow & splitRoute & MonthNow)

        If Dir(route & splitRoute & YearNow & splitRoute & MonthNow & splitRoute & folderName, vbDirectory) = Empty Then
          MkDir (route & splitRoute & YearNow & splitRoute & MonthNow & splitRoute & folderName)
          Application.ActiveWorkbook.SaveCopyAs Filename:=route & splitRoute & YearNow & splitRoute & MonthNow & splitRoute & folderName & splitRoute & workbookActive
          Application.StatusBar = "se guardo una copia en: " & route & splitRoute & YearNow & splitRoute & MonthNow & splitRoute & folderName & splitRoute & workbookActive
        Else
          Application.ActiveWorkbook.SaveCopyAs Filename:=route & splitRoute & YearNow & splitRoute & MonthNow & splitRoute & folderName & splitRoute & workbookActive
          Application.StatusBar = "se guardo una copia en: " & route & splitRoute & YearNow & splitRoute & MonthNow & splitRoute & folderName & splitRoute & workbookActive
        End If
End Sub

Sub clearContents()

  Dim rng, info, rngTrabajadores, rngEmo, rngAudio, rngVisio, rngOpto, rngEspiro, rngOsteo, rngComplementarios, rngPsicotecnica, rngPsicosensometrica, rngEnfasis, rngDiag, MyDay, MyMonth, MyYear As Integer
  Dim meses, finalRow, RowActive As Variant
  Dim nombre, orden, fecha, company, bookNow As String
  Dim libro, consolidado As Object

  meses = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
  formMix.Caption = "SIGAD Informe"
  formMix.lblMsg.Caption = "Ingrese el n" & Chr(250) & "mero de orden SIGAD"
  formMix.Show

  Set trabajadores = Worksheets("TRABAJADORES")
  Set emo = Worksheets("EMO")
  Set audio = Worksheets("AUDIO")
  Set visio = Worksheets("VISIO")
  Set opto = Worksheets("OPTO")
  Set espiro = Worksheets("ESPIRO")
  Set osteo = Worksheets("OSTEO")
  Set complementarios = Worksheets("COMPLEMENTARIOS")
  Set psicotecnica = Worksheets("PSICOTECNICA")
  Set psicosensometrica = Worksheets("PSICOSENSOMETRICA")
  Set enfasis = Worksheets("ENFASIS")
  Set diag = Worksheets("DIAGNOSTICOS")

  MyDay = Day(Date)
  MyMonth = Month(Date)
  MyYear = Year(Date)
  bookNow = Application.ActiveWorkbook.Name
  If trabajadores.Range("D6") <> Empty Or trabajadores.Range("D6") <> vbNullString Then: nombre = trabajadores.Range("B6").value & " - " & trabajadores.Range("D6").value & ".xlsb"
    If trabajadores.Range("D6") = Empty Or trabajadores.Range("D6") = vbNullString Then: nombre = trabajadores.Range("B6").value & ".xlsb"
      orden = trabajadores.Range("AX6").value

      If (Not IsEmpty(nombre)) And (Not IsEmpty(orden)) And (Not IsEmpty(sigad)) Then
        fecha = CStr(MyDay) + " " + CStr(meses(MyMonth - 1)) + " " + CStr(MyYear)

        route = CStr(Worksheets("RUTAS").Range("C6").value)

        trabajadores.Select
        Call folder(route, fecha, nombre, MyYear, meses(MyMonth - 1))

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False


        '' REGISTRO EN CONSOLIDADO ''
        info = Worksheets("TRABAJADORES").Range("A5", Worksheets("TRABAJADORES").Range("A5").End(xlDown)).Count

        libro = Worksheets("RUTAS").Range("C5").value
        If trabajadores.Range("D5").value = Empty Or trabajadores.Range("D5").value = vbNullString Then
          company = trabajadores.Range("B5").value
        Else
          company = trabajadores.Range("B5").value & " - " & trabajadores.Range("D5").value
        End If

        Set consolidado = Workbooks.Open(libro)
        '' TRABAJADORES ''

        consolidado.Worksheets("Registros").Select
        consolidado.ActiveSheet.Unprotect Password:="1024500065"
        Range("C3").End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        ActiveCell = Trim(UCase(company))
        ActiveCell.Offset(0, 1) = Trim(UCase("ICS-" & PadLeft(sigad, 4, "0")))
        ActiveCell.Offset(0, 2) = Trim(orden)
        ActiveCell.Offset(0, -1) = Date
        ActiveCell.Offset(0, 3) = Trim(info)

        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual

        consolidado.ActiveSheet.Protect Password:="1024500065", DrawingObjects:=False, Contents:=True, Scenarios:= _
        False, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:= _
        True
        consolidado.Save
        consolidado.Close

        Call AddRecordToGoogleSheet(Trim(UCase(company)), Trim(UCase("ICS-" & PadLeft(sigad, 4, "0"))), Trim(orden),Trim(info), libro, bookNow)

        Application.Calculation = xlCalculationManual
        trabajadores.Select
        route = Worksheets("RUTAS").Range("C8").value

        Range("A5").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando trabajadores por favor espere..."
          rng = Range("A5", Range("A5").End(xlDown)).Count - 2
          DoEvents
          Range("A5", Range("A5").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          ThisWorkbook.Worksheets("RUTAS").Range("$F$4") = CLngLng(Trim(Range("$AW$5").value)) + 1
        Else
          ThisWorkbook.Worksheets("RUTAS").Range("$F$4") = CLngLng(Trim(Range("$AW$5").value)) + 1
        End If

        enfasis.Select
        Range("A5").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando enfasis por favor espere..."
          rng = Range("A5", Range("A5").End(xlDown)).Count - 2
          DoEvents
          Range("A5", Range("A5").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          Range("tbl_enfasis[[ENFASIS_1]:[OBSERVACIONES_AL_ENFASIS_1]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_2]:[OBSERVACIONES AL ENFASIS_2]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_3]:[OBSERVACIONES AL ENFASIS_3]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_4]:[OBSERVACIONES AL ENFASIS_4]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_5]:[OBSERVACIONES AL ENFASIS_5]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_6]:[OBSERVACIONES AL ENFASIS_6]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_7]:[OBSERVACIONES AL ENFASIS_7]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_8]:[OBSERVACIONES AL ENFASIS_8]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_9]:[OBSERVACIONES AL ENFASIS_9]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_10]:[OBSERVACIONES AL ENFASIS_10]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_11]:[OBSERVACIONES AL ENFASIS_11]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_12]:[OBSERVACIONES AL ENFASIS_12]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_13]:[OBSERVACIONES AL ENFASIS_13]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_14]:[OBSERVACIONES AL ENFASIS_14]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_15]:[OBSERVACIONES AL ENFASIS_15]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_16]:[OBSERVACIONES AL ENFASIS_16]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_17]:[OBSERVACIONES AL ENFASIS_17]]").ClearContents
          Range("tbl_enfasis[[ENFASIS_18]:[OBSERVACIONES AL ENFASIS_18]]").ClearContents
        End If

        diag.Select
        Range("A5").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando diagnosticos por favor espere..."
          rng = Range("A5", Range("A5").End(xlDown)).Count - 2
          DoEvents
          Range("A5", Range("A5").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          Range("tbl_diagnosticos[[CODIGO DIAG PPAL]:[DIAG REL 20]]").ClearContents
        End If

        emo.Select
        Range("A5").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando emo por favor espere..."
          rng = Range("A5", Range("A5").End(xlDown)).Count - 2
          DoEvents
          Range("A5", Range("A5").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          ThisWorkbook.Worksheets("RUTAS").Range("$F$5") = CLngLng(Trim(Range("$EL$5").value)) + 1
        Else
          ThisWorkbook.Worksheets("RUTAS").Range("$F$5") = CLngLng(Trim(Range("$EL$5").value)) + 1
        End If

        audio.Select
        Range("A4").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando audio por favor espere..."
          rng = Range("A4", Range("A4").End(xlDown)).Count - 2
          DoEvents
          Range("A4", Range("A4").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          ThisWorkbook.Worksheets("RUTAS").Range("$F$6") = CLngLng(Trim(Range("$BG$4").value)) + 1
        Else
          ThisWorkbook.Worksheets("RUTAS").Range("$F$6") = CLngLng(Trim(Range("$BG$4").value)) + 1
        End If

        opto.Select
        Range("A4").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando opto por favor espere..."
          rng = Range("A4", Range("A4").End(xlDown)).Count - 2
          DoEvents
          Range("A4", Range("A4").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          ThisWorkbook.Worksheets("RUTAS").Range("$F$7") = CLngLng(Trim(Range("$BL$4").value)) + 1
          ThisWorkbook.Worksheets("RUTAS").Range("$F$8") = CLngLng(Trim(Range("$BM$4").value)) + 1
        Else
          ThisWorkbook.Worksheets("RUTAS").Range("$F$7") = CLngLng(Trim(Range("$BL$4").value)) + 1
          ThisWorkbook.Worksheets("RUTAS").Range("$F$8") = CLngLng(Trim(Range("$BM$4").value)) + 1
        End If

        visio.Select
        Range("A4").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando visio por favor espere..."
          rng = Range("A4", Range("A4").End(xlDown)).Count - 2
          DoEvents
          Range("A4", Range("A4").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          ThisWorkbook.Worksheets("RUTAS").Range("$F$9") = CLngLng(Trim(Range("$BS$4").value)) + 1
        Else
          ThisWorkbook.Worksheets("RUTAS").Range("$F$9") = CLngLng(Trim(Range("$BS$4").value)) + 1
        End If

        espiro.Select
        Range("A4").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando espiro por favor espere..."
          rng = Range("A4", Range("A4").End(xlDown)).Count - 2
          DoEvents
          Range("A4", Range("A4").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          ThisWorkbook.Worksheets("RUTAS").Range("$F$10") = CLngLng(Trim(Range("$BZ$4").value)) + 1
        Else
          ThisWorkbook.Worksheets("RUTAS").Range("$F$10") = CLngLng(Trim(Range("$BZ$4").value)) + 1
        End If

        osteo.Select
        Range("A4").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando osteo por favor espere..."
          rng = Range("A4", Range("A4").End(xlDown)).Count - 2
          DoEvents
          Range("A4", Range("A4").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          ThisWorkbook.Worksheets("RUTAS").Range("$F$11") = CLngLng(Trim(Range("$BG$4").value)) + 1
        Else
          ThisWorkbook.Worksheets("RUTAS").Range("$F$11") = CLngLng(Trim(Range("$BG$4").value)) + 1
        End If

        complementarios.Select
        Range("A4").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando complementarios por favor espere..."
          rng = Range("A4", Range("A4").End(xlDown)).Count - 2
          DoEvents
          Range("A4", Range("A4").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          ThisWorkbook.Worksheets("RUTAS").Range("$F$12") = CLngLng(Trim(Range("$J$4").value)) + 1
        Else
          ThisWorkbook.Worksheets("RUTAS").Range("$F$12") = CLngLng(Trim(Range("$J$4").value)) + 1
        End If

        psicotecnica.Select
        Range("A2").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando psicotecnica por favor espere..."
          rng = Range("A2", Range("A2").End(xlDown)).Count - 2
          DoEvents
          Range("A2", Range("A2").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          ThisWorkbook.Worksheets("RUTAS").Range("$F$13") = CLngLng(Trim(Range("$G$2").value)) + 1
        Else
          ThisWorkbook.Worksheets("RUTAS").Range("$F$13") = CLngLng(Trim(Range("$G$2").value)) + 1
        End If

        psicosensometrica.Select
        Range("A3").Select
        Selection.ListObject.Range.FormatConditions.Delete
        If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
          Application.StatusBar = "Limpiando psicosensometrica por favor espere..."
          rng = Range("A3", Range("A3").End(xlDown)).Count - 2
          DoEvents
          Range("A3", Range("A3").Offset(rng, 0)).Select
          Selection.EntireRow.Delete shift:=xlUp
          ThisWorkbook.Worksheets("RUTAS").Range("$F$14") = CLngLng(Trim(Range("$Q$3").value)) + 1
        Else
          ThisWorkbook.Worksheets("RUTAS").Range("$F$14") = CLngLng(Trim(Range("$Q$3").value)) + 1
        End If

        Application.ActiveWorkbook.SaveCopyAs Filename:=route & "\" & Application.ActiveWorkbook.Name
        Application.StatusBar = Empty
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True

        MsgBox "Almacenamiento terminado", vbOKOnly + vbInformation, "Almacenamiento"
      Else

        MsgBox "No hay datos para almacenar", vbOKOnly + vbInformation, "Almacenamiento"

      End If

      Call statusDesactivate(trabajadores.Name)
      Call statusDesactivate(emo.Name)
      Call statusDesactivate(audio.Name)
      Call statusDesactivate(visio.Name)
      Call statusDesactivate(opto.Name)
      Call statusDesactivate(espiro.Name)
      Call statusDesactivate(osteo.Name)
      Call statusDesactivate(complementarios.Name)
      Call statusDesactivate(psicotecnica.Name)
      Call statusDesactivate(psicosensometrica.Name)

      trabajadores.Select
      Range("A5").Select

End Sub

Sub Modification()

  Dim consolidado, libro, esLibro As Object
  Dim dateSmall As Date
  Dim Name, msg As String
  Dim patch As Variant

  libro = Worksheets("RUTAS").Range("C5").value

  If (Range("$B$6").value <> "") Then

    msg = Application.InputBox(prompt:="Indica el mensaje de la modificaci" & Chr(243) & "n efectuada", _
    Default:="", Type:=2)

    If (Trim(msg) = Empty) Then
      MsgBox prompt:="Las observaciones no pueden estar vacias", Buttons:=vbOKOnly, Title:="Error msg"
      Exit Sub
    End If

    Set esLibro = Application.ThisWorkbook
    patch = VBA.Split(esLibro.FullName, "\")
    Name = VBA.Split(esLibro.Name, ".")

    dateSmall = CDate(patch(8))
    Set consolidado = Workbooks.Open(libro)

    consolidado.Worksheets("Registros").Select
    consolidado.ActiveSheet.Unprotect Password:="1024500065"
    Range("B2").Select
    Cells.Find(What:=dateSmall, After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate

    Do While ActiveCell = dateSmall
      If ActiveCell = dateSmall And ActiveCell.Offset(, 1).value = Name(0) Then
        ActiveCell.Offset(, 7) = msg & " - Date Modified: " & Date
        Call UpdateGoogleSheetRecord(ActiveCell.Row - 1, msg & " - Date Modified: " & Date)
        consolidado.ActiveSheet.Protect Password:="1024500065", DrawingObjects:=False, Contents:=True, Scenarios:= _
        False, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:= _
        True
        consolidado.Save
        consolidado.Close
      End If
      ActiveCell.Offset(1, 0).Select
    Loop

    MsgBox prompt:="Se ha registrado la modificaci" & Chr(243) & "n", Buttons:=vbInformation + vbOKOnly, Title:="Registro Exitoso"

  End If
End Sub

Sub AddRecordToGoogleSheet(ByVal Company as String, ByVal sigad as String, ByVal orden as Integer, ByVal patience as Integer, ByVal libro As Variant, ByVal bookNow as String)

  '' ya funciona usa el token oAuth2

  Dim HttpReq As Variant
  Dim Json As Object
  Dim monthNow, yearNow As Integer
  Dim fullDate, dateNow,  bearerToken As String

  fullDate = Format(Now, "dd/mm/yyyy hh:mm:ss")
  dateNow = Format(Date, "dd-mmm-yyyy")
  monthNow = Month(Date)
  yearNow = Year(Date)
  bearerToken = Application.InputBox(prompt:="ingrese el Token de Acceso", title:="Acceso Google Sheet", Default:="", Type:=2)

  Set HttpReq = CreateObject("MSXML2.XMLHTTP")
  HttpReq.Open "POST", "https://sheets.googleapis.com/v4/spreadsheets/126vzNrB3mA-g-61ccgNyAz-ukhIIqg_Yn3JxzQljC5o/values/Registro!$A2:append?valueInputOption=RAW", False
  HttpReq.setRequestHeader "Authorization", "Bearer " & Trim(bearerToken)
  HttpReq.setRequestHeader "Content-Type", "application/json"

  Dim requestBody As String
  requestBody = "{""values"":[['" & fullDate & "','" & dateNow & "','" & Company & "','" & sigad & "'," & orden & "," & patience & "," & monthNow & "," & yearNow & "]]}"

  On Error Resume Next
  HttpReq.send (requestBody)

  If HttpReq.status = 200 Then
    MsgBox "Record added successfully:"+ vbNewLine + vbNewLine + Chr(32) +"code:" & HttpReq.status & ""+ vbNewLine + Chr(32)+"status:"& HttpReq.statusText
  ElseIf HttpReq.status = 12031 then
    MsgBox "Restriction by network administrator:"+ vbNewLine + vbNewLine + Chr(32) +"code:" & HttpReq.status
    Workbooks.Open(libro)
    Windows(bookNow).Activate
  Else
    MsgBox "Error adding record: " & HttpReq.status & vbNewLine & HttpReq.statusText & vbNewLine & HttpReq.responseText
    Workbooks.Open(libro)
    Windows(bookNow).Activate
  End If
End Sub

Sub UpdateGoogleSheetRecord(ByVal rowData As Integer, ByVal textModify As String)

  Dim sheetId, accessToken As String
  sheetId = "126vzNrB3mA-g-61ccgNyAz-ukhIIqg_Yn3JxzQljC5o"
  accessToken = Application.InputBox(prompt:="ingrese el Token de Acceso", Title:="Acceso Google Sheet", Default:="", Type:=2)

  Dim range As String
  range = "Registro!I" & rowData

  Dim requestBody As String
  requestBody = "{""values"": [['" & textModify & "']]}"

  Dim url As String
  url = "https://sheets.googleapis.com/v4/spreadsheets/" & sheetId & "/values/" & range & "?valueInputOption=RAW"

  Dim httpObject As Object
  Set httpObject = CreateObject("MSXML2.XMLHTTP")

  httpObject.Open "PUT", url, False
  httpObject.setRequestHeader "Content-Type", "application/json"
  httpObject.setRequestHeader "Authorization", "Bearer " & accessToken

  On Error Resume Next
  httpObject.send (requestBody)

  If (httpObject.status = 200) Then
    MsgBox "Record updated successfully:" + vbNewLine + vbNewLine + Chr(32) + "code:" & httpObject.status & "" + vbNewLine + Chr(32) + "status:" & httpObject.statusText
  ElseIf (httpObject.status = 12031) Then
    MsgBox "Restriction by network administrator:" + vbNewLine + vbNewLine + Chr(32) + "code:" & httpObject.status
  Else
    MsgBox "Error updated record: " & httpObject.status & vbNewLine & httpObject.statusText & vbNewLine & httpObject.responseText
  End If

End Sub

Sub ExportSQL()

  Dim origin As Workbook
  Dim comple_origin, worker_origin, osteo_origin, senso_origin, psico_origin, visio_origin, espiro_origin, opto_origin, audio_origin, emo_origin As Worksheet
  Dim sh, str, MyFile As Variant
  Dim num As Integer
  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")

  Set origin = Workbooks(ThisWorkbook.Name)
  Set worker_origin = origin.Worksheets("TRABAJADORES")
  Set emo_origin = origin.Worksheets("EMO")
  Set audio_origin = origin.Worksheets("AUDIO")
  Set opto_origin = origin.Worksheets("OPTO")
  Set espiro_origin = origin.Worksheets("ESPIRO")
  Set osteo_origin = origin.Worksheets("OSTEO")
  Set visio_origin = origin.Worksheets("VISIO")
  Set psico_origin = origin.Worksheets("PSICOTECNICA")
  Set senso_origin = origin.Worksheets("PSICOSENSOMETRICA")
  Set comple_origin = origin.Worksheets("COMPLEMENTARIOS")

  Set MyFile = FSO.OpenTextFile(ThisWorkbook.Worksheets("RUTAS").Range("C9").value &"testfile.sql", ForAppending, True, TristateTrue)
  For Each sh In origin.Worksheets
    If ActiveWorkbook.Sheets(sh.Name).Tab.ThemeColor = xlThemeColorAccent1 Then
      Select Case Trim(Ucase(sh.Name))
       Case "TRABAJADORES"
        ' orden lista trabajadores
        num = isEmptyValue(range("tbl_trabajadores[[SCRIPT orden_lista_trabajadores]]")) 
        If ( num > 0) Then
          MyFile.WriteLine "INSERT INTO orden_lista_trabajadores (`id`, `id_orden`, `estado`, `cedula`, `nombre`, `telefono`, `registro`, `ciudad_id`, `empresa_id`, `digitador_id`, `fecha_ingreso`, `id_cargo`, `fuente`, `edad`, `genero`, `estrato`, `id_raza`, `id_estado_civil`, `hijos`, `id_escolaridad`, `rango_edad`, `duracion`, `antiguedad`, `created_at`, `updated_at`, `id_tipo_actividad`, `id_tipo_examen`) VALUES"
          For Each Item In range("tbl_trabajadores[[SCRIPT orden_lista_trabajadores]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' paraclinicos
        num = isEmptyValue(range("tbl_trabajadores[[SCRIPT ordenes_trabajador_paraclinicos]]")) 
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ordenes_trabajador_paraclinicos (`id_orden_trabajador`, `id_paraclinico`, `estado`) VALUES"
          For Each Item In range("tbl_trabajadores[[SCRIPT ordenes_trabajador_paraclinicos]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If
       Case "EMO"
        ' ics_emo
        num = isEmptyValue(range("tbl_emo[[SCRIPT ics_emo]]")) 
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ics_emo (`id`, `id_orden_lista_trabajadores`, `id_concepto_evaluacion`, `observaciones`, `accidente_laboral`, `enfermedad_laboral`, `fecha_accidente`, `empresa`, `naturaleza_lesion`, `tipo_accidente`, `parte_afectada`, `dias_incapacidad`, `secuelas`, `enfermedad`, `etapa`, `observaciones_enfermedad`, `actividad_fisica`, `fuma`, `consumo_alcohol`, `peso`, `talla`, `tension_arterial`, `frecuencia_cardiaca`, `perimetro_abominal`, `lateralidad`, `frecuencia_respiratoria`, `imc2`, `clasificacion_imc`, `observacion_recomendacion`, `observacion_diagnostico`) VALUES"
          For Each Item In range("tbl_emo[[SCRIPT ics_emo]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' ics_emo_riesgos
        num = isEmptyValue(range("tbl_emo[[SCRIPT ics_emo_riesgos]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ics_emo_riesgos (`id_ics`, `id_riesgo`, `observaciones_otros`) VALUES"
          For Each Item In range("tbl_emo[[SCRIPT ics_emo_riesgos]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' ics_condiciones
        num = isEmptyValue(range("tbl_emo[[SCRIPT ics_condiciones]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ics_condiciones (`id_ics`, `id_condicion`, `condicion_seguridad`) VALUES"
          For Each Item In range("tbl_emo[[SCRIPT ics_condiciones]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' ics_cie (diagnosticos)
        num = isEmptyValue(range("tbl_emo[[script ics_cie (diagnosticos)]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ics_cie (`id_ics`, `id_cie`) VALUES"
          For Each Item In range("tbl_emo[[script ics_cie (diagnosticos)]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' ics_enfasis
        num = isEmptyValue(range("tbl_emo[[SCRIPT ics_enfasis]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ics_enfasis (`id_ics`, `id_enfasis`, `observacion`) VALUES"
          For Each Item In range("tbl_emo[[SCRIPT ics_enfasis]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

       Case "AUDIO"
        ' au_audiometria
        num = isEmptyValue(range("tbl_audio[[SCRIPT au_audiometria]]")) 
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO au_audiometria (`id`, `emo_id`, `auditivo`, `auditivo_copa`, `auditivo_insercion`, `auditivo_doble`, `diagnostico_interno`, `diagnostico_ppal`, `diagnostico_gati`, `status_obs`) VALUES"
          For Each Item In range("tbl_audio[[SCRIPT au_audiometria]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' au_audiometria_recomendacion
        num = isEmptyValue(range("tbl_audio[[SCRIPT au_audiometria_recomendacion]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO au_audiometria_recomendacion (`audiometria_id`, `recomendacion_id`, `fk_id_control`) VALUES"
          For Each Item In range("tbl_audio[[SCRIPT au_audiometria_recomendacion]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' au_oido
        num = isEmptyValue(range("tbl_audio[[SCRIPT au_oido]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO au_oido (`audiometria_id`, `tipo_oido_id`, `pabellon_id`, `auditivo_id`, `membrana_id`, `obs_pabellon`, `obs_auditivo`, `obs_membrana`, `frecuencia`, `pta`) VALUES"
          For Each Item In range("tbl_audio[[SCRIPT au_oido]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

       Case "OPTO"
        ' op_optometria
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_optometria]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_optometria (`id`, `emo_id`, `parpados`, `obs_parpados`, `conjuntivas`, `obs_conjuntivas`, `escleras`, `obs_escleras`, `pupilas`, `obs_pupilas`, `lejos`, `cerca`, `patologia_ocular`, `estado_correcion_id`, `otros_sintomas`, `recomendacion`, `remision`, `status_dig`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_optometria]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' op_optometria_riesgos
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_optometria_riesgos]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_optometria_riesgos (`optometria_id`, `riesgo_id`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_optometria_riesgos]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' op_optometria_sintomas
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_optometria_sintomas]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_optometria_sintomas (`optometria_id`, `sintomas_id`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_optometria_sintomas]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' op_diagnostico
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_diagnostico]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_diagnostico (`id`, `optometria_id`, `diagnostico_ppal`, `obs_ppal`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_diagnostico]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' op_diagnostico_cie
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_diagnostico_cie]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_diagnostico_cie (`diagnostico_id`, `cie_id`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_diagnostico_cie]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' op_optometria_recomendacion
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_optometria_recomendacion]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_optometria_recomendacion (`optometria_id`, `recomendacion_id`, `fk_control_id`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_optometria_recomendacion]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' op_optometria_remision
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_optometria_remision]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_optometria_remision (`optometria_id`, `remision_id`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_optometria_remision]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

       Case "VISIO"
        ' vi_visiometria
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_visiometria]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_visiometria (`id`, `emo_id`, `parpados`, `obs_parpados`, `conjuntivas`, `obs_conjuntivas`, `escleras`, `obs_escleras`, `pupilas`, `obs_pupilas`, `otros_sintomas`, `resultado`, `obs_resultado`, `recomendacion_general`, `remision`, `status_general`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_visiometria]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' vi_visiometria_antecedentes
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_visiometria_antecedentes]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_visiometria_antecedentes (`visiometria_id`, `antecedente_id`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_visiometria_antecedentes]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' vi_visiometria_sintomas
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_visiometria_sintomas]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_visiometria_sintomas (`visiometria_id`, `sintoma_id`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_visiometria_sintomas]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' vi_vl
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_vl]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_vl (`visiometria_id`, `oi`, `od`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_vl]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' vi_vp
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_vp]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_vp (`visiometria_id`, `oi`, `od`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_vp]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' vi_visiometria_recomendaciones
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_visiometria_recomendaciones]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_visiometria_recomendaciones (`visiometria_id`, `recomendacion_id`, `fk_id_control`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_visiometria_recomendaciones]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' vi_visiometria_remisiones
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_visiometria_remisiones]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_visiometria_remisiones (`visiometria_id`, `remision_id`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_visiometria_remisiones]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

       Case "ESPIRO"
        ' espirometria
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espirometria]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espirometria (`id`, `emo_id`, `observaciones_alergias`, `observaciones_cx_torax`, `observaciones_cancer`, `otros_respiratorios`, `otros_riesgos_quimicos`, `actividad_fisica`, `fuma`, `frecuencia_habito`, `numero_cigarrros`, `tiempo_anios`, `interpretaciones`, `tecnica_aceptable`, `calculos_diagnostico`, `diagnostico_ppal`, `observacion_ppal`, `tipo_interpretacion`, `tipo_grado`, `resultado_espiro`, `peso`, `talla`, `imc2`, `clasificacion_imc`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espirometria]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' espiro_antecedentes_pivot
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espiro_antecedentes_pivot]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espiro_antecedentes_pivot (`espiro_id`, `id_antecedente`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espiro_antecedentes_pivot]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' espiro_quimicos_pivot
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espiro_quimicos_pivot]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espiro_quimicos_pivot (`espiro_id`, `id_quimicos`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espiro_quimicos_pivot]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' espiro_riesgos_epp
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espiro_riesgos_epp]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espiro_riesgos_epp (`espiro_id`, `tapaboca`, `especifico`, `otro_tapaboca`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espiro_riesgos_epp]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' espiro_recomendaciones_pivot
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espiro_recomendaciones_pivot]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espiro_recomendaciones_pivot(`espiro_id`, `recomendaciones_id`, `fk_id_control`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espiro_recomendaciones_pivot]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' espiro_recomendaciones_lab_pivot
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espiro_recomendaciones_lab_pivot]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espiro_recomendaciones_lab_pivot (`espiro_id`, `recomendaciones_id`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espiro_recomendaciones_lab_pivot]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

       Case "OSTEO"
        ' osteomuscular
        num = isEmptyValue(range("tbl_osteo[[SCRIPT osteomuscular]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO osteomuscular (`id`, `emo_id`,`diagnostico_ppal`, `observacion_ppal`, `ocupacionales`, `generales`, `status_ppal`) VALUES"
          For Each Item In range("tbl_osteo[[SCRIPT osteomuscular]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' osteo_antecedentes_pivot
        num = isEmptyValue(range("tbl_osteo[[SCRIPT osteo_antecedentes_pivot]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO osteo_antecedentes_pivot (`osteo_id`, `id_antecedente_osteo`, `observacion_antecedente_sintoma`) VALUES"
          For Each Item In range("tbl_osteo[[SCRIPT osteo_antecedentes_pivot]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' osteo_cie_pivot
        num = isEmptyValue(range("tbl_osteo[[SCRIPT osteo_cie_pivot]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO osteo_cie_pivot (`osteo_id`, `cie_id`) VALUES"
          For Each Item In range("tbl_osteo[[SCRIPT osteo_cie_pivot]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' osteo_recomendaciones_pivot
        num = isEmptyValue(range("tbl_osteo[[SCRIPT osteo_recomendaciones_pivot]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO osteo_recomendaciones_pivot (`osteo_id`, `id_recomendaciones_osteo`) VALUES"
          For Each Item In range("tbl_osteo[[SCRIPT osteo_recomendaciones_pivot]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

       Case "COMPLEMENTARIOS","COMPLEMENTARIO"
        ' complementarios
        num = isEmptyValue(range("tbl_complementarios[[SCRIPT complementarios]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO complementarios (`id`, `emo_id`, `procedimiento_id`, `diagnostico_ppal`, `observacion_ppal`, `hallazgo`, `status_ppal`) VALUES"
          For Each Item In range("tbl_complementarios[[SCRIPT complementarios]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' complementarios_diagnos_observaciones_pivot
        num = isEmptyValue(range("tbl_complementarios[[SCRIPT complementarios_diagnos_observaciones_pivot]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO complementarios_diagnos_observaciones_pivot (`complementarios_id`, `diagnostico`) VALUES"
          For Each Item In range("tbl_complementarios[[SCRIPT complementarios_diagnos_observaciones_pivot]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

       Case "PSICOTECNICA","PSICOLOGIA"
        ' psicotecnica
        num = isEmptyValue(range("tbl_psicotecnica[[SCRIPT psicotecnica]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO psicotecnica (`id`, `emo_id`, `prueba`, `id_diagnostico_ppal`, `observacion_ppal`) VALUES"
          For Each Item In range("tbl_psicotecnica[[SCRIPT psicotecnica]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

       Case "PSICOSENSOMETRICA","PSICOMOTRIZ"
        ' psicosensometrica
        num = isEmptyValue(range("tbl_psicosensometrica[[SCRIPT psicosensometrica]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO psicosensometrica (`id`, `emo_id`, `prueba`, `id_diagnostico_ppal`, `observacion_ppal`) VALUES"
          For Each Item In range("tbl_psicosensometrica[[SCRIPT psicosensometrica]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' psicosenso_diagnos_observaciones_pivot
        num = isEmptyValue(range("tbl_psicosensometrica[[SCRIPT psicosenso_diagnos_observaciones_pivot]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO psicosenso_diagnos_observaciones_pivot (`psicosensometrica_id`, `diagnostico`) VALUES"
          For Each Item In range("tbl_psicosensometrica[[SCRIPT psicosenso_diagnos_observaciones_pivot]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If

        ' psicosensometricas_recomendaciones_pivot
        num = isEmptyValue(range("tbl_psicosensometrica[[SCRIPT psicosensometricas_recomendaciones_pivot]]"))
        If ( num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO psicosensometricas_recomendaciones_pivot (`psicosensometrica_id`, `recomendaciones_id`) VALUES"
          For Each Item In range("tbl_psicosensometrica[[SCRIPT psicosensometricas_recomendaciones_pivot]]")
            If Item <> "" And num <> 1 then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 then
              MyFile.WriteLine Item & ";"
              num = num - 1
            End If
          Next Item
        End If
      End Select
    End If
  Next sh
  MyFile.Close

  MsgBox "Se genero el archivo SQL textfile.sql" + vbNewLine + vbNewLine + Chr(32) + "Que se encuentra en la ruta: " + vbNewLine + vbNewLine + ThisWorkbook.Worksheets("RUTAS").Range("C9").value
End Sub

Public Function isEmptyValue(ByVal Ranges As Object) As Integer
  Dim num As Integer
  Dim Item As Variant
  
  num = 0
  For Each Item In Ranges
    If (Item <> "") Then: num = num + 1
  Next Item 
  isEmptyValue = num
  
End Function

