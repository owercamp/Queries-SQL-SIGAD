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

