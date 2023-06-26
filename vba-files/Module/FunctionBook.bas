Attribute VB_Name = "FunctionBook"
Option Explicit
Public sigad As Variant
Public trabajadores As Worksheet, emo As Worksheet, audio As Worksheet, visio As Worksheet, opto As Worksheet, espiro As Worksheet, osteo As Worksheet, complementarios As Worksheet, psicotecnica As Worksheet, psicosensometrica As Worksheet, enfasis As Worksheet, diag As Worksheet

'''? <summary>
'''? Abre un libro de Excel especificado en la hoja "RUTAS" y celda "C7".
'''? </summary>
Public Sub cargos()
Attribute cargos.VB_ProcData.VB_Invoke_Func = "k\n14"
  Workbooks.Open (ThisWorkbook.Worksheets("RUTAS").range("C7").value)
End Sub

'TODO: Crea una carpeta en la ruta especificada y guarda una copia del libro de trabajo activo en ella. Si la carpeta ya existe, simplemente guarda una copia del libro de trabajo activo en ella.

'* Argumentos:
'* @param route: la ruta donde se creara la carpeta y se guardara una copia del libro de trabajo activo.
'* @param folderName: el nombre de la carpeta que se creara en la ruta especificada.
'* @param workbookActive: el nombre del libro de trabajo activo que se guardara en la carpeta creada.
'* @param YearNow: el año actual, que se utilizara para crear una subcarpeta en la carpeta creada.
'* @param MonthNow: el mes actual, que se utilizara para crear una subcarpeta en la carpeta creada.
Public Sub folder(route, folderName, workbookActive, yearNow, monthNow)
  Dim splitRoute As String
  splitRoute = Application.PathSeparator

  If Dir(route, vbDirectory) = Empty Then
    MkDir route
  End If
  If Dir(route & splitRoute & yearNow, vbDirectory) = Empty Then
    MkDir (route & splitRoute & yearNow)
  End If
  If Dir(route & splitRoute & yearNow & splitRoute & monthNow, vbDirectory) = Empty Then
    MkDir (route & splitRoute & yearNow & splitRoute & monthNow)
  End If

  If Dir(route & splitRoute & yearNow & splitRoute & monthNow & splitRoute & folderName, vbDirectory) = Empty Then
    MkDir (route & splitRoute & yearNow & splitRoute & monthNow & splitRoute & folderName)
    Application.ActiveWorkbook.SaveCopyAs Filename:=route & splitRoute & yearNow & splitRoute & monthNow & splitRoute & folderName & splitRoute & workbookActive
    Application.StatusBar = "se guardo una copia en: " & route & splitRoute & yearNow & splitRoute & monthNow & splitRoute & folderName & splitRoute & workbookActive
  Else
    Application.ActiveWorkbook.SaveCopyAs Filename:=route & splitRoute & yearNow & splitRoute & monthNow & splitRoute & folderName & splitRoute & workbookActive
    Application.StatusBar = "se guardo una copia en: " & route & splitRoute & yearNow & splitRoute & monthNow & splitRoute & folderName & splitRoute & workbookActive
  End If
End Sub

Public Sub clearContents()

  Dim rng As Integer, info As Integer, rngTrabajadores As Integer, rngEmo As Integer, rngAudio As Integer, rngVisio As Integer, rngOpto As Integer, rngEspiro As Integer, rngOsteo As Integer, rngComplementarios As Integer, rngPsicotecnica As Integer, rngPsicosensometrica As Integer, rngEnfasis As Integer, rngDiag As Integer, MyDay As Integer, MyMonth As Integer, MyYear As Integer
  Dim meses As Variant, finalRow As Variant, libro As Variant, RowActive As Variant
  Dim nombre As String, orden As String, fecha As String, Company As String, bookNow As String
  Dim consolidado As Object

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
  If trabajadores.range("D6") <> Empty Or trabajadores.range("D6") <> vbNullString Then
    nombre = trabajadores.range("B6").value & " - " & trabajadores.range("D6").value & ".xlsb"
  End If
  If trabajadores.range("D6") = Empty Or trabajadores.range("D6") = vbNullString Then
    nombre = trabajadores.range("B6").value & ".xlsb"
  End If

  orden = trabajadores.range("AX6").value

  If (Not IsEmpty(nombre)) And (Not IsEmpty(orden)) And (Not IsEmpty(sigad)) Then
    fecha = CStr(MyDay) + " " + CStr(meses(MyMonth - 1)) + " " + CStr(MyYear)

    route = CStr(Worksheets("RUTAS").range("C6").value)

    trabajadores.Select
    Call folder(route, fecha, nombre, MyYear, meses(MyMonth - 1))

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False


    ''' REGISTRO EN CONSOLIDADO ''
    info = Worksheets("TRABAJADORES").range("A5", Worksheets("TRABAJADORES").range("A5").End(xlDown)).Count

    libro = Worksheets("RUTAS").range("C5").value
    If trabajadores.range("D5").value = Empty Or trabajadores.range("D5").value = vbNullString Then
      Company = trabajadores.range("B5").value
    Else
      Company = trabajadores.range("B5").value & " - " & trabajadores.range("D5").value
    End If

    Set consolidado = Workbooks.Open(libro)
    ''' TRABAJADORES ''

    consolidado.Worksheets("Registros").Select
    consolidado.ActiveSheet.Unprotect Password:="1024500065"
    range("C3").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell = Trim(UCase(Company))
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

    Call AddRecordToGoogleSheet(Trim(UCase(Company)), Trim(UCase("ICS-" & PadLeft(sigad, 4, "0"))), Trim(orden), Trim(info), libro, bookNow)

    Application.Calculation = xlCalculationManual
    trabajadores.Select
    route = Worksheets("RUTAS").range("C8").value

    range("A5").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando trabajadores por favor espere..."
      rng = range("A5", range("A5").End(xlDown)).Count - 2
      DoEvents
      range("A5", range("A5").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      ThisWorkbook.Worksheets("RUTAS").range("$F$4") = CLngPtr(Trim(range("$AW$5").value)) + 1
    Else
      ThisWorkbook.Worksheets("RUTAS").range("$F$4") = CLngPtr(Trim(range("$AW$5").value)) + 1
    End If

    enfasis.Select
    range("A5").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando enfasis por favor espere..."
      rng = range("A5", range("A5").End(xlDown)).Count - 2
      DoEvents
      range("A5", range("A5").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      range("tbl_enfasis[[ENFASIS_1]:[OBSERVACIONES_AL_ENFASIS_1]]").clearContents
      range("tbl_enfasis[[ENFASIS_2]:[OBSERVACIONES AL ENFASIS_2]]").clearContents
      range("tbl_enfasis[[ENFASIS_3]:[OBSERVACIONES AL ENFASIS_3]]").clearContents
      range("tbl_enfasis[[ENFASIS_4]:[OBSERVACIONES AL ENFASIS_4]]").clearContents
      range("tbl_enfasis[[ENFASIS_5]:[OBSERVACIONES AL ENFASIS_5]]").clearContents
      range("tbl_enfasis[[ENFASIS_6]:[OBSERVACIONES AL ENFASIS_6]]").clearContents
      range("tbl_enfasis[[ENFASIS_7]:[OBSERVACIONES AL ENFASIS_7]]").clearContents
      range("tbl_enfasis[[ENFASIS_8]:[OBSERVACIONES AL ENFASIS_8]]").clearContents
      range("tbl_enfasis[[ENFASIS_9]:[OBSERVACIONES AL ENFASIS_9]]").clearContents
      range("tbl_enfasis[[ENFASIS_10]:[OBSERVACIONES AL ENFASIS_10]]").clearContents
      range("tbl_enfasis[[ENFASIS_11]:[OBSERVACIONES AL ENFASIS_11]]").clearContents
      range("tbl_enfasis[[ENFASIS_12]:[OBSERVACIONES AL ENFASIS_12]]").clearContents
      range("tbl_enfasis[[ENFASIS_13]:[OBSERVACIONES AL ENFASIS_13]]").clearContents
      range("tbl_enfasis[[ENFASIS_14]:[OBSERVACIONES AL ENFASIS_14]]").clearContents
      range("tbl_enfasis[[ENFASIS_15]:[OBSERVACIONES AL ENFASIS_15]]").clearContents
      range("tbl_enfasis[[ENFASIS_16]:[OBSERVACIONES AL ENFASIS_16]]").clearContents
      range("tbl_enfasis[[ENFASIS_17]:[OBSERVACIONES AL ENFASIS_17]]").clearContents
      range("tbl_enfasis[[ENFASIS_18]:[OBSERVACIONES AL ENFASIS_18]]").clearContents
    End If

    diag.Select
    range("A5").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando diagnosticos por favor espere..."
      rng = range("A5", range("A5").End(xlDown)).Count - 2
      DoEvents
      range("A5", range("A5").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      range("tbl_diagnosticos[[CODIGO DIAG PPAL]:[DIAG REL 20]]").clearContents
    End If

    emo.Select
    range("A5").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando emo por favor espere..."
      rng = range("A5", range("A5").End(xlDown)).Count - 2
      DoEvents
      range("A5", range("A5").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      ThisWorkbook.Worksheets("RUTAS").range("$F$5") = CLngPtr(Trim(range("$EL$5").value)) + 1
    Else
      ThisWorkbook.Worksheets("RUTAS").range("$F$5") = CLngPtr(Trim(range("$EL$5").value)) + 1
    End If

    audio.Select
    range("A4").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando audio por favor espere..."
      rng = range("A4", range("A4").End(xlDown)).Count - 2
      DoEvents
      range("A4", range("A4").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      ThisWorkbook.Worksheets("RUTAS").range("$F$6") = CLngPtr(Trim(range("$BG$4").value)) + 1
    Else
      ThisWorkbook.Worksheets("RUTAS").range("$F$6") = CLngPtr(Trim(range("$BG$4").value)) + 1
    End If

    opto.Select
    range("A4").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando opto por favor espere..."
      rng = range("A4", range("A4").End(xlDown)).Count - 2
      DoEvents
      range("A4", range("A4").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      ThisWorkbook.Worksheets("RUTAS").range("$F$7") = CLngPtr(Trim(range("$BL$4").value)) + 1
      ThisWorkbook.Worksheets("RUTAS").range("$F$8") = CLngPtr(Trim(range("$BM$4").value)) + 1
    Else
      ThisWorkbook.Worksheets("RUTAS").range("$F$7") = CLngPtr(Trim(range("$BL$4").value)) + 1
      ThisWorkbook.Worksheets("RUTAS").range("$F$8") = CLngPtr(Trim(range("$BM$4").value)) + 1
    End If

    visio.Select
    range("A4").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando visio por favor espere..."
      rng = range("A4", range("A4").End(xlDown)).Count - 2
      DoEvents
      range("A4", range("A4").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      ThisWorkbook.Worksheets("RUTAS").range("$F$9") = CLngPtr(Trim(range("$BS$4").value)) + 1
    Else
      ThisWorkbook.Worksheets("RUTAS").range("$F$9") = CLngPtr(Trim(range("$BS$4").value)) + 1
    End If

    espiro.Select
    range("A4").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando espiro por favor espere..."
      rng = range("A4", range("A4").End(xlDown)).Count - 2
      DoEvents
      range("A4", range("A4").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      ThisWorkbook.Worksheets("RUTAS").range("$F$10") = CLngPtr(Trim(range("$BZ$4").value)) + 1
    Else
      ThisWorkbook.Worksheets("RUTAS").range("$F$10") = CLngPtr(Trim(range("$BZ$4").value)) + 1
    End If

    osteo.Select
    range("A4").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando osteo por favor espere..."
      rng = range("A4", range("A4").End(xlDown)).Count - 2
      DoEvents
      range("A4", range("A4").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      ThisWorkbook.Worksheets("RUTAS").range("$F$11") = CLngPtr(Trim(range("$BG$4").value)) + 1
    Else
      ThisWorkbook.Worksheets("RUTAS").range("$F$11") = CLngPtr(Trim(range("$BG$4").value)) + 1
    End If

    complementarios.Select
    range("A4").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando complementarios por favor espere..."
      rng = range("A4", range("A4").End(xlDown)).Count - 2
      DoEvents
      range("A4", range("A4").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      ThisWorkbook.Worksheets("RUTAS").range("$F$12") = CLngPtr(Trim(range("$J$4").value)) + 1
    Else
      ThisWorkbook.Worksheets("RUTAS").range("$F$12") = CLngPtr(Trim(range("$J$4").value)) + 1
    End If

    psicotecnica.Select
    range("A2").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando psicotecnica por favor espere..."
      rng = range("A2", range("A2").End(xlDown)).Count - 2
      DoEvents
      range("A2", range("A2").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      ThisWorkbook.Worksheets("RUTAS").range("$F$13") = CLngPtr(Trim(range("$G$2").value)) + 1
    Else
      ThisWorkbook.Worksheets("RUTAS").range("$F$13") = CLngPtr(Trim(range("$G$2").value)) + 1
    End If

    psicosensometrica.Select
    range("A3").Select
    Selection.ListObject.range.FormatConditions.Delete
    If (ActiveCell <> Empty Or ActiveCell <> vbNullString) And (ActiveCell.Offset(1, 0) <> Empty Or ActiveCell.Offset(1, 0) <> vbNullString) Then
      Application.StatusBar = "Limpiando psicosensometrica por favor espere..."
      rng = range("A3", range("A3").End(xlDown)).Count - 2
      DoEvents
      range("A3", range("A3").Offset(rng, 0)).Select
      Selection.EntireRow.Delete shift:=xlUp
      ThisWorkbook.Worksheets("RUTAS").range("$F$14") = CLngPtr(Trim(range("$Q$3").value)) + 1
    Else
      ThisWorkbook.Worksheets("RUTAS").range("$F$14") = CLngPtr(Trim(range("$Q$3").value)) + 1
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
  range("A5").Select

End Sub

'* Este procedimiento almacena en un libro de trabajo local y en una hoja de calculo de Google Sheet el mensaje de modificacion
'* y la fecha en que se realizo la modificacion para un registro especifico en la hoja "Registros" del libro de trabajo especificado
'* en la celda "C5" de la hoja "RUTAS". Si se introduce un mensaje de modificacion valido, el procedimiento busca la fecha y el nombre
'* del libro de trabajo actual en la hoja de calculo, para luego actualizar la celda correspondiente con el mensaje de modificacion
'* y la fecha actual, y tambien actualizar una hoja de calculo de Google con el mensaje de modificacion y la fecha actual.
Public Sub Modification()

  Dim consolidado As Object, libro As String, esLibro As Object
  Dim dateSmall As Date
  Dim name_book() As String, msg As String
  Dim patch As Variant

  libro = Worksheets("RUTAS").range("C5").value

  If (range("$B$6").value <> "") Then

    msg = Application.InputBox(prompt:="Indica el mensaje de la modificaci" & Chr(243) & "n efectuada", _
    Default:="", Type:=2)

    If (Trim(msg) = Empty) Then
      MsgBox prompt:="Las observaciones no pueden estar vacias", Buttons:=vbOKOnly, Title:="Error msg"
      Exit Sub
    End If

    Set esLibro = Application.ThisWorkbook
    patch = VBA.Split(esLibro.FullName, "\")
    name_book = VBA.Split(esLibro.Name, ".")

    dateSmall = CDate(patch(8))
    Set consolidado = Workbooks.Open(libro)

    consolidado.Worksheets("Registros").Select
    consolidado.ActiveSheet.Unprotect Password:="1024500065"
    range("B2").Select
    Cells.Find(What:=dateSmall, After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate

    Do While ActiveCell = dateSmall
      If ActiveCell = dateSmall And ActiveCell.Offset(, 1).value = name_book(0) Then
        ActiveCell.Offset(, 7) = msg & " - Date Modified: " & Date
        Call UpdateGoogleSheetRecord(ActiveCell.row - 1, msg & " - Date Modified: " & Date)
        consolidado.ActiveSheet.Protect Password:="1024500065", DrawingObjects:=False, Contents:=True, Scenarios:= _
        False, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:= _
        True
        consolidado.Save
      End If
      ActiveCell.Offset(1, 0).Select
    Loop
    Windows(consolidado.Name).Activate
    MsgBox prompt:="Se ha registrado la modificaci" & Chr(243) & "n", Buttons:=vbInformation + vbOKOnly, Title:="Registro Exitoso"

  End If
End Sub

'TODO: Agrega un registro a una hoja de Google Sheets

'* Entradas:
'* @param  Company - Cadena que representa el nombre de la empresa
'* @param  sigad - Cadena que representa sigad
'* @param  orden - Entero que representa el pedido
'* @param  patience - Entero que representa la paciencia
'* @param  libro - Un valor que representa el libro de Excel
'* @param  bookNow - Cadena que representa la hoja de trabajo
Public Sub AddRecordToGoogleSheet(ByVal Company As String, ByVal sigad As String, ByVal orden As Integer, ByVal patience As Integer, ByVal libro As Variant, ByVal bookNow As String)

  '' ya funciona usa el token oAuth2 ''

  Dim HttpReq As Variant
  Dim Json As Object
  Dim monthNow As Integer, yearNow As Integer
  Dim fullDate As String, dateNow As String, bearerToken As String

  fullDate = Format(Now, "dd/mm/yyyy hh:mm:ss")
  dateNow = Format(Date, "dd-mmm-yyyy")
  monthNow = Month(Date)
  yearNow = Year(Date)
  bearerToken = Application.InputBox(prompt:="ingrese el Token de Acceso", Title:="Acceso Google Sheet", Default:="", Type:=2)

  Set HttpReq = CreateObject("MSXML2.XMLHTTP")
  HttpReq.Open "POST", "https://sheets.googleapis.com/v4/spreadsheets/126vzNrB3mA-g-61ccgNyAz-ukhIIqg_Yn3JxzQljC5o/values/Registro!$A2:append?valueInputOption=RAW", False
  HttpReq.setRequestHeader "Authorization", "Bearer " & Trim(bearerToken)
  HttpReq.setRequestHeader "Content-Type", "application/json"

  Dim requestBody As String
  requestBody = "{""values"":[['" & fullDate & "','" & dateNow & "','" & Company & "','" & sigad & "'," & orden & "," & patience & "," & monthNow & "," & yearNow & "]]}"

  On Error Resume Next
  HttpReq.send (requestBody)

  If HttpReq.status = 200 Then
    MsgBox "Record added successfully:" + vbNewLine + vbNewLine + Chr(32) + "code:" & HttpReq.status & "" + vbNewLine + Chr(32) + "status:" & HttpReq.statusText
  ElseIf HttpReq.status = 12031 Then
    MsgBox "Restriction by network administrator:" + vbNewLine + vbNewLine + Chr(32) + "code:" & HttpReq.status
    Workbooks.Open (libro)
    Windows(bookNow).Activate
  Else
    MsgBox "Error adding record: " & HttpReq.status & vbNewLine & HttpReq.statusText & vbNewLine & HttpReq.responseText
    Workbooks.Open (libro)
    Windows(bookNow).Activate
  End If
End Sub

'TODO: Actualiza un registro en una hoja de calculo de Google
'* @param rowData: fila donde se encuentra el registro a actualizar
'* @param textModify: nuevo valor del registro a actualizar
Public Sub UpdateGoogleSheetRecord(ByVal rowData As Integer, ByVal textModify As String)

  Dim sheetId As String, accessToken As String
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

'''? <summary>
'''? Exporta la informacion de cada tabla en un archivo SQL.
'''? </summary>
Public Sub ExportSQL()

  Dim origin As Workbook
  Dim comple_origin As Worksheet, worker_origin As Worksheet, osteo_origin As Worksheet, senso_origin As Worksheet, psico_origin As Worksheet, visio_origin As Worksheet, espiro_origin As Worksheet, opto_origin As Worksheet, audio_origin As Worksheet, emo_origin As Worksheet
  Dim sh As Variant, str As Variant, MyFile As Variant
  Dim num As Integer
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")

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

  Set MyFile = fso.OpenTextFile(ThisWorkbook.Worksheets("RUTAS").range("C9").value & "testfile.sql", ForAppending, True, TristateTrue)
  For Each sh In origin.Worksheets
    If ActiveWorkbook.Sheets(sh.Name).Tab.ThemeColor = xlThemeColorAccent1 Then
      Select Case Trim(UCase(sh.Name))
       Case "TRABAJADORES"
        '' orden lista trabajadores
        num = isEmptyValue(range("tbl_trabajadores[[SCRIPT orden_lista_trabajadores]]"))
        If (num > 0) Then
          MyFile.WriteLine "INSERT INTO orden_lista_trabajadores (`id`, `id_orden`, `estado`, `cedula`, `nombre`, `telefono`, `registro`, `ciudad_id`, `empresa_id`, `digitador_id`, `fecha_ingreso`, `id_cargo`, `fuente`, `edad`, `genero`, `estrato`, `id_raza`, `id_estado_civil`, `hijos`, `id_escolaridad`, `rango_edad`, `duracion`, `antiguedad`, `created_at`, `updated_at`, `id_tipo_actividad`, `id_tipo_examen`) VALUES"
          For Each Item In range("tbl_trabajadores[[SCRIPT orden_lista_trabajadores]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' paraclinicos
        num = isEmptyValue(range("tbl_trabajadores[[SCRIPT ordenes_trabajador_paraclinicos]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ordenes_trabajador_paraclinicos (`id_orden_trabajador`, `id_paraclinico`, `estado`) VALUES"
          For Each Item In range("tbl_trabajadores[[SCRIPT ordenes_trabajador_paraclinicos]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If
       Case "EMO"
        '' ics_emo
        num = isEmptyValue(range("tbl_emo[[SCRIPT ics_emo]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ics_emo (`id`, `id_orden_lista_trabajadores`, `id_concepto_evaluacion`, `observaciones`, `accidente_laboral`, `enfermedad_laboral`, `fecha_accidente`, `empresa`, `naturaleza_lesion`, `tipo_accidente`, `parte_afectada`, `dias_incapacidad`, `secuelas`, `enfermedad`, `etapa`, `observaciones_enfermedad`, `actividad_fisica`, `fuma`, `consumo_alcohol`, `peso`, `talla`, `tension_arterial`, `frecuencia_cardiaca`, `perimetro_abominal`, `lateralidad`, `frecuencia_respiratoria`, `imc2`, `clasificacion_imc`, `observacion_recomendacion`, `observacion_diagnostico`) VALUES"
          For Each Item In range("tbl_emo[[SCRIPT ics_emo]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' ics_emo_riesgos
        num = isEmptyValue(range("tbl_emo[[SCRIPT ics_emo_riesgos]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ics_emo_riesgos (`id_ics`, `id_riesgo`, `observaciones_otros`) VALUES"
          For Each Item In range("tbl_emo[[SCRIPT ics_emo_riesgos]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' ics_condiciones
        num = isEmptyValue(range("tbl_emo[[SCRIPT ics_condiciones]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ics_condiciones (`id_ics`, `id_condicion`, `condicion_seguridad`) VALUES"
          For Each Item In range("tbl_emo[[SCRIPT ics_condiciones]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' ics_cie (diagnosticos)
        num = isEmptyValue(range("tbl_emo[[script ics_cie (diagnosticos)]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ics_cie (`id_ics`, `id_cie`) VALUES"
          For Each Item In range("tbl_emo[[script ics_cie (diagnosticos)]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' ics_enfasis
        num = isEmptyValue(range("tbl_emo[[SCRIPT ics_enfasis]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO ics_enfasis (`id_ics`, `id_enfasis`, `observacion`) VALUES"
          For Each Item In range("tbl_emo[[SCRIPT ics_enfasis]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

       Case "AUDIO"
        '' au_audiometria
        num = isEmptyValue(range("tbl_audio[[SCRIPT au_audiometria]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO au_audiometria (`id`, `emo_id`, `auditivo`, `auditivo_copa`, `auditivo_insercion`, `auditivo_doble`, `diagnostico_interno`, `diagnostico_ppal`, `diagnostico_gati`, `status_obs`) VALUES"
          For Each Item In range("tbl_audio[[SCRIPT au_audiometria]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' au_audiometria_recomendacion
        num = isEmptyValue(range("tbl_audio[[SCRIPT au_audiometria_recomendacion]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO au_audiometria_recomendacion (`audiometria_id`, `recomendacion_id`, `fk_id_control`) VALUES"
          For Each Item In range("tbl_audio[[SCRIPT au_audiometria_recomendacion]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' au_oido
        num = isEmptyValue(range("tbl_audio[[SCRIPT au_oido]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO au_oido (`audiometria_id`, `tipo_oido_id`, `pabellon_id`, `auditivo_id`, `membrana_id`, `obs_pabellon`, `obs_auditivo`, `obs_membrana`, `frecuencia`, `pta`) VALUES"
          For Each Item In range("tbl_audio[[SCRIPT au_oido]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

       Case "OPTO"
        '' op_optometria
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_optometria]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_optometria (`id`, `emo_id`, `parpados`, `obs_parpados`, `conjuntivas`, `obs_conjuntivas`, `escleras`, `obs_escleras`, `pupilas`, `obs_pupilas`, `lejos`, `cerca`, `patologia_ocular`, `estado_correcion_id`, `otros_sintomas`, `recomendacion`, `remision`, `status_dig`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_optometria]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' op_optometria_riesgos
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_optometria_riesgos]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_optometria_riesgos (`optometria_id`, `riesgo_id`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_optometria_riesgos]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' op_optometria_sintomas
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_optometria_sintomas]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_optometria_sintomas (`optometria_id`, `sintomas_id`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_optometria_sintomas]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' op_diagnostico
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_diagnostico]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_diagnostico (`id`, `optometria_id`, `diagnostico_ppal`, `obs_ppal`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_diagnostico]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' op_diagnostico_cie
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_diagnostico_cie]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_diagnostico_cie (`diagnostico_id`, `cie_id`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_diagnostico_cie]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' op_optometria_recomendacion
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_optometria_recomendacion]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_optometria_recomendacion (`optometria_id`, `recomendacion_id`, `fk_control_id`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_optometria_recomendacion]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' op_optometria_remision
        num = isEmptyValue(range("tbl_opto[[SCRIPT op_optometria_remision]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO op_optometria_remision (`optometria_id`, `remision_id`) VALUES"
          For Each Item In range("tbl_opto[[SCRIPT op_optometria_remision]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

       Case "VISIO"
        '' vi_visiometria
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_visiometria]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_visiometria (`id`, `emo_id`, `parpados`, `obs_parpados`, `conjuntivas`, `obs_conjuntivas`, `escleras`, `obs_escleras`, `pupilas`, `obs_pupilas`, `otros_sintomas`, `resultado`, `obs_resultado`, `recomendacion_general`, `remision`, `status_general`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_visiometria]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' vi_visiometria_antecedentes
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_visiometria_antecedentes]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_visiometria_antecedentes (`visiometria_id`, `antecedente_id`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_visiometria_antecedentes]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' vi_visiometria_sintomas
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_visiometria_sintomas]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_visiometria_sintomas (`visiometria_id`, `sintoma_id`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_visiometria_sintomas]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' vi_vl
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_vl]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_vl (`visiometria_id`, `oi`, `od`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_vl]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' vi_vp
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_vp]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_vp (`visiometria_id`, `oi`, `od`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_vp]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' vi_visiometria_recomendaciones
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_visiometria_recomendaciones]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_visiometria_recomendaciones (`visiometria_id`, `recomendacion_id`, `fk_id_control`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_visiometria_recomendaciones]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' vi_visiometria_remisiones
        num = isEmptyValue(range("tbl_visio[[SCRIPT vi_visiometria_remisiones]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO vi_visiometria_remisiones (`visiometria_id`, `remision_id`) VALUES"
          For Each Item In range("tbl_visio[[SCRIPT vi_visiometria_remisiones]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

       Case "ESPIRO"
        '' espirometria
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espirometria]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espirometria (`id`, `emo_id`, `observaciones_alergias`, `observaciones_cx_torax`, `observaciones_cancer`, `otros_respiratorios`, `otros_riesgos_quimicos`, `actividad_fisica`, `fuma`, `frecuencia_habito`, `numero_cigarrros`, `tiempo_anios`, `interpretaciones`, `tecnica_aceptable`, `calculos_diagnostico`, `diagnostico_ppal`, `observacion_ppal`, `tipo_interpretacion`, `tipo_grado`, `resultado_espiro`, `peso`, `talla`, `imc2`, `clasificacion_imc`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espirometria]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' espiro_antecedentes_pivot
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espiro_antecedentes_pivot]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espiro_antecedentes_pivot (`espiro_id`, `id_antecedente`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espiro_antecedentes_pivot]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' espiro_quimicos_pivot
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espiro_quimicos_pivot]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espiro_quimicos_pivot (`espiro_id`, `id_quimicos`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espiro_quimicos_pivot]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' espiro_riesgos_epp
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espiro_riesgos_epp]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espiro_riesgos_epp (`espiro_id`, `tapaboca`, `especifico`, `otro_tapaboca`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espiro_riesgos_epp]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' espiro_recomendaciones_pivot
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espiro_recomendaciones_pivot]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espiro_recomendaciones_pivot(`espiro_id`, `recomendaciones_id`, `fk_id_control`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espiro_recomendaciones_pivot]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' espiro_recomendaciones_lab_pivot
        num = isEmptyValue(range("tbl_espiro_info[[SCRIPT espiro_recomendaciones_lab_pivot]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO espiro_recomendaciones_lab_pivot (`espiro_id`, `recomendaciones_id`) VALUES"
          For Each Item In range("tbl_espiro_info[[SCRIPT espiro_recomendaciones_lab_pivot]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

       Case "OSTEO"
        '' osteomuscular
        num = isEmptyValue(range("tbl_osteo[[SCRIPT osteomuscular]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO osteomuscular (`id`, `emo_id`,`diagnostico_ppal`, `observacion_ppal`, `ocupacionales`, `generales`, `status_ppal`) VALUES"
          For Each Item In range("tbl_osteo[[SCRIPT osteomuscular]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' osteo_antecedentes_pivot
        num = isEmptyValue(range("tbl_osteo[[SCRIPT osteo_antecedentes_pivot]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO osteo_antecedentes_pivot (`osteo_id`, `id_antecedente_osteo`, `observacion_antecedente_sintoma`) VALUES"
          For Each Item In range("tbl_osteo[[SCRIPT osteo_antecedentes_pivot]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' osteo_cie_pivot
        num = isEmptyValue(range("tbl_osteo[[SCRIPT osteo_cie_pivot]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO osteo_cie_pivot (`osteo_id`, `cie_id`) VALUES"
          For Each Item In range("tbl_osteo[[SCRIPT osteo_cie_pivot]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' osteo_recomendaciones_pivot
        num = isEmptyValue(range("tbl_osteo[[SCRIPT osteo_recomendaciones_pivot]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO osteo_recomendaciones_pivot (`osteo_id`, `id_recomendaciones_osteo`) VALUES"
          For Each Item In range("tbl_osteo[[SCRIPT osteo_recomendaciones_pivot]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

       Case "COMPLEMENTARIOS", "COMPLEMENTARIO"
        '' complementarios
        num = isEmptyValue(range("tbl_complementarios[[SCRIPT complementarios]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO complementarios (`id`, `emo_id`, `procedimiento_id`, `diagnostico_ppal`, `observacion_ppal`, `hallazgo`, `status_ppal`) VALUES"
          For Each Item In range("tbl_complementarios[[SCRIPT complementarios]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' complementarios_diagnos_observaciones_pivot
        num = isEmptyValue(range("tbl_complementarios[[SCRIPT complementarios_diagnos_observaciones_pivot]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO complementarios_diagnos_observaciones_pivot (`complementarios_id`, `diagnostico`) VALUES"
          For Each Item In range("tbl_complementarios[[SCRIPT complementarios_diagnos_observaciones_pivot]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

       Case "PSICOTECNICA", "PSICOLOGIA"
        '' psicotecnica
        num = isEmptyValue(range("tbl_psicotecnica[[SCRIPT psicotecnica]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO psicotecnica (`id`, `emo_id`, `prueba`, `id_diagnostico_ppal`, `observacion_ppal`) VALUES"
          For Each Item In range("tbl_psicotecnica[[SCRIPT psicotecnica]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

       Case "PSICOSENSOMETRICA", "PSICOMOTRIZ"
        '' psicosensometrica
        num = isEmptyValue(range("tbl_psicosensometrica[[SCRIPT psicosensometrica]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO psicosensometrica (`id`, `emo_id`, `prueba`, `id_diagnostico_ppal`, `observacion_ppal`) VALUES"
          For Each Item In range("tbl_psicosensometrica[[SCRIPT psicosensometrica]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' psicosenso_diagnos_observaciones_pivot
        num = isEmptyValue(range("tbl_psicosensometrica[[SCRIPT psicosenso_diagnos_observaciones_pivot]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO psicosenso_diagnos_observaciones_pivot (`psicosensometrica_id`, `diagnostico`) VALUES"
          For Each Item In range("tbl_psicosensometrica[[SCRIPT psicosenso_diagnos_observaciones_pivot]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If

        '' psicosensometricas_recomendaciones_pivot
        num = isEmptyValue(range("tbl_psicosensometrica[[SCRIPT psicosensometricas_recomendaciones_pivot]]"))
        If (num > 0) Then
          MyFile.WriteLine ""
          MyFile.WriteLine "INSERT INTO psicosensometricas_recomendaciones_pivot (`psicosensometrica_id`, `recomendaciones_id`) VALUES"
          For Each Item In range("tbl_psicosensometrica[[SCRIPT psicosensometricas_recomendaciones_pivot]]")
            If Item <> "" And num <> 1 Then
              MyFile.WriteLine Item
              num = num - 1
            ElseIf Item <> "" And num = 1 Then
              MyFile.WriteLine reemplazarUltimoCaracter(Item)
              num = num - 1
            End If
          Next Item
        End If
      End Select
    End If
  Next sh
  MyFile.Close

  MsgBox "Se genero el archivo SQL textfile.sql" + vbNewLine + vbNewLine + Chr(32) + "Que se encuentra en la ruta: " + vbNewLine + vbNewLine + ThisWorkbook.Worksheets("RUTAS").range("C9").value
End Sub

Private Function reemplazarUltimoCaracter(ByVal texto As String) As String
  Dim ultimoCaracter As String

  ultimoCaracter = ","

  Dim posicion As Integer
  posicion = InStrRev(texto, ultimoCaracter)

  If posicion > 0 Then
    reemplazarUltimoCaracter = Left(texto, posicion - 1) & ";" & Right(texto, Len(texto) - posicion)
  End If
End Function

'TODO Funcion: isEmptyValue
'* Descripcion: Esta funcion toma un objeto Ranges como entrada y devuelve el numero de elementos que no estan vacios en el objeto.
'* Parametros:
'* @param Ranges: Objeto que contiene los elementos a verificar.
'* Valor de retorno:
'* @return Un entero que representa el numero de elementos que no estan vacios en Ranges.
Public Function isEmptyValue(ByVal Ranges As Object) As Integer
  Dim num As Integer
  Dim Item As Variant

  num = 0
  For Each Item In Ranges
    If (Item <> "") Then
      num = num + 1
    End If
  Next Item
  isEmptyValue = num

End Function

'TODO: Esta funcion convierte las fechas en un formato especifico a otro formato "dd/mm/yyyy".
'* Se recorre un rango de celdas hasta encontrar una celda vacia y se intenta convertir la fecha en esa celda al formato deseado.
'* Si la conversion es exitosa, se reemplaza el valor de la celda con el nuevo formato.
'* Si la conversion falla, se establece la celda en vacio.
'* Si se produce un error, se intenta convertir la fecha utilizando un formato diferente.
Public Sub DateAccident()
  Dim dateRange As range
  Dim dateCell As range
  Dim toDate As Date
  Dim dateString As String
  Dim arrayDateStringToNumber As Variant
  Dim element As Variant
  
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.StatusBar = "Procesando..."
  
  Set dateRange = range("$BC5")
  Set dateCell = dateRange
  
  arrayDateStringToNumber = Array("Jan,01", "Feb,02", "Mar,03", "Apr,04", "May,05", "Jun,06", "Jul,07", "Aug,08", "Sep,09", "Oct,10", "Nov,11", "Dec,12")
  
  Do Until IsEmpty(dateCell.Offset(0, -54))
    On Error Resume Next
    dateString = Trim$(dateCell.value)
    dateString = Replace(Replace(dateString, " ", "/"), "  ", " ")
    On Error GoTo 0
    
    For Each element In arrayDateStringToNumber
      If dateString = vbNullString Then
        Exit For
      ElseIf IsDate(dateString) Then
        toDate = dateString
        Exit For
      ElseIf UCase(Trim$(Split(dateString, "/")(0))) = UCase(Split(element, ",")(0)) Then
        toDate = Split(dateString, "/")(1) & "/" & Split(element, ",")(1) & "/" & Split(dateString, "/")(2)
        Exit For
      End If
    Next element
    
    If dateString <> vbNullString And IsDate(toDate) Then
      dateCell.value = Format(CDate(toDate), "dd/mm/yyyy")
      dateCell.NumberFormat = "dd/mm/yyyy"
    Else
      dateCell.value = ""
    End If
    
    Set dateCell = dateCell.Offset(1, 0)
    DoEvents
  Loop
  
  On Error Resume Next
  dateRange.Resize(dateCell.row - dateRange.row).TextToColumns Destination:=range("BC5"), DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 4), Array(20, 9)), TrailingMinusNumbers:=True
  On Error GoTo 0
  
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.StatusBar = ""

End Sub
