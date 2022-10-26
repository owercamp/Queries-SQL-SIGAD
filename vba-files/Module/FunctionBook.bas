Attribute VB_Name = "FunctionBook"
Option Explicit
Public sigad As Variant

Sub cargos()
Attribute cargos.VB_ProcData.VB_Invoke_Func = "k\n14"
  Workbooks.Open (ThisWorkbook.Worksheets("RUTAS").Range("C7").value)
End Sub

Sub folder(route, folderName, workbookActive)
  Dim splitRoute As String
  splitRoute = Application.PathSeparator

  If Dir(route, vbDirectory) = Empty Then
    MkDir route
  End If

  If Dir(route & splitRoute & folderName, vbDirectory) = Empty Then
    MkDir (route & splitRoute & folderName)
    Application.ActiveWorkbook.SaveCopyAs Filename:=route & splitRoute & folderName & splitRoute & workbookActive
    Application.StatusBar = "se guardo una copia en: " & route & splitRoute & folderName & splitRoute & workbookActive
  Else
    Application.ActiveWorkbook.SaveCopyAs Filename:=route & splitRoute & folderName & splitRoute & workbookActive
    Application.StatusBar = "se guardo una copia en: " & route & splitRoute & folderName & splitRoute & workbookActive
  End If
End Sub

Sub clearContents()

  Dim trabajadores, emo, audio, visio, opto, espiro, osteo, complementarios, psicotecnica, psicosensometrica, enfasis, diag As Worksheet
  Dim rng, info, rngTrabajadores, rngEmo, rngAudio, rngVisio, rngOpto, rngEspiro, rngOsteo, rngComplementarios, rngPsicotecnica, rngPsicosensometrica, rngEnfasis, rngDiag, MyDay, MyMonth, MyYear As Integer
  Dim meses, finalRow, RowActive As Variant
  Dim nombre, orden, fecha, company As String
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
  If trabajadores.Range("D6") <> Empty Or trabajadores.Range("D6") <> vbNullString Then: nombre = trabajadores.Range("B6").value & " - " & trabajadores.Range("D6").value & ".xlsb"
    If trabajadores.Range("D6") = Empty Or trabajadores.Range("D6") = vbNullString Then: nombre = trabajadores.Range("B6").value & ".xlsb"
      orden = trabajadores.Range("AX6").value

      If (Not IsEmpty(nombre)) And (Not IsEmpty(orden)) And (Not IsEmpty(sigad)) Then
        fecha = CStr(MyDay) + " " + CStr(meses(MyMonth - 1)) + " " + CStr(MyYear)

        route = CStr(Worksheets("RUTAS").Range("C6").value & "\" & MyYear & "\" & meses(MyMonth - 1))

        trabajadores.Select
        Call folder(route, fecha, nombre)

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False


        '' REGISTRO EN CONSOLIDADO ''
        info = Worksheets("TRABAJADORES").Range("A6", Worksheets("TRABAJADORES").Range("A6").End(xlDown)).Count

        libro = Worksheets("RUTAS").Range("C5").value
        If trabajadores.Range("D6").value = Empty Or trabajadores.Range("D6").value = vbNullString Then
          company = trabajadores.Range("B6").value
        Else
          company = trabajadores.Range("B6").value & " - " & trabajadores.Range("D6").value
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
          ThisWorkbook.Worksheets("RUTAS").Range("$F$4") = Trim(Range("$AW$5").value)
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
          ThisWorkbook.Worksheets("RUTAS").Range("$F$5") = Trim(Range("$EL$5").value)
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
          ThisWorkbook.Worksheets("RUTAS").Range("$F$6") = Trim(Range("$BG$4").value)
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
          ThisWorkbook.Worksheets("RUTAS").Range("$F$7") = Trim(Range("$BL$4").value)
          ThisWorkbook.Worksheets("RUTAS").Range("$F$8") = Trim(Range("$BM$4").value)
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
          ThisWorkbook.Worksheets("RUTAS").Range("$F$9") = Trim(Range("$BS$4").value)
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
          ThisWorkbook.Worksheets("RUTAS").Range("$F$10") = Trim(Range("$BZ$4").value)
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
          ThisWorkbook.Worksheets("RUTAS").Range("$F$11") = Trim(Range("$BG$4").value)
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
          ThisWorkbook.Worksheets("RUTAS").Range("$F$12") = Trim(Range("$J$4").value)
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
          ThisWorkbook.Worksheets("RUTAS").Range("$F$13") = Trim(Range("$G$2").value)
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
          ThisWorkbook.Worksheets("RUTAS").Range("$F$14") = Trim(Range("$Q$3").value)
        End If

        trabajadores.Select
        Range("A5").Select
        Application.ActiveWorkbook.SaveCopyAs Filename:=route & "\" & Application.ActiveWorkbook.Name
        Application.StatusBar = Empty
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True

        MsgBox "Almacenamiento terminado", vbOKOnly + vbInformation, "Almacenamiento"
      Else

        MsgBox "No hay datos para almacenar", vbOKOnly + vbInformation, "Almacenamiento"

      End If
End Sub
