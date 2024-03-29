Attribute VB_Name = "Initial"
'namespace=vba-files\Module
Option Explicit
'' Variables ''
Public origin As Workbook, destiny As Workbook
Public comple_destiny As Worksheet, osteo_destiny As Worksheet, senso_destiny As Worksheet, psico_destiny As Worksheet, visio_destiny As Worksheet, espiro_destiny As Worksheet, opto_destiny As Worksheet, audio_destiny As Worksheet, worker_destiny As Worksheet, emo_destiny As Worksheet, emphasis_destiny As Worksheet, diagnostics_destiny As Worksheet
Public route As String, nameCompany As String
Public variable As Object,ordenListaTrabajador As Long, item As Variant
Public vals As Double, valsGeneral As Double, porcentaje As Double, porcentajeGeneral As Double, counts As Double, totalData As Double, generalAll As Double, widthOneforOne As Double, widthGeneral As Double, oneForOne As Double
Public idOrden As LongPtr, numbers As LongPtr, numbersGeneral As LongPtr, sumOneforOne As LongPtr, sumGeneral As LongPtr, x As LongPtr, i As LongPtr, number_emphasis As LongPtr, number_diag As LongPtr
Public dateInitials As Date, dateFinals As Date,  withIncidence As Object, withoutIncidence As Object, search As String

Public Sub extraerdatos()

  Dim fso As Object
  Dim hora As Integer, min As Integer
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set withIncidence = CreateObject("Scripting.Dictionary")
  Set withoutIncidence = CreateObject("Scripting.Dictionary")
  
  '*Dictionary withIncidence and withoutIncidence

  withIncidence.CompareMode = vbTextCompare
  withIncidence.Add Trim$("OCASIONAL"),Trim$("OCASIONAL")
  withIncidence.Add Trim$("SI"),Trim$("SI")
  withIncidence.Add Trim$("1"),Trim$("1")
  
  withoutIncidence.CompareMode = vbTextCompare
  withoutIncidence.Add Trim$(""),Trim$("")
  withoutIncidence.Add Trim$("NO"),Trim$("NO")
  withoutIncidence.Add Trim$("0"),Trim$("0")

  numbers = 1
  numbersGeneral = 1
  porcentaje = 0
  porcentajeGeneral = 0
  totalData = 0
  dateInitials = VBA.Date

  If fso.FileExists(ThisWorkbook.Worksheets("RUTAS").range("C9").value & "/testfile.sql") Then
    fso.DeleteFile (ThisWorkbook.Worksheets("RUTAS").range("C9").value & "/testfile.sql")
  End If

  'route = ThisWorkbook.Worksheets("RUTAS").Range("C4").value

  '''''''''''''''''''''''''''''''''''''''''''''''''
  ''''        APERTURA DEL LIBRO ARCHIVO         ''
  '''''''''''''''''''''''''''''''''''''''''''''''''
  'Set origin = Workbooks.Open(route)


  'DATOS DESTINO

  Set destiny = Workbooks(ThisWorkbook.Name)
  Set worker_destiny = destiny.Worksheets("TRABAJADORES")
  Set emo_destiny = destiny.Worksheets("EMO")
  Set audio_destiny = destiny.Worksheets("AUDIO")
  Set opto_destiny = destiny.Worksheets("OPTO")
  Set espiro_destiny = destiny.Worksheets("ESPIRO")
  Set osteo_destiny = destiny.Worksheets("OSTEO")
  Set visio_destiny = destiny.Worksheets("VISIO")
  Set psico_destiny = destiny.Worksheets("PSICOTECNICA")
  Set senso_destiny = destiny.Worksheets("PSICOSENSOMETRICA")
  Set comple_destiny = destiny.Worksheets("COMPLEMENTARIOS")
  Set emphasis_destiny = destiny.Worksheets("ENFASIS")
  Set diagnostics_destiny = destiny.Worksheets("DIAGNOSTICOS")


  'DATOS ORIGEN

  With Application
    .StatusBar = "Importando informaci" & ChrW(243) & "n por favor espere"
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With
  totalData = total(origin)
  For Each variable In origin.Worksheets
    Select Case Trim(UCase(variable.Name))
     Case "EMO"
      If (variable.Visible = True) Then
        Call Workers(variable.Name)
        Call statusActivate(worker_destiny.Name)
        Call DataDiagnosticsEmo(variable.Name)
        Call DataEmphasisEmo(variable.Name)
        Call DataEmoWorkers(variable.Name)
        Call statusActivate(emo_destiny.Name)
      End If
     Case "AUDIO"
      If (variable.Visible = True) Then
        Call AudioData(variable.Name)
        Call statusActivate(audio_destiny.Name)
      End If
     Case "OPTO"
      If (variable.Visible = True) Then
        Call OptoData(variable.Name)
        Call statusActivate(opto_destiny.Name)
      End If
     Case "VISIO"
      If (variable.Visible = True) Then
        Call VisioData(variable.Name)
        Call statusActivate(visio_destiny.Name)
      End If
     Case "ESPIRO"
      If (variable.Visible = True) Then
        Call EspiroData(variable.Name)
        Call statusActivate(espiro_destiny.Name)
      End If
     Case "OSTEO"
      If (variable.Visible = True) Then
        Call OsteoData(variable.Name)
        Call statusActivate(osteo_destiny.Name)
      End If
     Case "COMPLEMENTARIO", "COMPLEMENTARIOS"
      If (variable.Visible = True) Then
        Call ComplementarioData(variable.Name)
        Call statusActivate(comple_destiny.Name)
      End If
     Case "PSICOTECNICA", "PSICOLOGIA", "PSICO"
      If (variable.Visible = True) Then
        Call PsicotecnicaData(variable.Name)
        Call statusActivate(psico_destiny.Name)
      End If
     Case "PSICOSENSOMETRICA", "PSICOMOTRIZ", "MOTRIZ"
      If (variable.Visible = True) Then
        Call PsicosensometricaData(variable.Name)
        Call statusActivate(senso_destiny.Name)
      End If
    End Select
  Next variable

  origin.Save
  origin.Close

  Windows(destiny.Name).Activate
  Worksheets("TRABAJADORES").Select
  With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
    .StatusBar = Empty
  End With
  Unload formImports

  hora = VBA.Hour(Time)
  min = VBA.Minute(Time)
  dateFinals = VBA.Date
  destiny.Save
  If (hora >= 17 And min >= 15) Or (dateInitials <> dateFinals) Then
    Call Shell("shutdown /s /t: 30 /f")
    destiny.Close
  Else
    MsgBox "Importe de informaci" & ChrW(243) & "n terminado", vbInformation + vbOKOnly, "Importaci" & ChrW(243) & "n Datos"
  End If

End Sub

Public Sub statusActivate(ByVal name_sheet As String)
  With Workbooks("Queries SQL SIGAD.xlsb").Sheets(name_sheet).Tab
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = -0.249977111117893
  End With
End Sub

Public Sub statusDesactivate(ByVal name_sheet As String)
  Sheets(name_sheet).Select
  With ActiveWorkbook.Sheets(name_sheet).Tab
    .Color = RGB(222, 222, 222)
    .TintAndShade = 0
  End With
End Sub

Public Sub info()
  formImports.Show
End Sub

Public Sub config()

  formControl.Show

End Sub

Public Sub cleanCaracthers()
Attribute cleanCaracthers.VB_ProcData.VB_Invoke_Func = "y\n14"
  formClear.Show
End Sub
