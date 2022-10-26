Attribute VB_Name = "Initial"
Option Explicit

'Variables
Public origin, destiny As Workbook
Public comple_origin, comple_destiny, osteo_origin, osteo_destiny, senso_destiny, senso_origin, psico_destiny, psico_origin, visio_destiny, visio_origin, espiro_destiny, espiro_origin, opto_origin, opto_destiny, audio_origin, audio_destiny, worker_destiny, emo_destiny, emo_origin, emphasis_destiny, diagnostics_destiny As Worksheet
Public route, nameCompany As String
Public variable, insertVisio, insertOpto, insertAudio, insertOsteo, insertSenso, inserEspiro, insertComple, insertPsico, insertEmo, dataInsert, ItemTitle, titulos, DatosOsteo, DatosSenso, DatosPsico, DatosComple, DatosOpto, DatosAudio, DatosEmo, DatosEspiro, DatosVisio As Object
Public ordenListaTrabajador As Long
Public Item As Variant
Public vals, valsGeneral, porcentaje, porcentajeGeneral As Double
Public idOrden, numbers, numbersGeneral, sumOneforOne, sumGeneral, x, i, number_emphasis, number_diag As Integer
Public dateInitials, dateFinals As Date
Public counts, totalData, generalAll, widthOneforOne, widthGeneral, oneForOne As Double

Sub extraerdatos()

  numbers = 1
  numbersGeneral = 1
  porcentaje = 0
  porcentajeGeneral = 0
  totalData = 0
  dateInitials = VBA.Date

  'route = ThisWorkbook.Worksheets("RUTAS").Range("C4").value

  '''''''''''''''''''''''''''''''''''''''''''''''''
  '''        APERTURA DEL LIBRO ARCHIVO         '''
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

  Application.StatusBar = "Importando informaci" & Chr(243) & "n por favor espere"
  Application.ScreenUpdating = False
  Application.Calculation = False
  Application.EnableEvents = False
  totalData = total(origin)
  For Each variable In origin.Worksheets
    Select Case Trim(UCase(variable.Name))
     Case "EMO"
      Call Workers
      Call DataEmoWorkers
      Call DataEmphasisEmo
      Call DataDiagnosticsEmo
     Case "AUDIO"
      Call AudioData
     Case "OPTO"
      Call OptoData
     Case "VISIO"
      Call VisioData
     Case "ESPIRO"
      Call EspiroData
     Case "OSTEO"
      Call OsteoData
     Case "COMPLEMENTARIO", "COMPLEMENTARIOS"
      Call ComplementarioData
     Case "PSICOTECNICA", "PSICOLOGIA"
      Call PsicotecnicaData
     Case "PSICOSENSOMETRICA", "PSICOMOTRIZ"
      Call PsicosensometricaData
    End Select
  Next variable

  origin.Save
  origin.Close

  Worksheets("TRABAJADORES").Select
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
  Application.StatusBar = Empty
  Unload formImports
  Dim hora, min As Integer

  hora = VBA.Hour(Time)
  min = VBA.Minute(Time)
  dateFinals = VBA.Date
  destiny.Save
  If (hora >= 17 And min >= 15) Or (dateInitials <> dateFinals) Then
    Call Shell("shutdown /s /t: 30 /f")
    destiny.Close
  Else
    MsgBox "Importe de informaci" & Chr(243) & "n terminado", vbInformation + vbOKOnly, "Importaci" & Chr(243) & "n Datos"
  End If

End Sub

Sub info()

  formImports.Show

End Sub

sub config()

  formControl.Show
  
End Sub
