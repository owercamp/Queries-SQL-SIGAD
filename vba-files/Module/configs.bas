Attribute VB_Name = "configs"
'namespace=vba-files\Module
Option Explicit

Public Sub init()
  
  Dim btnStatus As MsoButtonState, Sheet As Worksheet, th As Workbook
  
  Set th = ThisWorkbook
  
  btnStatus = MsgBox(prompt:="¿Desea realizar La importacion de la Hoja BASE P?", Buttons:=vbDefaultButton1 + vbYesNo + vbExclamation, _
  Title:="Consolidado Informaci" & ChrW(243) & "n")
  
  With Application
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
  End With
  
  If btnStatus = vbYes Then
    Dim book As Workbook, bookName As String
    
    bookName = Application.getOpenFilename
    
    Set book = Workbooks.Open(bookName)
    
    Windows(book.Name).Activate
    Sheets("BASE P").Select
    Sheets("BASE P").Copy Before:=Workbooks(th.Name).Sheets(1)
    book.Close
    th.Sheets(th.Worksheets.Count).Select
    Call routines(th)

  Else
    For Each Sheet In th.Worksheets
      If Sheet.Name = "BASE P" Then
        Call routines(th)
      End If
    Next Sheet
  End If
  
  With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
  End With
  
End Sub

Private Sub routines(ByVal th As Workbook)
  DoEvents
  Call addSheets
  th.Worksheets("RUTAS").Select
  Call configRoute
  th.Worksheets("DIAGNOSTICOS").Select
  Call configDiagnostics
  th.Worksheets("ENFASIS").Select
  Call configEmphasis
  th.Worksheets("TRABAJADORES").Select
  Call configWorkers
  th.Worksheets("EMO").Select
  Call configEmo
  th.Worksheets("AUDIO").Select
  Call configAudio
  th.Worksheets("VISIO").Select
  Call configVisio
  th.Worksheets("OPTO").Select
  Call configOpto
  th.Worksheets("ESPIRO").Select
  Call configEspiro
  th.Worksheets("OSTEO").Select
  Call configOsteo
  th.Worksheets("COMPLEMENTARIOS").Select
  Call configComple
  th.Worksheets("PSICOTECNICA").Select
  Call configPsico
  th.Worksheets("PSICOSENSOMETRICA").Select
  Call configSenso
  th.Worksheets("RUTAS").Visible = 2
  th.Worksheets("TRABAJADORES").Select
  Call btnCreate
  Call insertFunctions
End Sub
