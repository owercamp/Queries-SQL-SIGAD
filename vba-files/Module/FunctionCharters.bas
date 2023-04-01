Attribute VB_Name = "FunctionCharters"
Option Explicit

'/*
' ELIMINA LOS ESPACIOS AL INICIO Y AL FINAL DE CADA VALOR
'*/
Public Function charters(ByVal value As String) As String
  charters = Trim(UCase(value))
End Function

'/*
' ELIMINA LOS ESPACIOS AL INICIO Y AL FINAL DE CADA VALOR Y VERIFICA QUE NO SEA UN CAMPO VACIO
'*/
Public Function charters_empty(value)
  Select Case Trim(UCase(value))
   Case IsEmpty(Trim(UCase(value))), "", "NO"
    charters_empty = "0"
   Case "OCASIONAL", "SI"
    charters_empty = "1"
   Case Else
    charters_empty = Trim(UCase(value))
  End Select
End Function

Public Function PadLeft(text As Variant, totalLength As Integer, padCharacter As String) As String
  PadLeft = String(totalLength - Len(CStr(text)), padCharacter) & CStr(text)
End Function

Public Function PadRight(text As Variant, totalLength As Integer, padCharacter As String) As String
  PadRight = CStr(text) & String(totalLength - Len(CStr(text)), padCharacter)
End Function


'/*
' VALIDACION CIUDAD
'*/
Public Function city(ByVal value As String) As String
  Select Case value
   Case "BOGOTA", "BOGOTA, D.C.", "BOGOT" & Chr(193) & ", D.C.", "BOGOTA, D.C", "BOGOTA D.C","BOGOT"& Chr(193), "BOGOTA  D.C","BOGOTA, BOGOTA D.C"
    city = Trim("BOGOTA D.C.")
   Case "CARTAGENA DE INDIAS","CARTAGENA, BOLIVAR"
    city = Trim("CARTAGENA")
   Case "BUGA"
    city = Trim("GUADALAJARA DE BUGA")
   Case "MONTEL" & Chr(205) & "BANO"
    city = Trim("MONTELIBANO")
   Case "PUERTO GAIT" & Chr(193) & "N"
    city = Trim("PUERTO GAITAN")
   Case "PUERTO BOYAC" & Chr(193)
    city = Trim("PUERTO BOYACA")
   Case "PUERTO AS" & Chr(205) & "S"
    city = Trim("PUERTO ASIS")
   Case "TULU"&Chr(193)
    city = Trim("TULUA")
   Case "POPAY"&Chr(193)&"N"
    city =Trim("POPAYAN")
   Case "SAN JOSE DE GUAVIARE"
    city = Trim("SAN JOSE DEL GUAVIARE")
   Case "MANIZALEZ"
    city = Trim("MANIZALES")
   Case "QUIBD" & Chr(211)
    city = Trim("QUIBDO")
   Case "UBATE"
    city = Trim("VILLA DE SAN DIEGO DE UBATE")
   Case "CHIQUINQUIR" & Chr(193)
    city = Trim("CHIQUINQUIRA")
   Case "FACATATIV" & Chr(193)
    city = Trim("FACATATIVA")
   Case "BUCARAMANGA, SANTANDER"
    city = "BUCARAMANGA"
   Case "VILLAVICENCIO, META"
    city = "VILLAVICENCIO"
   Case "IBAGUE, TOLIMA"
    city = "IBAGUE"
   Case "BARRANQUILA"
    city = "BARRANQUILLA"
   Case "CALI, VALLE DEL CAUCA"
    city = "CALI"
   Case "MEDELLIN, ANTIOQUIA"
    city = "MEDELLIN"
   Case "TUMACO"
    city = "SAN ANDRES DE TUMACO"
   Case Else
    city = value
  End Select
End Function

'/*
' VALIDACION ESCOLARIDAD
'*/
Public Function school(ByVal value As String) As String
  Select Case value
   Case "POSTGRADO","POST GRADO"
    school = "POSGRADO"
   Case "PROFESIONAL"
    school = "UNIVERSITARIO"
   Case "BACHILLER"
    school = "SECUNDARIA"
   Case "MAGISTER"
    school = "MAESTRIA"
   Case Else
    school = value
  End Select

End Function

'/*
' VALIDACION EXAMEN MEDICO
'*/
Public Function typeExams(ByVal value As String) As String
  Select Case value
   Case "POST INCAPACIDAD","POST-INCAPACIDAD"
    typeExams = "POS INCAPACIDAD"
   Case "PERIODICO SEG"
    typeExams = "PERIODICO"
   Case "PERIODICO SEGUIMIENTO","PERIODICO CON RECOMENDACIONES","PERIODICO CON SEGUIMIENTO"
    typeExams = "PERIODICO DE SEGUIMIENTO"
   Case "CAMBIO OCUPACION", "CAMBIO DE OCUPACI" & Chr(211) & "N"
    typeExams = "CAMBIO DE OCUPACION"
   Case "REINTEGRO LABORAL", "OTROS REINTEGROS"
    typeExams = "EGRESO"
   Case "PRE-INGRESO", "PRE_INGRESO", "INGRESO"
    typeExams = "PRE-INGRESO"
   Case Else
    typeExams = value
  End Select
End Function

'/*
' VALIDACION RAZA
'*/
Public Function typeSex(ByVal value As String) As String
  Select Case value
   Case "COBRIZA", "COBRIZO"
    typeSex = Trim("COBRIZA")
   Case "NEGRA", "NEGRO"
    typeSex = Trim("NEGRA")
   Case "OTRO", "OTRA"
    typeSex = Trim("OTRO")
   Case "BLANCA", "CAUCASICA", "BLANCO", "CAUCASICO"
    typeSex = Trim("CAUCASICA")
   Case "MULATA", "MULATO"
    typeSex = Trim("MULATO")
   Case "MESTIZO", "MESTIZA"
    typeSex = Trim("MESTIZO")
   Case "SIN DATO", "SIN DATOS"
    typeSex = Trim("SIN DATO")
   Case "IND" & Chr(205) & "GENA"
    typeSex = Trim("INDIGENA")
   Case Else
    typeSex = value
  End Select
End Function

'/*
' VALIDACION ESTADO CIVIL
'*/
Public Function typeCivil(ByVal value As String) As String
  Select Case value
   Case "UNI" & Chr(211) & "N LIBRE"
    typeCivil = "UNION LIBRE"
   Case Else
    typeCivil = value
  End Select
End Function

'/*
' VALIDACION ACTIVIDAD FISICA
'*/
Public Function typeActivity(ByVal value As String) As String
  Select Case value
   Case "F" & Chr(205) & "SICAMENTE ACTIVO", "FISICAMENTE ACTIVO", "FISICAMENTE ACTIVO(A)", "F" & Chr(205) & "SICAMENTE ACTIVO(A)"
    typeActivity = "F" & Chr(205) & "SICAMENTE ACTIVO"
   Case Else
    typeActivity = value
  End Select
End Function

'/*
' VALIDACION FUMADOR O EXFUMADOR
'*/
Public Function typeSmoke(ByVal value As String) As String
  Select Case value
   Case "EX-FUMADOR", "EXFUMADOR"
    typeSmoke = "EXFUMADOR"
   Case "SI"
    typeSmoke = "FUMADOR"
   Case "NO"
    typeSmoke = "NO FUMA"
   Case Else
    typeSmoke = value
  End Select
End Function

'/*
' VALIDACION CORRECTION OPTO
'*/
Public Function correction(ByVal value As String) As String
  Select Case value
   Case "ANORMAL SIN CORRECCION"
    correction = "ANORMAL MAL CORREGIDO"
   Case Else
    correction = value
  End Select
End Function

Public Function typeComplements(ByVal value As String) As String
  Select Case value
   Case "ENCUESTA RESPIRATORIA","ENCUESTA DE SINTOMAS RESPIRATORIOS"
    typeComplements = "VALORACION RESPIRATORIA"
   Case Else
    typeComplements = value
  End Select

End Function

'/*
' REALIZA EL CONTEO TOTAL DE DATOS A IMPORTAR
'*/
Function total(ByVal book As Object) As Integer

  Dim emo, audio, opto, espiro, visio, complementarios, psicotecnica, psicosensometrica, osteo As Integer
  Dim Sheet As Object

  For Each Sheet In book.Worksheets

    Select Case Trim(UCase(Sheet.Name))
     Case "EMO"
      If Sheet.Range("A2") <> "" And Sheet.Range("A3") <> "" Then
        nameCompany = Sheet.Range("A2").value
        formImports.Caption = CStr(nameCompany)
        emo = Sheet.Range("A2", Sheet.Range("A2").End(xlDown)).Count
      Else
        emo = 1
      End If
     Case "AUDIO"
      If Sheet.Range("A2") <> "" And Sheet.Range("A3") <> "" Then
        audio = Sheet.Range("A2", Sheet.Range("A2").End(xlDown)).Count
      Else
        audio = 1
      End If
     Case "OPTO"
      If Sheet.Range("A2") <> "" And Sheet.Range("A3") <> "" Then
        opto = Sheet.Range("A2", Sheet.Range("A2").End(xlDown)).Count
      Else
        opto = 1
      End If
     Case "VISIO"
      If Sheet.Range("A2") <> "" And Sheet.Range("A3") <> "" Then
        visio = Sheet.Range("A2", Sheet.Range("A2").End(xlDown)).Count
      Else
        visio = 1
      End If
     Case "ESPIRO"
      If Sheet.Range("A2") <> "" And Sheet.Range("A3") <> "" Then
        espiro = Sheet.Range("A2", Sheet.Range("A2").End(xlDown)).Count
      Else
        espiro = 1
      End If
     Case "OSTEO"
      If Sheet.Range("A2") <> "" And Sheet.Range("A3") <> "" Then
        osteo = Sheet.Range("A2", Sheet.Range("A2").End(xlDown)).Count
      Else
        osteo = 1
      End If
     Case "COMPLEMENTARIO", "COMPLEMENTARIOS"
      If Sheet.Range("A2") <> "" And Sheet.Range("A3") <> "" Then
        complementarios = Sheet.Range("A2", Sheet.Range("A2").End(xlDown)).Count
      Else
        complementarios = 1
      End If
     Case "PSICOTECNICA", "PSICOLOGIA"
      If Sheet.Range("A2") <> "" And Sheet.Range("A3") <> "" Then
        psicotecnica = Sheet.Range("A2", Sheet.Range("A2").End(xlDown)).Count
      Else
        psicotecnica = 1
      End If
     Case "PSICOSENSOMETRICA", "PSICOMOTRIZ"
      If Sheet.Range("A2") <> "" And Sheet.Range("A3") <> "" Then
        psicosensometrica = Sheet.Range("A2", Sheet.Range("A2").End(xlDown)).Count
      Else
        psicosensometrica = 1
      End If
    End Select
  Next Sheet

  total = (emo * 4) + audio + visio + espiro + osteo + complementarios + psicotecnica + psicosensometrica + opto

End Function

Sub ClearCharter()
  Attribute ClearCharter.VB_ProcData.VB_Invoke_Func = "y\n14"

  Dim data As Variant

  data = Array(Chr(193), Chr(192), Chr(200), Chr(201), Chr(204), Chr(205), Chr(210), Chr(211), Chr(217), Chr(218), Chr(44), Chr(95), Chr(147), Chr(13), Chr(10), Chr(160) & Chr(160), Chr(92), Chr(47), Chr(45), Chr(46))

  ' Doble espaciado
  Selection.Replace What:=data(15), Replacement:=" ", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  Selection.Replace What:="  ", Replacement:=" ", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  If (ActiveSheet.Name = "COMPLEMENTARIOS" And Selection.Address = Range("tbl_complementarios[PROCEDIMIENTO]").Address) Then
    ' guion al medio
    Selection.Replace What:=data(18), Replacement:=" ", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
  End If
  ' Slach
  Selection.Replace What:=data(16), Replacement:=" ", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' Back Slach
  Selection.Replace What:=data(17), Replacement:=" ", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' A con tilde
  Selection.Replace What:=data(0), Replacement:="A", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' A con tilde invertida
  Selection.Replace What:=data(1), Replacement:="A", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' E con tilde invertida
  Selection.Replace What:=data(2), Replacement:="E", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' E con tilde
  Selection.Replace What:=data(3), Replacement:="E", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' I con tilde invertida
  Selection.Replace What:=data(4), Replacement:="I", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' I con tilde
  Selection.Replace What:=data(5), Replacement:="I", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' O con tilde invertida
  Selection.Replace What:=data(6), Replacement:="O", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' O con tilde
  Selection.Replace What:=data(7), Replacement:="O", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' U con tilde invertida
  Selection.Replace What:=data(8), Replacement:="U", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' U con tilde
  Selection.Replace What:=data(9), Replacement:="U", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  If (ActiveSheet.Name = "OPTO" And (Selection.Address = Range("tbl_opto[DIAG PPAL]").Address Or Selection.Address = Range("tbl_opto[DIAG OBS]").Address Or Selection.Address = Range("tbl_opto[DIAG REL/1]").Address Or Selection.Address = Range("tbl_opto[DIAG REL/2]").Address Or Selection.Address = Range("tbl_opto[DIAG Rel/3]").Address Or Selection.Address = Range("tbl_opto[[DIAG OBS]:[DIAG Rel/3]]").Address)) Then
    ' Coma
    Selection.Replace What:=data(10), Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
  End If
  ' Raya al piso
  Selection.Replace What:=data(11), Replacement:=" ", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' Doble commilla
  Selection.Replace What:=data(12), Replacement:="", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' Espaciado
  Selection.Replace What:=data(13), Replacement:=" ", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' Salto de linea
  Selection.Replace What:=data(14), Replacement:=" ", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  ' Punto
  If (ActiveSheet.Name = "DIAGNOSTICOS" Or ActiveSheet.Name = "ENFASIS") Then
    Selection.Replace What:=data(19), Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
  End If

  If (ActiveSheet.Name ="TRABAJADORES" and Selection.Address = Range("tbl_trabajadores[CARGO USUARIO]").Address) Then
    Call ClearNonAlphaNumeric
  End If
  If (ActiveSheet.Name ="TRABAJADORES" and Selection.Address = Range("tbl_trabajadores[PACIENTE]").Address) Then
    Call ClearNonAlphaNumeric
  End If


  MsgBox "Correcciones realizadas, exitosamente!!",vbInformation,"Correcciones"

End Sub

Sub ClearNonAlphaNumeric()

  Dim valor As String
  Dim ini As String

  Application.ScreenUpdating = False
  ini = ActiveCell.Address
  Do While Not IsEmpty(ActiveCell)
    valor = ActiveCell.value
    ActiveCell = Trim(ReplaceNonAlphaNumeric(valor))
    ActiveCell.Offset(1, 0).Select
  Loop
  Range(ini).Select
  Range(ActiveCell,ActiveCell.End(xlDown)).Select
  Application.ScreenUpdating = True

End Sub

Function ReplaceNonAlphaNumeric(str As String) As String
  Dim regEx As Object
  Set regEx = CreateObject("vbscript.regexp")

  ' Define la expresión regular para encontrar valores no alfanuméricos '
  regEx.Pattern = "[^a-zA-Z0-9"&Chr(209)&"]"
  regEx.Global = True

  ' Reemplaza cualquier valor no alfanumérico por un espacio '
  ReplaceNonAlphaNumeric = regEx.Replace(str, " ")
End Function
