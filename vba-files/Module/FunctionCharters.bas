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
   Case "BOGOTA", "BOGOTA, D.C.", "BOGOT" & Chr(193) & ", D.C.", "BOGOTA, D.C", "BOGOTA D.C"
    city = Trim("BOGOTA D.C.")
   Case "CARTAGENA DE INDIAS"
    city = Trim("CARTAGENA")
   Case "BUGA"
    city = Trim("GUADALAJARA DE BUGA")
   Case "MONTEL" & Chr(205) & "BANO"
    city = Trim("MONTELIBANO")
   Case "PUERTO GAIT" & Chr(193) & "N"
    city = Trim("PUERTO GAITAN")
   Case "PUERTO BOYAC" & Chr(193)
    city = Trim("PUERTO BOYACA")
   case "PUERTO AS"&Chr(205)&"S"
    city = Trim("PUERTO ASIS")
   Case Else
    city = value
  End Select
End Function

'/*
' VALIDACION ESCOLARIDAD
'*/
Public Function school(ByVal value As String) As String
  Select Case value
   Case "POSTGRADO"
    school = "POSGRADO"
   Case Else
    school = value
  End Select

End Function

'/*
' VALIDACION EXAMEN MEDICO
'*/
Public Function typeExams(ByVal value As String) As String
  Select Case value
   Case "POST INCAPACIDAD"
    typeExams = "POS INCAPACIDAD"
   Case "PERIODICO SEG", "PERIODICO SEGUIMIENTO"
    typeExams = "PERIODICO"
   Case "CAMBIO OCUPACION","CAMBIO DE OCUPACI" & Chr(211) & "N"
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
   Case "F"&Chr(205)&"SICAMENTE ACTIVO", "FISICAMENTE ACTIVO", "FISICAMENTE ACTIVO(A)", "F"&Chr(205)&"SICAMENTE ACTIVO(A)"
    typeActivity = "F"&Chr(205)&"SICAMENTE ACTIVO"
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

  total = (emo*4) + audio + visio + espiro + osteo + complementarios + psicotecnica + psicosensometrica + opto

End Function
