Attribute VB_Name = "Conditionals"

Sub dataDuplicate()

  Range(Selection, Selection.End(xlDown)).Select
  Selection.FormatConditions.AddUniqueValues
  Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
  Selection.FormatConditions(1).DupeUnique = xlDuplicate
  With Selection.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = -0.499984740745262
  End With
  With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15388336
    .TintAndShade = 0
  End With
  Selection.FormatConditions(1).StopIfTrue = False

End Sub

Sub iqualCero()

  Range(Selection, Selection.End(xlDown)).Select
  Select Case Trim(UCase(ActiveSheet.Name))
   Case "AUDIO"
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=Y($AT4=0;$AU4=0;$AV4=0;$AW4=0;$AX4=0)"
   Case "VISIO"
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=Y($BL4=0;$BM4=0;$BN4=0;$BO4=0;$BP4=0;$BQ4=0)"
   Case "OPTO"
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=Y($BF4=0;$BG4=0;$BH4=0;$BI4=0;$BD4=0;$BE4=0)"
   Case "PSICOSENSOMETRICA"
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=Y($I3=0;$J3=0;$K3=0;$L3=0;$M3=0;$N3=0)"
  End Select
  Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
  With Selection.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = -0.499984740745262
  End With
  With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15388336
    .TintAndShade = 0
  End With
  Selection.FormatConditions(1).StopIfTrue = False

End Sub

Sub meetsfails()

  Range(Selection, Selection.End(xlDown)).Select
  Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
  "=Y($D2<>""CUMPLE"";$D2<>""NO CUMPLE"")"
  Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
  With Selection.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = -0.499984740745262
  End With
  With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15388336
    .TintAndShade = 0
  End With
  Selection.FormatConditions(1).StopIfTrue = False

End Sub

Sub Risk()

  Range(Selection, Selection.End(xlDown)).Select
  Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
  "=Y($EO5="""";O(TRABAJADORES!$G5=""PERIODICO"";TRABAJADORES!$G5=""POS INCAPACIDAD"";TRABAJADORES!$G5=""PERIODICO DE SEGUIMIENTO"";TRABAJADORES!$G5=""ESPECIAL""))"
  Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
  With Selection.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = -0.499984740745262
  End With
  With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15388336
    .TintAndShade = 0
  End With
  Selection.FormatConditions(1).StopIfTrue = False

End Sub

Sub riskPre_ingreso()

  Range(Selection, Selection.End(xlDown)).Select
  Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
  "=Y($EO5<>"""";TRABAJADORES!$G5=""PRE-INGRESO"")"
  Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
  With Selection.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent4
    .TintAndShade = -0.499984740745262
  End With
  With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 11791359
    .TintAndShade = 0
  End With
  Selection.FormatConditions(1).StopIfTrue = False

End Sub

Sub formatter()

  Selection.NumberFormat = "0"
  Selection.RowHeight = 40

End Sub

Sub greaterThanOne()

  Range(Selection, Selection.End(xlDown)).Select
  Select Case Trim(UCase(ActiveSheet.Name))
   Case "AUDIO"
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=SUMA($AT4;$AU4;$AV4;$AW4;$AX4)>1"
   Case "VISIO"
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=SUMA($BL4;$BM4;$BN4;$BO4;$BP4;$BQ4)>1"
   Case "OPTO"
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=SUMA($BD4;$BE4;$BF4;$BG4;$BH4;$BI4)>1"
   Case "PSICOSENSOMETRICA"
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=SUMA($I3;$J3;$K3;$L3;$M3;$N3)>1"
  End Select
  Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
  With Selection.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = -0.499984740745262
  End With
  With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15388336
    .TintAndShade = 0
  End With
  Selection.FormatConditions(1).StopIfTrue = False

End Sub

Sub thisText()

  Range(Selection, Selection.End(xlDown)).Select
  Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
  "=ESTEXTO($BH5)"
  Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
  With Selection.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = -0.499984740745262
  End With
  With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15388336
    .TintAndShade = 0
  End With
  Selection.FormatConditions(1).StopIfTrue = False

End Sub

Sub thisEgreso()

  Range(Selection, Selection.End(xlDown)).Select
  Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
  "=$G5=""EGRESO"""
  Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
  With Selection.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .Color = -16777024
    .TintAndShade = 0
  End With
  With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15198207
    .TintAndShade = 0
  End With
  Selection.FormatConditions(1).StopIfTrue = False

End Sub

Sub ClearCharter()
  Attribute ClearCharter.VB_ProcData.VB_Invoke_Func = "y\n14"

  Dim data As Variant

  data = Array(Chr(193), Chr(192), Chr(200), Chr(201), Chr(204), Chr(205), Chr(210), Chr(211), Chr(217), Chr(218), Chr(44), Chr(46), Chr(147), Chr(13), Chr(10), Chr(160) & Chr(160), Chr(92), Chr(47), Chr(45))

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
  ' Punto
  ' Selection.Replace What:=data(11), Replacement:="", LookAt:=xlPart, _
  ' SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ' ReplaceFormat:=False
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

  MsgBox "Correcciones realizadas, exitosamente!!",vbInformation,"Correcciones"

End Sub
