Attribute VB_Name = "config_emph"
'namespace=vba-files\Module\configs
Option Explicit

Public Sub configEmphasis()
  '
  ' configEnfasis: realiza la configuracion inicial de las cabeceras asi como el nombre de la tabla
  '
  '
  range("A4") = "IDENTIFICACION"
  range("B4") = "id_emo"
  range("C4") = "ENFASIS_1"
  range("D4") = "CONCEPTO AL ENFASIS_1"
  range("E4") = "OBSERVACIONES_AL_ENFASIS_1"
  range("F4") = "SQL ENFASIS_1"
  range("G4") = "ENFASIS_2"
  range("H4") = "CONCEPTO AL ENFASIS_2"
  range("I4") = "OBSERVACIONES AL ENFASIS_2"
  range("J4") = "SQL ENFASIS_2"
  range("K4") = "ENFASIS_3"
  range("L4") = "CONCEPTO AL ENFASIS_3"
  range("M4") = "OBSERVACIONES AL ENFASIS_3"
  range("N4") = "SQL ENFASIS_3"
  range("O4") = "ENFASIS_4"
  range("P4") = "CONCEPTO AL ENFASIS_4"
  range("Q4") = "OBSERVACIONES AL ENFASIS_4"
  range("R4") = "SQL ENFASIS_4"
  range("S4") = "ENFASIS_5"
  range("T4") = "CONCEPTO AL ENFASIS_5"
  range("U4") = "OBSERVACIONES AL ENFASIS_5"
  range("V4") = "SQL ENFASIS_5"
  range("W4") = "ENFASIS_6"
  range("X4") = "CONCEPTO AL ENFASIS_6"
  range("Y4") = "OBSERVACIONES AL ENFASIS_6"
  range("Z4") = "SQL ENFASIS_6"
  range("AA4") = "ENFASIS_7"
  range("AB4") = "CONCEPTO AL ENFASIS_7"
  range("AC4") = "OBSERVACIONES AL ENFASIS_7"
  range("AD4") = "SQL ENFASIS_7"
  range("AE4") = "ENFASIS_8"
  range("AF4") = "CONCEPTO AL ENFASIS_8"
  range("AG4") = "OBSERVACIONES AL ENFASIS_8"
  range("AH4") = "SQL ENFASIS_8"
  range("AI4") = "ENFASIS_9"
  range("AJ4") = "CONCEPTO AL ENFASIS_9"
  range("AK4") = "OBSERVACIONES AL ENFASIS_9"
  range("AL4") = "SQL ENFASIS_9"
  range("AM4") = "ENFASIS_10"
  range("AN4") = "CONCEPTO AL ENFASIS_10"
  range("AO4") = "OBSERVACIONES AL ENFASIS_10"
  range("AP4") = "SQL ENFASIS_10"
  range("AQ4") = "ENFASIS_11"
  range("AR4") = "CONCEPTO AL ENFASIS_11"
  range("AS4") = "OBSERVACIONES AL ENFASIS_11"
  range("AT4") = "SQL ENFASIS_11"
  range("AU4") = "ENFASIS_12"
  range("AV4") = "CONCEPTO AL ENFASIS_12"
  range("AW4") = "OBSERVACIONES AL ENFASIS_12"
  range("AX4") = "SQL ENFASIS_12"
  range("AY4") = "ENFASIS_13"
  range("AZ4") = "CONCEPTO AL ENFASIS_13"
  range("BA4") = "OBSERVACIONES AL ENFASIS_13"
  range("BB4") = "SQL ENFASIS_13"
  range("BC4") = "ENFASIS_14"
  range("BD4") = "CONCEPTO AL ENFASIS_14"
  range("BE4") = "OBSERVACIONES AL ENFASIS_14"
  range("BF4") = "SQL ENFASIS_14"
  range("BG4") = "ENFASIS_15"
  range("BH4") = "CONCEPTO AL ENFASIS_15"
  range("BI4") = "OBSERVACIONES AL ENFASIS_15"
  range("BJ4") = "SQL ENFASIS_15"
  range("BK4") = "ENFASIS_16"
  range("BL4") = "CONCEPTO AL ENFASIS_16"
  range("BM4") = "OBSERVACIONES AL ENFASIS_16"
  range("BN4") = "SQL ENFASIS_16"
  range("BO4") = "ENFASIS_17"
  range("BP4") = "CONCEPTO AL ENFASIS_17"
  range("BQ4") = "OBSERVACIONES AL ENFASIS_17"
  range("BR4") = "SQL ENFASIS_17"
  range("BS4") = "ENFASIS_18"
  range("BT4") = "CONCEPTO AL ENFASIS_18"
  range("BU4") = "OBSERVACIONES AL ENFASIS_18"
  range("BV4") = "SQL ENFASIS_18"
  ActiveSheet.ListObjects.Add(xlSrcRange, range("$A$4:$BV$5"), , xlYes).Name = _
  "tbl_enfasis"
  ActiveSheet.ListObjects("tbl_enfasis").TableStyle = "TableStyleLight9"

  range("tbl_enfasis[[#Headers],[ENFASIS_1]:[SQL ENFASIS_1]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(55, 86, 35)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("C3") = "OSTEOMUSCULAR"
  range("C3:F3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_2]:[SQL ENFASIS_2]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(31, 78, 120)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("G3") = "ALTURAS"
  range("G3:J3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_3]:[SQL ENFASIS_3]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(128, 96, 0)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("K3") = "ALIMENTOS"
  range("K3:N3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_4]:[SQL ENFASIS_4]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(82, 82, 82)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("O3") = "ESPACIOS CONFINADOS"
  range("O3:R3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_5]:[SQL ENFASIS_5]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(131, 60, 12)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("S3") = "SEGURIDAD VIAL"
  range("S3:V3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_6]:[SQL ENFASIS_6]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(32, 55, 100)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("W3") = "BRIGADISTA"
  range("W3:Z3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_7]:[SQL ENFASIS_7]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(34, 43, 53)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("AA3") = "MEDICAMENTOS"
  range("AA3:AD3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_8]:[SQL ENFASIS_8]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(128, 96, 0)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("AE3") = "QUIMICOS"
  range("AE3:AH3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_9]:[SQL ENFASIS_9]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(13, 13, 13)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("AI3") = "ACTIVIDAD DEPORTIVA"
  range("AI3:AL3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_10]:[SQL ENFASIS_10]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(198, 89, 17)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("AM3") = "CARDIOVASCULAR"
  range("AM3:AP3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_11]:[SQL ENFASIS_11]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(31, 78, 120)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("AQ3") = "TRABAJO CON ENERGIAS PELIGROSAS ALTA TENSION"
  range("AQ3:AT3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_12]:[SQL ENFASIS_12]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(192, 0, 0)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("AU3") = "TRABAJO CON TEMPERATURAS EXTREMAS BAJAS"
  range("AU3:AX3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_13]:[SQL ENFASIS_13]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(89, 89, 89)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("AY3") = "TRABAJO EN ALTITUDES MAYORES A 2500 METROS SOBRE EL NIVEL DEL MAR"
  range("AY3:BB3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_14]:[SQL ENFASIS_14]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(55, 86, 35)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("BC3") = "TRABAJO EN AMBIENTES HIPERBARICOS"
  range("BC3:BF3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_15]:[SQL ENFASIS_15]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(32, 55, 100)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("BG3") = "TRABAJO EN AMBIENTES HIPERBARICOS"
  range("BG3:BJ3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_16]:[SQL ENFASIS_16]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(55, 86, 35)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("BK3") = "RADIACIONES IONIZANTES"
  range("BK3:BN3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_17]:[SQL ENFASIS_17]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(22, 22, 22)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("BO3") = "AEROPORTUARIO"
  range("BO3:BR3").Select
  selection.Merge
  range("tbl_enfasis[[#Headers],[ENFASIS_18]:[SQL ENFASIS_18]]").Select
  range(selection, selection.Offset(-1, 0)).Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = RGB(32, 55, 100)
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  With selection.Font
    .Color = RGB(255, 255, 255)
    .TintAndShade = 0
  End With
  Call center(selection)
  range("BS3") = "RESPIRATORIO"
  range("BS3:BV3").Select
  selection.Merge
  
  Call formatTable("tbl_enfasis")
  
End Sub

Private Function center(ByVal selection As range)
  With selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
  End With
End Function

Private Sub ConvertirNumeroARGB()
    Dim numero As Variant, arrayNumero As Variant
    Dim r As Integer, g As Integer, b As Integer

    arrayNumero = Array(2315831, 7884319, 24704, 5395026, 801923, 6567712, 3484450, 24704, 855309, 1137094, _
    7884319, 192, 5855577, 2315831, 6567712, 2315831, 1447446, 6567712)
    
    For Each numero In arrayNumero
      'Extraer los componentes RGB
      r = numero Mod 256
      g = (numero \ 256) Mod 256
      b = (numero \ 256 \ 256) Mod 256
      'Mostrar el c�digo RGB en la ventana de resultados
      Debug.Print CStr("El codigo RGB es: RGB(" & r & ", " & g & ", " & b & ")")
    Next numero
    
End Sub
