Attribute VB_Name = "config_diag"
'namespace=vba-files\Module\configs
Option Explicit

Public Sub configDiagnostics()
  '
  ' configDiagnosticos: realiza la configuracion inicial de las cabeceras asi como el nombre de la tabla
  '
  '
  range("A4") = "IDENTIFICACION"
  range("B4") = "id emo"
  range("C4") = "TODO"
  range("D4") = "CODIGO DIAG PPAL"
  range("E4") = "DIAG PPAL"
  range("F4") = "CODIGO DIAG REL1"
  range("G4") = "DIAG REL 1"
  range("H4") = "CODIGO DIAG REL2"
  range("I4") = "DIAG REL 2"
  range("J4") = "CODIGO DIAG REL3"
  range("K4") = "DIAG REL 3"
  range("L4") = "CODIGO DIAG REL4"
  range("M4") = "DIAG REL 4"
  range("N4") = "CODIGO DIAG REL5"
  range("O4") = "DIAG REL 5"
  range("P4") = "CODIGO DIAG REL6"
  range("Q4") = "DIAG REL 6"
  range("R4") = "CODIGO DIAG REL7"
  range("S4") = "DIAG REL 7"
  range("T4") = "CODIGO DIAG REL8"
  range("U4") = "DIAG REL 8"
  range("V4") = "CODIGO DIAG REL9"
  range("W4") = "DIAG REL 9"
  range("X4") = "CODIGO DIAG REL10"
  range("Y4") = "DIAG REL 10"
  range("Z4") = "CODIGO DIAG REL11"
  range("AA4") = "DIAG REL 11"
  range("AB4") = "CODIGO DIAG REL12"
  range("AC4") = "DIAG REL 12"
  range("AD4") = "CODIGO DIAG REL13"
  range("AE4") = "DIAG REL 13"
  range("AF4") = "CODIGO DIAG REL14"
  range("AG4") = "DIAG REL 14"
  range("AH4") = "CODIGO DIAG REL15"
  range("AI4") = "DIAG REL 15"
  range("AJ4") = "CODIGO DIAG REL16"
  range("AK4") = "DIAG REL 16"
  range("AL4") = "CODIGO DIAG REL17"
  range("AM4") = "DIAG REL 17"
  range("AN4") = "CODIGO DIAG REL18"
  range("AO4") = "DIAG REL 18"
  range("AP4") = "CODIGO DIAG REL19"
  range("AQ4") = "DIAG REL 19"
  range("AR4") = "CODIGO DIAG REL20"
  range("AS4") = "DIAG REL 20"
  ActiveSheet.ListObjects.Add(xlSrcRange, range("$A$4:$As$5"), , xlYes).Name = _
  "tbl_diagnosticos"
  ActiveSheet.ListObjects("tbl_diagnosticos").TableStyle = "TableStyleLight9"
  
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG PPAL]:[DIAG PPAL]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 15189684
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL1]:[DIAG REL 1]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 11389944
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL2]:[DIAG REL 2]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 14408667
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL3]:[DIAG REL 3]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 10086143
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL4]:[DIAG REL 4]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 15652797
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL5]:[DIAG REL 5]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 11854022
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL6]:[DIAG REL 6]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 15189684
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL7]:[DIAG REL 7]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 11389944
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL8]:[DIAG REL 8]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 14408667
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL9]:[DIAG REL 9]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 10086143
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL10]:[DIAG REL 10]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 15652797
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL11]:[DIAG REL 11]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 11854022
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL12]:[DIAG REL 12]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 15189684
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL13]:[DIAG REL 13]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 11389944
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL14]:[DIAG REL 14]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 14408667
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL15]:[DIAG REL 15]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 10086143
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL16]:[DIAG REL 16]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 15652797
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL17]:[DIAG REL 17]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 11854022
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL18]:[DIAG REL 18]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 15189684
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL19]:[DIAG REL 19]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 11389944
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG REL20]:[DIAG REL 20]]").Select
  With selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 14408667
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  range("tbl_diagnosticos[[#Headers],[CODIGO DIAG PPAL]:[DIAG REL 20]]").Select
  With selection.Font
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
  End With
  range("tbl_diagnosticos[[#headers],[IDENTIFICACION]]").Select
  
  Call formatTable("tbl_diagnosticos")

End Sub

