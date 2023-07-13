Attribute VB_Name = "config_workers"
'namespace=vba-files\Module\configs
Option Explicit

Public Sub addSheets()
  '
  ' Sheets: a�ade todas las hojas requeridas
  '
  '
  Dim arraySheets As Variant
  Dim sheetName As Variant

  arraySheets = Array("DIAGNOSTICOS", "ENFASIS", "TRABAJADORES", "EMO", "AUDIO", "VISIO", "OPTO", "ESPIRO", "OSTEO", "COMPLEMENTARIOS", _
  "PSICOTECNICA", "PSICOSENSOMETRICA", "RUTAS")

  For Each sheetName In arraySheets
    Sheets.Add After:=ActiveSheet
    Sheets(Sheets.Count).Name = sheetName
  Next sheetName
End Sub

Public Sub configWorkers()
Attribute configWorkers.VB_ProcData.VB_Invoke_Func = " \n14"
  '
  ' configTrabajadores: realiza la configuracion inicial de las cabeceras asi como el nombre de la tabla
  '
  '
  range("A4") = "estado"
  range("B4") = "NOMBRE CONTRATO"
  range("C4") = "LLAVE"
  range("D4") = "DESTINO"
  range("E4") = "CIUDAD"
  range("F4") = "INGRESO"
  range("G4") = "TIPO EXAMEN"
  range("H4") = "FECHA INGRESO"
  range("I4") = "PACIENTE"
  range("J4") = "NRO IDENFICACION"
  range("K4") = "EDAD"
  range("L4") = "rango_edad"
  range("M4") = "ESTRATO"
  range("N4") = "GENERO"
  range("O4") = "NRO HIJOS"
  range("P4") = "hijos"
  range("Q4") = "RAZA"
  range("R4") = "ESTADO CIVIL"
  range("S4") = "ESCOLARIDAD"
  range("T4") = "CARGO USUARIO"
  range("U4") = "CARGO_REC"
  range("V4") = "LAB DURACION EN A" & ChrW(209) & "OS"
  range("W4") = "ANTIGUEDAD"
  range("X4") = "FUENTE"
  range("Y4") = "TIPO ACTIVIDAD"
  range("Z4") = "analista"
  range("AA4") = "profesional"
  range("AB4") = "fecha_inicio"
  range("AC4") = "fecha_fin"
  range("AD4") = "tipo examen solicitud"
  range("AE4") = "CIUDAD_ID"
  range("AF4") = "id_tipo_examen"
  range("AG4") = "fecha_texto"
  range("AH4") = "id_raza"
  range("AI4") = "id_estado_civil"
  range("AJ4") = "id_escolaridad"
  range("AK4") = "id_cargo"
  range("AL4") = "fuente2"
  range("AM4") = "(id_tipo_actividad)"
  range("AN4") = "AUDIO"
  range("AO4") = "OPTO"
  range("AP4") = "ESPIRO"
  range("AQ4") = "VISIO"
  range("AR4") = "OSTEO"
  range("AS4") = "PSICOSENSOMETRICA"
  range("AT4") = "PSICOTECNICA"
  range("AU4") = "COMPLEMENTARIOS"
  range("AV4") = "EMO"
  range("AW4") = "idOrdenListaTrabajadores"
  range("AX4") = "idOrden"
  range("AY4") = "SCRIPT ordenes"
  range("AZ4") = "SCRIPT ordenes_tipo_actividad"
  range("BA4") = "SCRIPT ordenes_tipo_examen"
  range("BB4") = "SCRIPT orden_informe"
  range("BC4") = "SCRIPT orden_lista_trabajadores"
  range("BD4") = "SCRIPT ordenes_trabajador_paraclinicos"
  ActiveSheet.ListObjects.Add(xlSrcRange, range("$A$4:$BD$5"), , xlYes).Name = _
        "tbl_trabajadores"
  ActiveSheet.ListObjects("tbl_trabajadores").TableStyle = "TableStyleLight9"
        
  range("tbl_trabajadores[[#Headers],[LLAVE]]").Style = "Neutral"
  range("tbl_trabajadores[[#Data],[LLAVE]]").Style = "Notas"
  range("tbl_trabajadores[[#Headers],[rango_edad]]").Style = "Neutral"
  range("tbl_trabajadores[[#Data],[rango_edad]]").Style = "Notas"
  range("tbl_trabajadores[[#Headers],[hijos]]").Style = "Neutral"
  range("tbl_trabajadores[[#Data],[hijos]]").Style = "Notas"
  range("tbl_trabajadores[[#Headers],[CARGO_REC]]").Style = "Neutral"
  range("tbl_trabajadores[[#Data],[CARGO_REC]]").Style = "Notas"
  range("tbl_trabajadores[[#Headers],[ANTIGUEDAD]]").Style = "Neutral"
  range("tbl_trabajadores[[#Data],[ANTIGUEDAD]]").Style = "Notas"
  range("tbl_trabajadores[[#Headers],[CIUDAD_ID]:[EMO]]").Style = "Neutral"
  range("tbl_trabajadores[[#Data],[CIUDAD_ID]:[EMO]]").Style = "Notas"
  range("tbl_trabajadores[[#Headers],[SCRIPT ordenes]:[SCRIPT ordenes_trabajador_paraclinicos]]" _
        ).Style = "Celda de comprobaci" & ChrW(243) & "n"
  range("tbl_trabajadores[[#Data],[SCRIPT ordenes]:[SCRIPT ordenes_trabajador_paraclinicos]]" _
        ).Style = "Salida"
  
  Call formatTable("tbl_trabajadores")
  Rows("5:5").Select
  ActiveWindow.FreezePanes = True
  
        
End Sub
Public Sub formatTable(ByVal tblName As String)

  range(tblName & "[#Headers]").Select
  selection.ColumnWidth = 20
  selection.RowHeight = 30
  range(tblName & "[#Data]").Select
  selection.RowHeight = 40
  range(tblName & "[#All]").Select
  With selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
  End With
  
End Sub
