Attribute VB_Name = "Conditionals"

Public Sub dataDuplicate(startCell As String)
  ' Esta subrutina aplica formato a las celdas de una columna para resaltar valores duplicados.
  ' Solo se aplica formato a los valores que aparecen más de una vez en la columna.
  ' Los valores duplicados se resaltan en negrita y color de fondo.

  Dim lastRow As Long
  lastRow = Cells(Rows.Count, Range(startCell).Column).End(xlUp).Row
  Dim rng As Range
  Set rng = Range(startCell, Cells(lastRow, Range(startCell).Column))

  rng.FormatConditions.AddUniqueValues
  rng.FormatConditions(1).SetFirstPriority
  rng.FormatConditions(1).DupeUnique = xlDuplicate
  With rng.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = -0.499984740745262
  End With
  With rng.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15388336
    .TintAndShade = 0
  End With
  rng.FormatConditions(1).StopIfTrue = False
End Sub

Public Sub iqualCero(rngStr As String)
  ' Selecciona un rango y aplica una condición de formato en función del nombre de la hoja activa.
  '
  ' El rango a seleccionar comienza desde la celda actualmente seleccionada y se extiende hasta la última celda no vacía de la columna.
  ' La condición de formato se agrega en función de las siguientes reglas:
  '   - Si la hoja activa es "AUDIO" y las celdas AT4 a AX4 son todas iguales a cero, se formatea la fuente y el fondo de las celdas.
  '   - Si la hoja activa es "VISIO" y las celdas BL4 a BQ4 son todas iguales a cero, se formatea la fuente y el fondo de las celdas.
  '   - Si la hoja activa es "OPTO" y las celdas BD4 a BI4 son todas iguales a cero, se formatea la fuente y el fondo de las celdas.
  '   - Si la hoja activa es "PSICOSENSOMETRICA" y las celdas I3 a N3 son todas iguales a cero, se formatea la fuente y el fondo de las celdas.
  '   - Si la hoja activa es "ESPIRO" y las celdas BN4 a BS4 son todas iguales a cero, se formatea la fuente y el fondo de las celdas.
  '
  ' Se establece la fuente en negrita y se utiliza el color de tema accent1 con un tono y sombra de -0,5.
  ' El fondo se establece en un color sólido con valor RGB de 15388336 (un tono de naranja).
  '
  ' Nota: Esta función asume que la primera fila del rango seleccionado contiene los encabezados de las columnas.

  Dim ws As Worksheet
  Dim rng As Range
  Dim fc As FormatCondition

  Set ws = ActiveSheet
  If IsEmpty(ws.Range(rngStr)(1).offset(1, 0).value) Then
    Set rng = ws.Range(rngStr)
  else
    Set rng = ws.range(rngStr, ws.range(rngStr).End(xlDown))
  End If

  Select Case Trim$(UCase$(ws.Name))
   Case "AUDIO"
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=Y($AT4=0;$AU4=0;$AV4=0;$AW4=0;$AX4=0)")
   Case "VISIO"
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=Y($BL4=0;$BM4=0;$BN4=0;$BO4=0;$BP4=0;$BQ4=0)")
   Case "OPTO"
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=Y($BF4=0;$BG4=0;$BH4=0;$BI4=0;$BD4=0;$BE4=0)")
   Case "PSICOSENSOMETRICA"
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=Y($I3=0;$J3=0;$K3=0;$L3=0;$M3=0;$N3=0)")
   Case "ESPIRO"
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=Y($BN4=0;$BO4=0;$BP4=0;$BQ4=0;$BR4=0;$BS4=0)")
   Case Else
    Exit Sub
  End Select

  With fc
    .SetFirstPriority
    .StopIfTrue = False
    With .Font
      .Bold = True
      .Italic = False
      .ThemeColor = xlThemeColorAccent1
      .TintAndShade = -0.499984740745262
    End With
    With .Interior
      .PatternColorIndex = xlAutomatic
      .Color = 15388336
      .TintAndShade = 0
    End With
  End With
End Sub

Public Sub meetsfails(startCell As String)
  ' Seleccione un rango de celdas y aplique un formato condicional para resaltar las celdas que no contienen "CUMPLE" o "NO CUMPLE".
  '
  ' Este sub no toma argumentos y no devuelve un valor.
  '
  ' Ejemplo:
  '   meetsfails
  '
  ' Esta sub asume que una selección activa de celdas ya ha sido hecha en la hoja de cálculo activa.
  ' Si no se ha seleccionado un rango de celdas, se producirá un error en tiempo de ejecución.

  Dim lastRow As Long
  lastRow = Cells(Rows.Count, Range(startCell).Column).End(xlUp).Row
  Dim rng As Range
  Set rng = Range(startCell, Cells(lastRow, Range(startCell).Column))

  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=Y($D2<>""CUMPLE"";$D2<>""NO CUMPLE"")"
  rng.FormatConditions(1).SetFirstPriority
  With rng.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = -0.499984740745262
  End With
  With rng.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15388336
    .TintAndShade = 0
  End With
  rng.FormatConditions(1).StopIfTrue = False
End Sub

Public Sub Risk(startCell As String)
  'Aplica una regla de formato condicional a un rango seleccionado de celdas.
  'Las celdas que cumplen la condición especificada en la fórmula de la regla se formatean con fuente en negrita, color de tema xlThemeColorAccent1 y un tono de sombreado específico.
  'El color de fondo de la celda se establece en un valor específico.
  'La condición es verdadera si la celda EO5 está vacía y G5 en la hoja TRABAJADORES es PERIODICO, POS INCAPACIDAD, PERIODICO DE SEGUIMIENTO o ESPECIAL.

  Dim lastRow As Long
  lastRow = Cells(Rows.Count, Range(startCell).Column).End(xlUp).Row
  Dim rng As Range
  Set rng = Range(startCell, Cells(lastRow, Range(startCell).Column))

  rng.FormatConditions.Add Type:=xlExpression, Formula1:= _
  "=Y($EO5="""";O(TRABAJADORES!$G5=""PERIODICO"";TRABAJADORES!$G5=""POS INCAPACIDAD"";TRABAJADORES!$G5=""PERIODICO DE SEGUIMIENTO"";TRABAJADORES!$G5=""ESPECIAL""))"
  rng.FormatConditions(1).SetFirstPriority
  With rng.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = -0.499984740745262
  End With
  With rng.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15388336
    .TintAndShade = 0
  End With
  rng.FormatConditions(1).StopIfTrue = False

End Sub

Public Sub riskPre_ingreso(startCell As String)
  'Aplica formato condicional a las celdas seleccionadas que cumplan con la expresión especificada
  'para trabajadores en pre-ingreso

  Dim lastRow As Long
  lastRow = Cells(Rows.Count, Range(startCell).Column).End(xlUp).Row
  Dim rng As Range
  Set rng = Range(startCell, Cells(lastRow, Range(startCell).Column))

  rng.FormatConditions.Add Type:=xlExpression, Formula1:= _
  "=Y($EO5<>"""";TRABAJADORES!$G5=""PRE-INGRESO"")"
  rng.FormatConditions(1).SetFirstPriority
  With rng.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent4
    .TintAndShade = -0.499984740745262
  End With
  With rng.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 11791359
    .TintAndShade = 0
  End With
  rng.FormatConditions(1).StopIfTrue = False
End Sub

Public Sub formatter(startCell As String)
  ' Este subrutina formatea la selección en la hoja de cálculo activa
  ' El formato numérico se establece en "0" y la altura de fila se establece en 40

  Dim lastRow As Long
  lastRow = Cells(Rows.Count, Range(startCell).Column).End(xlUp).Row
  Dim rng As Range
  Set rng = Range(startCell, Cells(lastRow, Range(startCell).Column))

  With rng
    .NumberFormat = "0"
    .RowHeight = 40
  End With

End Sub

Public Sub greaterThanOne(rngStr As String)
  'Este sub selecciona un rango de celdas desde la celda activa hasta la última celda con datos en la columna hacia abajo.
  'Luego, agrega una condición de formato en función del nombre de la hoja activa y la suma de ciertas celdas del rango seleccionado.
  'Si la suma es mayor que 1, se aplica un formato de fuente en negrita y color de fondo en naranja.

  Dim ws As Worksheet
  Dim rng As Range
  Dim fc As FormatCondition

  Set ws = ActiveSheet
  If IsEmpty(ws.Range(rngStr)(1).offset(1, 0).value) Then
    Set rng = ws.range(rngStr)
  else
    Set rng = ws.range(rngStr, ws.range(rngStr).End(xlDown))
  End If

  Select Case Trim$(UCase$(ws.Name))
   Case "AUDIO"
    Set fc = rng.FormatConditions.Add (Type:=xlExpression, Formula1:="=SUMA($AT4;$AU4;$AV4;$AW4;$AX4)>1")
   Case "VISIO"
    Set fc = rng.FormatConditions.Add (Type:=xlExpression, Formula1:="=SUMA($BL4;$BM4;$BN4;$BO4;$BP4;$BQ4)>1")
   Case "OPTO"
    Set fc = rng.FormatConditions.Add (Type:=xlExpression, Formula1:="=SUMA($BD4;$BE4;$BF4;$BG4;$BH4;$BI4)>1")
   Case "PSICOSENSOMETRICA"
    Set fc = rng.FormatConditions.Add (Type:=xlExpression, Formula1:="=SUMA($I3;$J3;$K3;$L3;$M3;$N3)>1")
   Case "ESPIRO"
    Set fc = rng.FormatConditions.Add (Type:=xlExpression, Formula1:="=SUMA($BN4;$BO4;$BP4;$BQ4;$BR4;$BS4)>1")
   Case Else
    Exit Sub
  End Select

  With fc
    .SetFirstPriority
    .StopIfTrue = False
    With .Font
      .Bold = True
      .Italic = False
      .ThemeColor = xlThemeColorAccent1
      .TintAndShade = -0.499984740745262
    End With
    With .Interior
      .PatternColorIndex = xlAutomatic
      .Color = 15388336
      .TintAndShade = 0
    End With
  End With

End Sub

Public Sub thisText(startCell As String)
  ' Este procedimiento selecciona el rango desde la celda activa hacia abajo y agrega una condición de formato para resaltar las celdas que contienen texto en la columna BH.

  Dim lastRow As Long
  lastRow = Cells(Rows.Count, Range(startCell).Column).End(xlUp).Row
  Dim rng As Range
  Set rng = Range(startCell, Cells(lastRow, Range(startCell).Column))

  rng.FormatConditions.Add Type:=xlExpression, Formula1:= _
  "=ESTEXTO($BH5)"
  rng.FormatConditions(1).SetFirstPriority
  With rng.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = -0.499984740745262
  End With
  With rng.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15388336
    .TintAndShade = 0
  End With
  rng.FormatConditions(1).StopIfTrue = False

End Sub

Public Sub thisEgreso(startCell As String)
  'Aplica una condición de formato a la selección actual si el valor de la columna G es "EGRESO".

  Dim lastRow As Long
  lastRow = Cells(Rows.Count, Range(startCell).Column).End(xlUp).Row
  Dim rng As Range
  Set rng = Range(startCell, Cells(lastRow, Range(startCell).Column))

  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=$G5=""EGRESO"""
  rng.FormatConditions(1).SetFirstPriority
  With rng.FormatConditions(1).Font
    .Bold = True
    .Italic = False
    .Color = -16777024
    .TintAndShade = 0
  End With
  With rng.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 15198207
    .TintAndShade = 0
  End With
  rng.FormatConditions(1).StopIfTrue = False

End Sub

' Elimina todas las condiciones de formato de la hoja de calculo activa
Public Sub deleteFormatConditions()
  ' Declarar variables de hoja de calculo y rango
  Dim ws As Worksheet
  Dim rng As Range

  ' Establece la variable de hoja de calculo en la hoja activa y la variable de rango en todas las celdas de la hoja de calculo
  Set ws = ActiveSheet
  Set rng = ws.Cells

  ' Desactiva la actualizacion de pantalla para mejorar el rendimiento
  Application.ScreenUpdating = False

  ' Elimina todas las condiciones de formato del rango de celdas
  rng.FormatConditions.Delete

  ' Vuelve a activar la actualizacion de pantalla
  Application.ScreenUpdating = True
End Sub
