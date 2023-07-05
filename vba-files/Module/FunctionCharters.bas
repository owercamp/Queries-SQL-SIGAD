Attribute VB_Name = "FunctionCharters"
Option Explicit

'''? <summary>
'''? Esta funcion convierte una cadena de texto a mayusculas y elimina los espacios en blanco alrededor del valor de entrada.
'''? </summary>
'''? @param <param name="value">El valor de entrada como cadena de texto.</param>
'''? @return <returns>El valor de entrada en mayusculas y sin espacios en blanco.</returns>
Public Function charters(ByVal value As String) As String
  charters = Trim(UCase(value))
End Function

'TODO: Elimina los espacios al inicio y al final de cada valor y verifica que no sea un campo vacio.
'? Parametros:
'?@param - value: El valor que se va a verificar.
'? Retorno:
'? @return - Si el valor es un campo vacio o una cadena vacia o "NO", devuelve "0".
'? @return - Si el valor es "OCASIONAL" o "SI", devuelve "1".
'? @return - En cualquier otro caso, devuelve el valor sin espacios al inicio y al final en mayusculas.
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

'TODO: Devuelve una cadena de caracteres con el texto proporcionado rellenado a la izquierda con el caracter de relleno especificado
'? hasta alcanzar la longitud total especificada.
'? Parametros:
'? @param text: El texto que se va a rellenar a la izquierda.
'? @param totalLength: La longitud total de la cadena resultante, incluyendo el texto y los caracteres de relleno.
'? @param padCharacter: El caracter utilizado para rellenar a la izquierda el texto.
'? Devuelve:
'? @return Una cadena de caracteres con el texto proporcionado rellenado a la izquierda con el caracter de relleno especificado hasta alcanzar la longitud total especificada.
Public Function PadLeft(text As Variant, totalLength As Integer, padCharacter As String) As String
  PadLeft = String(totalLength - Len(CStr(text)), padCharacter) & CStr(text)
End Function

'TODO: Devuelve una cadena de caracteres con el texto proporcionado rellenado a la derecha con el caracter de relleno especificado
'? hasta alcanzar la longitud total especificada.
'? Parametros:
'? @param text: El texto que se va a rellenar a la derecha.
'? @param totalLength: La longitud total de la cadena resultante, incluyendo el texto y los caracteres de relleno.
'? @param padCharacter: El caracter utilizado para rellenar a la derecha el texto.
'? Devuelve:
'? @return Una cadena de caracteres con el texto proporcionado rellenado a la derecha con el caracter de relleno especificado hasta alcanzar la longitud total especificada.
Public Function PadRight(text As Variant, totalLength As Integer, padCharacter As String) As String
  PadRight = CStr(text) & String(totalLength - Len(CStr(text)), padCharacter)
End Function

'TODO Funcion: city
'? Descripcion: Esta funcion recibe una cadena de texto "value" que representa un nombre de ciudad y devuelve una cadena de texto representando una version estandarizada del nombre de la ciudad. Si el valor de entrada coincide con uno de los casos listados en la instruccion Select Case, se devuelve el nombre de ciudad estandarizado correspondiente. Si el valor de entrada no coincide con ninguno de los casos, se devuelve el valor de entrada original.
'? Parametros:
'? @param value: Cadena de texto que representa un nombre de ciudad.
'? Retorno:
'? @return Cadena de texto representando una version estandarizada del nombre de ciudad.
Public Function city(ByVal value As String) As String
  Select Case value
   Case "BOGOTA", "BOGOTA, D.C.", "BOGOT" & Chr(193) & ", D.C.", "BOGOTA, D.C", "BOGOTA D.C", "BOGOT" & Chr(193), "BOGOTA  D.C", "BOGOTA, BOGOTA D.C", "BOGOTA,D,C", "BOGOTA  D C", "BOGOTa, D,C,", "BOGOTA,D.C", "BOGOTA, DC"
    city = Trim("BOGOTA D.C.")
   Case "CARTAGENA DE INDIAS", "CARTAGENA, BOLIVAR"
    city = Trim("CARTAGENA")
   Case "BUGA", "GUADALAJARA"
    city = Trim("GUADALAJARA DE BUGA")
   Case "MONTEL" & Chr(205) & "BANO"
    city = Trim("MONTELIBANO")
   Case "PUERTO GAIT" & Chr(193) & "N"
    city = Trim("PUERTO GAITAN")
   Case "PUERTO BOYAC" & Chr(193)
    city = Trim("PUERTO BOYACA")
   Case "PUERTO AS" & Chr(205) & "S", "PUETRTO ASIS"
    city = Trim("PUERTO ASIS")
   Case "TULU" & Chr(193)
    city = Trim("TULUA")
   Case "POPAY" & Chr(193) & "N"
    city = Trim("POPAYAN")
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
   Case "RIO NEGRO"
    city = "RIONEGRO"
   Case Else
    city = value
  End Select
End Function

'TODO Funcion: school
'? Descripcion: Esta funcion recibe una cadena de texto "value" que representa un nivel de educacion y devuelve una cadena de texto representando una version estandarizada del nivel de educacion. Si el valor de entrada coincide con uno de los casos listados en la instruccion Select Case, se devuelve el nivel de educacion estandarizado correspondiente. Si el valor de entrada no coincide con ninguno de los casos, se devuelve el valor de entrada original.
'? Parametros:
'? @param value: Cadena de texto que representa un nivel de educacion.
'? Retorno:
'? @return Cadena de texto representando una version estandarizada del nivel de educacion.
Public Function school(ByVal value As String) As String
  Select Case value
   Case "POSTGRADO", "POST GRADO"
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

'TODO Funcion: typeExams
'? Descripcion: Esta funcion recibe una cadena de texto "value" que representa un tipo de examen y devuelve una cadena de texto representando una version estandarizada del tipo de examen. Si el valor de entrada coincide con uno de los casos listados en la instruccion Select Case, se devuelve el tipo de examen estandarizado correspondiente. Si el valor de entrada no coincide con ninguno de los casos, se devuelve el valor de entrada original.
'? Parametros:
'? @param value: Cadena de texto que representa un tipo de examen.
'? Retorno:
'? @return Cadena de texto representando una version estandarizada del tipo de examen.
Public Function typeExams(ByVal value As String) As String
  Select Case value
   Case "POST INCAPACIDAD", "POST-INCAPACIDAD"
    typeExams = "POS INCAPACIDAD"
   Case "PERIODICO SEG"
    typeExams = "PERIODICO"
   Case "PERIODICO SEGUIMIENTO", "PERIODICO CON RECOMENDACIONES", "PERIODICO CON SEGUIMIENTO"
    typeExams = "PERIODICO DE SEGUIMIENTO"
   Case "CAMBIO OCUPACION", "CAMBIO DE OCUPACI" & Chr(211) & "N"
    typeExams = "CAMBIO DE OCUPACION"
   Case "REINTEGRO LABORAL", "OTROS REINTEGROS"
    typeExams = "EGRESO"
   Case "PRE-INGRESO", "PRE_INGRESO", "INGRESO", "PRE INGRESO"
    typeExams = "PRE-INGRESO"
   Case Else
    typeExams = value
  End Select
End Function

'TODO Funcion: typeSex
'? Descripcion: Esta funcion recibe una cadena de texto "value" que representa un tipo de raza o etnia y devuelve una cadena de texto representando una version estandarizada del tipo de raza o etnia. Si el valor de entrada coincide con uno de los casos listados en la instruccion Select Case, se devuelve el tipo de raza o etnia estandarizado correspondiente. Si el valor de entrada no coincide con ninguno de los casos, se devuelve el valor de entrada original.
'? Parametros:
'? @param value: Cadena de texto que representa un tipo de raza o etnia.
'? Retorno:
'? @return Cadena de texto representando una version estandarizada del tipo de raza o etnia.
Public Function typeSex(ByVal value As String) As String
  Select Case value
   Case "COBRIZA", "COBRIZO"
    typeSex = Trim("COBRIZA")
   Case "NEGRA", "NEGRO", "PALENQUERO"
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

'TODO Funcion: typeCivil
'? Descripcion: Esta funcion recibe una cadena de texto "value" que representa un estado civil y devuelve una cadena de texto representando una version estandarizada del estado civil. Si el valor de entrada coincide con uno de los casos listados en la instruccion Select Case, se devuelve el estado civil estandarizado correspondiente. Si el valor de entrada no coincide con ninguno de los casos, se devuelve el valor de entrada original.
'? Parametros:
'? @param value: Cadena de texto que representa un estado civil.
'? Retorno:
'? @return Cadena de texto representando una version estandarizada del estado civil.
Public Function typeCivil(ByVal value As String) As String
  Select Case value
   Case "UNI" & Chr(211) & "N LIBRE"
    typeCivil = "UNION LIBRE"
   Case Else
    typeCivil = value
  End Select
End Function

'TODO Funcion: typeActivity
'? Descripcion: Esta funcion recibe una cadena de texto "value" que representa un actividad fisica y devuelve una cadena de texto representando una version estandarizada del actividad fisica. Si el valor de entrada coincide con uno de los casos listados en la instruccion Select Case, se devuelve el actividad fisica estandarizado correspondiente. Si el valor de entrada no coincide con ninguno de los casos, se devuelve el valor de entrada original.
'? Parametros:
'? @param value: Cadena de texto que representa un actividad fisica.
'? Retorno:
'? @return Cadena de texto representando una version estandarizada del actividad fisica.
Public Function typeActivity(ByVal value As String) As String
  Select Case value
   Case "F" & Chr(205) & "SICAMENTE ACTIVO", "FISICAMENTE ACTIVO", "FISICAMENTE ACTIVO(A)", "F" & Chr(205) & "SICAMENTE ACTIVO(A)"
    typeActivity = "F" & Chr(205) & "SICAMENTE ACTIVO"
   Case Else
    typeActivity = value
  End Select
End Function

'TODO Funcion: typeSmoke
'? Descripcion: Esta funcion toma un valor de cadena y devuelve una cadena que indica si el valor es un fumador, un exfumador o no fuma.
'? Parametros:
'? @param value: El valor de cadena que se evaluara para determinar si es un fumador, un exfumador o no fuma.
'? Retorno:
'? @return Una cadena que indica si el valor es un fumador, un exfumador o no fuma.
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

'TODO Funcion: correction
'? Descripcion: Esta funcion toma un valor de cadena y devuelve una cadena que indica si el valor esta corregido correctamente o no.
'? Parametros:
'? @param value: El valor de cadena que se evaluara para determinar si esta corregido correctamente o no.
'? Retorno:
'? @return Una cadena que indica si el valor esta corregido correctamente o no.
Public Function correction(ByVal value As String) As String
  Select Case value
   Case "ANORMAL SIN CORRECCION"
    correction = "ANORMAL MAL CORREGIDO"
   Case Else
    correction = value
  End Select
End Function

'TODO Funcion: typeComplements
'? Descripcion: Esta funcion toma un valor de cadena y devuelve una cadena que indica si el valor es una encuesta respiratoria o una valoracion respiratoria.
'? Parametros:
'? @param value: El valor de cadena que se evaluara para determinar si es una encuesta respiratoria o una valoracion respiratoria.
'? Retorno:
'? @return Una cadena que indica si el valor es una encuesta respiratoria o una valoracion respiratoria.
Public Function typeComplements(ByVal value As String) As String
  Select Case value
   Case "ENCUESTA RESPIRATORIA", "ENCUESTA DE SINTOMAS RESPIRATORIOS"
    typeComplements = "VALORACION RESPIRATORIA"
   Case Else
    typeComplements = value
  End Select
End Function

'TODO Funcion: total
'? Descripcion: Esta funcion toma un objeto de libro de Excel y cuenta el numero de filas en cada hoja de trabajo con un nombre especifico para calcular un total.
'? Parametros:
'? @param book: El objeto de libro de Excel que se utilizara para contar el numero de filas en cada hoja de trabajo.
'? Retorno:
'? @return Un numero entero que indica el total calculado.
Public Function total(ByVal book As Object) As LongPtr

  Dim emo As LongPtr, audio As LongPtr, opto As LongPtr, espiro As LongPtr, visio As LongPtr, complementarios As LongPtr, psicotecnica As LongPtr, psicosensometrica As LongPtr, osteo As LongPtr
  Dim Sheet As Object

  For Each Sheet In book.Worksheets

    Select Case Trim(UCase(Sheet.Name))
     Case "EMO"
      If Sheet.range("A2") <> "" And Sheet.range("A3") <> "" Then
        nameCompany = Sheet.range("A2").value
        formImports.Caption = CStr(nameCompany)
        emo = Sheet.range("A2", Sheet.range("A2").End(xlDown)).Count
      Else
        emo = 1
      End If
     Case "AUDIO"
      If Sheet.range("A2") <> "" And Sheet.range("A3") <> "" Then
        audio = Sheet.range("A2", Sheet.range("A2").End(xlDown)).Count
      Else
        audio = 1
      End If
     Case "OPTO"
      If Sheet.range("A2") <> "" And Sheet.range("A3") <> "" Then
        opto = Sheet.range("A2", Sheet.range("A2").End(xlDown)).Count
      Else
        opto = 1
      End If
     Case "VISIO"
      If Sheet.range("A2") <> "" And Sheet.range("A3") <> "" Then
        visio = Sheet.range("A2", Sheet.range("A2").End(xlDown)).Count
      Else
        visio = 1
      End If
     Case "ESPIRO"
      If Sheet.range("A2") <> "" And Sheet.range("A3") <> "" Then
        espiro = Sheet.range("A2", Sheet.range("A2").End(xlDown)).Count
      Else
        espiro = 1
      End If
     Case "OSTEO"
      If Sheet.range("A2") <> "" And Sheet.range("A3") <> "" Then
        osteo = Sheet.range("A2", Sheet.range("A2").End(xlDown)).Count
      Else
        osteo = 1
      End If
     Case "COMPLEMENTARIO", "COMPLEMENTARIOS"
      If Sheet.range("A2") <> "" And Sheet.range("A3") <> "" Then
        complementarios = Sheet.range("A2", Sheet.range("A2").End(xlDown)).Count
      Else
        complementarios = 1
      End If
     Case "PSICOTECNICA", "PSICOLOGIA"
      If Sheet.range("A2") <> "" And Sheet.range("A3") <> "" Then
        psicotecnica = Sheet.range("A2", Sheet.range("A2").End(xlDown)).Count
      Else
        psicotecnica = 1
      End If
     Case "PSICOSENSOMETRICA", "PSICOMOTRIZ"
      If Sheet.range("A2") <> "" And Sheet.range("A3") <> "" Then
        psicosensometrica = Sheet.range("A2", Sheet.range("A2").End(xlDown)).Count
      Else
        psicosensometrica = 1
      End If
    End Select
  Next Sheet

  total = (emo * 4) + audio + visio + espiro + osteo + complementarios + psicotecnica + psicosensometrica + opto

End Function

'TODO: Esta macro elimina los caracteres no alfanumericos de una columna
Public Sub ClearNonAlphaNumeric()
  Dim valor As String
  Dim data As Variant
  Dim i As Long, j As Long

  '? Desactivar la actualizacion de pantalla, el calculo y los eventos
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  '? Leer los valores de la seleccion en una matriz
  data = Selection.value

  '? Iterar a traves de la matriz y realizar los reemplazos
  On Error Resume Next
  For i = 1 To UBound(data, 1)
    For j = 1 To UBound(data, 2)
      valor = VBA.Replace(data(i, j), "  ", " ", , , vbTextCompare)
      data(i, j) = Trim(ReplaceNonAlphaNumeric(valor))
    Next j
    DoEvents
  Next i
  On Error GoTo 0

  '? Escribir la matriz de vuelta en la hoja de calculo
  Selection.value = data

  '? Activar la actualizacion de pantalla, el calculo y los eventos
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

  formClear.Hide
  MsgBox "Correcciones realizadas, exitosamente!!", vbInformation, "Correcciones"
End Sub

'TODO: Esta funcion reemplaza los caracteres no alfanumericos y las letras con acentos en una cadena de texto
Public Function ReplaceNonAlphaNumeric(str As String) As String

  Dim regEx As Object, letter As String, accent As Variant, accentPairs As Variant

  Set regEx = CreateObject("vbscript.regexp")
  accentPairs = Array(ChrW(192) & ",A", ChrW(200) & ",E", ChrW(204) & ",I", ChrW(210) & ",O", ChrW(217) & ",U", ChrW(193) & ",A", ChrW(201) & ",E", ChrW(205) & ",I", ChrW(211) & ",O", ChrW(218) & ",U")

  '? Recorre el array de pares de acentos y letras, aplicando las expresiones regulares correspondientes
  For Each accent In accentPairs
    letter = Split(accent, ",")(0)
    regEx.Pattern = "[" & letter & ChrW(AscW(letter) + 1) & "]"
    regEx.Global = True
    str = regEx.Replace(str, Split(accent, ",")(1))
    DoEvents
  Next accent

  '? Define la expresion regular para encontrar valores no alfanumericos
  regEx.Pattern = "[^a-zA-Z0-9/" & ChrW(209) & ChrW(45) & "]"
  regEx.Global = True

  '? Reemplaza cualquier valor no alfanumerico por un espacio
  ReplaceNonAlphaNumeric = regEx.Replace(str, " ")
End Function

'TODO: Este Subrutina asigna un numero aleatorio entre 60 y 80 a las celdas vacias en la columna activa, siempre y cuando el valor de la celda no sea "SIN DATO".
'?Variables:
'? @param num: longptr - Almacena el numero aleatorio generado.
'?Instrucciones:
'?   1. Inicio del bucle hasta que la celda activa en la columna anterior este vacia.
'?   2. Si la celda activa esta vacia o el valor en mayusculas es "SIN DATO", entonces genera un numero aleatorio entre 60 y 80 y lo asigna a la celda activa.
'?   3. Selecciona la celda siguiente en la columna activa.
'?   4. Fin del bucle.
Public Sub Peso()
  Dim num As LongPtr

  Do While Not IsEmpty(ActiveCell.Offset(, -35))
    If IsEmpty(ActiveCell) Or Trim$(UCase$(ActiveCell.value)) = "SIN DATO" Then
      num = Int((80 - 60 + 1) * Rnd + 60)
      ActiveCell.value = num
    End If
    ActiveCell.Offset(1, 0).Select
  Loop

End Sub

'TODO: ajustarTallas ajusta la informacion de altura en la columna activa.
'? Las celdas vacias o con el valor "sin dato" se reemplazan con una altura aleatoria entre 1.6 y 1.8 metros.
'? Las celdas que contienen un numero entero se dividen por 100 para convertirlos a metros.
'? Este subrutina continua hasta que se encuentra una celda vacia en la columna -36.
Public Sub ajustarTallas()
  Dim talla As Double

  '? Recorre todas las celdas hacia abajo hasta encontrar una vacia en la columna -36
  Do While Not IsEmpty(ActiveCell.Offset(0, -36))
    If Trim$(ActiveCell.value) = "" Or Trim$(UCase$(ActiveCell.value)) = "SIN DATO" Then
      '? Genera una talla aleatoria entre 1.6 y 1.8 metros
      talla = CDec((Int((180 - 160 + 1) * Rnd + 160)) / 100)
      ActiveCell.value = talla
    ElseIf ActiveCell.value = Int(ActiveCell.value) Then
      '? Divide el numero entero de la celda por 100
      talla = CDec(ActiveCell.value / 100)
      ActiveCell.value = talla
    End If

    '? Selecciona la celda siguiente
    ActiveCell.Offset(1, 0).Select
  Loop
End Sub
