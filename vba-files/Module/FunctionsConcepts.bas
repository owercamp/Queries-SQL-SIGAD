Attribute VB_Name = "FunctionsConcepts"

Public Function validateConcepts(value)

    Select Case Trim(UCase(value))
        '' CONCEPTO DE EVALUACION: PUEDE CONTINUAR REALIZANDO SU LABOR ''
     Case "PUEDE CONTINUAR LABORANDO", "PUEDE CONTINUAR REALIANZADO LA LABOR", "PUEDE CONTINUAR REALIZANDO LA  LABOR","PUEDE CONTINUAR REALIZANDO SU LABOR", "PUEDE CONTINUAR REALIZANDO LA LABORANDO", "PUEDE CONTINUAR DESEMPE" & Chr(209) & "ANDO SU LABOR","CON PATOLOGIA QUE NO LIMITA LA LABOR","PUEDE CONTINUAR REALIZANDO  LA LABOR"
        validateConcepts = Trim("PUEDE CONTINUAR REALIZANDO LA LABOR")

        '' CONCEPTO DE EVALUACION: CUMPLE PARA DESEMPE"&Chr(209)&"AR EL CARGO ''
     Case "CUMPLE PARA DESEMPE" & Chr(209) & "AR EL CARGO", "APTO PARA DESEMPE" & Chr(209) & "AR CARGO", "CUMPLE PARA DESEMPE" & Chr(209) & "AR  CARGO","APTO PARA DESEMPE" & Chr(209) & "AR LABOR", "APTO PARA DESEMPE" & Chr(209) & "AR EL CARGO", "APTO PARA DESEMPE" & Chr(209) & "AR EL  CARGO","APTO (SIN PATOLOG" & Chr(205) & "AS EVIDENTES; CUMPLE CON LOS CRITERIOS M" & Chr(201) & "DICOS PARA EL CARGO)", "CUMPLE PARA REALIZAR EL CARGO", "APTO PARA EL CARGO SIN PATOLOGIA APARENTE"
        validateConcepts = Trim("CUMPLE PARA DESEMPE" & Chr(209) & "AR EL CARGO")

        '' CONCEPTO DE EVALUACION: SIN RESTRICCIONES LABORALES PARA EL CARGO ''
     Case "SIN RESTRICCIONES LABORALES PARA EL CARGO", "SIN RESTRICCIONES PARA DESEMPE" & Chr(209) & "AR LA LABOR"
        validateConcepts = Trim("SIN RESTRICCIONES LABORALES PARA EL CARGO")

        '' CONCEPTO DE EVALUACION: APLAZADO ''
     Case "APLAZADO", "APLAZADA","NO CUMPLE"
        validateConcepts = Trim("APLAZADO")

        '' CONCEPTO DE EVALUACION: SIN RESTRICCIONES PARA DESEMPE"&Chr(209)&"ARSE EN EL NUEVO CARGO ''
     Case "SIN RESTRICCIONES PARA DESEMPE" & Chr(209) & "ARSE EN EL NUEVO CARGO", "SIN RESTRICCIONES PARA DESEMPE" & Chr(209) & "ARSE EL CARGO"
        validateConcepts = Trim("SIN RESTRICCIONES PARA DESEMPE" & Chr(209) & "ARSE EN EL NUEVO CARGO")

        '' CONCEPTO DE EVALUACION: REALIZADO ''
     Case "REALIZADO", "REALIZADA"
        validateConcepts = Trim("REALIZADO")

        '' CONCEPTO DE EVALUACION: REASIGNACI"&Chr(211)&"N DE FUNCIONES Y TAREAS ''
     Case "REASIGNACION DE TAREAS", "REASIGNACI" & Chr(211) & "N DE TAREAS", "REASIGNACI" & Chr(211) & "N DE FUNCIONES Y TAREAS", "REASIGNACI" & Chr(211) & "N","REASIGNACION"
        validateConcepts = Trim("REASIGNACION DE TAREAS")

        '' CONCEPTO DE EVALUACION: REINCORPORACION ''
     Case "REINCORPORACION", "REINCORPORACI" & Chr(211) & "N", "REINCORPORACI" & Chr(211) & "N AL PUESTO DE TRABAJO","REINTEGRO AL CARGO ACTUAL CON RECOMENDACIONES Y RESTRICCIONES", "REINTEGRO TENIENDO EN CUENTA RECOMENDACIONES ESPECIFICAS", "REINCORPORACION AL PUESTO DE TRABAJO"
        validateConcepts = Trim("REINCORPORACION")

        '' CONCEPTO DE EVALUACION: PRESENTA RESTRICCION ''
     Case "PRESENTA RESTRICCION", "PRESENTA RESTRICCI" & Chr(211) & "N", "APTO CON LIMITACI" & Chr(211) & "N O RESTRICCI" & Chr(211) & "N QUE S" & Chr(205) & " INTERFIERE PARA EL CARGO", "CON RESTRICCION PARA EL CARGO"
        validateConcepts = Trim("PRESENTA RESTRICCION")

        '' CONCEPTO DE EVALUACION: APTO CON RECOMENDACION ''
     Case "APTO CON RECOMENDACION", "APTO CON RESTRICCION", "CON PATOLOGIA QUE NO LIMITA LA LABOR","APTO CON DEFECTO F" & Chr(205) & "SICO O PATOLOGIA QUE NO LIMITA SU LABOR", "REINTEGRO CON RECOMENDACIONES","REINTEGRO CON MODIFICACIONES/RESTRICCIONES", "PRESENTA RESTRICCIONES", "REINTEGRO CON RESTRICCIONES", "APTO CON RESTRICCIONES", "APTO CON PATOLOG" & Chr(205) & "AS (QUE NO LIMITAN SU CAPACIDAD LABORAL)", "APTO CON RECOMENDACIONES", "EXAMEN PERIODICO CON RECOMENDACIONES","CUMPLE PARA DESEMPEÑAR EL CARGO CON RECOMENDACIONES","PUEDE CONTINUAR REALIZANDO LA LABOR CON RECOMENDACIONES"
        validateConcepts = Trim("APTO CON RECOMENDACION")

     Case Else
        validateConcepts = Trim(UCase(value))
    End Select

End Function

Public Function emphasisConcepts(value, emphasis)

    Dim No As Integer
    Dim status, activity As String

    No = VBA.InStr(Trim(UCase(value)), "NO")

    '' VERIFICA SI VIENE LA PALABRA APTO O CUMPLE EN EL CONCEPTO ''
    If VBA.InStr(Trim(UCase(value)), "APTO") > 0 Then

        status = "APTO"

    ElseIf VBA.InStr(Trim(UCase(value)), "CUMPLE") > 0 Then

        status = "CUMPLE"

    ElseIf VBA.InStr(Trim(UCase(value)), "APLAZADO") > 0 Then

        status = "APLAZADO"

    End If

    '' SEPARA LA ACTIVIDAD CORRESPONDIENTE ''
    If VBA.InStr(Trim(UCase(value)), "ESPACIOS CONFINADOS") > 0 Or VBA.InStr(Trim(UCase(value)), "ESPACIO CONFINADO") > 0 Then

        activity = "ESPACIOS CONFINADOS"

    ElseIf VBA.InStr(Trim(UCase(value)), "SEGURIDAD VIAL") > 0 Then

        activity = "SEGURIDAD VIAL"

    ElseIf VBA.InStr(Trim(UCase(value)), "BRIGADISTA") > 0 Then

        activity = "BRIGADISTA"

    ElseIf VBA.InStr(Trim(UCase(value)), "ACTIVIDADES DEPORTIVAS") > 0 Or VBA.InStr(Trim(UCase(value)), "ACTIVIDAD DEPORTIVA") > 0 Then

        activity = "ACTIVIDAD DEPORTIVA"

    ElseIf VBA.InStr(Trim(UCase(value)), "ALTURA") > 0 Or VBA.InStr(Trim(UCase(value)), "ALTURAS") > 0 Then

        activity = "ALTURA"

    ElseIf VBA.InStr(Trim(UCase(value)), "ACTIVIDADES DEPORTIVAS") > 0 Or VBA.InStr(Trim(UCase(value)), "ACTIVIDAD DEPORTIVA") > 0 Then

        activity = "ACTIVIDAD DEPORTIVA"

    ElseIf VBA.InStr(Trim(UCase(value)), "ALIMENTOS") > 0 Or VBA.InStr(Trim(UCase(value)), "ALIMENTO") > 0 Then

        activity = "ALIMENTOS"

    ElseIf VBA.InStr(Trim(UCase(value)), "MEDICAMENTOS") > 0 Or VBA.InStr(Trim(UCase(value)), "MEDICAMENTO") > 0 Then

        activity = "MEDICAMENTO"

    ElseIf VBA.InStr(Trim(UCase(value)), "QUIMICOS") > 0 Or VBA.InStr(Trim(UCase(value)), "QUIMICO") > 0 Then

        activity = "QUIMICOS"

    ElseIf VBA.InStr(Trim(UCase(value)), "ALTA TENSION") > 0 Or VBA.InStr(Trim(UCase(value)), "ALTAS TENSIONES") > 0 Then

        activity = "ALTA TENSION"

    ElseIf VBA.InStr(Trim(UCase(value)), "ACTIVIDADES DEPORTIVAS") > 0 Or VBA.InStr(Trim(UCase(value)), "ACTIVIDAD DEPORTIVA") > 0 Then

        activity = "ACTIVIDAD DEPORTIVA"

    ElseIf VBA.InStr(Trim(UCase(value)), "BAJA") > 0 Or VBA.InStr(Trim(UCase(value)), "BAJAS") > 0 Then

        activity = "BAJAS"

    ElseIf VBA.InStr(Trim(UCase(value)), "ALTA") > 0 Or VBA.InStr(Trim(UCase(value)), "ALTAS") > 0 Then

        activity = "ALTAS"

    ElseIf VBA.InStr(Trim(UCase(value)), "NIVEL DEL MAR") > 0 Then

        activity = "NIVEL DEL MAR"

    ElseIf VBA.InStr(Trim(UCase(value)), "HIPERBARICO") > 0 Or VBA.InStr(Trim(UCase(value)), "HIPERBARICOS") > 0 Then

        activity = "HIPERBARICOS"

    End If

    '' CONCEPTOS AL ENFASIS DE ESPACIOS CONFINADOS ''
    If No = 0 And (status = "CUMPLE" Or status = "APTO") And (activity = "ESPACIOS CONFINADOS" Or activity = "" Or activity = Empty) And _
        (Trim(UCase(emphasis)) = "ESPACIOS CONFINADOS" Or Trim(UCase(emphasis)) = "ESPACIO CONFINADO") Then

        emphasisConcepts = "Apto para trabajo en espacios confinados"

    ElseIf No <> 0 And (status = "CUMPLE" Or status = "APTO") And (activity = "ESPACIOS CONFINADOS" Or activity = "" Or activity = Empty) And _
        (Trim(UCase(emphasis)) = "ESPACIOS CONFINADOS" Or Trim(UCase(emphasis)) = "ESPACIO CONFINADO") Then

        emphasisConcepts = "No cumple para trabajar en espacios confinados"

    ElseIf No = 0 And status = "APLAZADO" And (activity = "ESPACIOS CONFINADOS" Or activity = "" Or activity = Empty) And _
        (Trim(UCase(emphasis)) = "ESPACIOS CONFINADOS" Or Trim(UCase(emphasis)) = "ESPACIO CONFINADO") Then

        emphasisConcepts = "Aplazado para trabajar en espacios confinados"

        '' CONCEPTOS AL ENFASIS DE SEGURIDAD VIAL ''
    ElseIf No = 0 And (status = "CUMPLE" Or status = "APTO") And (activity = "SEGURIDAD VIAL" Or activity = "" Or activity = Empty) And Trim(UCase(emphasis)) = "SEGURIDAD VIAL" Then

        emphasisConcepts = "Apto para seguridad vial"

    ElseIf (No = 0 And status = "APLAZADO" And (activity = "SEGURIDAD VIAL" Or activity = "" Or actibity = Empty) And Trim(UCase(emphasis)) = "SEGURIDAD VIAL") Or (No <> 0 And (status = "CUMPLE" Or status = "APTO") And (activity = "SEGURIDAD VIAL" Or activity = "" Or activity = Empty) And Trim(UCase(emphasis)) = "SEGURIDAD VIAL") Then

        emphasisConcepts = "Aplazado para seguridad vial"

        '' CONCEPTOS AL ENFASIS DE BRIGADISTA ''
    ElseIf No = 0 And (status = "CUMPLE" Or status = "APTO") And (activity = "BRIGADISTA" Or activity = "" Or activity = Empty) And Trim(UCase(emphasis)) = "BRIGADISTA" Then

        emphasisConcepts = "Apto para brigadista"

        '' CONCEPTOS AL ENFASIS DE ACTIVIDADES DEPORTIVAS ''
    ElseIf No = 0 And (status = "CUMPLE" Or status = "APTO") And (activity = "ACTIVIDAD DEPORTIVA" Or activity = "" Or activity = Empty) And _
        (Trim(UCase(emphasis)) = "ACTIVIDAD DEPORTIVA" Or Trim(UCase(emphasis)) = "ACTIVIDADES DEPORTIVAS") Then

        emphasisConcepts = "Apto para actividad deportiva"

    ElseIf No = 0 And status = "APLAZADO" And (activity = "ACTIVIDAD DEPORTIVA" Or activity = "" Or activity = Empty) And _
        (Trim(UCase(emphasis)) = "ACTIVIDAD DEPORTIVA" Or Trim(UCase(emphasis)) = "ACTIVIDADES DEPORTIVAS") Then

        emphasisConcepts = "Aplazado para actividad deportiva"

    ElseIf No <> 0 And (status = "CUMPLE" Or status = "APTO") And (activity = "ACTIVIDAD DEPORTIVA" Or activity = "" Or activity = Empty) And _
        (Trim(UCase(emphasis)) = "ACTIVIDAD DEPORTIVA" Or Trim(UCase(emphasis)) = "ACTIVIDADES DEPORTIVAS") Then

        emphasisConcepts = "No cumple para actividad deportiva"

        '' CONCEPTOS AL ENFASIS DE ALTURAS ''
    ElseIf No = 0 And (status = "CUMPLE" Or status = "APTO") And (activity = "ALTURA" Or activity = "" Or activity = Empty) And _
        (Trim(UCase(emphasis)) = "ALTURA" Or Trim(UCase(emphasis)) = "ALTURAS") Then

        emphasisConcepts = "Apto para trabajo en alturas"

    ElseIf No <> 0 And (status = "CUMPLE" Or status = "APTO") And (activity = "ALTURA" Or activity = "" Or activity = Empty) And _
        (Trim(UCase(emphasis)) = "ALTURA" Or Trim(UCase(emphasis)) = "ALTURAS") Then

        emphasisConcepts = "No cumple para trabajar en alturas"

    ElseIf (No = 0 And status = "APLAZADO" And (activity = "ALTURA" Or activity = "" Or activity = Empty) And (Trim(UCase(emphasis)) = "ALTURAS" Or Trim(UCase(emphasis)) = "ALTURA")) Or _
        (No <> 0 And status = "APLAZADO" And (activity = "ALTURA" Or activity = "" Or activity = Empty) And (Trim(UCase(emphasis)) = "ALTURAS" Or Trim(UCase(emphasis)) = "ALTURA")) Then

        emphasisConcepts = "Aplazado para trabajar en alturas"

        '' CONCEPTOS AL ENFASIS DE ALIMENTOS ''
    ElseIf No = 0 And (status = "CUMPLE" Or status = "APTO") And (activity = "ALIMENTOS" Or activity = "" Or activity = Empty) And _
        (Trim(UCase(emphasis)) = "ALIMENTO" Or Trim(UCase(emphasis)) = "ALIMENTOS") Then

        emphasisConcepts = "Apto para manipular alimentos"

    ElseIf (No = 0 And status = "APLAZADO" And (activity = "ALIMENTOS" Or activity = "" Or activity = Empty) And (Trim(UCase(emphasis)) = "ALIMENTOS" Or Trim(UCase(emphasis)) = "ALIMENTO")) Or _
        (No <> 0 And status = "APLAZADO" And (activity = "ALIMENTOS" Or activity = "" Or activity = Empty) And (Trim(UCase(emphasis)) = "ALIMENTOS" Or Trim(UCase(emphasis)) = "ALIMENTO")) Then

        emphasisConcepts = "Aplazado"

        '' CONCEPTOS AL ENFASIS DE MEDICAMENTOS ''
    ElseIf No = 0 And (status = "CUMPLE" Or status = "APTO") And (activity = "MEDICAMENTO" Or activity = "" Or activity = Empty) And _
        (Trim(UCase(emphasis)) = "MEDICAMENTO" Or Trim(UCase(emphasis)) = "MEDICAMENTOS") Then

        emphasisConcepts = "Apto para manipular medicamentos"

        '' CONCEPTOS AL ENFASIS QUIMICOS ''
    ElseIf Trim(UCase(value)) = UCase("Apto") Then

        emphasisConcepts = "Apto"

        '' CONCEPTOS AL ENFASIS PARA ALTA TENSION, TEMPERATURAS ALTAS - BAJAS, NIVEL DEL MAR, HIPERBARICOS, CARDIOVASCULAR Y IONIZANTES ''
    ElseIf Trim(UCase(value)) = UCase("Cumple") Then

        emphasisConcepts = "Cumple"

        '' CONCEPTO MULTIPLE APLAZADO PARA LOS CONCEPTOS BRIGADISTA, ALIMENTOS, MEDICAMENTOS, QUIMICOS, ''
        '' ALTA TENSION, TEMPERATURAS ALTAS - BAJAS, NIVEL DEL MAR, HIPERBARICOS, CARDIOVASCULAR, DERMATOLOGICO Y IONIZANTES ''
    ElseIf Trim(UCase(value)) = UCase("Aplazado") Then

        emphasisConcepts = "Aplazado"

        '' CONCEPTO MULTIPLE NO CUMPLE PARA TODOS LOS CONCEPTOS ''
    ElseIf Trim(UCase(value)) = UCase("No cumple") Then

        emphasisConcepts = "No cumple"

    Else

        emphasisConcepts = value

    End If

End Function
