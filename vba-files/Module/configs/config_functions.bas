Attribute VB_Name = "config_functions"
'namespace=vba-files\Module\configs
Option Explicit

Public Sub insertFunctions()

  Dim book As Workbook, tbl_Formulas As Object, only As Workbook, Sheet As Variant
  
  Set only = ThisWorkbook
  Set book = Workbooks.Open(ThisWorkbook.Worksheets("RUTAS").range("$C$7").value)
  book.Worksheets("Funciones").Select
  Set tbl_Formulas = book.Worksheets("Funciones").ListObjects("tbl_formulas").DataBodyRange
  
  Windows(only.Name).Activate
  
  For Each Sheet In only.Worksheets
    Call assignFuncions(Sheet, tbl_Formulas)
  Next Sheet
  
  book.Save
  book.Close
  
End Sub

Private Sub assignFuncions(ByVal Sheet As Worksheet, ByVal formulas As Object)

  Dim counter As Integer
  
  Select Case Trim(Sheet.Name)
   Case "TRABAJADORES"
    Worksheets(Sheet.Name).Select
    range("tbl_trabajadores[[#Headers],[LLAVE]]").Offset(1, 0) = formulas(1, 2)
    range("tbl_trabajadores[[#Headers],[rango_edad]]").Offset(1, 0) = formulas(2, 2)
    range("tbl_trabajadores[[#Headers],[hijos]]").Offset(1, 0) = formulas(3, 2)
    range("tbl_trabajadores[[#Headers],[CARGO_REC]]").Offset(1, 0) = formulas(4, 2)
    range("tbl_trabajadores[[#Headers],[ANTIGUEDAD]]").Offset(1, 0) = formulas(5, 2)
    range("tbl_trabajadores[[#Headers],[CIUDAD_ID]]").Offset(1, 0) = formulas(6, 2)
    range("tbl_trabajadores[[#Headers],[id_tipo_examen]]").Offset(1, 0) = formulas(7, 2)
    range("tbl_trabajadores[[#Headers],[fecha_texto]]").Offset(1, 0) = formulas(8, 2)
    range("tbl_trabajadores[[#Headers],[id_raza]]").Offset(1, 0) = formulas(9, 2)
    range("tbl_trabajadores[[#Headers],[id_estado_civil]]").Offset(1, 0) = formulas(10, 2)
    range("tbl_trabajadores[[#Headers],[id_escolaridad]]").Offset(1, 0) = formulas(11, 2)
    range("tbl_trabajadores[[#Headers],[id_cargo]]").Offset(1, 0) = formulas(12, 2)
    range("tbl_trabajadores[[#Headers],[fuente2]]").Offset(1, 0) = formulas(13, 2)
    range("tbl_trabajadores[[#Headers],[(id_tipo_actividad)]]").Offset(1, 0) = formulas(14, 2)
    range("tbl_trabajadores[[#Headers],[AUDIO]]").Offset(1, 0) = formulas(15, 2)
    range("tbl_trabajadores[[#Headers],[OPTO]]").Offset(1, 0) = formulas(16, 2)
    range("tbl_trabajadores[[#Headers],[ESPIRO]]").Offset(1, 0) = formulas(17, 2)
    range("tbl_trabajadores[[#Headers],[VISIO]]").Offset(1, 0) = formulas(18, 2)
    range("tbl_trabajadores[[#Headers],[OSTEO]]").Offset(1, 0) = formulas(19, 2)
    range("tbl_trabajadores[[#Headers],[PSICOSENSOMETRICA]]").Offset(1, 0) = formulas(20, 2)
    range("tbl_trabajadores[[#Headers],[PSICOTECNICA]]").Offset(1, 0) = formulas(21, 2)
    range("tbl_trabajadores[[#Headers],[COMPLEMENTARIOS]]").Offset(1, 0) = formulas(22, 2)
    range("tbl_trabajadores[[#Headers],[EMO]]").Offset(1, 0) = formulas(23, 2)
    range("tbl_trabajadores[[#Headers],[SCRIPT orden_lista_trabajadores]]").Offset(1, 0) = formulas(24, 2)
    range("tbl_trabajadores[[#Headers],[SCRIPT ordenes_trabajador_paraclinicos]]").Offset(1, 0) = formulas(25, 2)
   Case "EMO"
    Worksheets(Sheet.Name).Select
    range("tbl_emo[[#Headers],[IMC2]]").Offset(1, 0) = formulas(26, 2)
    range("tbl_emo[[#Headers],[CLASIFICACION_IMC]]").Offset(1, 0) = formulas(27, 2)
    range("tbl_emo[[#Headers],[CODIGO DIAG PPAL]]").Offset(1, 0) = formulas(28, 2)
    range("tbl_emo[[#Headers],[DIAG PPAL]]").Offset(1, 0) = formulas(29, 2)
    range("tbl_emo[[#Headers],[CODIGO DIAG REL1]]").Offset(1, 0) = formulas(30, 2)
    range("tbl_emo[[#Headers],[DIAG REL 1]]").Offset(1, 0) = formulas(31, 2)
    range("tbl_emo[[#Headers],[CODIGO DIAG REL2]]").Offset(1, 0) = formulas(32, 2)
    range("tbl_emo[[#Headers],[DIAG REL 2]]").Offset(1, 0) = formulas(33, 2)
    range("tbl_emo[[#Headers],[CODIGO DIAG REL3]]").Offset(1, 0) = formulas(34, 2)
    range("tbl_emo[[#Headers],[DIAG REL 3]]").Offset(1, 0) = formulas(35, 2)
    range("tbl_emo[[#Headers],[CODIGO DIAG REL4]]").Offset(1, 0) = formulas(36, 2)
    range("tbl_emo[[#Headers],[DIAG REL 4]]").Offset(1, 0) = formulas(37, 2)
    range("tbl_emo[[#Headers],[CODIGO DIAG REL5]]").Offset(1, 0) = formulas(38, 2)
    range("tbl_emo[[#Headers],[DIAG REL 5]]").Offset(1, 0) = formulas(39, 2)
    range("tbl_emo[[#Headers],[CODIGO DIAG REL6]]").Offset(1, 0) = formulas(40, 2)
    range("tbl_emo[[#Headers],[DIAG REL 6]]").Offset(1, 0) = formulas(41, 2)
    range("tbl_emo[[#Headers],[CODIGO DIAG REL7]]").Offset(1, 0) = formulas(42, 2)
    range("tbl_emo[[#Headers],[DIAG REL 7]]").Offset(1, 0) = formulas(43, 2)
    range("tbl_emo[[#Headers],[CODIGO DIAG REL8]]").Offset(1, 0) = formulas(44, 2)
    range("tbl_emo[[#Headers],[DIAG REL 8]]").Offset(1, 0) = formulas(45, 2)
    range("tbl_emo[[#Headers],[CODIGO DIAG REL9]]").Offset(1, 0) = formulas(46, 2)
    range("tbl_emo[[#Headers],[DIAG REL 9]]").Offset(1, 0) = formulas(47, 2)
    range("tbl_emo[[#Headers],[ENFASIS_1]]").Offset(1, 0) = formulas(48, 2)
    range("tbl_emo[[#Headers],[CONCEPTO AL ENFASIS_1]]").Offset(1, 0) = formulas(49, 2)
    range("tbl_emo[[#Headers],[OBSERVACIONES_AL_ENFASIS_1]]").Offset(1, 0) = formulas(50, 2)
    range("tbl_emo[[#Headers],[ENFASIS_2]]").Offset(1, 0) = formulas(51, 2)
    range("tbl_emo[[#Headers],[CONCEPTO AL ENFASIS_2]]").Offset(1, 0) = formulas(52, 2)
    range("tbl_emo[[#Headers],[OBSERVACIONES AL ENFASIS_2]]").Offset(1, 0) = formulas(53, 2)
    range("tbl_emo[[#Headers],[ENFASIS_3]]").Offset(1, 0) = formulas(54, 2)
    range("tbl_emo[[#Headers],[CONCEPTO AL ENFASIS_3]]").Offset(1, 0) = formulas(55, 2)
    range("tbl_emo[[#Headers],[OBSERVACIONES AL ENFASIS_3]]").Offset(1, 0) = formulas(56, 2)
    range("tbl_emo[[#Headers],[ENFASIS_4]]").Offset(1, 0) = formulas(57, 2)
    range("tbl_emo[[#Headers],[CONCEPTO AL ENFASIS_4]]").Offset(1, 0) = formulas(58, 2)
    range("tbl_emo[[#Headers],[OBSERVACIONES AL ENFASIS_4]]").Offset(1, 0) = formulas(59, 2)
    range("tbl_emo[[#Headers],[ACCIDENTE SI NO]]").Offset(1, 0) = formulas(60, 2)
    range("tbl_emo[[#Headers],[ics_emofecha_accidente]]").Offset(1, 0) = formulas(61, 2)
    range("tbl_emo[[#Headers],[ics_emonaturaleza_lesion_id]]").Offset(1, 0) = formulas(62, 2)
    range("tbl_emo[[#Headers],[ics_emoparte_afectada]]").Offset(1, 0) = formulas(63, 2)
    range("tbl_emo[[#Headers],[ics_emoenfermedad_laboral]]").Offset(1, 0) = formulas(64, 2)
    range("tbl_emo[[#Headers],[ics_emoetapa]]").Offset(1, 0) = formulas(65, 2)
    range("tbl_emo[[#Headers],[ics_emoid_concepto_evaluacion]]").Offset(1, 0) = formulas(66, 2)
    range("tbl_emo[[#Headers],[orden_lista_trabajadoresid]]").Offset(1, 0) = formulas(67, 2)
    range("tbl_emo[[#Headers],[EMPRESA]]").Offset(1, 0) = formulas(68, 2)
    range("tbl_emo[[#Headers],[SCRIPT ics_emo]]").Offset(1, 0) = formulas(69, 2)
    range("tbl_emo[[#Headers],[SCRIPT ics_emo_riesgos]]").Offset(1, 0) = formulas(70, 2)
    range("tbl_emo[[#Headers],[SCRIPT ics_condiciones]]").Offset(1, 0) = formulas(71, 2)
    range("tbl_emo[[#Headers],[SCRIPT ics_cie (diagnosticos)]]").Offset(1, 0) = formulas(72, 2)
    range("tbl_emo[[#Headers],[SCRIPT ics_enfasis]]").Offset(1, 0) = formulas(73, 2)
    range("tbl_emo[[#Headers],[SCRIPT ics_recomendacion_general]]").Offset(1, 0) = formulas(74, 2)
    range("tbl_emo[[#Headers],[SCRIPT ics_recomendacion_ocupacional]]").Offset(1, 0) = formulas(75, 2)
    range("tbl_emo[[#Headers],[LLAVE]]").Offset(1, 0) = formulas(76, 2)
   Case "AUDIO"
    Worksheets(Sheet.Name).Select
    range("tbl_audio[[#Headers],[PTA OD]]").Offset(1, 0) = formulas(77, 2)
    range("tbl_audio[[#Headers],[PTA OI]]").Offset(1, 0) = formulas(78, 2)
    range("tbl_audio[[#Headers],[tipo1 au_oido]]").Offset(1, 0) = formulas(79, 2)
    range("tbl_audio[[#Headers],[tipo2 au_oido]]").Offset(1, 0) = formulas(80, 2)
    range("tbl_audio[[#Headers],[frecuencia OD]]").Offset(1, 0) = formulas(81, 2)
    range("tbl_audio[[#Headers],[frecuencia OI]]").Offset(1, 0) = formulas(82, 2)
    range("tbl_audio[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Offset(1, 0) = formulas(83, 2)
    range("tbl_audio[[#Headers],[SCRIPT au_audiometria]]").Offset(1, 0) = formulas(84, 2)
    range("tbl_audio[[#Headers],[SCRIPT au_audiometria_recomendacion]]").Offset(1, 0) = formulas(85, 2)
    range("tbl_audio[[#Headers],[SCRIPT au_oido]]").Offset(1, 0) = formulas(86, 2)
    range("tbl_audio[[#Headers],[LLAVE]]").Offset(1, 0) = formulas(87, 2)
   Case "VISIO"
    Worksheets(Sheet.Name).Select
    range("tbl_visio[[#Headers],[0]]").Offset(1, 0) = formulas(88, 2)
    range("tbl_visio[[#Headers],[RESULTADO VISIO]]").Offset(1, 0) = formulas(89, 2)
    range("tbl_visio[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Offset(1, 0) = formulas(90, 2)
    range("tbl_visio[[#Headers],[SCRIPT vi_visiometria]]").Offset(1, 0) = formulas(91, 2)
    range("tbl_visio[[#Headers],[SCRIPT vi_visiometria_antecedentes]]").Offset(1, 0) = formulas(92, 2)
    range("tbl_visio[[#Headers],[SCRIPT vi_visiometria_sintomas]]").Offset(1, 0) = formulas(93, 2)
    range("tbl_visio[[#Headers],[SCRIPT vi_vl]]").Offset(1, 0) = formulas(94, 2)
    range("tbl_visio[[#Headers],[SCRIPT vi_vp]]").Offset(1, 0) = formulas(95, 2)
    range("tbl_visio[[#Headers],[SCRIPT vi_visiometria_recomendaciones]]").Offset(1, 0) = formulas(96, 2)
    range("tbl_visio[[#Headers],[SCRIPT vi_visiometria_remisiones]]").Offset(1, 0) = formulas(97, 2)
   Case "OPTO"
    Worksheets(Sheet.Name).Select
    range("tbl_opto[[#Headers],[CODIGO CIE10 DIAG PPAL]]").Offset(1, 0) = formulas(98, 2)
    range("tbl_opto[[#Headers],[estado_correccion_id]]").Offset(1, 0) = formulas(99, 2)
    range("tbl_opto[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Offset(1, 0) = formulas(100, 2)
    range("tbl_opto[[#Headers],[SCRIPT op_optometria]]").Offset(1, 0) = formulas(101, 2)
    range("tbl_opto[[#Headers],[SCRIPT op_optometria_riesgos]]").Offset(1, 0) = formulas(102, 2)
    range("tbl_opto[[#Headers],[SCRIPT op_optometria_sintomas]]").Offset(1, 0) = formulas(103, 2)
    range("tbl_opto[[#Headers],[SCRIPT op_diagnostico]]").Offset(1, 0) = formulas(104, 2)
    range("tbl_opto[[#Headers],[SCRIPT op_diagnostico_cie]]").Offset(1, 0) = formulas(105, 2)
    range("tbl_opto[[#Headers],[SCRIPT op_optometria_recomendacion]]").Offset(1, 0) = formulas(106, 2)
    range("tbl_opto[[#Headers],[SCRIPT op_optometria_remision]]").Offset(1, 0) = formulas(107, 2)
   Case "ESPIRO"
    Worksheets(Sheet.Name).Select
    range("tbl_espiro_info[[#Headers],[IMC2 (imc2)]]").Offset(1, 0) = formulas(108, 2)
    range("tbl_espiro_info[[#Headers],[CLASIFICACION_IMC (clasificacion_imc)]]").Offset(1, 0) = formulas(109, 2)
    range("tbl_espiro_info[[#Headers],[RESULTADO_ESPIROMETRIA]]").Offset(1, 0) = formulas(110, 2)
    range("tbl_espiro_info[[#Headers],[calculos_diagnostico]]").Offset(1, 0) = formulas(111, 2)
    range("tbl_espiro_info[[#Headers],[RESULTADO_ESPIROMETRIA_2 (resultado_espiro)]]").Offset(1, 0) = formulas(112, 2)
    range("tbl_espiro_info[[#Headers],[espirometriatipo_interpretacion]]").Offset(1, 0) = formulas(113, 2)
    range("tbl_espiro_info[[#Headers],[espirometriatipo_grado]]").Offset(1, 0) = formulas(114, 2)
    range("tbl_espiro_info[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Offset(1, 0) = formulas(115, 2)
    range("tbl_espiro_info[[#Headers],[SCRIPT espirometria]]").Offset(1, 0) = formulas(116, 2)
    range("tbl_espiro_info[[#Headers],[SCRIPT espiro_antecedentes_pivot]]").Offset(1, 0) = formulas(117, 2)
    range("tbl_espiro_info[[#Headers],[SCRIPT espiro_quimicos_pivot]]").Offset(1, 0) = formulas(118, 2)
    range("tbl_espiro_info[[#Headers],[SCRIPT espiro_riesgos_epp]]").Offset(1, 0) = formulas(119, 2)
    range("tbl_espiro_info[[#Headers],[SCRIPT espiro_recomendaciones_pivot]]").Offset(1, 0) = formulas(120, 2)
    range("tbl_espiro_info[[#Headers],[SCRIPT espiro_recomendaciones_lab_pivot]]").Offset(1, 0) = formulas(121, 2)
    range("tbl_espiro_info[[#Headers],[LLAVE]]").Offset(1, 0) = formulas(122, 2)
   Case "OSTEO"
    Worksheets(Sheet.Name).Select
    range("tbl_osteo[[#Headers],[IMC]]").Offset(1, 0) = formulas(123, 2)
    range("tbl_osteo[[#Headers],[CLASIFICACION IMC]]").Offset(1, 0) = formulas(124, 2)
    range("tbl_osteo[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Offset(1, 0) = formulas(125, 2)
    range("tbl_osteo[[#Headers],[SCRIPT osteomuscular]]").Offset(1, 0) = formulas(126, 2)
    range("tbl_osteo[[#Headers],[SCRIPT osteo_antecedentes_pivot]]").Offset(1, 0) = formulas(127, 2)
    range("tbl_osteo[[#Headers],[SCRIPT osteo_recomendaciones_pivot]]").Offset(1, 0) = formulas(128, 2)
   Case "COMPLEMENTARIOS"
    Worksheets(Sheet.Name).Select
    range("tbl_complementarios[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Offset(1, 0) = formulas(129, 2)
    range("tbl_complementarios[[#Headers],[SCRIPT complementarios]]").Offset(1, 0) = formulas(130, 2)
    range("tbl_complementarios[[#Headers],[SCRIPT complementarios_diagnos_observaciones_pivot]]").Offset(1, 0) = formulas(131, 2)
   Case "PSICOTECNICA"
    Worksheets(Sheet.Name).Select
    range("tbl_psicotecnica[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Offset(1, 0) = formulas(132, 2)
    range("tbl_psicotecnica[[#Headers],[SCRIPT psicotecnica]]").Offset(1, 0) = formulas(133, 2)
   Case "PSICOSENSOMETRICA"
    Worksheets(Sheet.Name).Select
    range("tbl_psicosensometrica[[#Headers],[emo_id(orden_lista_trabajadoresid)]]").Offset(1, 0) = formulas(134, 2)
    range("tbl_psicosensometrica[[#Headers],[SCRIPT psicosensometrica]]").Offset(1, 0) = formulas(135, 2)
    range("tbl_psicosensometrica[[#Headers],[SCRIPT psicosenso_diagnos_observaciones_pivot]]").Offset(1, 0) = formulas(136, 2)
    range("tbl_psicosensometrica[[#Headers],[SCRIPT psicosensometricas_recomendaciones_pivot]]").Offset(1, 0) = formulas(137, 2)
   Case "DIAGNOSTICOS"
    Worksheets(Sheet.Name).Select
    range("tbl_diagnosticos[[#Headers],[id emo]]").Offset(1, 0) = formulas(138, 2)
    range("tbl_diagnosticos[[#Headers],[TODO]]").Offset(1, 0) = formulas(139, 2)
   Case "ENFASIS"
    Worksheets(Sheet.Name).Select
    range("tbl_enfasis[[#Headers],[id_emo]]").Offset(1, 0) = formulas(140, 2)
    range("tbl_enfasis[[#Headers],[SQL ENFASIS_1]]").Offset(1, 0) = formulas(141, 2)
    For counter = 2 To 18 Step 1
      range("tbl_enfasis[[#Headers],[SQL ENFASIS_" & counter & "]]").Offset(1, 0) = VBA.Replace(formulas(142, 2), "_W", "_" & counter)
    Next counter
  End Select

End Sub
