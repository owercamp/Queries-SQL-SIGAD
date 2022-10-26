VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formControl
Caption         =   "Config"
ClientHeight    =   3930
ClientLeft      =   45
ClientTop       =   390
ClientWidth     =   8265.001
OleObjectBlob   =   "formControl.frx":0000
StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "formControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btn_informe_Click()
    If (Trim(Me.txt_informes.value) <> Trim(ThisWorkbook.Worksheets("RUTAS").Range("$C$4").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$C$4") = Trim(Me.txt_informes.value)
End Sub

Private Sub btn_consolidado_Click()
    If (Trim(Me.txt_consolidado.value) <> Trim(ThisWorkbook.Worksheets("RUTAS").Range("$C$5").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$C$5") = Trim(Me.txt_consolidado.value)
End Sub

Private Sub btn_script_Click()
    If (Trim(Me.txt_script.value) <> Trim(ThisWorkbook.Worksheets("RUTAS").Range("$C$6").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$C$6") = Trim(Me.txt_script.value)
End Sub

Private Sub btn_cargos_Click()
    If (Trim(Me.txt_cargo.value) <> Trim(ThisWorkbook.Worksheets("RUTAS").Range("$C$7").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$C$7") = Trim(Me.txt_cargo.value)
End Sub

Private Sub btn_backup_Click()
    If (Trim(Me.txt_backup.value) <> Trim(ThisWorkbook.Worksheets("RUTAS").Range("$C$8").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$C$8") = Trim(Me.txt_backup.value)
End Sub

Private Sub txt_trabajadores_Change()
    If CLngLng(Trim(txt_trabajadores.value)) <> CLngLng(Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$4").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$F$4") = Trim(txt_trabajadores.value)
End Sub

Private Sub txt_emo_Change()
    If CLngLng(Trim(txt_emo.value)) <> CLngLng(Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$5").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$F$5") = Trim(txt_emo.value)
End Sub

Private Sub txt_audio_Change()
    If CLngLng(Trim(txt_audio.value)) <> CLngLng(Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$6").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$F$6") = Trim(txt_audio.value)
End Sub

Private Sub txt_opto_Change()
    If CLngLng(Trim(txt_opto.value)) <> CLngLng(Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$7").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$F$7") = Trim(txt_opto.value)
End Sub

Private Sub txt_diag_Change()
    If CLngLng(Trim(txt_diag.value)) <> CLngLng(Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$8").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$F$8") = Trim(txt_diag.value)
End Sub

Private Sub txt_visio_Change()
    If CLngLng(Trim(txt_visio.value)) <> CLngLng(Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$9").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$F$9") = Trim(txt_visio.value)
End Sub

Private Sub txt_espiro_Change()
    If CLngLng(Trim(txt_espiro.value)) <> CLngLng(Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$10").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$F$10") = Trim(txt_espiro.value)
End Sub

Private Sub txt_osteo_Change()
    If CLngLng(Trim(txt_osteo.value)) <> CLngLng(Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$11").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$F$11") = Trim(txt_osteo.value)
End Sub

Private Sub txt_comple_Change()
    If CLngLng(Trim(txt_comple.value)) <> CLngLng(Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$12").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$F$12") = Trim(txt_comple.value)
End Sub

Private Sub txt_psico_Change()
    If CLngLng(Trim(txt_psico.value)) <> CLngLng(Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$13").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$F$13") = Trim(txt_psico.value)
End Sub

Private Sub txt_senso_Change()
    If CLngLng(Trim(txt_senso.value)) <> CLngLng(Trim(ThisWorkbook.Worksheets("RUTAS").Range("$F$14").value)) Then: ThisWorkbook.Worksheets("RUTAS").Range("$F$14") = Trim(txt_senso.value)
End Sub

Private Sub UserForm_Initialize()

    Me.txt_informes.value = ThisWorkbook.Worksheets("RUTAS").Range("$C$4").value
    Me.txt_consolidado.value = ThisWorkbook.Worksheets("RUTAS").Range("$C$5").value
    Me.txt_script.value = ThisWorkbook.Worksheets("RUTAS").Range("$C$6").value
    Me.txt_cargo.value = ThisWorkbook.Worksheets("RUTAS").Range("$C$7").value
    Me.txt_backup.value = ThisWorkbook.Worksheets("RUTAS").Range("$C$8").value
    Me.txt_trabajadores.value = ThisWorkbook.Worksheets("RUTAS").Range("$F$4").value
    Me.txt_emo.value = ThisWorkbook.Worksheets("RUTAS").Range("$F$5").value
    Me.txt_audio.value = ThisWorkbook.Worksheets("RUTAS").Range("$F$6").value
    Me.txt_opto.value = ThisWorkbook.Worksheets("RUTAS").Range("$F$7").value
    Me.txt_diag.value = ThisWorkbook.Worksheets("RUTAS").Range("$F$8").value
    Me.txt_visio.value = ThisWorkbook.Worksheets("RUTAS").Range("$F$9").value
    Me.txt_espiro.value = ThisWorkbook.Worksheets("RUTAS").Range("$F$10").value
    Me.txt_osteo.value = ThisWorkbook.Worksheets("RUTAS").Range("$F$11").value
    Me.txt_comple.value = ThisWorkbook.Worksheets("RUTAS").Range("$F$12").value
    Me.txt_psico.value = ThisWorkbook.Worksheets("RUTAS").Range("$F$13").value
    Me.txt_senso.value = ThisWorkbook.Worksheets("RUTAS").Range("$F$14").value
    Me.MultiPage1.Pages.Item(1)

End Sub
