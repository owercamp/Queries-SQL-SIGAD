VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formControl 
   Caption         =   "Config"
   ClientHeight    =   3924
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   8268.001
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
  If (Trim(Me.txt_informes.value) <> Trim(ThisWorkbook.Worksheets("RUTAS").range("$C$4").value)) Then: ThisWorkbook.Worksheets("RUTAS").range("$C$4") = Trim(Me.txt_informes.value)
End Sub

Private Sub btn_consolidado_Click()
  If (Trim(Me.txt_consolidado.value) <> Trim(ThisWorkbook.Worksheets("RUTAS").range("$C$5").value)) Then: ThisWorkbook.Worksheets("RUTAS").range("$C$5") = Trim(Me.txt_consolidado.value)
End Sub

Private Sub btn_script_Click()
  If (Trim(Me.txt_script.value) <> Trim(ThisWorkbook.Worksheets("RUTAS").range("$C$6").value)) Then: ThisWorkbook.Worksheets("RUTAS").range("$C$6") = Trim(Me.txt_script.value)
End Sub

Private Sub btn_cargos_Click()
  If (Trim(Me.txt_cargo.value) <> Trim(ThisWorkbook.Worksheets("RUTAS").range("$C$7").value)) Then: ThisWorkbook.Worksheets("RUTAS").range("$C$7") = Trim(Me.txt_cargo.value)
End Sub

Private Sub btn_backup_Click()
  If (Trim(Me.txt_backup.value) <> Trim(ThisWorkbook.Worksheets("RUTAS").range("$C$8").value)) Then: ThisWorkbook.Worksheets("RUTAS").range("$C$8") = Trim(Me.txt_backup.value)
End Sub

Private Sub txt_trabajadores_Change()
  If CLngPtr(Trim(txt_trabajadores.value)) <> CLngPtr(Trim(ThisWorkbook.Worksheets("RUTAS").range("$F$4").value)) Then
    ThisWorkbook.Worksheets("RUTAS").range("$F$4") = Trim(txt_trabajadores.value)
    ThisWorkbook.Worksheets("TRABAJADORES").range("$AW$5") = Trim(txt_trabajadores.value)
  End If
End Sub

Private Sub txt_emo_Change()
  If CLngPtr(Trim(txt_emo.value)) <> CLngPtr(Trim(ThisWorkbook.Worksheets("RUTAS").range("$F$5").value)) Then
    ThisWorkbook.Worksheets("RUTAS").range("$F$5") = Trim(txt_emo.value)
    ThisWorkbook.Worksheets("EMO").range("$EL$5") = Trim(txt_emo.value)
  End If
End Sub

Private Sub txt_audio_Change()
  If CLngPtr(Trim(txt_audio.value)) <> CLngPtr(Trim(ThisWorkbook.Worksheets("RUTAS").range("$F$6").value)) Then
    ThisWorkbook.Worksheets("RUTAS").range("$F$6") = Trim(txt_audio.value)
    ThisWorkbook.Worksheets("AUDIO").range("$BG$4") = Trim(txt_audio.value)
  End If
End Sub

Private Sub txt_opto_Change()
  If CLngPtr(Trim(txt_opto.value)) <> CLngPtr(Trim(ThisWorkbook.Worksheets("RUTAS").range("$F$7").value)) Then
    ThisWorkbook.Worksheets("RUTAS").range("$F$7") = Trim(txt_opto.value)
    ThisWorkbook.Worksheets("OPTO").range("$BL$4") = Trim(txt_opto.value)
  End If
End Sub

Private Sub txt_diag_Change()
  If CLngPtr(Trim(txt_diag.value)) <> CLngPtr(Trim(ThisWorkbook.Worksheets("RUTAS").range("$F$8").value)) Then
    ThisWorkbook.Worksheets("RUTAS").range("$F$8") = Trim(txt_diag.value)
    ThisWorkbook.Worksheets("OPTO").range("$BM$4") = Trim(txt_diag.value)
  End If
End Sub

Private Sub txt_visio_Change()
  If CLngPtr(Trim(txt_visio.value)) <> CLngPtr(Trim(ThisWorkbook.Worksheets("RUTAS").range("$F$9").value)) Then
    ThisWorkbook.Worksheets("RUTAS").range("$F$9") = Trim(txt_visio.value)
    ThisWorkbook.Worksheets("VISIO").range("$BS$4") = Trim(txt_visio.value)
  End If
End Sub

Private Sub txt_espiro_Change()
  If CLngPtr(Trim(txt_espiro.value)) <> CLngPtr(Trim(ThisWorkbook.Worksheets("RUTAS").range("$F$10").value)) Then
    ThisWorkbook.Worksheets("RUTAS").range("$F$10") = Trim(txt_espiro.value)
    ThisWorkbook.Worksheets("ESPIRO").range("$BZ$4") = Trim(txt_espiro.value)
  End If
End Sub

Private Sub txt_osteo_Change()
  If CLngPtr(Trim(txt_osteo.value)) <> CLngPtr(Trim(ThisWorkbook.Worksheets("RUTAS").range("$F$11").value)) Then
    ThisWorkbook.Worksheets("RUTAS").range("$F$11") = Trim(txt_osteo.value)
    ThisWorkbook.Worksheets("OSTEO").range("$BG$4") = Trim(txt_osteo.value)
  End If
End Sub

Private Sub txt_comple_Change()
  If CLngPtr(Trim(txt_comple.value)) <> CLngPtr(Trim(ThisWorkbook.Worksheets("RUTAS").range("$F$12").value)) Then
    ThisWorkbook.Worksheets("RUTAS").range("$F$12") = Trim(txt_comple.value)
    ThisWorkbook.Worksheets("COMPLEMENTARIOS").range("$J$4") = Trim(txt_comple.value)
  End If
End Sub

Private Sub txt_psico_Change()
  If CLngPtr(Trim(txt_psico.value)) <> CLngPtr(Trim(ThisWorkbook.Worksheets("RUTAS").range("$F$13").value)) Then
    ThisWorkbook.Worksheets("RUTAS").range("$F$13") = Trim(txt_psico.value)
    ThisWorkbook.Worksheets("PSICOTECNICA").range("$G$2") = Trim(txt_psico.value)
  End If
End Sub

Private Sub txt_senso_Change()
  If CLngPtr(Trim(txt_senso.value)) <> CLngPtr(Trim(ThisWorkbook.Worksheets("RUTAS").range("$F$14").value)) Then
    ThisWorkbook.Worksheets("RUTAS").range("$F$14") = Trim(txt_senso.value)
    ThisWorkbook.Worksheets("PSICOSENSOMETRICA").range("$Q$3") = Trim(txt_senso.value)
  End If
End Sub

Private Sub UserForm_Initialize()

  Me.txt_informes.value = ThisWorkbook.Worksheets("RUTAS").range("$C$4").value
  Me.txt_consolidado.value = ThisWorkbook.Worksheets("RUTAS").range("$C$5").value
  Me.txt_script.value = ThisWorkbook.Worksheets("RUTAS").range("$C$6").value
  Me.txt_cargo.value = ThisWorkbook.Worksheets("RUTAS").range("$C$7").value
  Me.txt_backup.value = ThisWorkbook.Worksheets("RUTAS").range("$C$8").value
  Me.txt_trabajadores.value = ThisWorkbook.Worksheets("RUTAS").range("$F$4").value
  Me.txt_emo.value = ThisWorkbook.Worksheets("RUTAS").range("$F$5").value
  Me.txt_audio.value = ThisWorkbook.Worksheets("RUTAS").range("$F$6").value
  Me.txt_opto.value = ThisWorkbook.Worksheets("RUTAS").range("$F$7").value
  Me.txt_diag.value = ThisWorkbook.Worksheets("RUTAS").range("$F$8").value
  Me.txt_visio.value = ThisWorkbook.Worksheets("RUTAS").range("$F$9").value
  Me.txt_espiro.value = ThisWorkbook.Worksheets("RUTAS").range("$F$10").value
  Me.txt_osteo.value = ThisWorkbook.Worksheets("RUTAS").range("$F$11").value
  Me.txt_comple.value = ThisWorkbook.Worksheets("RUTAS").range("$F$12").value
  Me.txt_psico.value = ThisWorkbook.Worksheets("RUTAS").range("$F$13").value
  Me.txt_senso.value = ThisWorkbook.Worksheets("RUTAS").range("$F$14").value
  Me.MultiPage1.Pages.Item (1)

End Sub
