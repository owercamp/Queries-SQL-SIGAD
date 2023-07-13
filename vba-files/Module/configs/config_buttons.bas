Attribute VB_Name = "config_buttons"
'namespace=vba-files\Module\configs
Option Explicit

Public Sub btnCreate()
    '
    ' Buttons: a" & ChrW(243) & "ade todos los botones requeridas
    '
    '
    Dim names As Variant, positions As Variant, item As Variant, colorText As Variant, x As Integer
    Dim action As Variant

    names = Array("Traer Informaci" & ChrW(243) & "n", "Archivar Contenido", "Configuraci" & ChrW(243) & "n", "Modificaci" & ChrW(243) & "n", "Generar SQL")
    positions = Array(120, 238, 356, 474, 592)
    colorText = Array(RGB(183, 149, 11), RGB(131, 97, 141), RGB(123, 36, 28), RGB(133, 70, 61), RGB(135, 54, 0))
    action = Array("info", "clearContents", "config", "Modification", "ExportSQL")
    x = 0

    For Each item In names
      Dim btn As Variant
      btn = ActiveSheet.Buttons.Add(positions(x), 10, 113, 24).Select
      On Error Resume Next
      selection.OnAction = action(x)
      On Error GoTo 0
      With selection
        .Name = item
        .Caption = item
        .Font.Name = "Bahnschrift"
        .Font.FontStyle = "Negrita"
        .Font.Size = 11
        .Font.Strikethrough = False
        .Font.Superscript = False
        .Font.Subscript = False
        .Font.OutlineFont = False
        .Font.Shadow = False
        .Font.Underline = xlUnderlineStyleNone
        .Font.Color = colorText(x)
      End With
'        esta es una rutina para objectos de ActiveX
'        ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, Left:=positions(x), Top:=10, Width:=113, Height:=24).Select
'        Set btn = selection.Object
'        With btn
'            .Caption = Item ' Cambia el texto del bot" & ChrW(243) & "n
'            .ForeColor = colorText(x) ' Cambia el color del texto del bot" & ChrW(243) & "n
'            .FontBold = True
'            .Font.Name = "Bahnschrift" ' Cambia la tipografia del texto del bot" & ChrW(243) & "n
'        End With
'        selection.Name = Item
        x = x + 1
    Next item
    
    ActiveSheet.Shapes.range(Array("Generar SQL", "Traer Informaci" & ChrW(243) & "n", "Archivar Contenido", "Configuraci" & ChrW(243) & "n", "Modificaci" & ChrW(243) & "n")).Select
    selection.ShapeRange.Align msoAlignTops, msoFalse
    selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
    selection.ShapeRange.Group.Select
    
    range("tbl_trabajadores[[#Headers],[estado]]").Select
    
End Sub

