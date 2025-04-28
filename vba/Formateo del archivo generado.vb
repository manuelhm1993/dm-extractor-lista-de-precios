Private Sub Workbook_Open()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LISTA")

    ' 1. Ajustar altura exacta del banner
    ws.Rows(1).RowHeight = 141

    ' 2. Forzar fusión A1:H1 para banner
    ws.Rows(1).UnMerge
    With ws.Range("A1:H1")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 3. Restaurar encabezados desde archivo original
    Dim encabezados As Variant
    encabezados = Array("CODIGO", "DESCRIPCION", "REFERENCIA", "MARCA", "PRECIO $", "STOCK", "PEDIDO", "IMAGEN")

    Dim i As Integer
    For i = LBound(encabezados) To UBound(encabezados)
        ws.Cells(2, i + 1).Value = encabezados(i)
        ws.Cells(2, i + 1).Font.Bold = True
        ' ws.Cells(2, i + 1).Interior.Color = RGB(255, 255, 0) ' <-- Color original amarillo de lista de precios
        ws.Cells(2, i + 1).Interior.Color = RGB(200, 200, 200) ' Color suave de fondo opcional
    Next i

    ' 4. Verificar y crear botón "Procesar Pedido" en I1 si no existe
    ' Dim botonExiste As Boolean: botonExiste = False
    ' Dim s As Shape

    ' For Each s In ws.Shapes
    '     If s.Name = "ProcesarPedido" Then
    '         botonExiste = True
    '         Exit For
    '     End If
    ' Next s

    ' If Not botonExiste Then
    '     Set s = ws.Shapes.AddFormControl(xlButtonControl, ws.Range("I1").Left, ws.Range("I1").Top, 100, 30)
    '     With s
    '         .Name = "ProcesarPedido"
    '         .TextFrame.Characters.Text = "Procesar Pedido"
    '         .OnAction = "ProcesarPedidoMacro" ' Cambia al nombre de la macro real
    '     End With
    ' End If
End Sub
