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

    Dim bannerShape As Shape
    
    On Error Resume Next
    Set bannerShape = ws.Shapes("Banner") ' Asegurar que el banner se llame exactamente "Banner"
    On Error GoTo 0

    If Not bannerShape Is Nothing Then
        With bannerShape
            ' Anclar en esquina superior izquierda de A1
            .Top = ws.Range("A1").Top
            .Left = ws.Range("A1").Left
            ' Ajustar ancho exacto hasta el borde derecho de H1
            .Width = ws.Range("H1").Left + ws.Range("H1").Width - ws.Range("A1").Left
            ' Ajustar alto exacto de la fila 1
            .Height = ws.Rows(1).Height
            .Placement = xlMoveAndSize
        End With
    End If
End Sub

