Sub transferirdatosotrahoja()
    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim filaOrigen As Long, filaDestino As Long
    Dim ultimaFila As Long
    Dim i As Long

    ' Establecer hojas
    Set wsOrigen = ThisWorkbook.Sheets("LISTA")
    Set wsDestino = ThisWorkbook.Sheets("PEDIDO")

    ' 1. Borrar datos existentes desde A3:F última fila en PEDIDO
    With wsDestino
        If Application.WorksheetFunction.CountA(.Range("A3:F" & .Rows.Count)) > 0 Then
            .Range("A3:F" & .Cells(.Rows.Count, "A").End(xlUp).Row).ClearContents
        End If
    End With

    ' 2. Ajustar fila de inicio
    filaOrigen = 3 ' Antes era 7
    filaDestino = 3

    ' 3. Calcular última fila con datos en LISTA
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row

    ' 4. Transferir datos
    For i = filaOrigen To ultimaFila
        ' Si hay algo escrito en la columna "G" (pedido), se transfiere
        If wsOrigen.Cells(i, "G").Value <> "" Then
            wsDestino.Cells(filaDestino, 1).Value = wsOrigen.Cells(i, 1).Value ' CODIGO
            wsDestino.Cells(filaDestino, 2).Value = wsOrigen.Cells(i, 2).Value ' DESCRIPCION
            wsDestino.Cells(filaDestino, 3).Value = wsOrigen.Cells(i, 5).Value ' PRECIO
            wsDestino.Cells(filaDestino, 4).Value = wsOrigen.Cells(i, 6).Value ' STOCK
            wsDestino.Cells(filaDestino, 5).Value = wsOrigen.Cells(i, 7).Value ' PEDIDO
            ' Importante: agrega aquí más columnas si fuese necesario
            filaDestino = filaDestino + 1
        End If
    Next i

    MsgBox "Datos transferidos correctamente.", vbInformation
End Sub
