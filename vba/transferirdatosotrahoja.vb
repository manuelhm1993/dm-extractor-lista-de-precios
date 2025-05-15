Sub transferirdatosotrahoja()
    Dim codigo As String
    Dim descrip As String
    Dim ref As String
    Dim marca As String
    Dim precio As Double
    Dim pedido As Double

    Dim ultimaFila As Long
    Dim ultimafilaauxiliar As Long
    Dim cont As Long

    ' 1. Limpiar el rango A3:G en hoja PEDIDO, incluyendo fila de total si existe
    With Sheets("PEDIDO")
        ultimaFila = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        ' Si hay registros desde la fila 3, limpiar el rango A3:G última fila
        If ultimaFila >= 3 Then
            .Range("A3:G" & ultimaFila).ClearContents
            
            ' Verificar si la siguiente fila contiene "TOTAL" en la columna F y eliminarla si es así
            If .Cells(ultimaFila + 1, 6).Value = "TOTAL" Then
                .Rows(ultimaFila + 1).Delete
            End If
        End If
    End With

    ' 2. Obtener la última fila de la hoja LISTA
    ultimaFila = Sheets("LISTA").Range("A" & Rows.Count).End(xlUp).Row
    If ultimaFila < 3 Then Exit Sub

    ' 3. Transferir datos de LISTA a PEDIDO
    For cont = 3 To ultimaFila
        If Sheets("LISTA").Cells(cont, 7) >= 1 Then

            codigo = Sheets("LISTA").Cells(cont, 1)
            descrip = Sheets("LISTA").Cells(cont, 2)
            ref = Sheets("LISTA").Cells(cont, 3)
            marca = Sheets("LISTA").Cells(cont, 4)
            precio = Sheets("LISTA").Cells(cont, 5)
            pedido = Sheets("LISTA").Cells(cont, 7)

            ' Calcular la última fila disponible en PEDIDO
            ultimafilaauxiliar = Sheets("PEDIDO").Range("A" & Rows.Count).End(xlUp).Row
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 1) = codigo
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 2) = descrip
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 3) = ref
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 4) = marca
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 5) = pedido
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 6) = precio
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 7) = pedido * precio
        End If
    Next cont

    ' 4. Recalcular la última fila después de transferir los datos
    ultimafilaauxiliar = Sheets("PEDIDO").Range("A" & Rows.Count).End(xlUp).Row

    ' 5. Agregar la fila de total al final
    With Sheets("PEDIDO")
        .Cells(ultimafilaauxiliar + 1, 6).Value = "TOTAL"
        .Cells(ultimafilaauxiliar + 1, 6).Font.Bold = True
        .Cells(ultimafilaauxiliar + 1, 7).Formula = "=SUM(G3:G" & ultimafilaauxiliar & ")"
        .Cells(ultimafilaauxiliar + 1, 7).Font.Bold = True
    End With

    MsgBox "Proceso Culminado", vbInformation, "Resultado"
End Sub
