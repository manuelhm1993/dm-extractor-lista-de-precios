Sub transferirdatosotrahoja()
    Dim codigo As String
    Dim descrip As String
    Dim ref As String
    Dim marca As String
    Dim precio As Double
    Dim pedido As Double
    
    Dim ultimafila As Long
    Dim ultimafilaauxiliar As Long
    Dim cont As Long
    
    ultimafila = Sheets("LISTA").Range("A" & Rows.Count).End(xlUp).Row
    
    If ultimafila < 2 Then
        Exit Sub
    End If
    
    For cont = 3 To ultimafila
        If Sheets("LISTA").Cells(cont, 7) >= 1 Then

            codigo = Sheets("LISTA").Cells(cont, 1)
            descrip = Sheets("LISTA").Cells(cont, 2)
            ref = Sheets("LISTA").Cells(cont, 3)
            marca = Sheets("LISTA").Cells(cont, 4)
            precio = Sheets("LISTA").Cells(cont, 5)
            pedido = Sheets("LISTA").Cells(cont, 7)
            
            ultimafilaauxiliar = Sheets("PEDIDO").Range("A" & Rows.Count).End(xlUp).Row
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 1) = codigo
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 2) = descrip
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 3) = ref
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 4) = marca
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 5) = precio
            Sheets("PEDIDO").Cells(ultimafilaauxiliar + 1, 6) = pedido
        End If
    Next cont
    
    MsgBox "Proceso Culminado", vbInformation, "Resultado"
    
End Sub
