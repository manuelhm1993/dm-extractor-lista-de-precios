Private Sub Workbook_Open()
    ' Llamar al procedimiento genérico para las columnas E y F
    convertirTextoANumero ThisWorkbook.Sheets("LISTA").Range("E3")
    convertirTextoANumero ThisWorkbook.Sheets("LISTA").Range("F3")
End Sub

Sub convertirTextoANumero(celdaInicio As Range)
    Dim ws As Worksheet
    Dim fila As Long
    Dim ultimaFila As Long
    Dim celda As Range
    Dim valor As String
    Dim numero As Double

    ' Obtener la hoja desde la celda inicial
    Set ws = celdaInicio.Worksheet

    ' Determinar la última fila de la columna especificada
    ultimaFila = ws.Cells(ws.Rows.Count, celdaInicio.Column).End(xlUp).Row

    ' Recorrer el rango desde la celda inicial hasta la última fila detectada
    For fila = celdaInicio.Row To ultimaFila
        Set celda = ws.Cells(fila, celdaInicio.Column)
        valor = Trim(celda.Value)

        ' Verificar si el valor contiene números
        If Len(valor) > 0 And IsNumeric(Replace(valor, ",", ".")) Then
            ' Reemplazar coma decimal por punto para conversión
            valor = Replace(valor, ",", ".")
            
            ' Intentar convertir a número de forma segura
            On Error Resume Next
            numero = CDbl(valor)
            On Error GoTo 0
            
            ' Verificación final de número correcto
            If numero <> 0 Or valor = "0" Then
                ' Asignar el valor convertido y aplicar formato numérico
                celda.Value = numero
                celda.NumberFormat = "0.00"
            End If
        End If
    Next fila

    MsgBox "Conversión completada para la columna " & celdaInicio.Address(0, 0), vbInformation
End Sub
