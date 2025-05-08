Private Sub Workbook_Open()
    ' Llamar a la función al abrir el archivo
    Call convertirTextoANumero
End Sub

Sub convertirTextoANumero()
    Dim ws As Worksheet
    Dim ultimaFilaE As Long, ultimaFilaF As Long
    Dim rangoE As Range, rangoF As Range
    Dim celda As Range

    ' Definir la hoja de trabajo (ajusta el nombre según tu hoja)
    Set ws = ThisWorkbook.Sheets("LISTA")

    ' Encontrar la última fila con datos en columna E y F
    ultimaFilaE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    ultimaFilaF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    ' Verificar que haya datos a partir de E3 y F3
    If ultimaFilaE >= 3 Then
        Set rangoE = ws.Range("E3:E" & ultimaFilaE)
        For Each celda In rangoE
            If IsNumeric(Replace(celda.Value, ",", ".")) Then
                celda.Value = CDbl(Replace(celda.Value, ",", "."))
                celda.NumberFormat = "0.00" ' Formato numérico con dos decimales
            End If
        Next celda
    End If

    If ultimaFilaF >= 3 Then
        Set rangoF = ws.Range("F3:F" & ultimaFilaF)
        For Each celda In rangoF
            If IsNumeric(Replace(celda.Value, ",", ".")) Then
                celda.Value = CDbl(Replace(celda.Value, ",", "."))
                celda.NumberFormat = "0.00" ' Formato numérico con dos decimales
            End If
        Next celda
    End If

    MsgBox "Conversión de texto a número completada en columnas E y F.", vbInformation
End Sub
