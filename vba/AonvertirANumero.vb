Private Sub Workbook_Open()
    ' Llamar a la función para columnas E y F
    Call procesarRangoTextoANumero(ThisWorkbook.Sheets("LISTA").Range("E3"))
    Call procesarRangoTextoANumero(ThisWorkbook.Sheets("LISTA").Range("F3"))
End Sub

Sub procesarRangoTextoANumero(celdaInicio As Range)
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim celda As Range
    Dim valor As String
    Dim numero As Double

    ' Definir la hoja desde la celda inicial
    Set ws = celdaInicio.Worksheet

    ' Determinar la última fila con datos en la columna
    ultimaFila = ws.Cells(ws.Rows.Count, celdaInicio.Column).End(xlUp).Row

    ' Recorrer el rango desde la celda inicial hasta la última fila detectada
    For Each celda In ws.Range(celdaInicio, ws.Cells(ultimaFila, celdaInicio.Column))
        ' Paso 1: Reemplazar comas por puntos
        valor = Trim(celda.Value)
        valor = reemplazarComasPorPuntos(valor)

        ' Paso 2: Intentar convertir a número si es válido
        If IsNumeric(valor) Then
            numero = CDbl(valor)
            celda.Value = numero
            celda.NumberFormat = "0.00"
        End If
    Next celda

    MsgBox "Procesamiento completado para la columna " & celdaInicio.Address(0, 0), vbInformation
End Sub

Function reemplazarComasPorPuntos(valor As String) As String
    ' Función que reemplaza comas con puntos en el valor dado
    If InStr(valor, ",") > 0 Then
        valor = Replace(valor, ",", ".")
    End If
    reemplazarComasPorPuntos = valor
End Function
