Private Sub Workbook_Open()
    ' Llamar a la función al abrir el archivo
    Call convertirTextoANumero
End Sub

Sub convertirTextoANumero()
    Dim ws As Worksheet
    Dim ultimaFilaE As Long, ultimaFilaF As Long
    Dim rangoE As Range, rangoF As Range
    Dim celda As Range
    Dim valor As String
    Dim numero As Double

    ' Definir la hoja de trabajo (ajusta el nombre según tu hoja)
    Set ws = ThisWorkbook.Sheets("LISTA")

    ' Encontrar la última fila con datos en columna E y F
    ultimaFilaE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    ultimaFilaF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    ' Convertir texto a número en columna E
    If ultimaFilaE >= 3 Then
        Set rangoE = ws.Range("E3:E" & ultimaFilaE)
        For Each celda In rangoE
            valor = Trim(celda.Value)
            If InStr(valor, ",") > 0 Then
                ' Eliminar cualquier separador de miles y mantener el decimal
                If Len(Split(valor, ",")(1)) > 2 Then
                    valor = Replace(valor, ".", "") ' Eliminar punto si es de miles
                    valor = Replace(valor, ",", ".") ' Reemplazar coma decimal
                Else
                    valor = Replace(valor, ",", ".") ' Solo reemplazo de decimal
                End If
                ' Intentar la conversión a número
                If IsNumeric(valor) Then
                    numero = CDbl(valor)
                    celda.Value = numero
                    celda.NumberFormat = "0.00"
                End If
            End If
        Next celda
    End If

    ' Convertir texto a número en columna F
    If ultimaFilaF >= 3 Then
        Set rangoF = ws.Range("F3:F" & ultimaFilaF)
        For Each celda In rangoF
            valor = Trim(celda.Value)
            If InStr(valor, ",") > 0 Then
                ' Eliminar cualquier separador de miles y mantener el decimal
                If Len(Split(valor, ",")(1)) > 2 Then
                    valor = Replace(valor, ".", "") ' Eliminar punto si es de miles
                    valor = Replace(valor, ",", ".") ' Reemplazar coma decimal
                Else
                    valor = Replace(valor, ",", ".") ' Solo reemplazo de decimal
                End If
                ' Intentar la conversión a número
                If IsNumeric(valor) Then
                    numero = CDbl(valor)
                    celda.Value = numero
                    celda.NumberFormat = "0.00"
                End If
            End If
        Next celda
    End If

    MsgBox "Conversión de texto a número completada en columnas E y F.", vbInformation
End Sub
