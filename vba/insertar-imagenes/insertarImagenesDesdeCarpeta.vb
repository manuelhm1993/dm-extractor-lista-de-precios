Sub insertarImagenesDesdeCarpeta()
    Dim ws As Worksheet
    Dim fila As Long
    Dim codigo As String
    Dim rutaBase As String
    Dim rutaImagen As String
    Dim img As Picture
    Dim insertadas As Long
    Dim noEncontradas As Long
    Dim imagenFaltante As String

    ' VALIDAR hoja
    If Not WorksheetExists("IMAGEN") Then
        MsgBox "La hoja 'IMAGEN' no existe. Inserción cancelada.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets("IMAGEN")
    rutaBase = "C:\EXPORTA\" ' AJUSTA esta ruta

    ' LIMPIAR imágenes anteriores opcional (descomenta si quieres esto)
    ' Call eliminarTodasLasImagenes(ws)

    fila = 2
    insertadas = 0
    noEncontradas = 0

    Do While ws.Cells(fila, 1).Value <> ""
        codigo = Trim(ws.Cells(fila, 1).Value)
        rutaImagen = rutaBase & codigo & ".bmp"

        If Dir(rutaImagen) <> "" Then
            Set img = ws.Pictures.Insert(rutaImagen)
            With img
                .Top = ws.Cells(fila, 2).Top
                .Left = ws.Cells(fila, 2).Left
                .Width = ws.Cells(fila, 2).Width
                .Height = ws.Cells(fila, 2).Height
                .Placement = xlMoveAndSize
                .Name = "img_" & codigo
            End With
            insertadas = insertadas + 1
        Else
            noEncontradas = noEncontradas + 1
            imagenFaltante = imagenFaltante & vbNewLine & "- " & codigo & ".bmp"
        End If
        fila = fila + 1
    Loop

    ' Resultado final
    MsgBox "Insertadas: " & insertadas & vbCrLf & "No encontradas: " & noEncontradas, vbInformation, "Resultado"

    If noEncontradas > 0 Then
        MsgBox "No se encontraron las siguientes imágenes:" & vbNewLine & imagenFaltante, vbExclamation, "Imágenes Faltantes"
    End If
End Sub

' Función para validar existencia de hoja
Function WorksheetExists(wsName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not ThisWorkbook.Sheets(wsName) Is Nothing
    On Error GoTo 0
End Function

' Limpieza total de imágenes (opcional)
Sub eliminarTodasLasImagenes(ws As Worksheet)
    Dim sh As Shape
    For Each sh In ws.Shapes
        If Left(sh.Name, 4) = "img_" Then sh.Delete
    Next sh
End Sub
