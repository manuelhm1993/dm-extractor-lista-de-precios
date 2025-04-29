Sub insertarImagenesDesdeCarpeta()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("IMAGEN")

    Dim fila As Long
    Dim codigo As String
    Dim rutaBase As String
    Dim rutaImagen As String
    Dim img As Picture

    rutaBase = "C:\EXPORTA\" ' Ruta local con las imágenes

    fila = 2 ' Asumimos encabezado en fila 1
    Do While ws.Cells(fila, 1).Value <> ""
        codigo = ws.Cells(fila, 1).Value
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
        End If
        fila = fila + 1
    Loop

    MsgBox "Imágenes insertadas exitosamente.", vbInformation
End Sub