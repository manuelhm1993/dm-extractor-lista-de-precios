Sub exportarHojaPedido()
    Dim wbNuevo As Workbook
    Dim rutaTemporal As String

    ' Usar el directorio actual del archivo que contiene esta macro
    rutaTemporal = ThisWorkbook.Path & "\pedido.xlsx"

    ' Copiar la hoja PEDIDO a un nuevo libro
    ThisWorkbook.Sheets("PEDIDO").Copy
    Set wbNuevo = ActiveWorkbook

    ' Limpiar la celda H1 antes de guardar
    Call limpiarCeldaH1(wbNuevo)

    ' Guardar como .xlsx
    Application.DisplayAlerts = False
    wbNuevo.SaveAs fileName:=rutaTemporal, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbNuevo.Close SaveChanges:=False

    MsgBox "Archivo guardado correctamente en:" & vbCrLf & rutaTemporal, vbInformation, "Exportación completa"
End Sub

Private Sub limpiarCeldaH1(ByVal wb As Workbook)
    On Error Resume Next
    wb.Sheets(1).Range("H1").ClearContents
    On Error GoTo 0
End Sub
