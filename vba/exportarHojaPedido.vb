Sub exportarHojaPedido()
    Dim wbNuevo As Workbook
    Dim rutaTemporal As String

    ' Usar el directorio actual del archivo que contiene esta macro
    rutaTemporal = ThisWorkbook.Path & "\pedido.xlsx"

    ' Copiar la hoja PEDIDO a un nuevo libro
    ThisWorkbook.Sheets("PEDIDO").Copy
    Set wbNuevo = ActiveWorkbook

    ' Guardar como .xlsx
    Application.DisplayAlerts = False
    wbNuevo.SaveAs fileName:=rutaTemporal, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbNuevo.Close SaveChanges:=False

    MsgBox "Archivo guardado correctamente en:" & vbCrLf & rutaTemporal, vbInformation, "Exportación completa"
End Sub
