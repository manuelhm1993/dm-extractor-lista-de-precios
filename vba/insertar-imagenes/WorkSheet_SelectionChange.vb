Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim codigo As String
    Dim img As Shape

    If Target.Column = 1 And Target.Row > 1 Then ' Columna A
        codigo = Target.Value
        On Error Resume Next
        Set img = Me.Shapes("img_" & codigo)
        On Error GoTo 0
        
        If Not img Is Nothing Then
            img.Select
            MsgBox "Imagen seleccionada para " & codigo, vbInformation
            ' También podrías mostrarla en un UserForm con Image control
        End If
    End If
End Sub
