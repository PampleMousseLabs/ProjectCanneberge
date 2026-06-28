Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("BD9")) Is Nothing Then
        Call BuildCompChart
    End If
End Sub
