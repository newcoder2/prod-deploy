# prod-deploy
master repository

Private Sub Workbook_Open()
On Error GoTo ErrHandler
Dim Flag As Boolean
    LoadCombo
    Flag = FindSheet("Template")
    If Not Flag Then
        MsgBox "Please confirm if 'Template' sheet exists or has been rename by mistake", vbCritical, "Forecast Tracker"
    Else
        ThisWorkbook.Worksheets("Template").Visible = False
    End If

ErrHandler:
If Err.Number Then
    MsgBox "Workbook_open process-> " & vbCrLf & Err.Description, vbCritical, "Forecast Tracker"
End If
End Sub


