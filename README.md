Private Sub cmdDBRefresh_Click()
    cmdDBRefresh.Caption = "Updating..."
    Sheet2.Cells(1, 13).Value = "Last Update:"
    Sheet2.Cells(1, 14).Value = Time
    Sheet2.Cells(1, 14).Font.Color = RGB(255, 0, 0): Sheet2.Cells(1, 14).Font.Bold = True: Sheet2.Cells(1, 14).Font.Size = 7
    ThisWorkbook.RefreshAll
    cmdDBRefresh.Caption = "DB Refresh"
End Sub


Private Sub cmdTrackingList_Click()
    ActiveWorkbook.FollowHyperlink Address:="http://sharepoint2.bankofamerica.com/sites/WebservicesTeam/Lists/OOS%20Tracking%20%20Assignments/Excel%20View.aspx"
End Sub


Private Sub cbxChart_Click()
Dim i As Integer: i = 1
With Sheet1
    .Cells(14, 4).Value = .cbxChart.Value
End With

With Sheet3
    Do While .Cells(i, 1).Value <> ""
        If .Cells(i, 1).Value = Sheet1.cbxChart.Value Then
            COUNTERS .Cells(i, 10)
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
End With
PrintChart
End Sub

