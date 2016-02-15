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

Option Explicit
Dim sNewName As String
Dim WS, myWS As Object


Private Sub cmdExecute_Click()
On Error GoTo ErrHandler
Dim sName As String
Dim ifValue As Integer
Dim WB As Object

Set WB = ThisWorkbook.Worksheets

sName = Sheet1.cbxExSht.Text
ifValue = MsgBox("Check Total process for -" & sName & "- month will start, please confirm.", vbOKCancel, "Forecast Tracker")

If ifValue = 1 Then
    If sName <> "CPT List" And sName <> "Template" And sName <> "Reports" And sName <> "Settings" And sName <> "Select a Month" Then
        With WB(sName)
            CalProc sName
        End With
    Else
        MsgBox "Please select a correct sheet", vbExclamation, "Forecast Tracker"
    End If
Else
    Exit Sub
End If

Set WB = Nothing

ErrHandler:
If Err.Number Then
    MsgBox "cmdExucute process: " & vbCrLf & "Description: " & Err.Description, vbExclamation, "Forecast Tracker"
End If
End Sub

Private Sub cmdNewSheet_Click()
On Error GoTo ErrHandler
Dim nValue, nflag As Boolean
Dim ii As Integer
Dim oWS, nWS As Object

Set nWS = ThisWorkbook.Worksheets

nflag = CheckTemplate(nWS, "Template")

If Not nflag Then
    MsgBox "Please hide the template sheet or check its name", vbExclamation, "Forecast Counter"
    nflag = True
    Set oWS = Nothing
    Set nWS = Nothing
    Exit Sub
End If

'*** Will create a new worksheet with the month indicated by user
'*** The template will be hardcoded based on the list in CPT List sheet.
'*** the number of days will be hardcoded base on the month entered by user.

sNewName = ucase(InputBox("Enter the month you will create, i.e. January = JAN"))
If Len(sNewName) >= 3 Then
    sNewName = Left(sNewName, 3)
Else
    MsgBox "Add a complete Month name for sheet", vbInformation, "Forecast Tracker"
    Exit Sub
End If
'************ COPY TEMPLATE SHEET & VALIDATE SHEET NAME *****************
If sNewName <> "" Then
    nValue = IsNumeric(sNewName)
    If Not nValue Then
        For ii = 1 To nWS.Count
            If nWS(ii).Name = sNewName Then
                MsgBox WS(ii).Name & " sheet already exists, please enter a different one", vbExclamation, "Forecast Counter"
                Set nWS = Nothing
                Set oWS = Nothing
                Exit Sub
            Else
                nflag = False
            End If
        Next
        If nflag = False Then
            nWS("Template").Copy After:=nWS(nWS.Count) 'Issues with object
            nWS(nWS.Count).Name = sNewName
            nWS(sNewName).Visible = True
            nWS(sNewName).cmdSaveTemplate.Visible = False
            nWS(sNewName).Select
        End If
    Else
        MsgBox "Enter a valid month name", vbInformation, "Forecast Counter"
        Exit Sub
    End If
Else
    MsgBox "Please entere a Month in order to create a new template", vbInformation, "Forecast Counter"
    Exit Sub
End If

'***********************************
Template sNewName   'Calling template procedure
LoadCombo                 'Loading combobox in CPT List sheet
'***********************************
Set nWS = Nothing
Set oWS = Nothing

ErrHandler: 'HANDLING ERROR MESSAGE========================
If Err.Number Then
    MsgBox "cmdNewSheet Process." & vbCrLf & "Description: " & Err.Description, vbExclamation, "Forecast Counter"
    Set nWS = Nothing
    Set oWS = Nothing
    Err.Clear
End If

End Sub

Private Sub CleanWS()
Dim S, i As Object

Set S = ThisWorkbook.Worksheets

For Each i In ThisWorkbook.Worksheets
    MsgBox i.CodeName & " - " & i.Name
Next

End Sub

Private Sub cmdTemplate_Click()
On Error GoTo ErrHandler
Dim mWS As Object
Dim Flag As Boolean

Set mWS = ThisWorkbook.Worksheets

'*** Check if there is a template sheet
Flag = FindSheet("Template")
If Not Flag Then
    MsgBox "There is no Template sheet, please veirfy.", vbCritical, "Forecast Tracker"
Else
    With mWS("Template")
        .Visible = True
        .cmdSaveTemplate.Visible = True
        .Select
    End With
End If

Set mWS = Nothing

ErrHandler:
If Err.Number Then
    MsgBox "cmdTemplate - " & Err.Description, vbExclamation, "Forecast Tracker"
End If
End Sub

Private Sub Worksheet_Activate()
LoadCombo
End Sub



