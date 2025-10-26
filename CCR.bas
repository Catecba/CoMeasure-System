Attribute VB_Name = "CCR"
Sub CCRNewRow()

Dim act As Worksheet, ws As Worksheet
Set ws = Worksheets("Input")
Set act = Worksheets("Activity list")
Dim tbl As ListObject, tblact As ListObject
Set tbl = ws.ListObjects("CCRs")
Set tblact = act.ListObjects("Activities")


'add a row at the end of the table
Application.EnableEvents = False
tbl.ListRows.Add
tbl.DataBodyRange.Cells(tbl.ListRows.Count, 1).Value = "CCR" & tbl.ListRows.Count
Application.EnableEvents = True

'add collumn in activity table
Application.EnableEvents = False
tblact.ListColumns.Add.Name = "#" & tbl.ListRows.Count
With tblact.ListColumns("#" & tbl.ListRows.Count)
    .DataBodyRange.Validation.Delete  'remove the dropdown listof the prob distr
    .Range.EntireColumn.Columns.AutoFit 'autofit the col
End With
tblact.ListColumns.Add.Name = "CCR" & tbl.ListRows.Count
tblact.ListColumns("CCR" & tbl.ListRows.Count).Range.ColumnWidth = 10
With tblact.ListColumns("CCR" & tbl.ListRows.Count).DataBodyRange.Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:="=ProbDist"
    .IgnoreBlank = True
    .InCellDropdown = True
    .ShowInput = True
    .ShowError = False
End With
Application.EnableEvents = True
End Sub
Sub CCRDeleteRow()

Dim act As Worksheet, ws As Worksheet
Set ws = Worksheets("Input")
Set act = Worksheets("Activity list")
Dim tbl As ListObject, tblact As ListObject
Set tbl = ws.ListObjects("CCRs")
Set tblact = act.ListObjects("Activities")


'delete row
If tbl.ListRows.Count > 1 Then
    Application.EnableEvents = False
    tbl.ListRows(tbl.ListRows.Count).Delete
    'Delete the name of the CCR above the activity table column
    act.Cells(tblact.HeaderRowRange.row - 1, tblact.ListColumns(tblact.ListColumns.Count).Range.Column).ClearContents
    'delete respective collumn in activity table + the # column
    For i = 0 To 1
    tblact.ListColumns(tblact.ListColumns.Count).Delete
    Next i
    Application.EnableEvents = True
End If

End Sub




