Attribute VB_Name = "ProcMapAreas"
Sub NewSwimArea()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = Worksheets("Structuring")
    Set tbl = ws.ListObjects("Swimlane")

    ' Stop if nextRow exceeds last used row
    If tbl.ListRows.Count > 7 Then
        MsgBox "Limit of rows reached", vbInformation
        Exit Sub
    End If

    'Add a new area
    tbl.ListRows.Add
    tbl.DataBodyRange.Cells(tbl.ListRows.Count, 1).Value = "AREA " & tbl.ListRows.Count
    tbl.ListRows(tbl.ListRows.Count).Range.RowHeight = 178



End Sub
Sub DeleteSwimArea()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = Worksheets("Structuring")
    Set tbl = ws.ListObjects("Swimlane")

    ' Stop if lastrow is the 1
    If tbl.ListRows.Count = 1 Then
        MsgBox "This area can not be deleted", vbInformation
        Exit Sub
    End If

    'Delete the last area
    tbl.ListRows(tbl.ListRows.Count).Range.RowHeight = 16
    tbl.ListRows(tbl.ListRows.Count).Delete
    
End Sub






