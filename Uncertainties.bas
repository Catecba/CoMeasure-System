Attribute VB_Name = "Uncertainties"
Sub UncNewRow()

Dim ws As Worksheet
Set ws = Worksheets("Input")
Dim tbl As ListObject
Set tbl = ws.ListObjects("UncTable")
'Dim C As Integer

'add a row at the end of the table
Application.EnableEvents = False
tbl.ListRows.Add
Application.EnableEvents = True

End Sub
Sub UncDeleteRow()

Dim ws As Worksheet
Set ws = Worksheets("Input")
Dim tbl As ListObject
Set tbl = ws.ListObjects("UncTable")

Application.EnableEvents = False
If tbl.ListRows.Count > 1 Then 'delete a row at the end of the table
    tbl.ListRows(tbl.ListRows.Count).Delete
End If
Application.EnableEvents = True

End Sub


