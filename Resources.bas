Attribute VB_Name = "Resources"

Sub ResourceNewRow()

Dim ws As Worksheet
Set ws = Worksheets("Input")
Dim tbl As ListObject
Set tbl = ws.ListObjects("Resource")
'Dim C As Integer

'add a row at the end of the table
tbl.ListRows.Add

End Sub
Sub ResourceDeleteRow()

Dim ws As Worksheet
Set ws = Worksheets("Input")
Dim tbl As ListObject
Set tbl = ws.ListObjects("Resource")
'Dim C As Integer
If tbl.ListRows.Count > 1 Then 'delete a row at the end of the table
    tbl.ListRows(tbl.ListRows.Count).Delete
End If

End Sub


Sub ResetInputTables()

Dim act As Worksheet, ws As Worksheet
Set ws = Worksheets("Input")
Set act = Worksheets("Activity list")
Dim tbl As ListObject, tbl2 As ListObject, tblact As ListObject
Set tbl = ws.ListObjects("Resource")
Set tblccr = ws.ListObjects("CCRs")
Set tblact = act.ListObjects("Activities")
Set tblunc = ws.ListObjects("UncTable")

tbl.DataBodyRange.ClearContents 'clear table
tblccr.DataBodyRange.Cells(1, 2).ClearContents 'to activate cleaning the label above the CCR1 column in the activities table
tblccr.DataBodyRange.ClearContents
tblccr.DataBodyRange.Cells(1, 1).Value = "CCR1"
tblunc.DataBodyRange.ClearContents

Do While tbl.ListRows.Count > 1 'delete extra rows
    Call ResourceDeleteRow
Loop

 'delete extra rows in CCR and the CCR columns in activity list
Do While tblccr.ListRows.Count > 1
    Call ccr.CCRDeleteRow
Loop



Do While tblunc.ListRows.Count > 1 'delete extra rows in UncTable
    Call Uncertainties.UncDeleteRow
Loop


End Sub

