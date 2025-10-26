Attribute VB_Name = "SSIM"

Sub DefineSSIM()
    Dim a As Variant
    Dim b As Integer
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = Worksheets("Structuring")
    Set tbl = ws.ListObjects("SSIM")
    
    b = tbl.DataBodyRange.Rows.Count
    
    a = InputBox("How many uncertainties would you like to report?", "Quantity of uncertainties", 1)
    
    If a = "" Then
        MsgBox "No uncertainties received"
        Exit Sub
        
    End If
    
    
    If a < 2 Then
        MsgBox ("The minimum number of uncertainties is 2")
        Exit Sub
    ElseIf a > 15 Then
        MsgBox ("You have exceeded the limit of uncertainties (#15)")
        Exit Sub
    End If
    
    Do While b < a
        b = b + 1
        Call AddLastRow
        Call AddLastCol(b)
        Call Dropdown
    Loop
    
    

End Sub

Sub ResetSSIM()

Dim ws As Worksheet
Set ws = Worksheets("Structuring")
Dim tbl As ListObject
Set tbl = ws.ListObjects("SSIM")

If ws.ChartObjects.Count <> 0 Then 'delete any chart
    ws.ChartObjects.Delete 'delete chart
End If

tbl.DataBodyRange.ClearContents 'clear table

Do While tbl.ListRows.Count > 2 'return format to normal
    ws.Rows(18 + tbl.ListRows.Count).RowHeight = 16
    tbl.ListRows(tbl.ListRows.Count).Delete
    tbl.ListColumns(tbl.ListColumns.Count).Delete
Loop



End Sub

Sub AddLastCol(col As Integer)
Dim ws As Worksheet
Set ws = Worksheets("Structuring")
Dim tbl As ListObject
Set tbl = ws.ListObjects("SSIM")


'add a new column at the end of the table
tbl.ListColumns.Add.Name = col
tbl.ListColumns(tbl.ListColumns.Count).DataBodyRange.Cells.HorizontalAlignment = xlHAlignCenter

'clear fill colour from last cell with dropdown in the new collumn
    With tbl.DataBodyRange.Cells(tbl.ListRows.Count - 1, tbl.ListColumns.Count).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
End Sub

Sub AddLastRow()
Dim ws As Worksheet
Set ws = Worksheets("Structuring")
Dim tbl As ListObject
Set tbl = ws.ListObjects("SSIM")
Dim c As Integer

'add a row at the end of the table
tbl.ListRows.Add
c = tbl.ListRows.Count
ws.Rows(18 + c).RowHeight = 25



End Sub

Sub Dropdown()
Dim ws As Worksheet
Set ws = Worksheets("Structuring")
Dim tbl As ListObject
Set tbl = ws.ListObjects("SSIM")

    tbl.ListColumns(tbl.ListColumns.Count).DataBodyRange.Select

    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="= SSIM_Values"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    f = tbl.ListRows(tbl.ListRows.Count).Range.Select
    Selection.Validation.Delete
    
 
End Sub





