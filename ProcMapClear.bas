Attribute VB_Name = "ProcMapClear"
Sub ClearShapes()

Dim ws As Worksheet
Dim tbl As ListObject
Dim shp As Shape
Dim config As Worksheet


Set ws = Worksheets("Structuring")
Set tbl = ws.ListObjects("Swimlane")

'loops through all shapes
For Each shp In ws.Shapes
    If ProcMapToList.AuthorizedShape(shp) Then 'if the shape is within the process map area
        If shp.Name = "1" And shp.AutoShapeType = msoShapeFlowchartAlternateProcess Then
            shp.TextFrame.Characters.Delete 'clear of content the first node
        Else
            shp.Delete 'and remove any other
        End If
    End If
Next shp

End Sub

