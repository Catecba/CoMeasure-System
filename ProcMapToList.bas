Attribute VB_Name = "ProcMapToList"
Sub SwimlaneDone()

    
    Dim wsDiagram As Worksheet
    Dim wsList As Worksheet
    Dim swim As ListObject
    Dim tbl As ListObject
    Dim node As Variant
    
    Set wsDiagram = Worksheets("Structuring")
    Set swim = wsDiagram.ListObjects("Swimlane")
    Set wsList = Worksheets("Activity list")
    Set tbl = wsList.ListObjects("Activities")
    
    Dim shp As Shape
    Dim i As Integer



    ' Clear previous diagram output
    
    Application.EnableEvents = False
    tbl.DataBodyRange.ClearContents
    tbl.DataBodyRange.Cells(1, 1).Value = "1."
    Do While tbl.ListRows.Count > 1
        tbl.ListRows(tbl.ListRows.Count).Delete
    Loop
    Application.EnableEvents = True
    
    wsList.Range("SimResCol").ClearContents

For Each shp In wsDiagram.Shapes 'loop all shapes
    If AuthorizedShape(shp) And (shp.AutoShapeType = msoShapeFlowchartProcess Or shp.AutoShapeType = msoShapeFlowchartAlternateProcess) Then 'add to the table of activities if the shape is of an autorized activity
        If IsEmpty(tbl.DataBodyRange.Cells(1, 2)) Then
                tbl.DataBodyRange.Cells(tbl.ListRows.Count, 2).Value = shp.TextFrame.Characters.text

        Else
                tbl.ListRows.Add
                tbl.DataBodyRange.Cells(tbl.ListRows.Count, 2).Value = shp.TextFrame.Characters.text
                tbl.DataBodyRange.Cells(tbl.ListRows.Count, 1).Value = tbl.ListRows.Count
        End If
    End If
Next shp


End Sub


Function AuthorizedShape(shp As Shape) As Boolean

Dim ws As Worksheet
Dim swim As ListObject

Set ws = Worksheets("Structuring")
Set swim = ws.ListObjects("Swimlane")

'get the boundaries of the swimlane
leftBoundary = swim.ListColumns(2).Range.Cells(1, 1).Left
    With swim.ListColumns(swim.ListColumns.Count).Range.Cells(1, 1)
        rightBoundary = .Left + .Width
    End With
topBoundary = swim.Range.Rows(1).Top

bottomBoundary = swim.Range.Rows(swim.ListRows.Count).Top + swim.Range.Rows(swim.ListRows.Count).Height

'check if the shp is within the swimlane
If shp.Top >= topBoundary And (shp.Top + shp.Height) <= bottomBoundary And shp.Left >= leftBoundary Then
    AuthorizedShape = True
Else
    AuthorizedShape = False
End If

End Function

Sub branch(BegNod As Variant)

    Dim branch As New Collection
    Dim ws As Worksheet
    Dim node As String
    Set ws = Worksheets("Structuring")
    
    If BegNod = "1" Then
         branch.Add ws.Shapes("1").TextFrame.Characters.text
    Else
        branch.Add BegNod
    End If
    
    node = BegNod
    'while there is more to go
    Do While Fathers.Exists(node)
        If ws.Shapes(Fathers(node)).AutoShapeType = msoShapeFlowchartProcess Then 'if the next shape is an activity
            branch.Add Fathers(node) 'add the new connection
            node = Fathers.Item(node) 'change the son to the father node to look for
        Else
            branch.Add Fathers(node) 'add the decision point and exit loop
            Exit Do
        End If
    Loop
    Set Branches(BegNod) = branch 'save the branch
  
End Sub


Sub BranchProb(shp As Shape, prob As String)


If shp.AutoShapeType = msoShapeFlowchartDecision Then
        If Probabilities.Exists(shp.Name) Then 'If a branch has been added for this decision
           Probabilities(shp.Name).Add CDbl(prob)  'get the collection of probs
        Else 'if it is the first branch
                Dim probs As New Collection
                probs.Add CDbl(prob)
                Probabilities.Add shp.Name, probs
         End If
Else
        Exit Sub
End If

End Sub


Function Isconnected(shp As Shape) As Boolean

    Dim shp1 As Shape, shp2 As Shape
    
    Isconnected = True
    
    On Error GoTo NoConnection
    Set shp1 = shp.ConnectorFormat.BeginConnectedShape
    Set shp2 = shp.ConnectorFormat.EndConnectedShape
    
    Exit Function
    'if there is an error in setting any of the shape ->The connector is not connected at that point
NoConnection:
        Isconnected = False

End Function


