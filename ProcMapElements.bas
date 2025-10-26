Attribute VB_Name = "ProcMapElements"
Sub Addactivity()

Dim text As String


'user writes the name of the activity to be put in the shape
text = InputBox("Write the name of the activity to be added", "Activity name", "Activity Name")
If text = "" Then
    MsgBox "No input was submitted"
    Exit Sub
End If

'create the shape
    ActiveSheet.Shapes.AddShape(msoShapeFlowchartProcess, 326, 217, 146, 29).Select
    With Selection.ShapeRange
    .Line.Visible = msoTrue
    .Line.ForeColor.RGB = RGB(63, 71, 81)
    .Line.Weight = 1
    .Shadow.Visible = msoFalse
    .Fill.Visible = msoFalse
    .TextFrame2.TextRange.Font.Bold = msoFalse
    .TextFrame2.VerticalAnchor = msoAnchorMiddle
    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    .TextFrame2.TextRange.Characters.text = text
    .Name = text
    End With
End Sub
    

Sub Adddecision()

Dim text As String


'usewrites the name of the activity to be put in the shape
text = InputBox("Write the name of the decision to be added", "Decision name", 2)
If text = "" Then
    MsgBox "No input was submitted"
    Exit Sub
End If

'create a decision
    ActiveSheet.Shapes.AddShape(msoShapeFlowchartDecision, 340, 281, 108, 59).Select
    With Selection.ShapeRange
    .Line.Visible = msoTrue
    .Line.ForeColor.RGB = RGB(63, 71, 81)
    .Line.Weight = 1
    .Shadow.Visible = msoFalse
    .Fill.Visible = msoFalse
    .TextFrame2.TextRange.Font.Bold = msoFalse
    .TextFrame2.VerticalAnchor = msoAnchorMiddle
    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    .TextFrame2.TextRange.Characters.text = text
    .Name = text
    End With

End Sub

Sub Addlink()

    Dim ws As Worksheet
    Dim shp2con As ShapeRange
    Dim connector As Shape
    Dim config As Worksheet
    Dim ar() As String
    Dim Father As String
    Dim Son As String
    Dim prob As String
    
    On Error GoTo Errorhandler
    
    Set ws = Worksheets("Structuring")
    Set config = Worksheets("Configuration")
    
    '
    Set shp2con = Selection.ShapeRange
        ' Create connector
        Father = shp2con(1).Name
        Son = shp2con(2).Name
        ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 1, 1, 1, 1).Select '
        If shp2con(1).AutoShapeType = msoShapeFlowchartDecision Then 'if it is a divergency input the probability
            prob = InputBox("Write the probability of this branch in decimal form", "Branch Probability", 2)
            With Selection
                .Name = prob
            End With
        End If
        With Selection.ShapeRange
        .ConnectorFormat.BeginConnect ws.Shapes(Father), 1
        .ConnectorFormat.EndConnect ws.Shapes(Son), 1
        .RerouteConnections
        .Line.EndArrowheadStyle = msoArrowheadTriangle
        End With
        

    Exit Sub

  
Errorhandler:
    MsgBox "You didn't select authorized shape(s)"
    Exit Sub

End Sub

Sub Addnote()
Attribute Addnote.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveSheet.Shapes.AddShape(msoShapeRound2DiagRectangle, 367.4, 346.1, 69.7, 23.6).Select
    With Selection.ShapeRange
    .Line.Visible = msoTrue
    .Line.ForeColor.RGB = RGB(63, 71, 81)
    .Line.Weight = 1
    .Shadow.Visible = msoFalse
    .Fill.Visible = msoFalse
    .TextFrame2.TextRange.Font.Bold = msoFalse
    .TextFrame2.VerticalAnchor = msoAnchorMiddle
    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    .TextFrame2.TextRange.Font.Size = 8
    End With
    
    
End Sub
