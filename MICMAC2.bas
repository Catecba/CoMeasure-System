Attribute VB_Name = "MICMAC2"
Sub MICMAC(names() As String, Ind_power() As Double, Dep_power() As Double)

Dim ws As Worksheet
Dim config As Worksheet
Set config = Worksheets("Configuration")
Set ws = Worksheets("Structuring")

If ws.ChartObjects.Count <> 0 Then 'delete previous chart
    ws.ChartObjects.Delete
End If

Dim MICMAC As ChartObject, i As Long
Set MICMAC = Sheets("Structuring").ChartObjects.Add(1300, 140, 600, 500)

'max = UBound(Dep_power) + 1


With MICMAC.Chart
    .ChartType = xlXYScatter
    .SeriesCollection.NewSeries
    .Legend.Delete
    .HasTitle = True
    .ChartTitle.text = "MICMAC"
    .ChartTitle.Left = 527.845
    .ChartTitle.Top = 2
    .PlotArea.Height = .PlotArea.Height - 40
    'Defining Axis
    With .Axes(xlCategory)
        .MaximumScale = 10
        .MinimumScale = 0
        .MajorUnit = 1
        .Format.Line.Weight = 2
        .HasTitle = True
        .AxisTitle.Characters.text = "Dependence power"
    End With
    
    With .Axes(xlValue)
        .MaximumScale = 10
        .MinimumScale = 0
        .MajorUnit = 1
        .HasMajorGridlines = False
        .Format.Line.Weight = 2
        .HasTitle = True
        .AxisTitle.Characters.text = "Independence power"
    End With

    With .SeriesCollection(1)
    .XValues = Dep_power
    .values = Ind_power
    .HasDataLabels = True
    .DataLabels.Position = xlLabelPositionAbove
    'Design markers
    .MarkerStyle = 1
    .MarkerSize = 10
    .Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent6
    'Add the labels
    For Each Name In names
        aux = aux + 1
        .Points(aux).DataLabel.text = Name
    Next Name
    End With
    
    'create horizontal line
    .SeriesCollection.NewSeries
    With .SeriesCollection(2)
    .XValues = Array(0, 10)
    .values = Array(5, 5)
    .MarkerStyle = xlMarkerStyleNone
    .Border.Color = vbBlack
    .Format.Line.Weight = 2
    End With
    
    'create vertical line
    .SeriesCollection.NewSeries
    With .SeriesCollection(3)
    .XValues = Array(5, 5)
    .values = Array(0, 10)
    .MarkerStyle = xlMarkerStyleNone
    .Border.Color = vbBlack
    .Format.Line.Weight = 2
    End With
    
    
    Dim captionShape As Shape
    Dim captionTop As Double, captionHeight As Double
    
    captionHeight = 40
    
    ' Position caption below plot area + some margin
    captionTop = .PlotArea.Top + .PlotArea.Height + 30
    
    Set captionShape = .Shapes.AddTextbox(msoTextOrientationHorizontal, _
        .PlotArea.Left, captionTop, .PlotArea.Width, captionHeight)
        
    With captionShape.TextFrame
        .Characters.text = "MICMAC Quadrants: Top-left = Driving Variables; Top-right = Linkage Variables; Bottom-left = Autonomous Variables; Bottom-right = Dependent Variables."
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .AutoSize = False
        '.WordWrap = True
    End With
    captionShape.Line.Visible = msoFalse
End With




End Sub
