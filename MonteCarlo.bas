Attribute VB_Name = "MonteCarlo"
Sub MainSimulation()

Dim config As Worksheet
Dim check As Boolean

Set config = Worksheets("Configuration")

check = config.Range("Checkbox").Value

'if the checkbox is checked
If check Then
    Call standardTDABC
Else
    Call MCSimulation 'perform monte carlo simulation if not checked
End If

End Sub

Sub standardTDABC()

Dim ws As Worksheet, inputws As Worksheet
Dim act_table As ListObject, ccr_table As ListObject


Set ws = Worksheets("Activity list")
Set inputws = Worksheets("Input")

Set act_table = ws.ListObjects("Activities")
Set ccr_table = inputws.ListObjects("CCRs")

On Error GoTo Errorhandler



Total = 0
'loop through all rows of the table
For i = 1 To act_table.ListRows.Count
    entry = act_table.ListColumns("Relevant CCRs").DataBodyRange.Cells(i) 'get the relevant CCRs of the activity i
            If Trim(entry) <> "" Then  'If the there are relevant ccrs
                used_ccrs = Split(entry, " ") 'get the CCRs enumerated
        
                For Each rel_ccr In used_ccrs
                    ccr_i = Mid(rel_ccr, 4) 'only keep the number from the CCR#
                    ccr_value = ccr_table.DataBodyRange(ccr_i, 4)
                    Time_value = act_table.DataBodyRange(i, ccr_i * 2 + 3)
                    ccr_quant = act_table.DataBodyRange(i, ccr_i * 2 + 2)
                    Total = Total + ccr_quant * Time_value * ccr_value '#ccr *time distribution*ccr value
                Next
            End If

  Next i


 
MsgBox "Completed"

ws.Range("SimMean").Value = Total
ws.Range("SimStdDev").Value = CVErr(xlErrNA)
ws.Range("Sim5P").Value = CVErr(xlErrNA)
ws.Range("Sim95P").Value = CVErr(xlErrNA)
ws.Range("SimMin").Value = CVErr(xlErrNA)
ws.Range("SimMax").Value = CVErr(xlErrNA)
ws.Range("SimValues").Value = 1

Exit Sub

Errorhandler:
MsgBox "An error occurred: " & Err.Description, vbExclamation


End Sub

Sub MCSimulation()

NumIterations = InputBox("Introduce the number of iterations the Monte Carlo Simulation will have", "MC Iterations") 'Number of iterations as user input

If NumIterations = "" Then 'if no input received
    MsgBox "No input received"
    Exit Sub
End If

Dim ws As Worksheet, inputws As Worksheet, wsDiagram As Worksheet, config As Worksheet
Dim act_table As ListObject, ccr_table As ListObject, unc_table As ListObject, swim As ListObject
Dim current_branch As Collection
Dim Pathway_Completed As Boolean
Dim node As Variant
Dim shp As Shape
Dim MCdata() As Variant
ReDim MCdata(1 To NumIterations)
Dim StartTime As Double
Dim SecondsElapsed As Double


Set ws = Worksheets("Activity list")
Set inputws = Worksheets("Input")
Set wsDiagram = Worksheets("Structuring")
Set config = Worksheets("Configuration")

Set act_table = ws.ListObjects("Activities")
Set ccr_table = inputws.ListObjects("CCRs")
Set unc_table = inputws.ListObjects("UncTable")
Set swim = wsDiagram.ListObjects("Swimlane")

StartTime = Timer

On Error GoTo Errorhandler

' Clear previous dictionaries
    Fathers.RemoveAll
    Branches.RemoveAll
    Probabilities.RemoveAll
    ShapeT.RemoveAll
    
'build the branches, branch probability, shape type and father-son dictionaries from theprocess map
For Each shp In wsDiagram.Shapes
    If shp.connector Then 'if we checking a connector
        If ProcMapToList.Isconnected(shp) Then 'and it is connected to smth
            With shp.ConnectorFormat
                    If ProcMapToList.AuthorizedShape(.BeginConnectedShape) And ProcMapToList.AuthorizedShape(.EndConnectedShape) Then
                        Call ConnectionTree.ConnsDic(.BeginConnectedShape, .EndConnectedShape)
                        Call ProcMapToList.BranchProb(.BeginConnectedShape, shp.Name)
                    End If
            End With
        End If
    End If
Next shp
    
For Each node In Branches.Keys
    Call ProcMapToList.branch(node)
Next node


SimMin = 0
SimMax = 0

Application.ScreenUpdating = False  'turn off worksheet updating to improve performance
Application.Calculation = xlCalculationManual

For Iteration = 1 To NumIterations
    Application.Calculate
    Total = 0

    Set current_branch = Branches("1") 'we start with the main branch
    Pathway_Incompleted = True
    i = 1

    Do While Pathway_Incompleted 'While there is no activity as the last node of a branch
        node = current_branch(i)  ' node of the  branch
        activity_cost = 0
        
        If ShapeT(node) = msoShapeFlowchartDecision Then 'if we are at a decision node
            
            Next_b = InvFunction.DISCR_INV(Rnd(), NumberSeqString(Fathers(node).Count), CollToString(Probabilities(node)))  'use a discrete inv distribution to choose the next branch to go to byoutputting the index of the respective branch in the Fathers' Item collection
            Set current_branch = Branches(Fathers(node)(Next_b)) 'get the new branch, which is named after the node that it starts it

            i = 1 'start from the begining of the branch
             
        ElseIf ShapeT(node) = msoShapeFlowchartProcess Or ShapeT(node) = msoShapeFlowchartAlternateProcess Then
            RowNum = Application.Match(node, act_table.ListColumns("Activities").DataBodyRange, 0) 'get the table row number of the node
            entry = act_table.ListColumns("Relevant CCRs").DataBodyRange.Cells(RowNum) 'get the relevant CCRs of the node
            
            If Trim(entry) <> "" Then  'If the there are relevant ccrs
                used_ccrs = Split(entry, " ") 'get the CCRs enumerated
        
                For Each rel_ccr In used_ccrs
                    ccr_i = Mid(rel_ccr, 4) 'only keep the number from the CCR#
                    ccr_value = ccr_table.DataBodyRange(ccr_i, 4)
                    Time_value = act_table.DataBodyRange(RowNum, ccr_i * 2 + 3)
                    ccr_quant = act_table.DataBodyRange(RowNum, ccr_i * 2 + 2)
                    activity_cost = activity_cost + ccr_quant * Time_value * ccr_value '#ccr *time distribution*ccr value
                Next
                
                If inputws.Range("Cond") = "Yes" Then 'if the user wants to use the uncertainty table
                    UncRow = Application.Match(node, unc_table.ListColumns("Targeted Activity ").DataBodyRange, 0) 'check if there is any uncertainty related with this node
                    If Not IsError(UncRow) Then 'if there is a match
                        activity_cost = activity_cost * unc_table.ListColumns("Probability Distribution").DataBodyRange.Cells(UncRow) 'multiply that activity equation by the uncertainty value
                    End If
                End If
                Total = Total + activity_cost 'add activity cost to the total of that iteration
            End If
            
            If i = current_branch.Count Then 'If the last node of the branch is an activity node = end of branch
                Pathway_Incompleted = False
                Exit Do
            End If
            
            i = i + 1 'next node
        Else
            MsgBox "Error on the shape type"
            Exit Sub
        End If
        
    Loop

    MCdata(Iteration) = Total
        
Next

    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True 'turn back to normal

SecondsElapsed = Round(Timer - StartTime, 2) '
 
MsgBox "Completed in " & SecondsElapsed & " seconds"

'Print the results from the simulation
ws.Range("SimMean").Value = WorksheetFunction.Average(MCdata)
ws.Range("SimStdDev").Value = WorksheetFunction.StDevP(MCdata)
ws.Range("Sim5P").Value = Application.WorksheetFunction.Percentile(MCdata, 0.05)
ws.Range("Sim95P").Value = Application.WorksheetFunction.Percentile(MCdata, 0.95)
ws.Range("SimMin").Value = Application.WorksheetFunction.Min(MCdata)
ws.Range("SimMax").Value = Application.WorksheetFunction.Max(MCdata)
ws.Range("SimValues").Value = NumIterations 'to include the number of iterations in the report

'If the results aren't constant
If WorksheetFunction.StDevP(MCdata) <> 0 Then
    ' Write values for histogram
    config.Columns("U:U").ClearContents
    config.Range("U3").Resize(CLng(NumIterations), 1).Value = Application.Transpose(MCdata)
    Call Histogram
End If

Exit Sub

Errorhandler:
MsgBox "An error occurred: " & Err.Description, vbExclamation

End Sub


Function NumberSeqString(n As Long) As String
Dim i As Long
Dim parts() As String
    
ReDim parts(1 To n)
    
For i = 1 To n
    parts(i) = CStr(i)
Next i
    
NumberSeqString = Join(parts, ";")    ' Join into one string with ;
    
End Function


Function CollToString(coll As Collection) As String
Dim i As Long
Dim arr() As String
    
ReDim arr(1 To coll.Count)
    
For i = 1 To coll.Count    ' Fill array from collection
    arr(i) = CStr(coll.Item(i))
Next i
    
CollToString = Join(arr, ";")    ' Create the string with ;

End Function


Sub Histogram()
'
    Sheets("Activity list").Activate
    Sheets("Activity list").Shapes.AddChart2(366, xlHistogram).Select
    ActiveChart.ChartTitle.Caption = "Simulation Distribution"
    ActiveChart.Axes(xlCategory).Select
    Selection.SetProperty "NumberFormat_FormatCategory", "Number"
    On Error Resume Next
    ActiveChart.SetSourceData (Sheets("Configuration").Range("U3", Sheets("Configuration").Range("U3").End(xlDown)))
End Sub


