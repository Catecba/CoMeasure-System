Attribute VB_Name = "Report"
Sub ExportPDF()
    Dim wsReport As Worksheet
    Dim ws As Worksheet, wsim As Worksheet, wsInput As Worksheet, wsact As Worksheet
    Dim tbl As ListObject
    Dim eName As Variant
    Dim elemNames As Variant
    Dim userFolder As Variant
    Dim pic As Shape
    Dim topPos As Double
    
    Set wsim = Worksheets("Structuring")
    Set wsInput = Worksheets("Input")
    Set wsact = Worksheets("Activity list")
    

    ' Create a temporary sheet for the report
    Set wsReport = ThisWorkbook.Sheets.Add
    
    On Error GoTo ErrHandler  ' Ensure temp sheet is deleted even if error occurs
    
    ' Add project information
    wsReport.Range("A2").Value = "Project Name:"
    wsReport.Range("B2").Value = Range("PrjName").Value
    wsReport.Range("A3").Value = "Perspective used:"
    wsReport.Range("B3").Value = Range("PrjPersp").Value
    
    wsReport.Columns("A:B").AutoFit

    ' Initialize top position in points (below project info) and get the page bottom
    topPos = wsReport.Range("A5").Top
    pageHeight = wsReport.Range("A32").Top + wsReport.Range("A32").Height
    Npage = pageHeight
    
    ' Ask user for folder path
    userFolder = InputBox("Please enter the folder path where you want to save the PDF", "Save PDF Folder")
    If userFolder = "" Then
        GoTo Cleanup
    End If
    
    userFolder = userFolder & Application.PathSeparator & Range("PrjName").Value & Format(Now(), "dd_mm_yyyy")
    
    ' List of element/table names
    elemNames = Array("Swimlane", "MICMAC", "Resource", "UncTable", "CCRs", "Activities", "SimRe", "Distribution")
    
    wsInput.Activate
    ActiveWindow.DisplayFormulas = True
    wsact.Activate
    ActiveWindow.DisplayFormulas = True
    wsim.Activate
    ActiveWindow.DisplayFormulas = True
    Application.Calculation = xlCalculationManual


    ' Loop through each element
    For Each eName In elemNames
        
        Select Case eName
            Case "Swimlane"
                Set tbl = wsim.ListObjects(eName)
                ' Copy table as picture
                tbl.Range.CopyPicture
                wsReport.Activate
                wsReport.Paste
                Set pic = wsReport.Shapes(wsReport.Shapes.Count)
                
            Case "Distribution"
                If wsact.ChartObjects.Count = 0 Then
                    GoTo NextIteration
                End If
                wsact.ChartObjects(1).CopyPicture
                wsReport.Paste
                Set pic = wsReport.Shapes(wsReport.Shapes.Count)
                pic.LockAspectRatio = msoTrue
                pic.ScaleHeight 0.6, msoTrue
   
                
            Case "MICMAC"
                If wsim.ChartObjects.Count = 0 Then
                    GoTo NextIteration
                End If
                wsim.ChartObjects(1).CopyPicture
                wsReport.Paste
                Set pic = wsReport.Shapes(wsReport.Shapes.Count)
                    
                pic.ScaleHeight 0.5, msoTrue
                
                
            Case "Resource"
                Set tbl = wsInput.ListObjects(eName)
                wsInput.Activate
                ' Copy table as picture
                tbl.Range.CopyPicture
                wsReport.Activate
                wsReport.Paste
                Set pic = wsReport.Shapes(wsReport.Shapes.Count)
                
            Case "CCRs"
                Set tbl = wsInput.ListObjects(eName)
                'wsInput.Activate
                'Application.Calculation = xlCalculationManual
                'ActiveWindow.DisplayFormulas = True
                ' Copy table as picture
                tbl.Range.CopyPicture
                wsReport.Activate
                wsReport.Paste
                Set pic = wsReport.Shapes(wsReport.Shapes.Count)
                
            Case "UncTable"
                If Range("Cond").Value <> "Yes" Then
                    GoTo NextIteration
                End If
                Set tbl = wsInput.ListObjects(eName)
                'wsInput.Activate
                'Application.Calculation = xlCalculationManual
                'ActiveWindow.DisplayFormulas = True
                ' Copy table as picture
                tbl.Range.CopyPicture
                wsReport.Activate
                wsReport.Paste
                Set pic = wsReport.Shapes(wsReport.Shapes.Count)
                
            Case "Activities"
                Set tbl = wsact.ListObjects(eName)
                'wsact.Activate
                'Application.Calculation = xlCalculationManual
                'ActiveWindow.DisplayFormulas = True
                ' Copy table as picture
                tbl.Range.CopyPicture
                wsReport.Activate
                wsReport.Paste
                Set pic = wsReport.Shapes(wsReport.Shapes.Count)
            
            Case "SimRe"
                Range("SimRe").CopyPicture
                wsReport.Paste
                Set pic = wsReport.Shapes(wsReport.Shapes.Count)

        End Select
        
        ' Check if picture surpasses the page limit
        
        pic.LockAspectRatio = msoTrue
        
        If topPos + pic.Height > Npage Then 'if the bottom of the pic goes bellow the page break
            pic.Top = Npage + 15
            topPos = Npage + 15 + pic.Height
            Npage = Npage + pageHeight 'update to next page break position
        Else
            pic.Top = topPos + 15
            topPos = topPos + 15 + pic.Height
        End If
        pic.Left = wsReport.Range("A1").Left


NextIteration:
    Next eName
    
    ' Page in landscape
    With wsReport.PageSetup
        .Orientation = xlLandscape
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        
    End With
    
    ' Export to PDF
    wsReport.ExportAsFixedFormat Type:=xlTypePDF, fileName:=userFolder, Quality:=xlQualityStandard
    MsgBox "PDF successfully saved to: " & userFolder
    
Cleanup:
    Application.DisplayAlerts = False 'to not get a message on deleting the sheet
    wsReport.Delete
    Application.DisplayAlerts = True
    
    ' Restore normal view
    
    wsInput.Activate
    ActiveWindow.DisplayFormulas = False
    wsact.Activate
    ActiveWindow.DisplayFormulas = False
    wsim.Activate
    ActiveWindow.DisplayFormulas = False

    Application.Calculation = xlCalculationAutomatic

    Exit Sub
    
ErrHandler:
    MsgBox "Error exporting PDF. Check folder path and write permissions."
    Resume Cleanup
End Sub




