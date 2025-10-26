VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    
    Dim wsInput As Worksheet, wsList As Worksheet
    Dim pathway As String
    
    Set wsInput = Worksheets("Input")
    Set wsList = Worksheets("Activity list")
    
    'Work around to remove the pathway for the add-in file from the functions in the cell
    
' Find TRIANG_INV inside the formula
pos = InStr(1, Range("Pathway").Formula, "TRIANG_INV(0.1,1,2,3)", vbTextCompare)

If pos > 0 Then
    ' Everything before TRIANG_INV is the "pathway" part (until pos -1)
    pathway = Left(Range("Pathway").Formula, pos - 1)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    If Len(pathway) > 0 Then
        ' Remove pathway in input and list sheets
        wsInput.Cells.Replace What:=pathway, Replacement:="=", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        wsList.Cells.Replace What:=pathway, Replacement:="=", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    End If
    
    ' Clean the formula itself
    Range("Pathway").Formula = "=TRIANG_INV(0.1,1,2,3)" '& Mid(Range("Pathway").Formula, pos)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End If

End Sub


